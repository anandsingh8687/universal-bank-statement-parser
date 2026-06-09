"""
Universal Indian Bank Statement Parser - v13

Key capabilities:
1. Multi-tier extraction pipeline: regex → x-position → Gemini LLM → Gemini Vision
2. EMI and bounce detection using rule-based classification and optional LLM pass
3. Cashflow metrics: turnover, ABB, cash deposit ratio, disbursement filtering
4. CC/OD account support with utilization-aware average balance
5. Stable ordering preserves source sequence within the same date.

Setup:
- Set GEMINI_API_KEY environment variable (required for Tier 3 and Tier 4)
- Install dependencies: pip install -r requirements.txt
- Run standalone: python universal_bank_parser_v13.py statement.pdf
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from typing import Optional

import pandas as pd
import pdfplumber
from google import genai
from google.genai import types

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
# Set GEMINI_API_KEY in your environment (e.g. in .env):
#   GEMINI_API_KEY=your_key_here
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
MODEL_NAME = "gemini-2.5-flash-lite"
MAX_OUTPUT_TOKENS = 8192
MAX_WORKERS = 5
MAX_RETRIES = 4
RETRY_DELAY_SEC = 8
INPUT_PRICE_PER_1M = 0.10
OUTPUT_PRICE_PER_1M = 0.40
USD_TO_INR = 84.0
CONFIDENCE_THRESHOLD = 0.95
DIGITAL_CHARS_PER_PG = 100

TOKEN_LOG: list[tuple[int, int]] = []
EMI_BOUNCE_DICT_PATH = os.path.join(
    os.path.dirname(__file__),
    "emi_bounce_master_dictionary.md",
)

# ---------------------------------------------------------------------------
# EMI / BOUNCE KEYWORD LISTS
# (These lists power the rule-based classifier. Extend them as needed.)
# ---------------------------------------------------------------------------

EMI_KW = [
    "emi", "loan", "instalment", "installment", "nach", "ecs debit",
    "equated", "mortgage", "housing loan", "car loan", "auto loan",
    "personal loan", "home loan", "loan repay", "achd-",
    # Add common NBFC / lender tokens here
]

BOUNCE_KW = [
    "bounce", "dishonour", "dishonored", "unpaid", "insufficient", "insuff",
    "ecs return", "nach return", "cheque return", "chq ret", "si return",
    "mandate return", "payment failed", "rejected", "inward return",
    "outward return", "clg return",
]


def normalize_rule_text(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(value).lower())).strip()


def _normalize_rule_phrase(value: str) -> str:
    return normalize_rule_text(value.replace("#", " "))


def load_master_rulebook(path: str) -> dict:
    """Load the EMI/bounce keyword rulebook from the markdown dictionary file."""
    if not os.path.exists(path):
        return {}
    rules = {}
    current = None
    section = None
    with open(path, encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.rstrip()
            if line.startswith("## "):
                current = line[3:].strip()
                rules[current] = {"allowed_flows": set(), "prefixes": [], "keywords": []}
                section = None
            elif current and line.startswith("- Allowed flows:"):
                flows = [p.strip() for p in line.split(":", 1)[1].split(",")]
                rules[current]["allowed_flows"] = {f for f in flows if f}
            elif current and line.startswith("### Strong prefixes"):
                section = "prefixes"
            elif current and line.startswith("### Top keywords"):
                section = "keywords"
            elif current and line.startswith("### "):
                section = None
            elif current and section == "prefixes" and line.startswith("- "):
                prefix = _normalize_rule_phrase(line[2:].split("(", 1)[0].strip())
                if prefix:
                    rules[current]["prefixes"].append(prefix)
            elif current and section == "keywords" and line.startswith("- "):
                keyword = _normalize_rule_phrase(line[2:].split("(", 1)[0].strip())
                if keyword:
                    rules[current]["keywords"].append(keyword)
    return rules


MASTER_RULEBOOK = load_master_rulebook(EMI_BOUNCE_DICT_PATH)

# Extended token sets used by the classifier
EMI_SPECIFIC_TOKENS = (
    "emi", "nach", "ecs", "mandate", "instal", "install", "loan", "autopay",
    # Add lender-specific tokens here as needed
)
EMI_LENDER_TOKENS = (
    "bajaj finance", "bajaj finserv", "muthoot", "neogrowth", "smfg",
    "hdbfs", "fullerton", "manappuram", "aditya birla", "tata capital",
    "kreditbee", "loan",
    # Add more NBFC/lender name tokens here
)
BNPL_AUTOPAY_TOKENS = ("simpl", "getsimpl", "lazypay", "postpe", "axio")
BOUNCE_SIGNAL_TOKENS = (
    "bounce", "dishonour", "dishonor", "unpaid", "insufficient", "insuff",
    "rtn", "rejected", "decline", "declined", "failed", "failure",
)
RETURN_LIKE_TOKENS = ("return", "reversal", "reversed", "reverse")
CHARGE_SIGNAL_TOKENS = ("charge", "charges", "chrg", "chg", "fee")

# ---------------------------------------------------------------------------
# GEMINI CLIENT HELPERS
# ---------------------------------------------------------------------------

_client = None


def _get_client():
    global _client
    if not GEMINI_API_KEY:
        return None
    if _client is None:
        _client = genai.Client(api_key=GEMINI_API_KEY)
    return _client


def _gemini_text(prompt: str) -> str:
    client = _get_client()
    if client is None:
        return ""
    cfg = types.GenerateContentConfig(
        temperature=0.0,
        max_output_tokens=MAX_OUTPUT_TOKENS,
        thinking_config=types.ThinkingConfig(thinking_budget=0),
    )
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME, contents=[prompt], config=cfg
            )
            if response.usage_metadata:
                TOKEN_LOG.append((
                    response.usage_metadata.prompt_token_count or 0,
                    response.usage_metadata.candidates_token_count or 0,
                ))
            return (response.text or "").strip()
        except Exception as exc:
            if "503" in str(exc) and attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY_SEC * attempt)
            else:
                return ""
    return ""


_LLM_EMI_BATCH_SIZE = 50

_LLM_EMI_PROMPT = (
    "You are analyzing Indian bank statement transactions. "
    "For each debit transaction below, classify if it is:\n"
    " EMI – a loan/credit card EMI payment\n"
    " EMI_BOUNCE – an EMI bounce or return\n"
    " BOUNCE_CHARGE – a bounce/penalty charge\n"
    " OTHER – none of the above\n\n"
    "Return ONLY a JSON array: [{\"idx\": 0, \"classification\": \"EMI\"}]\n\n"
    "Transactions:\n"
)

_LLM_CLASS_MAP = {
    "EMI": {"txn_category": "EMI", "is_emi": True, "is_bounce": False, "bounce_category": None},
    "EMI_BOUNCE": {"txn_category": "EMI Payment Bounce", "is_emi": False, "is_bounce": True, "bounce_category": "EMI Payment Bounce"},
    "BOUNCE_CHARGE": {"txn_category": "EMI Bounce Charges", "is_emi": False, "is_bounce": True, "bounce_category": "EMI Bounce Charges"},
}


def _llm_emi_config() -> tuple[bool, int, int]:
    """Return (enabled, max_transactions, min_debit) for the LLM EMI pass.

    Control via environment variables:
      BANK_PARSER_LLM_EMI_ENABLED=true
      BANK_PARSER_LLM_EMI_MAX_TRANSACTIONS=120
      BANK_PARSER_LLM_EMI_MIN_DEBIT=500
    """
    enabled = os.getenv("BANK_PARSER_LLM_EMI_ENABLED", "false").lower() in {"1", "true", "yes", "on"}
    max_transactions = int(os.getenv("BANK_PARSER_LLM_EMI_MAX_TRANSACTIONS", "120") or "120")
    min_debit = int(os.getenv("BANK_PARSER_LLM_EMI_MIN_DEBIT", "500") or "500")
    try:
        # If integrated into a larger app, you can override from app settings here
        # from app.core.config import get_settings
        # settings = get_settings()
        # enabled = bool(settings.bank_parser_llm_emi_enabled)
        pass
    except Exception:
        pass
    return enabled, max(0, max_transactions), max(0, min_debit)


def _redact_llm_narration(text: str) -> str:
    """Minimize PII before sending narrations to the LLM."""
    text = re.sub(r"\b[A-Z]{5}[0-9]{4}[A-Z]\b", "[PAN]", str(text), flags=re.I)
    text = re.sub(r"\b\d{9,18}\b", "[NUMBER]", text)
    text = re.sub(r"\b\d{4}\s?\d{4}\s?\d{4}\b", "[AADHAAR]", text)
    return text[:180]


def _llm_classify_emi_bounce(narrations: list[dict]) -> list[dict]:
    if not narrations:
        return []
    results: list[dict] = []
    for start in range(0, len(narrations), _LLM_EMI_BATCH_SIZE):
        batch = narrations[start: start + _LLM_EMI_BATCH_SIZE]
        prompt = _LLM_EMI_PROMPT + json.dumps(batch, ensure_ascii=False)
        raw = _gemini_text(prompt)
        if not raw:
            continue
        raw = raw.strip()
        if raw.startswith("```"):
            raw = re.sub(r"^```[a-z]*\n?", "", raw)
            raw = re.sub(r"\n?```$", "", raw)
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, list):
                for item in parsed:
                    if isinstance(item, dict) and "idx" in item and "classification" in item:
                        results.append({"idx": int(item["idx"]), "classification": str(item["classification"]).upper()})
        except (json.JSONDecodeError, ValueError):
            continue
    return results


def _gemini_vision(image_bytes: bytes, prompt: str) -> str:
    client = _get_client()
    if client is None:
        return ""
    cfg = types.GenerateContentConfig(temperature=0.0, max_output_tokens=MAX_OUTPUT_TOKENS)
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=[types.Part.from_bytes(data=image_bytes, mime_type="image/png"), prompt],
                config=cfg,
            )
            if response.usage_metadata:
                TOKEN_LOG.append((
                    response.usage_metadata.prompt_token_count or 0,
                    response.usage_metadata.candidates_token_count or 0,
                ))
            return (response.text or "").strip()
        except Exception as exc:
            if "503" in str(exc) and attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY_SEC * attempt)
            else:
                return ""
    return ""


def _parse_json(raw: str) -> dict:
    raw = re.sub(r"^```json\s*|^```\s*|\s*```$", "", raw.strip(), flags=re.MULTILINE)
    try:
        return json.loads(raw)
    except Exception:
        match = re.search(r"{[\s\S]*}", raw)
        if match:
            try:
                return json.loads(match.group())
            except Exception:
                pass
    return {}


# ---------------------------------------------------------------------------
# LAYER 1 - PDF TYPE DETECTION
# ---------------------------------------------------------------------------

def detect_pdf(pdf_path: str, password: Optional[str] = None) -> dict:
    try:
        kwargs = {"password": password} if password else {}
        with pdfplumber.open(pdf_path, **kwargs) as pdf:
            pages_text = []
            total_chars = 0
            for page in pdf.pages:
                raw = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                pages_text.append(raw)
                total_chars += len(raw)
            avg = total_chars / len(pages_text) if pages_text else 0
            return {
                "type": "digital" if avg >= DIGITAL_CHARS_PER_PG else "scanned",
                "pages": pages_text,
                "num_pages": len(pages_text),
                "avg_chars": round(avg, 1),
                "pdf_path": pdf_path,
            }
    except Exception as exc:
        err = str(exc).lower()
        if any(t in err for t in ("password", "encrypt", "decrypt")):
            return {"type": "encrypted", "pages": [], "num_pages": 0, "error": str(exc), "pdf_path": pdf_path}
        return {"type": "error", "pages": [], "num_pages": 0, "error": str(exc), "pdf_path": pdf_path}


# ---------------------------------------------------------------------------
# LAYER 2 - SCHEMA DISCOVERY
# ---------------------------------------------------------------------------

_SCHEMA_PROMPT = """
You are analyzing the first page of an Indian bank account statement.
Study the column headers and first 2-3 transaction rows in the text below.
Return ONLY a JSON object.

{
  "bank_name": "full bank name",
  "account_holder": "name as shown",
  "account_number_last4": "last 4 digits",
  "ifsc_code": "full 11-char IFSC code or null",
  "account_type": one of ["regular","cc_od"],
  "statement_from": "YYYY-MM-DD or null",
  "statement_to": "YYYY-MM-DD or null",
  "date_format": one of ["DD/MM/YY","DD/MM/YYYY","DD-MM-YYYY","DD MMM YYYY","DDMMYYYY","DD MMMM YYYY","MMM DD, YYYY","MMMM DD, YYYY"],
  "has_value_date_col": true or false,
  "has_ref_col": true or false,
  "amount_mode": one of ["dual","single_with_drcr","single_signed"],
  "debit_col_name": "column name for money OUT",
  "credit_col_name": "column name for money IN",
  "multiline_narration": true or false,
  "opening_balance": number or null
}

FIRST PAGE TEXT:
"""

# ... (heuristic schema detection, date parsing, and all tier functions follow)
# See full implementation below


_DATE_FORMATS = [
    ("DD/MM/YY",     r"\d{2}/\d{2}/\d{2}",           "%d/%m/%y"),
    ("DD/MM/YYYY",   r"\d{2}/\d{2}/\d{4}",           "%d/%m/%Y"),
    ("DD-MM-YYYY",   r"\d{2}-\d{2}-\d{4}",           "%d-%m-%Y"),
    ("DD-MMM-YYYY",  r"\d{2}-[A-Za-z]{3}-\d{4}",      "%d-%b-%Y"),
    ("DD-MMM-YY",    r"\d{2}-[A-Za-z]{3}-\d{2}",      "%d-%b-%y"),
    ("DD MMM YYYY",  r"\d{2}\s[A-Za-z]{3}\s\d{4}",  "%d %b %Y"),
    ("DD MMMM YYYY", r"\d{2}\s[A-Za-z]{4,9}\s\d{4}","%d %B %Y"),
    ("MMM DD, YYYY", r"[A-Za-z]{3}\s\d{2},\s\d{4}", "%b %d, %Y"),
    ("YYYY-MM-DD",   r"\d{4}-\d{2}-\d{2}",           "%Y-%m-%d"),
    ("DDMMYYYY",     r"\d{8}",                          "%d%m%Y"),
]
_DATE_PAT = {name: pat for name, pat, _ in _DATE_FORMATS}
_FALLBACK_DATE_PAT = r"\d{2}[/-]\d{2}[/-]\d{2,4}|\d{2}\s[A-Za-z]{3,9}\s\d{4}|[A-Za-z]{3,9}\s\d{2},\s\d{4}|\d{2}-[A-Za-z]{3}-\d{2,4}|\d{4}[/-]\d{2}[/-]\d{2}"

AMOUNT_RE = re.compile(r"(?<!\d)(-?\d{1,3}(?:,\d{3})*\.\d{2})(?!\d)")


def parse_date_universal(raw: str) -> Optional[str]:
    raw = raw.strip()
    for _, _, fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    for fmt in ("%d %B %Y", "%d-%b-%Y", "%d/%b/%Y", "%d-%B-%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return None


def pa(value) -> float:
    return float(str(value).replace(",", "")) if value else 0.0


def compact_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def clean_description_text(value: str) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return re.sub(r"\s+", " ", text).strip()


# ---------------------------------------------------------------------------
# NOTE: The full implementation of all tier functions (tier1_regex,
# tier1_flexible_text, tier2_xpos, tier3_gemini_pages, tier4_vision),
# the validate_and_fix pipeline, classify_transaction, compute_metrics,
# and parse_bank_statement are available in the private source repo.
#
# This file shows the architecture and key components. To integrate:
# 1. Set up GEMINI_API_KEY in your environment.
# 2. Install requirements: pip install -r requirements.txt
# 3. Call parse_bank_statement(pdf_path) from your application.
# ---------------------------------------------------------------------------


def parse_bank_statement(
    pdf_path: str,
    password: Optional[str] = None,
    save_json: bool = True,
    output_dir: Optional[str] = None,
) -> dict:
    """
    Main entry point. Parse a bank statement PDF.

    Returns a dict with keys:
      status, account_info, metrics, transactions, emi_details,
      bounce_details, monthly_breakdown, cost, confidence, extraction_route

    Example::

        result = parse_bank_statement("statement.pdf")
        print(result["metrics"]["monthly_turnover"])
        print(result["metrics"]["bounce_count"])
    """
    # Full implementation: see private source repo.
    # Steps:
    # 1. detect_pdf() -> determine digital / scanned / encrypted
    # 2. discover_schema() -> extract bank name, date format, column layout
    # 3. Tier 1a regex -> Tier 1b flexible -> Tier 2 xpos -> Tier 3 LLM -> Tier 4 Vision
    # 4. validate_and_fix() -> balance-continuity check, confidence score
    # 5. compute_metrics() -> EMI/bounce classification, turnover, ABB, etc.
    raise NotImplementedError(
        "Full implementation is in the private source. "
        "Set up your environment (GEMINI_API_KEY, dependencies) and integrate."
    )


if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog, simpledialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    pdf_path = filedialog.askopenfilename(
        title="Select Bank Statement PDF",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    if not pdf_path:
        print("No file selected.")
        sys.exit(0)
    entered = simpledialog.askstring(
        "Password",
        f"Password for {os.path.basename(pdf_path)} (blank if none):",
        parent=root,
        show="*",
    )
    password = entered.strip() if entered and entered.strip() else None
    root.destroy()
    result = parse_bank_statement(pdf_path, password=password)
    print(json.dumps(result, indent=2, default=str))
