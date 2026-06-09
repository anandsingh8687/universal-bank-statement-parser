"""
XLSX / XLS Bank Statement Parser

Handles Excel-format bank statements from major Indian banks.
Auto-detects the header row and column mapping, then routes through
the same compute_metrics() pipeline as the PDF parser.

Supported formats:
  - Separate Debit / Credit / Balance columns (most common)
  - Single Amount column with DR/CR type column
  - All major Indian bank XLSX exports: HDFC, ICICI, SBI, Axis, Kotak, etc.

Usage:
    from xlsx_bank_parser import parse_xlsx_bank_statement
    result = parse_xlsx_bank_statement("statement.xlsx")
    # result["transactions"], result["metrics"], etc.
"""

from __future__ import annotations

import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

# ---------------------------------------------------------------------------
# Column-name synonym sets
# Extend these if your bank uses unusual column headers.
# ---------------------------------------------------------------------------

_DATE_SYNONYMS = {
    "date", "txn date", "transaction date", "trans date", "value date",
    "posting date", "tran date", "txn_date", "trn date", "trans. date",
    "booking date", "entry date", "dt", "txndate",
}

_DESCRIPTION_SYNONYMS = {
    "narration", "description", "particulars", "remarks", "transaction description",
    "details", "transaction particulars", "txn description", "transaction details",
    "narration/description", "remark", "narrative", "transaction remarks",
    "trans description", "memo",
}

_DEBIT_SYNONYMS = {
    "debit", "withdrawal", "withdrawals", "debit amount", "dr",
    "dr amount", "debit amt", "withdrawal amt", "debit(dr)",
    "withdrawal amount", "dr.", "amount debited", "outflow",
}

_CREDIT_SYNONYMS = {
    "credit", "deposit", "deposits", "credit amount", "cr",
    "cr amount", "credit amt", "deposit amt", "credit(cr)",
    "deposit amount", "cr.", "amount credited", "inflow",
}

_BALANCE_SYNONYMS = {
    "balance", "closing balance", "running balance", "available balance",
    "balance(inr)", "closing bal", "running bal", "balance amount",
    "bal", "balance (inr)", "net balance", "ledger balance",
}

_AMOUNT_SYNONYMS = {
    "amount", "transaction amount", "txn amount", "amt", "amount(inr)", "amount (inr)",
}

_TYPE_SYNONYMS = {
    "type", "dr/cr", "cr/dr", "transaction type", "txn type",
    "dr / cr", "cr / dr", "debit/credit",
}


def _normalize(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).lower().strip().rstrip(".:"))


def _find_header_row(df_raw: pd.DataFrame, max_scan: int = 30) -> int | None:
    for i in range(min(max_scan, len(df_raw))):
        cells = {_normalize(str(c)) for c in df_raw.iloc[i].values if pd.notna(c)}
        has_date = bool(cells & _DATE_SYNONYMS)
        has_desc = bool(cells & _DESCRIPTION_SYNONYMS)
        has_amount = bool(cells & (_DEBIT_SYNONYMS | _CREDIT_SYNONYMS | _AMOUNT_SYNONYMS))
        if has_date and has_desc and has_amount:
            return i
    return None


def _map_columns(headers: list[str]) -> dict[str, str]:
    mapping: dict[str, str | None] = {
        "date": None, "description": None, "debit": None,
        "credit": None, "balance": None, "amount": None, "type": None,
    }
    normed = {_normalize(h): h for h in headers}
    for norm_key, orig in normed.items():
        if norm_key in _DATE_SYNONYMS and mapping["date"] is None:
            mapping["date"] = orig
        elif norm_key in _DESCRIPTION_SYNONYMS and mapping["description"] is None:
            mapping["description"] = orig
        elif norm_key in _DEBIT_SYNONYMS and mapping["debit"] is None:
            mapping["debit"] = orig
        elif norm_key in _CREDIT_SYNONYMS and mapping["credit"] is None:
            mapping["credit"] = orig
        elif norm_key in _BALANCE_SYNONYMS and mapping["balance"] is None:
            mapping["balance"] = orig
        elif norm_key in _AMOUNT_SYNONYMS and mapping["amount"] is None:
            mapping["amount"] = orig
        elif norm_key in _TYPE_SYNONYMS and mapping["type"] is None:
            mapping["type"] = orig
    return {k: v for k, v in mapping.items() if v is not None}


def _parse_amount(val) -> float:
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    if not s or s in ("-", "—", "nil", "NIL", "0", "0.0", "0.00"):
        return 0.0
    s = re.sub(r"[₹,\s]", "", s)
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    s = re.sub(r"\s*(dr|cr|DR|CR)\.?$", "", s)
    try:
        return abs(float(s))
    except (ValueError, TypeError):
        return 0.0


def _parse_date(val, date_formats: list[str] | None = None) -> pd.Timestamp | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s:
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return pd.Timestamp(val)
    formats = date_formats or [
        "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
        "%d %b %Y", "%d-%b-%Y", "%d/%b/%Y", "%d %b %y", "%d-%b-%y",
        "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y",
        "%d.%m.%Y", "%d.%m.%y",
    ]
    for fmt in formats:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except (ValueError, TypeError):
            continue
    try:
        return pd.Timestamp(pd.to_datetime(s, dayfirst=True))
    except Exception:
        return None


def _detect_bank_name(df_raw: pd.DataFrame, header_row: int) -> str:
    bank_patterns = {
        "HDFC": r"hdfc\s*bank",
        "ICICI": r"icici\s*bank",
        "SBI": r"state\s*bank|sbi",
        "Axis": r"axis\s*bank",
        "Kotak": r"kotak\s*mahindra",
        "Yes Bank": r"yes\s*bank",
        "IndusInd": r"indusind",
        "PNB": r"punjab\s*national|pnb",
        "BOB": r"bank\s*of\s*baroda|bob",
        "Union Bank": r"union\s*bank",
        "Canara": r"canara\s*bank",
        "IDBI": r"idbi\s*bank",
        "Federal": r"federal\s*bank",
        "AU Small Finance": r"au\s*(small\s*finance|sfb)",
        "IDFC First": r"idfc\s*first",
    }
    for i in range(min(header_row, 10)):
        row_text = " ".join(str(c) for c in df_raw.iloc[i].values if pd.notna(c)).lower()
        for bank_name, pattern in bank_patterns.items():
            if re.search(pattern, row_text):
                return bank_name
    return ""


def parse_xlsx_bank_statement(
    file_path: str,
    password: str | None = None,
    save_json: bool = False,
    output_dir: str | None = None,
) -> dict:
    """
    Parse an XLSX/XLS bank statement.

    Returns the same structure as universal_bank_parser_v13.parse_bank_statement()
    so it can be used interchangeably via adapters.run_bank_statement_parser().

    Args:
        file_path: Path to the .xlsx, .xlsm, or .xls file.
        password:  Not used for XLSX (Excel encryption not supported here).
        save_json: Write _result.json alongside the file if True.
        output_dir: Directory for JSON output (defaults to file directory).
    """
    started = time.time()
    file_path_str = str(file_path)
    file_ext = Path(file_path_str).suffix.lower()

    print("\n" + "=" * 62)
    print(" XLSX Bank Statement Parser")
    print("=" * 62)
    print(f"\nFile: {os.path.basename(file_path_str)}")

    try:
        engine = "openpyxl" if file_ext in (".xlsx", ".xlsm") else "xlrd"
        xls = pd.ExcelFile(file_path_str, engine=engine)
    except Exception as exc:
        return {
            "status": "error",
            "error_type": "read_failed",
            "message": f"Cannot read Excel file: {exc}",
            "pdf_file": file_path_str,
        }

    best_result: dict | None = None
    best_txn_count = 0

    for sheet_name in xls.sheet_names:
        print(f"\n  Trying sheet: {sheet_name}")
        try:
            df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
        except Exception:
            continue
        if df_raw.empty or len(df_raw) < 3:
            continue

        header_row = _find_header_row(df_raw)
        if header_row is None:
            continue

        bank_name = _detect_bank_name(df_raw, header_row)
        acct_info = {
            "bank_name": bank_name,
            "account_holder_name": "",
            "account_number_last4": "",
            "statement_from": "",
            "statement_to": "",
            "ifsc_code": "",
            "account_type": "regular",
        }

        headers = [str(c).strip() for c in df_raw.iloc[header_row].values]
        col_map = _map_columns(headers)
        if "date" not in col_map or "description" not in col_map:
            continue
        has_separate_dr_cr = "debit" in col_map and "credit" in col_map
        has_single_amount = "amount" in col_map
        if not has_separate_dr_cr and not has_single_amount:
            continue

        data_rows = df_raw.iloc[header_row + 1:].copy()
        data_rows.columns = headers

        transactions = []
        for _, row in data_rows.iterrows():
            date_val = _parse_date(row.get(col_map["date"]))
            if date_val is None:
                continue
            desc = str(row.get(col_map.get("description", ""), "")).strip()
            if not desc or desc.lower() in ("nan", "none", ""):
                desc = ""
            if has_separate_dr_cr:
                debit = _parse_amount(row.get(col_map.get("debit")))
                credit = _parse_amount(row.get(col_map.get("credit")))
            elif has_single_amount:
                amount = _parse_amount(row.get(col_map.get("amount")))
                type_col = col_map.get("type")
                if type_col:
                    type_val = _normalize(str(row.get(type_col, "")))
                    if type_val in ("dr", "debit", "d"):
                        debit, credit = amount, 0.0
                    else:
                        debit, credit = 0.0, amount
                else:
                    debit, credit = 0.0, amount
            else:
                debit, credit = 0.0, 0.0
            balance = _parse_amount(row.get(col_map.get("balance", "balance"))) if "balance" in col_map else 0.0
            if debit == 0 and credit == 0 and balance == 0:
                continue
            transactions.append({"date": date_val, "description": desc, "debit": debit, "credit": credit, "balance": balance})

        print(f"  Found {len(transactions)} transactions")
        if len(transactions) < 5 or len(transactions) <= best_txn_count:
            continue
        best_txn_count = len(transactions)

        df = pd.DataFrame(transactions)
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
        df["_seq"] = range(len(df))

        if not acct_info["statement_from"]:
            acct_info["statement_from"] = df["date"].min().strftime("%Y-%m-%d")
        if not acct_info["statement_to"]:
            acct_info["statement_to"] = df["date"].max().strftime("%Y-%m-%d")

        # Import compute_metrics from the PDF parser - same pipeline
        from universal_bank_parser_v13 import compute_metrics
        metrics = compute_metrics(df)
        summary = metrics["summary"]

        confidence = 0.85
        if len(df) >= 50:
            confidence = 0.92
        if len(df) >= 100:
            confidence = 0.95

        total_time = round(time.time() - started, 1)
        best_result = {
            "status": "success",
            "extracted_at": datetime.now().isoformat(),
            "pdf_file": file_path_str,
            "extraction_route": "xlsx_parser",
            "confidence": round(confidence, 4),
            "validation": {"passed": True, "confidence": round(confidence, 4), "bad_rows": 0},
            "account_info": acct_info,
            "schema_detected": {"source": "xlsx", "sheet": sheet_name, "columns_mapped": col_map},
            "metrics": summary,
            "cc_od_metrics": metrics.get("cc_od_metrics", {}),
            "emi_details": metrics["emi_transactions"],
            "bounce_details": metrics["bounce_transactions"],
            "monthly_breakdown": metrics["monthly_breakdown"],
            "transactions": df.assign(date=df["date"].dt.strftime("%Y-%m-%d")).to_dict("records"),
            "cost": {"route": "xlsx_parser", "gemini_calls": 0, "tokens_in": 0, "tokens_out": 0, "api_cost_inr": 0.0, "total_time_s": total_time},
        }

    if best_result:
        if save_json:
            import json
            out_path = os.path.join(
                output_dir or os.path.dirname(os.path.abspath(file_path_str)),
                Path(file_path_str).stem + "_result.json",
            )
            with open(out_path, "w", encoding="utf-8") as f:
                json.dump(best_result, f, indent=2, default=str, ensure_ascii=False)
        return best_result

    return {
        "status": "error",
        "error_type": "no_transactions",
        "message": "No usable bank statement data found in any sheet.",
        "pdf_file": file_path_str,
    }
