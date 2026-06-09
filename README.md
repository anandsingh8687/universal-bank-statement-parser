# Universal Indian Bank Statement Parser

A multi-tier PDF and XLSX bank statement parser for Indian banks with EMI detection, bounce classification, and cashflow intelligence.

## Overview

This parser extracts structured transaction data from Indian bank statements (PDF and Excel formats) and computes financial intelligence metrics including:

- **EMI detection** – identifies loan repayments using lender name matching, mandate patterns (NACH/ECS), and recurring amount analysis
- **Bounce classification** – detects cheque bounces, EMI payment bounces, inward/outward payment returns, and associated charges
- **Cashflow metrics** – monthly turnover, average bank balance (weighted daily ABB), net cash flow, cash deposit ratio
- **Disbursement filtering** – separates loan credits and inter-account transfers from real business turnover
- **CC/OD account support** – handles Cash Credit and Overdraft accounts with utilization-aware ABB

## Architecture

The parser uses a 4-tier fallback strategy:

```
Tier 1a  Schema-guided regex (date-on-line layout)
Tier 1b  Flexible text rows (pre-date narration layout)
Tier 2   X-position column classifier (pdfplumber words API)
Tier 3   Gemini LLM fallback on bad pages (selective, targeted)
Tier 4   Gemini Vision for scanned/image-only PDFs
```

Each tier produces a candidate set of transactions. The pipeline selects the best candidate using a confidence score derived from balance continuity validation.

## Setup

### Requirements

```
pdfplumber>=0.11
pandas>=2.0
google-genai>=1.38
pdf2image          # for scanned PDFs (Tier 4)
openpyxl           # for XLSX parsing
```

Install with:

```bash
pip install -r requirements.txt
```

### Environment Variables

Copy `.env.example` to `.env` and fill in the required values:

```bash
cp .env.example .env
```

| Variable | Required | Description |
|---|---|---|
| `GEMINI_API_KEY` | Yes (for LLM tiers) | Google Gemini API key. Get one at https://aistudio.google.com/ |
| `BANK_PARSER_LLM_EMI_ENABLED` | No | Set to `true` to enable LLM secondary EMI pass (default: `false`) |
| `BANK_PARSER_LLM_EMI_MAX_TRANSACTIONS` | No | Max transactions to send to LLM per statement (default: `120`) |
| `BANK_PARSER_LLM_EMI_MIN_DEBIT` | No | Minimum debit amount to consider for LLM EMI pass (default: `500`) |

> **Note:** The parser works without a Gemini API key for digital PDFs with good text quality. Tier 3 (selective LLM) and Tier 4 (vision) require the key.

## Usage

### CLI

```bash
python universal_bank_parser_v13.py statement.pdf
python universal_bank_parser_v13.py statement.pdf password123
```

A file picker dialog will open if no arguments are provided.

### Python API

```python
from universal_bank_parser_v13 import parse_bank_statement

result = parse_bank_statement(
    pdf_path="statement.pdf",
        password=None,       # PDF password if encrypted
            save_json=True,      # write _result.json alongside the PDF
            )

            # result["status"] == "success"
            # result["account_info"]    -> bank name, holder, account last 4, IFSC
            # result["metrics"]         -> turnover, ABB, EMI outflow, bounce count, etc.
            # result["transactions"]    -> list of dicts: date, description, debit, credit, balance
            # result["emi_details"]     -> EMI transactions
            # result["bounce_details"]  -> bounce transactions
            # result["monthly_breakdown"] -> per-month credits/debits/net/bounce count
            ```

            ### XLSX Statements

            ```python
            from xlsx_bank_parser import parse_xlsx_bank_statement

            result = parse_xlsx_bank_statement("statement.xlsx")
            ```

            The XLSX parser returns the same output structure as the PDF parser.

            ### Adapter (auto-routing)

            ```python
            from adapters import run_bank_statement_parser

            result = run_bank_statement_parser("statement.pdf")   # or .xlsx
            ```

            ## Output Schema

            ```json
            {
              "status": "success",
                "account_info": {
                    "bank_name": "HDFC Bank",
                        "account_holder_name": "JOHN DOE",
                            "account_number_last4": "1234",
                                "statement_from": "2024-01-01",
                                    "statement_to": "2024-03-31",
                                        "ifsc_code": "HDFC0001234",
                                            "account_type": "regular"
                                              },
                                                "metrics": {
                                                    "monthly_turnover": 150000.0,
                                                        "avg_monthly_balance": 45000.0,
                                                            "emi_outflow_latest_month": 18000.0,
                                                                "bounce_count": 2,
                                                                    "total_credits": 450000.0,
                                                                        "total_debits": 420000.0,
                                                                            "months_covered": 3,
                                                                                "cash_deposit_ratio": 0.12
                                                                                  },
                                                                                    "transactions": [
                                                                                        {
                                                                                              "date": "2024-01-05",
                                                                                                    "description": "NACH/EMI BAJAJ FINANCE",
                                                                                                          "debit": 9500.0,
                                                                                                                "credit": 0.0,
                                                                                                                      "balance": 38200.0,
                                                                                                                            "txn_category": "EMI",
                                                                                                                                  "is_emi": true,
                                                                                                                                        "is_bounce": false
                                                                                                                                            }
                                                                                                                                              ]
                                                                                                                                              }
                                                                                                                                              ```
                                                                                                                                              
                                                                                                                                              ## Supported Banks
                                                                                                                                              
                                                                                                                                              Tested with statements from HDFC Bank, ICICI Bank, SBI, Axis Bank, Kotak Mahindra Bank, Yes Bank, IndusInd Bank, AU Small Finance Bank, Bandhan Bank, PNB, Bank of Baroda, Union Bank, Canara Bank, IDFC First Bank, and more.
                                                                                                                                              
                                                                                                                                              ## Files
                                                                                                                                              
                                                                                                                                              | File | Description |
                                                                                                                                              |---|---|
                                                                                                                                              | `universal_bank_parser_v13.py` | Main PDF parser (4-tier pipeline) |
                                                                                                                                              | `xlsx_bank_parser.py` | XLSX/XLS parser (same output schema) |
                                                                                                                                              | `adapters.py` | Thin routing adapter (PDF vs XLSX) |
                                                                                                                                              | `gst_intelligence.py` | GST taxpayer data fetcher and scorer |
                                                                                                                                              | `emi_bounce_master_dictionary.md` | Curated EMI/bounce keyword rulebook |
                                                                                                                                              | `requirements.txt` | Python dependencies |
                                                                                                                                              | `.env.example` | Environment variable template |
                                                                                                                                              
                                                                                                                                              ## License
                                                                                                                                              
                                                                                                                                              MIT
