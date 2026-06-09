"""
Adapter layer – routes bank statement files to the appropriate parser.

Usage:
    from adapters import run_bank_statement_parser, run_gst_lookup

    # Auto-routes PDF -> universal_bank_parser_v13
    # Auto-routes XLSX/XLS -> xlsx_bank_parser
    result = run_bank_statement_parser("statement.pdf")
    result = run_bank_statement_parser("statement.xlsx")

    # GST intelligence lookup
    gst_data = run_gst_lookup("29ABCDE1234F1Z5")
"""

from pathlib import Path

from gst_intelligence import fetch_gst_intelligence
from universal_bank_parser_v13 import parse_bank_statement
from xlsx_bank_parser import parse_xlsx_bank_statement

_XLSX_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def run_gst_lookup(gstin: str) -> dict:
    """Fetch and score GST taxpayer intelligence for the given GSTIN."""
    return fetch_gst_intelligence(gstin)


def run_bank_statement_parser(
    file_path: str,
    password: str | None = None,
    **kwargs,
) -> dict:
    """
    Route to the correct parser based on file extension.

    Args:
        file_path: Path to the bank statement (PDF or Excel).
        password:  PDF password if the file is password-protected.
        **kwargs:  Additional arguments forwarded to the underlying parser.

    Returns:
        Parsed result dict (same schema for both PDF and XLSX).
    """
    ext = Path(file_path).suffix.lower()
    if ext in _XLSX_EXTENSIONS:
        return parse_xlsx_bank_statement(
            file_path=file_path,
            password=password,
            save_json=False,
            output_dir=str(Path(file_path).parent),
        )
    # Default: PDF parser
    return parse_bank_statement(
        pdf_path=file_path,
        password=password,
        save_json=False,
        output_dir=str(Path(file_path).parent),
    )
