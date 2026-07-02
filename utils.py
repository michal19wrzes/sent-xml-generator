# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import html as html_lib
import re
import sys

def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent

def resource_path(*parts: str) -> Path:
    return app_dir().joinpath(*parts)

def clean_html_text(text) -> str:
    if text is None:
        return ""
    text = html_lib.unescape(str(text)).replace("\xa0", " ")
    return " ".join(text.split()).strip()

def normalize_excel_key(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return " ".join(text.split()).upper()

def normalize_document_id(value) -> str:
    return " ".join(str(value or "").upper().split())

def safe_file_token(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper().replace(" ", "")) or "BRAKREJ"

def parse_registration(value: str) -> tuple[str, str]:
    if not value or not isinstance(value, str):
        return "", ""
    cleaned = value.strip().replace("  ", " ").replace(" /", "/").replace("/ ", "/")
    if "/" in cleaned:
        truck, trailer = cleaned.split("/", 1)
        return truck.strip().upper(), trailer.strip().upper()
    return cleaned.strip().upper(), ""

def get_registration_from_row(row, primary: str, fallback_columns=(), split_truck=False) -> str:
    raw = str(row.get(primary, "") or "").strip()
    for col in fallback_columns or ():
        if raw:
            break
        raw = str(row.get(col, "") or "").strip()
    if split_truck:
        return parse_registration(raw)[0]
    return raw.strip().upper()
