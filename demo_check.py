# -*- coding: utf-8 -*-
"""Smoke test for the anonymized GitHub demo package.

Run from repository root:
    python demo_check.py
"""
from __future__ import annotations

from pathlib import Path

from config import ClientRegistry
from excel_service import detect_sent_type, find_row_by_lp, load_excel
from utils import app_dir, get_registration_from_row, safe_file_token
from xml_builder import create_xml


def main() -> None:
    base_dir = app_dir()
    registry = ClientRegistry.load(base_dir)
    client = registry.get("DEMO_CLIENT")

    df = load_excel(base_dir, client)
    row = find_row_by_lp(df, "1")
    sent_type = detect_sent_type(row, client)

    xml_tree = create_xml(base_dir, row, client, sent_type)
    reg = get_registration_from_row(
        row,
        client.registration_column,
        client.registration_fallback_columns,
        client.registration_use_split_truck,
    )

    outdir = base_dir / "output"
    outdir.mkdir(exist_ok=True)
    outpath = outdir / f"{sent_type}_{client.name}_LP1_{safe_file_token(reg)}.xml"
    xml_tree.write(str(outpath), encoding="utf-8", xml_declaration=True, pretty_print=True)

    print("OK: demo package is runnable")
    print(f"Client: {client.name}")
    print(f"SENT type: {sent_type}")
    print(f"Generated: {outpath.relative_to(base_dir)}")


if __name__ == "__main__":
    main()
