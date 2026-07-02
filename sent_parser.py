# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from lxml import etree
from lxml import html as lxml_html
from utils import clean_html_text
import re

def value_after_phrase_id(doc, phrase_ids):
    if isinstance(phrase_ids, str):
        phrase_ids = [phrase_ids]
    for phrase_id in phrase_ids:
        nodes = doc.xpath(f'//*[@data-phrase-id="{phrase_id}"]')
        for node in nodes:
            parent = node.getparent()
            if parent is not None:
                value = clean_html_text(parent.text_content())
                if value:
                    return value
            value = clean_html_text(node.tail)
            if value:
                return value
    return ""

def value_from_table_column_by_phrase_id(doc, phrase_id):
    headers = doc.xpath(f"//*[@data-phrase-id='{phrase_id}']")
    for header in headers:
        th = header if header.tag.lower() in ("td", "th") else header.getparent()
        if th is None:
            continue
        row = th.getparent()
        if row is None:
            continue
        cells = row.xpath("./td|./th")
        try:
            col_idx = cells.index(th)
        except ValueError:
            continue
        next_row = row.xpath("./following-sibling::tr[1]")
        if not next_row:
            continue
        value_cells = next_row[0].xpath("./td|./th")
        if len(value_cells) > col_idx:
            return clean_html_text(" ".join(value_cells[col_idx].xpath(".//text()")))
    return ""

def parse_sent_html_file(html_path: Path):
    parser = lxml_html.HTMLParser(encoding="utf-8")
    doc = lxml_html.parse(str(html_path), parser=parser)
    body_text = clean_html_text(" ".join(doc.xpath("//body//text()")))
    sent_raw = value_after_phrase_id(doc, ["sentMultipleNumber", "sentNumber"]) or body_text
    m_sent = re.search(r"\bSENT(?:ZB)?\d+\b", sent_raw, re.IGNORECASE)
    sent_number = m_sent.group(0).upper() if m_sent else ""
    key_raw = value_after_phrase_id(doc, "sentRecipientKey")
    m_key = re.search(r"\bKR-[A-Z0-9]+\b", key_raw, re.IGNORECASE)
    recipient_key = m_key.group(0).upper() if m_key else clean_html_text(key_raw)
    own_number_raw = value_after_phrase_id(doc, ["sentSourceDocumentID", "sentSelfNumber"])
    cmr_number = clean_html_text(value_after_phrase_id(doc, "sentNumberOfTransportDocument"))
    m_lp = re.search(r"\b(\d+)\s*/\s*\d+\b", own_number_raw) or re.search(r"\b(\d+)\s*/\s*\d+\b", cmr_number)
    lp = m_lp.group(1) if m_lp else ""
    weight_raw = value_after_phrase_id(doc, ["sentTotalGrossWeightOfGoods", "sentGrossWeightOfGoods"])
    if not weight_raw:
        weight_raw = value_from_table_column_by_phrase_id(doc, "sentGrossWeightOfGoods")
    weight_value = ""
    m_weight = re.search(r"(\d+(?:[,.]\d+)?)\s*(kg|t)?\b", weight_raw.replace(" ", ""), re.IGNORECASE)
    if m_weight:
        numeric = float(m_weight.group(1).replace(",", "."))
        unit = (m_weight.group(2) or "").lower()
        if unit == "kg" or numeric > 1000:
            numeric = numeric / 1000
        weight_value = round(numeric, 3)
    return {"file": html_path.name, "lp": lp, "sent": sent_number, "recipient_key": recipient_key, "cmr": cmr_number or clean_html_text(own_number_raw), "weight": weight_value, "own_number": clean_html_text(own_number_raw)}

def parse_sent_xml_file(xml_path: Path):
    sent_match = re.search(r"(SENT\d+)", xml_path.name, re.IGNORECASE)
    sent_number = sent_match.group(1).upper() if sent_match else ""
    tree = etree.parse(str(xml_path))
    root = tree.getroot()
    def first_text_by_localname(names):
        if isinstance(names, str):
            names = [names]
        for name in names:
            for node in root.xpath(f".//*[local-name()='{name}']"):
                value = clean_html_text(node.text)
                if value:
                    return value
        return ""
    recipient_key = first_text_by_localname(["RecipientKey", "GoodsRecipientKey", "SentRecipientKey"])
    cmr_number = first_text_by_localname(["NumberOfTransportDocument", "TransportDocumentNumber"])
    own_number = first_text_by_localname(["DocumentId", "SourceDocumentID", "SelfNumber"])
    lp = ""
    for raw in [own_number, cmr_number]:
        m_lp = re.search(r"\b(\d+)\s*/\s*\d+\b", str(raw or ""))
        if m_lp:
            lp = m_lp.group(1)
            break
    weight_value = ""
    weight_raw = first_text_by_localname(["AmountOfGoods", "GrossWeightOfGoods", "TotalGrossWeightOfGoods"])
    if weight_raw:
        try:
            numeric = float(str(weight_raw).replace(",", "."))
            if numeric > 1000:
                numeric = numeric / 1000
            weight_value = round(numeric, 3)
        except Exception:
            pass
    return {"file": xml_path.name, "lp": lp, "sent": sent_number, "recipient_key": recipient_key, "cmr": cmr_number or clean_html_text(own_number), "weight": weight_value, "own_number": clean_html_text(own_number)}
