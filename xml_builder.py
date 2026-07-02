# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from datetime import timedelta
import re
import pandas as pd
import yaml
from lxml import etree
from utils import parse_registration, get_registration_from_row
from excel_service import read_sender_from_excel

TP_NS = "http://www.mf.gov.pl/SENT/2017/12/08/STypes.xsd"
TP_ELEMENTS = {
    "TraderInfo", "TraderAddress", "TraderName", "TraderIdentityType", "TraderIdentityNumber",
    "Street", "HouseNumber", "FlatNumber", "City", "Country", "PostalCode",
    "TruckOrTrainNumber", "TrailerOrWagonNumber", "PermitRoad", "GeoLocatorNumber",
    "FailoverGeoLocatorNumber", "FailoverCarrierEmail", "NotificationSMSPhone", "DriverKey",
    "PlannedStartCarriageDate", "StartTransportDate", "EndTransportDate",
    "PlaceOfLoading", "PlaceOfDelivery", "ExitFromPoland", "EntranceToPoland", "RoutePlace", "RouteNumber",
    "CodeTERC", "Latitude", "Longitude", "ElementNumber", "CodeCnClassification", "GoodsName",
    "AmountOfGoods", "UnitOfMeasure", "VATRate", "WasteCode", "GoodsCoveredByContract",
    "TypeOfTransportDocument", "NumberOfTransportDocument", "EmailChannel", "EmailAddress1", "EmailAddress2",
    "EmailAddress3", "WebServiceChannel", "WsFromSISC", "Statement1", "FirstName", "LastName",
    "Statement2", "Statement3", "Statement6", "Statement7",
}

def load_mapping(path: Path) -> dict[str, str]:
    if not path.exists():
        raise FileNotFoundError(f"Plik mapowania nie istnieje: {path}")
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return {str(k).strip().strip("/"): v for k, v in data.items()}

def detect_root_and_ns(xsd_path: Path) -> tuple[str, str]:
    tree = etree.parse(str(xsd_path))
    ns = tree.getroot().get("targetNamespace")
    root = tree.xpath("//xs:schema/xs:element[1]", namespaces={"xs": "http://www.w3.org/2001/XMLSchema"})
    return (root[0].get("name") if root else xsd_path.stem), ns

def ns(root):
    return {"ns": root.nsmap.get(None), "tp": root.nsmap.get("tp") or TP_NS}

def ensure_path(root: etree._Element, path: str) -> etree._Element:
    ns_sent = root.nsmap.get(None)
    parts = [p for p in path.split("/") if p]
    cur = root
    for p in parts:
        if "[" in p and p.endswith("]"):
            name, idx_str = p[:-1].split("[")
            idx = int(idx_str)
        else:
            name, idx = p, None
        matches = [ch for ch in cur if etree.QName(ch).localname == name]
        target = matches[0] if idx is None and matches else None
        if idx is not None and len(matches) >= idx:
            target = matches[idx - 1]
        if target is None:
            uri = ns_sent if name == "DocumentId" or name not in TP_ELEMENTS else TP_NS
            target = etree.Element(etree.QName(uri, name), nsmap=root.nsmap)
            cur.append(target)
        cur = target
    return cur

def reorder_trader_addresses(root):
    n = ns(root)
    for addr in root.xpath(".//tp:TraderAddress", namespaces=n):
        street_el = addr.find(f"{{{TP_NS}}}Street")
        house_el = addr.find(f"{{{TP_NS}}}HouseNumber")
        if street_el is not None and house_el is None:
            street_txt = (street_el.text or "").strip()
            m = re.match(r"^(.*\D)\s+(\d+[A-Za-z]?)\s*$", street_txt)
            if m:
                street_el.text = m.group(1).strip()
                house_el = etree.Element(f"{{{TP_NS}}}HouseNumber")
                house_el.text = m.group(2).strip()
                addr.insert(list(addr).index(street_el) + 1, house_el)
        order = ["Street", "HouseNumber", "FlatNumber", "City", "Country", "PostalCode"]
        addr[:] = sorted(list(addr), key=lambda el: order.index(etree.QName(el).localname) if etree.QName(el).localname in order else 99)

def fix_root_order(root, root_name: str):
    if root_name in ("SENT_105", "SENT_100"):
        order = ["TypeOfTransport", "TypeOfDeclaration", "GoodsSender", "Carrier", "GoodsRecipient", "MeansOfTransport", "Transport", "GoodsInformation", "PurposeOfTheGoods", "GoodsTransportDocuments", "Comments", "DocumentId", "ResponseAddress", "Statements", "AdditionalStatements"]
    else:
        order = ["TypeOfTransport", "TypeOfDeclaration", "GoodsRecipient", "Carrier", "GoodsSender", "MeansOfTransport", "Transport", "GoodsInformation", "GoodsTransportDocuments", "Comments", "DocumentId", "ResponseAddress", "Statements", "AdditionalStatements"]
    existing = {name: root.find(f"ns:{name}", namespaces=ns(root)) for name in order}
    for child in list(root):
        root.remove(child)
    for name in order:
        if existing.get(name) is not None:
            root.append(existing[name])

def parse_carrier_data(cell_value: str) -> dict[str, str]:
    if not isinstance(cell_value, str) or not cell_value.strip():
        return {}
    lines = [ln.strip() for ln in cell_value.strip().replace("\r", "\n").split("\n") if ln.strip()]
    country_map = {"Polska": "PL", "Poland": "PL", "Norge": "NO", "Norway": "NO", "Deutschland": "DE", "Germany": "DE", "": "PL"}
    data = {"Carrier/TraderInfo/TraderName": "", "Carrier/TraderInfo/TraderIdentityType": "", "Carrier/TraderInfo/TraderIdentityNumber": "", "Carrier/TraderAddress/Street": "", "Carrier/TraderAddress/HouseNumber": "", "Carrier/TraderAddress/PostalCode": "", "Carrier/TraderAddress/City": "", "Carrier/TraderAddress/Country": "PL", "ResponseAddress/EmailChannel/EmailAddress2": ""}
    for line in lines:
        lower = line.lower()
        if lower.startswith("name:"):
            data["Carrier/TraderInfo/TraderName"] = line.split(":", 1)[1].strip()
        elif lower.startswith("registrienummer:"):
            nip_raw = re.sub(r"[^A-Z0-9]", "", line.split(":", 1)[1].strip().upper())
            data["Carrier/TraderInfo/TraderIdentityType"] = "NIP"
            data["Carrier/TraderInfo/TraderIdentityNumber"] = re.sub(r"^[A-Z]{2}", "", nip_raw)
        elif lower.startswith("anschrift:"):
            parts = [p.strip() for p in line.split(":", 1)[1].strip().split(",") if p.strip()]
            country_raw = parts[-1] if parts else ""
            country = country_map.get(country_raw, country_raw[:2].upper() or "PL")
            street = house = postal = city = ""
            if parts:
                m = re.match(r"^(.*?)(\d+[A-Za-z]?)?$", parts[0])
                if m:
                    street = (m.group(1) or "").strip(" ,")
                    house = (m.group(2) or "").strip()
            if len(parts) >= 2:
                m2 = re.match(r"^(\d{2}-\d{3})?\s*(.*)$", parts[1])
                if m2:
                    postal = (m2.group(1) or "").strip()
                    city = (m2.group(2) or "").strip()
            data.update({"Carrier/TraderAddress/Street": street, "Carrier/TraderAddress/HouseNumber": house, "Carrier/TraderAddress/PostalCode": postal, "Carrier/TraderAddress/City": city, "Carrier/TraderAddress/Country": country})
        elif "mail" in lower:
            email_match = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", line)
            if email_match:
                data["ResponseAddress/EmailChannel/EmailAddress2"] = email_match.group(0).strip()
    return data

def resolve_mapping_value(row, xml_path: str, col):
    if isinstance(col, str) and col.startswith("="):
        return col[1:]
    return row.get(col, "")

def build_document_id(row, client) -> str | None:
    if not client.document_id_base:
        return None
    col = client.document_id_registration_column or client.registration_column
    reg = get_registration_from_row(row, col, client.registration_fallback_columns, client.document_id_use_split_truck)
    if reg:
        return f"{client.document_id_base} {reg.replace(' ', '').upper()}"
    return client.document_id_base

def build_xml_from_row(row, mapping, schema, target_ns, root_name, client, sender_overrides=None):
    root = etree.Element(etree.QName(target_ns, root_name), nsmap={None: target_ns, "tp": TP_NS})
    sender_overrides = sender_overrides or {}
    for xml_path, col in mapping.items():
        value = resolve_mapping_value(row, xml_path, col)
        if pd.isna(value) or str(value).strip() == "":
            continue
        el = ensure_path(root, xml_path)
        if xml_path.startswith("Carrier/"):
            carrier_col = mapping.get("Carrier/TraderInfo/TraderName")
            if carrier_col and carrier_col in row.index:
                for k, v in parse_carrier_data(str(row[carrier_col])).items():
                    if v:
                        ensure_path(root, k).text = str(v)
                continue
        if xml_path.endswith(("/RegistrationNumber", "/TruckOrTrainNumber", "/TrailerOrWagonNumber")) and isinstance(value, str):
            value = value.replace(" ", "").strip().upper()
        if xml_path.endswith("/TraderIdentityNumber") and isinstance(value, str) and value.upper().startswith("PL"):
            value = value[2:].strip()
        if xml_path.endswith("/Statement1"):
            ensure_path(root, "Statements/Statement1").text = "true"
            continue
        if xml_path.endswith("EndTransportDate"):
            try:
                base = pd.to_datetime(str(row["DATUM"]).strip(), dayfirst=True)
                value = (base + timedelta(days=1)).date().isoformat()
            except Exception:
                value = str(row.get("DATUM", ""))
        if client.registration_use_split_truck and xml_path == "MeansOfTransport/TruckOrTrainNumber":
            raw = get_registration_from_row(row, client.registration_column, client.registration_fallback_columns, False)
            value = parse_registration(raw)[0]
        if client.registration_use_split_truck and xml_path == "MeansOfTransport/TrailerOrWagonNumber":
            raw = get_registration_from_row(row, client.registration_column, client.registration_fallback_columns, False)
            value = parse_registration(raw)[1]
        if xml_path.endswith("/AmountOfGoods"):
            try:
                value = str(round(float(str(value).replace(",", ".")) * 1000, 3))
            except Exception:
                pass
        if xml_path == "DocumentId":
            dynamic = build_document_id(row, client)
            if dynamic:
                value = dynamic
        if "ResponseAddress" in xml_path:
            email_match = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", str(value))
            value = email_match.group(0).strip() if email_match else ""
            if not value:
                continue
        if xml_path.endswith(("PlannedStartCarriageDate", "StartTransportDate")):
            try:
                value = pd.to_datetime(str(value), dayfirst=True).date().isoformat()
            except Exception:
                pass
        el.text = str(value)
    # Dane nadawcy z Excela maja pierwszenstwo nad mappingiem.
    for sender_path, sender_value in sender_overrides.items():
        if sender_value is not None and str(sender_value).strip():
            ensure_path(root, sender_path).text = str(sender_value).strip()

    tree = etree.ElementTree(root)
    reorder_trader_addresses(root)
    fix_root_order(root, root_name)
    try:
        schema.assertValid(tree)
    except etree.DocumentInvalid as e:
        log_msg = str(e.error_log)
        if "EmailAddress" in log_msg and "pattern" in log_msg:
            pass
        else:
            raise
    return tree

def create_xml(base_dir: Path, row, client, sent_type: str):
    xsd_path = base_dir / "schemas" / f"{sent_type}.xsd"
    if not xsd_path.exists():
        raise FileNotFoundError(f"Nie znaleziono schematu: {xsd_path}")
    mapping_name = client.mappings.get(sent_type)
    if not mapping_name:
        raise FileNotFoundError(f"Brak mapowania dla {client.name} / {sent_type}")
    mapping_path = base_dir / "mappings" / mapping_name
    parser = etree.XMLParser(load_dtd=True, no_network=False)
    schema = etree.XMLSchema(etree.parse(str(xsd_path), parser))
    root_name, target_ns = detect_root_and_ns(xsd_path)
    mapping = load_mapping(mapping_path)
    sender_overrides = read_sender_from_excel(base_dir, client)
    return build_xml_from_row(row, mapping, schema, target_ns, root_name, client, sender_overrides)
