# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import shutil
import yaml
from config import ClientConfig, SentTypeRule
from excel_service import get_headers_from_excel


def const_or_todo(value: str, todo_column: str) -> str:
    """Zwraca stałą YAML (=wartość) albo nazwę kolumny TODO do ręcznego mapowania."""
    value = str(value or "").strip()
    return f"={value}" if value else todo_column


def pick_trailer_column(headers: list[str], reg_col: str) -> str:
    """Dobiera kolumnę naczepy/wagonu na podstawie nagłówków Excela."""
    normalized = {str(h).strip().upper(): h for h in headers}
    candidates = [
        "TRAILER",
        "PRZYCZEPA",
        "NACZEPA",
        "NACZEPA / WAGON",
        "NACZEPA/WAGON",
        "TRAILER/WAGON",
        "TRAILER / WAGON",
        "WAGON",
    ]
    for c in candidates:
        if c in normalized:
            return normalized[c]

    # W starym układzie AUTO/ WAGON bywa rozbijane z jednej komórki, np. AUTO/PRZYCZEPA.
    # Dla Excela typu AUTO bez oddzielnej kolumny zostawiamy AUTO jako awaryjny fallback.
    if reg_col == "AUTO/ WAGON":
        return "PRZYCZEPA" if "PRZYCZEPA" in headers else reg_col
    return reg_col


def build_base_mapping_200(sender: dict[str, str] | None = None) -> dict[str, str]:
    sender = sender or {}
    return {
        "TypeOfTransport": "=1",
        "GoodsRecipient/TraderInfo/TraderName": "=EXAMPLE RECIPIENT SP. Z O.O.",
        "GoodsRecipient/TraderInfo/TraderIdentityType": "=NIP",
        "GoodsRecipient/TraderInfo/TraderIdentityNumber": "=0000000000",
        "GoodsRecipient/TraderAddress/Street": "=Example Street",
        "GoodsRecipient/TraderAddress/HouseNumber": "=1",
        "GoodsRecipient/TraderAddress/City": "=Example City",
        "GoodsRecipient/TraderAddress/Country": "=PL",
        "GoodsRecipient/TraderAddress/PostalCode": "=00-000",
        "Transport/PlaceOfDelivery/Street": "=Example Street",
        "Transport/PlaceOfDelivery/HouseNumber": "=1",
        "Transport/PlaceOfDelivery/City": "=Example City",
        "Transport/PlaceOfDelivery/Country": "=PL",
        "Transport/PlaceOfDelivery/PostalCode": "=00-000",
        "Transport/PlaceOfDelivery/CodeTERC": "=00",
        "Transport/PlaceOfDelivery/Latitude": "=52.000000",
        "Transport/PlaceOfDelivery/Longitude": "=21.000000",
        "GoodsInformation/ElementNumber": "=1",
        "GoodsInformation/CodeCnClassification": "=0005",
        "GoodsInformation/GoodsName": "=EXAMPLE GOODS",
        "GoodsInformation/AmountOfGoods": "POTWIERDZONA MASA",
        "GoodsInformation/UnitOfMeasure": "=kg",
        "GoodsInformation/VATRate": "=23",
        "GoodsInformation/WasteCode": "=191207",
        "DocumentId": "AUTO",
        "ResponseAddress/EmailChannel/EmailAddress1": "=notifications@example.com",
        "ResponseAddress/EmailChannel/EmailAddress2": "TRÄGER",
        "Statements/Statement1": "=true",
        "Statements/FirstName": "=JANE",
        "Statements/LastName": "=DOE",
    }


def build_mapping_205(reg_col: str, trailer_col: str, sender: dict[str, str] | None = None) -> dict[str, str]:
    m = build_base_mapping_200(sender)
    m.update({
        "Carrier/TraderInfo/TraderName": "TRÄGER",
        "Carrier/TraderInfo/TraderIdentityType": "=NIP",
        "Carrier/TraderInfo/TraderIdentityNumber": "=0000000000",
        "Carrier/TraderAddress/Street": "",
        "Carrier/TraderAddress/HouseNumber": "",
        "Carrier/TraderAddress/City": "",
        "Carrier/TraderAddress/Country": "",
        "Carrier/TraderAddress/PostalCode": "",
        "MeansOfTransport/TruckOrTrainNumber": reg_col,
        "MeansOfTransport/TrailerOrWagonNumber": trailer_col,
        "MeansOfTransport/PermitRoad": "NUMER ZEZWOLENIA DROGOWEGO (LICENCJA)",
        "MeansOfTransport/GeoLocatorNumber": "NADAJNIK GPS",
        "Transport/StartTransportDate": "DATUM",
        "Transport/EndTransportDate": "DATUM",
        "Transport/EntranceToPoland/RoutePlace": "=EXAMPLE_BORDER_CROSSING",
        "GoodsTransportDocuments/TypeOfTransportDocument": "=INNY",
        "GoodsTransportDocuments/NumberOfTransportDocument": "NR CMR PRZYPISANY DO SENT",
        "DocumentId": reg_col,
    })
    return m


def save_yaml(path: Path, data: dict):
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")


def create_client(
    base_dir: Path,
    registry,
    name: str,
    excel_source: Path,
    excel_type: str,
    document_id: str = "",
    sent_rule_index: int | None = 7,
    sender_name_cell: str = "A1",
    sender_identity_cell: str = "A2",
    sender_address_cell: str = "A3",
) -> ClientConfig:
    name = name.strip().upper()
    if not name:
        raise ValueError("Podaj nazwę klienta.")
    if not excel_source.exists():
        raise FileNotFoundError(f"Nie znaleziono Excela: {excel_source}")

    excel_type = excel_type.strip().upper()
    reg_col = "AUTO/ WAGON" if "WAGON" in excel_type else "AUTO"
    headers = get_headers_from_excel(excel_source)
    if reg_col not in headers:
        raise ValueError(f"Wybrano typ {reg_col}, ale Excel nie ma takiej kolumny. Dostępne: {', '.join(headers)}")

    trailer_col = pick_trailer_column(headers, reg_col)

    dest_excel = base_dir / "excels" / f"AWIZACJA {name}.xlsx"
    dest_excel.parent.mkdir(exist_ok=True)
    if excel_source.resolve() != dest_excel.resolve():
        shutil.copy2(excel_source, dest_excel)

    safe_name = name.replace(" ", "")
    map200 = f"mapping_{safe_name}_200.yaml"
    map205 = f"mapping_{safe_name}_205.yaml"
    mappings_dir = base_dir / "mappings"
    mappings_dir.mkdir(exist_ok=True)

    save_yaml(mappings_dir / map200, build_base_mapping_200() | {"DocumentId": reg_col})
    save_yaml(mappings_dir / map205, build_mapping_205(reg_col, trailer_col))

    rule = SentTypeRule(column_index=sent_rule_index, contains="EXPRO", true_type="SENT_200", false_type="SENT_205") if sent_rule_index is not None else None
    client = ClientConfig(
        name=name,
        excel_file=dest_excel.name,
        sent_folder=f"SENTY {name}",
        outlook_document_ids=(document_id,) if document_id else (),
        document_id_base=document_id or None,
        document_id_registration_column=reg_col,
        document_id_use_split_truck=reg_col == "AUTO/ WAGON",
        registration_column=reg_col,
        registration_use_split_truck=reg_col == "AUTO/ WAGON",
        default_sent_type="SENT_205",
        sent_type_rule=rule,
        mappings={"SENT_200": map200, "SENT_205": map205},
        sender_source="excel_top_cells",
        sender_name_cell=sender_name_cell or "A1",
        sender_identity_cell=sender_identity_cell or "A2",
        sender_address_cell=sender_address_cell or "A3",
    )
    registry.add_client(client)
    (base_dir / client.sent_folder).mkdir(exist_ok=True)
    return client
