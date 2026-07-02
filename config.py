# -*- coding: utf-8 -*-
from __future__ import annotations
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any
import yaml

@dataclass(frozen=True)
class SentTypeRule:
    column_index: int | None = None
    column_name: str | None = None
    contains: str = "EXPRO"
    true_type: str = "SENT_200"
    false_type: str = "SENT_205"

@dataclass(frozen=True)
class ClientConfig:
    name: str
    excel_file: str
    sent_folder: str
    default_sent_type: str = "SENT_205"
    sent_type_rule: SentTypeRule | None = None
    mappings: dict[str, str] = field(default_factory=dict)
    document_id_base: str | None = None
    document_id_registration_column: str | None = None
    document_id_use_split_truck: bool = False
    registration_column: str = "AUTO"
    registration_fallback_columns: tuple[str, ...] = ()
    registration_use_split_truck: bool = False
    outlook_document_ids: tuple[str, ...] = ()
    carrier_registration_split: bool = False
    sender_source: str = "mapping"
    sender_name_cell: str = "A1"
    sender_identity_cell: str = "A2"
    sender_address_cell: str = "A3"

    @staticmethod
    def from_dict(raw: dict[str, Any]) -> "ClientConfig":
        rule_raw = raw.get("sent_type_rule")
        rule = SentTypeRule(**rule_raw) if isinstance(rule_raw, dict) else None
        return ClientConfig(
            name=str(raw["name"]).upper(),
            excel_file=str(raw["excel_file"]),
            sent_folder=str(raw.get("sent_folder") or f"SENTY {raw['name']}").upper(),
            default_sent_type=str(raw.get("default_sent_type", "SENT_205")),
            sent_type_rule=rule,
            mappings={str(k): str(v) for k, v in (raw.get("mappings") or {}).items()},
            document_id_base=raw.get("document_id_base"),
            document_id_registration_column=raw.get("document_id_registration_column"),
            document_id_use_split_truck=bool(raw.get("document_id_use_split_truck", False)),
            registration_column=str(raw.get("registration_column", "AUTO")),
            registration_fallback_columns=tuple(raw.get("registration_fallback_columns") or ()),
            registration_use_split_truck=bool(raw.get("registration_use_split_truck", False)),
            outlook_document_ids=tuple(str(x) for x in raw.get("outlook_document_ids", ())),
            carrier_registration_split=bool(raw.get("carrier_registration_split", False)),
            sender_source=str(raw.get("sender_source", "mapping")),
            sender_name_cell=str(raw.get("sender_name_cell", "A1")),
            sender_identity_cell=str(raw.get("sender_identity_cell", "A2")),
            sender_address_cell=str(raw.get("sender_address_cell", "A3")),
        )

    def as_yaml_dict(self) -> dict[str, Any]:
        data = {
            "name": self.name,
            "excel_file": self.excel_file,
            "sent_folder": self.sent_folder,
            "default_sent_type": self.default_sent_type,
            "mappings": dict(self.mappings),
            "registration_column": self.registration_column,
        }
        if self.registration_fallback_columns:
            data["registration_fallback_columns"] = list(self.registration_fallback_columns)
        if self.registration_use_split_truck:
            data["registration_use_split_truck"] = True
        if self.carrier_registration_split:
            data["carrier_registration_split"] = True
        if self.sender_source != "mapping":
            data["sender_source"] = self.sender_source
            data["sender_name_cell"] = self.sender_name_cell
            data["sender_identity_cell"] = self.sender_identity_cell
            data["sender_address_cell"] = self.sender_address_cell
        if self.document_id_base:
            data["document_id_base"] = self.document_id_base
        if self.document_id_registration_column:
            data["document_id_registration_column"] = self.document_id_registration_column
        if self.document_id_use_split_truck:
            data["document_id_use_split_truck"] = True
        if self.outlook_document_ids:
            data["outlook_document_ids"] = list(self.outlook_document_ids)
        if self.sent_type_rule:
            rule = asdict(self.sent_type_rule)
            data["sent_type_rule"] = {k: v for k, v in rule.items() if v is not None}
        return data

class ClientRegistry:
    def __init__(self, clients: list[ClientConfig], config_path: Path):
        self.config_path = config_path
        self._clients = {c.name.upper(): c for c in clients}

    @classmethod
    def load(cls, base_dir: Path) -> "ClientRegistry":
        path = base_dir / "clients.yaml"
        if not path.exists():
            raise FileNotFoundError(f"Nie znaleziono clients.yaml: {path}")
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        return cls([ClientConfig.from_dict(x) for x in raw.get("clients", [])], path)

    def names(self) -> list[str]:
        return sorted(self._clients)

    def get(self, name: str) -> ClientConfig:
        key = str(name).upper()
        if key not in self._clients:
            raise KeyError(f"Nie ma klienta w konfiguracji: {name}")
        return self._clients[key]

    def add_client(self, client: ClientConfig) -> None:
        if client.name.upper() in self._clients:
            raise ValueError(f"Klient już istnieje: {client.name}")
        self._clients[client.name.upper()] = client
        self.save()

    def remove_client(self, name: str) -> None:
        key = str(name).upper()
        if key not in self._clients:
            raise KeyError(f"Nie ma klienta w konfiguracji: {name}")
        if len(self._clients) <= 1:
            raise ValueError("Nie można usunąć ostatniego klienta z konfiguracji.")
        del self._clients[key]
        self.save()

    def provider_from_document_id(self, document_id: str, normalizer) -> str:
        normalized = normalizer(document_id)
        for client in self._clients.values():
            for doc_id in client.outlook_document_ids:
                if normalized.startswith(normalizer(doc_id)):
                    return client.name
        return ""

    def save(self) -> None:
        payload = {"clients": [self._clients[k].as_yaml_dict() for k in self.names()]}
        self.config_path.write_text(yaml.safe_dump(payload, allow_unicode=True, sort_keys=False), encoding="utf-8")
