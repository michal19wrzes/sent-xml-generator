# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
from datetime import datetime, timedelta
import re
import os
from utils import normalize_document_id

OUTLOOK_CLIENT_EMAIL = os.getenv("OUTLOOK_CLIENT_EMAIL", "sent@example.gov")
OUTLOOK_PROCESSED_CATEGORY = os.getenv("OUTLOOK_PROCESSED_CATEGORY", "PROCESSED_XML")
OUTLOOK_LOOKBACK_DAYS = 14

def extract_document_id_from_subject(subject: str) -> str:
    match = re.search(r"Id\s+dokumentu\s*:\s*(.+)$", str(subject or ""), re.IGNORECASE)
    if not match:
        return ""
    return normalize_document_id(re.split(r"[,;]", match.group(1).strip(), maxsplit=1)[0].strip())

def extract_sent_filename_from_subject(subject: str) -> str:
    text = str(subject or "")
    sent_number_match = re.search(r"\b(SENT\d+)\b", text, re.IGNORECASE)
    sent_type_match = re.search(r"\[(SENT)_?(\d+)\]", text, re.IGNORECASE)
    if not sent_number_match or not sent_type_match:
        return ""
    sent_number = sent_number_match.group(1).upper()
    sent_type = f"{sent_type_match.group(1).upper()}{sent_type_match.group(2)}"
    return f"{sent_number}_{sent_type}.xml"

def get_outlook_sender_smtp(mail):
    try:
        sender_email_type = str(getattr(mail, "SenderEmailType", "") or "").upper()
        sender_email = str(getattr(mail, "SenderEmailAddress", "") or "")
        if sender_email_type == "EX":
            sender = getattr(mail, "Sender", None)
            if sender is not None:
                for method in ("GetExchangeUser", "GetExchangeDistributionList"):
                    try:
                        obj = getattr(sender, method)()
                        smtp = str(getattr(obj, "PrimarySmtpAddress", "") or "") if obj is not None else ""
                        if smtp:
                            return smtp.lower()
                    except Exception:
                        pass
        return sender_email.lower()
    except Exception:
        return str(getattr(mail, "SenderEmailAddress", "") or "").lower()

def transfer_xml_from_outlook(base_dir: Path, registry, log=lambda m: None) -> str:
    try:
        import win32com.client
    except ImportError:
        raise RuntimeError("Brak biblioteki pywin32. Zainstaluj ją komendą: pip install pywin32")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    mailbox = None
    for store in namespace.Folders:
        name = str(store.Name).strip().upper()
        mailbox_keywords = [x.strip().upper() for x in os.getenv("OUTLOOK_MAILBOX_KEYWORDS", "EXAMPLE,MAILBOX").split(",") if x.strip()]
        if all(keyword in name for keyword in mailbox_keywords):
            mailbox = store
            break
    if mailbox is None:
        raise RuntimeError("Nie znaleziono skrzynki Outlook zgodnej z OUTLOOK_MAILBOX_KEYWORDS.")
    try:
        inbox = mailbox.Folders["Skrzynka odbiorcza"]
    except Exception:
        inbox = mailbox.Folders["Inbox"]
    try:
        folder = inbox.Folders["SENT"]
    except Exception:
        raise RuntimeError("Nie znaleziono folderu SENT w skrzynce odbiorczej.")
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    since = datetime.now() - timedelta(days=OUTLOOK_LOOKBACK_DAYS)
    saved = checked = matched = processed = xml_seen = skipped_processed = skipped_no_mapping = skipped_no_subject = 0
    for mail in items:
        try:
            if getattr(mail, "Class", None) != 43:
                continue
            received_time = getattr(mail, "ReceivedTime", None)
            if received_time is not None:
                try:
                    if received_time.replace(tzinfo=None) < since:
                        break
                except Exception:
                    pass
            checked += 1
            sender = get_outlook_sender_smtp(mail)
            if sender != OUTLOOK_CLIENT_EMAIL.lower():
                continue
            matched += 1
            categories = str(getattr(mail, "Categories", "") or "")
            if OUTLOOK_PROCESSED_CATEGORY in categories:
                skipped_processed += 1
                continue
            subject = str(getattr(mail, "Subject", "") or "")
            document_id = extract_document_id_from_subject(subject)
            provider = registry.provider_from_document_id(document_id, normalize_document_id)
            subject_file_name = extract_sent_filename_from_subject(subject)
            if not provider:
                skipped_no_mapping += 1
                log(f"Brak mapowania Id dokumentu: {document_id or subject}")
                continue
            if not subject_file_name:
                skipped_no_subject += 1
                log(f"Nie udało się utworzyć nazwy z tematu: {subject}")
                continue
            client = registry.get(provider)
            target_folder = base_dir / client.sent_folder
            target_folder.mkdir(parents=True, exist_ok=True)
            if getattr(mail, "Attachments", None) is None or mail.Attachments.Count == 0:
                continue
            mail_saved = 0
            for i in range(1, mail.Attachments.Count + 1):
                attachment = mail.Attachments.Item(i)
                file_name = str(attachment.FileName)
                if not file_name.lower().endswith(".xml"):
                    continue
                xml_seen += 1
                save_path = target_folder / subject_file_name
                if save_path.exists() or mail_saved > 0:
                    stem, suffix = Path(subject_file_name).stem, Path(subject_file_name).suffix
                    counter = 2
                    while True:
                        candidate = target_folder / f"{stem}_{counter}{suffix}"
                        if not candidate.exists():
                            save_path = candidate
                            break
                        counter += 1
                attachment.SaveAsFile(str(save_path))
                saved += 1
                mail_saved += 1
                log(f"OK {provider}: zapisano {save_path.name}")
            if mail_saved > 0:
                mail.Categories = (categories + "; " if categories else "") + OUTLOOK_PROCESSED_CATEGORY
                mail.Save()
                processed += 1
        except Exception as exc:
            log(f"Pominięto jedną wiadomość: {exc}")
    return (f"Transfer zakończony.\n\nSprawdzone wiadomości: {checked}\nWiadomości od nadawcy: {matched}\n"
            f"Pominięte jako przetworzone: {skipped_processed}\nPominięte bez mapowania: {skipped_no_mapping}\n"
            f"Pominięte bez nazwy z tematu: {skipped_no_subject}\nZnalezione XML: {xml_seen}\n"
            f"Przetworzone wiadomości: {processed}\nZapisane pliki XML: {saved}")
