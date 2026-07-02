# -*- coding: utf-8 -*-
from __future__ import annotations
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import traceback
import re
from config import ClientRegistry
from utils import app_dir, safe_file_token, get_registration_from_row
from excel_service import load_excel, find_row_by_lp, detect_sent_type, import_sent_data, update_close_dates
from xml_builder import create_xml
from outlook_service import transfer_xml_from_outlook
from client_wizard import create_client

APP_TITLE = "SENT 2.0 - Generator XML"

LIGHT_THEME = {
    "bg": "#f4f6fb",
    "panel": "#ffffff",
    "panel2": "#eef2f8",
    "fg": "#172033",
    "muted": "#5f6f89",
    "entry": "#ffffff",
    "text_bg": "#ffffff",
    "text_fg": "#172033",
    "accent": "#2563eb",
    "danger": "#dc2626",
}

DARK_THEME = {
    "bg": "#111827",
    "panel": "#1f2937",
    "panel2": "#273449",
    "fg": "#f9fafb",
    "muted": "#cbd5e1",
    "entry": "#111827",
    "text_bg": "#0b1220",
    "text_fg": "#e5e7eb",
    "accent": "#60a5fa",
    "danger": "#f87171",
}

class TextPopup(tk.Toplevel):
    def __init__(self, parent, title, message, width=760, height=520):
        super().__init__(parent)
        self.parent = parent
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.configure(bg=parent.colors["bg"])
        frame = ttk.Frame(self, padding=12, style="Panel.TFrame")
        frame.pack(fill="both", expand=True, padx=12, pady=12)
        text = tk.Text(
            frame,
            wrap="word",
            bg=parent.colors["text_bg"],
            fg=parent.colors["text_fg"],
            insertbackground=parent.colors["text_fg"],
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=12,
        )
        text.insert("1.0", message)
        text.configure(state="disabled")
        text.pack(fill="both", expand=True)
        ttk.Button(frame, text="OK", command=self.destroy, style="Accent.TButton").pack(pady=(10, 0))
        self.grab_set()
        self.focus_force()

class AddClientDialog(tk.Toplevel):
    def __init__(self, parent, on_created):
        super().__init__(parent)
        self.parent = parent
        self.on_created = on_created
        self.title("Dodaj klienta")
        self.geometry("760x560")
        self.resizable(False, False)
        self.configure(bg=parent.colors["bg"])
        self.excel_path = tk.StringVar()
        self.name = tk.StringVar()
        self.excel_type = tk.StringVar(value="AUTO/ WAGON")
        self.document_id = tk.StringVar()
        self.rule_index = tk.StringVar(value="7")

        # Dane nadawcy sa czytane automatycznie z gornych komorek Excela.
        self.sender_name_cell = tk.StringVar(value="A1")
        self.sender_identity_cell = tk.StringVar(value="A2")
        self.sender_address_cell = tk.StringVar(value="A3")

        frm = ttk.Frame(self, padding=20, style="Panel.TFrame")
        frm.pack(fill="both", expand=True, padx=16, pady=16)
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text="Dodaj klienta", style="Title.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        ttk.Label(frm, text="Nazwa klienta:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.name).grid(row=1, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(frm, text="Excel awizacji:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.excel_path).grid(row=2, column=1, sticky="ew", pady=5)
        ttk.Button(frm, text="Wybierz...", command=self.choose_excel).grid(row=2, column=2, padx=(8, 0))
        ttk.Label(frm, text="Typ Excela / rejestracji:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Combobox(frm, textvariable=self.excel_type, values=["AUTO", "AUTO/ WAGON"], state="readonly").grid(row=3, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(frm, text="Bazowy DocumentId:").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.document_id).grid(row=4, column=1, columnspan=2, sticky="ew", pady=5)
        ttk.Label(frm, text="Indeks kolumny EXPRO:").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(frm, textvariable=self.rule_index, width=12).grid(row=5, column=1, sticky="w", pady=5)

        help_text = (
            "Indeks kolumny liczony od 0. Najczęściej 7; dopasuj do układu konkretnego Excela. "
            "Puste pole = zawsze domyślny SENT_205. Dla Excela typu AUTO program sam spróbuje użyć kolumny TRAILER/PRZYCZEPA na naczepę."
        )
        ttk.Label(frm, text=help_text, style="Muted.TLabel", wraplength=690).grid(row=6, column=0, columnspan=3, sticky="w", pady=(6, 14))

        sep = ttk.Separator(frm, orient="horizontal")
        sep.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        ttk.Label(frm, text="GoodsSender z Excela", style="Title.TLabel").grid(row=8, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Label(frm, text="Komórka z nazwą nadawcy:").grid(row=9, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.sender_name_cell, width=12).grid(row=9, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Komórka z typem i numerem ID:").grid(row=10, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.sender_identity_cell, width=12).grid(row=10, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Komórka z adresem:").grid(row=11, column=0, sticky="w", pady=4)
        ttk.Entry(frm, textvariable=self.sender_address_cell, width=12).grid(row=11, column=1, sticky="w", pady=4)

        ttk.Label(
            frm,
            text="Domyślnie program czyta A1/A2/A3. Przykład: A1 = nazwa firmy, A2 = INNY: numer albo VAT UE: numer, A3 = ulica nr, kod miasto. Te dane są pobierane przy generowaniu XML, więc nie trzeba uzupełniać ich ręcznie w YAML.",
            style="Muted.TLabel",
            wraplength=690,
        ).grid(row=12, column=0, columnspan=3, sticky="w", pady=(8, 14))

        ttk.Button(frm, text="Utwórz klienta", command=self.create, style="Accent.TButton").grid(row=13, column=0, columnspan=3, sticky="ew", ipady=4)
        self.grab_set()
        self.focus_force()

    def choose_excel(self):
        path = filedialog.askopenfilename(title="Wybierz Excel awizacji", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.excel_path.set(path)
            if not self.name.get().strip():
                stem = Path(path).stem.replace("AWIZACJA", "").strip()
                self.name.set(stem.upper())

    def create(self):
        try:
            raw_idx = self.rule_index.get().strip()
            rule_index = int(raw_idx) if raw_idx else None
            client = create_client(
                self.parent.base_dir,
                self.parent.registry,
                self.name.get(),
                Path(self.excel_path.get()),
                self.excel_type.get(),
                self.document_id.get().strip(),
                rule_index,
                self.sender_name_cell.get().strip() or "A1",
                self.sender_identity_cell.get().strip() or "A2",
                self.sender_address_cell.get().strip() or "A3",
            )
            self.on_created(client.name)
            messagebox.showinfo("Dodano klienta", f"Dodano klienta: {client.name}\nUtworzono mappingi w folderze mappings.")
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Błąd", str(exc))

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.base_dir = app_dir()
        self.registry = ClientRegistry.load(self.base_dir)
        self.previewed_row = None
        self.previewed_sent = None
        self.previewed_lp = None
        self.dark_mode = tk.BooleanVar(value=True)
        self.colors = DARK_THEME
        self.title(APP_TITLE)
        self.geometry("900x690")
        self.minsize(820, 620)
        self.option_add("*Font", ("Segoe UI", 10))
        self._configure_styles()
        self._build_ui()
        self.apply_theme()
        if self.provider.get() == "DEMO_CLIENT":
            self.ent_lp.insert(0, "1")
            self.log("Tryb demo: wybrano DEMO_CLIENT i wpisano LP=1. Kliknij 'Podgląd wiersza', a potem 'Generuj XML'.")
        else:
            self.log("Gotowe. Wybierz klienta, wpisz LP i użyj podglądu wiersza przed generowaniem XML.")

    def _configure_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

    def apply_theme(self):
        self.colors = DARK_THEME if self.dark_mode.get() else LIGHT_THEME
        c = self.colors
        self.configure(bg=c["bg"])
        self.style.configure("TFrame", background=c["bg"])
        self.style.configure("Panel.TFrame", background=c["panel"], relief="flat")
        self.style.configure("Card.TFrame", background=c["panel2"], relief="flat")
        self.style.configure("TLabel", background=c["panel"], foreground=c["fg"])
        self.style.configure("Muted.TLabel", background=c["panel"], foreground=c["muted"])
        self.style.configure("Title.TLabel", background=c["panel"], foreground=c["fg"], font=("Segoe UI", 18, "bold"))
        self.style.configure("Header.TLabel", background=c["bg"], foreground=c["fg"], font=("Segoe UI", 20, "bold"))
        self.style.configure("Subheader.TLabel", background=c["bg"], foreground=c["muted"], font=("Segoe UI", 10))
        self.style.configure("TButton", padding=(12, 7), background=c["panel2"], foreground=c["fg"], borderwidth=0)
        self.style.map("TButton", background=[("active", c["accent"])] , foreground=[("active", "#ffffff")])
        self.style.configure("Accent.TButton", padding=(12, 8), background=c["accent"], foreground="#ffffff", borderwidth=0, font=("Segoe UI", 10, "bold"))
        self.style.map("Accent.TButton", background=[("active", c["accent"])] , foreground=[("active", "#ffffff")])
        self.style.configure("Danger.TButton", padding=(12, 7), background=c["danger"], foreground="#ffffff", borderwidth=0)
        self.style.configure("TEntry", fieldbackground=c["entry"], foreground=c["fg"], insertcolor=c["fg"])
        self.style.configure("TCombobox", fieldbackground=c["entry"], foreground=c["fg"], arrowcolor=c["fg"])
        self.style.map("TCombobox", fieldbackground=[("readonly", c["entry"])] , foreground=[("readonly", c["fg"])])
        if hasattr(self, "txt_status"):
            self.txt_status.configure(bg=c["text_bg"], fg=c["text_fg"], insertbackground=c["text_fg"])
        if hasattr(self, "theme_btn"):
            self.theme_btn.configure(text="Tryb jasny" if self.dark_mode.get() else "Tryb ciemny")

    def _build_ui(self):
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        header = ttk.Frame(outer)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="SENT 2.0", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Generator XML z konfiguracją klientów", style="Subheader.TLabel").grid(row=1, column=0, sticky="w")
        self.theme_btn = ttk.Button(header, text="Tryb jasny", command=self.toggle_theme)
        self.theme_btn.grid(row=0, column=1, rowspan=2, sticky="e")

        top = ttk.Frame(outer, padding=16, style="Panel.TFrame")
        top.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=0)
        ttk.Label(top, text="Klient").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.provider = tk.StringVar()
        self.cb_provider = ttk.Combobox(top, textvariable=self.provider, values=self.registry.names(), state="readonly")
        self.cb_provider.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=5)
        if self.registry.names():
            self.cb_provider.current(0)
        ttk.Button(top, text="Dodaj klienta", command=self.open_add_client).grid(row=0, column=2, padx=4, pady=5)
        ttk.Button(top, text="Usuń klienta", command=self.delete_client, style="Danger.TButton").grid(row=0, column=3, padx=(4, 0), pady=5)
        ttk.Label(top, text="LP").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.ent_lp = ttk.Entry(top)
        self.ent_lp.grid(row=1, column=1, sticky="ew", padx=(0, 10), pady=5)
        ttk.Button(top, text="Podgląd wiersza", command=self.preview_row).grid(row=1, column=2, padx=4, pady=5)
        ttk.Button(top, text="Generuj XML", command=self.generate, style="Accent.TButton").grid(row=1, column=3, padx=(4, 0), pady=5)

        body = ttk.Frame(outer)
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        actions = ttk.Frame(body, padding=14, style="Panel.TFrame")
        actions.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        ttk.Label(actions, text="Operacje", style="Title.TLabel").pack(anchor="w", pady=(0, 10))
        action_buttons = [
            ("Zaciągnij dane z folderu SENTY", self.import_sent),
            ("Transfer XML z e-mail", self.transfer_xml),
            ("Uzupełnij daty zamknięcia SENT", self.update_close_dates),
            ("Pokaż konfigurację klienta", self.show_client_config),
        ]
        for label, command in action_buttons:
            ttk.Button(actions, text=label, command=command).pack(fill="x", pady=5)
        ttk.Label(actions, text="Dodawanie i usuwanie klientów działa na clients.yaml. Pliki Excela i mappingi zostają w folderach projektu.", style="Muted.TLabel", wraplength=240).pack(anchor="w", pady=(18, 0))

        status_panel = ttk.Frame(body, padding=14, style="Panel.TFrame")
        status_panel.grid(row=0, column=1, sticky="nsew")
        status_panel.rowconfigure(1, weight=1)
        status_panel.columnconfigure(0, weight=1)
        ttk.Label(status_panel, text="Status", style="Title.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 10))
        self.txt_status = tk.Text(status_panel, height=15, wrap="word", relief="flat", borderwidth=0, padx=12, pady=12)
        self.txt_status.grid(row=1, column=0, sticky="nsew")

    def toggle_theme(self):
        self.dark_mode.set(not self.dark_mode.get())
        self.apply_theme()

    def refresh_clients(self, select_name=None):
        self.registry = ClientRegistry.load(self.base_dir)
        names = self.registry.names()
        self.cb_provider.configure(values=names)
        if select_name and select_name in names:
            self.provider.set(select_name)
        elif names:
            self.provider.set(names[0])
        else:
            self.provider.set("")

    def get_client(self):
        return self.registry.get(self.provider.get())

    def log(self, msg):
        self.txt_status.insert("end", str(msg) + "\n")
        self.txt_status.see("end")
        self.update_idletasks()

    def open_add_client(self):
        AddClientDialog(self, self.refresh_clients)

    def delete_client(self):
        try:
            name = self.provider.get().strip()
            if not name:
                raise ValueError("Nie wybrano klienta do usunięcia.")
            if not messagebox.askyesno("Usuń klienta", f"Usunąć klienta {name} z listy?\n\nPliki Excel, YAML i folder SENTY nie zostaną skasowane."):
                return
            self.registry.remove_client(name)
            self.refresh_clients()
            self.previewed_row = None
            self.previewed_sent = None
            self.previewed_lp = None
            self.log(f"Usunięto klienta z konfiguracji: {name}")
        except Exception as exc:
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", str(exc))

    def preview_row(self):
        try:
            lp = self.ent_lp.get().strip()
            if not lp:
                raise ValueError("Podaj LP.")
            client = self.get_client()
            df = load_excel(self.base_dir, client)
            row = find_row_by_lp(df, lp)
            sent_type = detect_sent_type(row, client)
            self.previewed_row = row
            self.previewed_sent = sent_type
            self.previewed_lp = lp
            data = "\n".join([f"{col}: {row[col]}" for col in df.columns])
            data += f"\n\nPropozycja typu SENT: {sent_type}"
            TextPopup(self, f"Wiersz LP={lp}", data)
        except Exception as exc:
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", str(exc))

    def generate(self):
        try:
            if self.previewed_row is None:
                raise ValueError("Najpierw wykonaj podgląd wiersza (LP).")
            client = self.get_client()
            row = self.previewed_row.copy()
            sent_type = self.previewed_sent or detect_sent_type(row, client)
            self.log(f"Generowanie XML: klient={client.name}, LP={self.previewed_lp}, typ={sent_type}")
            xml_tree = create_xml(self.base_dir, row, client, sent_type)
            reg = get_registration_from_row(row, client.registration_column, client.registration_fallback_columns, client.registration_use_split_truck)
            outdir = self.base_dir / "output"
            outdir.mkdir(exist_ok=True)
            fname = f"{sent_type}_{client.name}_LP{self.previewed_lp}_{safe_file_token(reg)}.xml"
            fname = re.sub(r"[^A-Za-z0-9_.-]", "_", fname)
            outpath = outdir / fname
            xml_tree.write(str(outpath), encoding="utf-8", xml_declaration=True, pretty_print=True)
            self.log(f"Zapisano: {outpath}")
            messagebox.showinfo("Sukces", f"Wygenerowano plik:\n{outpath}")
        except Exception as exc:
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", str(exc))

    def import_sent(self):
        try:
            msg = import_sent_data(self.base_dir, self.get_client(), self.log)
            self.log(msg.replace("\n", " | "))
            TextPopup(self, "Import zakończony", msg)
        except Exception as exc:
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", str(exc))

    def update_close_dates(self):
        try:
            msg = update_close_dates(self.base_dir, self.get_client(), self.log)
            self.log(msg.replace("\n", " | "))
            TextPopup(self, "Daty zamknięcia SENT", msg)
        except Exception as exc:
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", str(exc))

    def transfer_xml(self):
        try:
            msg = transfer_xml_from_outlook(self.base_dir, self.registry, self.log)
            self.log(msg.replace("\n", " | "))
            TextPopup(self, "Transfer XML", msg)
        except Exception as exc:
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", str(exc))

    def show_client_config(self):
        try:
            client = self.get_client()
            import yaml
            TextPopup(self, f"Konfiguracja {client.name}", yaml.safe_dump(client.as_yaml_dict(), allow_unicode=True, sort_keys=False))
        except Exception as exc:
            messagebox.showerror("Błąd", str(exc))
