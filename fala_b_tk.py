import base64
import binascii
import json
import hashlib
import hmac
import os
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from copy import deepcopy
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Dict


def _excel_fixed(value: float, digits: int) -> float:
    """Replikacja działania funkcji FIXED z Excela."""
    if digits < 0:
        raise ValueError("Liczba miejsc po przecinku nie może być ujemna.")
    quant = Decimal("1").scaleb(-digits)
    return float(Decimal(str(value)).quantize(quant, rounding=ROUND_HALF_UP))


CONFIG_DIR_NAME = "kalkulator_retruso"
CONFIG_FILE_NAME = "config.json"
DEFAULT_MARGIN_RULES = [
    {"max_quantity": 300, "margin_percent": 100.0},
    {"max_quantity": 500, "margin_percent": 70.0},
    {"max_quantity": 1000, "margin_percent": 25.0},
]
PBKDF2_ITERATIONS = 120_000


def _get_config_dir() -> Path:
    local_app_data = os.getenv("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / CONFIG_DIR_NAME
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / CONFIG_DIR_NAME
    return Path.home() / ".config" / CONFIG_DIR_NAME


class ConfigManager:
    def __init__(self) -> None:
        self.config_dir = _get_config_dir()
        self.config_file = self.config_dir / CONFIG_FILE_NAME
        self.data: Dict[str, Any] = {
            "password": None,
            "margin_rules": deepcopy(DEFAULT_MARGIN_RULES),
        }
        self.load()

    def load(self) -> None:
        try:
            with self.config_file.open("r", encoding="utf-8") as file:
                raw_data = json.load(file)
        except FileNotFoundError:
            return
        except (json.JSONDecodeError, OSError):
            return

        password_info = raw_data.get("password")
        if isinstance(password_info, dict):
            salt = password_info.get("salt")
            hash_value = password_info.get("hash")
            iterations = password_info.get("iterations", PBKDF2_ITERATIONS)
            if salt and hash_value:
                try:
                    iterations_int = int(iterations)
                except (TypeError, ValueError):
                    iterations_int = PBKDF2_ITERATIONS
                self.data["password"] = {
                    "salt": str(salt),
                    "hash": str(hash_value),
                    "iterations": iterations_int,
                }

        margin_rules = self._sanitize_margin_rules(raw_data.get("margin_rules"))
        if margin_rules is not None:
            self.data["margin_rules"] = margin_rules

    def save(self) -> None:
        try:
            self.config_dir.mkdir(parents=True, exist_ok=True)
            with self.config_file.open("w", encoding="utf-8") as file:
                json.dump(
                    {
                        "password": self.data.get("password"),
                        "margin_rules": self.data.get("margin_rules", []),
                    },
                    file,
                    ensure_ascii=False,
                    indent=2,
                )
        except OSError:
            pass

    def has_password(self) -> bool:
        password_info = self.data.get("password")
        return bool(
            isinstance(password_info, dict)
            and password_info.get("salt")
            and password_info.get("hash")
        )

    def set_password(self, password: str) -> None:
        salt_b64, hash_b64 = self._hash_password(password)
        self.data["password"] = {
            "salt": salt_b64,
            "hash": hash_b64,
            "iterations": PBKDF2_ITERATIONS,
        }
        self.save()

    def verify_password(self, password: str) -> bool:
        password_info = self.data.get("password")
        if not isinstance(password_info, dict):
            return False
        salt_b64 = password_info.get("salt")
        hash_b64 = password_info.get("hash")
        iterations = password_info.get("iterations", PBKDF2_ITERATIONS)
        if not salt_b64 or not hash_b64:
            return False
        try:
            salt = base64.b64decode(str(salt_b64))
            stored_hash = base64.b64decode(str(hash_b64))
            iterations_int = int(iterations)
        except (binascii.Error, TypeError, ValueError):
            return False
        new_hash = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), salt, iterations_int
        )
        return hmac.compare_digest(new_hash, stored_hash)

    def get_margin_rules(self) -> list[dict[str, float]]:
        return deepcopy(self.data.get("margin_rules", []))

    def update_margin_rules(
        self, rules: list[dict[str, float]]
    ) -> list[dict[str, float]]:
        sanitized = self._sanitize_margin_rules(rules)
        if sanitized is None:
            sanitized = deepcopy(DEFAULT_MARGIN_RULES)
        self.data["margin_rules"] = sanitized
        self.save()
        return deepcopy(self.data["margin_rules"])

    def _hash_password(self, password: str) -> tuple[str, str]:
        salt = os.urandom(16)
        password_hash = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS
        )
        return (
            base64.b64encode(salt).decode("ascii"),
            base64.b64encode(password_hash).decode("ascii"),
        )

    def _sanitize_margin_rules(
        self, value: Any
    ) -> list[dict[str, float]] | None:
        if not isinstance(value, list):
            return None
        sanitized: list[dict[str, float]] = []
        for item in value:
            if not isinstance(item, dict):
                continue
            try:
                max_quantity = int(item.get("max_quantity"))
                margin_percent = float(item.get("margin_percent"))
            except (TypeError, ValueError):
                continue
            sanitized.append(
                {"max_quantity": max_quantity, "margin_percent": margin_percent}
            )
        sanitized.sort(key=lambda rule: rule["max_quantity"])
        return sanitized

def oblicz_fala_b(
    dl: float,
    sz: float,
    wys: float,
    gramatura: float,
    cena_m2: float,
    dodatkowe_koszty: float,
    stawka_transport_km: float,
    dystans_km: float,
    transport_powrot: bool = True,
) -> Dict[str, Any]:
    """
    Przelicza wszystkie zależności z arkusza "FALA B".
    """

    # --- BIGI I BIGOWE (wiersze 8–9) ---
    c8 = (sz / 2.0) + 2.0
    d8 = wys + 10.0
    e8 = (sz / 2.0) + 2.0
    f8 = sz + 1.0
    g8 = dl + 3.0
    h8 = sz + 3.0
    i8 = dl + 3.0
    j8 = 35.0

    c9 = c8
    d9 = c8 + d8
    e9 = c8 + d8 + e8
    f9 = f8
    g9 = f8 + g8
    h9 = f8 + g8 + h8
    i9 = f8 + g8 + h8 + i8
    j9 = f8 + g8 + h8 + i8 + j8

    # --- FORMATKA I WYMIAR ZEWNĘTRZNY ---
    formatka_c11 = (((dl + sz) * 2.0) + 35.0 + 12.0) - 2.0
    wymiar_zewnetrzny_e11 = c8 + d8 + e8

    # --- ZUŻYCIE M2 ORAZ WAGA ---
    zuzycie_m2 = _excel_fixed((formatka_c11 * wymiar_zewnetrzny_e11) / 1_000_000.0, 3)
    waga_kg = (gramatura * zuzycie_m2) / 1000.0 if gramatura else 0.0

    # --- PODSTAWOWE KOSZTY ---
    koszt_mat_na_szt = zuzycie_m2 * cena_m2

    # --- MINIMUM PRODUKCYJNE I WERYFIKACJA ---
    min_aq = 500.0 / formatka_c11 * 1000.0 if formatka_c11 else 0.0
    min_con = 300.0 / zuzycie_m2 if zuzycie_m2 else 0.0
    min_pg = 500.0 / zuzycie_m2 if zuzycie_m2 else 0.0

    weryfikacja_dl = i8 + 3.0
    weryfikacja_sz = h8 + 3.0
    weryfikacja_wys = d8 + 2.0

    paletyzacja_dlugosc = f8 + g8
    paletyzacja_szerokosc = wymiar_zewnetrzny_e11

    # --- TRANSPORT ---
    stawka_pelna = stawka_transport_km * (2.0 if transport_powrot else 1.0)
    koszt_transport_calk = stawka_pelna * max(dystans_km, 0.0)

    return {
        "bigi": {"c8": c8, "d8": d8, "e8": e8},
        "bigowe": {"f8": f8, "g8": g8, "h8": h8, "i8": i8, "j8": j8},
        "sumy_bigowe": {"c9": c9, "d9": d9, "e9": e9, "f9": f9, "g9": g9, "h9": h9, "i9": i9, "j9": j9},
        "formatka_mm": formatka_c11,
        "wymiar_zewnetrzny_mm": wymiar_zewnetrzny_e11,
        "zuzycie_m2_na_szt": zuzycie_m2,
        "waga_kg_na_szt": waga_kg,
        "koszt_mat_na_szt": koszt_mat_na_szt,
        "koszty_dodatkowe": dodatkowe_koszty,
        "minimum_produkcji": {"aq": min_aq, "con": min_con, "pg": min_pg},
        "weryfikacja_zewnetrzna": {"dl": weryfikacja_dl, "sz": weryfikacja_sz, "wys": weryfikacja_wys},
        "paletyzacja": {"dlugosc": paletyzacja_dlugosc, "szerokosc": paletyzacja_szerokosc},
        "transport": {
            "stawka_pelna": stawka_pelna,
            "koszt_calkowity": koszt_transport_calk,
            "dystans": dystans_km,
            "powrot": transport_powrot,
        },
    }


class FalaBApp(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=12)
        self.master.title("Kalkulator Rekruso")
        self.grid(sticky="nsew")
        self.config = ConfigManager()
        self.margin_rules: list[dict[str, float]] = self.config.get_margin_rules()
        self.settings_unlocked = False
        self._init_variables()
        self.last_results: Dict[str, Any] = {}
        self.create_widgets()
        self.master.bind("<Return>", lambda _event: self.policz())

    def _init_variables(self) -> None:
        self.var_client_name = tk.StringVar()
        self.var_client_address = tk.StringVar()
        self.var_client_nip = tk.StringVar()
        self.var_client_email = tk.StringVar()

        self.var_dl = tk.StringVar(value="400")
        self.var_sz = tk.StringVar(value="300")
        self.var_wys = tk.StringVar(value="200")
        self.var_gram = tk.StringVar(value="675")
        self.var_cena_m2 = tk.StringVar(value="1.11")

        self.var_inne = tk.StringVar(value="0")

        self.var_transport_stawka = tk.StringVar(value="2.5")
        self.var_transport_km = tk.StringVar(value="0")
        self.var_transport_powrot = tk.BooleanVar(value=True)

        self.var_minimum_aq = tk.StringVar(value=": – szt.")
        self.var_minimum_con = tk.StringVar(value=": – szt.")
        self.var_minimum_pg = tk.StringVar(value=": – szt.")
        self.var_wymiar_h12 = tk.StringVar(value=": – mm")
        self.var_wymiar_i12 = tk.StringVar(value=": – mm")
        self.var_wymiar_j12 = tk.StringVar(value=": – mm")
        self.var_paletyzacja_dl = tk.StringVar(value=": – mm")
        self.var_paletyzacja_sz = tk.StringVar(value=": – mm")
        self.var_costs = tk.StringVar(value="–")
        self.var_transport_info = tk.StringVar(value="–")

        self.var_bigi_row1 = tk.StringVar(value="– mm | – mm | – mm")
        self.var_bigi_row2 = tk.StringVar(
            value="Pozycje bigów: – mm | – mm | – mm"
        )
        self.var_bigowanie_row1 = tk.StringVar(
            value="Szerokości segmentów: – mm | – mm | – mm | – mm | – mm"
        )
        self.var_bigowanie_row2 = tk.StringVar(
            value="Pozycje segmentów: – mm | – mm | – mm | – mm | – mm"
        )
        self.var_formatka_dims = tk.StringVar(
            value="Długość: – mm | Szerokość: – mm"
        )

    def create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.tab_calculator = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_calculator, text="Kalkulator")
        self.notebook.add(self.tab_settings, text="Ustawienia")

        self._build_calculator_tab()
        self._build_settings_tab()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _build_calculator_tab(self) -> None:
        tab = self.tab_calculator
        tab.columnconfigure(0, weight=1, uniform="col")
        tab.columnconfigure(1, weight=1, uniform="col")

        frame_client = ttk.LabelFrame(tab, text="Dane klienta")
        frame_client.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        for col in (1, 3):
            frame_client.columnconfigure(col, weight=1)

        ttk.Label(frame_client, text="Nazwa firmy").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_client, textvariable=self.var_client_name).grid(
            row=0, column=1, sticky="we", padx=(0, 8)
        )
        ttk.Label(frame_client, text="Adres").grid(row=0, column=2, sticky="w")
        ttk.Entry(frame_client, textvariable=self.var_client_address).grid(
            row=0, column=3, sticky="we"
        )
        ttk.Label(frame_client, text="NIP").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frame_client, textvariable=self.var_client_nip).grid(
            row=1, column=1, sticky="we", padx=(0, 8), pady=(6, 0)
        )
        ttk.Label(frame_client, text="E-mail").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(frame_client, textvariable=self.var_client_email).grid(
            row=1, column=3, sticky="we", pady=(6, 0)
        )

        frame_inputs = ttk.LabelFrame(tab, text="Parametry kartonu i nakłady")
        frame_inputs.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=(0, 8))
        for col in (1, 3, 5):
            frame_inputs.columnconfigure(col, weight=1)

        bold_font = ("TkDefaultFont", 10, "bold")
        ttk.Label(frame_inputs, text="DŁUGOŚĆ (mm)", font=bold_font).grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_inputs, textvariable=self.var_dl, width=10).grid(row=0, column=1, sticky="we", padx=(0, 6))
        ttk.Label(frame_inputs, text="SZEROKOŚĆ (mm)", font=bold_font).grid(row=0, column=2, sticky="w", padx=(6, 0))
        ttk.Entry(frame_inputs, textvariable=self.var_sz, width=10).grid(row=0, column=3, sticky="we", padx=(0, 6))
        ttk.Label(frame_inputs, text="WYSOKOŚĆ (mm)", font=bold_font).grid(row=0, column=4, sticky="w")
        ttk.Entry(frame_inputs, textvariable=self.var_wys, width=10).grid(row=0, column=5, sticky="we")

        ttk.Label(frame_inputs, text="Gramatura [g/m²]").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frame_inputs, textvariable=self.var_gram, width=10).grid(row=1, column=1, sticky="we", padx=(0, 6), pady=(6, 0))
        ttk.Label(frame_inputs, text="Cena surowca 1 m² [zł]").grid(row=1, column=2, sticky="w", padx=(6, 0), pady=(6, 0))
        ttk.Entry(frame_inputs, textvariable=self.var_cena_m2, width=10).grid(row=1, column=3, sticky="we", padx=(0, 6), pady=(6, 0))

        ttk.Separator(frame_inputs).grid(row=2, column=0, columnspan=6, sticky="we", pady=8)
        summary_frame = ttk.Frame(frame_inputs)
        summary_frame.grid(row=3, column=0, columnspan=6, sticky="nsew", pady=(8, 0))
        for col in range(3):
            summary_frame.columnconfigure(col, weight=1)

        minimum_frame = ttk.LabelFrame(summary_frame, text="Minimum produkcyjne")
        minimum_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        for row, (label_text, var) in enumerate(
            (
                ("AQ", self.var_minimum_aq),
                ("CON", self.var_minimum_con),
                ("PG", self.var_minimum_pg),
            )
        ):
            ttk.Label(minimum_frame, text=label_text, font=bold_font).grid(
                row=row, column=0, sticky="w"
            )
            ttk.Label(minimum_frame, textvariable=var).grid(
                row=row, column=1, sticky="w", padx=(4, 0)
            )

        wymiar_frame = ttk.LabelFrame(summary_frame, text="Wymiar zewnętrzny")
        wymiar_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        for row, (label_text, var) in enumerate(
            (
                ("Długość zewnętrzna", self.var_wymiar_h12),
                ("Szerokość zewnętrzna", self.var_wymiar_i12),
                ("Wysokość zewnętrzna", self.var_wymiar_j12),
            )
        ):
            ttk.Label(wymiar_frame, text=label_text, font=bold_font).grid(
                row=row, column=0, sticky="w"
            )
            ttk.Label(wymiar_frame, textvariable=var).grid(
                row=row, column=1, sticky="w", padx=(4, 0)
            )

        paletyzacja_frame = ttk.LabelFrame(summary_frame, text="Paletyzacja")
        paletyzacja_frame.grid(row=0, column=2, sticky="nsew")
        for row, (label_text, var) in enumerate(
            (
                ("Długość paletyzacyjna", self.var_paletyzacja_dl),
                ("Szerokość paletyzacyjna", self.var_paletyzacja_sz),
            )
        ):
            ttk.Label(paletyzacja_frame, text=label_text, font=bold_font).grid(
                row=row, column=0, sticky="w"
            )
            ttk.Label(paletyzacja_frame, textvariable=var).grid(
                row=row, column=1, sticky="w", padx=(4, 0)
            )

        frame_right = ttk.Frame(tab)
        frame_right.grid(row=1, column=1, sticky="nsew", pady=(0, 8))
        frame_right.columnconfigure(0, weight=1)
        frame_right.rowconfigure(0, weight=1)

        frame_machine = ttk.LabelFrame(frame_right, text="Ustawienia maszyny")
        frame_machine.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        frame_machine.columnconfigure(1, weight=1)

        ttk.Label(frame_machine, text="⚙", font=("TkDefaultFont", 20)).grid(
            row=0,
            column=0,
            rowspan=8,
            sticky="n",
            padx=(0, 8),
            pady=(0, 4),
        )
        ttk.Label(frame_machine, text="Bigi", font=bold_font).grid(
            row=0, column=1, sticky="w"
        )
        ttk.Label(frame_machine, textvariable=self.var_bigi_row1).grid(
            row=1, column=1, sticky="w"
        )
        ttk.Separator(frame_machine, orient="horizontal").grid(
            row=2, column=1, sticky="we", pady=(2, 4)
        )
        ttk.Label(frame_machine, textvariable=self.var_bigi_row2).grid(
            row=3, column=1, sticky="w"
        )
        ttk.Label(frame_machine, text="Bigowanie", font=bold_font).grid(
            row=4, column=1, sticky="w", pady=(6, 0)
        )
        ttk.Label(frame_machine, textvariable=self.var_bigowanie_row1).grid(
            row=5, column=1, sticky="w"
        )
        ttk.Label(frame_machine, textvariable=self.var_bigowanie_row2).grid(
            row=6, column=1, sticky="w"
        )

        frame_formatka = ttk.LabelFrame(frame_right, text="Formatka (mm)")
        frame_formatka.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=(8, 0))
        ttk.Label(frame_formatka, textvariable=self.var_formatka_dims).grid(
            row=0, column=0, sticky="w"
        )

        frame_costs = ttk.LabelFrame(tab, text="Koszty dodatkowe i transport")
        frame_costs.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        for col in (1, 3):
            frame_costs.columnconfigure(col, weight=1)

        ttk.Label(frame_costs, text="Dodatkowe koszty [zł/partia]").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_costs, textvariable=self.var_inne, width=10).grid(
            row=0, column=1, sticky="we", padx=(0, 8)
        )
        ttk.Label(frame_costs, text="Transport – stawka zł/km").grid(
            row=0, column=2, sticky="w"
        )
        ttk.Entry(frame_costs, textvariable=self.var_transport_stawka, width=10).grid(
            row=0, column=3, sticky="we"
        )
        ttk.Label(frame_costs, text="Dystans [km]").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Entry(frame_costs, textvariable=self.var_transport_km, width=10).grid(
            row=1, column=1, sticky="we", padx=(0, 8), pady=(6, 0)
        )
        ttk.Checkbutton(
            frame_costs,
            text="Uwzględnij powrót",
            variable=self.var_transport_powrot,
        ).grid(row=1, column=2, columnspan=2, sticky="w", pady=(6, 0))

        frame_actions = ttk.Frame(tab)
        frame_actions.grid(row=3, column=0, columnspan=2, sticky="we", pady=(0, 8))
        frame_actions.columnconfigure(0, weight=1)
        frame_actions.columnconfigure(1, weight=1)

        ttk.Button(frame_actions, text="Policz", command=self.policz).grid(
            row=0, column=0, sticky="we", padx=(0, 4)
        )
        ttk.Button(
            frame_actions,
            text="Drukuj podsumowanie",
            command=self.print_summary,
        ).grid(row=0, column=1, sticky="we", padx=(4, 0))

        frame_results = ttk.LabelFrame(tab, text="Wyniki")
        frame_results.grid(row=4, column=0, columnspan=2, sticky="nsew")
        frame_results.columnconfigure(0, weight=1)

        ttk.Label(frame_results, textvariable=self.var_costs, justify="left").grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(frame_results, textvariable=self.var_transport_info, justify="left").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )

        tab.rowconfigure(4, weight=1)

    def _build_settings_tab(self) -> None:
        tab = self.tab_settings
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)

        self.settings_locked_frame = ttk.Frame(tab, padding=20)
        self.settings_locked_frame.grid(row=0, column=0, sticky="nsew")
        self.settings_locked_frame.columnconfigure(1, weight=1)

        self.settings_content_frame = ttk.Frame(tab, padding=20)
        self.settings_content_frame.grid(row=0, column=0, sticky="nsew")
        self.settings_content_frame.columnconfigure(0, weight=1)
        self.settings_content_frame.columnconfigure(1, weight=1)
        self.settings_content_frame.columnconfigure(2, weight=1)
        self.settings_content_frame.columnconfigure(3, weight=1)

        self.var_settings_password = tk.StringVar()
        self.var_settings_password_confirm = tk.StringVar()
        self.settings_message_var = tk.StringVar()

        self.settings_locked_title = ttk.Label(
            self.settings_locked_frame, text=""
        )
        self.settings_locked_title.grid(row=0, column=0, columnspan=2, sticky="w")

        ttk.Label(self.settings_locked_frame, text="Hasło").grid(
            row=1, column=0, sticky="w", pady=(12, 0)
        )
        self.entry_settings_password = ttk.Entry(
            self.settings_locked_frame,
            textvariable=self.var_settings_password,
            show="*",
        )
        self.entry_settings_password.grid(row=1, column=1, sticky="we", pady=(12, 0))

        self.settings_locked_confirm_label = ttk.Label(
            self.settings_locked_frame, text="Powtórz hasło"
        )
        self.settings_locked_confirm_entry = ttk.Entry(
            self.settings_locked_frame,
            textvariable=self.var_settings_password_confirm,
            show="*",
        )

        self.settings_message_label = ttk.Label(
            self.settings_locked_frame, textvariable=self.settings_message_var
        )
        self.settings_message_label.grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(12, 0)
        )

        self.settings_action_button = ttk.Button(
            self.settings_locked_frame, command=self._handle_locked_action
        )
        self.settings_action_button.grid(
            row=4, column=0, columnspan=2, sticky="we", pady=(12, 0)
        )

        ttk.Label(
            self.settings_content_frame,
            text="Konfiguracja marży",
            font=("TkDefaultFont", 11, "bold"),
        ).grid(row=0, column=0, columnspan=4, sticky="w")

        self.margin_tree = ttk.Treeview(
            self.settings_content_frame,
            columns=("quantity", "margin"),
            show="headings",
            selectmode="browse",
            height=10,
        )
        self.margin_tree.heading("quantity", text="Do ilości [szt.]")
        self.margin_tree.heading("margin", text="Marża [%]")
        self.margin_tree.column("quantity", anchor="center", width=160)
        self.margin_tree.column("margin", anchor="center", width=120)
        self.margin_tree.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(12, 8))

        margin_scroll = ttk.Scrollbar(
            self.settings_content_frame,
            orient="vertical",
            command=self.margin_tree.yview,
        )
        margin_scroll.grid(row=1, column=3, sticky="nsw", pady=(12, 8))
        self.margin_tree.configure(yscrollcommand=margin_scroll.set)
        self.margin_tree.bind("<<TreeviewSelect>>", self._on_margin_tree_select)

        input_frame = ttk.Frame(self.settings_content_frame)
        input_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(4, 0))
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)

        ttk.Label(input_frame, text="Maks. ilość [szt.]").grid(
            row=0, column=0, sticky="w"
        )
        self.var_margin_quantity = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.var_margin_quantity).grid(
            row=0, column=1, sticky="ew", padx=(4, 12)
        )

        ttk.Label(input_frame, text="Marża [%]").grid(row=0, column=2, sticky="w")
        self.var_margin_value = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.var_margin_value).grid(
            row=0, column=3, sticky="ew", padx=(4, 0)
        )

        button_frame = ttk.Frame(self.settings_content_frame)
        button_frame.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)

        ttk.Button(
            button_frame,
            text="Dodaj / zaktualizuj",
            command=self._add_or_update_margin_rule,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(
            button_frame,
            text="Usuń zaznaczoną",
            command=self._remove_margin_rule,
        ).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(
            button_frame,
            text="Przywróć domyślne",
            command=self._reset_margin_rules,
        ).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        self.margin_message_var = tk.StringVar()
        self.margin_message_label = ttk.Label(
            self.settings_content_frame, textvariable=self.margin_message_var
        )
        self.margin_message_label.grid(
            row=4, column=0, columnspan=4, sticky="w", pady=(10, 0)
        )

        self.settings_content_frame.grid_remove()
        self._show_settings_locked()

    def _on_tab_changed(self, event: tk.Event) -> None:
        widget = event.widget
        if not isinstance(widget, ttk.Notebook):
            return
        current = widget.select()
        if current == str(self.tab_settings):
            if self.settings_unlocked:
                self._show_settings_content()
            else:
                self._show_settings_locked()

    def _show_settings_locked(self) -> None:
        self.settings_unlocked = False
        self.settings_content_frame.grid_remove()
        self.settings_locked_frame.grid()
        self.var_settings_password.set("")
        self.var_settings_password_confirm.set("")
        self._set_settings_message("")
        self._set_margin_message("")
        self._refresh_locked_frame_mode()

    def _show_settings_content(self) -> None:
        self.settings_locked_frame.grid_remove()
        self.settings_content_frame.grid()
        self.settings_unlocked = True
        self._set_settings_message("")
        self._set_margin_message("")
        self.margin_rules = self.config.get_margin_rules()
        self._refresh_margin_tree()

    def _refresh_locked_frame_mode(self) -> None:
        if self.config.has_password():
            self.settings_locked_title.config(
                text="Podaj hasło, aby otworzyć ustawienia."
            )
            self.settings_locked_confirm_label.grid_remove()
            self.settings_locked_confirm_entry.grid_remove()
            self.settings_action_button.config(text="Zaloguj")
        else:
            self.settings_locked_title.config(
                text="Utwórz hasło, aby zabezpieczyć ustawienia."
            )
            self.settings_locked_confirm_label.grid(
                row=2, column=0, sticky="w", pady=(6, 0)
            )
            self.settings_locked_confirm_entry.grid(
                row=2, column=1, sticky="we", pady=(6, 0)
            )
            self.settings_action_button.config(text="Zapisz hasło")
        self.entry_settings_password.focus_set()

    def _handle_locked_action(self) -> None:
        password = self.var_settings_password.get().strip()
        if not password:
            self._set_settings_message("Hasło nie może być puste.", error=True)
            return

        if self.config.has_password():
            if not self.config.verify_password(password):
                self._set_settings_message("Nieprawidłowe hasło.", error=True)
                self.var_settings_password.set("")
                return
            self._set_settings_message("")
            self.var_settings_password.set("")
            self._show_settings_content()
            return

        confirm = self.var_settings_password_confirm.get().strip()
        if password != confirm:
            self._set_settings_message("Hasła muszą być identyczne.", error=True)
            return
        self.config.set_password(password)
        self.var_settings_password.set("")
        self.var_settings_password_confirm.set("")
        self._set_settings_message("")
        messagebox.showinfo(
            "Hasło zapisane",
            "Hasło zostało utworzone. Zapisz je w bezpiecznym miejscu.",
        )
        self._show_settings_content()

    def _set_settings_message(self, text: str, error: bool = False) -> None:
        self.settings_message_var.set(text)
        self.settings_message_label.configure(foreground="red" if error else "")

    def _refresh_margin_tree(self) -> None:
        for item in self.margin_tree.get_children():
            self.margin_tree.delete(item)
        for index, rule in enumerate(self.margin_rules):
            qty = rule.get("max_quantity")
            margin = rule.get("margin_percent")
            qty_display = "-"
            try:
                if qty is not None:
                    qty_float = float(qty)
                    qty_display = (
                        f"{int(qty_float)}"
                        if float(qty_float).is_integer()
                        else f"{qty_float:.2f}"
                    )
            except (TypeError, ValueError):
                qty_display = "-"
            try:
                margin_float = float(margin) if margin is not None else None
            except (TypeError, ValueError):
                margin_float = None
            margin_display = "-" if margin_float is None else f"{margin_float:.2f}"
            self.margin_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(qty_display, margin_display),
            )
        for item in self.margin_tree.selection():
            self.margin_tree.selection_remove(item)

    def _on_margin_tree_select(self, _event: tk.Event) -> None:
        selection = self.margin_tree.selection()
        if not selection:
            return
        item_id = selection[0]
        values = self.margin_tree.item(item_id, "values")
        if len(values) >= 2:
            self.var_margin_quantity.set(str(values[0]))
            self.var_margin_value.set(str(values[1]))

    def _add_or_update_margin_rule(self) -> None:
        qty_text = self.var_margin_quantity.get().strip().replace(",", ".")
        margin_text = self.var_margin_value.get().strip().replace(",", ".")
        if not qty_text or not margin_text:
            self._set_margin_message("Podaj wartości ilości i marży.", error=True)
            return

        try:
            qty_value = int(float(qty_text))
        except ValueError:
            self._set_margin_message("Nieprawidłowa liczba sztuk.", error=True)
            return
        if qty_value <= 0:
            self._set_margin_message(
                "Maksymalna ilość musi być większa od zera.", error=True
            )
            return

        try:
            margin_value = float(margin_text)
        except ValueError:
            self._set_margin_message("Nieprawidłowa wartość marży.", error=True)
            return
        if margin_value < 0:
            self._set_margin_message("Marża nie może być ujemna.", error=True)
            return

        existing = next(
            (rule for rule in self.margin_rules if rule["max_quantity"] == qty_value),
            None,
        )
        if existing is not None:
            existing["margin_percent"] = margin_value
        else:
            self.margin_rules.append(
                {"max_quantity": qty_value, "margin_percent": margin_value}
            )
        self.margin_rules.sort(key=lambda rule: rule["max_quantity"])
        self._persist_margin_rules()
        self._set_margin_message("Zapisano konfigurację marży.")
        self._clear_margin_form()

    def _remove_margin_rule(self) -> None:
        selection = self.margin_tree.selection()
        if not selection:
            self._set_margin_message("Zaznacz pozycję do usunięcia.", error=True)
            return
        item_id = selection[0]
        try:
            index = int(item_id)
        except ValueError:
            self._set_margin_message("Nie udało się usunąć pozycji.", error=True)
            return
        if not 0 <= index < len(self.margin_rules):
            self._set_margin_message("Nie udało się usunąć pozycji.", error=True)
            return
        del self.margin_rules[index]
        self._persist_margin_rules()
        self._set_margin_message("Usunięto pozycję marży.")
        self._clear_margin_form()

    def _reset_margin_rules(self) -> None:
        if not messagebox.askyesno(
            "Przywracanie", "Czy na pewno chcesz przywrócić domyślne wartości marży?"
        ):
            return
        self.margin_rules = deepcopy(DEFAULT_MARGIN_RULES)
        self._persist_margin_rules()
        self._set_margin_message("Przywrócono domyślne wartości marży.")
        self._clear_margin_form()

    def _persist_margin_rules(self) -> None:
        self.margin_rules = self.config.update_margin_rules(self.margin_rules)
        self._refresh_margin_tree()

    def _clear_margin_form(self) -> None:
        self.var_margin_quantity.set("")
        self.var_margin_value.set("")
        for item in self.margin_tree.selection():
            self.margin_tree.selection_remove(item)

    def _set_margin_message(self, text: str, error: bool = False) -> None:
        self.margin_message_var.set(text)
        self.margin_message_label.configure(foreground="red" if error else "")

    def _parse_float(self, var: tk.StringVar, name: str, default: float | None = None) -> float:
        text = str(var.get()).strip().replace(",", ".")
        if not text:
            if default is not None:
                return default
            raise ValueError(f"Wymagana wartość w polu: {name}")
        try:
            return float(text)
        except ValueError as exc:
            raise ValueError(f"Nieprawidłowa wartość w polu: {name}") from exc

    def _parse_float_optional(self, var: tk.StringVar, name: str) -> float:
        return self._parse_float(var, name, default=0.0)

    def policz(self) -> None:
        try:
            dl = self._parse_float(self.var_dl, "DŁ")
            sz = self._parse_float(self.var_sz, "SZ")
            wys = self._parse_float(self.var_wys, "WYS")
            gram = self._parse_float(self.var_gram, "Gramatura")
            cena_m2 = self._parse_float(self.var_cena_m2, "Cena 1 m²")
            dodatkowe = self._parse_float_optional(
                self.var_inne, "Dodatkowe koszty"
            )
            stawka_km = self._parse_float_optional(self.var_transport_stawka, "Stawka transport")
            dystans = self._parse_float_optional(self.var_transport_km, "Dystans km")
            powrot = bool(self.var_transport_powrot.get())
        except ValueError as exc:
            messagebox.showerror("Błąd danych", str(exc))
            return

        wyniki = oblicz_fala_b(
            dl=dl,
            sz=sz,
            wys=wys,
            gramatura=gram,
            cena_m2=cena_m2,
            dodatkowe_koszty=dodatkowe,
            stawka_transport_km=stawka_km,
            dystans_km=dystans,
            transport_powrot=powrot,
        )

        bigi = wyniki["bigi"]
        bigowe = wyniki["bigowe"]
        sumy_bigowe = wyniki["sumy_bigowe"]

        self.var_bigi_row1.set(
            f"{bigi['c8']:.2f} mm | {bigi['d8']:.2f} mm | {bigi['e8']:.2f} mm"
        )
        self.var_bigi_row2.set(
            "Pozycje bigów: "
            + " | ".join(
                [
                    f"{sumy_bigowe['c9']:.2f} mm",
                    f"{sumy_bigowe['d9']:.2f} mm",
                    f"{sumy_bigowe['e9']:.2f} mm",
                ]
            )
        )
        self.var_bigowanie_row1.set(
            "Szerokości segmentów: "
            + " | ".join(
                [
                    f"{bigowe['f8']:.2f} mm",
                    f"{bigowe['g8']:.2f} mm",
                    f"{bigowe['h8']:.2f} mm",
                    f"{bigowe['i8']:.2f} mm",
                    f"{bigowe['j8']:.2f} mm",
                ]
            )
        )
        self.var_bigowanie_row2.set(
            "Pozycje segmentów: "
            + " | ".join(
                [
                    f"{sumy_bigowe['f9']:.2f} mm",
                    f"{sumy_bigowe['g9']:.2f} mm",
                    f"{sumy_bigowe['h9']:.2f} mm",
                    f"{sumy_bigowe['i9']:.2f} mm",
                    f"{sumy_bigowe['j9']:.2f} mm",
                ]
            )
        )
        self.var_formatka_dims.set(
            f"Długość: {wyniki['formatka_mm']:.2f} mm | "
            f"Szerokość: {wyniki['wymiar_zewnetrzny_mm']:.2f} mm"
        )

        minimum = wyniki["minimum_produkcji"]
        self.var_minimum_aq.set(f": {minimum['aq']:.0f} szt.")
        self.var_minimum_con.set(f": {minimum['con']:.0f} szt.")
        self.var_minimum_pg.set(f": {minimum['pg']:.0f} szt.")

        weryf = wyniki["weryfikacja_zewnetrzna"]
        self.var_wymiar_h12.set(f": {weryf['dl']:.2f} mm")
        self.var_wymiar_i12.set(f": {weryf['sz']:.2f} mm")
        self.var_wymiar_j12.set(f": {weryf['wys']:.2f} mm")

        paletyzacja = wyniki["paletyzacja"]
        self.var_paletyzacja_dl.set(f": {paletyzacja['dlugosc']:.2f} mm")
        self.var_paletyzacja_sz.set(f": {paletyzacja['szerokosc']:.2f} mm")

        koszt_lines = [
            f"Zużycie m²/szt.: {wyniki['zuzycie_m2_na_szt']:.3f}",
            f"Waga kg/szt.: {wyniki['waga_kg_na_szt']:.3f}",
            f"Koszt materiału/szt.: {wyniki['koszt_mat_na_szt']:.4f} zł",
        ]
        if wyniki["koszty_dodatkowe"]:
            koszt_lines.append(
                f"Dodatkowe koszty (partia): {wyniki['koszty_dodatkowe']:.2f} zł"
            )
        self.var_costs.set("\n".join(koszt_lines))

        transport = wyniki["transport"]
        powrot_txt = "tak" if transport["powrot"] else "nie"
        self.var_transport_info.set(
            f"Transport: stawka {transport['stawka_pelna']:.2f} zł/km, "
            f"dystans {transport['dystans']:.2f} km, powrót: {powrot_txt}. "
            f"Koszt łączny: {transport['koszt_calkowity']:.2f} zł"
        )

        self.last_results = {
            "client": {
                "nazwa": self.var_client_name.get().strip(),
                "adres": self.var_client_address.get().strip(),
                "nip": self.var_client_nip.get().strip(),
                "email": self.var_client_email.get().strip(),
            },
            "inputs": {
                "dl": dl,
                "sz": sz,
                "wys": wys,
                "gramatura": gram,
                "cena_m2": cena_m2,
                "dodatkowe_koszty": dodatkowe,
                "stawka_transport": stawka_km,
                "dystans": dystans,
                "powrot": powrot,
            },
            "wyniki": wyniki,
            "margin_rules": self.config.get_margin_rules(),
        }

    def _build_print_summary(self) -> str:
        if not self.last_results:
            raise ValueError("Brak danych do wydruku. Najpierw wykonaj obliczenia.")

        inputs = self.last_results.get("inputs", {})
        wyniki = self.last_results.get("wyniki", {})

        def fmt(value: Any, digits: int = 2) -> str:
            try:
                return f"{float(value):.{digits}f}"
            except (TypeError, ValueError):
                return "-"

        rows: list[tuple[str, str, str]] = []

        def add_row(section: str, label: str, value: str) -> None:
            rows.append((section, label, value))

        add_row("Informacje", "Nazwa programu", "Kalkulator Rekruso")
        add_row(
            "Informacje",
            "Data wydruku",
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        )

        add_row(
            "Parametry wejściowe",
            "Długość (DL) [mm]",
            fmt(inputs.get("dl")),
        )
        add_row(
            "Parametry wejściowe",
            "Szerokość (SZ) [mm]",
            fmt(inputs.get("sz")),
        )
        add_row(
            "Parametry wejściowe",
            "Wysokość (WYS) [mm]",
            fmt(inputs.get("wys")),
        )
        add_row(
            "Parametry wejściowe",
            "Gramatura [g/m²]",
            fmt(inputs.get("gramatura"), 0),
        )
        add_row(
            "Parametry wejściowe",
            "Cena surowca 1 m² [zł]",
            fmt(inputs.get("cena_m2"), 4),
        )
        add_row(
            "Parametry wejściowe",
            "Dodatkowe koszty (partia) [zł]",
            fmt(inputs.get("dodatkowe_koszty")),
        )
        add_row(
            "Parametry wejściowe",
            "Transport – stawka [zł/km]",
            fmt(inputs.get("stawka_transport")),
        )
        add_row(
            "Parametry wejściowe",
            "Transport – dystans [km]",
            fmt(inputs.get("dystans")),
        )
        add_row(
            "Parametry wejściowe",
            "Transport – powrót",
            "tak" if inputs.get("powrot") else "nie",
        )

        bigi = wyniki.get("bigi", {})
        bigowe = wyniki.get("bigowe", {})
        sumy = wyniki.get("sumy_bigowe", {})

        add_row("Bigi", "C8 [mm]", fmt(bigi.get("c8")))
        add_row("Bigi", "D8 [mm]", fmt(bigi.get("d8")))
        add_row("Bigi", "E8 [mm]", fmt(bigi.get("e8")))

        add_row("Pozycje bigów", "C9 [mm]", fmt(sumy.get("c9")))
        add_row("Pozycje bigów", "D9 [mm]", fmt(sumy.get("d9")))
        add_row("Pozycje bigów", "E9 [mm]", fmt(sumy.get("e9")))

        add_row("Szerokości segmentów", "F8 [mm]", fmt(bigowe.get("f8")))
        add_row("Szerokości segmentów", "G8 [mm]", fmt(bigowe.get("g8")))
        add_row("Szerokości segmentów", "H8 [mm]", fmt(bigowe.get("h8")))
        add_row("Szerokości segmentów", "I8 [mm]", fmt(bigowe.get("i8")))
        add_row("Szerokości segmentów", "J8 [mm]", fmt(bigowe.get("j8")))

        add_row("Pozycje segmentów", "F9 [mm]", fmt(sumy.get("f9")))
        add_row("Pozycje segmentów", "G9 [mm]", fmt(sumy.get("g9")))
        add_row("Pozycje segmentów", "H9 [mm]", fmt(sumy.get("h9")))
        add_row("Pozycje segmentów", "I9 [mm]", fmt(sumy.get("i9")))
        add_row("Pozycje segmentów", "J9 [mm]", fmt(sumy.get("j9")))

        add_row("Formatka", "Długość [mm]", fmt(wyniki.get("formatka_mm")))
        add_row(
            "Formatka",
            "Szerokość [mm]",
            fmt(wyniki.get("wymiar_zewnetrzny_mm")),
        )

        weryf = wyniki.get("weryfikacja_zewnetrzna", {})
        add_row("Weryfikacja", "DL [mm]", fmt(weryf.get("dl")))
        add_row("Weryfikacja", "SZ [mm]", fmt(weryf.get("sz")))
        add_row("Weryfikacja", "WYS [mm]", fmt(weryf.get("wys")))

        paletyzacja = wyniki.get("paletyzacja", {})
        add_row(
            "Paletyzacja",
            "Długość paletyzacyjna [mm]",
            fmt(paletyzacja.get("dlugosc")),
        )
        add_row(
            "Paletyzacja",
            "Szerokość paletyzacyjna [mm]",
            fmt(paletyzacja.get("szerokosc")),
        )

        minimum = wyniki.get("minimum_produkcji", {})
        add_row("Minimum produkcyjne", "AQ [szt.]", fmt(minimum.get("aq"), 0))
        add_row("Minimum produkcyjne", "CON [szt.]", fmt(minimum.get("con"), 0))
        add_row("Minimum produkcyjne", "PG [szt.]", fmt(minimum.get("pg"), 0))

        add_row(
            "Zużycie i koszty",
            "Zużycie m²/szt.",
            fmt(wyniki.get("zuzycie_m2_na_szt"), 3),
        )
        add_row(
            "Zużycie i koszty",
            "Waga kg/szt.",
            fmt(wyniki.get("waga_kg_na_szt"), 3),
        )
        add_row(
            "Zużycie i koszty",
            "Koszt materiału/szt. [zł]",
            fmt(wyniki.get("koszt_mat_na_szt"), 4),
        )
        add_row(
            "Zużycie i koszty",
            "Dodatkowe koszty (partia) [zł]",
            fmt(wyniki.get("koszty_dodatkowe")),
        )

        transport = wyniki.get("transport", {})
        add_row(
            "Transport (wynik)",
            "Stawka końcowa [zł/km]",
            fmt(transport.get("stawka_pelna")),
        )
        add_row(
            "Transport (wynik)",
            "Dystans [km]",
            fmt(transport.get("dystans")),
        )
        add_row(
            "Transport (wynik)",
            "Powrót",
            "tak" if transport.get("powrot") else "nie",
        )
        add_row(
            "Transport (wynik)",
            "Koszt łączny [zł]",
            fmt(transport.get("koszt_calkowity")),
        )

        margin_rules = self.last_results.get("margin_rules") or self.config.get_margin_rules()
        if margin_rules:
            for rule in margin_rules:
                qty_label = "Próg marży"
                qty_value = rule.get("max_quantity")
                try:
                    qty_int = int(float(qty_value))
                    qty_label = f"Do {qty_int} szt."
                except (TypeError, ValueError):
                    pass
                try:
                    margin_value = float(rule.get("margin_percent"))
                    margin_text = f"{margin_value:.2f} %"
                except (TypeError, ValueError):
                    margin_text = "-"
                add_row("Marża", qty_label, margin_text)

        csv_lines = ["sep=;", "Sekcja;Parametr;Wartość"]
        csv_lines.extend(";".join(row) for row in rows)
        return "\n".join(csv_lines)

    @staticmethod
    def _send_to_printer(text: str) -> None:
        suffix = ".txt" if sys.platform.startswith("win") else ".csv"
        temp_path: Path | None = None
        with tempfile.NamedTemporaryFile(
            "w", delete=False, encoding="utf-8", suffix=suffix
        ) as temp_file:
            temp_file.write(text)
            temp_path = Path(temp_file.name)

        def _cleanup(path: Path) -> None:
            try:
                path.unlink()
            except FileNotFoundError:
                pass

        try:
            if sys.platform.startswith("win"):
                try:
                    subprocess.run(["notepad.exe", "/p", str(temp_path)], check=True)
                except FileNotFoundError as exc:
                    if hasattr(os, "startfile"):
                        try:
                            os.startfile(str(temp_path), "print")  # type: ignore[attr-defined]
                        except OSError as inner_exc:
                            raise RuntimeError(
                                "Nie udało się uruchomić systemowego polecenia drukowania."
                            ) from inner_exc
                    else:
                        raise RuntimeError(
                            "Nie znaleziono programu Notepad wymaganego do wydruku."
                        ) from exc
                except subprocess.CalledProcessError as exc:
                    raise RuntimeError(
                        "System Windows zgłosił błąd podczas drukowania."
                    ) from exc
            elif sys.platform == "darwin":
                subprocess.run(["lp", str(temp_path)], check=True)
            else:
                subprocess.run(["lpr", str(temp_path)], check=True)
        except FileNotFoundError as exc:
            raise RuntimeError("Nie znaleziono polecenia drukarki w systemie.") from exc
        except subprocess.CalledProcessError as exc:
            raise RuntimeError("Polecenie drukarki zakończyło się niepowodzeniem.") from exc
        except Exception as exc:
            raise RuntimeError("Nie udało się wysłać danych do drukarki.") from exc
        finally:
            if temp_path is not None:
                threading.Timer(10.0, _cleanup, args=(temp_path,)).start()

    def print_summary(self) -> None:
        try:
            summary = self._build_print_summary()
        except ValueError as exc:
            messagebox.showinfo("Brak danych", str(exc))
            return
        except Exception as exc:
            messagebox.showerror(
                "Błąd", f"Nie udało się przygotować danych do wydruku.\n{exc}"
            )
            return

        try:
            self._send_to_printer(summary)
        except RuntimeError as exc:
            messagebox.showerror("Błąd drukowania", str(exc))
            return
        messagebox.showinfo("Drukowanie", "Podsumowanie zostało wysłane do drukarki.")


def main() -> None:
    root = tk.Tk()
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root.geometry("980x640")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    app = FalaBApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
