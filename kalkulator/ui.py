"""Warstwa interfejsu użytkownika aplikacji Kalkulator Rekruso."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Dict

from .calculations import oblicz_fala_b
from .config import ConfigManager, DEFAULT_MARGIN_RULES
from .printing import (
    PrinterError,
    build_summary_csv,
    print_text_document,
)


WAVE_TABS = [
    "FALA B",
    "FALA E",
    "FALA C+EB",
    "FALA EB+B 203",
    "FALA BC",
    "F203 FALA BC",
    "F200 B+ EB",
    "F200 BC",
]


class CalculatorTab(ttk.Frame):
    """Pojedyncza zakładka kalkulatora odpowiadająca konkretnej fali."""

    def __init__(self, master: ttk.Notebook, app: "FalaBApp", wave_name: str):
        super().__init__(master)
        self.app = app
        self.wave_name = wave_name
        self.last_results: Dict[str, Any] = {}

        self._init_variables()
        self._build_ui()

    # ------------------------------------------------------------------
    # Budowanie interfejsu użytkownika
    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1, uniform="col")
        self.columnconfigure(1, weight=1, uniform="col")

        frame_client = ttk.LabelFrame(self, text="Dane klienta")
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
        ttk.Label(frame_client, text="E-mail").grid(
            row=1, column=2, sticky="w", pady=(6, 0)
        )
        ttk.Entry(frame_client, textvariable=self.var_client_email).grid(
            row=1, column=3, sticky="we", pady=(6, 0)
        )

        frame_inputs = ttk.LabelFrame(self, text="Parametry kartonu i nakłady")
        frame_inputs.grid(row=1, column=0, sticky="nsew", padx=(0, 10), pady=(0, 8))
        for col in (1, 3, 5):
            frame_inputs.columnconfigure(col, weight=1)

        bold_font = ("TkDefaultFont", 10, "bold")
        ttk.Label(frame_inputs, text="DŁUGOŚĆ (mm)", font=bold_font).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Entry(frame_inputs, textvariable=self.var_dl, width=10).grid(
            row=0, column=1, sticky="we", padx=(0, 6)
        )
        ttk.Label(frame_inputs, text="SZEROKOŚĆ (mm)", font=bold_font).grid(
            row=0, column=2, sticky="w", padx=(6, 0)
        )
        ttk.Entry(frame_inputs, textvariable=self.var_sz, width=10).grid(
            row=0, column=3, sticky="we", padx=(0, 6)
        )
        ttk.Label(frame_inputs, text="WYSOKOŚĆ (mm)", font=bold_font).grid(
            row=0, column=4, sticky="w"
        )
        ttk.Entry(frame_inputs, textvariable=self.var_wys, width=10).grid(
            row=0, column=5, sticky="we"
        )

        ttk.Label(frame_inputs, text="Gramatura [g/m²]").grid(
            row=1, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Entry(frame_inputs, textvariable=self.var_gram, width=10).grid(
            row=1, column=1, sticky="we", padx=(0, 6), pady=(6, 0)
        )
        ttk.Label(frame_inputs, text="Cena surowca 1 m² [zł]").grid(
            row=1, column=2, sticky="w", padx=(6, 0), pady=(6, 0)
        )
        ttk.Entry(frame_inputs, textvariable=self.var_cena_m2, width=10).grid(
            row=1, column=3, sticky="we", padx=(0, 6), pady=(6, 0)
        )

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

        frame_right = ttk.Frame(self)
        frame_right.grid(row=1, column=1, sticky="nsew", pady=(0, 8))
        frame_right.columnconfigure(0, weight=1)
        frame_right.rowconfigure(0, weight=1)

        frame_bigi = ttk.LabelFrame(frame_right, text="Bigowanie")
        frame_bigi.grid(row=0, column=0, sticky="nsew")
        frame_bigi.columnconfigure(0, weight=1)
        frame_bigi.columnconfigure(1, weight=1)

        ttk.Label(frame_bigi, textvariable=self.var_bigi_row1, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(frame_bigi, textvariable=self.var_bigi_row2, justify="left").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        ttk.Label(frame_bigi, textvariable=self.var_bigowanie_row1, justify="left").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        ttk.Label(frame_bigi, textvariable=self.var_bigowanie_row2, justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )
        ttk.Label(frame_bigi, textvariable=self.var_formatka_dims, justify="left").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

        frame_costs = ttk.LabelFrame(self, text="Koszty dodatkowe i transport")
        frame_costs.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        for col in (1, 3):
            frame_costs.columnconfigure(col, weight=1)

        ttk.Label(frame_costs, text="Dodatkowe koszty [zł/partia]").grid(
            row=0, column=0, sticky="w"
        )
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

        frame_actions = ttk.Frame(self)
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

        frame_results = ttk.LabelFrame(self, text="Wyniki")
        frame_results.grid(row=4, column=0, columnspan=2, sticky="nsew")
        frame_results.columnconfigure(0, weight=1)

        ttk.Label(frame_results, textvariable=self.var_costs, justify="left").grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        ttk.Label(frame_results, textvariable=self.var_transport_info, justify="left").grid(
            row=1, column=0, sticky="w", pady=(0, 8)
        )

        self.rowconfigure(4, weight=1)

    # ------------------------------------------------------------------
    # Logika kalkulatora
    # ------------------------------------------------------------------
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
            dodatkowe = self._parse_float_optional(self.var_inne, "Dodatkowe koszty")
            stawka_km = self._parse_float_optional(
                self.var_transport_stawka, "Stawka transport"
            )
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
                "fala": self.wave_name,
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
            "margin_rules": self.app.config.get_margin_rules(),
        }

    def print_summary(self) -> None:
        try:
            summary = build_summary_csv(
                self.last_results,
                fallback_margin_rules=self.app.config.get_margin_rules(),
            )
        except ValueError as exc:
            messagebox.showinfo("Brak danych", str(exc))
            return
        except Exception as exc:  # pragma: no cover - defensywnie
            messagebox.showerror(
                "Błąd",
                f"Nie udało się przygotować danych do wydruku.\n{exc}",
            )
            return

        try:
            print_text_document(summary, suffix=".csv", prefer_notepad=True)
        except PrinterError as exc:
            messagebox.showerror("Błąd drukowania", str(exc))
            return
        messagebox.showinfo("Drukowanie", "Podsumowanie zostało wysłane do drukarki.")


class FalaBApp(ttk.Frame):
    """Główne okno aplikacji kalkulatora."""

    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=12)
        self.master.title("Kalkulator Rekruso")
        self.grid(sticky="nsew")

        self.config = ConfigManager()
        self.margin_rules: list[dict[str, float]] = self.config.get_margin_rules()
        self.settings_unlocked = False
        self.calculator_tabs: dict[str, CalculatorTab] = {}

        self.create_widgets()
        self.master.bind("<Return>", self._handle_return)

    # ------------------------------------------------------------------
    # Budowanie interfejsu użytkownika
    # ------------------------------------------------------------------
    def create_widgets(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        for wave_name in WAVE_TABS:
            tab = CalculatorTab(self.notebook, self, wave_name)
            self.notebook.add(tab, text=wave_name)
            self.calculator_tabs[str(tab)] = tab

        self.tab_settings = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_settings, text="Ustawienia")

        self._build_settings_tab()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _build_settings_tab(self) -> None:
        tab = self.tab_settings
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(0, weight=1)

        self.settings_locked_frame = ttk.Frame(tab, padding=20)
        self.settings_locked_frame.grid(row=0, column=0, sticky="nsew")
        self.settings_locked_frame.columnconfigure(1, weight=1)

        self.settings_content_frame = ttk.Frame(tab, padding=20)
        self.settings_content_frame.grid(row=0, column=0, sticky="nsew")
        for col in range(4):
            self.settings_content_frame.columnconfigure(col, weight=1)

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
        for col in range(3):
            button_frame.columnconfigure(col, weight=1)

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

    # ------------------------------------------------------------------
    # Obsługa zakładki ustawień
    # ------------------------------------------------------------------
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

    # ------------------------------------------------------------------
    # Operacje na konfiguracji marży
    # ------------------------------------------------------------------
    def _refresh_margin_tree(self) -> None:
        for item in self.margin_tree.get_children():
            self.margin_tree.delete(item)
        for index, rule in enumerate(self.margin_rules):
            self.margin_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(rule.get("max_quantity"), rule.get("margin_percent")),
            )

    def _on_margin_tree_select(self, _event: tk.Event) -> None:
        selection = self.margin_tree.selection()
        if not selection:
            return
        item_id = selection[0]
        try:
            index = int(item_id)
        except ValueError:
            return
        if not 0 <= index < len(self.margin_rules):
            return
        rule = self.margin_rules[index]
        self.var_margin_quantity.set(str(rule.get("max_quantity", "")))
        self.var_margin_value.set(str(rule.get("margin_percent", "")))

    def _add_or_update_margin_rule(self) -> None:
        quantity_text = self.var_margin_quantity.get().strip()
        margin_text = self.var_margin_value.get().strip()

        try:
            quantity = int(float(quantity_text))
            margin = float(margin_text.replace(",", "."))
        except ValueError:
            self._set_margin_message("Podaj poprawne wartości liczbowe.", error=True)
            return

        existing_index = None
        for idx, rule in enumerate(self.margin_rules):
            if int(rule.get("max_quantity", -1)) == quantity:
                existing_index = idx
                break

        new_rule = {"max_quantity": quantity, "margin_percent": margin}
        if existing_index is not None:
            self.margin_rules[existing_index] = new_rule
        else:
            self.margin_rules.append(new_rule)
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
            "Przywracanie",
            "Czy na pewno chcesz przywrócić domyślne wartości marży?",
        ):
            return
        self.margin_rules = [rule.copy() for rule in DEFAULT_MARGIN_RULES]
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

    # ------------------------------------------------------------------
    # Integracja z zakładkami kalkulatora
    # ------------------------------------------------------------------
    def _get_current_calculator_tab(self) -> CalculatorTab | None:
        current = self.notebook.select()
        return self.calculator_tabs.get(current)

    def _handle_return(self, _event: tk.Event) -> None:
        tab = self._get_current_calculator_tab()
        if tab is not None:
            tab.policz()


def main() -> None:
    root = tk.Tk()
    try:
        from ctypes import windll

        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # pragma: no cover - zależne od systemu
        pass
    root.geometry("980x640")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    FalaBApp(root)
    root.mainloop()


__all__ = ["FalaBApp", "main"]
