import os
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from tkinter import messagebox, ttk
from typing import Any, Dict


def _excel_fixed(value: float, digits: int) -> float:
    """Replikacja działania funkcji FIXED z Excela."""
    if digits < 0:
        raise ValueError("Liczba miejsc po przecinku nie może być ujemna.")
    quant = Decimal("1").scaleb(-digits)
    return float(Decimal(str(value)).quantize(quant, rounding=ROUND_HALF_UP))

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
        ttk.Label(frame_client, text="E-mail").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(frame_client, textvariable=self.var_client_email).grid(
            row=1, column=3, sticky="we", pady=(6, 0)
        )

        frame_inputs = ttk.LabelFrame(self, text="Parametry kartonu i nakłady")
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

        frame_right = ttk.Frame(self)
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

        frame_costs = ttk.LabelFrame(self, text="Koszty dodatkowe i transport")
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

        csv_lines = ["sep=;", "Sekcja;Parametr;Wartość"]
        csv_lines.extend(";".join(row) for row in rows)
        return "\n".join(csv_lines)

    @staticmethod
    def _send_to_printer(text: str) -> None:
        temp_path = ""
        with tempfile.NamedTemporaryFile(
            "w", delete=False, encoding="utf-8", suffix=".csv"
        ) as temp_file:
            temp_file.write(text)
            temp_path = temp_file.name

        def _cleanup(path: str) -> None:
            try:
                os.remove(path)
            except OSError:
                pass

        try:
            if sys.platform.startswith("win"):
                if hasattr(os, "startfile"):
                    os.startfile(temp_path, "print")  # type: ignore[attr-defined]
                else:
                    raise RuntimeError("Brak obsługi drukowania w tym systemie.")
            elif sys.platform == "darwin":
                subprocess.run(["lp", temp_path], check=True)
            else:
                subprocess.run(["lpr", temp_path], check=True)
        except FileNotFoundError as exc:
            raise RuntimeError("Nie znaleziono polecenia drukarki w systemie.") from exc
        except Exception as exc:
            raise RuntimeError("Nie udało się wysłać danych do drukarki.") from exc
        finally:
            if temp_path:
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
