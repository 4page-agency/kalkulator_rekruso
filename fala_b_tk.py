import tkinter as tk
from decimal import Decimal, ROUND_HALF_UP
from tkinter import messagebox, ttk
from typing import Dict, List


def _excel_fixed(value: float, digits: int) -> float:
    """Replikacja działania funkcji FIXED z Excela."""
    if digits < 0:
        raise ValueError("Liczba miejsc po przecinku nie może być ujemna.")
    quant = Decimal("1").scaleb(-digits)
    return float(Decimal(str(value)).quantize(quant, rounding=ROUND_HALF_UP))


def _safe_division(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else 0.0


def _ensure_len(values: List[float], size: int, fill: float = 0.0) -> List[float]:
    out = list(values)
    if len(out) < size:
        out.extend([fill] * (size - len(out)))
    return out[:size]


def oblicz_fala_b(
    dl: float,
    sz: float,
    wys: float,
    gramatura: float,
    cena_m2: float,
    naklad1: int,
    naklad2: int,
    naklad3: int,
    stale: List[float],
    nasza: List[float],
    kolory: List[float],
    sloter_total: float,
    rotacja_total: float,
    druk_total: float,
    inne_total: float,
    klejenie_na_szt: float,
    stawka_transport_km: float,
    dystans_km: float,
    podatek_multiplier: float = 0.38,
    slot_klejenie_proc: float = 0.35,
    transport_powrot: bool = True,
) -> Dict:
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
    koszt_sklejenia_na_szt = koszt_mat_na_szt * slot_klejenie_proc

    # --- MINIMUM PRODUKCYJNE I WERYFIKACJA ---
    min_aq = 500.0 / formatka_c11 * 1000.0 if formatka_c11 else 0.0
    min_con = 300.0 / zuzycie_m2 if zuzycie_m2 else 0.0
    min_pg = 500.0 / zuzycie_m2 if zuzycie_m2 else 0.0

    weryfikacja_dl = i8 + 3.0
    weryfikacja_sz = h8 + 3.0
    weryfikacja_wys = d8 + 2.0

    # --- NAKŁADY I LISTY WEJŚCIOWE ---
    naklady = _ensure_len([int(naklad1), int(naklad2), int(naklad3)], 3, 0)
    stale = _ensure_len(stale, 3, 0.0)
    nasza = _ensure_len(nasza, 3, 0.0)
    kolory = _ensure_len(kolory, 3, 0.0)
    podatek = [nasza[i] * podatek_multiplier for i in range(3)]

    # --- TRANSPORT ---
    stawka_pelna = stawka_transport_km * (2.0 if transport_powrot else 1.0)
    koszt_transport_calk = stawka_pelna * max(dystans_km, 0.0)

    # --- SUMY DLA KAŻDEGO NAKŁADU ---
    total_operations = [sloter_total, rotacja_total, druk_total, inne_total]
    wyniki_nakladow = []
    for idx, naklad in enumerate(naklady):
        zuzycie_total = zuzycie_m2 * naklad
        koszt_mat_total = koszt_mat_na_szt * naklad
        transport_na_szt = _safe_division(koszt_transport_calk, naklad)
        operacje_na_szt = sum(_safe_division(val, naklad) for val in total_operations)
        marza_suma = (
            stale[idx]
            + nasza[idx]
            + podatek[idx]
            + kolory[idx]
            + operacje_na_szt
            + klejenie_na_szt
            + transport_na_szt
        )
        koszt_mat_jednostkowy = _safe_division(koszt_mat_total, naklad)
        cena_opakowania = marza_suma + koszt_mat_jednostkowy
        wyniki_nakladow.append(
            {
                "naklad": naklad,
                "zuzycie_m2_total": zuzycie_total,
                "koszt_materialu_total": koszt_mat_total,
                "koszt_materialu_na_szt": koszt_mat_jednostkowy,
                "marza_suma_na_szt": marza_suma,
                "transport_na_szt": transport_na_szt,
                "cena_opakowania_na_szt": cena_opakowania,
                "cena_opakowania_total": cena_opakowania * naklad,
                "stala": stale[idx],
                "nasza": nasza[idx],
                "podatek": podatek[idx],
                "kolor": kolory[idx],
            }
        )

    return {
        "bigi": {"c8": c8, "d8": d8, "e8": e8},
        "bigowe": {"f8": f8, "g8": g8, "h8": h8, "i8": i8, "j8": j8},
        "sumy_bigowe": {"c9": c9, "d9": d9, "e9": e9, "f9": f9, "g9": g9, "h9": h9, "i9": i9, "j9": j9},
        "formatka_mm": formatka_c11,
        "wymiar_zewnetrzny_mm": wymiar_zewnetrzny_e11,
        "zuzycie_m2_na_szt": zuzycie_m2,
        "waga_kg_na_szt": waga_kg,
        "koszt_mat_na_szt": koszt_mat_na_szt,
        "koszt_sklejenia_na_szt": koszt_sklejenia_na_szt,
        "slot_klejenie_proc": slot_klejenie_proc,
        "minimum_produkcji": {"aq": min_aq, "con": min_con, "pg": min_pg},
        "weryfikacja_zewnetrzna": {"dl": weryfikacja_dl, "sz": weryfikacja_sz, "wys": weryfikacja_wys},
        "naklady": wyniki_nakladow,
        "transport": {
            "stawka_pelna": stawka_pelna,
            "koszt_calkowity": koszt_transport_calk,
            "dystans": dystans_km,
            "powrot": transport_powrot,
        },
        "inne_parametry": {
            "podatek_multiplier": podatek_multiplier,
            "sloter_total": sloter_total,
            "rotacja_total": rotacja_total,
            "druk_total": druk_total,
            "inne_total": inne_total,
            "klejenie_na_szt": klejenie_na_szt,
        },
    }


class FalaBApp(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=12)
        self.master.title("Kalkulator – FALA B (handlowiec)")
        self.grid(sticky="nsew")
        self._init_variables()
        self.create_widgets()
        self.master.bind("<Return>", lambda _event: self.policz())

    def _init_variables(self) -> None:
        self.var_dl = tk.StringVar(value="400")
        self.var_sz = tk.StringVar(value="300")
        self.var_wys = tk.StringVar(value="200")
        self.var_gram = tk.StringVar(value="675")
        self.var_cena_m2 = tk.StringVar(value="1.11")

        self.vars_naklad = [tk.StringVar(value=wartosc) for wartosc in ("2000", "1500", "500")]

        self.vars_stale = [tk.StringVar(value="0.02") for _ in range(3)]
        self.vars_nasza = [tk.StringVar(value=wartosc) for wartosc in ("0.93", "0.25", "0.28")]
        self.vars_kolor = [tk.StringVar(value="0.0") for _ in range(3)]
        self.vars_podatek = [tk.StringVar(value="0.0") for _ in range(3)]

        self.var_podatek_multiplier = tk.StringVar(value="0.38")
        self.var_slot_klejenie_proc = tk.StringVar(value="0.35")

        self.var_sloter = tk.StringVar(value="40")
        self.var_rotacja = tk.StringVar(value="0")
        self.var_druk = tk.StringVar(value="0")
        self.var_inne = tk.StringVar(value="0")
        self.var_klejenie_szt = tk.StringVar(value="0.03")

        self.var_transport_stawka = tk.StringVar(value="2.5")
        self.var_transport_km = tk.StringVar(value="0")
        self.var_transport_powrot = tk.BooleanVar(value=True)

        self.var_machine = tk.StringVar(value="–")
        self.var_minimum = tk.StringVar(value="–")
        self.var_weryfikacja = tk.StringVar(value="–")
        self.var_costs = tk.StringVar(value="–")
        self.var_transport_info = tk.StringVar(value="–")

        self.table_params = (
            "Nakład [szt.]",
            "Zużycie m² (razem)",
            "Koszt materiału (zł)",
            "Koszt materiału / szt. (zł)",
            "Transport / szt. (zł)",
            "Marża suma / szt. (zł)",
            "Cena opakowania / szt. (zł)",
            "Cena partii (zł)",
        )
        self.table_vars: Dict[str, List[tk.StringVar]] = {
            param: [tk.StringVar(value="–") for _ in range(3)]
            for param in self.table_params
        }

    def create_widgets(self) -> None:
        self.columnconfigure(0, weight=1, uniform="col")
        self.columnconfigure(1, weight=1, uniform="col")

        frame_inputs = ttk.LabelFrame(self, text="Parametry kartonu i nakłady")
        frame_inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=(0, 8))
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
        ttk.Label(frame_inputs, text="Nakłady [szt.]").grid(row=3, column=0, columnspan=6, sticky="w")
        ttk.Label(frame_inputs, text="Nakład 1").grid(row=4, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(frame_inputs, textvariable=self.vars_naklad[0], width=10).grid(row=4, column=1, sticky="we", padx=(0, 6), pady=(4, 0))
        ttk.Label(frame_inputs, text="Nakład 2").grid(row=4, column=2, sticky="w", padx=(6, 0), pady=(4, 0))
        ttk.Entry(frame_inputs, textvariable=self.vars_naklad[1], width=10).grid(row=4, column=3, sticky="we", padx=(0, 6), pady=(4, 0))
        ttk.Label(frame_inputs, text="Nakład 3").grid(row=4, column=4, sticky="w", padx=(6, 0), pady=(4, 0))
        ttk.Entry(frame_inputs, textvariable=self.vars_naklad[2], width=10).grid(row=4, column=5, sticky="we")

        frame_costs = ttk.LabelFrame(self, text="Składowe ceny za 1 szt.")
        frame_costs.grid(row=0, column=1, sticky="nsew", pady=(0, 8))
        frame_costs.columnconfigure(0, weight=1)
        for col in range(1, 4):
            frame_costs.columnconfigure(col, weight=1)

        headers = ["Parametr", "Nakład 1", "Nakład 2", "Nakład 3"]
        for col, text in enumerate(headers):
            ttk.Label(frame_costs, text=text, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=col, sticky="we", padx=2, pady=(0, 4))

        rows = [
            ("Stała (C20/D20/E20)", self.vars_stale),
            ("Nasza (C21/D21/E21)", self.vars_nasza),
            ("Kolor (C23/D23/E23)", self.vars_kolor),
        ]
        for r, (label, vars_row) in enumerate(rows, start=1):
            ttk.Label(frame_costs, text=label).grid(row=r, column=0, sticky="w", padx=2, pady=2)
            for idx, var in enumerate(vars_row):
                ttk.Entry(frame_costs, textvariable=var, width=10).grid(row=r, column=idx + 1, sticky="we", padx=2, pady=2)

        ttk.Label(frame_costs, text="Podatek = A24 × Nasza").grid(row=4, column=0, sticky="w", padx=2, pady=(8, 2))
        for idx, var in enumerate(self.vars_podatek):
            ttk.Label(frame_costs, textvariable=var, foreground="#555555").grid(row=4, column=idx + 1, sticky="we", padx=2, pady=(8, 2))

        ttk.Label(frame_costs, text="Mnożnik podatku (A24)").grid(row=5, column=0, sticky="w", padx=2, pady=(6, 0))
        ttk.Entry(frame_costs, textvariable=self.var_podatek_multiplier, width=10).grid(row=5, column=1, sticky="we", padx=2, pady=(6, 0))

        ttk.Label(frame_costs, text="S+K [% od materiału] (A19)").grid(row=6, column=0, sticky="w", padx=2, pady=(6, 0))
        ttk.Entry(frame_costs, textvariable=self.var_slot_klejenie_proc, width=10).grid(row=6, column=1, sticky="we", padx=2, pady=(6, 0))

        frame_ops = ttk.LabelFrame(self, text="Operacje i transport")
        frame_ops.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        for col in range(0, 6):
            frame_ops.columnconfigure(col, weight=1)

        ttk.Label(frame_ops, text="SLOTER (C25) [zł/partia]").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_sloter, width=10).grid(row=0, column=1, sticky="we", padx=(0, 8))
        ttk.Label(frame_ops, text="ROTACJA (D25)").grid(row=0, column=2, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_rotacja, width=10).grid(row=0, column=3, sticky="we", padx=(0, 8))
        ttk.Label(frame_ops, text="Inne (E25/F25) [zł/partia]").grid(row=0, column=4, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_druk, width=10).grid(row=0, column=5, sticky="we")

        ttk.Label(frame_ops, text="Dodatkowe koszty (F25)").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frame_ops, textvariable=self.var_inne, width=10).grid(row=1, column=1, sticky="we", padx=(0, 8), pady=(6, 0))
        ttk.Label(frame_ops, text="KLEJENIE / szt. (G25)").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(frame_ops, textvariable=self.var_klejenie_szt, width=10).grid(row=1, column=3, sticky="we", padx=(0, 8), pady=(6, 0))

        ttk.Separator(frame_ops).grid(row=2, column=0, columnspan=6, sticky="we", pady=8)
        ttk.Label(frame_ops, text="Transport – stawka zł/km").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_transport_stawka, width=10).grid(row=3, column=1, sticky="we", padx=(0, 8))
        ttk.Label(frame_ops, text="Dystans [km] (C28)").grid(row=3, column=2, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_transport_km, width=10).grid(row=3, column=3, sticky="we", padx=(0, 8))
        ttk.Checkbutton(frame_ops, text="Uwzględnij powrót", variable=self.var_transport_powrot).grid(row=3, column=4, columnspan=2, sticky="w")

        ttk.Button(self, text="Policz", command=self.policz).grid(row=2, column=0, columnspan=2, sticky="we", pady=4)

        frame_results = ttk.LabelFrame(self, text="Wyniki")
        frame_results.grid(row=3, column=0, columnspan=2, sticky="nsew")
        frame_results.columnconfigure(0, weight=1)

        ttk.Label(frame_results, textvariable=self.var_machine, justify="left").grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(frame_results, textvariable=self.var_minimum, justify="left").grid(row=1, column=0, sticky="w", pady=(0, 4))
        ttk.Label(frame_results, textvariable=self.var_weryfikacja, justify="left").grid(row=2, column=0, sticky="w", pady=(0, 4))
        ttk.Label(frame_results, textvariable=self.var_costs, justify="left").grid(row=3, column=0, sticky="w", pady=(0, 4))
        ttk.Label(frame_results, textvariable=self.var_transport_info, justify="left").grid(row=4, column=0, sticky="w", pady=(0, 8))

        table_frame = ttk.Frame(frame_results)
        table_frame.grid(row=5, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=2)
        for col in range(1, 4):
            table_frame.columnconfigure(col, weight=1)

        for col, heading in enumerate(["Parametr", "Nakład 1", "Nakład 2", "Nakład 3"]):
            ttk.Label(table_frame, text=heading, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=col, sticky="we", padx=2, pady=(0, 4))

        for r, param in enumerate(self.table_params, start=1):
            ttk.Label(table_frame, text=param).grid(row=r, column=0, sticky="w", padx=2, pady=2)
            for idx, var in enumerate(self.table_vars[param]):
                ttk.Label(table_frame, textvariable=var).grid(row=r, column=idx + 1, sticky="we", padx=2, pady=2)

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

    def _parse_int(self, var: tk.StringVar, name: str) -> int:
        text = str(var.get()).strip()
        if not text:
            return 0
        try:
            return int(float(text.replace(",", ".")))
        except ValueError as exc:
            raise ValueError(f"Nieprawidłowa wartość w polu: {name}") from exc

    def _format_number(self, value: float, decimals: int = 2, placeholder: str = "–") -> str:
        if value is None:
            return placeholder
        try:
            return f"{value:.{decimals}f}"
        except (TypeError, ValueError):
            return placeholder

    def policz(self) -> None:
        try:
            dl = self._parse_float(self.var_dl, "DŁ")
            sz = self._parse_float(self.var_sz, "SZ")
            wys = self._parse_float(self.var_wys, "WYS")
            gram = self._parse_float(self.var_gram, "Gramatura")
            cena_m2 = self._parse_float(self.var_cena_m2, "Cena 1 m²")

            naklady = [self._parse_int(var, f"Nakład {idx + 1}") for idx, var in enumerate(self.vars_naklad)]

            stale = [self._parse_float_optional(var, f"Stała {idx + 1}") for idx, var in enumerate(self.vars_stale)]
            nasza = [self._parse_float_optional(var, f"Nasza {idx + 1}") for idx, var in enumerate(self.vars_nasza)]
            kolory = [self._parse_float_optional(var, f"Kolor {idx + 1}") for idx, var in enumerate(self.vars_kolor)]

            podatek_multiplier = self._parse_float_optional(self.var_podatek_multiplier, "Mnożnik podatku")
            slot_proc = self._parse_float_optional(self.var_slot_klejenie_proc, "S+K [%]")

            sloter = self._parse_float_optional(self.var_sloter, "SLOTER")
            rotacja = self._parse_float_optional(self.var_rotacja, "ROTACJA")
            druk = self._parse_float_optional(self.var_druk, "Dodatkowe koszty (E25)")
            inne = self._parse_float_optional(self.var_inne, "Dodatkowe koszty (F25)")
            klejenie_szt = self._parse_float_optional(self.var_klejenie_szt, "KLEJENIE / szt.")

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
            naklad1=naklady[0],
            naklad2=naklady[1],
            naklad3=naklady[2],
            stale=stale,
            nasza=nasza,
            kolory=kolory,
            sloter_total=sloter,
            rotacja_total=rotacja,
            druk_total=druk,
            inne_total=inne,
            klejenie_na_szt=klejenie_szt,
            stawka_transport_km=stawka_km,
            dystans_km=dystans,
            podatek_multiplier=podatek_multiplier,
            slot_klejenie_proc=slot_proc,
            transport_powrot=powrot,
        )

        bigi = wyniki["bigi"]
        bigowe = wyniki["bigowe"]
        self.var_machine.set(
            "\n".join(
                [
                    "Ustawienia maszyny:",
                    (
                        "Bigi: "
                        f"{bigi['c8']:.2f} mm | {bigi['d8']:.2f} mm | {bigi['e8']:.2f} mm"
                    ),
                    (
                        "Bigowe: "
                        f"{bigowe['f8']:.2f} mm | {bigowe['g8']:.2f} mm | {bigowe['h8']:.2f} mm | "
                        f"{bigowe['i8']:.2f} mm | {bigowe['j8']:.2f} mm"
                    ),
                    f"Formatka: {wyniki['formatka_mm']:.2f} mm",
                ]
            )
        )

        minimum = wyniki["minimum_produkcji"]
        self.var_minimum.set(
            "Minimum produkcyjne: "
            f"AQ (A11) {minimum['aq']:.0f} szt | "
            f"CON (A12) {minimum['con']:.0f} szt | "
            f"PG (A13) {minimum['pg']:.0f} szt"
        )

        weryf = wyniki["weryfikacja_zewnetrzna"]
        self.var_weryfikacja.set(
            "Weryfikacja wymiarów (H12/I12/J12): "
            f"dł {weryf['dl']:.2f} mm | sz {weryf['sz']:.2f} mm | wys {weryf['wys']:.2f} mm"
        )

        self.var_costs.set(
            "\n".join(
                [
                    f"Zużycie m²/szt. (G5): {wyniki['zuzycie_m2_na_szt']:.3f}",
                    f"Waga kg/szt. (J5): {wyniki['waga_kg_na_szt']:.3f}",
                    (
                        "Koszt materiału/szt. (A18): "
                        f"{wyniki['koszt_mat_na_szt']:.4f} zł"
                    ),
                    (
                        "S+K {wyniki['slot_klejenie_proc']*100:.1f}% → "
                        f"{wyniki['koszt_sklejenia_na_szt']:.4f} zł/szt."
                    ),
                ]
            )
        )

        transport = wyniki["transport"]
        powrot_txt = "tak" if transport["powrot"] else "nie"
        self.var_transport_info.set(
            f"Transport: stawka {transport['stawka_pelna']:.2f} zł/km, "
            f"dystans {transport['dystans']:.2f} km, powrót: {powrot_txt}. "
            f"Koszt łączny: {transport['koszt_calkowity']:.2f} zł"
        )

        for idx, dane in enumerate(wyniki["naklady"]):
            self.vars_podatek[idx].set(self._format_number(dane["podatek"], 6))

        for param in self.table_params:
            for var in self.table_vars[param]:
                var.set("–")

        for idx, dane in enumerate(wyniki["naklady"]):
            self.table_vars["Nakład [szt.]"][idx].set(f"{int(dane['naklad'])}")
            self.table_vars["Zużycie m² (razem)"][idx].set(self._format_number(dane["zuzycie_m2_total"], 3))
            self.table_vars["Koszt materiału (zł)"][idx].set(self._format_number(dane["koszt_materialu_total"], 2))
            self.table_vars["Koszt materiału / szt. (zł)"][idx].set(self._format_number(dane["koszt_materialu_na_szt"], 4))
            self.table_vars["Transport / szt. (zł)"][idx].set(self._format_number(dane["transport_na_szt"], 4))
            self.table_vars["Marża suma / szt. (zł)"][idx].set(self._format_number(dane["marza_suma_na_szt"], 4))
            self.table_vars["Cena opakowania / szt. (zł)"][idx].set(self._format_number(dane["cena_opakowania_na_szt"], 4))
            self.table_vars["Cena partii (zł)"][idx].set(self._format_number(dane["cena_opakowania_total"], 2))


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
