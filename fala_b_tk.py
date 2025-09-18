import tkinter as tk
from decimal import Decimal, ROUND_HALF_UP
from tkinter import messagebox, ttk
from typing import Dict


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
    sloter_total: float,
    rotacja_total: float,
    druk_total: float,
    inne_total: float,
    klejenie_na_szt: float,
    stawka_transport_km: float,
    dystans_km: float,
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
        "koszt_sklejenia_na_szt": koszt_sklejenia_na_szt,
        "slot_klejenie_proc": slot_klejenie_proc,
        "minimum_produkcji": {"aq": min_aq, "con": min_con, "pg": min_pg},
        "weryfikacja_zewnetrzna": {"dl": weryfikacja_dl, "sz": weryfikacja_sz, "wys": weryfikacja_wys},
        "paletyzacja": {"dlugosc": paletyzacja_dlugosc, "szerokosc": paletyzacja_szerokosc},
        "transport": {
            "stawka_pelna": stawka_pelna,
            "koszt_calkowity": koszt_transport_calk,
            "dystans": dystans_km,
            "powrot": transport_powrot,
        },
        "inne_parametry": {
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

        self.var_slot_klejenie_proc = tk.StringVar(value="0.35")

        self.var_sloter = tk.StringVar(value="40")
        self.var_rotacja = tk.StringVar(value="0")
        self.var_druk = tk.StringVar(value="0")
        self.var_inne = tk.StringVar(value="0")
        self.var_klejenie_szt = tk.StringVar(value="0.03")

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
            value="Formatka – długość: – mm | Formatka – wysokość: – mm"
        )

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
        frame_right.grid(row=0, column=1, sticky="nsew", pady=(0, 8))
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

        frame_ops = ttk.LabelFrame(self, text="Operacje i transport")
        frame_ops.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 8))
        for col in range(0, 6):
            frame_ops.columnconfigure(col, weight=1)

        ttk.Label(frame_ops, text="SLOTER [zł/partia]").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_sloter, width=10).grid(row=0, column=1, sticky="we", padx=(0, 8))
        ttk.Label(frame_ops, text="ROTACJA [zł/partia]").grid(row=0, column=2, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_rotacja, width=10).grid(row=0, column=3, sticky="we", padx=(0, 8))
        ttk.Label(frame_ops, text="DRUK [zł/partia]").grid(row=0, column=4, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_druk, width=10).grid(row=0, column=5, sticky="we")

        ttk.Label(frame_ops, text="Pozostałe koszty [zł/partia]").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(frame_ops, textvariable=self.var_inne, width=10).grid(row=1, column=1, sticky="we", padx=(0, 8), pady=(6, 0))
        ttk.Label(frame_ops, text="KLEJENIE / szt.").grid(row=1, column=2, sticky="w", pady=(6, 0))
        ttk.Entry(frame_ops, textvariable=self.var_klejenie_szt, width=10).grid(row=1, column=3, sticky="we", padx=(0, 8), pady=(6, 0))
        ttk.Label(frame_ops, text="S+K [% od materiału]").grid(
            row=1, column=4, sticky="w", pady=(6, 0)
        )
        ttk.Entry(frame_ops, textvariable=self.var_slot_klejenie_proc, width=10).grid(
            row=1, column=5, sticky="we", pady=(6, 0)
        )

        ttk.Separator(frame_ops).grid(row=2, column=0, columnspan=6, sticky="we", pady=8)
        ttk.Label(frame_ops, text="Transport – stawka zł/km").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_transport_stawka, width=10).grid(row=3, column=1, sticky="we", padx=(0, 8))
        ttk.Label(frame_ops, text="Dystans [km]").grid(row=3, column=2, sticky="w")
        ttk.Entry(frame_ops, textvariable=self.var_transport_km, width=10).grid(row=3, column=3, sticky="we", padx=(0, 8))
        ttk.Checkbutton(frame_ops, text="Uwzględnij powrót", variable=self.var_transport_powrot).grid(row=3, column=4, columnspan=2, sticky="w")

        ttk.Button(self, text="Policz", command=self.policz).grid(row=2, column=0, columnspan=2, sticky="we", pady=4)

        frame_results = ttk.LabelFrame(self, text="Wyniki")
        frame_results.grid(row=3, column=0, columnspan=2, sticky="nsew")
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

            slot_proc = self._parse_float_optional(self.var_slot_klejenie_proc, "S+K [%]")

            sloter = self._parse_float_optional(self.var_sloter, "SLOTER")
            rotacja = self._parse_float_optional(self.var_rotacja, "ROTACJA")
            druk = self._parse_float_optional(self.var_druk, "Koszt druku")
            inne = self._parse_float_optional(self.var_inne, "Pozostałe koszty")
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
            sloter_total=sloter,
            rotacja_total=rotacja,
            druk_total=druk,
            inne_total=inne,
            klejenie_na_szt=klejenie_szt,
            stawka_transport_km=stawka_km,
            dystans_km=dystans,
            slot_klejenie_proc=slot_proc,
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
            f"Formatka – długość: {wyniki['formatka_mm']:.2f} mm | "
            f"Formatka – wysokość: {wyniki['wymiar_zewnetrzny_mm']:.2f} mm"
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

        self.var_costs.set(
            "\n".join(
                [
                    f"Zużycie m²/szt.: {wyniki['zuzycie_m2_na_szt']:.3f}",
                    f"Waga kg/szt.: {wyniki['waga_kg_na_szt']:.3f}",
                    (
                        "Koszt materiału/szt.: "
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
