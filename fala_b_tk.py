
import math
import tkinter as tk
from tkinter import ttk, messagebox

# =============== LOGIKA OBLICZEŃ (zachowana jak poprzednio) ===============

def oblicz_fala_b(dl, sz, wys, gramatura, cena_m2, naklad1, naklad2, naklad3, marza_proc, koszt_transport_km, dystans_km, koszt_proc_sklejenie=0.35):
    """
    Zwraca słownik z wynikami obliczeń dla FALA B.
    """
    # Wiersz 8 – BIGI / BIGOWE
    c8 = (sz / 2.0) + 2.0
    d8 = wys + 10.0
    e8 = (sz / 2.0) + 2.0
    f8 = sz + 1.0
    g8 = dl + 3.0
    h8 = sz + 3.0
    i8 = dl + 3.0
    j8 = 35.0  # stała

    # Wiersz 9 – sumy narastająco
    c9 = c8
    d9 = c8 + d8
    e9 = c8 + d8 + e8
    f9 = f8
    g9 = f8 + g8
    h9 = f8 + g8 + h8
    i9 = f8 + g8 + h8 + i8
    j9 = f8 + g8 + h8 + i8 + j8

    # Wiersz 11 – formatka i wymiar zewnętrzny
    # C11 = (((C5 + D5) * 2) + 35 + 12) - 2
    formatka_c11 = (((dl + sz) * 2.0) + 35.0 + 12.0) - 2.0
    # E11 = SUM(C8:E8)
    wymiar_zewnetrzny_e11 = c8 + d8 + e8

    # G5 – zużycie m2/szt. (3 miejsca po przecinku jak w Excelu)
    zuzycie_m2 = round((formatka_c11 * wymiar_zewnetrzny_e11) / 1_000_000.0, 3)

    # J5 – waga opakowania kg/szt.
    waga_kg = (gramatura * zuzycie_m2) / 1000.0

    # Koszt materiału / szt.
    koszt_mat_szt = zuzycie_m2 * cena_m2

    # Sloter + klejenie (przyjęty współczynnik 0.35 od kosztu materiału)
    koszt_sklejenia = koszt_mat_szt * koszt_proc_sklejenie

    # Transport – prosty model: koszt = stawka * km; jeśli podano nakład,
    # rozbijamy koszt na 1 szt. przez (nakład1 + nakład2 + nakład3) > 0
    laczny_naklad = max(int(naklad1) + int(naklad2) + int(naklad3), 1)
    koszt_transport_calk = (koszt_transport_km * dystans_km)
    koszt_transport_szt = koszt_transport_calk / laczny_naklad

    # Suma techniczna (bez marży) / szt.
    koszt_techniczny = koszt_mat_szt + koszt_sklejenia + koszt_transport_szt

    # Marża
    cena_netto = koszt_techniczny * (1.0 + marza_proc / 100.0)

    # Minimum produkcyjne (z arkusza)
    # A11 – =500 / C11 * 1000
    min_prod_arkusz_1 = (500.0 / formatka_c11 * 1000.0) if formatka_c11 else 0.0
    # A12 – =300 / G5
    min_prod_arkusz_2 = (300.0 / zuzycie_m2) if zuzycie_m2 > 0 else 0.0

    # KLUCZE ZWRACANE — NAZYWAMY JEDNOZNACZNIE
    return {
        "formatka_mm": formatka_c11,
        "wymiar_zewnetrzny_mm": wymiar_zewnetrzny_e11,  # <— TEN KLUCZ
        "zuzycie_m2_na_szt": zuzycie_m2,
        "waga_kg_na_szt": waga_kg,
        "koszt_mat_na_szt": koszt_mat_szt,
        "koszt_sklejenia_na_szt": koszt_sklejenia,
        "koszt_transport_na_szt": koszt_transport_szt,
        "cena_sugerowana_netto": cena_netto,
        "min_prod_1": min_prod_arkusz_1,
        "min_prod_2": min_prod_arkusz_2,
    }

# =============== TKINTER UI ===============

class FalaBApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=12)
        self.master.title("Kalkulator – FALA B (handlowiec)")
        self.grid(sticky="nsew")
        self.create_widgets()
        self.master.bind("<Return>", lambda e: self.policz())

    def create_widgets(self):
        # Konfiguracja siatki
        for i in range(0, 8):
            self.columnconfigure(i, weight=1)
        for r in range(0, 20):
            self.rowconfigure(r, weight=0)

        # Sekcja wejścia
        lbl = ttk.Label(self, text="Wymiary kartonu [mm]")
        lbl.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0,6))

        self.var_dl = tk.StringVar(value="400")
        self.var_sz = tk.StringVar(value="300")
        self.var_wys = tk.StringVar(value="200")
        self.var_gram = tk.StringVar(value="675")
        self.var_cena_m2 = tk.StringVar(value="1.11")
        self.var_n1 = tk.StringVar(value="0")
        self.var_n2 = tk.StringVar(value="0")
        self.var_n3 = tk.StringVar(value="0")
        self.var_marza = tk.StringVar(value="20")
        self.var_km_stawka = tk.StringVar(value="2.5")
        self.var_km = tk.StringVar(value="0")

        row = 1
        ttk.Label(self, text="DŁ (C5)").grid(row=row, column=0, sticky="w"); ttk.Entry(self, textvariable=self.var_dl, width=10).grid(row=row, column=1, sticky="we")
        ttk.Label(self, text="SZ (D5)").grid(row=row, column=2, sticky="w"); ttk.Entry(self, textvariable=self.var_sz, width=10).grid(row=row, column=3, sticky="we")

        row += 1
        ttk.Label(self, text="WYS (E5)").grid(row=row, column=0, sticky="w"); ttk.Entry(self, textvariable=self.var_wys, width=10).grid(row=row, column=1, sticky="we")
        ttk.Label(self, text="Gramatura g/m² (H5)").grid(row=row, column=2, sticky="w"); ttk.Entry(self, textvariable=self.var_gram, width=10).grid(row=row, column=3, sticky="we")

        row += 1
        ttk.Label(self, text="Cena surowca 1 m² (I5) [PLN]").grid(row=row, column=0, sticky="w"); ttk.Entry(self, textvariable=self.var_cena_m2, width=10).grid(row=row, column=1, sticky="we")
        ttk.Label(self, text="Marża [%]").grid(row=row, column=2, sticky="w"); ttk.Entry(self, textvariable=self.var_marza, width=10).grid(row=row, column=3, sticky="we")

        row += 1
        ttk.Label(self, text="Nakład 1 (C15)").grid(row=row, column=0, sticky="w"); ttk.Entry(self, textvariable=self.var_n1, width=10).grid(row=row, column=1, sticky="we")
        ttk.Label(self, text="Nakład 2 (D15)").grid(row=row, column=2, sticky="w"); ttk.Entry(self, textvariable=self.var_n2, width=10).grid(row=row, column=3, sticky="we")

        row += 1
        ttk.Label(self, text="Nakład 3 (E15)").grid(row=row, column=0, sticky="w"); ttk.Entry(self, textvariable=self.var_n3, width=10).grid(row=row, column=1, sticky="we")
        ttk.Label(self, text="Transport [zł/km] • km").grid(row=row, column=2, sticky="w")
        frm_tr = ttk.Frame(self)
        frm_tr.grid(row=row, column=3, sticky="we")
        ttk.Entry(frm_tr, textvariable=self.var_km_stawka, width=7).pack(side="left", padx=(0,6))
        ttk.Entry(frm_tr, textvariable=self.var_km, width=7).pack(side="left")

        # Przycisk
        row += 1
        ttk.Button(self, text="Policz", command=self.policz).grid(row=row, column=0, columnspan=4, sticky="we", pady=8)

        # Wyniki
        sep = ttk.Separator(self); sep.grid(row=row+1, column=0, columnspan=4, sticky="we", pady=4)

        row += 2
        self.txt_wymiary = ttk.Label(self, text="Wymiary zewnętrzne (E11): – mm   |  Formatka (C11): – mm")
        self.txt_wymiary.grid(row=row, column=0, columnspan=4, sticky="w")

        row += 1
        self.txt_min = ttk.Label(self, text="Minimum produkcyjne: A11 –, A12 – [szt]")
        self.txt_min.grid(row=row, column=0, columnspan=4, sticky="w")

        row += 1
        self.txt_koszty = ttk.Label(self, text="Koszt materiału/szt.: – zł  |  SLOTER+KLEJENIE/szt.: – zł  |  Transport/szt.: – zł")
        self.txt_koszty.grid(row=row, column=0, columnspan=4, sticky="w")

        row += 1
        self.txt_cena = ttk.Label(self, text="Cena sugerowana netto/szt.: – zł")
        self.txt_cena.grid(row=row, column=0, columnspan=4, sticky="w")

    def _parse_float(self, var, name):
        try:
            return float(str(var.get()).replace(",", "."))
        except Exception:
            raise ValueError(f"Nieprawidłowa wartość w polu: {name}")

    def _parse_int(self, var, name):
        try:
            return int(float(str(var.get()).replace(",", ".")))
        except Exception:
            raise ValueError(f"Nieprawidłowa wartość w polu: {name}")

    def policz(self):
        try:
            dl = self._parse_float(self.var_dl, "DŁ")
            sz = self._parse_float(self.var_sz, "SZ")
            wys = self._parse_float(self.var_wys, "WYS")
            gram = self._parse_float(self.var_gram, "Gramatura")
            cena_m2 = self._parse_float(self.var_cena_m2, "Cena m2")
            n1 = self._parse_int(self.var_n1, "Nakład 1")
            n2 = self._parse_int(self.var_n2, "Nakład 2")
            n3 = self._parse_int(self.var_n3, "Nakład 3")
            marza = self._parse_float(self.var_marza, "Marża")
            km_st = self._parse_float(self.var_km_stawka, "Stawka km")
            km = self._parse_float(self.var_km, "km")
        except ValueError as e:
            messagebox.showerror("Błąd danych", str(e))
            return

        wyn = oblicz_fala_b(
            dl=dl, sz=sz, wys=wys,
            gramatura=gram, cena_m2=cena_m2,
            naklad1=n1, naklad2=n2, naklad3=n3,
            marza_proc=marza,
            koszt_transport_km=km_st, dystans_km=km
        )

        # Pobranie wartości z bezpiecznym .get (na wypadek starych wersji pliku)
        wym_zew = wyn.get("wymiar_zewnetrzny_mm", wyn.get("wymiar_zewnetrzny_e11", 0.0))
        formatka = wyn.get("formatka_mm", 0.0)
        min1 = wyn.get("min_prod_1", 0.0)
        min2 = wyn.get("min_prod_2", 0.0)
        koszt_mat = wyn.get("koszt_mat_na_szt", 0.0)
        koszt_skl = wyn.get("koszt_sklejenia_na_szt", 0.0)
        koszt_tr = wyn.get("koszt_transport_na_szt", 0.0)
        cena_netto = wyn.get("cena_sugerowana_netto", 0.0)

        # Aktualizacja etykiet
        self.txt_wymiary.configure(text=f"Wymiary zewnętrzne (E11): {wym_zew:.2f} mm   |  Formatka (C11): {formatka:.2f} mm")
        self.txt_min.configure(text=f"Minimum produkcyjne: A11 {min1:.0f} szt,  A12 {min2:.0f} szt")
        self.txt_koszty.configure(text=(
            f"Koszt materiału/szt.: {koszt_mat:.4f} zł  |  "
            f"SLOTER+KLEJENIE/szt.: {koszt_skl:.4f} zł  |  "
            f"Transport/szt.: {koszt_tr:.4f} zł"
        ))
        self.txt_cena.configure(text=f"Cena sugerowana netto/szt.: {cena_netto:.4f} zł")

def main():
    root = tk.Tk()
    # Styl
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)  # lepsze skalowanie na Windows
    except Exception:
        pass
    root.geometry("700x330")
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)
    app = FalaBApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
