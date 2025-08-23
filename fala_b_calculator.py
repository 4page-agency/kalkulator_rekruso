
"""
FALA B – kalkulator w Pythonie (wersja robocza)
------------------------------------------------
Odwzorowanie kluczowych obliczeń z arkusza "FALA B" na podstawie formuł odczytanych z Excela.

Wejścia (jednostki tak, jak w Excelu):
- dl (C5)  : DŁ (mm)
- sz (D5)  : SZ (mm)
- wys (E5) : WYS (mm)
- gramatura (H5): Surowiec / gramatura (g/m2)
- cena_m2 (I5): Cena surowca za 1 m2 (PLN)
- naklad1 (C15), naklad2 (D15), naklad3 (E15): sztuk
- wspolczynnik (A26): domyślnie 5 (z arkusza)
- koszt_transport_km (nagłówek w A27): 2.5 PLN/km (stała z arkusza)
- dystans_km (D28): ile km (domyślnie 0 w arkuszu)

Wyniki (wybrane):
- zuzycie_m2 (G5)
- waga_kg (J5)
- formatka (C11)
- wymiar_zewnetrzny (E11)
- koszty materiału dla nakładów (wiersz 17)
- pomocnicze: bigi/bigowe (wiersze 8-9)
"""

from dataclasses import dataclass

@dataclass
class FalaBInputs:
    dl: float              # C5 (mm)
    sz: float              # D5 (mm)
    wys: float             # E5 (mm)
    gramatura: float       # H5 (g/m2)
    cena_m2: float         # I5 (PLN/m2)
    naklad1: int = 0       # C15
    naklad2: int = 0       # D15
    naklad3: int = 0       # E15
    wspolczynnik: float = 5.0   # A26
    koszt_transport_km: float = 2.5  # A27 opis
    dystans_km: float = 0.0      # D28

@dataclass
class FalaBOutputs:
    c8: float; d8: float; e8: float; f8: float; g8: float; h8: float; i8: float; j8: float
    c9: float; d9: float; e9: float; f9: float; g9: float; h9: float; i9: float; j9: float
    formatka_c11: float
    wymiar_zewnetrzny_e11: float
    zuzycie_m2_g5: float
    waga_kg_j5: float
    koszt_mat_na_szt_a18: float
    sloter_klejenie_coef_a19: float
    sloter_klejenie_b19: float
    min_prod_a11: float
    min_prod_a12: float
    zuzycie_m2_n1: float
    zuzycie_m2_n2: float
    zuzycie_m2_n3: float
    koszt_mat_n1: float
    koszt_mat_n2: float
    koszt_mat_n3: float

def kalkuluj_fala_b(inp: FalaBInputs) -> FalaBOutputs:
    # Wiersz 8 – BIGI / BIGOWE
    c8 = (inp.sz / 2.0) + 2.0
    d8 = inp.wys + 10.0
    e8 = (inp.sz / 2.0) + 2.0
    f8 = inp.sz + 1.0
    g8 = inp.dl + 3.0
    h8 = inp.sz + 3.0
    i8 = inp.dl + 3.0
    j8 = 35.0  # stała w J8
    
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
    formatka_c11 = (((inp.dl + inp.sz) * 2.0) + 35.0 + 12.0) - 2.0
    # E11 = SUM(C8:E8)
    wymiar_zewnetrzny_e11 = c8 + d8 + e8
    
    # G5 – zużycie m2 = FIXED(((C11*E11)/1000000),3)
    zuzycie_m2_g5 = round((formatka_c11 * wymiar_zewnetrzny_e11) / 1_000_000.0, 3)
    
    # J5 – waga opakowania kg = H5*G5/1000
    waga_kg_j5 = (inp.gramatura * zuzycie_m2_g5) / 1000.0
    
    # Wiersz 18 – A18 = G5 * I5  (koszt materiału na sztukę?)
    koszt_mat_na_szt_a18 = zuzycie_m2_g5 * inp.cena_m2
    
    # Wiersz 19 – B19 = A18 * A19, gdzie A19 = 0.35 (odczyt w arkuszu)
    sloter_klejenie_coef_a19 = 0.35
    sloter_klejenie_b19 = koszt_mat_na_szt_a18 * sloter_klejenie_coef_a19
    
    # A11 – Minimum produkcyjne = 500 / C11 * 1000
    min_prod_a11 = 500.0 / formatka_c11 * 1000.0 if formatka_c11 else 0.0
    # A12 – = 300 / G5
    min_prod_a12 = 300.0 / zuzycie_m2_g5 if zuzycie_m2_g5 else 0.0
    
    # Wiersz 16 – zużycie m2 dla nakładów
    zuzycie_m2_n1 = zuzycie_m2_g5 * inp.naklad1
    zuzycie_m2_n2 = zuzycie_m2_g5 * inp.naklad2
    zuzycie_m2_n3 = zuzycie_m2_g5 * inp.naklad3
    
    # Wiersz 17 – koszt materiału dla nakładów
    koszt_mat_n1 = zuzycie_m2_g5 * inp.cena_m2 * inp.naklad1
    koszt_mat_n2 = zuzycie_m2_g5 * inp.cena_m2 * inp.naklad2
    koszt_mat_n3 = zuzycie_m2_g5 * inp.cena_m2 * inp.naklad3
    
    # TODO: MARŻA 1/2/3 (C19, D19, E19) – nie udało się odczytać formuł (prawdopodobnie procenty od A18 albo od całości).
    # TODO: "MARŻA SUMA" i "CENA OPAKOWANIA" (wiersz 30) – brak formuł widocznych w pliku.
    # TODO: TRANSPORT – w arkuszu jest opis 2,5 zł/km i komórka D28 (km). W Pythonie można to policzyć jako koszt_transport_km * dystans_km.
    
    return FalaBOutputs(
        c8=c8, d8=d8, e8=e8, f8=f8, g8=g8, h8=h8, i8=i8, j8=j8,
        c9=c9, d9=d9, e9=e9, f9=f9, g9=g9, h9=h9, i9=i9, j9=j9,
        formatka_c11=formatka_c11,
        wymiar_zewnetrzny_e11=wymiar_zewnetrzny_e11,
        zuzycie_m2_g5=zuzycie_m2_g5,
        waga_kg_j5=waga_kg_j5,
        koszt_mat_na_szt_a18=koszt_mat_na_szt_a18,
        sloter_klejenie_coef_a19=sloter_klejenie_coef_a19,
        sloter_klejenie_b19=sloter_klejenie_b19,
        min_prod_a11=min_prod_a11,
        min_prod_a12=min_prod_a12,
        zuzycie_m2_n1=zuzycie_m2_n1,
        zuzycie_m2_n2=zuzycie_m2_n2,
        zuzycie_m2_n3=zuzycie_m2_n3,
        koszt_mat_n1=koszt_mat_n1,
        koszt_mat_n2=koszt_mat_n2,
        koszt_mat_n3=koszt_mat_n3,
    )

if __name__ == "__main__":
    # Przykład użycia
    inp = FalaBInputs(
        dl=0, sz=0, wys=0,
        gramatura=675, cena_m2=1.11,
        naklad1=0, naklad2=0, naklad3=0
    )
    out = kalkuluj_fala_b(inp)
    from pprint import pprint
    pprint(out)
