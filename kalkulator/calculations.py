"""Moduł odpowiedzialny za obliczenia dla fali B."""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Dict


def excel_fixed(value: float, digits: int) -> float:
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
    """Przelicza wszystkie zależności z arkusza "FALA B"."""

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
    zuzycie_m2 = excel_fixed((formatka_c11 * wymiar_zewnetrzny_e11) / 1_000_000.0, 3)
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


__all__ = ["excel_fixed", "oblicz_fala_b"]
