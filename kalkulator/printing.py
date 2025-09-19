"""Funkcje związane z przygotowaniem wydruków."""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from typing import Any


class PrinterError(RuntimeError):
    """Wyjątek zgłaszany, gdy wysyłanie wydruku się nie powiedzie."""


def build_summary_csv(
    last_results: dict[str, Any],
    fallback_margin_rules: list[dict[str, float]] | None = None,
) -> str:
    """Zwraca treść raportu w formacie CSV na podstawie ostatnich wyników."""

    if not last_results:
        raise ValueError("Brak danych do wydruku. Najpierw wykonaj obliczenia.")

    inputs = last_results.get("inputs", {})
    wyniki = last_results.get("wyniki", {})

    def fmt(value: Any, digits: int = 2) -> str:
        try:
            return f"{float(value):.{digits}f}"
        except (TypeError, ValueError):
            return "-"

    rows: list[tuple[str, str, str]] = []

    def add_row(section: str, label: str, value: str) -> None:
        rows.append((section, label, value))

    add_row("Informacje", "Nazwa programu", "Kalkulator Rekruso")
    fala = inputs.get("fala")
    if fala:
        add_row("Informacje", "Rodzaj fali", str(fala))
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
        "Stawka transportowa [zł/km]",
        fmt(inputs.get("stawka_transport")),
    )
    add_row(
        "Parametry wejściowe",
        "Dystans transportu [km]",
        fmt(inputs.get("dystans")),
    )
    add_row(
        "Parametry wejściowe",
        "Transport z powrotem",
        "tak" if inputs.get("powrot") else "nie",
    )

    add_row(
        "Wymiary",
        "Formatka [mm]",
        fmt(wyniki.get("formatka_mm")),
    )
    add_row(
        "Wymiary",
        "Wymiar zewnętrzny [mm]",
        fmt(wyniki.get("wymiar_zewnetrzny_mm")),
    )

    bigi = wyniki.get("bigi", {})
    add_row(
        "Bigowanie",
        "Segmenty bigów (mm)",
        " | ".join(
            [
                fmt(bigi.get("c8")),
                fmt(bigi.get("d8")),
                fmt(bigi.get("e8")),
            ]
        ),
    )

    sumy_bigowe = wyniki.get("sumy_bigowe", {})
    add_row(
        "Bigowanie",
        "Pozycje bigów (mm)",
        " | ".join(
            [
                fmt(sumy_bigowe.get("c9")),
                fmt(sumy_bigowe.get("d9")),
                fmt(sumy_bigowe.get("e9")),
                fmt(sumy_bigowe.get("f9")),
                fmt(sumy_bigowe.get("g9")),
                fmt(sumy_bigowe.get("h9")),
                fmt(sumy_bigowe.get("i9")),
                fmt(sumy_bigowe.get("j9")),
            ]
        ),
    )

    minimum = wyniki.get("minimum_produkcji", {})
    add_row("Minimum produkcyjne", "AQ [szt.]", fmt(minimum.get("aq"), 0))
    add_row("Minimum produkcyjne", "CON [szt.]", fmt(minimum.get("con"), 0))
    add_row("Minimum produkcyjne", "PG [szt.]", fmt(minimum.get("pg"), 0))

    paletyzacja = wyniki.get("paletyzacja", {})
    add_row("Paletyzacja", "Długość [mm]", fmt(paletyzacja.get("dlugosc")))
    add_row("Paletyzacja", "Szerokość [mm]", fmt(paletyzacja.get("szerokosc")))

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

    margin_rules = last_results.get("margin_rules")
    if not margin_rules and fallback_margin_rules:
        margin_rules = fallback_margin_rules
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


def print_text_document(content: str, suffix: str = ".txt", *, prefer_notepad: bool = False) -> None:
    """Zapisuje treść do pliku tymczasowego i wysyła go na drukarkę."""

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "w", delete=False, encoding="utf-8", suffix=suffix
        ) as temp_file:
            temp_file.write(content)
            temp_path = Path(temp_file.name)
    except OSError as exc:
        raise PrinterError("Nie udało się przygotować pliku do wydruku.") from exc

    try:
        _send_to_printer(temp_path, prefer_notepad=prefer_notepad)
    finally:
        if temp_path is not None:
            _schedule_cleanup(temp_path)


def _send_to_printer(path: Path, *, prefer_notepad: bool) -> None:
    try:
        if sys.platform.startswith("win"):
            if prefer_notepad:
                try:
                    subprocess.run(["notepad.exe", "/p", str(path)], check=True)
                    return
                except FileNotFoundError:
                    if hasattr(os, "startfile"):
                        try:
                            os.startfile(str(path), "print")  # type: ignore[attr-defined]
                            return
                        except OSError as exc:
                            raise PrinterError(
                                "Nie udało się uruchomić systemowego polecenia drukowania."
                            ) from exc
                    raise PrinterError(
                        "Nie znaleziono programu Notepad wymaganego do wydruku."
                    )
                except subprocess.CalledProcessError as exc:
                    raise PrinterError(
                        "System Windows zgłosił błąd podczas drukowania."
                    ) from exc
            else:
                if hasattr(os, "startfile"):
                    try:
                        os.startfile(str(path), "print")  # type: ignore[attr-defined]
                        return
                    except OSError as exc:
                        raise PrinterError(
                            "Nie udało się uruchomić systemowego polecenia drukowania."
                        ) from exc
                try:
                    subprocess.run(["print", str(path)], check=True)
                    return
                except (FileNotFoundError, subprocess.CalledProcessError) as exc:
                    raise PrinterError(
                        "System Windows zgłosił błąd podczas drukowania."
                    ) from exc
        elif sys.platform == "darwin":
            subprocess.run(["lp", str(path)], check=True)
            return
        else:
            subprocess.run(["lpr", str(path)], check=True)
            return
    except FileNotFoundError as exc:
        raise PrinterError("Nie znaleziono polecenia drukarki w systemie.") from exc
    except subprocess.CalledProcessError as exc:
        raise PrinterError("Polecenie drukarki zakończyło się niepowodzeniem.") from exc


def _schedule_cleanup(path: Path) -> None:
    def _cleanup() -> None:
        try:
            path.unlink()
        except FileNotFoundError:
            pass

    timer = threading.Timer(10.0, _cleanup)
    timer.daemon = True
    timer.start()


__all__ = [
    "PrinterError",
    "build_summary_csv",
    "print_text_document",
]
