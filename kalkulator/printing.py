"""Funkcje związane z przygotowaniem wydruków."""

from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import threading
import unicodedata
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

    sections = build_summary_sections(
        last_results,
        fallback_margin_rules=fallback_margin_rules,
    )

    csv_lines = ["sep=;", "Sekcja;Parametr;Wartość"]
    for section, rows in sections:
        for label, value in rows:
            csv_lines.append(";".join((section, label, value)))
    return "\n".join(csv_lines)


def build_summary_sections(
    last_results: dict[str, Any],
    fallback_margin_rules: list[dict[str, float]] | None = None,
) -> list[tuple[str, list[tuple[str, str]]]]:
    """Buduje strukturę danych z sekcjami podsumowania dla wydruków."""

    if not last_results:
        raise ValueError("Brak danych do wydruku. Najpierw wykonaj obliczenia.")

    inputs = last_results.get("inputs", {})
    wyniki = last_results.get("wyniki", {})
    klient = last_results.get("client", {})

    def fmt(value: Any, digits: int = 2) -> str:
        try:
            return f"{float(value):.{digits}f}"
        except (TypeError, ValueError):
            return "-"

    def fmt_int(value: Any) -> str:
        return fmt(value, digits=0)

    def txt(value: Any) -> str:
        text = str(value).strip()
        return text or "-"

    sections: list[tuple[str, list[tuple[str, str]]]] = []

    sections.append(
        (
            "Dane klienta",
            [
                ("Nazwa firmy", txt(klient.get("nazwa", ""))),
                ("Adres", txt(klient.get("adres", ""))),
                ("NIP", txt(klient.get("nip", ""))),
                ("E-mail", txt(klient.get("email", ""))),
                (
                    "Data wydruku",
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                ),
            ],
        )
    )

    sections.append(
        (
            "Parametry kartonu i nakłady",
            [
                ("Rodzaj fali", txt(inputs.get("fala", "-"))),
                ("Długość (DL) [mm]", fmt(inputs.get("dl"))),
                ("Szerokość (SZ) [mm]", fmt(inputs.get("sz"))),
                ("Wysokość (WYS) [mm]", fmt(inputs.get("wys"))),
                ("Gramatura [g/m²]", fmt(inputs.get("gramatura"), 0)),
                ("Cena surowca 1 m² [zł]", fmt(inputs.get("cena_m2"), 4)),
            ],
        )
    )

    bigi = wyniki.get("bigi", {})
    bigowe = wyniki.get("bigowe", {})
    sumy_bigowe = wyniki.get("sumy_bigowe", {})
    sections.append(
        (
            "Bigowanie",
            [
                (
                    "Segmenty bigów [mm]",
                    " | ".join(
                        [
                            fmt(bigi.get("c8")),
                            fmt(bigi.get("d8")),
                            fmt(bigi.get("e8")),
                        ]
                    ),
                ),
                (
                    "Pozycje bigów [mm]",
                    " | ".join(
                        [
                            fmt(sumy_bigowe.get("c9")),
                            fmt(sumy_bigowe.get("d9")),
                            fmt(sumy_bigowe.get("e9")),
                        ]
                    ),
                ),
                (
                    "Szerokości segmentów [mm]",
                    " | ".join(
                        [
                            fmt(bigowe.get("f8")),
                            fmt(bigowe.get("g8")),
                            fmt(bigowe.get("h8")),
                            fmt(bigowe.get("i8")),
                            fmt(bigowe.get("j8")),
                        ]
                    ),
                ),
                (
                    "Pozycje segmentów [mm]",
                    " | ".join(
                        [
                            fmt(sumy_bigowe.get("f9")),
                            fmt(sumy_bigowe.get("g9")),
                            fmt(sumy_bigowe.get("h9")),
                            fmt(sumy_bigowe.get("i9")),
                            fmt(sumy_bigowe.get("j9")),
                        ]
                    ),
                ),
            ],
        )
    )

    minimum = wyniki.get("minimum_produkcji", {})
    sections.append(
        (
            "Minimum produkcyjne",
            [
                ("AQ [szt.]", fmt_int(minimum.get("aq"))),
                ("CON [szt.]", fmt_int(minimum.get("con"))),
                ("PG [szt.]", fmt_int(minimum.get("pg"))),
            ],
        )
    )

    weryfikacja = wyniki.get("weryfikacja_zewnetrzna", {})
    sections.append(
        (
            "Wymiar zewnętrzny",
            [
                ("Długość zewnętrzna [mm]", fmt(weryfikacja.get("dl"))),
                ("Szerokość zewnętrzna [mm]", fmt(weryfikacja.get("sz"))),
                ("Wysokość zewnętrzna [mm]", fmt(weryfikacja.get("wys"))),
            ],
        )
    )

    paletyzacja = wyniki.get("paletyzacja", {})
    sections.append(
        (
            "Paletyzacja",
            [
                ("Długość paletyzacyjna [mm]", fmt(paletyzacja.get("dlugosc"))),
                ("Szerokość paletyzacyjna [mm]", fmt(paletyzacja.get("szerokosc"))),
            ],
        )
    )

    transport = wyniki.get("transport", {})
    sections.append(
        (
            "Koszty dodatkowe i transport",
            [
                (
                    "Dodatkowe koszty (partia) [zł]",
                    fmt(wyniki.get("koszty_dodatkowe")),
                ),
                (
                    "Stawka transportowa (wejściowa) [zł/km]",
                    fmt(inputs.get("stawka_transport")),
                ),
                ("Dystans [km]", fmt(inputs.get("dystans"))),
                ("Powrót", "tak" if inputs.get("powrot") else "nie"),
                ("Stawka końcowa [zł/km]", fmt(transport.get("stawka_pelna"))),
                ("Koszt łączny transportu [zł]", fmt(transport.get("koszt_calkowity"))),
            ],
        )
    )

    margin_rules = last_results.get("margin_rules")
    if not margin_rules and fallback_margin_rules:
        margin_rules = fallback_margin_rules
    if margin_rules:
        margin_rows: list[tuple[str, str]] = []
        for rule in margin_rules:
            qty_value = rule.get("max_quantity")
            try:
                qty_int = int(float(qty_value))
                qty_label = f"Do {qty_int} szt."
            except (TypeError, ValueError):
                qty_label = "Próg marży"
            try:
                margin_value = float(rule.get("margin_percent"))
                margin_text = f"{margin_value:.2f} %"
            except (TypeError, ValueError):
                margin_text = "-"
            margin_rows.append((qty_label, margin_text))
        if margin_rows:
            sections.append(("Progi marży", margin_rows))

    sections.append(
        (
            "Wyniki",
            [
                ("Formatka [mm]", fmt(wyniki.get("formatka_mm"))),
                (
                    "Wymiar zewnętrzny [mm]",
                    fmt(wyniki.get("wymiar_zewnetrzny_mm")),
                ),
                ("Zużycie m²/szt.", fmt(wyniki.get("zuzycie_m2_na_szt"), 3)),
                ("Waga kg/szt.", fmt(wyniki.get("waga_kg_na_szt"), 3)),
                (
                    "Koszt materiału/szt. [zł]",
                    fmt(wyniki.get("koszt_mat_na_szt"), 4),
                ),
            ],
        )
    )

    return sections


def build_summary_pdf(
    last_results: dict[str, Any],
    fallback_margin_rules: list[dict[str, float]] | None = None,
) -> bytes:
    """Tworzy plik PDF z podsumowaniem kalkulacji."""

    sections = build_summary_sections(
        last_results,
        fallback_margin_rules=fallback_margin_rules,
    )

    pdf = _SummaryPDFBuilder()
    pdf.add_title("Kalkulator Rekruso — podsumowanie")
    for section_name, rows in sections:
        pdf.add_section(section_name, rows)
    return pdf.render()


def print_text_document(
    content: str, suffix: str = ".txt", *, prefer_notepad: bool = False
) -> None:
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


def print_pdf_document(content: bytes) -> None:
    """Zapisuje dokument PDF w pliku tymczasowym i drukuje go przez Adobe Reader."""

    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("wb", delete=False, suffix=".pdf") as temp_file:
            temp_file.write(content)
            temp_path = Path(temp_file.name)
    except OSError as exc:
        raise PrinterError("Nie udało się przygotować pliku PDF do wydruku.") from exc

    try:
        _send_to_printer(temp_path, use_adobe_reader=True)
    finally:
        if temp_path is not None:
            _schedule_cleanup(temp_path)


def _send_to_printer(
    path: Path, *, prefer_notepad: bool = False, use_adobe_reader: bool = False
) -> None:
    try:
        if sys.platform.startswith("win"):
            if use_adobe_reader:
                reader_path = _find_adobe_reader()
                if reader_path is None:
                    raise PrinterError(
                        "Nie znaleziono instalacji Adobe Reader potrzebnej do wydruku PDF."
                    )
                try:
                    subprocess.run(
                        [str(reader_path), "/N", "/T", str(path)],
                        check=True,
                    )
                    return
                except subprocess.CalledProcessError as exc:
                    raise PrinterError(
                        "Adobe Reader zgłosił błąd podczas drukowania dokumentu."
                    ) from exc
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


def _find_adobe_reader() -> Path | None:
    """Próbuje odnaleźć plik wykonywalny Adobe Reader na systemie Windows."""

    possible_names = ("AcroRd32.exe", "Acrobat.exe")
    search_roots: list[Path] = []
    for env_var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        value = os.environ.get(env_var)
        if value:
            search_roots.append(Path(value) / "Adobe")

    for root in list(search_roots):
        if root.exists():
            for name in possible_names:
                candidate = root / "Acrobat Reader DC" / "Reader" / name
                if candidate.exists():
                    return candidate
                candidate = root / "Reader 11.0" / "Reader" / name
                if candidate.exists():
                    return candidate
                candidate = root / "Acrobat DC" / "Acrobat" / name
                if candidate.exists():
                    return candidate

    for root in search_roots:
        if root.exists():
            for name in possible_names:
                try:
                    found = next(root.rglob(name))
                except StopIteration:
                    continue
                else:
                    return found
    return None


class _SummaryPDFBuilder:
    """Pomocnicza klasa tworząca prosty dokument PDF z sekcjami tabel."""

    page_width = 595.28  # A4 w punktach PDF
    page_height = 841.89
    margin_x = 36.0
    margin_top = 36.0
    margin_bottom = 40.0
    section_spacing = 18.0
    header_height = 24.0
    row_height = 20.0
    padding_y = 6.0
    column_ratio = 0.45

    def __init__(self) -> None:
        self.pages: list[list[str]] = []
        self.cursor_y: float = 0.0
        self.current_page: list[str] | None = None
        self._new_page()

    def add_title(self, text: str) -> None:
        if self.cursor_y - 30.0 < self.margin_bottom:
            self._new_page()
        y = self.cursor_y
        self._text(self.margin_x, y, text, font="F2", size=16)
        self.cursor_y = y - 32.0

    def add_section(self, title: str, rows: list[tuple[str, str]]) -> None:
        if not rows:
            return

        section_height = (
            self.header_height
            + len(rows) * self.row_height
            + 2 * self.padding_y
        )
        if self.cursor_y - section_height < self.margin_bottom:
            self._new_page()

        top = self.cursor_y
        bottom = top - section_height
        left = self.margin_x
        width = self.page_width - 2 * self.margin_x
        header_bottom = top - self.header_height
        row_top_line = header_bottom - self.padding_y
        row_bottom_line = row_top_line - len(rows) * self.row_height
        column_split = left + width * self.column_ratio

        self._append(f"{left:.2f} {bottom:.2f} {width:.2f} {section_height:.2f} re S")

        self._append("q")
        self._append("0.9 g")
        self._append(
            f"{left:.2f} {header_bottom:.2f} {width:.2f} {self.header_height:.2f} re f"
        )
        self._append("Q")

        title_y = top - (self.header_height / 2.0) - 4.0
        self._text(left + 8.0, title_y, title, font="F2", size=12)

        self._append(
            f"{column_split:.2f} {row_top_line:.2f} m {column_split:.2f} {row_bottom_line:.2f} l S"
        )
        for index in range(len(rows) + 1):
            line_y = row_top_line - index * self.row_height
            self._append(
                f"{left:.2f} {line_y:.2f} m {left + width:.2f} {line_y:.2f} l S"
            )

        for idx, (label, value) in enumerate(rows):
            baseline = row_top_line - idx * self.row_height - 12.0
            self._text(left + 8.0, baseline, label, font="F1", size=10)
            value_lines = value.splitlines() or [""]
            for line_offset, line in enumerate(value_lines):
                self._text(
                    column_split + 8.0,
                    baseline - line_offset * 12.0,
                    line,
                    font="F1",
                    size=10,
                )

        self.cursor_y = bottom - self.section_spacing

    def render(self) -> bytes:
        if not self.pages:
            self._new_page()

        font_objects = [
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
        ]

        num_pages = len(self.pages)
        font_start = 3
        page_start = font_start + len(font_objects)
        content_start = page_start + num_pages
        total_objects = content_start + num_pages - 1

        objects: list[bytes] = [b""] * (total_objects + 1)

        pages_obj = 2
        objects[1] = f"<< /Type /Catalog /Pages {pages_obj} 0 R >>".encode("ascii")

        kids = " ".join(f"{page_start + i} 0 R" for i in range(num_pages))
        media_box = (
            f"[0 0 {self.page_width:.2f} {self.page_height:.2f}]".encode("ascii")
        )
        objects[pages_obj] = (
            b"<< /Type /Pages /Kids ["
            + kids.encode("ascii")
            + b"] /Count "
            + str(num_pages).encode("ascii")
            + b" /MediaBox "
            + media_box
            + b" >>"
        )

        for index, font in enumerate(font_objects):
            objects[font_start + index] = font

        for page_index, commands in enumerate(self.pages):
            page_obj_num = page_start + page_index
            content_obj_num = content_start + page_index
            resources = (
                f"<< /Font << /F1 {font_start} 0 R /F2 {font_start + 1} 0 R >> >>"
            )
            objects[page_obj_num] = (
                f"<< /Type /Page /Parent {pages_obj} 0 R /Resources {resources}"
                f" /Contents {content_obj_num} 0 R >>"
            ).encode("ascii")

            stream_text = "\n".join(commands).encode("ascii")
            stream = (
                b"<< /Length "
                + str(len(stream_text)).encode("ascii")
                + b" >>\nstream\n"
                + stream_text
                + b"\nendstream"
            )
            objects[content_obj_num] = stream

        buffer = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]

        for obj_num in range(1, len(objects)):
            offsets.append(len(buffer))
            buffer.extend(f"{obj_num} 0 obj\n".encode("ascii"))
            buffer.extend(objects[obj_num])
            buffer.extend(b"\nendobj\n")

        xref_pos = len(buffer)
        buffer.extend(f"xref\n0 {len(objects)}\n".encode("ascii"))
        buffer.extend(b"0000000000 65535 f \n")
        for offset in offsets[1:]:
            buffer.extend(f"{offset:010} 00000 n \n".encode("ascii"))

        buffer.extend(b"trailer\n")
        buffer.extend(
            f"<< /Size {len(objects)} /Root 1 0 R >>\n".encode("ascii")
        )
        buffer.extend(b"startxref\n")
        buffer.extend(f"{xref_pos}\n".encode("ascii"))
        buffer.extend(b"%%EOF")
        return bytes(buffer)

    def _append(self, command: str) -> None:
        if self.current_page is None:
            self._new_page()
        self.current_page.append(command)

    def _text(self, x: float, y: float, text: str, *, font: str, size: int) -> None:
        escaped = _pdf_escape_text(text)
        self._append(
            f"BT /{font} {size} Tf {x:.2f} {y:.2f} Td ({escaped}) Tj ET"
        )

    def _new_page(self) -> None:
        self.current_page = []
        self.pages.append(self.current_page)
        self.cursor_y = self.page_height - self.margin_top
        self._append("0.5 w")


def _pdf_escape_text(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    filtered = []
    for char in normalized:
        if unicodedata.combining(char):
            continue
        if 32 <= ord(char) <= 126:
            filtered.append(char)
        elif char == "\t":
            filtered.append("    ")
        else:
            continue
    escaped = "".join(filtered)
    escaped = escaped.replace("\\", "\\\\")
    escaped = escaped.replace("(", "\\(")
    escaped = escaped.replace(")", "\\)")
    return escaped


__all__ = [
    "PrinterError",
    "build_summary_csv",
    "build_summary_pdf",
    "build_summary_sections",
    "print_text_document",
    "print_pdf_document",
]
