"""
Skrypt testowy do sprawdzania ekstrakcji kwot brutto z faktur PDF.
Uruchom: py test_extract.py

WORKFLOW — jak dodać nową fakturę:
  1. Wrzuć plik PDF do folderu  testowe faktury kosztowe/
  2. Uruchom skrypt — zobaczysz wynik i linie z faktury
  3. Skrypt zapyta o prawidłową kwotę i sam dopisze ją do EXPECTED
  4. Jeśli wynik jest błędny: dodaj nowy wzorzec w extract_gross_amount()
  5. Kiedy test przechodzi — skopiuj wzorzec do app.py (ta sama funkcja)
"""

import io
import os
import re
import sys
import pdfplumber

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR    = os.path.join(SCRIPT_DIR, "testowe faktury kosztowe")

# ----------------------------------------------------------------
# OCZEKIWANE WYNIKI
# Klucz = nazwa pliku PDF, wartość = oczekiwana kwota (string "-X,XX")
# None  = skan bez tekstu (program nie może odczytać — OK jeśli zwróci "")
# "?"   = nieznana — skrypt zapyta przy pierwszym uruchomieniu
# ----------------------------------------------------------------
EXPECTED = {
    "Agata meble 9.3.2026.pdf":                          None,
    "Allegro Rejda 1.4.2026.pdf":                        "?",
    "Allegro Savinga 8.4.2026.pdf":                      "-339,45",
    "Eon Omulewska 02032026 _229758775727.pdf":          "-320,75",
    "Mafika 032026 ksiegowosc.pdf":                      "-824,10",
    "Netia zbiorcza 032026 5911178724_22_0.pdf":         "-129,00",
    "Pgnig Perzynskiego 032026 5681995_41_2026_f.pdf":   "-48,59",
    "Play naleczowska 032026 .pdf":                      "-138,84",
}


# ----------------------------------------------------------------
# FUNKCJA EKSTRAKCJI — kopia z app.py; tu ją ulepszamy, potem
# wklejamy poprawki do app.py (funkcja extract_gross_amount)
# ----------------------------------------------------------------

def extract_gross_amount(pdf_bytes):
    _NUM = r"([\d ]+[,.][\d]{2})"

    patterns = [
        # 1. "Do zapłaty" / "Pozostaje do zapłaty" / "Razem do zapłaty"
        #    — pomijamy 0,00 (Allegro: zapłacone przy zakupie)
        r"(?:razem\s+|pozostaje\s+)?do\s+zap[lł]aty[^\d]*?" + _NUM,
        # 2. "Wartość brutto X,XX" jako osobna linia (np. Allegro)
        r"warto[śs][ćc]\s+brutto\s+" + _NUM,
        # 3. "Należność X,XX" — np. E.ON
        r"nale[żz]no[śs][ćc]\s+" + _NUM,
        # 4. "Razem brutto" / "Suma brutto" / "Ogółem brutto"
        r"(?:razem|suma|og[oó][lł]em)\s+brutto[^\d]*?" + _NUM,
        # 5. "Kwota brutto: X,XX"
        r"kwota\s+brutto\s*[:\-]\s*" + _NUM,
        # 6. "Łączna kwota" / "Total"
        r"(?:[lł][aą]czna\s+kwota|total)\s*[:\-]?\s*" + _NUM,
    ]

    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = "".join(page.extract_text() or "" for page in pdf.pages)

        if not text.strip():
            return ""

        tl = text.lower()

        for p in patterns:
            m = re.search(p, tl)
            if m:
                val = m.group(1).strip().replace(" ", "")
                if val in ("0,00", "0.00"):   # zapłacone z góry — szukaj dalej
                    continue
                return "-" + val

        # Ostatnia szansa: linia "Razem netto VAT brutto" — ostatnia liczba
        for line in tl.splitlines():
            if re.match(r"\s*(?:\d+\.\s+)?razem\b", line):
                nums = re.findall(r"\d+[,.]\d{2}", line)
                if len(nums) >= 2:
                    return "-" + nums[-1]

    except Exception as e:
        print(f"    [wyjątek: {e}]")

    return ""


# ----------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------

def _matching_lines(text):
    keywords = {"zapłat", "zaplat", "brutto", "razem", "suma", "łącznie",
                "total", "netto", "vat", "należ", "kwota", "ogółem", "ogolem"}
    return [
        line for line in text.splitlines()
        if any(k in line.lower() for k in keywords)
        or re.search(r"\d+[,.]\d{2}", line)
    ]


def _save_expected(name, value):
    """Dopisuje/aktualizuje wpis w słowniku EXPECTED w tym pliku."""
    path = os.path.abspath(__file__)
    with open(path, encoding="utf-8") as f:
        src = f.read()

    # Szukamy linii z danym plikiem lub wstawiamy nową przed zamknięciem "}"
    entry = f'    "{name}":'
    new_line = f'    "{name}":  "{value}",'

    if entry in src:
        # Zastąp istniejący wpis
        src = re.sub(
            rf'^    "{re.escape(name)}".*$',
            new_line,
            src,
            flags=re.MULTILINE,
        )
    else:
        # Wstaw przed zamknięciem słownika "}"
        src = src.replace(
            "\n}\n\n\n# ----",
            f"\n{new_line}\n}}\n\n\n# ----",
        )

    with open(path, "w", encoding="utf-8") as f:
        f.write(src)


# ----------------------------------------------------------------
# Runner
# ----------------------------------------------------------------

def process_pdf(path):
    name = os.path.basename(path)
    with open(path, "rb") as f:
        data = f.read()

    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = "".join(page.extract_text() or "" for page in pdf.pages)

    result = extract_gross_amount(data)
    exp    = EXPECTED.get(name, "NOWY")

    print("=" * 70)
    print(f"PLIK:  {name}")

    # --- Nowa faktura — zapytaj o prawidłową kwotę ---
    if exp == "NOWY":
        print(f"WYNIK: {result!r}   (nowy plik — brak oczekiwanej wartości)")
        print("-" * 70)
        if not text.strip():
            print("  (brak tekstu — skan/obraz)")
        else:
            for line in _matching_lines(text):
                print(f"  | {line}")
        print()
        ans = input("  Wpisz prawidłową kwotę (np. -339,45) lub ENTER jeśli skan: ").strip()
        if ans:
            val = ans if ans.startswith("-") else "-" + ans
            _save_expected(name, val)
            EXPECTED[name] = val
            print(f"  → Zapisano: {val!r}")
            ok = (result == val)
        else:
            _save_expected(name, "?")
            EXPECTED[name] = "?"
            ok = True  # nieznana — nie liczymy jako błąd
        print()
        return "new"

    # --- Znana faktura — sprawdź wynik ---
    ok = (exp is None and result == "") or result == exp
    status = "✓ OK" if ok else f"✗ BŁĄD (oczekiwano: {exp!r})"
    print(f"WYNIK: {result!r}   {status}")

    if not ok:
        print("-" * 70)
        if not text.strip():
            print("  (brak tekstu — skan/obraz)")
        else:
            for line in _matching_lines(text):
                print(f"  | {line}")
    print()
    return "ok" if ok else "fail"


if __name__ == "__main__":
    pdfs = sorted(
        os.path.join(PDF_DIR, f)
        for f in os.listdir(PDF_DIR)
        if f.lower().endswith(".pdf")
    )
    if not pdfs:
        print(f"Brak plików PDF w: {PDF_DIR}")
        sys.exit(1)

    counts = {"ok": 0, "fail": 0, "new": 0, "skip": 0}
    for p in pdfs:
        name = os.path.basename(p)
        if EXPECTED.get(name) is None:
            # skan — pokaż krótko i pomiń
            print("=" * 70)
            print(f"PLIK:  {name}")
            print("WYNIK: ''   ✓ OK  (skan — pomijam)")
            print()
            counts["skip"] += 1
            continue
        status = process_pdf(p)
        counts[status] = counts.get(status, 0) + 1

    print("=" * 70)
    print(f"PODSUMOWANIE:  {counts['ok']} OK  |  {counts['fail']} błędów  "
          f"|  {counts['new']} nowych  |  {counts['skip']} skanów pominięto")
