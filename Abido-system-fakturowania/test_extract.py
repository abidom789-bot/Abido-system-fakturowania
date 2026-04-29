"""
Skrypt testowy do sprawdzania ekstrakcji kwot brutto z faktur PDF.
Uruchom: py test_extract.py
"""

import io
import os
import re
import sys
import pdfplumber

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PDF_DIR = os.path.join(os.path.dirname(__file__), "testowe faktury kosztowe")

# ----------------------------------------------------------------
# OCZEKIWANE WYNIKI (uzupełnij ręcznie po weryfikacji faktur)
# ----------------------------------------------------------------
EXPECTED = {
    "Agata meble 9.3.2026.pdf":                          None,      # skan — brak tekstu
    "Eon Omulewska 02032026 _229758775727.pdf":          "-320,75",
    "Mafika 032026 ksiegowosc.pdf":                      "-824,10",
    "Netia zbiorcza 032026 5911178724_22_0.pdf":         "-129,00",
    "Pgnig Perzynskiego 032026 5681995_41_2026_f.pdf":   "-48,59",
    "Play naleczowska 032026 .pdf":                      "-138,84",
}

# ----------------------------------------------------------------
# ULEPSZONA FUNKCJA — tu rozwijamy i testujemy
# ----------------------------------------------------------------

def extract_gross_amount(pdf_bytes):
    _NUM = r"([\d ]+[,.][\d]{2})"

    patterns = [
        # 1. "Do zapłaty" / "Pozostaje do zapłaty" / "Razem do zapłaty"
        r"(?:razem\s+|pozostaje\s+)?do\s+zap[lł]aty[^\d]*?" + _NUM,
        # 2. "Należność X,XX" — np. E.ON
        r"nale[żz]no[śs][ćc]\s+" + _NUM,
        # 3. "Razem brutto" / "Suma brutto" / "Ogółem brutto"
        r"(?:razem|suma|og[oó][lł]em)\s+brutto[^\d]*?" + _NUM,
        # 4. "Kwota brutto: X,XX" — wartość w osobnej linii po nagłówku
        r"kwota\s+brutto\s*[:\-]\s*" + _NUM,
        # 5. "Łączna kwota" / "Total"
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
                return "-" + val

        # Ostatnia szansa: linia zaczynająca się od "Razem" z co najmniej 2 liczbami
        # Format: "Razem netto VAT brutto" — bierzemy OSTATNIĄ liczbę w linii
        for line in tl.splitlines():
            if re.match(r"\s*(?:\d+\.\s+)?razem\b", line):
                nums = re.findall(r"\d+[,.]\d{2}", line)
                if len(nums) >= 2:
                    return "-" + nums[-1]

    except Exception as e:
        print(f"    [wyjątek: {e}]")

    return ""


# ----------------------------------------------------------------
# Runner
# ----------------------------------------------------------------

def dump_pdf(path):
    with open(path, "rb") as f:
        data = f.read()

    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text = "".join(page.extract_text() or "" for page in pdf.pages)

    result  = extract_gross_amount(data)
    name    = os.path.basename(path)
    exp     = EXPECTED.get(name, "?")
    ok      = (exp is None and result == "") or result == exp
    status  = "✓ OK" if ok else f"✗ BŁĄD (oczekiwano: {exp!r})"

    print("=" * 70)
    print(f"PLIK:   {name}")
    print(f"WYNIK:  {result!r}   {status}")

    if not ok or not result:
        print("-" * 70)
        if not text.strip():
            print("  (brak tekstu — skan/obraz PDF)")
        else:
            keywords = {"zapłat", "zaplat", "brutto", "razem", "suma",
                        "łącznie", "total", "netto", "vat", "należ",
                        "kwota", "ogółem", "ogolem"}
            print("PASUJĄCE LINIE:")
            for line in text.splitlines():
                ll = line.lower()
                if any(k in ll for k in keywords) or re.search(r"\d+[,.]\d{2}", ll):
                    print(f"  | {line}")
    print()


if __name__ == "__main__":
    pdfs = sorted(
        os.path.join(PDF_DIR, f)
        for f in os.listdir(PDF_DIR)
        if f.lower().endswith(".pdf")
    )
    if not pdfs:
        print(f"Brak plików PDF w: {PDF_DIR}")
        sys.exit(1)

    passed = failed = skipped = 0
    for p in pdfs:
        name = os.path.basename(p)
        if EXPECTED.get(name) is None and name in EXPECTED:
            skipped += 1
        dump_pdf(p)
        with open(p, "rb") as f:
            r = extract_gross_amount(f.read())
        exp = EXPECTED.get(name, "?")
        if exp == "?":
            pass
        elif exp is None:
            skipped += 1
        elif r == exp:
            passed += 1
        else:
            failed += 1

    print("=" * 70)
    print(f"PODSUMOWANIE: {passed} OK  |  {failed} błędów  |  {skipped} pominięto (skany)")
