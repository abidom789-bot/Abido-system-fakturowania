# Instrukcja — Abido System Fakturowania

## Spis treści
1. [Uruchamianie aplikacji](#1-uruchamianie-aplikacji)
2. [Codzienny workflow miesięczny](#2-codzienny-workflow-miesięczny)
3. [Testowanie ekstrakcji kwot z faktur PDF](#3-testowanie-ekstrakcji-kwot-z-faktur-pdf)
4. [Jak nauczyć program nowego wzorca](#4-jak-nauczyć-program-nowego-wzorca)
5. [Wdrażanie zmian na serwer (git push)](#5-wdrażanie-zmian-na-serwer-git-push)
6. [Struktura plików projektu](#6-struktura-plików-projektu)
7. [Dane dostępowe](#7-dane-dostępowe)

---

## 1. Uruchamianie aplikacji

Aplikacja działa w chmurze na **Streamlit Cloud** — nie trzeba nic instalować lokalnie.

**Adres:** `abido-system-fakturowania-eplj5de35kzs6mgy3na5yj.streamlit.app`

---

## 2. Codzienny workflow miesięczny

### Krok po kroku (np. miesiąc `042026`):

```
1. Wgraj faktury PDF do folderu [FVS] Kosztowe w podfolderze 042026 na Drive

2. Kliknij "Zaczytaj faktury kosztowe"
   → Program odczyta kwoty z PDF i doda wiersze do arkusza ze statusem 0

3. W Google Sheets (zakładka 042026):
   - Sprawdź każdą fakturę
   - Popraw kwotę w kolumnie B jeśli program źle odczytał
   - Wpisz "1" w kolumnie C gdy faktura potwierdzona

4. Pobierz wyciąg bankowy XLS, wgraj do folderu Listy_operacji_abido

5. Kliknij "Paruj wyciąg bankowy z arkuszem"
   → Program dopasuje transakcje do faktur i wypełni kolumny E–Q

6. Sprawdź dopasowania:
   - Jeśli OK → wpisz "2" w kolumnie C (zatwierdzenie)
   - Jeśli błąd → popraw ręcznie, zostaw status 1

7. Kliknij "Generuj faktury sprzedaży PDF"
   → Faktury trafią na Drive do Faktury-sprzedazy / Faktury sprzedazy 042026
```

### Statusy wierszy (kolumna C):

| Status | Znaczenie | Zachowanie programu |
|--------|-----------|---------------------|
| `0` | Niezweryfikowany | Nadpisywany przy ponownym "Zaczytaj" |
| `1` | Zweryfikowany | Chroniony; parowany z wyciągiem |
| `2` | Zatwierdzony | Całkowicie nietykalny |

### Podgląd arkusza:

Przycisk **"Wyświetl ex"** obok pola miesiąca otwiera edytowalny widok arkusza.
Można zmieniać statusy bezpośrednio w tabeli i kliknąć **"Zapisz zmiany"**.

---

## 3. Testowanie ekstrakcji kwot z faktur PDF

Lokalny skrypt `test_extract.py` pozwala sprawdzać i ulepszać czytanie kwot
**bez uruchamiania całej aplikacji** i bez wgrywania plików na Drive.

### Wymagania (jednorazowo):

```
py -m pip install pdfplumber
```

### Uruchomienie:

```
cd "C:\repos abidom789\Abido-system-fakturowania"
py test_extract.py
```

### Co robi skrypt:

- Wczytuje wszystkie PDF z folderu `testowe faktury kosztowe/`
- Dla każdego pliku pokazuje wynik ekstrakcji i linie z faktury
- Porównuje z oczekiwaną kwotą ze słownika `EXPECTED`
- Wyświetla `✓ OK` lub `✗ BŁĄD`
- Na końcu: podsumowanie ile OK / ile błędów

---

## 4. Jak nauczyć program nowego wzorca

### Krok 1 — dodaj fakturę do folderu testowego

Skopiuj plik PDF do:
```
testowe faktury kosztowe\
```

### Krok 2 — uruchom skrypt i podaj prawidłową kwotę

```
py test_extract.py
```

Dla nowego pliku skrypt zapyta:
```
Wpisz prawidłową kwotę (np. -339,45) lub ENTER jeśli skan:
```

Wpisz kwotę z faktury (ze znakiem minus, z przecinkiem, np. `-320,75`).
Skrypt automatycznie zapisze ją do słownika `EXPECTED` w pliku `test_extract.py`.

Jeśli to skan (zdjęcie — brak tekstu), naciśnij tylko ENTER.

### Krok 3 — sprawdź wynik

- Jeśli pokazuje `✓ OK` — wzorzec już działa, przejdź do kroku 5
- Jeśli pokazuje `✗ BŁĄD` — patrz na "PASUJĄCE LINIE" i dodaj wzorzec (krok 4)

### Krok 4 — dodaj nowy wzorzec (jeśli potrzeba)

Otwórz `test_extract.py` i znajdź funkcję `extract_gross_amount()`.
W bloku `patterns = [...]` dodaj nową linię, np.:

```python
# Nowy wzorzec dla faktury XYZ
r"twoja\s+fraza[^\d]*?" + _NUM,
```

Gdzie `_NUM = r"([\d ]+[,.][\d]{2})"` to przechwytywanie liczby w formacie `1 234,56`.

Uruchom skrypt ponownie i sprawdź czy wynik jest poprawny.

### Krok 5 — skopiuj wzorzec do app.py

Gdy test przechodzi (`✓ OK`), skopiuj nowy wzorzec z `test_extract.py`
do tej samej funkcji `extract_gross_amount()` w pliku `app.py`.

Obie funkcje muszą być identyczne.

### Krok 6 — wyślij na serwer

```
git add app.py test_extract.py
git commit -m "Dodaj wzorzec dla faktury XYZ"
git push origin master:main
```

---

## 5. Wdrażanie zmian na serwer (git push)

Wszystkie zmiany kodu trafiają na Streamlit Cloud przez GitHub.

```
cd "C:\repos abidom789\Abido-system-fakturowania"

# Sprawdź co zostało zmienione
git status

# Dodaj zmienione pliki (nigdy nie dodawaj plików .json z kluczami!)
git add app.py
git add test_extract.py   # opcjonalnie

# Utwórz commit z opisem zmiany
git commit -m "Opis zmiany"

# Wyślij na GitHub → Streamlit automatycznie się przeładuje
git push origin master:main
```

**Uwaga:** Nie dodawaj do git plików:
- `*.json` (klucze Google Cloud)
- `*.xls`, `*.xlsx`, `*.pdf` (dane)
- `client_secret_*.json`

---

## 6. Struktura plików projektu

```
Abido-system-fakturowania/        ← katalog roboczy (lokalnie)
│
├── app.py                        ← cały kod aplikacji Streamlit
├── requirements.txt              ← biblioteki Python (Streamlit Cloud)
├── packages.txt                  ← pakiety systemowe (Streamlit Cloud)
├── INSTRUKCJA.md                 ← ten plik
├── CLAUDE.md                     ← instrukcje dla AI (Claude Code)
│
├── fonts/                        ← czcionki DejaVu do generowania PDF
│   ├── DejaVuSans.ttf
│   └── DejaVuSans-Bold.ttf
│
├── test_extract.py               ← lokalny tester ekstrakcji kwot
├── testowe faktury kosztowe/     ← PDF do testów (nie trafiają na Drive)
│
├── get_refresh_token.py          ← jednorazowy skrypt OAuth2 (lokalne)
└── docs/
    └── opis_programu.md          ← szczegółowy opis logiki programu
```

### Architektura app.py:

```
app.py
├── KONFIGURACJA          stałe: FOLDER_ID, SPREADSHEET_ID, separatory
├── FUNKCJE BACKENDOWE    czyste Python, zero Streamlit
│   ├── Google Drive      find_subfolder, list_pdfs, download_pdf
│   ├── Google Sheets     read_all_sections, rebuild_sheet, apply_sync_logic
│   ├── PDF               extract_gross_amount, build_invoice_pdf_bytes
│   ├── Parowanie         pair_transactions, sync_parowanie
│   └── Statystyki        count_kosztowe_statuses, count_parowanie_statuses
└── INTERFEJS STREAMLIT   przyciski + bloki akcji (if btn_xxx:)
```

---

## 7. Dane dostępowe

| Element | Wartość |
|---------|---------|
| Aplikacja Streamlit | `abido-system-fakturowania-eplj5de35kzs6mgy3na5yj.streamlit.app` |
| Konto Streamlit | `abidom789-bot` |
| Repo GitHub | `abidom789-bot/Abido-system-fakturowania` |
| Google Cloud projekt | `regal-river-494622-c7` |
| Service Account | `fakturowanie-bot@regal-river-494622-c7.iam.gserviceaccount.com` |
| SPREADSHEET_ID | `1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0` |
| FOLDER_ID (Faktury) | `1kwY6tOalKS2jnidABw6uUV23ykMj1iR2` |

### Sekrety Streamlit Cloud (`[gcp_service_account]`):

Klucz JSON Service Account Google — edytuj w:
Streamlit Cloud → aplikacja → Settings → Secrets

### OAuth2 Drive Upload (`[google_drive_oauth]`):

Potrzebny do wgrywania faktur sprzedaży na Drive.
Jeśli wygaśnie lub trzeba odświeżyć:
```
py get_refresh_token.py
```
i wklej nowy `refresh_token` do Secrets.
