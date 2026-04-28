# Opis programu — Abido System Fakturowania

## Co robi program

Streamlit Cloud aplikacja do zarządzania fakturami miesięcznymi.
Łączy trzy źródła danych:
- **Google Drive** — pliki PDF faktur + pliki XLS wyciągów bankowych
- **Google Sheets** — arkusz z danymi miesięcznymi (jeden arkusz = jeden miesiąc)
- **Interfejs użytkownika** — przyciski w przeglądarce

Główne funkcje:
1. Odczytuje faktury PDF z Google Drive i wyciąga z nich kwotę brutto
2. Zapisuje dane faktur do arkusza Google Sheets
3. Paruje pozycje z wyciągu bankowego (XLS) z fakturami w arkuszu
4. Pokazuje statusy i podsumowania

---

## Struktura Google Drive

```
Faktury/                          ← folder główny (FOLDER_ID)
├── 032026/                       ← podfolder miesięczny (marzec 2026)
│   ├── [FVS] Kosztowe/           ← faktury kosztowe PDF
│   └── [FVS] Sprzedaz/           ← faktury sprzedaży PDF (opcjonalnie)
├── 042026/                       ← podfolder miesięczny (kwiecień 2026)
└── Listy_operacji_abido/         ← wyciągi bankowe XLS
    ├── lista_operacji_032026.xls
    └── lista_operacji_042026.xls
```

Nazwa podfolderu miesięcznego: `MMRRRR` np. `032026` = marzec 2026.
Nazwa wyciągu: zawsze `lista_operacji_MMRRRR.xls`.

---

## Struktura Google Sheets

Jeden arkusz (zakładka) = jeden miesiąc. Nazwa zakładki = nazwa podfolderu (np. `032026`).

### Kolumny A–P

| Kol | Nazwa              | Kto wypełnia                        |
|-----|--------------------|-------------------------------------|
| A   | Nazwa / Plik       | Program (z nazwy pliku PDF)         |
| B   | Kwota brutto       | Program (z treści PDF), zawsze ujemna dla kosztów |
| C   | Status             | Użytkownik ręcznie (0 / 1 / 2)      |
| D   | Adres              | Program (Tworz wiersze sprzedazy)   |
| E   | Klucz_Ksiegowy     | Program (Paruj wyciag, auto)        |
| F   | wyciag_Kontrahent  | Program (Paruj wyciag)              |
| G   | wyciag_Kwota       | Program (Paruj wyciag)              |
| H   | Kwota_raport_kasowy| Program (tylko BLIK/gotówka)        |
| I   | Data_ksiegowania   | Program (Paruj wyciag)              |
| J   | wyciag_Tytul       | Program (Paruj wyciag)              |
| K   | wyciag_Data_op     | Program (Paruj wyciag)              |
| L   | wyciag_Rodzaj      | Program (Paruj wyciag)              |
| M   | wyciag_Waluta      | Program (Paruj wyciag)              |
| N   | wyciag_Nr_rachunku | Program (Paruj wyciag)              |
| O   | wyciag_Imie_Nazwisko | Program (Paruj wyciag)            |
| P   | Uwagi              | Program (auto) lub użytkownik       |

### Sekcje arkusza (separatory)

Arkusz jest podzielony na 4 sekcje kolorowanymi wierszami-separatorami:

| Separator (kolor)           | Zawiera                                    |
|-----------------------------|--------------------------------------------|
| FAKTURY KOSZTOWE (czerwony) | Wydatki firmy — zakupy, media, usługi      |
| FAKTURY SPRZEDAZY NAJEMCOM (zielony) | Przychody — czynsze, kaucje     |
| FAKTURY WLASCICIELE I SPOLDZIELNIE (pomarańczowy) | Przelewy do właścicieli i spółdzielni |
| NIEZNANE / NIESPAROWANE Z WYCIAGU (szary) | Transakcje z wyciągu bez pary w arkuszu |

Program zawsze zapisuje sekcje w tej kolejności.

---

## System statusów (kolumna C)

| Status | Znaczenie | Co robi program |
|--------|-----------|-----------------|
| 0 | Niezweryfikowany — świeżo dodany | Nadpisywany przy ponownym "Zaczytaj" |
| 1 | Zweryfikowany przez użytkownika | Chroniony — nigdy nie nadpisywany; parowany z wyciągiem |
| 2 | Sparowany i zatwierdzony | Całkowicie nietykalny — program go nigdy nie rusza |

**Zasada ochrony:**
- `status=1` lub `status=2` → program nie nadpisuje wiersza, nawet jeśli plik PDF został usunięty z Drive
- `status=0` → program nadpisuje przy każdym "Zaczytaj faktury kosztowe"

---

## Przyciski — co robi każdy

### Zaczytaj faktury kosztowe

**Kiedy używać:** gdy wgrałeś nowe pliki PDF do folderu `[FVS] Kosztowe` na Drive.

**Co robi:**
1. Wyszukuje rekurencyjnie folder `[FVS]` w podfolderze miesięcznym
2. Listuje wszystkie pliki PDF
3. Dla każdego nowego pliku (nie ma go w arkuszu lub ma status=0):
   - Pobiera PDF
   - Wyciąga kwotę brutto z tekstu (szuka wzorców: "Razem", "Do zapłaty", "Brutto")
   - Wpisuje wiersz do sekcji FAKTURY KOSZTOWE z status=0
4. Pliki z nazwą zawierającą słowo **"cash"** → trafiają na koniec sekcji kosztowej
5. Wiersze z status=1 lub status=2 → pozostają bez zmian

**Wynik:** nowe wiersze w sekcji FAKTURY KOSZTOWE.

---

### Sprawdz stan faktur kosztowych

**Kiedy używać:** aby zobaczyć ile faktur jest zweryfikowanych vs niezweryfikowanych.

**Co robi:**
- Liczy wiersze w sekcji FAKTURY KOSZTOWE według statusu
- Pokazuje tabelę: ile plików na Drive, ile w arkuszu, ile ze statusem 0/1/2

---

### Tworz wiersze faktur sprzedazy

**Kiedy używać:** raz na początku miesiąca, aby przygotować wiersze dla najemców.

**Co robi:**
1. Czyta listę najemców (folder Mieszkania na Drive)
2. Tworzy wiersze w sekcji FAKTURY SPRZEDAZY NAJEMCOM
3. Wypełnia kolumnę D (Adres) na podstawie danych najemcy
4. Wiersze z status=1 lub status=2 → chronione

---

### Paruj wyciag bankowy z arkuszem

**Kiedy używać:** po pobraniu wyciągu bankowego XLS i wgraniu go do `Listy_operacji_abido`.

**Co robi:**
1. Pobiera plik `lista_operacji_MMRRRR.xls` z Drive
2. Parsuje transakcje (wiersze od 7. w górę, nagłówek w wierszu 6)
3. Dla wszystkich wierszy z status=1 we wszystkich sekcjach (Kosztowe, Sprzedaz, Wlasciciele):
   - Uruchamia algorytm parowania wieloprzebiegowego (4 przebiegi)
   - Wypełnia kolumny E–P dopasowaną transakcją
4. Transakcje bez pary w arkuszu → dopisywane na dół jako sekcja NIEZNANE

**Algorytm parowania (4 przebiegi):**

| Przebieg | Kryterium | Pewność |
|----------|-----------|---------|
| 1 | Nazwisko + kwota | Najsilniejsze |
| 2 | Imię + kwota | Mocne |
| 3 | Słowa kluczowe z nazwy pliku + kwota | Dla faktur PDF (Allegro, Netia...) |
| 4 | Sama kwota | Ostatnia deska ratunku; dodaje uwagę "sprawdź" |

**Specjalne przypadki:**
- **Nest Bank S.A.** — płatności kartą; rzeczywisty sklep jest w Tytule operacji, nie w Kontrahencie
- **Revolut Bank UAB** — przelewy zagraniczne; rzeczywista osoba w Tytule operacji

**Ponowne parowanie:**
- status=2 → pominięty
- status=1 z istniejącym dopasowaniem → NIE nadpisywany
- status=1 bez dopasowania → próbuje ponownie (nowy, większy wyciąg)
- Sekcja NIEZNANE → zawsze zastępowana świeżą listą

---

### Status parowania

**Kiedy używać:** aby zobaczyć podsumowanie stanu parowania w danym miesiącu.

**Co robi:**
- Pokazuje tabelę dla każdej sekcji (Kosztowe, Sprzedaz, Wlasciciele):

| Wiersz | Znaczenie |
|--------|-----------|
| Status 0 — niezweryfikowane | Czekają na sprawdzenie przez użytkownika |
| Status 1 — bez pary z wyciągiem | Zweryfikowane ale brak transakcji na wyciągu |
| Status 1 — sparowane (czeka na '2') | Mają parę z wyciągu, czekają na zatwierdzenie |
| Status 2 — zatwierdzone | Zamknięte, nie będą już ruszane |

---

## Klucze księgowe (kolumna E)

Przypisywane automatycznie przy parowaniu na podstawie sekcji + kierunku transakcji:

| Klucz | Znaczenie |
|-------|-----------|
| `kos_pr_out` | Kosztowe — przelew wychodzący (zakup) |
| `kos_med_pr_out` | Kosztowe — media (Netia, EON, PGNiG, Play...) |
| `kos_pr_in` | Kosztowe — przychodzący (zwrot) |
| `kos_rk_kw` | Kosztowe — gotówka wydatek (raport kasowy) |
| `kos_rk_kp` | Kosztowe — gotówka przychód |
| `prz_naj_pr_in` | Sprzedaż najem — przelew przychodzący (czynsz) |
| `prz_naj_rk_kp` | Sprzedaż najem — gotówka przychód |
| `wla_pr_out` | Właściciele — przelew wychodzący |
| `wla_med_pr_out` | Właściciele — media wychodzące |
| `wla_pr_in` | Właściciele — przychodzący |
| `nieznany_out` | Niesparowana transakcja wychodząca |
| `nieznany_in` | Niesparowana transakcja przychodząca |

**Słowa kluczowe mediów:** Netia, EON, E.ON, PGNiG, Play, P4, energia, prąd, gaz, internet, woda, SBM

**Faktury gotówkowe (cash):**
Jeśli nazwa pliku zawiera słowo "cash" (dowolna wielkość liter) → klucz zawsze `kos_rk_kw`, niezależnie od wyciągu. Takie faktury trafiają na koniec sekcji FAKTURY KOSZTOWE.

---

## Typowy workflow miesięczny

```
1. Wgraj faktury PDF do [FVS] Kosztowe w podfolderze np. 032026
2. Kliknij "Zaczytaj faktury kosztowe" → wiersze status=0 w arkuszu
3. Sprawdź każdą fakturę:
   - Popraw kwotę w kolumnie B jeśli program źle odczytał
   - Wpisz "1" w kolumnie C gdy faktura potwierdzona
4. Pobierz wyciąg bankowy, wgraj do Listy_operacji_abido
5. Kliknij "Paruj wyciag bankowy" → kolumny E-P wypełnione
6. Sprawdź dopasowania:
   - Jeśli OK → wpisz "2" w kolumnie C (zatwierdzenie)
   - Jeśli błąd → popraw ręcznie, zostaw status=1
7. W połowie miesiąca wgraj nowy (większy) wyciąg → paruj ponownie
   - Status=2 → pominięte
   - Status=1 → nowe dopasowania z nowego wyciągu
8. Kliknij "Status parowania" aby zobaczyć podsumowanie
```

---

## Architektura kodu (app.py)

```
app.py
├── KONFIGURACJA           stałe: FOLDER_ID, SPREADSHEET_ID, separatory, słowa kluczowe
├── FUNKCJE BACKENDOWE     czyste Python, zero Streamlit
│   ├── Google Drive       find_subfolder, list_pdfs_from_drive, download_pdf
│   ├── Google Sheets      read_all_sections, rebuild_sheet, apply_sync_logic
│   ├── PDF                extract_gross_amount
│   ├── Parowanie          find_bank_file, parse_bank_statement, pair_transactions
│   │                      assign_klucz_ksiegowy, sync_parowanie
│   └── Statystyki         count_kosztowe_statuses, count_parowanie_statuses
└── INTERFEJS STREAMLIT
    ├── UI (przyciski)     deklaracje st.button(...)
    └── AKCJE              bloki if btn_xxx: — po jednym na przycisk
```

Każdy przycisk jest niezależny: osobna funkcja backendowa + osobny blok akcji.
Można modyfikować jeden przycisk bez ryzyka zepsucia pozostałych.

---

## Środowisko techniczne

| Element | Wartość |
|---------|---------|
| Platforma | Streamlit Cloud |
| Konto Streamlit | `abidom789-bot` |
| Repo GitHub | `abidom789-bot/Abido-system-fakturowania` |
| Branch lokalny | `master` → push do `main` na GitHub |
| Secrets | `[gcp_service_account]` — JSON Service Account Google |
| Google Cloud projekt | `regal-river-494622-c7` |
| Service Account | `fakturowanie-bot@regal-river-494622-c7.iam.gserviceaccount.com` |
| APIs | Google Drive API, Google Sheets API |
| SPREADSHEET_ID | `1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0` |
| FOLDER_ID (Faktury) | `1kwY6tOalKS2jnidABw6uUV23ykMj1iR2` |
