# Klucze księgowe, parowanie i workflow miesięczny

## 1. Struktura klucza księgowego (kolumna D)

Klucz składa się z segmentów oddzielonych `_`:

```
[kategoria]_[podkategoria]_[sposób płatności]_[kierunek]
```

### Kategoria (prefiks)

| Prefiks | Znaczenie |
|---------|-----------|
| `kos_` | Faktura kosztowa (wydatek firmy) |
| `prz_naj_` | Przychód od najemcy (faktura sprzedaży) |
| `wla_` | Właściciel lub spółdzielnia |
| `roz_` | Rozrachunki międzyokresowe / kaucje / bankomat |
| `nieznany_` | Transakcja z wyciągu bez pary w arkuszu |

### Podkategoria (opcjonalna)

| Człon | Znaczenie |
|-------|-----------|
| `med_` | Media: gaz, prąd, woda, internet (Netia, EON, PGNiG, Play, P4, SBM...) |
| `depo_all_` | Kaucja — cała kwota |
| `depo_part_` | Kaucja — część kwoty |
| `bankomat_` | Wypłata / wpłata bankomatu |
| `zus_` | Składki ZUS |
| `pod_` | Podatki |

### Sposób płatności

| Człon | Znaczenie |
|-------|-----------|
| `pr_` | Przelew bankowy |
| `rk_` | Raport kasowy (gotówka) |
| `fak_` | Sama faktura — brak płatności w tym miesiącu |

### Kierunek

| Człon | Znaczenie |
|-------|-----------|
| `_out` | Pieniądze wychodzą (płatność, wydatek) |
| `_in` | Pieniądze wchodzą (wpływ, przychód) |
| `_kw` | Kasa wydała (gotówka wychodzi) |
| `_kp` | Kasa przyjęła (gotówka wchodzi) |
| `_no_pay` | Faktura istnieje, płatność będzie w innym miesiącu |
| `_no_fak` | Płatność istnieje, faktura była w innym miesiącu |

---

## 2. Pełna lista kluczy

### Faktury kosztowe (`kos_`)

| Klucz | Kiedy używać |
|-------|-------------|
| `kos_pr_out` | Faktura kosztowa opłacona przelewem wychodzącym |
| `kos_pr_in` | Zwrot / refaktura — przelew przychodzi |
| `kos_med_pr_out` | Media (gaz, prąd, woda, internet) — przelew wychodzący |
| `kos_rk_kw` | Faktura kosztowa opłacona gotówką (kasa wydała) |
| `kos_rk_kp` | Zwrot gotówkowy — kasa przyjęła |
| `kos_fak_no_pay` | Faktura jest w tym miesiącu, płatność wyjdzie w innym miesiącu |
| `kos_pr_out_no_fak` | Przelew wychodzi, faktura była w innym miesiącu |

> **Reguła automatyczna:** faktury PDF z `cash` w nazwie pliku → zawsze `kos_rk_kw`

### Przychody od najemców (`prz_naj_`)

| Klucz | Kiedy używać |
|-------|-------------|
| `prz_naj_pr_in` | Czynsz / wpłata od najemcy przelewem |
| `prz_naj_rk_kp` | Czynsz / wpłata od najemcy gotówką (kasa przyjęła) |
| `prz_naj_fak_no_pay` | Faktura wystawiona, płatność wpłynie w innym miesiącu |
| `prz_naj_pr_in_no_fak` | Wpłata od najemcy, faktura była wystawiona w innym miesiącu |
| `prz_pr_in` | Zwrot / refaktura przychodząca (w sekcji kosztowej) |

### Właściciele i spółdzielnie (`wla_`)

| Klucz | Kiedy używać |
|-------|-------------|
| `wla_pr_out` | Przelew do właściciela lub spółdzielni |
| `wla_med_pr_out` | Opłata za media dla właściciela — przelew wychodzący |
| `wla_pr_in` | Przelew od właściciela (zwrot, korekta) |
| `wla_fak_no_pay` | Faktura właściciela w tym miesiącu, płatność w innym |
| `wla_pr_out_no_fak` | Przelew do właściciela, faktura była w innym miesiącu |

### Rozrachunki (`roz_`)

| Klucz | Kiedy używać |
|-------|-------------|
| `roz_depo_all_pr_in` | Kaucja przyjęta w całości — przelew |
| `roz_depo_part_pr_in` | Kaucja przyjęta częściowo — przelew |
| `roz_depo_all_pr_out` | Zwrot całej kaucji — przelew |
| `roz_depo_part_pr_out` | Zwrot części kaucji — przelew |
| `roz_depo_all_rk_kw` | Zwrot całej kaucji — gotówka |
| `roz_depo_part_rk_kw` | Zwrot części kaucji — gotówka |
| `roz_depo_all_rk_kp` | Kaucja przyjęta w całości — gotówka |
| `roz_depo_part_rk_kp` | Kaucja przyjęta częściowo — gotówka |
| `roz_bankomat_rk_kw` | Wypłata z bankomatu (pieniądze lądują w kasie) |
| `roz_bankomat_rk_kp` | Wpłata do bankomatu (pieniądze wychodzą z kasy) |
| `roz_pr_in` | Rozrachunki ogólne — przelew przychodzący |
| `roz_pr_out` | Rozrachunki ogólne — przelew wychodzący |
| `roz_rk_kw` | Rozrachunki gotówkowe — kasa wydała |
| `roz_rk_kp` | Rozrachunki gotówkowe — kasa przyjęła |

### Niesparowane z wyciągu (`nieznany_`)

| Klucz | Kiedy używać |
|-------|-------------|
| `nieznany_in` | Transakcja przychodząca z wyciągu bez pary faktury |
| `nieznany_out` | Transakcja wychodząca z wyciągu bez pary faktury |

> Klucze `nieznany_` są nadawane automatycznie przez program dla transakcji z wyciągu
> które nie zostały dopasowane do żadnego wiersza w arkuszu.

---

## 3. Przypadki międzyokresowe

Sytuacja gdy **miesiąc faktury ≠ miesiąc płatności**.

### Zasada ogólna
> Gdzie faktura — tam liczymy koszt lub przychód.
> Płatność w innym miesiącu dostaje klucz `_no_fak` (brak faktury w tym miesiącu).

### Przypadek A — faktura w lipcu, płatność w sierpniu

| Miesiąc | Wiersz | Klucz | Opis |
|---------|--------|-------|------|
| Lipiec | Faktura PDF | `kos_fak_no_pay` | Koszt zaksięgowany, brak płatności |
| Sierpień | Przelew bankowy | `kos_pr_out_no_fak` | Płatność wychodzi, faktura była w lipcu |

### Przypadek B — płatność w lutym, faktura w marcu

| Miesiąc | Wiersz | Klucz | Opis |
|---------|--------|-------|------|
| Luty | Przelew bankowy | `kos_pr_out_no_fak` | Przedpłata, faktura jeszcze nie istnieje |
| Marzec | Faktura PDF | `kos_fak_no_pay` | Faktura pojawia się, płatność już była |

### Analogicznie dla przychodów

| Sytuacja | Miesiąc faktury | Miesiąc wpłaty | Klucz faktury | Klucz wpłaty |
|----------|----------------|----------------|---------------|--------------|
| Najem — wpłata spóźniona | Sierpień | Wrzesień | `prz_naj_fak_no_pay` | `prz_naj_pr_in_no_fak` |
| Najem — przedpłata | Sierpień | Lipiec | `prz_naj_fak_no_pay` | `prz_naj_pr_in_no_fak` |

---

## 4. System statusów (kolumna C)

| Status | Znaczenie | Przy „Zaczytaj faktury" | Przy „Paruj wyciąg" |
|--------|-----------|------------------------|---------------------|
| `0` | Nowy — do weryfikacji kwoty | Nadpisywany | Pomijany |
| `1` | Gotowy do parowania | Chroniony | Parowany z wyciągiem |
| `2` | Zamrożony — sparowany i zatwierdzony | Chroniony | Pomijany (TX w pre_used) |
| `3` | Absolutny beton — nic go nie rusza | Chroniony | Pomijany |
| `9` | Złe parowanie — ponów przy następnym parowaniu | Chroniony | Traktowany jak `1` |

**Workflow statusów:**
```
PDF dodany → status 0
Użytkownik weryfikuje kwotę → wpisuje 1
Program paruje z wyciągiem → kolumny E-N wypełnione
Użytkownik zatwierdza dopasowanie → wpisuje 2
```

---

## 5. Algorytm parowania wyciągu bankowego (6 przebiegów)

Program wykonuje 6 przebiegów w kolejności priorytetu. Każda transakcja z wyciągu może być dopasowana **tylko raz**.

### Przebieg 1 — Nazwisko + kwota (najsilniejsze)
- Wyodrębnij ostatni token z kolumny A (nazwisko)
- Szukaj w wyciągu: `|kwota_wyciag| == |kwota_B|` ORAZ nazwisko w Kontrahencie lub Tytule
- Dla Nest Bank / Revolut → szukaj w Tytule operacji (kontrahent pośredni)

### Przebieg 2 — Imię + kwota
- Wyodrębnij pierwszy token z kolumny A (imię)
- Ta sama logika co przebieg 1

### Przebieg 3 — Nazwisko bez kwoty (słabe — fioletowe)
- Nazwisko pasuje, kierunek zgodny, kwota może się różnić
- Wiersz oznaczany kolorem fioletowym

### Przebieg 4 — Imię bez kwoty (słabe — fioletowe)
- Imię pasuje, kierunek zgodny, kwota może się różnić

### Przebieg 5 — Sama kwota
- Dokładnie jedna wolna transakcja o tej kwocie i kierunku
- Dodaje uwagę: „Dopasowanie tylko po kwocie — sprawdź"

### Przebieg 6 — Multi-parowanie (sub-wiersze)
- Dla wierszy sparowanych w przebiegach 1/2: szuka dodatkowych TX od tej samej osoby
- Dodaje **sub-wiersze** bezpośrednio pod wierszem głównym (col A puste, col D-N wypełnione)
- Kolor `_MULTI_BG` (jasny fiolet)
- Dotyczy tylko sekcji SPRZEDAZ i WLASC (KOSZTOWE wykluczone)

### Specjalne przypadki kontrahentów

| Kontrahent w wyciągu | Problem | Rozwiązanie |
|---------------------|---------|-------------|
| Nest Bank S.A. | Płatności kartą — rzeczywisty sklep w Tytule | Szukaj w Tytule: „Allegro Poznan...", „AGATA S.A..." |
| Revolut Bank UAB | Przelewy zagraniczne — osoba w Tytule | Wyodrębnij imię/nazwisko z Tytułu |

### Kierunek transakcji

| Sekcja | Oczekiwany kierunek |
|--------|---------------------|
| KOSZTOWE | Ujemny (wydatek), wyjątek: kolumna B > 0 → `prz_pr_in` |
| SPRZEDAZ | Dodatni (wpływ) |
| WLASC | Ujemny (wydatek) |

---

## 6. Automatyczne przypisywanie klucza księgowego

Program nadaje klucz automatycznie podczas parowania na podstawie:
1. Sekcji arkusza (KOSZTOWE / SPRZEDAZ / WLASC)
2. Kierunku transakcji (dodatnia / ujemna kwota)
3. Słów kluczowych mediów w nazwie kontrahenta lub tytule

**Słowa kluczowe mediów:** Netia, EON, E.ON, PGNiG, P4, Play, energia, prąd, gaz, internet, woda, SBM

---

## 7. Podsumowanie segmentów — struktura

Przycisk „Dodaj podsumowanie segmentów" generuje trzy tabele na dole aktywnego arkusza.

### Tabela 1 — Faktury główne (KOSZTOWE + SPRZEDAZ + WLASC)

| Kolumna | Źródło | Opis |
|---------|--------|------|
| Ilość pozycji | col A niepuste | Liczba faktur w sekcji |
| Suma kol. B (faktura) | col B, col A niepuste | Suma kwot brutto z faktur |
| Suma wyciąg_Kwota | col F, wszystkie wiersze | Suma dopasowanych transakcji bankowych |
| Suma RK | col B, klucz zawiera `_rk_` | Suma płatności gotówkowych |
| Bilans | = kol.B − wyciąg − RK | 0 = w pełni pokryte |

> **Bilans = 0** oznacza że wszystkie faktury mają pokrycie w przelewach lub gotówce.
> Bilans ≠ 0 sygnalizuje faktury bez płatności lub nadpłaty.

### Tabela 2 — Pozostałe (INNE_RK + NIEZNANE)

| Kolumna | Opis |
|---------|------|
| klucz_kos_kolB | Suma col B gdzie klucz zaczyna się od `kos_` |
| klucz_prz_kolB | Suma col B gdzie klucz zaczyna się od `prz_` |
| klucz_roz_kolB | Suma col B gdzie klucz zaczyna się od `roz_` |

> Dla sekcji NIEZNANE źródłem jest `wyciag_Kwota` (col F) zamiast col B,
> bo wiersze NIEZNANE nie mają faktur (col B pusta).
> Kategoryzacja nadal po prefiksie klucza, fallback: ujemne → kos_, dodatnie → prz_.

### Tabela 3 — Matryca kluczy × segmentów

Wiersze: `Koszty (kos_)` | `Przychody (prz_)` | `Rozrachunkowe (roz_)`

Kolumny: każdy z 5 segmentów + `Bilans` (suma wiersza)

Daje obraz gdzie w arkuszu są koszty, przychody i rozrachunki.

### Wiersz końcowy — Bilans miesiąca

```
Bilans miesiąca prz-kos MMYYYY = mat_bil[prz_] + mat_bil[kos_]
```

Jasno zielony. Suma wszystkich przychodów i kosztów ze wszystkich segmentów.

### Wiersze diagnostyczne (powyżej tabel)

| Wiersz | Znaczenie | Kolor gdy OK |
|--------|-----------|--------------|
| Status 0 | Wiersze niezweryfikowane (czekają na wpisanie 1) | Zielony gdy = 0 |
| Status 1 | Wiersze gotowe, niesprawdzone jeszcze po parowaniu | Zielony gdy = 0 |
| Status Null | Wiersze bez żadnego statusu (pominięte przez użytkownika) | Żółty gdy = 0 |
| Klucz Null | Wiersze z pustym kluczem księgowym | Żółty gdy = 0 |
| Klucz Nieznane | Wiersze z kluczem `nieznany_*` | Żółty gdy = 0 |

---

## 8. Struktura plików i folderów

### Google Drive

```
Faktury/                              ← FOLDER_ID
├── 032026/                           ← podfolder miesięczny (MMRRRR)
│   ├── [FVS] Kosztowe/               ← faktury kosztowe PDF
│   └── [FVS] Sprzedaz/               ← faktury sprzedaży PDF (scalony lub osobne)
├── 042026/
│   └── ...
├── Listy_operacji_abido/             ← wyciągi bankowe XLS
│   ├── lista_operacji_032026.xls
│   └── lista_operacji_042026.xls
└── Mieszkania/                       ← dane najemców (Google Sheets)
```

**Konwencja nazw:**
- Podfolder miesięczny: `MMRRRR` (np. `032026` = marzec 2026)
- Wyciąg bankowy: `lista_operacji_MMRRRR.xls`
- Faktura gotówkowa: nazwa pliku zawiera słowo `cash`

### Google Sheets — arkusz główny

Jeden arkusz (zakładka) = jeden miesiąc. Nazwa zakładki = nazwa podfolderu (np. `032026`).

**Kolumny (0-indexed):**

| Idx | Kolumna | Nazwa | Kto wypełnia |
|-----|---------|-------|--------------|
| 0 | A | Nazwa / Plik | Program (z nazwy PDF) |
| 1 | B | Kwota brutto | Program (z PDF), zawsze ujemna dla kosztów |
| 2 | C | Status | Użytkownik ręcznie (0/1/2/3) |
| 3 | D | Klucz_Ksiegowy | Program (przy parowaniu) |
| 4 | E | wyciag_Kontrahent | Program (parowanie) |
| 5 | F | wyciag_Kwota | Program (parowanie) |
| 6 | G | Data_ksiegowania | Program (parowanie) |
| 7 | H | wyciag_Tytul | Program (parowanie) |
| 8 | I | wyciag_Data_op | Program (parowanie) |
| 9 | J | wyciag_Rodzaj | Program (parowanie) |
| 10 | K | wyciag_Waluta | Program (parowanie) |
| 11 | L | wyciag_Nr_rachunku | Program (parowanie) |
| 12 | M | wyciag_Imie_Nazwisko | Program (parowanie) |
| 13 | N | Uwagi | Program (auto) lub użytkownik |

### Google Sheets — arkusz najemców

- ID: `ABIDO_NAJEMCY_ID`
- Zakładka: `Aktualni najemcy`
- Kolumny odczytywane po nazwie nagłówka (nie po indeksie)
- Filtr: tylko wiersze gdzie `Status do tworzenia == 1.0`

---

## 9. Sekcje arkusza

| Separator | Kolor | Zawartość |
|-----------|-------|-----------|
| `--- FAKTURY KOSZTOWE ---` | Czerwony | Wydatki firmy |
| `--- FAKTURY SPRZEDAZY NAJEMCOM ---` | Zielony | Przychody od najemców |
| `--- FAKTURY WLASCICIELE I SPOLDZIELNIE ---` | Pomarańczowy | Przelewy do właścicieli |
| `--- INNE RAPORTY KASOWE ---` | Turkus | Operacje gotówkowe bez faktury |
| `--- NIEZNANE / NIESPAROWANE Z WYCIAGU ---` | Szary | TX bankowe bez pary |

**Kolejność sekcji jest zawsze stała** — program odtwarza ją przy każdym zapisie.

---

## 10. Typowy workflow miesięczny

```
1. Wgraj faktury PDF do [FVS] Kosztowe w podfolderze np. 032026
2. Kliknij „Zaczytaj faktury kosztowe" → wiersze status=0 w arkuszu
3. Sprawdź każdą fakturę:
   - Popraw kwotę w kolumnie B jeśli program źle odczytał
   - Wpisz „1" w kolumnie C gdy faktura potwierdzona
4. Kliknij „Tworz wiersze faktur sprzedazy" → wiersze najemców status=1
5. Pobierz wyciąg bankowy XLS, wgraj do Listy_operacji_abido
6. Kliknij „Paruj wyciag bankowy" → kolumny D-N wypełnione
7. Sprawdź dopasowania:
   - Jeśli OK → wpisz „2" w kolumnie C (zatwierdzenie)
   - Jeśli błąd → wpisz „9" → przy następnym parowaniu szuka ponownie
8. W połowie miesiąca wgraj nowy (większy) wyciąg → paruj ponownie
   - Status 2/3 → pomijane
   - Status 1 bez dopasowania → próbuje ponownie
9. Kliknij „Dodaj podsumowanie segmentów" → kontrola bilansu miesiąca
10. Gdy Bilans = 0 i Status 0/1/Null = 0 → miesiąc zamknięty
```

---

## 11. Środowisko techniczne

| Element | Wartość |
|---------|---------|
| Platforma | Streamlit Cloud |
| Konto Streamlit | `abidom789-bot` |
| Repo GitHub | `abidom789-bot/Abido-system-fakturowania` |
| Branch lokalny | `master` → push do `main` na GitHub |
| Secrets | `[gcp_service_account]` — JSON Service Account Google |
| Google Cloud | `regal-river-494622-c7` |
| Service Account | `fakturowanie-bot@regal-river-494622-c7.iam.gserviceaccount.com` |
| SPREADSHEET_ID | `1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0` |
| FOLDER_ID | `1kwY6tOalKS2jnidABw6uUV23ykMj1iR2` |
| ABIDO_NAJEMCY_ID | `1TuHpPvdZmGN_kXbAuhdA72hs8AKxaiOLQrOUpXh3uYA` |
