# Logika statusów i parowania wyciągu bankowego

## System statusów (kolumna C)

| Status | Znaczenie | Zachowanie przy "Zaczytaj faktury kosztowe" | Zachowanie przy "Paruj wyciąg" |
|--------|-----------|---------------------------------------------|-------------------------------|
| 0 | Niezweryfikowany — wiersz świeżo dodany lub niepewny | Nadpisywany (re-read z PDF) | NIE parowany — pomijany |
| 1 | Zweryfikowany przez użytkownika — faktura potwierdzona | Nietykalny | Parowany z wyciągiem bankowym |
| 2 | Sparowany i zatwierdzony — użytkownik potwierdził dopasowanie | Nietykalny | Nietykalny — nigdy się nie zmienia |

## Przepływ pracy użytkownika

1. **Zaczytaj faktury kosztowe** → nowe pliki z Drive lądują ze statusem 0
2. Użytkownik sprawdza fakturę → wpisuje **1** (potwierdzam, wiem co to jest)
   - Nawet jeśli plik nie ma jeszcze odpowiednika na wyciągu (spóźniona faktura) → status 1 chroni wiersz
3. **Paruj wyciąg** (np. 10. dnia miesiąca) → paruje wszystkie wiersze ze statusem 1
4. Użytkownik sprawdza czy dopasowania są poprawne
5. Jeśli dopasowanie dobre → wpisuje **2** → wiersz zablokowany na zawsze
6. **Paruj wyciąg** ponownie (np. 20. dnia miesiąca) → wyciąg jest większy, więcej wierszy ze statusem 1 dostaje dopasowanie
7. Wiersze ze statusem 2 są zawsze pomijane — nie dotykamy ich nigdy

## Zasada ochrony przy "Zaczytaj faktury kosztowe"

```
Status 0 → re-read z PDF (nadpisz)
Status 1 → nietykalny
Status 2 → nietykalny
```

## Kolumny arkusza po wdrożeniu parowania

| Kol | Nazwa | Kto wypełnia |
|-----|-------|--------------|
| A | Nazwa / Plik | Zaczytaj faktury kosztowe |
| B | Kwota brutto | Zaczytaj faktury kosztowe (z PDF) |
| C | Status (0/1/2) | Użytkownik ręcznie |
| D | Adres | Tworz wiersze sprzedazy |
| E | Klucz_Ksiegowy | Paruj wyciąg (auto) |
| F | wyciag_Dane kontrahenta | Paruj wyciąg |
| G | wyciag_Kwota | Paruj wyciąg |
| H | Kwota_raport_kasowy | Paruj wyciąg (tylko BLIK/gotówka) |
| I | Data_ksiegowania | Paruj wyciąg |
| J | wyciag_Tytul operacji | Paruj wyciąg |
| K | wyciag_Data operacji | Paruj wyciąg |
| L | wyciag_Rodzaj operacji | Paruj wyciąg |
| M | wyciag_Waluta | Paruj wyciąg |
| N | wyciag_Numer rachunku | Paruj wyciąg |
| O | wyciag_Imie_Nazwisko | Paruj wyciąg |
| P | Uwagi | Paruj wyciąg (auto) lub użytkownik |

---

## Algorytm parowania — wieloprzebiegowy

### Wejście
- Arkusz Google Sheets: wiersze ze statusem 1 z kolumn A i B
- Plik XLS z wyciągiem bankowym (z Google Drive, folder Listy_operacji_abido)

### Zasada naczelna
Każda transakcja z wyciągu może być dopasowana **tylko raz**.
Pozycje z wyciągu które nie zostały sparowane — trafiają na dół arkusza jako osobne wiersze (A i B puste, kolumny F+ wypełnione). ZAKAZ usuwania jakichkolwiek pozycji z wyciągu.

### Specjalne przypadki kontrahentów — wyciąg bankowy

Dwie grupy transakcji wymagają szczególnego traktowania bo kontrahent NIE jest rzeczywistą osobą:

**1. Nest Bank S.A.** — wszystkie płatności kartą przechodzą przez Nest Bank
   - Kontrahent = "Nest Bank S.A." (bezużyteczny do dopasowania)
   - Rzeczywisty sprzedawca jest w **Tytule operacji**: "Allegro Poznan Nr karty...", "AGATA S.A. KIELCE..."
   - Dopasowanie: wyodrębnij pierwsze słowo z tytułu → porównaj z nazwą faktury

**2. Revolut Bank UAB** — przelewy zagraniczne przez Revolut
   - Kontrahent = "Revolut Bank UAB" (bezużyteczny)
   - Rzeczywista osoba jest w **Tytule operacji**: "Palash Tripathi Sent from Revolut", "CHIADIKA OBENWA..."
   - Dopasowanie: wyodrębnij imię/nazwisko z tytułu → porównaj z kol A

### Przebieg 1 — Nazwisko + kwota (najsilniejsze dopasowanie)

Dla każdego wiersza ze statusem 1:
1. Wyodrębnij tokeny z kol A (słowa, minimum 3 znaki, bez cyfr)
2. Znajdź token który wygląda jak nazwisko (ostatni token lub token po którym nie ma cyfr)
3. Szukaj w wyciągu: transakcje gdzie |kwota_wyciag| == |kwota_B| ORAZ nazwisko występuje w polu Dane_kontrahenta lub Tytul
   - Uwaga: dla Nest Bank i Revolut → szukaj w Tytule
4. Jeśli dokładnie jeden wynik → dopasuj (pewne)

### Przebieg 2 — Imię + kwota

Dla wierszy niedopasowanych po Przebiegu 1:
1. Wyodrębnij pierwsze słowo z kol A jako potencjalne imię
2. Szukaj w wolnych transakcjach: |kwota_wyciag| == |kwota_B| ORAZ imię w kontrahent/tytul
3. Jeśli dokładnie jeden wynik → dopasuj (pewne)

### Przebieg 3 — Słowa kluczowe z nazwy faktury + kwota

Dla wierszy niedopasowanych po Przebiegu 2 (głównie kosztowe — pliki PDF):
1. Wyodrębnij znaczące słowa z nazwy pliku (kol A): np. "Allegro", "Agata", "Netia", "EON", "PGNiG", "Play"
2. Szukaj w wolnych transakcjach: |kwota_wyciag| == |kwota_B| ORAZ dowolne słowo kluczowe w tytule/kontrahencie
3. Dopasuj wynik z największą liczbą pasujących słów kluczowych
4. Jeśli remis (dwa wyniki z tym samym score) → zostaw jako niedopasowany (wymaga ręcznego przeglądu)

### Przebieg 4 — Sama kwota (ostatnia deska ratunku)

Dla wierszy ciągle niedopasowanych:
1. Szukaj w wolnych transakcjach tylko po kwocie: |kwota_wyciag| == |kwota_B|
2. Jeśli dokładnie jeden wynik → dopasuj, ale w kolumnie Uwagi wpisz "Dopasowanie tylko po kwocie — sprawdź"
3. Jeśli zero lub wiele wyników → zostaw niedopasowany

### Po wszystkich przebiegach

- Niedopasowane wiersze arkusza (status 1, brak pary): kolumny E-P pozostają puste
- Niedopasowane transakcje z wyciągu: dopisywane na dół arkusza (A i B puste, E-P wypełnione z wyciągu)

---

## Przypadki wielokrotnych przelewów od tej samej osoby (z danych marcowych)

Przypadki zaobserwowane w lista_operacji_032026.xls — wymagają szczególnej uwagi:

### Przypadek A — różne kwoty (bezproblemowy)
Jeden kontrahent, wiele przelewów, ale każdy innej kwoty.
Dopasowanie po nazwisku + kwocie działa jednoznacznie.

Przykłady:
- BARANOVSKA OKSANA: 3 przelewy (1100 kaucja, 450 pokój Bazylińska, 1150 pokój Umińskiego)
- HELMAN DMYTRO: 2 przelewy (1700 czynsz, 510.49 woda)
- GRZEGORZ KWASIBORSKI: 5 przelewów wszystkie różne kwoty
- JOLANTA KOWALCZYK: 3 przelewy różne kwoty
- ANNA MOSSAKOWSKA: 3 przelewy różne kwoty
- DOMHUT (spółdzielnia): 2 przelewy (-517.77 woda, -1583.15 czynsz)

Reguła: jeśli arkusz ma np. "Baranovska Oksana" z kwotą -1150 → Przebieg 1 dopasuje po nazwisku + kwocie jednoznacznie.

### Przypadek B — ta sama kwota od tej samej osoby (trudny)
Ten sam kontrahent, ta sama kwota. Algorytm nie może rozróżnić po kwocie.

Przykłady z danych:
- **SEBASTIAN OSETEK**: 2x 4000 PLN, oba z tytułem "Czynsz" (daty: 31-03 i 03-03)
- **IVANNA PASHKO**: 2x 1300 PLN, oba z tytułem "Przelew środków z miszkanie za grydzie"

Reguła dla przypadku B:
1. Jeśli arkusz ma JEDEN wiersz "Sebastian Osetek" z kwotą 4000 → dopasuj do PIERWSZEJ (chronologicznie wcześniejszej) transakcji. Druga transakcja idzie na dół jako niedopasowana.
2. Jeśli arkusz ma DWA wiersze "Sebastian Osetek" z kwotą 4000 → pierwszy wiersz = pierwsza transakcja, drugi wiersz = druga transakcja (dopasowanie po kolejności).
3. W kolumnie Uwagi wpisz: "Uwaga: wiele przelewów tej samej kwoty — sprawdź datę"

### Przypadek C — kontrahent pośredni (Nest Bank / Revolut)

Płatności kartą — kontrahent to zawsze "Nest Bank S.A.", faktyczny sklep jest w tytule:
- "Allegro Poznan Nr karty ...3752 119,00PLN" → słowo kluczowe: Allegro
- "AGATA S.A. KIELCE KIELCE Nr karty ...5279 131,00PLN" → słowo kluczowe: Agata
- "PEPCO 110139 WARSZAWA" → słowo kluczowe: Pepco
- "MARKET OBI 020 KIELCE" → słowo kluczowe: Obi, OBI

Przelewy via Revolut — kontrahent to "Revolut Bank UAB", faktyczna osoba w tytule:
- "Palash Tripathi Sent from Revolut" → nazwisko: Tripathi
- "CHIADIKA OBENWA Afrykanska pokoj .1 za Marzec 2026" → nazwisko: Chiadika lub Obenwa

Reguła: dla tych kontrahentów Przebieg 1 i 2 muszą szukać w polu Tytul_operacji, nie w Dane_kontrahenta.

---

## Klucz_Ksiegowy (kol E) — zasady auto-wyznaczania

Zależy od: sekcji arkusza + kierunku transakcji + rodzaju płatności

| Sekcja | Kierunek | Rodzaj | Klucz |
|--------|----------|--------|-------|
| Kosztowe | wychodzący | przelew/karta | kos_pr_out |
| Kosztowe | wychodzący | media | kos_med_pr_out |
| Kosztowe | przychodzący | przelew/zwrot | kos_pr_in |
| Kosztowe | brak wyciągu | gotówka wydatek | kos_rk_kw |
| Kosztowe | brak wyciągu | gotówka przychód | kos_rk_kp |
| Sprzedaz (najem) | przychodzący | przelew | prz_naj_pr_in |
| Sprzedaz (najem) | brak wyciągu | gotówka | prz_naj_rk_kp |
| Wlasciciele | wychodzący | przelew | wla_pr_out |
| Wlasciciele | wychodzący | media | wla_med_pr_out |
| Niedopasowana z wyciągu | wychodzący | - | nieznany_out |
| Niedopasowana z wyciągu | przychodzący | - | nieznany_in |

Słowa kluczowe mediów (decydują o kos_med vs kos): Netia, EON, E.ON, PGNiG, Play, P4, prąd, gaz, energia, internet

---

## Wielokrotne parowanie (ten sam miesiąc)

- Wyciąg importowany np. 10. i 20. dnia miesiąca
- Przy ponownym parowaniu:
  - Status 2 → pomijany (nietykalny)
  - Status 1 bez dopasowania → próbuj ponownie z nowym (większym) wyciągiem
  - Status 1 z istniejącym dopasowaniem (kol F niepuste) → NIE nadpisuj (użytkownik mógł już sprawdzić)
- Wiersze "niedopasowane z wyciągu" na dole: przy kolejnym parowaniu są czyszczone i zastępowane nową listą niedopasowanych

---

## Odpowiedzi na pytania konfiguracyjne

- [x] Q5: Parowane są **wszystkie sekcje** — kosztowe, najem (sprzedaż) i właściciele
- [x] Nazwa pliku XLS: zawsze `lista_operacji_MMRRRR.xls` np. `lista_operacji_032026.xls`
       Program szuka pliku po nazwie skonstruowanej z numeru miesiąca wpisanego przez użytkownika
- [x] Folder: `Listy_operacji_abido` jest bezpośrednio w folderze głównym `Faktury` (FOLDER_ID)
       Szukamy go przez `find_subfolder(service, FOLDER_ID, "Listy_operacji_abido")`
