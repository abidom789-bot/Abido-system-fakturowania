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

Zmiana względem poprzedniego kodu: status 2 musi być traktowany tak samo jak status 1 w funkcji apply_sync_logic.

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

## Algorytm parowania (projekt)

### Wejście
- Arkusz Google Sheets: wiersze ze statusem 1 (z kolumn A i B)
- Plik XLS z wyciągiem bankowym (z Google Drive, folder Listy_operacji_abido)

### Dopasowanie wiersza faktury do transakcji bankowej
1. Weź kwotę z kol B (np. -824.10)
2. Szukaj w wyciągu transakcji gdzie |kwota| == |B| (dokładne dopasowanie)
3. Jeśli jeden wynik → dopasuj
4. Jeśli wiele wyników o tej samej kwocie → użyj słów kluczowych:
   - Wyodrębnij słowa z nazwy pliku (kol A)
   - Porównaj z "Dane kontrahenta" i "Tytuł operacji" z wyciągu
   - Dopasuj ten z najlepszym score'em słów kluczowych
5. Jeśli zero wyników → zostaw E-P puste (płatność w innym miesiącu lub gotówka)

### Transakcja już użyta
Każda transakcja z wyciągu może być dopasowana tylko raz.
Jeśli dwie faktury walczą o tę samą transakcję → przydziel tej z lepszym dopasowaniem słów kluczowych, drugą zostaw jako niedopasowaną.

### Klucz_Ksiegowy (kol E) — zasady auto-wyznaczania
Zależy od: sekcji arkusza + kierunku transakcji + rodzaju płatności

| Sekcja | Kierunek | Rodzaj | Klucz |
|--------|----------|--------|-------|
| Kosztowe | wychodzący | przelew/karta | kos_pr_out |
| Kosztowe | wychodzący | media (Netia/EON/PGNiG/Play) | kos_med_pr_out |
| Kosztowe | przychodzący | przelew | kos_pr_in |
| Kosztowe | brak wyciągu | BLIK/gotówka wydatek | kos_rk_kw |
| Kosztowe | brak wyciągu | BLIK/gotówka przychód | kos_rk_kp |
| Sprzedaz (najem) | przychodzący | przelew | prz_naj_pr_in |
| Sprzedaz (najem) | brak wyciągu | gotówka | prz_naj_rk_kp |
| Wlasciciele | wychodzący | przelew | wla_pr_out |

Słowa kluczowe mediów: Netia, EON, E.ON, PGNiG, Play, P4

## Wielokrotne parowanie (ten sam miesiąc)

- Wyciąg importowany np. 10. i 20. dnia miesiąca
- Przy ponownym parowaniu:
  - Status 2 → pomijany (już zablokowany)
  - Status 1 bez dopasowania → próbuj ponownie z nowym (większym) wyciągiem
  - Status 1 z dopasowaniem → czy nadpisywać? TBD — prawdopodobnie NIE (użytkownik już mógł sprawdzić)

## Otwarte pytania (do uzupełnienia)

- [ ] Q4b: Co z transakcjami z wyciągu które nie mają żadnego odpowiednika w arkuszu?
       (np. Allegro 639,20 — brak faktury w folderze, brak wiersza w arkuszu)
       Opcja A: dodaj jako nowy wiersz ze statusem 0 i pustym A
       Opcja B: ignoruj
- [ ] Q5: Które sekcje są parowane? Tylko kosztowe, czy też najem i właściciele?
- [ ] Skąd brana jest nazwa pliku XLS wyciągu? Stała nazwa, czy program szuka najnowszego?
