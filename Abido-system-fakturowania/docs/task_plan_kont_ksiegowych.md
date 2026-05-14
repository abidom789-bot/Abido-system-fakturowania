# TASK: Plan kont księgowych — analiza i propozycja

> Status: DO ANALIZY — nie kodować dopóki nie zatwierdzone
> Cel: zastąpić obecny klucz (kos_pr_out, prz_naj_pr_in itp.)
>      profesjonalnym zapisem księgowym wg standardów polskiej rachunkowości

---

## Zatwierdzone decyzje projektowe

### 1. Jeden wiersz = jedna faktura (architektura bez zmian)
Kolumny A+B = faktura, kolumny E-N = sparowana transakcja z wyciągu. Bez podziału na FZ+WB.

### 2. Słownik kontrahentów — zakładka w głównym arkuszu GSheets
Zakładka np. `"Kontrahenci"` w `SPREADSHEET_ID`. Zawiera:
- Kody dostawców (Castorama → 201-CAST, Netia → 201-NETI...)
- Kody najemców i mieszkań — pobierane z arkusza `"Aktualni najemcy"` w `ABIDO_NAJEMCY_ID`
  (tam już będą kody mieszkań i kody najemców)
- Wszystko w jednym miejscu: zakupy, media, najemcy, właściciele, ZUS, US

### 3. Automatyczne rozpoznawanie kont przez program
Program na podstawie nazwy pliku PDF + kontrahenta z wyciągu nadaje Konto_Wn i Konto_Ma automatycznie.
Użytkownik sprawdza wyniki i zgłasza błędy.

### 4. Pętla uczenia się — historia uwag i korekt (KLUCZOWE)
Mechanizm ciągłego doskonalenia rozpoznawania:

```
[Program nadaje konta automatycznie]
         ↓
[Użytkownik sprawdza w arkuszu]
         ↓
[Użytkownik zgłasza błędy do Claude]
         ↓
[Claude zapisuje błędy do pliku historii: docs/historia_uwag_kont.md]
         ↓
[Claude poprawia logikę rozpoznawania w kodzie]
         ↓
[Historia pozostaje na zawsze w projekcie jako dokumentacja decyzji]
```

Plik historii: `docs/historia_uwag_kont.md`
- Każda korekta zapisywana z datą, opisem błędu i zmianą w kodzie
- Nigdy nie kasowany — jest dowodem dlaczego kod działa tak jak działa
- Przy nowej sesji Claude czyta ten plik i zna kontekst wszystkich poprzednich poprawek

---

## Decyzja architektoniczna (zatwierdzona — jeden wiersz)

**Jeden wiersz = jeden koszt/faktura.**
- Kolumny A i B = faktura (nazwa pliku PDF, kwota brutto)
- Kolumny E-N = sparowana transakcja z wyciągu bankowego (jak teraz)
- NIE dzielimy na dwa wiersze FZ + WB — wszystko zostaje w jednym wierszu

Konsekwencja: konta Wn/Ma opisują **istotę kosztu lub przychodu**, nie pełny zapis dwustronny.
Konto Wn = co się stało z pieniędzmi (jaki koszt/przychód)
Konto Ma = kto jest kontrahentem (z kogo zobowiązania / na czyją korzyść)

---

## Symbole dokumentów (standard rynkowy)

| Symbol | Nazwa | Kiedy używany w tym systemie |
|--------|-------|------------------------------|
| **FZ** | Faktura Zakupu | każda faktura kosztowa (materiały, usługi, media) |
| **FS** | Faktura Sprzedaży | faktura wystawiona najemcy (czynsz miesięczny) |
| **WB** | Wyciąg Bankowy | wiersz bez faktury — transakcja z wyciągu bez pary (SEP_NIEZNANE, bankomat) |
| **KP** | Kasa Przyjmie | wpłata gotówkowa (czynsz gotówką, kaucja gotówką) |
| **KW** | Kasa Wypłaci | wypłata gotówkowa (zakup za gotówkę, faktura "cash") |
| **PK** | Polecenie Księgowania | wynagrodzenia, ZUS, korekty ręczne |

---

## Propozycja planu kont

### Zespół 1 — Środki pieniężne
```
100   Kasa (gotówka)
131   Rachunek bankowy (mBank bieżący)
```

### Zespół 2 — Rozrachunki
```
200   Rozrachunki z najemcami (należności za czynsz — strona FS)
      200-M01        Mieszkanie 01 (np. Bazylińska 5/3)
      200-M01-01     Najemca bieżący w M01
      200-M01-02     Kolejny najemca (zmiana lokatora)
      200-M02        Mieszkanie 02
      200-M02-01     Najemca w M02
      ... itd dla każdego mieszkania

201   Rozrachunki z dostawcami (zobowiązania za FZ)
      201-CAST       Castorama
      201-LERO       Leroy Merlin
      201-AGAT       Agata S.A.
      201-PEPC       PEPCO
      201-ALLE       Allegro
      201-OBI        OBI
      201-NETI       Netia
      201-EONS       E.ON / EON Sprzedaż
      201-PGNIG      PGNiG / Polska Spółka Gazownictwa
      201-PLAY       Play / P4
      201-MAFI       Mafika (lub inny dostawca usług)
      201-INNE       Inny dostawca (do uzupełnienia)

225   Rozrachunki publicznoprawne — podatki
      225-CIT        CIT / podatek dochodowy
      225-PON        Podatek od nieruchomości
      225-VAT        VAT (jeśli płatnik)

229   Rozrachunki z ZUS
      229-ZUS        Zakład Ubezpieczeń Społecznych

234   Rozrachunki z pracownikami / zleceniobiorcami
      234-ZAP        Kajetan Zapała
      234-DYB        Milena Dybalska-Stypułko

240   Kaucje / depozyty najemców (zobowiązanie — do zwrotu)
      240-M01-01     Kaucja najemcy M01-01
      240-M02-01     Kaucja najemcy M02-01
      ... itd

250   Rozrachunki z właścicielami / spółdzielniami
      250-DOMHUT     DOMHUT sp. z o.o.
      250-SBM        Spółdzielnia Budowlano-Mieszkaniowa
      250-WL01       Właściciel 01
      250-WL02       Właściciel 02
```

### Zespół 4 — Koszty wg rodzajów
```
401   Zużycie materiałów i energii
      401-MAT        Materiały budowlane i remontowe (farba, narzędzia)
      401-MED-EL     Media — prąd (EON, Enea)
      401-MED-GAZ    Media — gaz (PGNiG, PSG)
      401-MED-INT    Media — internet / telefon (Netia, Play)
      401-MED-WOD    Media — woda / ścieki

402   Usługi obce
      402-REM        Remonty i naprawy (hydraulik, elektryk)
      402-SPRZ       Sprzątanie i porządkowanie
      402-ADM        Usługi administracyjne / biurowe
      402-INNE       Inne usługi obce

403   Podatki i opłaty
      403-CIT        CIT
      403-PON        Podatek od nieruchomości
      403-INNE       Inne opłaty publiczne

404   Wynagrodzenia
      404-ZAP        Kajetan Zapała
      404-DYB        Milena Dybalska-Stypułko

405   Składki ZUS (część pracodawcy)
```

### Zespół 7 — Przychody
```
700   Przychody z najmu
      700-M01        Mieszkanie 01 (ogółem)
      700-M01-01     Czynsz najemca M01-01
      700-M02-01     Czynsz najemca M02-01
      ... itd
```

---

## Przykłady transakcji — jeden wiersz, zapis Wn/Ma

### FZ — Faktura kosztowa, zakup materiałów (Castorama, farba)
```
Dok:      FZ
Konto Wn: 401-MAT
Konto Ma: 201-CAST
Kontrah.: 201-CAST (Castorama)
WB:       przelew wychodzący widoczny w kolumnach E-N
```

### FZ — Faktura kosztowa, media — prąd (EON)
```
Dok:      FZ
Konto Wn: 401-MED-EL
Konto Ma: 201-EONS
WB:       przelew wychodzący E-N
```

### FZ — Faktura kosztowa, media — internet (Netia)
```
Dok:      FZ
Konto Wn: 401-MED-INT
Konto Ma: 201-NETI
WB:       przelew wychodzący E-N
```

### FZ — Faktura kosztowa, media — gaz (PGNiG)
```
Dok:      FZ
Konto Wn: 401-MED-GAZ
Konto Ma: 201-PGNIG
WB:       przelew wychodzący E-N
```

### FZ — Faktura kosztowa, usługi remontowe
```
Dok:      FZ
Konto Wn: 402-REM
Konto Ma: 201-INNE  (lub konkretny kod wykonawcy)
WB:       przelew wychodzący E-N
```

### PK — Wynagrodzenie (Kajetan Zapała)
```
Dok:      PK
Konto Wn: 404-ZAP
Konto Ma: 234-ZAP
WB:       przelew wychodzący E-N
```

### PK — ZUS
```
Dok:      PK
Konto Wn: 405
Konto Ma: 229-ZUS
WB:       przelew wychodzący E-N
```

### PK — Podatek (US / Urząd Skarbowy)
```
Dok:      PK
Konto Wn: 403-CIT  lub  403-PON
Konto Ma: 225-CIT  lub  225-PON
WB:       przelew wychodzący E-N
```

### KW — Zakup gotówkowy (faktura "cash", np. Leroy Merlin)
```
Dok:      KW
Konto Wn: 401-MAT  (lub 402-REM)
Konto Ma: 100       (kasa — wypłata gotówki)
WB:       brak (brak pary z wyciągu — kolumny E-N puste)
```

### FS — Faktura sprzedaży, czynsz od najemcy (przelew)
```
Dok:      FS
Konto Wn: 200-M03-01   (należność od najemcy Baranovska, M03)
Konto Ma: 700-M03-01   (przychód z najmu M03)
WB:       przelew przychodzący E-N
```

### FS — Faktura sprzedaży, czynsz od najemcy (gotówka)
```
Dok:      FS / KP
Konto Wn: 200-M05-01
Konto Ma: 700-M05-01
WB:       brak (brak pary z wyciągu)
```

### WB — Przelew do właściciela / spółdzielni (DOMHUT)
```
Dok:      WB
Konto Wn: 250-DOMHUT
Konto Ma: 131
(brak faktury w kolumnie A — wiersz tworzony ręcznie lub przez SEP_WLASC)
```

### WB — Kaucja, wpłata od najemcy
```
Dok:      WB / KP
Konto Wn: 131  lub  100
Konto Ma: 240-M03-01   (zobowiązanie — kaucja do zwrotu)
WB:       przelew przychodzący E-N
```

### WB — Kaucja, zwrot do najemcy
```
Dok:      WB
Konto Wn: 240-M03-01
Konto Ma: 131
WB:       przelew wychodzący E-N
```

### WB — Bankomat (SEP_INNE_RK)
```
Dok:      WB
Konto Wn: 100    (gotówka weszła do kasy)
Konto Ma: 131    (wyszła z rachunku bankowego)
WB:       przelew wychodzący E-N
```

### FZ — Faktura z płatnością w następnym miesiącu
```
Dok:      FZ
Konto Wn: 401-...
Konto Ma: 201-...
WB:       brak (przelew będzie w kolejnym miesiącu)
Uwagi:    kos_fak_no_pay
```

### FZ — Zwrot / refaktura (kwota dodatnia w kosztowych)
```
Dok:      FZ / NK
Konto Wn: 131
Konto Ma: 401-...   (zmniejszenie kosztu)
WB:       przelew przychodzący E-N
```

---

## Nowe kolumny — propozycja

Obecne kolumny A-N (14 kolumn). Propozycja dołożenia 3 na końcu:

| Kolumna | Idx | Nazwa | Przykład |
|---------|-----|-------|---------|
| O | 14 | Dok | FZ / FS / WB / KP / KW / PK |
| P | 15 | Konto_Wn | 401-MAT, 200-M03-01 |
| Q | 16 | Konto_Ma | 201-CAST, 700-M03-01 |

Kolumna N (Uwagi, idx=13) zostaje na swoim miejscu.

---

## Pytania do rozstrzygnięcia

1. **Ile masz mieszkań i jakie mają nazwy / adresy?**
   Żeby ustalić kody M01, M02... i przypisać je do adresów z arkusza najemców.

2. **Słownik kontrahentów** — gdzie trzymać mapowanie "Netia" → "201-NETI"?
   a) Osobna zakładka w Google Sheets (edytowalna bez deployu) — **rekomendowane**
   b) Stała w kodzie (szybciej działa, ale zmiana = deploy na Streamlit)

3. **Automatyczne rozpoznawanie** — czy program ma sam nadawać Konto_Wn/Ma
   na podstawie nazwy pliku PDF / kontrahenta z wyciągu,
   czy użytkownik wpisuje ręcznie po sparowaniu?

4. **Czy obecny klucz** (kos_pr_out itp.) ma zostać jako pole pomocnicze
   obok nowych kont, czy całkowicie zniknąć?

5. **Symbol Dok** — czy ma być nadawany automatycznie przez program
   (FZ dla kosztowych, FS dla sprzedaży) czy wpisywany ręcznie?
