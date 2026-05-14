# TASK: Plan kont księgowych — analiza i propozycja

> Status: DO ANALIZY — nie kodować dopóki nie zatwierdzone
> Cel: zastąpić obecny klucz (kos_pr_out, prz_naj_pr_in itp.)
>      profesjonalnym zapisem księgowym wg standardów polskiej rachunkowości

---

## Symbole dokumentów (standard rynkowy)

| Symbol | Nazwa | Kiedy używany |
|--------|-------|---------------|
| **FZ** | Faktura Zakupu | każda faktura kosztowa (zakup materiałów, usług, mediów) |
| **FS** | Faktura Sprzedaży | faktura wystawiona najemcy (czynsz) |
| **WB** | Wyciąg Bankowy | przelew bankowy (wpływ lub wypływ z rachunku) |
| **KP** | Kasa Przyjmie | wpłata gotówkowa (czynsz gotówką, wpłata kaucji) |
| **KW** | Kasa Wypłaci | wypłata gotówkowa (wynagrodzenia, zakupy za gotówkę) |
| **PK** | Polecenie Księgowania | korekty, przeksięgowania, noty |
| **NK** | Nota Korygująca | korekta faktury |

---

## Propozycja planu kont

### Zespół 1 — Środki pieniężne
```
100   Kasa (gotówka)
131   Rachunek bankowy (mBank)
```

### Zespół 2 — Rozrachunki
```
200   Rozrachunki z najemcami (należności za czynsz)
      200-M01        Mieszkanie 01 (np. Bazylińska 5/3)
      200-M01-01     Najemca w mieszkaniu M01 (bieżący)
      200-M01-02     Następny najemca (zmiana lokatora)
      200-M02        Mieszkanie 02
      200-M02-01     Najemca w mieszkaniu M02
      ... itd dla każdego mieszkania

201   Rozrachunki z dostawcami (zobowiązania za FZ)
      201-CAST       Castorama
      201-LERO       Leroy Merlin
      201-AGAT       Agata S.A.
      201-PEPC       PEPCO
      201-ALLE       Allegro
      201-NETI       Netia
      201-EONS       E.ON / EON Sprzedaż
      201-PGNIG      PGNiG / Polska Spółka Gazownictwa
      201-PLAY       Play / P4
      201-INNE       Inny dostawca (do uzupełnienia)

225   Rozrachunki publicznoprawne — podatki
      225-CIT        CIT (podatek dochodowy)
      225-PON        Podatek od nieruchomości (US / Urząd Skarbowy)
      225-VAT        VAT (jeśli płatnik)

229   Rozrachunki publicznoprawne — ZUS
      229-ZUS        Zakład Ubezpieczeń Społecznych

234   Rozrachunki z pracownikami / zleceniobiorcami
      234-ZAP        Kajetan Zapała
      234-DYB        Milena Dybalska-Stypułko

240   Rozrachunki różne — kaucje / depozyty
      240-M01-01     Kaucja najemcy M01-01
      240-M02-01     Kaucja najemcy M02-01
      ... itd

250   Rozrachunki z właścicielami / spółdzielniami
      250-WL01       Właściciel 01 (imię/nazwisko)
      250-WL02       Właściciel 02
      250-DOMHUT     DOMHUT sp. z o.o.
      250-SBM        Spółdzielnia Budowlano-Mieszkaniowa
```

### Zespół 4 — Koszty wg rodzajów
```
401   Zużycie materiałów i energii
      401-MAT        Materiały budowlane, remontowe (farba, narzędzia)
      401-MED-EL     Media — prąd (EON, Enea)
      401-MED-GAZ    Media — gaz (PGNiG, PSG)
      401-MED-INT    Media — internet (Netia)
      401-MED-TEL    Media — telefon (Play)
      401-MED-WOD    Media — woda / ścieki (SBM)

402   Usługi obce
      402-REM        Remonty i naprawy
      402-SPRZ       Sprzątanie
      402-KSIEG      Usługi księgowe / administracyjne
      402-INNE       Inne usługi obce

403   Podatki i opłaty
      403-CIT        CIT
      403-PON        Podatek od nieruchomości
      403-INNE       Inne opłaty publiczne

404   Wynagrodzenia
      404-ZAP        Kajetan Zapała
      404-DYB        Milena Dybalska-Stypułko

405   Ubezpieczenia społeczne (ZUS pracodawcy)

409   Pozostałe koszty rodzajowe
```

### Zespół 7 — Przychody
```
700   Przychody ze sprzedaży usług najmu
      700-M01        Przychody z mieszkania M01
      700-M01-01     Czynsz najemca M01-01
      700-M02        Przychody z mieszkania M02
      ... itd

750   Przychody finansowe (odsetki, kaucje do rozliczenia)
```

---

## Przykłady transakcji z Twojego systemu — propozycja zapisu

### 1. Faktura kosztowa — zakup materiałów (farba, Castorama)
```
Dokument:  FZ
Opis:      Castorama — zakup farby
Konto Wn:  401-MAT        (koszt materiałów)
Konto Ma:  201-CAST       (zobowiązanie wobec Castoramy)
---
Dokument:  WB
Opis:      Zapłata FZ Castorama
Konto Wn:  201-CAST       (rozliczamy zobowiązanie)
Konto Ma:  131             (wychodzi z rachunku bankowego)
```

### 2. Faktura kosztowa — media (Netia, internet)
```
Dokument:  FZ
Opis:      Netia — faktura za internet MM/RRRR
Konto Wn:  401-MED-INT
Konto Ma:  201-NETI
---
Dokument:  WB
Opis:      Zapłata FZ Netia
Konto Wn:  201-NETI
Konto Ma:  131
```

### 3. Faktura kosztowa — media (EON, prąd)
```
Dokument:  FZ
Opis:      E.ON — faktura za prąd MM/RRRR
Konto Wn:  401-MED-EL
Konto Ma:  201-EONS
---
Dokument:  WB
Opis:      Zapłata FZ EON
Konto Wn:  201-EONS
Konto Ma:  131
```

### 4. Faktura kosztowa — wynagrodzenie (Kajetan Zapała)
```
Dokument:  FZ / PK
Opis:      Wynagrodzenie za kwiecień 2026 — Kajetan Zapała
Konto Wn:  404-ZAP        (koszt wynagrodzenia)
Konto Ma:  234-ZAP        (zobowiązanie wobec pracownika)
---
Dokument:  WB
Opis:      Wypłata wynagrodzenia Kajetan Zapała
Konto Wn:  234-ZAP
Konto Ma:  131
```

### 5. Faktura kosztowa — ZUS
```
Dokument:  PK
Opis:      ZUS za kwiecień 2026
Konto Wn:  405             (składki ZUS)
Konto Ma:  229-ZUS
---
Dokument:  WB
Opis:      Przelew ZUS
Konto Wn:  229-ZUS
Konto Ma:  131
```

### 6. Faktura kosztowa — podatek (US / Urząd Skarbowy)
```
Dokument:  PK
Opis:      CIT za 2026 / podatek od nieruchomości
Konto Wn:  403-CIT
Konto Ma:  225-CIT
---
Dokument:  WB
Opis:      Przelew podatku
Konto Wn:  225-CIT
Konto Ma:  131
```

### 7. Faktura sprzedaży — czynsz od najemcy (przelew)
```
Dokument:  FS
Opis:      Czynsz kwiecień 2026 — Baranovska Oksana, M03
Konto Wn:  200-M03-01     (należność od najemcy)
Konto Ma:  700-M03-01     (przychód)
---
Dokument:  WB
Opis:      Wpływ czynszu — Baranovska Oksana
Konto Wn:  131
Konto Ma:  200-M03-01     (rozliczamy należność)
```

### 8. Faktura sprzedaży — czynsz gotówką
```
Dokument:  FS
Opis:      Czynsz kwiecień 2026 — [Najemca], M05
Konto Wn:  200-M05-01
Konto Ma:  700-M05-01
---
Dokument:  KP
Opis:      Wpłata gotówkowa czynsz — [Najemca]
Konto Wn:  100             (kasa)
Konto Ma:  200-M05-01
```

### 9. Przelew do właściciela / spółdzielni (DOMHUT)
```
Dokument:  WB
Opis:      Przelew DOMHUT — czynsz administracyjny MM/RRRR
Konto Wn:  250-DOMHUT
Konto Ma:  131
```

### 10. Kaucja — wpłata od najemcy
```
Dokument:  KP / WB
Opis:      Kaucja — Baranovska Oksana, M03
Konto Wn:  131 / 100
Konto Ma:  240-M03-01     (zobowiązanie — kaucja do zwrotu)
```

### 11. Kaucja — zwrot do najemcy
```
Dokument:  WB
Opis:      Zwrot kaucji — [Najemca], M03
Konto Wn:  240-M03-01
Konto Ma:  131
```

### 12. Zakup gotówkowy (faktura "cash")
```
Dokument:  FZ
Opis:      [Sklep] — zakup gotówkowy
Konto Wn:  401-MAT / 402-REM
Konto Ma:  100             (wypływ z kasy, nie z banku)
```

### 13. Bankomat — podjęcie gotówki
```
Dokument:  WB
Opis:      Bankomat — podjęcie gotówki
Konto Wn:  100             (gotówka weszła do kasy)
Konto Ma:  131             (wyszła z banku)
```

### 14. Faktura kosztowa z płatnością w następnym miesiącu
```
Dokument:  FZ
Opis:      [Dostawca] — faktura MM/RRRR (płatność następny miesiąc)
Konto Wn:  401-...
Konto Ma:  201-...         (zobowiązanie pozostaje otwarte)
— brak WB w tym miesiącu —
```

### 15. Zwrot / refaktura (kwota dodatnia w kosztowych)
```
Dokument:  NK / FZ-korekta
Opis:      Zwrot od [Dostawca]
Konto Wn:  131
Konto Ma:  401-...         (zmniejszenie kosztu)
  lub
Konto Ma:  201-...         (rozliczenie nadpłaty)
```

---

## Pytania do analizy przed kodowaniem

1. **Ile masz mieszkań?** — żeby ustalić schemat numeracji M01, M02...
2. **Czy każde mieszkanie ma stały numer** (przypisany raz na zawsze) czy numerowane wg kolejności dodania?
3. **Słownik kontrahentów** — czy ma być:
   a) W osobnym arkuszu Google Sheets (edytowalny przez użytkownika)?
   b) W kodzie jako stała (szybciej, ale zmiana = deploy)?
   c) Generowany automatycznie z nazw plików PDF?
4. **Czy obecny klucz** (kos_pr_out itp.) ma zniknąć całkowicie, czy zostać jako pomocnicze pole?
5. **Ile kolumn dołożyć?** Propozycja:
   - Kolumna O: `Konto_Wn`
   - Kolumna P: `Konto_Ma`
   - Kolumna Q: `Kod_Kontrahenta`
   (N = Uwagi przesuwa się lub zostaje jako R)
6. **Czy chcesz pełny zapis dwustronny** (FZ + WB jako dwa osobne wiersze), czy jeden wiersz łączący fakturę z przelewem (jak teraz)?
7. **VAT** — czy firma jest płatnikiem VAT? Czy potrzebne konto 222/223?

---

## Uwaga dotycząca obecnego systemu

Obecny klucz (`kos_pr_out`, `prz_naj_pr_in` itp.) to uproszczony jednowymiarowy opis.
Profesjonalny zapis Wn/Ma to dwuwymiarowy: **skąd → dokąd** idą pieniądze.

Obecny system łączy fakturę i przelew w **jednym wierszu** (faktura PDF + dane z wyciągu w tych samych kolumnach).
W klasycznej księgowości to **dwa oddzielne zdarzenia** (dwa wiersze: FZ i WB).

To jest kluczowa decyzja architektoniczna do podjęcia przed kodowaniem.
