# Przykłady wierszy arkusza po wprowadzeniu kont księgowych

> Nowe kolumny dodawane na końcu: O = Dok, P = Konto_Wn, Q = Konto_Ma
> Kolumny A-N pozostają bez zmian.

---

## Przykład 1 — FAKTURA KOSZTOWA (zakup materiałów, Allegro)

| Kol | Nazwa | Wartość |
|-----|-------|---------|
| A | Nazwa / Plik | `Allegro Rejda 1.4.2026.pdf` |
| B | Kwota brutto | `-92,91` |
| C | Status | `3` |
| D | Klucz_Ksiegowy | `kos_pr_out` |
| E | wyciag_Kontrahent | `Allegro Poznan 4824 xxxx xxxx 3752` |
| F | wyciag_Kwota | `92,91` |
| G | Data_ksiegowania | `2026-04-03` |
| H | wyciag_Tytul | `Allegro Poznan Nr karty ...3752 92,91PLN` |
| I | wyciag_Data_op | `2026-04-03` |
| J | wyciag_Rodzaj | `Płatności kartą` |
| K | wyciag_Waluta | `PLN` |
| L | wyciag_Nr_rachunku | `41253000087082010009480002` |
| M | wyciag_Imie_Nazwisko | `Allegro Poznan` |
| N | Uwagi | *(puste)* |
| **O** | **Dok** | **`FZ`** |
| **P** | **Konto_Wn** | **`401-MAT`** |
| **Q** | **Konto_Ma** | **`201-ALLE`** |

---

## Przykład 2 — FAKTURA SPRZEDAŻY (czynsz najemcy, Hasan Garip, Perzyńskiego)

| Kol | Nazwa | Wartość |
|-----|-------|---------|
| A | Nazwa / Plik | `Hasan Garip` |
| B | Kwota brutto | `1 250,00` |
| C | Status | `2` |
| D | Klucz_Ksiegowy | `prz_naj_pr_in` |
| E | wyciag_Kontrahent | `HASAN GARIP UL. JÓZEFA BEMA 5/48/2` |
| F | wyciag_Kwota | `1 250,00` |
| G | Data_ksiegowania | `2026-04-03` |
| H | wyciag_Tytul | `pokoj1 perzynskiego kwiecien` |
| I | wyciag_Data_op | `2026-04-03` |
| J | wyciag_Rodzaj | `Przelewy przychodzące` |
| K | wyciag_Waluta | `PLN` |
| L | wyciag_Nr_rachunku | `56109025900000000147810984` |
| M | wyciag_Imie_Nazwisko | `HASAN GARIP UL. JÓZEFA BEMA 5/48/2` |
| N | Uwagi | *(puste)* |
| **O** | **Dok** | **`FS`** |
| **P** | **Konto_Wn** | **`200-M01-01`** |
| **Q** | **Konto_Ma** | **`700-M01-01`** |

*(M01 = Perzyńskiego 11a/28, -01 = Hasan Garip)*

---

## Przykład 3 — WŁAŚCICIELE (przelew do właścicielki, Agnieszka Borkowska)

| Kol | Nazwa | Wartość |
|-----|-------|---------|
| A | Nazwa / Plik | `Agnieszka Borkowska` |
| B | Kwota brutto | `-2 100,00` |
| C | Status | `2` |
| D | Klucz_Ksiegowy | `wla_pr_out` |
| E | wyciag_Kontrahent | `Agnieszka Borkowska` |
| F | wyciag_Kwota | `-2 100,00` |
| G | Data_ksiegowania | `2026-04-03` |
| H | wyciag_Tytul | `Czynsz za wynajem mieszkania za aktualny miesiąc - Perzyńskiego 11a/28` |
| I | wyciag_Data_op | `2026-04-03` |
| J | wyciag_Rodzaj | `Przelewy wychodzące` |
| K | wyciag_Waluta | `PLN` |
| L | wyciag_Nr_rachunku | `24102010970000780200488726` |
| M | wyciag_Imie_Nazwisko | `Agnieszka Borkowska` |
| N | Uwagi | *(puste)* |
| **O** | **Dok** | **`WB`** |
| **P** | **Konto_Wn** | **`250-WL01`** |
| **Q** | **Konto_Ma** | **`131`** |

*(250-WL01 = Agnieszka Borkowska właścicielka M01, 131 = rachunek bankowy mBank)*

---

## Przykład 4 — INNE RK (wypłata z bankomatu BLIK, Euronet)

| Kol | Nazwa | Wartość |
|-----|-------|---------|
| A | Nazwa / Plik | `Bankomat Euronet` |
| B | Kwota brutto | `-200,00` |
| C | Status | `2` |
| D | Klucz_Ksiegowy | `roz_bankomat_rk_kp` |
| E | wyciag_Kontrahent | `Bankomat Euronet` |
| F | wyciag_Kwota | `-200,00` |
| G | Data_ksiegowania | `2026-04-20` |
| H | wyciag_Tytul | `Wypłata BLIK z bankomatu Bankomat Euronet UL RYNEK 13 14 RZESZOW` |
| I | wyciag_Data_op | `2026-04-20` |
| J | wyciag_Rodzaj | `Płatności Blik` |
| K | wyciag_Waluta | `PLN` |
| L | wyciag_Nr_rachunku | `41253000087082010009480002` |
| M | wyciag_Imie_Nazwisko | `Bankomat Euronet` |
| N | Uwagi | *(puste)* |
| **O** | **Dok** | **`WB`** |
| **P** | **Konto_Wn** | **`100`** |
| **Q** | **Konto_Ma** | **`131`** |

*(100 = kasa — gotówka wchodzi do kasy, 131 = bank — pieniądze wychodzą z rachunku)*

---

## Przykład 5 — KAUCJA (wpłata depozytu od najemcy, Lopez Villacis)

| Kol | Nazwa | Wartość |
|-----|-------|---------|
| A | Nazwa / Plik | *(puste — brak faktury)* |
| B | Kwota brutto | *(puste)* |
| C | Status | `3` |
| D | Klucz_Ksiegowy | `roz_depo_part_pr_in` |
| E | wyciag_Kontrahent | `LOPEZ VILLACIS JONATHAN XAVIER` |
| F | wyciag_Kwota | `500,00` |
| G | Data_ksiegowania | `2026-04-13` |
| H | wyciag_Tytul | `REF: 0114696010005667 deposit rent /ES0400490114...` |
| I | wyciag_Data_op | `2026-04-13` |
| J | wyciag_Rodzaj | `Przelewy przychodzące` |
| K | wyciag_Waluta | `PLN` |
| L | wyciag_Nr_rachunku | `86109000040000001178610035` |
| M | wyciag_Imie_Nazwisko | `LOPEZ VILLACIS JONATHAN XAVIER` |
| N | Uwagi | *(puste)* |
| **O** | **Dok** | **`WB`** |
| **P** | **Konto_Wn** | **`131`** |
| **Q** | **Konto_Ma** | **`240-M04-03`** |

*(131 = bank, 240 = kaucja do zwrotu, M04-03 = Lopez Villacis, Balkonowa 1 pokój 3)*

---

---

## Słownik kontrahentów — propozycja zakładki "Kontrahenci"

> Zakładka w głównym arkuszu Google Sheets (SPREADSHEET_ID).
> Program czyta ją przy każdym parowaniu.
> Użytkownik edytuje bez potrzeby deployu.

---

### A. Mieszkania (kody obiektów)

| Kod | Adres | Właściciel | Spółdzielnia |
|-----|-------|-----------|-------------|
| M01 | Perzyńskiego 11a/28 | WL01 | WL02 |
| M02 | Omulewska 18 | WL03 | WL04 |
| M03 | Nałęczowska 62/66 | WL05 | WL06 |
| M04 | Balkonowa 1 | WL08 | WL09 |
| M05 | Górczewska 210/58 | WL12 | WL12-A |
| M06 | Etiudy Rewolucyjnej 44 | WL10 | WL11 |
| M07 | Ryżowa 43b/5 | WL14 | WL14-A |
| M08 | Dembowskiego | WL13 | — |
| M09 | Nowowiejska 15/44 | WL15 | WL15-A |
| M10 | Krysta 5 | WL16 | WL16-A |
| M11 | Bazylińska 7/41 | WL17 | — |
| M12 | Afrykańska 16E | WL18 | — |
| M13 | Umińskiego 20/28 | WL19 | — |
| M14 | Małej Łanki 23 | WL20 | — |
| M15 | Kondratowicza 25A | WL21 | — |
| M16 | Ringelbluma | WL07 | — |

---

### B. Najemcy (konta 200-Mxx-xx)

| Kod | Imię i nazwisko | Mieszkanie | Pokój | Konto należności | Konto przychodu |
|-----|----------------|-----------|-------|-----------------|----------------|
| M01-01 | Hasan Garip | M01 Perzyńskiego | 1 | 200-M01-01 | 700-M01-01 |
| M01-02 | Hromovenko Natalia | M01 Perzyńskiego | 2 | 200-M01-02 | 700-M01-02 |
| M01-03 | Zuzanna Sarnowska | M01 Perzyńskiego | 3 | 200-M01-03 | 700-M01-03 |
| M01-04 | Harikrishnan Madhusoodhanan Nair | M01 Perzyńskiego | 4 | 200-M01-04 | 700-M01-04 |
| M01-05 | Illia Shcherbyna | M01 Perzyńskiego | 5 | 200-M01-05 | 700-M01-05 |
| M02-01 | Stepan Seremchuk | M02 Omulewska | 1 | 200-M02-01 | 700-M02-01 |
| M02-02 | Ejikam Michael Chidiebere | M02 Omulewska | 2 | 200-M02-02 | 700-M02-02 |
| M02-03 | Łukasz Chinek | M02 Omulewska | 9 | 200-M02-03 | 700-M02-03 |
| M02-04 | Mateusz Gruszczyński | M02 Omulewska | 6 | 200-M02-04 | 700-M02-04 |
| M02-05 | Nadiia Sanchuk | M02 Omulewska | — | 200-M02-05 | 700-M02-05 |
| M02-06 | Stanisław Turenko | M02 Omulewska | — | 200-M02-06 | 700-M02-06 |
| M02-07 | Huseyn Farajzade | M02 Omulewska | — | 200-M02-07 | 700-M02-07 |
| M03-01 | Asadbek Yadgorov / Dilshodbek Zokirov | M03 Nałęczowska | — | 200-M03-01 | 700-M03-01 |
| M03-02 | Maksym Ruban / Khrushch Iryna | M03 Nałęczowska | — | 200-M03-02 | 700-M03-02 |
| M03-03 | Vasyl Vitiuk | M03 Nałęczowska | — | 200-M03-03 | 700-M03-03 |
| M03-04 | Maryam Salayeva | M03 Nałęczowska | — | 200-M03-04 | 700-M03-04 |
| M04-01 | Fuzi Yang | M04 Balkonowa | — | 200-M04-01 | 700-M04-01 |
| M04-02 | Rizvan Mamedov | M04 Balkonowa | 3 | 200-M04-02 | 700-M04-02 |
| M04-03 | Nika Kokauri | M04 Balkonowa | — | 200-M04-03 | 700-M04-03 |
| M04-04 | Sofia Antoniv | M04 Balkonowa | — | 200-M04-04 | 700-M04-04 |
| M04-05 | Gagandeep Kaur | M04 Balkonowa | — | 200-M04-05 | 700-M04-05 |
| M05-01 | Avetis Kyubelyan | M05 Górczewska | 3 | 200-M05-01 | 700-M05-01 |
| M05-02 | Kostenko Volodymyr | M05 Górczewska | 2 | 200-M05-02 | 700-M05-02 |
| M06-01 | Sairam Valisetty | M06 Etiudy | — | 200-M06-01 | 700-M06-01 |
| M06-02 | Sasidhar Mahesh Avvaru | M06 Etiudy | 4 | 200-M06-02 | 700-M06-02 |
| M06-03 | Katsiaryna Krasautsava | M06 Etiudy | 3 | 200-M06-03 | 700-M06-03 |
| M06-04 | Oleh Vovchenko | M06 Etiudy | — | 200-M06-04 | 700-M06-04 |
| M07-01 | Akalezi Cajethan | M07 Ryżowa | — | 200-M07-01 | 700-M07-01 |
| M07-02 | Mateusz Nizio | M07 Ryżowa | — | 200-M07-02 | 700-M07-02 |
| M07-03 | Pashkevych Rehina | M07 Ryżowa | — | 200-M07-03 | 700-M07-03 |
| M08-01 | Matsvei Kazhadub | M08 Dembowskiego | — | 200-M08-01 | 700-M08-01 |
| M09-01 | Magdalena Zych | M09 Nowowiejska | — | 200-M09-01 | 700-M09-01 |
| M09-02 | Mirosław Wróbel | M09 Nowowiejska | — | 200-M09-02 | 700-M09-02 |
| M10-01 | Siijn Na | M10 Krysta | — | 200-M10-01 | 700-M10-01 |
| M11-01 | Onyskina Iryna | M11 Bazylińska | 1 | 200-M11-01 | 700-M11-01 |
| M11-02 | Ivanna Hyryk Pashko | M11 Bazylińska | — | 200-M11-02 | 700-M11-02 |
| M11-03 | Arun Kuttiyani | M11 Bazylińska | — | 200-M11-03 | 700-M11-03 |
| M11-04 | Dmytro Shevchenko Helman | M11 Bazylińska | — | 200-M11-04 | 700-M11-04 |
| M12-01 | Chiadika Obenwa | M12 Afrykańska | 1 | 200-M12-01 | 700-M12-01 |
| M12-02 | Viktoriia / Alona Maksymenko | M12 Afrykańska | 2 | 200-M12-02 | 700-M12-02 |
| M12-03 | Ladouce Irakoze | M12 Afrykańska | 3 | 200-M12-03 | 700-M12-03 |
| M13-01 | Oksana Baranovska | M13 Umińskiego | 1 | 200-M13-01 | 700-M13-01 |
| M13-02 | Władysław Kasteczka | M13 Umińskiego | 2 | 200-M13-02 | 700-M13-02 |
| M13-03 | Maksym Khorzhevskyi | M13 Umińskiego | 3 | 200-M13-03 | 700-M13-03 |
| M14-01 | Rumbidzo Nehanda | M14 Małej Łanki | 1 | 200-M14-01 | 700-M14-01 |
| M14-02 | Namatirai Gangata Glenda | M14 Małej Łanki | 2 | 200-M14-02 | 700-M14-02 |
| M14-03 | Catherine Malaki Yohana | M14 Małej Łanki | 3 | 200-M14-03 | 700-M14-03 |
| M15-01 | Artom Dionisiadis | M15 Kondratowicza | — | 200-M15-01 | 700-M15-01 |
| M15-02 | Vitalij Hyz | M15 Kondratowicza | — | 200-M15-02 | 700-M15-02 |
| M15-03 | Saru Magar Biraj | M15 Kondratowicza | 4 | 200-M15-03 | 700-M15-03 |
| M15-04 | Saira Fatima | M15 Kondratowicza | — | 200-M15-04 | 700-M15-04 |
| M15-05 | Mehdi Edbouche | M15 Kondratowicza | — | 200-M15-05 | 700-M15-05 |
| M15-06 | Nazarii Ohonovskyi | M15 Kondratowicza | — | 200-M15-06 | 700-M15-06 |
| M16-01 | Tripathi Palash | M16 Ringelbluma | — | 200-M16-01 | 700-M16-01 |
| M16-02 | Joseph Sagayara | M16 Ringelbluma | — | 200-M16-02 | 700-M16-02 |
| M16-03 | Malhan Aarnav | M16 Ringelbluma | — | 200-M16-03 | 700-M16-03 |
| M16-04 | Ivan Jarashuk | M16 Ringelbluma | — | 200-M16-04 | 700-M16-04 |

---

### C. Właściciele i spółdzielnie (konta 250-WLxx)

| Kod | Nazwa | Adres | Konto |
|-----|-------|-------|-------|
| WL01 | Agnieszka Borkowska | Perzyńskiego 11a/28 | 250-WL01 |
| WL02 | Domhut sp. z o.o. | spółdzielnia Perzyńskiego | 250-WL02 |
| WL03 | Elżbieta Biller | Omulewska 18 | 250-WL03 |
| WL04 | S.B.M. Grenadierów | spółdzielnia Omulewska 18/9 | 250-WL04 |
| WL05 | Ewa Derenowska | Nałęczowska 62/66 | 250-WL05 |
| WL06 | Spółdzielnia Sztuka Nałęczowska | spółdzielnia Nałęczowska 62/66 | 250-WL06 |
| WL07 | Mariusz Myszkiewicz | Ringelbluma | 250-WL07 |
| WL08 | Jolanta Kowalczyk | Balkonowa 1 | 250-WL08 |
| WL09 | Jolanta Kowalczyk (spółdzielnia) | spółdzielnia Balkonowa 1 | 250-WL09 |
| WL10 | Sławomir Stefański | Etiudy Rewolucyjnej 44 | 250-WL10 |
| WL11 | SBM Politechnika | spółdzielnia Etiudy | 250-WL11 |
| WL12 | Danuta Kędzior | Górczewska 210/58 | 250-WL12 |
| WL13 | Jan Laszuk | Dembowskiego | 250-WL13 |
| WL14 | Maciej Warowny | Ryżowa 43b/5 | 250-WL14 |
| WL15 | Anna Zawrzykraj | Nowowiejska 15/44 | 250-WL15 |
| WL16 | Seweryn Brzozowski | Krysta 5 | 250-WL16 |

---

### D. Dostawcy / sklepy (konta 201-XXXX)

| Kod | Nazwa kontrahenta | Konto | Typ kosztów |
|-----|-------------------|-------|-------------|
| 201-ALLE | Allegro (przez Nest Bank / karty) | 201-ALLE | 401-MAT materiały |
| 201-AGAT | Agata S.A. | 201-AGAT | 401-MAT materiały |
| 201-CAST | Castorama | 201-CAST | 401-MAT materiały |
| 201-LERO | Leroy Merlin | 201-LERO | 401-MAT materiały |
| 201-NETI | Netia S.A. | 201-NETI | 401-MED-INT internet |
| 201-EONS | E.ON Polska S.A. | 201-EONS | 401-MED-EL prąd |
| 201-PGNIG | PGNiG Obrót Detaliczny | 201-PGNIG | 401-MED-GAZ gaz |
| 201-PLAY | P4 Sp. z o.o. (Play) | 201-PLAY | 401-MED-INT telefon |
| 201-MAFI | MAFIKA Accounting | 201-MAFI | 402-KSIEG księgowość |
| 201-TBBS | TBBS Polska (dezynfekcja) | 201-TBBS | 402-REM usługi |
| 201-PAYU | PayU (OLX) | 201-PAYU | 402-ADM administracja |
| 201-ELKAB | Elkabel | 201-ELKAB | 401-MAT materiały |
| 201-MULTI | Multiserwis Bielany | 201-MULTI | 402-REM usługi |
| 201-RAYPAT | Raypath (Katarzyna Gliwka) | 201-RAYPAT | 402-REM usługi |

---

### E. Instytucje publiczne

| Kod | Nazwa | Konto zobowiązań | Konto kosztu |
|-----|-------|-----------------|-------------|
| 229-ZUS | Zakład Ubezpieczeń Społecznych | 229-ZUS | 405 |
| 225-US | Urząd Skarbowy / Karbowy Łódź | 225-CIT / 225-PON | 403-CIT / 403-PON |

---

### F. Pracownicy / zleceniobiorcy (konta 234-XXX)

| Kod | Imię i nazwisko | Konto zobowiązań | Konto kosztu |
|-----|----------------|-----------------|-------------|
| 234-ZAP | Kajetan Zapała | 234-ZAP | 404-ZAP |
| 234-DYB | Milena Dybalska-Stypułko | 234-DYB | 404-DYB |

---

### G. Konta własne (bez kontrahenta)

| Konto | Opis |
|-------|------|
| 100 | Kasa (gotówka fizyczna) |
| 131 | Rachunek bankowy mBank |
| 240-Mxx-xx | Kaucja najemcy do zwrotu (depozyt) |
