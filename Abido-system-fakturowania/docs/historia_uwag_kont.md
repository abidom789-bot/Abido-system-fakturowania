# Historia uwag i korekt — rozpoznawanie kont księgowych

> Plik NIGDY nie jest kasowany. Każda korekta zapisywana chronologicznie.
> Cel: Claude czyta ten plik na początku sesji i zna powody wszystkich decyzji w kodzie.
> Gdy użytkownik zgłosi błąd → Claude dopisuje wpis tutaj → poprawia kod → commit.

---

## Jak zgłaszać błąd

Powiedz Claude:
> "Program przypisał [co przypisał] dla faktury [nazwa/opis], a powinno być [co powinno być]"

Claude dopisze wpis do tego pliku i poprawi logikę rozpoznawania.

---

## Format wpisu

```
### RRRR-MM-DD — [krótki opis błędu]
**Faktura / transakcja:** [nazwa pliku lub opis]
**Program przypisał:** Dok=X, Konto_Wn=X, Konto_Ma=X
**Powinno być:**        Dok=X, Konto_Wn=X, Konto_Ma=X
**Powód błędu:** [dlaczego program się pomylił]
**Zmiana w kodzie:** [co zostało zmienione, w której funkcji]
```

---

## Wpisy korekt

*(brak wpisów — plik gotowy do zapisu pierwszej korekty)*

---

## Wzorce rozpoznawania — aktualny stan

> Ta sekcja aktualizowana przez Claude po każdej korekcie.
> Opisuje AKTUALNĄ logikę programu — co rozpoznaje i jak.

### Rozpoznawanie symbolu Dok
*(do uzupełnienia po implementacji)*

### Rozpoznawanie Konto_Wn
*(do uzupełnienia po implementacji)*

### Rozpoznawanie Konto_Ma / Kod kontrahenta
*(do uzupełnienia po implementacji)*

### Przypadki niejednoznaczne — wymagają ręcznego sprawdzenia
*(do uzupełnienia po implementacji)*
