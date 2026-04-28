# CLAUDE.md — Abido System Fakturowania

## Ogolne zasady pracy z tym kodem

- Zmieniaj tylko to, o co prosisz. Nie poprawiaj sasiadujacego kodu.
- Nie dodawaj ficzorow, konfigurowalnosci ani obslugi bledow, o ktore nie prosiles.
- Dopasuj styl do istniejacego kodu.

---

## Architektura — jeden plik, trzy warstwy

```
app.py
├── KONFIGURACJA        (stale: FOLDER_ID, SPREADSHEET_ID, separatory sekcji)
├── FUNKCJE BACKENDOWE  (logika, Google API — bez Streamlit)
└── INTERFEJS STREAMLIT (UI: przyciski, inputy, wyniki)
```

Kazdy przycisk = klocek LEGO:
1. Funkcja(e) backendowe (czyste Python, bez `st.*`)
2. Deklaracja przycisku w UI (`st.button(...)`)
3. Blok akcji (`if btn_xxx: ...`) na dole pliku

Dzieki temu kazdy przycisk mozna zmieniac niezaleznie.

---

## Wzorzec dodawania nowego przycisku

### 1. Funkcja backendowa
```python
def moja_nowa_logika(worksheet, dane):
    # tylko logika, zero st.*
    ...
    return wynik
```

### 2. Przycisk w UI (w odpowiedniej kolumnie)
```python
with left_col:   # lub right_col
    btn_nowy = st.button("Nazwa przycisku", use_container_width=True)
```

### 3. Blok akcji na dole pliku
```python
# ----------------------------------------------------------------
# AKCJA: Nazwa przycisku
# ----------------------------------------------------------------
if btn_nowy:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            ...
            st.success("Gotowe!")
        except Exception as e:
            st.error(f"Wystapil blad: {e}")
```

---

## Wspoldzielone funkcje — nie modyfikuj bez potrzeby

| Funkcja | Co robi |
|---|---|
| `read_all_sections(worksheet)` | Czyta caly arkusz, zwraca slownik sekcji |
| `rebuild_sheet(worksheet, sections)` | Zapisuje arkusz w poprawnej kolejnosci + koloruje separatory |
| `apply_sync_logic(existing, new_data)` | Laczy stare (C=1 chronione) z nowymi danymi |
| `_match_separator(val)` | Rozpoznaje separator sekcji (takze legacy formaty) |

Jesli dodajesz nowa sekcje do arkusza:
- Dodaj stala `SEP_NOWA = "--- NOWA SEKCJA ---"`
- Dodaj do `SECTION_ORDER` w odpowiednim miejscu
- Dodaj kolor w `SEP_COLORS`

---

## Zasady ochrony danych uzytkownika

- Wiersz z `status=1` jest **nietykalny** — nigdy nie nadpisuj, nawet jesli plik zniknal z Drive
- `worksheet.clear()` NIE czysci formatowania — uzyj `worksheet.format(...)` z bialym tlem przed zapisem
- Porownuj status jako `str(row[2]).strip() == "1"` (nie `row[2] == "1"`) — unika problemow z typem liczbowym

---

## Uklad UI

```
[ input: subfolder_name (wycentrowany) ]

[ left_col: Faktury kosztowe ]  [ right_col: Faktury sprzedazy ]
  - Zaczytaj faktury kosztowe     - Tworz wiersze faktur sprzedazy
  - Sprawdz stan faktur kosztowych

[ wyniki akcji (pelna szerokosc) ]
```

Nowe przyciski: dodaj do odpowiedniej kolumny lub dodaj trzecia kolumne.

---

## Srodowisko

- Streamlit Cloud, konto: `abidom789-bot`
- Secrets: `[gcp_service_account]` (JSON Service Account)
- Push: `git push origin master:main`
- Lokalnie: `C:/repos abidom789/Abido-system-fakturowania/`
