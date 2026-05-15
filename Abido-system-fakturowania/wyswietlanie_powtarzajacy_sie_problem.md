# Powtarzający się problem z wyświetlaniem

## Objaw

Wewnątrz `st.expander` pojawia się ogromna pusta przestrzeń ponad tabelą wyników.
Tabela renderuje się na dole expandra, ale większość obszaru jest pusta.

## Przyczyna

`st.dataframe` z obiektem `pandas Styler` (`.style.apply(...)`) wewnątrz `st.expander`
rezerwuje domyślnie dużo więcej wysokości niż potrzeba — Streamlit nie wie ile
wierszy będzie, więc przydziela maksimum.

## Rozwiązanie (stałe)

Zawsze dodawaj `height=` wyliczone dynamicznie przy każdym `st.dataframe` ze Stylerem:

```python
_df_height = min((len(df) + 1) * 35 + 3, 600)
st.dataframe(df.style.apply(...), use_container_width=True, hide_index=True, height=_df_height)
```

Wzór: `(liczba_wierszy + 1) * 35 + 3` — +1 na nagłówek, +3 px margines, max 600.

## Miejsca w kodzie gdzie to wystąpiło

| Komponent | Linia (orientacyjna) | Status |
|-----------|----------------------|--------|
| Szukanie Google Sheets — wyniki (`_sh_df_sum.style`) | ~3614 | naprawione 15.05.2026 |

## Zasada na przyszłość

Przy dodawaniu nowego `st.dataframe` wewnątrz `st.expander`:
- Jeśli używasz `.style.apply(...)` — **zawsze** dodaj `height=`
- Jeśli nie używasz Styler — domyślna wysokość zwykle jest OK

## Inne powtarzające się problemy z layoutem

### st.columns wewnątrz st.expander

`st.columns()` na poziomie root po `st.expander` z wewnętrznymi kolumnami
powoduje wizualne nakładanie/duplikowanie elementów przy rerunie.

Rozwiązanie: wyniki renderować przez `@st.dialog` (modal overlay oddzielony od layoutu strony).
Zastosowane w: `_dialog_kosztowe_status`, `_dialog_sprzedaz_status`.

### Confirmation dialog — miganie

Po kliknięciu "Tak" komunikat potwierdzający nie znikał natychmiast.
Rozwiązanie: `st.rerun()` bezpośrednio po `st.session_state` zmianie,
przed właściwą akcją.
