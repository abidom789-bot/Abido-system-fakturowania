# Plan: Logowanie z rolami — dostęp dla księgowej

## Decyzje

| Pytanie | Odpowiedź |
|---------|-----------|
| Biblioteka | `streamlit-authenticator` (Wariant B) |
| Login księgowej | `ksiegowa` |
| Widoczne sekcje | Szukanie Google Drive, Szukanie Google Sheets, Bilans najemcy |
| Zapamiętanie sesji | Tak (cookie) |

---

## Sekcje widoczne dla każdej roli

| Sekcja | Admin | Ksiegowa |
|--------|-------|----------|
| Input subfolder_name | TAK | NIE |
| Przyciski akcji (left_col / right_col) | TAK | NIE |
| Szukanie Google Drive | TAK | TAK |
| Szukanie Google Sheets | TAK | TAK |
| Bilans najemcy | TAK | TAK |
| Szukanie najemcy (osobna sekcja pod bilansem) | TAK | TAK |
| KP i KW | TAK | NIE |
| Dodaj podsumowanie | TAK | NIE |

---

## Kroki implementacji

### KROK 1 — Dodaj zależność

W `requirements.txt` dodaj:
```
streamlit-authenticator==0.3.3
```
(sprawdź aktualną wersję na PyPI przed dodaniem)

### KROK 2 — Wygeneruj hasła (lokalnie, jednorazowo)

Uruchom lokalnie w terminalu:
```python
import streamlit_authenticator as stauth
print(stauth.Hasher(['HASLO_ADMINA', 'HASLO_KSIEGOWEJ']).generate())
```
Wynik: dwa hasze bcrypt — skopiuj je do secrets.

### KROK 3 — Dodaj do Streamlit Secrets

W share.streamlit.io → Settings → Secrets dodaj blok:
```toml
[auth]
cookie_name = "abido_auth"
cookie_key  = "LOSOWY_STRING_32_ZNAKI"
cookie_expiry_days = 30

[auth.credentials.usernames.admin]
name     = "Admin"
password = "$2b$12$HASH_ADMINA..."

[auth.credentials.usernames.ksiegowa]
name     = "Ksiegowa"
password = "$2b$12$HASH_KSIEGOWEJ..."
```

### KROK 4 — Inicjalizacja authenticatora w app.py (na górze, przed UI)

```python
import streamlit_authenticator as stauth
import yaml

# Wczytaj credentials z secrets
_auth_cfg = {
    "credentials": {
        "usernames": {
            "admin":    {"name": "Admin",    "password": st.secrets["auth"]["credentials"]["usernames"]["admin"]["password"]},
            "ksiegowa": {"name": "Ksiegowa", "password": st.secrets["auth"]["credentials"]["usernames"]["ksiegowa"]["password"]},
        }
    },
    "cookie": {
        "name":         st.secrets["auth"]["cookie_name"],
        "key":          st.secrets["auth"]["cookie_key"],
        "expiry_days":  st.secrets["auth"]["cookie_expiry_days"],
    },
}

authenticator = stauth.Authenticate(
    _auth_cfg["credentials"],
    _auth_cfg["cookie"]["name"],
    _auth_cfg["cookie"]["key"],
    _auth_cfg["cookie"]["expiry_days"],
)

name, authentication_status, username = authenticator.login("Logowanie", "main")

if authentication_status is False:
    st.error("Nieprawidłowy login lub hasło.")
    st.stop()
if authentication_status is None:
    st.info("Wprowadź login i hasło.")
    st.stop()

# Rola na podstawie loginu
_role = "admin" if username == "admin" else "ksiegowa"
```

### KROK 5 — Przycisk wylogowania w nagłówku

```python
authenticator.logout("Wyloguj", "sidebar")
# lub w nagłówku strony jako st.button w kolumnie
```

### KROK 6 — Warunkowe renderowanie UI

```python
# Input subfolder + przyciski akcji — tylko admin
if _role == "admin":
    subfolder_name = st.text_input(...)
    left_col, right_col = st.columns(2)
    with left_col:
        btn_czytaj = st.button(...)
    # ... itd.

# Szukanie Google Drive — wszyscy
with st.expander("Szukanie Google Drive", ...):
    ...

# Szukanie Google Sheets — wszyscy
with st.expander("Szukanie Google Sheets", ...):
    ...

# Bilans najemcy — wszyscy
with st.expander("Bilans najemcy", ...):
    ...

# Reszta — tylko admin
if _role == "admin":
    with st.expander("KP i KW", ...):
        ...
    # ... dodaj podsumowanie itd.
```

---

## Kolejność prac w kodzie

1. `requirements.txt` — dodaj `streamlit-authenticator`
2. Wygeneruj hasze lokalnie (skrypt jednorazowy)
3. Wklej hasze + cookie config do Streamlit Secrets
4. Na początku app.py (po importach, przed jakimkolwiek `st.*`) — blok authenticatora
5. Owinąć istniejące sekcje w `if _role == "admin":` lub zostawić bez warunku
6. Test lokalny: `streamlit run app.py` (lokalnie secrets z `.streamlit/secrets.toml`)
7. Push → deploy

---

## Uwagi

- `cookie_key` musi być losowy string min. 32 znaki — wygeneruj np. przez `secrets.token_hex(32)` w Pythonie
- Lokalnie potrzebny plik `.streamlit/secrets.toml` z tym samym blokiem `[auth]`
- Biblioteka `streamlit-authenticator 0.3.x` zmieniała API między wersjami — sprawdź docs dla konkretnej wersji przed kodowaniem
- Jeśli zmieniasz hasło w przyszłości — wystarczy zaktualizować hash w Secrets (bez redeploy kodu)
