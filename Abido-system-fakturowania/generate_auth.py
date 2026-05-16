"""
Uruchom lokalnie JEDNORAZOWO, zeby wygenerowac hashe hasel do Streamlit Secrets.
Uzycie: py generate_auth.py

Wymaga: pip install streamlit-authenticator==0.3.3
"""
import secrets as _secrets

try:
    import streamlit_authenticator as stauth
except ImportError:
    print("Zainstaluj biblioteke: pip install streamlit-authenticator==0.3.3")
    raise

admin_haslo    = input("Haslo dla 'admin':    ")
ksiegowa_haslo = input("Haslo dla 'ksiegowa': ")

hashed = stauth.Hasher([admin_haslo, ksiegowa_haslo]).generate()
cookie_key = _secrets.token_hex(32)

print("\n" + "=" * 60)
print("Wklej do Streamlit Secrets (share.streamlit.io -> Settings -> Secrets):")
print("=" * 60)
print(f"""
[auth]
cookie_name        = "abido_auth"
cookie_key         = "{cookie_key}"
cookie_expiry_days = 30

[auth.credentials.usernames.admin]
name     = "Admin"
password = "{hashed[0]}"

[auth.credentials.usernames.ksiegowa]
name     = "Ksiegowa"
password = "{hashed[1]}"
""")
print("=" * 60)
print("Lokalne testy: wklej to samo do .streamlit/secrets.toml")
