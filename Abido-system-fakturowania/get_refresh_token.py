"""
Skrypt do jednorazowego uzyskania refresh_token dla Google Drive.
Uruchom lokalnie, NIE na Streamlit Cloud.

INSTRUKCJA:
1. Wejdz na https://console.cloud.google.com/
2. Projekt: regal-river-494622-c7
3. APIs & Services -> Credentials -> Create Credentials -> OAuth 2.0 Client ID
4. Application type: Desktop app, Nazwa: np. "Abido Drive Upload"
5. Pobierz JSON z credentials (przycisk Download)
6. Skopiuj client_id i client_secret z pobranego pliku
7. Wpisz je ponizej i uruchom: python get_refresh_token.py
8. Zaloguj sie w przegladarce i zgódz sie na dostep do Drive
9. Skopiuj wynik do Streamlit secrets

pip install google-auth-oauthlib
"""

CLIENT_ID     = "WPISZ_SWOJ_CLIENT_ID"
CLIENT_SECRET = "WPISZ_SWOJ_CLIENT_SECRET"

# ----------------------------------------------------------------

from google_auth_oauthlib.flow import InstalledAppFlow

flow = InstalledAppFlow.from_client_config(
    {
        "installed": {
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
            "token_uri":     "https://oauth2.googleapis.com/token",
            "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob", "http://localhost"],
        }
    },
    scopes=["https://www.googleapis.com/auth/drive"],
)

creds = flow.run_local_server(port=0)

print("\n" + "=" * 60)
print("Skopiuj te wartosci do Streamlit Cloud secrets:")
print("Ustawienia -> Secrets -> Edit (format TOML)")
print("=" * 60)
print()
print("[google_drive_oauth]")
print(f'client_id = "{CLIENT_ID}"')
print(f'client_secret = "{CLIENT_SECRET}"')
print(f'refresh_token = "{creds.refresh_token}"')
print()
print("=" * 60)
