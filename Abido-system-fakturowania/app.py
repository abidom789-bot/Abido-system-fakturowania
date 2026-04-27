import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ----------------------------------------------------------------
# KONFIGURACJA - uzupelnij te dwie wartosci
# ----------------------------------------------------------------
FOLDER_ID = "1kwY6tOalKS2jnidABw6uUV23ykMj1iR2"
SPREADSHEET_ID = "1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0"
# ----------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


def get_credentials():
    """Pobiera dane logowania z st.secrets (Streamlit Cloud)."""
    service_account_info = dict(st.secrets["gcp_service_account"])
    credentials = Credentials.from_service_account_info(
        service_account_info, scopes=SCOPES
    )
    return credentials


def find_subfolder(credentials, parent_folder_id, subfolder_name):
    """Szuka podfolderu o podanej nazwie wewnatrz folderu nadrzednego."""
    service = build("drive", "v3", credentials=credentials)
    query = (
        f"'{parent_folder_id}' in parents "
        f"and name = '{subfolder_name}' "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    folders = results.get("files", [])
    return folders[0] if folders else None


def list_pdfs_from_drive(credentials, folder_id):
    """Zwraca liste nazw plikow PDF z podanego folderu Google Drive."""
    service = build("drive", "v3", credentials=credentials)
    query = (
        f"'{folder_id}' in parents "
        "and mimeType='application/pdf' "
        "and trashed=false"
    )
    results = (
        service.files()
        .list(
            q=query,
            fields="files(id, name, createdTime, size)",
            orderBy="name",
        )
        .execute()
    )
    return results.get("files", [])


def write_to_sheets(credentials, spreadsheet_id, files):
    """Dopisuje nazwy plikow do Google Sheets (Sheet1)."""
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.sheet1

    # Naglowki jesli arkusz jest pusty
    if worksheet.row_count == 0 or worksheet.cell(1, 1).value is None:
        worksheet.append_row(["Nazwa pliku", "Data utworzenia", "Rozmiar (bytes)"])

    rows = [
        [
            f["name"],
            f.get("createdTime", ""),
            f.get("size", ""),
        ]
        for f in files
    ]
    if rows:
        worksheet.append_rows(rows)

    return len(rows)


# ----------------------------------------------------------------
# INTERFEJS STREAMLIT
# ----------------------------------------------------------------

st.set_page_config(
    page_title="System Fakturowania",
    page_icon=":page_facing_up:",
    layout="centered",
)

st.title("System Fakturowania")
st.markdown("---")

st.markdown(
    """
    Kliknij przycisk ponizej, aby zaczytac liste plikow PDF
    z folderu Google Drive i zapisac je do Arkusza Google Sheets.
    """
)

st.markdown("<br>", unsafe_allow_html=True)

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    subfolder_name = st.text_input(
        "Nazwa podfolderu (np. 032026)",
        placeholder="wpisz nazwe podfolderu...",
    )
    run = st.button(
        "Zaczytaj faktury kosztowe do Google Sheets",
        use_container_width=True,
        type="primary",
    )

if run:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed uruchomieniem.")
    elif FOLDER_ID.startswith("WKLEJ") or SPREADSHEET_ID.startswith("WKLEJ"):
        st.error(
            "Uzupelnij zmienne FOLDER_ID oraz SPREADSHEET_ID w pliku app.py przed uruchomieniem."
        )
    else:
        with st.spinner(f"Szukam podfolderu '{subfolder_name}'..."):
            try:
                creds = get_credentials()
                subfolder = find_subfolder(creds, FOLDER_ID, subfolder_name.strip())

                if subfolder is None:
                    st.error(f"Nie znaleziono podfolderu o nazwie '{subfolder_name}' w folderze glownym.")
                else:
                    st.info(f"Znaleziono podfolder: {subfolder['name']} — czytam pliki PDF...")
                    files = list_pdfs_from_drive(creds, subfolder["id"])

                    if not files:
                        st.warning(f"Brak plikow PDF w podfolderze '{subfolder_name}'.")
                    else:
                        count = write_to_sheets(creds, SPREADSHEET_ID, files)
                        st.success(
                            f"Gotowe! Dopisano {count} plik(ow) PDF z folderu '{subfolder_name}' do Google Sheets."
                        )
                        st.dataframe(
                            [{"Nazwa pliku": f["name"], "Data": f.get("createdTime", ""), "Rozmiar": f.get("size", "")} for f in files],
                            use_container_width=True,
                        )
            except Exception as e:
                st.error(f"Wystapil blad: {e}")
