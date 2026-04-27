import io
import re
import streamlit as st
import gspread
import pdfplumber
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ----------------------------------------------------------------
# KONFIGURACJA
# ----------------------------------------------------------------
FOLDER_ID = "1kwY6tOalKS2jnidABw6uUV23ykMj1iR2"
SPREADSHEET_ID = "1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0"
# ----------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]


def get_credentials():
    service_account_info = dict(st.secrets["gcp_service_account"])
    return Credentials.from_service_account_info(service_account_info, scopes=SCOPES)


def find_subfolder(service, parent_folder_id, subfolder_name):
    query = (
        f"'{parent_folder_id}' in parents "
        f"and name = '{subfolder_name}' "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    folders = results.get("files", [])
    return folders[0] if folders else None


def list_pdfs_from_drive(service, folder_id):
    query = (
        f"'{folder_id}' in parents "
        "and mimeType='application/pdf' "
        "and trashed=false"
    )
    results = service.files().list(
        q=query, fields="files(id, name)", orderBy="name"
    ).execute()
    return results.get("files", [])


def download_pdf(service, file_id):
    request = service.files().get_media(fileId=file_id)
    return request.execute()


def extract_gross_amount(pdf_bytes):
    """
    Wyciaga kwote brutto z tekstu PDF i zwraca jako ujemna wartosc string.
    """
    patterns = [
        r"do\s+zap[lł]aty[^\d]*?([\d\s]+[,.][\d]{2})",
        r"razem\s+brutto[^\d]*?([\d\s]+[,.][\d]{2})",
        r"kwota\s+brutto[^\d]*?([\d\s]+[,.][\d]{2})",
        r"[lł][aą]czna\s+kwota[^\d]*?([\d\s]+[,.][\d]{2})",
        r"suma\s+brutto[^\d]*?([\d\s]+[,.][\d]{2})",
        r"ogó?[lł]em\s+brutto[^\d]*?([\d\s]+[,.][\d]{2})",
        r"total[^\d]*?([\d\s]+[,.][\d]{2})",
    ]
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = ""
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
        text_lower = text.lower()
        for pattern in patterns:
            match = re.search(pattern, text_lower)
            if match:
                amount = match.group(1).strip().replace(" ", "")
                return f"-{amount}"
    except Exception:
        pass
    return ""


def get_or_create_worksheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=5)


def read_existing_rows(worksheet):
    """
    Czyta istniejace dane z arkusza.
    Zwraca slownik: nazwa_pliku -> {"brutto": str, "status": str, "row_index": int}
    Pomija wiersz naglowkowy (wiersz 1).
    """
    all_rows = worksheet.get_all_values()
    existing = {}
    for i, row in enumerate(all_rows[1:], start=2):  # start=2 bo wiersz 1 to naglowek
        if len(row) >= 1 and row[0]:
            name = row[0]
            brutto = row[1] if len(row) > 1 else ""
            status = row[2] if len(row) > 2 else "0"
            existing[name] = {"brutto": brutto, "status": status, "row_index": i}
    return existing


def sync_to_sheets(credentials, spreadsheet_id, drive_files_data, sheet_name):
    """
    Synchronizuje dane z Drive do arkusza z zachowaniem wierszy z C=1.

    Logika:
    - C=1 (zweryfikowane przez uzytkownika): wiersze NIETYKALANE
    - C=0 (do weryfikacji): nadpisane nowym odczytem
    - Nowe pliki: dodane na koniec z C=0
    """
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = get_or_create_worksheet(spreadsheet, sheet_name)

    # Sprawdz czy arkusz ma naglowek
    first_row = worksheet.row_values(1)
    if not first_row or first_row[0] != "Nazwa pliku":
        worksheet.insert_row(["Nazwa pliku", "Kwota brutto", "Status"], index=1)

    existing = read_existing_rows(worksheet)

    # Buduj slownik nowych danych z Drive (nazwa -> brutto)
    drive_data = {f["name"]: f["brutto"] for f in drive_files_data}

    updates = []   # (row_index, brutto) — do nadpisania wierszy z C=0
    new_rows = []  # nowe pliki nieobecne w arkuszu

    for name, brutto in drive_data.items():
        if name in existing:
            if existing[name]["status"] == "1":
                # Zweryfikowany — pomijamy
                pass
            else:
                # C=0 — nadpisz kwote, status pozostaje 0
                updates.append((existing[name]["row_index"], name, brutto))
        else:
            # Nowy plik — dodaj
            new_rows.append([name, brutto, "0"])

    # Zastosuj aktualizacje wierszy C=0
    for row_index, name, brutto in updates:
        worksheet.update(f"A{row_index}:C{row_index}", [[name, brutto, "0"]])

    # Dodaj nowe wiersze
    if new_rows:
        worksheet.append_rows(new_rows)

    return len(updates), len(new_rows)


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
    "Wpisz nazwe podfolderu. Nowe faktury zostana dodane z C=0. "
    "Wiersze z C=1 (zweryfikowane recznie) pozostana niezmienione."
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
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)

            with st.spinner(f"Szukam podfolderu '{name}'..."):
                subfolder = find_subfolder(drive_service, FOLDER_ID, name)

            if subfolder is None:
                st.error(f"Nie znaleziono podfolderu '{name}' w folderze glownym.")
            else:
                with st.spinner("Pobieram liste plikow PDF..."):
                    files = list_pdfs_from_drive(drive_service, subfolder["id"])

                if not files:
                    st.warning(f"Brak plikow PDF w podfolderze '{name}'.")
                else:
                    progress = st.progress(0, text="Analizuje faktury...")
                    files_data = []
                    for i, f in enumerate(files):
                        progress.progress(
                            (i + 1) / len(files),
                            text=f"Analizuje: {f['name']}"
                        )
                        pdf_bytes = download_pdf(drive_service, f["id"])
                        brutto = extract_gross_amount(pdf_bytes)
                        files_data.append({"name": f["name"], "brutto": brutto})

                    progress.empty()

                    with st.spinner("Synchronizuje z Google Sheets..."):
                        updated, added = sync_to_sheets(
                            creds, SPREADSHEET_ID, files_data, name
                        )

                    st.success(
                        f"Gotowe! Zaktualizowano: {updated} | Dodano nowych: {added} | "
                        f"Pominieto (C=1): {len(files) - updated - added}"
                    )
                    st.dataframe(
                        [{"Nazwa pliku": d["name"], "Kwota brutto": d["brutto"]} for d in files_data],
                        use_container_width=True,
                    )

        except Exception as e:
            st.error(f"Wystapil blad: {e}")
