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
FOLDER_ID            = "1kwY6tOalKS2jnidABw6uUV23ykMj1iR2"           # Faktury (root)
FAKTURY_KOSZTOWE_ID  = "12RxQDakB6y9pxURM_Z73sS0fLNQyGtm1"           # Faktury-kosztowe
MIESZKANIA_FOLDER_ID = "1mvVZN6y2vaKyWGV6SIWd7FuK38T2DHAI"           # 01_MIESZKANIA
SPREADSHEET_ID       = "1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0"
# ----------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

SEP_KOSZTOWE   = "--- FAKTURY KOSZTOWE ---"
SEP_WLASCICIEL = "--- FAKTURY WLASCICIELE I SPOLDZIELNIE ---"
SEP_SPRZEDAZ   = "--- FAKTURY SPRZEDAZY NAJEMCOM ---"


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


def list_fvs_folders(service, parent_id):
    """Zwraca liste folderow najemcow oznaczonych tagiem [FVS]."""
    query = (
        f"'{parent_id}' in parents "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    folders = results.get("files", [])
    return [f for f in folders if f["name"].startswith("[FVS]")]


def parse_fvs_folder(folder_name):
    """
    Parsuje nazwe folderu najemcy.
    Format: [FVS] Imie Nazwisko | Adres | Cena | DataOd-DataDo
    Zwraca dict: name, address, price, period
    """
    # Usun tag [FVS] z poczatku
    text = folder_name.replace("[FVS]", "").strip()
    parts = [p.strip() for p in text.split("|")]
    return {
        "name":    parts[0] if len(parts) > 0 else "",
        "address": parts[1] if len(parts) > 1 else "",
        "price":   parts[2] if len(parts) > 2 else "",
        "period":  parts[3] if len(parts) > 3 else "",
    }


# ----------------------------------------------------------------
# GOOGLE SHEETS — funkcje pomocnicze
# ----------------------------------------------------------------

def get_or_create_worksheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=6)


def is_separator(row):
    """Sprawdza czy wiersz jest separatorem sekcji."""
    return row and row[0].startswith("---")


def read_section_rows(worksheet, section_marker):
    """
    Czyta wiersze nalezace do danej sekcji (miedzy jej separatorem a nastepnym).
    Zwraca dict: nazwa -> {brutto, status, row_index}
    """
    all_rows = worksheet.get_all_values()
    in_section = False
    existing = {}
    for i, row in enumerate(all_rows):
        val = row[0] if row else ""
        if val == section_marker:
            in_section = True
            continue
        if in_section:
            if val.startswith("---"):
                break  # nastepna sekcja
            if not val:
                continue
            brutto  = row[1] if len(row) > 1 else ""
            status  = row[2] if len(row) > 2 else "0"
            address = row[3] if len(row) > 3 else ""
            existing[val] = {
                "brutto": brutto, "status": status,
                "address": address, "row_index": i + 1  # 1-based
            }
    return existing


def ensure_sheet_structure(worksheet):
    """Upewnia sie ze arkusz ma poprawna strukture z separatorami sekcji."""
    all_rows = worksheet.get_all_values()
    flat = [r[0] for r in all_rows if r]

    # Naglowek kolumn
    if not flat or flat[0] != "Nazwa / Plik":
        worksheet.insert_row(
            ["Nazwa / Plik", "Kwota brutto", "Status", "Adres"], index=1
        )
        all_rows = worksheet.get_all_values()
        flat = [r[0] for r in all_rows if r]

    # Dodaj brakujace separatory na koncu jesli ich nie ma
    for sep in [SEP_KOSZTOWE, SEP_WLASCICIEL, SEP_SPRZEDAZ]:
        if sep not in flat:
            worksheet.append_row([sep, "", "", ""])


def sync_section(worksheet, section_marker, drive_data, has_address=False):
    """
    Synchronizuje wiersze w danej sekcji arkusza.
    - Usuwa wiersze C=0 w tej sekcji
    - Zachowuje C=1
    - Dodaje nowe pozycje z Drive ktore nie maja C=1
    Zwraca (skipped, added).
    """
    all_rows = worksheet.get_all_values()

    # Znajdz zakres sekcji (indeksy 1-based)
    sep_row_idx   = None
    next_sep_idx  = None
    for i, row in enumerate(all_rows):
        val = row[0] if row else ""
        if val == section_marker:
            sep_row_idx = i + 1
        elif sep_row_idx and val.startswith("---") and (i + 1) > sep_row_idx:
            next_sep_idx = i + 1
            break

    if sep_row_idx is None:
        return 0, 0

    # Zbierz wiersze sekcji
    end_idx = next_sep_idx - 1 if next_sep_idx else len(all_rows)
    section_rows = []
    for i in range(sep_row_idx, end_idx):  # 0-based slice
        row = all_rows[i] if i < len(all_rows) else []
        key = row[0] if row else ""
        if not key:
            continue
        status = row[2] if len(row) > 2 else "0"
        section_rows.append({"key": key, "status": status, "row_index": i + 1})

    verified_keys = {r["key"] for r in section_rows if r["status"] == "1"}

    # Usun wiersze C=0 od konca
    to_delete = sorted(
        [r["row_index"] for r in section_rows if r["status"] != "1"],
        reverse=True
    )
    for idx in to_delete:
        worksheet.delete_rows(idx)

    # Dodaj nowe pozycje nie bedace wsrod zweryfikowanych
    new_rows = []
    for item in drive_data:
        key = item["key"]
        if key not in verified_keys:
            if has_address:
                new_rows.append([key, item.get("brutto", ""), "0", item.get("address", "")])
            else:
                new_rows.append([key, item.get("brutto", ""), "0", ""])

    if new_rows:
        # Wstaw przed nastepnym separatorem lub na koncu sekcji
        # Najlatwiej: append (separatory sa zawsze na koncu)
        worksheet.append_rows(new_rows)

    return len(verified_keys), len(new_rows)


def count_sheet_statuses(credentials, spreadsheet_id, sheet_name):
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return None
    all_rows = worksheet.get_all_values()
    counts = {"0": 0, "1": 0, "inne": 0}
    for row in all_rows[1:]:
        if not row or not row[0] or row[0].startswith("---"):
            continue
        status = row[2] if len(row) > 2 else ""
        if status == "0":
            counts["0"] += 1
        elif status == "1":
            counts["1"] += 1
        else:
            counts["inne"] += 1
    return counts


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

col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    subfolder_name = st.text_input(
        "Miesiac (np. 032026)",
        placeholder="wpisz nazwe podfolderu miesiacowego...",
    )
    btn_czytaj = st.button(
        "Zaczytaj faktury kosztowe do Google Sheets",
        use_container_width=True,
        type="primary",
    )
    btn_sprzedaz = st.button(
        "Tworz wstepne wiersze faktur sprzedazy",
        use_container_width=True,
        type="primary",
    )
    btn_sprawdz = st.button(
        "Sprawdz ilosc pozycji",
        use_container_width=True,
    )

# ----------------------------------------------------------------
# AKCJA: Sprawdz ilosc pozycji
# ----------------------------------------------------------------
if btn_sprawdz:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed sprawdzeniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)

            with st.spinner("Sprawdzam..."):
                subfolder = find_subfolder(drive_service, FAKTURY_KOSZTOWE_ID, name)
                drive_count = 0
                if subfolder:
                    files = list_pdfs_from_drive(drive_service, subfolder["id"])
                    drive_count = len(files)
                sheet_counts = count_sheet_statuses(creds, SPREADSHEET_ID, name)

            st.subheader(f"Wyniki dla: {name}")
            col_a, col_b = st.columns(2)

            with col_a:
                st.markdown("**Google Drive**")
                with st.container(border=True):
                    if subfolder is None:
                        st.warning("Nie znaleziono folderu na Drive.")
                    else:
                        st.metric("Pliki PDF w folderze", drive_count)

            with col_b:
                st.markdown("**Google Sheets**")
                with st.container(border=True):
                    if sheet_counts is None:
                        st.warning("Brak arkusza o tej nazwie.")
                    else:
                        total = sum(sheet_counts.values())
                        st.metric("Wierszy lacznie", total)
                        st.markdown(
                            f"- Status **0** (do weryfikacji): **{sheet_counts['0']}**  \n"
                            f"- Status **1** (zweryfikowane): **{sheet_counts['1']}**  \n"
                            f"- Brak statusu / inne: **{sheet_counts['inne']}**"
                        )
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Zaczytaj faktury kosztowe
# ----------------------------------------------------------------
if btn_czytaj:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed uruchomieniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)

            with st.spinner(f"Szukam podfolderu '{name}'..."):
                subfolder = find_subfolder(drive_service, FAKTURY_KOSZTOWE_ID, name)

            if subfolder is None:
                st.error(f"Nie znaleziono podfolderu '{name}' w Faktury-kosztowe.")
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
                        files_data.append({"key": f["name"], "brutto": brutto})
                    progress.empty()

                    with st.spinner("Synchronizuje z Google Sheets..."):
                        client = gspread.authorize(creds)
                        spreadsheet = client.open_by_key(SPREADSHEET_ID)
                        worksheet = get_or_create_worksheet(spreadsheet, name)
                        ensure_sheet_structure(worksheet)
                        skipped, added = sync_section(
                            worksheet, SEP_KOSZTOWE, files_data, has_address=False
                        )

                    st.success(
                        f"Gotowe! Dodano/odswiezono: {added} | "
                        f"Zachowano zweryfikowanych (C=1): {skipped}"
                    )
                    st.dataframe(
                        [{"Nazwa pliku": d["key"], "Kwota brutto": d["brutto"]} for d in files_data],
                        use_container_width=True,
                    )
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Tworz wiersze faktur sprzedazy najemcom
# ----------------------------------------------------------------
if btn_sprzedaz:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed uruchomieniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)

            with st.spinner("Szukam folderow najemcow [FVS]..."):
                fvs_folders = list_fvs_folders(drive_service, MIESZKANIA_FOLDER_ID)

            if not fvs_folders:
                st.warning("Nie znaleziono zadnych folderow z tagiem [FVS] w 01_MIESZKANIA.")
            else:
                tenants = [parse_fvs_folder(f["name"]) for f in fvs_folders]
                drive_data = [
                    {"key": t["name"], "brutto": t["price"], "address": t["address"]}
                    for t in tenants
                ]

                with st.spinner("Synchronizuje z Google Sheets..."):
                    client = gspread.authorize(creds)
                    spreadsheet = client.open_by_key(SPREADSHEET_ID)
                    worksheet = get_or_create_worksheet(spreadsheet, name)
                    ensure_sheet_structure(worksheet)
                    skipped, added = sync_section(
                        worksheet, SEP_SPRZEDAZ, drive_data, has_address=True
                    )

                st.success(
                    f"Gotowe! Dodano/odswiezono: {added} najemcow | "
                    f"Zachowano zweryfikowanych (C=1): {skipped}"
                )
                st.dataframe(
                    [{"Najemca": t["name"], "Kwota": t["price"],
                      "Adres": t["address"], "Okres": t["period"]} for t in tenants],
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"Wystapil blad: {e}")
