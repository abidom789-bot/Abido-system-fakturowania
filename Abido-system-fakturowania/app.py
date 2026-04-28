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
FOLDER_ID            = "1kwY6tOalKS2jnidABw6uUV23ykMj1iR2"
FAKTURY_KOSZTOWE_ID  = "12RxQDakB6y9pxURM_Z73sS0fLNQyGtm1"
MIESZKANIA_FOLDER_ID = "1mvVZN6y2vaKyWGV6SIWd7FuK38T2DHAI"
SPREADSHEET_ID       = "1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0"
# ----------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

HEADER_ROW   = ["Nazwa / Plik", "Kwota brutto", "Status", "Adres"]
SEP_KOSZTOWE = "--- FAKTURY KOSZTOWE ---"
SEP_WLASC    = "--- FAKTURY WLASCICIELE I SPOLDZIELNIE ---"
SEP_SPRZEDAZ = "--- FAKTURY SPRZEDAZY NAJEMCOM ---"


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
    return service.files().get_media(fileId=file_id).execute()


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
            text = "".join(page.extract_text() or "" for page in pdf.pages)
        for pattern in patterns:
            match = re.search(pattern, text.lower())
            if match:
                return "-" + match.group(1).strip().replace(" ", "")
    except Exception:
        pass
    return ""


def list_fvs_folders(service):
    """Szuka folderow [FVS] rekurencyjnie na calym dysku bota."""
    query = (
        "name contains '[FVS]' "
        "and mimeType = 'application/vnd.google-apps.folder' "
        "and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    return results.get("files", [])


def parse_fvs_folder(folder_name):
    """
    Format: [FVS] Imie Nazwisko | Adres | Cena | DataOd-DataDo
    Zwraca: name, address, price
    """
    text = folder_name.replace("[FVS]", "").strip()
    parts = [p.strip() for p in text.split("|")]
    return {
        "name":    parts[0] if len(parts) > 0 else "",
        "address": parts[1] if len(parts) > 1 else "",
        "price":   parts[2] if len(parts) > 2 else "",
    }


SECTION_ORDER = [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC]


def get_or_create_worksheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=6)


def read_all_sections(worksheet):
    """
    Czyta caly arkusz i zwraca slownik sekcji:
    { SEP_KOSZTOWE: [rows...], SEP_SPRZEDAZ: [rows...], ... }
    Wiersze separatorow sa pomijane.
    """
    all_rows = worksheet.get_all_values()
    sections = {sep: [] for sep in SECTION_ORDER}
    current  = None
    for row in all_rows:
        val = row[0] if row else ""
        if val in SECTION_ORDER:
            current = val
        elif current and val:
            sections[current].append(row)
    return sections


def rebuild_sheet(worksheet, sections):
    """
    Zapisuje caly arkusz w poprawnej kolejnosci sekcji:
    Header → Kosztowe → Wlasciciele → Sprzedaz.
    Pomija puste sekcje.
    """
    all_new = [HEADER_ROW]
    for sep in SECTION_ORDER:
        if sections[sep]:
            all_new.append([sep, "", "", ""])
            all_new.extend(sections[sep])
    worksheet.clear()
    if all_new:
        worksheet.update("A1", all_new)


def apply_sync_logic(existing_rows, new_data, has_address=False):
    """
    Laczy istniejace zweryfikowane wiersze (C=1) z nowymi danymi.
    Zwraca (nowe_wiersze, skipped, added).
    """
    verified = {
        row[0]: row
        for row in existing_rows
        if len(row) > 2 and row[2] == "1"
    }
    result = []
    for item in new_data:
        key = item["key"]
        if key in verified:
            result.append(verified[key])
        else:
            addr = item.get("address", "") if has_address else ""
            result.append([key, item.get("brutto", ""), "0", addr])
    return result, len(verified), len(new_data) - len(verified)


def sync_kosztowe(worksheet, files_data):
    sections                  = read_all_sections(worksheet)
    new_rows, skipped, added  = apply_sync_logic(sections[SEP_KOSZTOWE], files_data)
    sections[SEP_KOSZTOWE]    = new_rows
    rebuild_sheet(worksheet, sections)
    return skipped, added


def sync_sprzedaz(worksheet, tenants_data):
    sections                  = read_all_sections(worksheet)
    new_rows, skipped, added  = apply_sync_logic(
        sections[SEP_SPRZEDAZ], tenants_data, has_address=True
    )
    sections[SEP_SPRZEDAZ]    = new_rows
    rebuild_sheet(worksheet, sections)
    return skipped, added


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
                subfolder    = find_subfolder(drive_service, FAKTURY_KOSZTOWE_ID, name)
                drive_count  = len(list_pdfs_from_drive(drive_service, subfolder["id"])) if subfolder else 0
                sheet_counts = count_sheet_statuses(creds, SPREADSHEET_ID, name)

            st.subheader(f"Wyniki dla: {name}")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Google Drive**")
                with st.container(border=True):
                    if not subfolder:
                        st.warning("Nie znaleziono folderu na Drive.")
                    else:
                        st.metric("Pliki PDF w folderze", drive_count)
            with col_b:
                st.markdown("**Google Sheets**")
                with st.container(border=True):
                    if sheet_counts is None:
                        st.warning("Brak arkusza o tej nazwie.")
                    else:
                        st.metric("Wierszy lacznie", sum(sheet_counts.values()))
                        st.markdown(
                            f"- Status **0**: **{sheet_counts['0']}**  \n"
                            f"- Status **1**: **{sheet_counts['1']}**  \n"
                            f"- Inne: **{sheet_counts['inne']}**"
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
                        progress.progress((i + 1) / len(files), text=f"Analizuje: {f['name']}")
                        brutto = extract_gross_amount(download_pdf(drive_service, f["id"]))
                        files_data.append({"key": f["name"], "brutto": brutto})
                    progress.empty()

                    with st.spinner("Zapisuje do Google Sheets..."):
                        client = gspread.authorize(creds)
                        worksheet = get_or_create_worksheet(
                            client.open_by_key(SPREADSHEET_ID), name
                        )
                        skipped, added = sync_kosztowe(worksheet, files_data)

                    st.success(f"Gotowe! Odswiezono: {added} | Zachowano (C=1): {skipped}")
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
                fvs_folders = list_fvs_folders(drive_service)

            if not fvs_folders:
                st.warning("Nie znaleziono zadnych folderow z tagiem [FVS].")
            else:
                tenants = [parse_fvs_folder(f["name"]) for f in fvs_folders]
                tenants_data = [
                    {"key": t["name"], "brutto": t["price"], "address": t["address"]}
                    for t in tenants
                ]

                with st.spinner("Zapisuje do Google Sheets..."):
                    client = gspread.authorize(creds)
                    worksheet = get_or_create_worksheet(
                        client.open_by_key(SPREADSHEET_ID), name
                    )
                    skipped, added = sync_sprzedaz(worksheet, tenants_data)

                st.success(f"Gotowe! Dodano: {added} najemcow | Zachowano (C=1): {skipped}")
                st.dataframe(
                    [{"Najemca": t["name"], "Kwota": t["price"], "Adres": t["address"]}
                     for t in tenants],
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"Wystapil blad: {e}")
