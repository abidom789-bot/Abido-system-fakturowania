import io
import re
import xlrd
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

HEADER_ROW = [
    "Nazwa / Plik", "Kwota brutto", "Status", "Adres",
    "Klucz_Ksiegowy", "wyciag_Kontrahent", "wyciag_Kwota",
    "Kwota_raport_kasowy", "Data_ksiegowania", "wyciag_Tytul",
    "wyciag_Data_op", "wyciag_Rodzaj", "wyciag_Waluta",
    "wyciag_Nr_rachunku", "wyciag_Imie_Nazwisko", "Uwagi",
]
SEP_KOSZTOWE = "--- FAKTURY KOSZTOWE ---"
SEP_WLASC    = "--- FAKTURY WLASCICIELE I SPOLDZIELNIE ---"
SEP_SPRZEDAZ = "--- FAKTURY SPRZEDAZY NAJEMCOM ---"
SEP_NIEZNANE = "--- NIEZNANE / NIESPAROWANE Z WYCIAGU ---"

LISTY_OPERACJI_FOLDER_NAME = "Listy_operacji_abido"

MEDIA_KW = {"netia", "eon", "e.on", "pgnig", "p4", "play",
            "energia", "prad", "gaz", "internet", "woda", "sbm"}
POSREDNI_KW = ("NEST BANK", "REVOLUT")


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


SECTION_ORDER = [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC, SEP_NIEZNANE]


def get_or_create_worksheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=6)


def _match_separator(val):
    """Rozpoznaje separator sekcji - dokladnie lub po slowach kluczowych (legacy)."""
    if val in SECTION_ORDER:
        return val
    # Dopasowanie czesciowe dla starych/polaczonych separatorow
    if "---" not in val:
        return None
    v = val.upper()
    if "NIEZNANE" in v or "NIESPAROWANE" in v:
        return SEP_NIEZNANE
    if "WLASCICIELE" in v or "SPOLDZIELNIE" in v:
        return SEP_WLASC
    if "SPRZEDAZ" in v or "NAJEMCOM" in v:
        return SEP_SPRZEDAZ
    if "KOSZTOWE" in v:
        return SEP_KOSZTOWE
    return None


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
        matched = _match_separator(val)
        if matched:
            current = matched
        elif current:
            # SEP_NIEZNANE moze miec puste kol A (brak faktury) — wlaczamy jesli cokolwiek wypelnione
            if current == SEP_NIEZNANE:
                if any(c for c in row):
                    sections[current].append(row)
            elif val:
                sections[current].append(row)
    return sections


SEP_COLORS = {
    SEP_KOSZTOWE: {"red": 0.90, "green": 0.22, "blue": 0.22},  # czerwony
    SEP_SPRZEDAZ: {"red": 0.18, "green": 0.65, "blue": 0.32},  # zielony
    SEP_WLASC:    {"red": 0.95, "green": 0.55, "blue": 0.10},  # pomaranczowy
    SEP_NIEZNANE: {"red": 0.50, "green": 0.50, "blue": 0.50},  # szary
}


def rebuild_sheet(worksheet, sections):
    """
    Zapisuje caly arkusz w poprawnej kolejnosci sekcji:
    Header → Kosztowe → Sprzedaz → Wlasciciele.
    Pomija puste sekcje. Koloruje wiersze separatorow.
    """
    all_new = [HEADER_ROW]
    sep_row_nums = {}  # sep -> numer wiersza (1-based)
    for sep in SECTION_ORDER:
        if sections[sep]:
            all_new.append([sep, "", "", ""])
            sep_row_nums[sep] = len(all_new)  # aktualny ostatni wiersz
            all_new.extend(sections[sep])
    worksheet.clear()
    # Reset formatowania calego arkusza (clear() nie czysci kolorow)
    worksheet.format("A1:P500", {
        "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
        "textFormat": {"bold": False},
        "horizontalAlignment": "LEFT",
    })
    # Kolumna A (nazwy) — wyśrodkowana
    worksheet.format("A1:A500", {"horizontalAlignment": "CENTER"})
    if all_new:
        worksheet.update("A1", all_new, value_input_option="USER_ENTERED")
    for sep, row_num in sep_row_nums.items():
        worksheet.format(f"A{row_num}:P{row_num}", {
            "backgroundColor": SEP_COLORS[sep],
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
            },
        })


def apply_sync_logic(existing_rows, new_data, has_address=False):
    """
    Laczy istniejace zweryfikowane wiersze (C=1) z nowymi danymi.
    Wiersze z C=1 sa zachowane nawet jesli plik zostal usuniety z Drive.
    Zwraca (nowe_wiersze, skipped, added).
    """
    verified = {
        row[0]: row
        for row in existing_rows
        if len(row) > 2 and str(row[2]).strip() in ("1", "2")
    }
    new_keys = {item["key"] for item in new_data}
    result = []
    for item in new_data:
        key = item["key"]
        if key in verified:
            result.append(verified[key])
        else:
            addr = item.get("address", "") if has_address else ""
            result.append([key, item.get("brutto", ""), "0", addr])
    # Zachowaj zweryfikowane wiersze ktorych plik zostal usuniety z Drive
    for key, row in verified.items():
        if key not in new_keys:
            result.append(row)
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


def count_kosztowe_statuses(credentials, spreadsheet_id, sheet_name):
    """Liczy tylko wiersze z sekcji KOSZTOWE w arkuszu."""
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return None
    rows = read_all_sections(worksheet)[SEP_KOSZTOWE]
    counts = {"0": 0, "1": 0, "inne": 0}
    for row in rows:
        if not row or not row[0]:
            continue
        status = str(row[2]).strip() if len(row) > 2 else ""
        if status == "0":
            counts["0"] += 1
        elif status == "1":
            counts["1"] += 1
        else:
            counts["inne"] += 1
    return counts


def count_parowanie_statuses(credentials, spreadsheet_id, sheet_name):
    """Liczy wiersze wg statusu i sparowania, per sekcja."""
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return None
    sections = read_all_sections(worksheet)
    result = {}
    for sep in [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC]:
        counts = {"s0": 0, "s1_bez": 0, "s1_para": 0, "s2": 0}
        for row in sections[sep]:
            status = str(row[2]).strip() if len(row) > 2 else ""
            has_pair = len(row) > 5 and str(row[5]).strip() != ""
            if status == "0":
                counts["s0"] += 1
            elif status == "1":
                counts["s1_para" if has_pair else "s1_bez"] += 1
            elif status == "2":
                counts["s2"] += 1
        result[sep] = counts
    return result


# ----------------------------------------------------------------
# PAROWANIE WYCIAGU BANKOWEGO
# ----------------------------------------------------------------

def find_bank_file(service, subfolder_name):
    """Szuka pliku lista_operacji_MMRRRR.xls w folderze Listy_operacji_abido."""
    filename = f"lista_operacji_{subfolder_name}.xls"
    folder = find_subfolder(service, FOLDER_ID, LISTY_OPERACJI_FOLDER_NAME)
    if not folder:
        return None
    query = (
        f"'{folder['id']}' in parents "
        f"and name = '{filename}' "
        "and trashed = false"
    )
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])
    return files[0] if files else None


def parse_bank_statement(xls_bytes):
    """Parsuje XLS wyciagu bankowego. Zwraca liste slownikow transakcji."""
    wb = xlrd.open_workbook(file_contents=xls_bytes)
    ws = wb.sheet_by_index(0)
    transactions = []
    for i in range(7, ws.nrows):   # wiersze 0-5 = naglowek banku, 6 = nazwy kolumn
        row = ws.row_values(i)
        if not row[0]:
            continue
        try:
            kwota = float(row[3])
        except (ValueError, TypeError):
            kwota = 0.0
        transactions.append({
            "data_ks":    str(row[0]),
            "data_op":    str(row[1]),
            "rodzaj":     str(row[2]),
            "kwota":      kwota,
            "waluta":     str(row[4]),
            "kontrahent": str(row[5]),
            "nr_rachunku":str(row[6]),
            "tytul":      str(row[7]),
        })
    return transactions


def _parse_amount(s):
    """Parsuje kwote z komorki arkusza (obsluguje przecinek i minus)."""
    try:
        return abs(float(str(s).replace(",", ".")))
    except (ValueError, TypeError):
        return None


def _is_posredni(kontrahent):
    """Zwraca True jesli kontrahent to posrednik (Nest Bank, Revolut)."""
    k = kontrahent.upper()
    return any(p in k for p in POSREDNI_KW)


def _search_token(tx, token):
    """Sprawdza czy token wystepuje w kontrahencie lub tytule transakcji."""
    token_up = token.upper()
    if _is_posredni(tx["kontrahent"]):
        return token_up in tx["tytul"].upper()
    return token_up in tx["kontrahent"].upper() or token_up in tx["tytul"].upper()


def _extract_name_tokens(text):
    """Wyodrebnia tokeny alfabetyczne z nazwy (min 3 znaki, bez cyfr)."""
    text = re.sub(r"\.\w{2,4}$", "", text, flags=re.IGNORECASE)
    tokens = re.findall(r"[a-zA-Z\u00C0-\u024F]{3,}", text)
    stopwords = {"pdf", "xls", "xlsx", "dla", "the", "and", "von", "van",
                 "brak", "cash", "bez", "per", "via"}
    return [t for t in tokens if t.lower() not in stopwords]


def _extract_name_from_tx(tx):
    """Wyodrebnia imie/nazwisko z transakcji bankowej."""
    kontrahent = tx["kontrahent"]
    if _is_posredni(kontrahent):
        tytul = tx["tytul"]
        tytul = re.sub(r"Nr karty.*", "", tytul, flags=re.IGNORECASE)
        tytul = re.sub(r"Sent from Revolut.*", "", tytul, flags=re.IGNORECASE)
        return tytul.strip()[:60]
    return kontrahent.split("|")[0].strip()


def _is_media(tx):
    """Zwraca True jesli transakcja dotyczy mediow."""
    text = (tx["kontrahent"] + " " + tx["tytul"]).lower()
    return any(kw in text for kw in MEDIA_KW)


def assign_klucz_ksiegowy(section, tx, amount_b_str):
    """Wyznacza Klucz_Ksiegowy na podstawie sekcji i transakcji."""
    if tx is None:
        if section == SEP_SPRZEDAZ:
            return "prz_naj_rk_kp"
        try:
            val = float(str(amount_b_str).replace(",", "."))
        except (ValueError, TypeError):
            val = -1
        return "kos_rk_kp" if val > 0 else "kos_rk_kw"

    kwota = tx["kwota"]
    if section == SEP_KOSZTOWE:
        if kwota > 0:
            return "kos_pr_in"
        return "kos_med_pr_out" if _is_media(tx) else "kos_pr_out"
    if section == SEP_SPRZEDAZ:
        return "prz_naj_pr_in" if kwota > 0 else "prz_naj_rk_kp"
    if section == SEP_WLASC:
        return ("wla_med_pr_out" if _is_media(tx) else "wla_pr_out") if kwota < 0 else "wla_pr_in"
    return "nieznany_out" if kwota < 0 else "nieznany_in"


def pair_transactions(candidates, transactions):
    """
    Paruje kandydatow (wiersze arkusza) z transakcjami bankowymi w 4 przebiegach.
    candidates: lista (idx, name, amount_float)
    transactions: lista slownikow transakcji
    Zwraca: matched {cand_idx: tx_idx}, used_tx set(tx_idx)
    """
    matched = {}
    used_tx = set()

    def free_by_amount(amount):
        return [i for i, tx in enumerate(transactions)
                if i not in used_tx and _parse_amount(tx["kwota"]) == amount]

    def assign(cand_idx, tx_idx):
        matched[cand_idx] = tx_idx
        used_tx.add(tx_idx)

    # Przebieg 1: nazwisko (ostatni token) + kwota
    for idx, name, amount in candidates:
        if idx in matched or amount is None:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        last_name = tokens[-1]
        hits = [i for i in free_by_amount(amount) if _search_token(transactions[i], last_name)]
        if hits:
            assign(idx, hits[0])

    # Przebieg 2: imie (pierwszy token) + kwota
    for idx, name, amount in candidates:
        if idx in matched or amount is None:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        first_name = tokens[0]
        hits = [i for i in free_by_amount(amount) if _search_token(transactions[i], first_name)]
        if hits:
            assign(idx, hits[0])

    # Przebieg 3: wszystkie tokeny (slowa kluczowe) + kwota, najlepszy score
    for idx, name, amount in candidates:
        if idx in matched or amount is None:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        pool = free_by_amount(amount)
        if not pool:
            continue
        scored = [(i, sum(1 for t in tokens if _search_token(transactions[i], t))) for i in pool]
        scored = [(i, s) for i, s in scored if s > 0]
        if not scored:
            continue
        max_score = max(s for _, s in scored)
        best = [i for i, s in scored if s == max_score]
        if len(best) == 1:
            assign(idx, best[0])

    # Przebieg 4: sama kwota (ostatnia szansa)
    for idx, name, amount in candidates:
        if idx in matched or amount is None:
            continue
        pool = free_by_amount(amount)
        if len(pool) == 1:
            assign(idx, pool[0])

    return matched, used_tx


def _build_paired_row(existing_row, tx, klucz, uwagi=""):
    """Uzupelnia wiersz arkusza danymi z transakcji bankowej."""
    row = list(existing_row) + [""] * max(0, 16 - len(existing_row))
    row[4]  = klucz
    row[5]  = tx["kontrahent"].split("|")[0]
    row[6]  = tx["kwota"]
    row[7]  = ""   # raport kasowy — puste dla transakcji bankowych
    row[8]  = tx["data_ks"]
    row[9]  = tx["tytul"][:100]
    row[10] = tx["data_op"]
    row[11] = tx["rodzaj"]
    row[12] = tx["waluta"]
    row[13] = tx["nr_rachunku"]
    row[14] = _extract_name_from_tx(tx)
    row[15] = uwagi
    return row


def _build_unmatched_row(tx):
    """Buduje wiersz dla niesparowanej transakcji z wyciagu (A i B puste)."""
    klucz = "nieznany_out" if tx["kwota"] < 0 else "nieznany_in"
    return [
        "", "", "", "",          # A=nazwa, B=kwota, C=status, D=adres
        klucz,
        tx["kontrahent"].split("|")[0],
        tx["kwota"],
        "",                      # raport kasowy
        tx["data_ks"],
        tx["tytul"][:100],
        tx["data_op"],
        tx["rodzaj"],
        tx["waluta"],
        tx["nr_rachunku"],
        _extract_name_from_tx(tx),
        "",                      # uwagi
    ]


def sync_parowanie(worksheet, transactions):
    """
    Paruje wiersze ze statusem 1 z transakcjami bankowymi.
    Status 2 — nietykalny. Status 0 — pomijany.
    Niesparowane transakcje bankowe laduja na dole w sekcji SEP_NIEZNANE.
    Zwraca (sparowane, niesparowane_z_wyciagu).
    """
    sections = read_all_sections(worksheet)

    # Zbierz kandydatow ze statusem 1 ze wszystkich sekcji
    candidates = []   # (flat_idx, section, row_idx_in_section, name, amount)
    for sep in [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC]:
        for i, row in enumerate(sections[sep]):
            if str(row[2]).strip() == "1":
                amount = _parse_amount(row[1] if len(row) > 1 else "")
                candidates.append((len(candidates), sep, i, row[0], amount))

    flat = [(c[0], c[3], c[4]) for c in candidates]
    matched, used_tx = pair_transactions(flat, transactions)

    # Zapisz wyniki parowania do wierszy
    for flat_idx, sep, row_idx, name, amount in candidates:
        row = sections[sep][row_idx]
        tx_idx = matched.get(flat_idx)
        if tx_idx is not None:
            tx = transactions[tx_idx]
            klucz = assign_klucz_ksiegowy(sep, tx, row[1] if len(row) > 1 else "")
            sections[sep][row_idx] = _build_paired_row(row, tx, klucz)
        else:
            # Brak pary — klucz ksiegowy + wyczysc kolumny wyciagu
            klucz = assign_klucz_ksiegowy(sep, None, row[1] if len(row) > 1 else "")
            r = list(row) + [""] * max(0, 16 - len(row))
            r[4] = klucz
            for col in range(5, 16):
                r[col] = ""
            sections[sep][row_idx] = r

    # Niesparowane transakcje z wyciagu → SEP_NIEZNANE (zawsze zastepowane)
    sections[SEP_NIEZNANE] = [
        _build_unmatched_row(transactions[i])
        for i in range(len(transactions))
        if i not in used_tx
    ]

    rebuild_sheet(worksheet, sections)
    return len(matched), len(sections[SEP_NIEZNANE])


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

left_col, right_col = st.columns(2)
with left_col:
    st.markdown("#### Faktury kosztowe")
    btn_czytaj = st.button(
        "Zaczytaj faktury kosztowe",
        use_container_width=True,
        type="primary",
    )
    btn_sprawdz = st.button(
        "Sprawdz stan faktur kosztowych",
        use_container_width=True,
    )
with right_col:
    st.markdown("#### Faktury sprzedazy")
    btn_sprzedaz = st.button(
        "Tworz wstepne wiersze faktur sprzedazy",
        use_container_width=True,
        type="primary",
    )

st.markdown("---")
paruj_col, status_col = st.columns(2)
with paruj_col:
    btn_paruj = st.button(
        "Paruj wyciag bankowy z arkuszem",
        use_container_width=True,
        type="primary",
    )
with status_col:
    btn_status_parowania = st.button(
        "Status parowania",
        use_container_width=True,
    )

# ----------------------------------------------------------------
# AKCJA: Paruj wyciag bankowy
# ----------------------------------------------------------------
if btn_paruj:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed parowaniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)

            with st.spinner(f"Szukam pliku lista_operacji_{name}.xls ..."):
                bank_file = find_bank_file(drive_service, name)

            if bank_file is None:
                st.error(
                    f"Nie znaleziono pliku 'lista_operacji_{name}.xls' "
                    f"w folderze '{LISTY_OPERACJI_FOLDER_NAME}'."
                )
            else:
                with st.spinner("Pobierام plik wyciagu..."):
                    xls_bytes = download_pdf(drive_service, bank_file["id"])

                with st.spinner("Parsuje transakcje..."):
                    transactions = parse_bank_statement(xls_bytes)

                with st.spinner("Paruje z arkuszem..."):
                    client = gspread.authorize(creds)
                    worksheet = get_or_create_worksheet(
                        client.open_by_key(SPREADSHEET_ID), name
                    )
                    sparowane, niesparowane = sync_parowanie(worksheet, transactions)

                st.success(
                    f"Gotowe! Sparowano: {sparowane} pozycji | "
                    f"Niesparowane z wyciagu: {niesparowane}"
                )
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Status parowania
# ----------------------------------------------------------------
if btn_status_parowania:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed sprawdzeniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            with st.spinner("Czytam arkusz..."):
                data = count_parowanie_statuses(creds, SPREADSHEET_ID, name)
            if data is None:
                st.warning(f"Brak arkusza '{name}'.")
            else:
                st.subheader(f"Status parowania — {name}")
                SEKCJE = {
                    SEP_KOSZTOWE: "Kosztowe",
                    SEP_SPRZEDAZ: "Sprzedaz",
                    SEP_WLASC:    "Wlasciciele",
                }
                rows = [
                    {"Wiersz": "Status 0 — niezweryfikowane",         **{v: data[k]["s0"]     for k, v in SEKCJE.items()}},
                    {"Wiersz": "Status 1 — bez pary z wyciagiem",     **{v: data[k]["s1_bez"] for k, v in SEKCJE.items()}},
                    {"Wiersz": "Status 1 — sparowane (czeka na '2')", **{v: data[k]["s1_para"]for k, v in SEKCJE.items()}},
                    {"Wiersz": "Status 2 — zatwierdzone",             **{v: data[k]["s2"]     for k, v in SEKCJE.items()}},
                ]
                for r in rows:
                    r["Razem"] = r["Kosztowe"] + r["Sprzedaz"] + r["Wlasciciele"]
                st.dataframe(rows, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Sprawdz stan faktur kosztowych
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
                sheet_counts = count_kosztowe_statuses(creds, SPREADSHEET_ID, name)

            st.subheader(f"Faktury kosztowe — {name}")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Google Drive**")
                with st.container(border=True):
                    if not subfolder:
                        st.warning("Nie znaleziono folderu na Drive.")
                    else:
                        st.metric("Pliki PDF", drive_count)
            with col_b:
                st.markdown("**Google Sheets (sekcja kosztowa)**")
                with st.container(border=True):
                    if sheet_counts is None:
                        st.warning("Brak arkusza o tej nazwie.")
                    else:
                        st.metric("Wierszy lacznie", sum(sheet_counts.values()))
                        st.markdown(
                            f"- Niezweryfikowane (0): **{sheet_counts['0']}**  \n"
                            f"- Zweryfikowane (1): **{sheet_counts['1']}**  \n"
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
