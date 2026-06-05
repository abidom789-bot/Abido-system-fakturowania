import io
import os
import re
import zipfile
import time
import calendar
import unicodedata
from collections import Counter
import xlrd
import streamlit as st
import hashlib
import gspread
import pdfplumber
from datetime import date, datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ----------------------------------------------------------------
# HELPERS — retry dla Google Sheets API (limit 60 write req/min)
# ----------------------------------------------------------------
def _api(fn, *args, **kwargs):
    """Wywołuje funkcję gspread z exponential backoff przy błędzie 429."""
    for attempt in range(7):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            if "429" in str(e) and attempt < 6:
                wait = min(2 ** attempt, 64)   # 1, 2, 4, 8, 16, 32, 64 sek
                time.sleep(wait)
            else:
                raise


def _norm_date(s):
    """Normalizuje date do formatu DD-MM-YYYY (obsluguje format GSheets '2026-03-03 00:00:00')."""
    s = str(s or "").strip()
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', s)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})', s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return s


def _read_col_b_notes(worksheet):
    """Czyta notatki z kol B dla wierszy z niepustą kol A.
    Zwraca {col_a_value: note_text}.
    """
    try:
        sid  = worksheet.spreadsheet.id
        name = worksheet.title
        resp = worksheet.spreadsheet.client.request(
            "GET",
            f"https://sheets.googleapis.com/v4/spreadsheets/{sid}",
            params={
                "ranges": f"'{name}'!A:B",
                "includeGridData": "true",
                "fields": "sheets.data.rowData.values.note,sheets.data.rowData.values.userEnteredValue",
            },
        )
        rows_data = (resp.json()
                     .get("sheets", [{}])[0]
                     .get("data",   [{}])[0]
                     .get("rowData", []))
        notes = {}
        for rd in rows_data:
            vals = rd.get("values", [])
            if len(vals) < 2:
                continue
            uev   = vals[0].get("userEnteredValue", {})
            col_a = str(uev.get("stringValue", uev.get("numberValue", ""))).strip()
            note  = vals[1].get("note", "").strip()
            if col_a and note:
                notes[col_a] = note
        return notes
    except Exception:
        return {}


def _write_col_b_notes(worksheet, notes, all_new):
    """Zapisuje notatki do kol B na podstawie nowych pozycji wierszy."""
    if not notes:
        return
    sheet_id = worksheet._properties["sheetId"]
    col_a_to_rownum = {}
    for i, row in enumerate(all_new):
        col_a = str(row[0]).strip() if row else ""
        if col_a and col_a in notes and col_a not in col_a_to_rownum:
            col_a_to_rownum[col_a] = i + 1  # 1-based
    reqs = []
    for col_a, note in notes.items():
        row_num = col_a_to_rownum.get(col_a)
        if row_num:
            reqs.append({"updateCells": {
                "range": {
                    "sheetId":        sheet_id,
                    "startRowIndex":  row_num - 1,
                    "endRowIndex":    row_num,
                    "startColumnIndex": 1,
                    "endColumnIndex":   2,
                },
                "rows":   [{"values": [{"note": note}]}],
                "fields": "note",
            }})
    if reqs:
        _api(worksheet.spreadsheet.batch_update, {"requests": reqs})


def _batch_format_rows(worksheet, row_formats):
    """Wysyła wszystkie formatowania wierszy w jednym API call (batchUpdate).

    row_formats: lista (row_num_1based, format_dict)
    format_dict: np. {"backgroundColor": {...}} lub {"backgroundColor": ..., "textFormat": ...}
    """
    if not row_formats:
        return
    sheet_id = worksheet._properties["sheetId"]
    requests = []
    for row_num, fmt in row_formats:
        fields = ",".join(f"userEnteredFormat.{k}" for k in fmt.keys())
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_num - 1,
                    "endRowIndex": row_num,
                    "startColumnIndex": 0,
                    "endColumnIndex": 14,
                },
                "cell": {"userEnteredFormat": fmt},
                "fields": fields,
            }
        })
    _api(worksheet.spreadsheet.batch_update, {"requests": requests})


# ----------------------------------------------------------------
# KONFIGURACJA
# ----------------------------------------------------------------
FOLDER_ID            = "1kwY6tOalKS2jnidABw6uUV23ykMj1iR2"
FAKTURY_KOSZTOWE_ID  = "12RxQDakB6y9pxURM_Z73sS0fLNQyGtm1"
MIESZKANIA_FOLDER_ID = "1mvVZN6y2vaKyWGV6SIWd7FuK38T2DHAI"
SPREADSHEET_ID       = "1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0"
ABIDO_NAJEMCY_ID     = "1TuHpPvdZmGN_kXbAuhdA72hs8AKxaiOLQrOUpXh3uYA"
ABIDO_NAJEMCY_SHEET  = "Aktualni najemcy"
# ----------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

HEADER_ROW = [
    "Nazwa / Plik", "Kwota brutto", "Status",
    "Klucz_Ksiegowy", "wyciag_Kontrahent", "wyciag_Kwota",
    "Data_ksiegowania", "wyciag_Tytul",
    "wyciag_Data_op", "wyciag_Rodzaj", "wyciag_Waluta",
    "wyciag_Nr_rachunku", "wyciag_Imie_Nazwisko", "Uwagi",
]
SEP_KOSZTOWE = "--- FAKTURY KOSZTOWE ---"
SEP_WLASC    = "--- FAKTURY WLASCICIELE I SPOLDZIELNIE ---"
SEP_SPRZEDAZ = "--- FAKTURY SPRZEDAZY NAJEMCOM ---"
SEP_INNE_RK  = "--- INNE RAPORTY KASOWE ---"
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


def list_pdfs_from_drive(service, folder_id, include_images=False):
    if include_images:
        mime_filter = "(mimeType='application/pdf' or mimeType='image/jpeg' or mimeType='image/jpg')"
    else:
        mime_filter = "mimeType='application/pdf'"
    query = f"'{folder_id}' in parents and {mime_filter} and trashed=false"
    results = service.files().list(
        q=query, fields="files(id, name)", orderBy="name"
    ).execute()
    return results.get("files", [])


def download_pdf(service, file_id):
    return service.files().get_media(fileId=file_id).execute()


def extract_gross_amount(pdf_bytes):
    _NUM = r"(-?[\d ]+[,.][\d]{2})"
    patterns = [
        # "Do zapłaty" / "Pozostaje do zapłaty" / "Razem do zapłaty" — NAJWYZSZY PRIORYTET
        #   — pomijamy 0,00 (Allegro: zapłacone przy zakupie)
        r"(?:razem\s+|pozostaje\s+)?do\s+zap[lł]aty[^\d]*?" + _NUM,
        # "Kwota zapłaty (zaliczki) dokumentowana fakturą: X,XX" — KSeF EON faktury zaliczkowe
        # Tylko gdy nie ma "do zapłaty". Uzywamy [^\d:]* bo pdfplumber niekiedy
        # zwraca ł jako corrupted char. Wzorzec wymaga "(zaliczki)" aby uniknac false positive.
        r"kwota\s+zap[^\d(]*\([^\)]*zaliczki[^\)]*\)[^\d:]*:\s*" + _NUM,
        # "Wartość zamówienia lub umowy z uwzględnieniem kwoty podatku: X,XX" — KSeF EON fallback
        r"warto[^\d]*?kwoty\s+podatku\s*:\s*" + _NUM,
        # "Wartość brutto X,XX" jako osobna linia (np. Allegro)
        r"warto[śs][ćc]\s+brutto\s+" + _NUM,
        # "Należność X,XX" — np. E.ON
        r"nale[żz]no[śs][ćc]\s+" + _NUM,
        # "Razem brutto" / "Suma brutto" / "Ogółem brutto"
        r"(?:razem|suma|og[oó][lł]em)\s+brutto[^\d]*?" + _NUM,
        # "Kwota brutto: X,XX"
        r"kwota\s+brutto\s*[:\-]\s*" + _NUM,
        # "Kwota należności ogółem: X,XX PLN" — KSeF
        r"kwota\s+nale[żz]no[śs]ci\s+og[oó][lł]em\s*[:\-]\s*" + _NUM,
        # "Łączna kwota" / "Total"
        r"(?:[lł][aą]czna\s+kwota|total)\s*[:\-]?\s*" + _NUM,
    ]
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = "".join(page.extract_text() or "" for page in pdf.pages)
        if not text.strip():
            return ""
        tl = text.lower()
        for p in patterns:
            m = re.search(p, tl)
            if m:
                val = m.group(1).strip().replace(" ", "")
                if val in ("0,00", "0.00"):   # zapłacone z góry — szukaj dalej
                    continue
                # PDF z minusem (faktura korygujaca/zwrot) → w arkuszu na plusie (wpływ)
                # PDF bez minusa (zwykly koszt)            → w arkuszu na minusie (koszt)
                if val.startswith("-"):
                    return val[1:]
                return "-" + val
        # Ostatnia szansa: linia "Razem netto VAT brutto" — ostatnia liczba w linii
        for line in tl.splitlines():
            if re.match(r"\s*(?:\d+\.\s+)?razem\b", line):
                nums = re.findall(r"-?\d+[,.]\d{2}", line)
                if len(nums) >= 2:
                    val = nums[-1]
                    if val.startswith("-"):
                        return val[1:]
                    return "-" + val
    except Exception:
        pass
    return ""


def _clean_for_filename(name):
    """Czyści napis do użycia w nazwie pliku: polskie znaki → ASCII, spacje → _, usuwa znaki spec."""
    transl = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")
    name = name.translate(transl)
    name = re.sub(r"[^\w\s]", "", name)        # zostaw litery, cyfry, spacje, podkreślniki
    name = re.sub(r"\s+", "_", name.strip())   # spacje → _
    name = re.sub(r"_+", "_", name)            # wielokrotne _ → jedno
    return name


def extract_ksef_metadata(pdf_bytes):
    """
    Wyciąga z faktury KSeF:
      (seller, date_str, amount_str, numer_faktury, ksef_num)
    date_str: 'DD.MM.YYYY' lub None.
    amount_str: np. '379,91' (bez minusa) lub None.
    numer_faktury: numer nadany przez wystawce (str) lub None.
    ksef_num: numer KSeF (NIP-DATA-HEX-CRC) lub None.
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text = "".join(p.extract_text() or "" for p in pdf.pages)
    except Exception:
        return None, None, None, None, None

    # Sprzedawca: pierwsza "Nazwa: ..." nie zawierająca ABIDO (nabywca)
    seller = None
    for m in re.finditer(r"Nazwa:\s*(.+?)(?:\n|$)", text):
        raw = m.group(1).strip()
        # Format K365: dwa "Nazwa:" na jednej linii — weź fragment przed drugim
        if "Nazwa:" in raw:
            raw = raw[:raw.index("Nazwa:")].strip()
        if raw and "ABIDO" not in raw.upper():
            seller = raw
            break

    # Data wystawienia
    date_str = None
    md = re.search(r"Data wystawienia[^:]*:\s*(\d{2}\.\d{2}\.\d{4})", text)
    if md:
        date_str = md.group(1)

    # Kwota brutto (reużywamy istniejącej logiki)
    amount_raw = extract_gross_amount(pdf_bytes)   # zwraca np. "-379,91"
    amount_str = amount_raw.lstrip("-") if amount_raw else None

    # Numer Faktury wystawcy (linia po "Numer Faktury:")
    numer_faktury = None
    mf = re.search(r"Numer\s+Faktury\s*:?\s*\n?\s*([^\n]+)", text, re.IGNORECASE)
    if mf:
        numer_faktury = mf.group(1).strip()

    # Numer KSeF systemowy
    ksef_num = None
    mk = re.search(r"Numer\s+KSEF\s*:?\s*(\d{10}-\d{8}-[0-9A-Fa-f]+-[0-9A-Fa-f]+)",
                   text, re.IGNORECASE)
    if mk:
        ksef_num = mk.group(1).strip()

    return seller, date_str, amount_str, numer_faktury, ksef_num




# Wzorzec numeru KSeF: NIP(10)-DATA(8)-HEX-CRC
_KSEF_NUM_PAT = re.compile(r"\d{10}-\d{8}-[0-9A-Fa-f]+-[0-9A-Fa-f]+")
# Wzorzec do wyciagania numeru faktury z nazwy pliku KSeF (segment miedzy data a numerem KSeF)
# Wymaga pelnego numeru KSeF (NIP-YYYYMMDD-HEX-HEX) po numerze faktury —
# dzieki temu nie pasuje do zwyklych faktur z przypadkowym 10+8-cyfrowym ciagiem.
_KSEF_NR_IN_FNAME = re.compile(
    r"_\d{2}-\d{2}-\d{4}_(.+?)_\d{10}-\d{8}-[0-9A-Fa-f]+-[0-9A-Fa-f]+",
    re.IGNORECASE
)


def _build_ksef_map(drive_sa, folder_id):
    """
    Skanuje folder Drive, zwraca {numer_ksef: (file_id, file_name)}.
    Pliki z KSeF w nazwie (nowa konwencja) — wyciaga KSeF z nazwy (szybko).
    Pozostale (stara konwencja) — pobiera PDF i wyciaga KSeF z tresci (wolno, raz).
    """
    ksef_map = {}
    for f in list_pdfs_from_drive(drive_sa, folder_id):
        name = f["name"]
        m = _KSEF_NUM_PAT.search(name)
        if m:
            ksef_map[m.group()] = (f["id"], name)
        else:
            try:
                _, _, _, _, ksef_num = extract_ksef_metadata(
                    download_pdf(drive_sa, f["id"])
                )
                if ksef_num:
                    ksef_map[ksef_num] = (f["id"], name)
            except Exception:
                pass
    return ksef_map


def upload_ksef_from_zip_bytes(zip_bytes, drive_sa, drive_oauth, folder_id):
    """
    Rozpakowuje ZIP z fakturami KSeF w pamieci.
    Deduplikuje po Numerze KSeF wyciagnietym z tresci PDF.
    Pliki w starej konwencji (bez KSeF w nazwie) sa przemianowywane na nowa.
    Nowa nazwa: {Sprzedawca}_{DD-MM-YYYY}_{NumerFaktury}_{KSeF}.pdf
    Zwraca (uploaded, skipped, renamed, errors).
    """
    ksef_map = _build_ksef_map(drive_sa, folder_id)
    uploaded, skipped, renamed, errors = [], [], [], []

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
        for member in z.namelist():
            if not member.lower().endswith(".pdf"):
                continue
            original_name = os.path.basename(member)
            pdf_bytes     = z.read(member)

            seller, date_str, _, numer_faktury, ksef_num = extract_ksef_metadata(pdf_bytes)

            # Fallback KSeF z nazwy pliku jezeli PDF go nie zawiera
            if not ksef_num:
                m = _KSEF_NUM_PAT.search(original_name)
                if m:
                    ksef_num = m.group()

            if not ksef_num:
                errors.append(f"{original_name}: nie znaleziono Numeru KSeF")
                continue

            # Fallback daty z nazwy pliku: {NIP}-{YYYYMMDD}-...
            if not date_str:
                m = re.match(r"^\d+-(\d{4})(\d{2})(\d{2})-", original_name)
                if m:
                    date_str = f"{m.group(3)}.{m.group(2)}.{m.group(1)}"

            seller     = seller        or "KSeF"
            date_str   = date_str      or "brak-daty"
            nr_fakt    = _clean_for_filename(numer_faktury) if numer_faktury else "brak-nr"

            new_name = (
                f"{_clean_for_filename(seller)}_"
                f"{date_str.replace('.', '-')}_"
                f"{nr_fakt}_"
                f"{ksef_num}.pdf"
            )

            if ksef_num in ksef_map:
                existing_id, existing_name = ksef_map[ksef_num]
                if _KSEF_NUM_PAT.search(existing_name):
                    # Juz w nowej konwencji — pomij
                    skipped.append(original_name)
                else:
                    # Stara konwencja — przemianuj na nowa
                    try:
                        drive_sa.files().update(
                            fileId=existing_id, body={"name": new_name}
                        ).execute()
                        ksef_map[ksef_num] = (existing_id, new_name)
                        renamed.append(f"{existing_name} → {new_name}")
                    except Exception as exc:
                        errors.append(f"Rename '{existing_name}': {exc}")
                continue

            # Nowy plik — wgraj
            try:
                media = MediaIoBaseUpload(
                    io.BytesIO(pdf_bytes), mimetype="application/pdf", resumable=False
                )
                drive_oauth.files().create(
                    body={"name": new_name, "parents": [folder_id]},
                    media_body=media,
                ).execute()
                ksef_map[ksef_num] = (None, new_name)
                uploaded.append(new_name)
            except Exception as exc:
                errors.append(f"{original_name}: {exc}")

    return uploaded, skipped, renamed, errors


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
    Zwraca: name, address, price, dates
    """
    text = folder_name.replace("[FVS]", "").strip()
    parts = [p.strip() for p in text.split("|")]
    return {
        "name":    parts[0] if len(parts) > 0 else "",
        "address": parts[1] if len(parts) > 1 else "",
        "price":   parts[2] if len(parts) > 2 else "",
        "dates":   parts[3] if len(parts) > 3 else "",
    }


# ----------------------------------------------------------------
# GENEROWANIE FAKTUR SPRZEDAZY PDF
# ----------------------------------------------------------------

SELLER_NAME    = "ABIDO SP. Z O.O."
SELLER_ADDR1   = "ul. Henryka Sienkiewicza 85/87"
SELLER_ADDR2   = "90-057 Łódź"
SELLER_NIP     = "NIP: 7252283544"
SELLER_ACCOUNT = "Nr konta: 98 1870 1045 2078 1071 3878 0001"
FAKTURY_KOSZTOWE_SUFFIX  = "Faktury kosztowe"
FAKTURY_SPRZEDAZY_SUFFIX = "Faktury sprzedazy"

_PDF_FONTS_CACHED: dict = {}
_FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")


def _get_pdf_fonts():
    """Rejestruje czcionki DejaVu. Szuka w fonts/ obok app.py, potem w systemie."""
    global _PDF_FONTS_CACHED
    if _PDF_FONTS_CACHED:
        return _PDF_FONTS_CACHED["reg"], _PDF_FONTS_CACHED["bold"]

    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.pdfmetrics import registerFontFamily

    FONT_DIRS = [
        _FONTS_DIR,
        "/usr/share/fonts/truetype/dejavu",
        "/usr/share/fonts/dejavu",
    ]

    def find_font(fname):
        for d in FONT_DIRS:
            p = os.path.join(d, fname)
            if os.path.exists(p):
                return p
        return None

    reg_path  = find_font("DejaVuSans.ttf")
    bold_path = find_font("DejaVuSans-Bold.ttf")

    try:
        pdfmetrics.registerFont(TTFont("DejaVu", reg_path))
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", bold_path or reg_path))
        registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold",
                           italic="DejaVu", boldItalic="DejaVu-Bold")
        _PDF_FONTS_CACHED = {"reg": "DejaVu", "bold": "DejaVu-Bold"}
    except Exception:
        _PDF_FONTS_CACHED = {"reg": "Helvetica", "bold": "Helvetica-Bold"}

    return _PDF_FONTS_CACHED["reg"], _PDF_FONTS_CACHED["bold"]


def _gsheets_date_to_str(s):
    """Konwertuje date z gspread (rozne formaty) do 'D.M.YYYY' dla _parse_contract_start."""
    if not s:
        return ""
    s = str(s).strip()
    if re.match(r"\d{1,2}\.\d{1,2}\.\d{4}", s):
        return s
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        try:
            d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            return f"{d.day}.{d.month}.{d.year}"
        except ValueError:
            pass
    return s


def _read_najemcy_all(credentials):
    """Czyta caly arkusz Abido najemcy. Zwraca (invoice_rows, lookup_dict)."""
    client = gspread.authorize(credentials)
    ws = client.open_by_key(ABIDO_NAJEMCY_ID).worksheet(ABIDO_NAJEMCY_SHEET)
    all_rows = ws.get_all_values()
    if not all_rows:
        return [], {}

    header = [h.lower().strip() for h in all_rows[0]]

    def _ci(keywords):
        for kw in keywords:
            for i, h in enumerate(header):
                if kw in h:
                    return i
        return None

    idx_status   = _ci(["status do tworzenia", "status"])
    idx_name     = _ci(["imię i nazwisko", "nazwisko najemcy", "najemca", "najemc"])
    idx_lokal    = _ci(["adres mieszkania", "lokal mieszkalny", "adres", "lokal"])
    idx_koszt    = _ci(["koszt wynajmu", "koszt najmu", "koszt"])
    idx_umowa_od = _ci(["umowa od"])

    invoice_rows = []
    lookup       = {}

    for row in all_rows[1:]:
        if idx_name is None or len(row) <= idx_name or not row[idx_name].strip():
            continue
        name    = row[idx_name].strip()
        address = row[idx_lokal].strip() if idx_lokal is not None and len(row) > idx_lokal else ""
        dates   = _gsheets_date_to_str(row[idx_umowa_od]) if idx_umowa_od is not None and len(row) > idx_umowa_od else ""
        lookup[name] = {"address": address, "dates": dates}

        if idx_status is None or idx_koszt is None:
            continue
        try:
            status = float(row[idx_status])
        except (ValueError, TypeError):
            continue
        if status != 1.0:
            continue
        invoice_rows.append({
            "key":     name,
            "brutto":  row[idx_koszt].strip() if len(row) > idx_koszt else "",
            "address": "",
            "dates":   "",
            "_dates":  dates,  # tylko do sortowania, nie trafia do arkusza
        })

    return invoice_rows, lookup


def read_najemcy_for_invoices(credentials):
    """Czyta Status=1 z Abido najemcy. Kolejnosc: taka jak w pliku."""
    rows, _ = _read_najemcy_all(credentials)
    return rows


def read_najemcy_lookup(credentials):
    """Zwraca slownik {najemca: {address, dates}} z Abido najemcy."""
    _, lookup = _read_najemcy_all(credentials)
    return lookup


def _parse_contract_start(dates_str):
    """Parsuje date poczatku umowy z formatu 'D.M.YYYY-D.M.YYYY'. Zwraca date lub None."""
    if not dates_str:
        return None
    m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", dates_str.strip())
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def _normalize_name_for_filename(name):
    """Normalizuje nazwe do uzycia w nazwie pliku: male litery, bez polskich znakow, _ zamiast spacji."""
    name = name.replace("Ł", "L").replace("ł", "l")   # Ł nie rozklada sie przez NFKD
    nfkd = unicodedata.normalize("NFKD", name)
    ascii_str = nfkd.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", ascii_str.lower()).strip("_")


def _format_pln(amount):
    """Formatuje kwote jako polski string: 1 150,00 PLN"""
    formatted = f"{amount:,.2f}".replace(",", " ").replace(".", ",")
    return f"{formatted} PLN"


def _amount_words_pl(amount):
    """Kwota slownie po polsku."""
    try:
        from num2words import num2words
        whole = int(round(amount))
        gr    = round((amount - int(amount)) * 100)
        text  = num2words(whole, lang="pl")
        return f"{text} złotych {gr:02d}/100"
    except Exception:
        return f"{amount:.2f} PLN"


def build_invoice_pdf_bytes(invoice_data):
    """Generuje jedna fakture sprzedazy jako bajty PDF."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT, TA_CENTER

    fn, fnb = _get_pdf_fonts()

    def ps(name, **kw):
        base = {"fontName": fn, "fontSize": 9, "leading": 13, "spaceAfter": 0}
        base.update(kw)
        return ParagraphStyle(name, **base)

    s_n   = ps("n")
    s_r   = ps("r",  alignment=TA_RIGHT)
    s_c   = ps("c",  alignment=TA_CENTER)
    s_rb  = ps("rb", fontName=fnb, alignment=TA_RIGHT)
    s_cb  = ps("cb", fontName=fnb, alignment=TA_CENTER)
    s_sm  = ps("sm", fontSize=7.5, leading=10,
               textColor=colors.Color(0.35, 0.35, 0.35))
    s_ttl = ps("ttl", fontName=fnb, fontSize=16, leading=20)
    s_wb  = ps("wb", fontName=fnb, fontSize=13, leading=18,
               textColor=colors.white)
    s_wn  = ps("wn", textColor=colors.white)

    inv     = invoice_data
    amt     = inv["amount"]
    amt_str = _format_pln(amt)

    GREY  = colors.Color(0.90, 0.90, 0.90)
    DARK  = colors.Color(0.15, 0.15, 0.15)
    BDR   = colors.Color(0.70, 0.70, 0.70)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    W = A4[0] - 4*cm  # ~481 pt

    story = []

    # ── 1. Naglowek ───────────────────────────────────────────────────
    hdr = Table([[
        Paragraph("FAKTURA", s_ttl),
        Paragraph(
            f"<b>{inv['invoice_nr']}</b><br/>"
            f"Data wystawienia: {inv['issue_date'].strftime('%d.%m.%Y')}<br/>"
            f"Data sprzedaży: {inv['sale_date'].strftime('%d.%m.%Y')}",
            s_r
        ),
    ]], colWidths=[W * 0.5, W * 0.5])
    hdr.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW",    (0, 0), (-1,  0), 1.5, colors.black),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.35 * cm))

    # ── 2. Sprzedawca / Nabywca ───────────────────────────────────────
    seller = (
        f"<b>SPRZEDAWCA</b><br/><b>{SELLER_NAME}</b><br/>"
        f"{SELLER_ADDR1}<br/>{SELLER_ADDR2}<br/>"
        f"{SELLER_NIP}<br/>{SELLER_ACCOUNT}"
    )
    buyer = f"<b>NABYWCA</b><br/><b>{inv['buyer_name']}</b>"

    parties = Table(
        [[Paragraph(seller, s_n), Paragraph(buyer, s_n)]],
        colWidths=[W * 0.55, W * 0.45]
    )
    parties.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("BOX",          (0, 0), ( 0,  0), 0.5, BDR),
        ("BOX",          (1, 0), ( 1,  0), 0.5, BDR),
        ("BACKGROUND",   (0, 0), ( 0,  0), GREY),
        ("LEFTPADDING",  (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING",   (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
    ]))
    story.append(parties)
    story.append(Spacer(1, 0.4 * cm))

    # ── 3. Opis uslugi ────────────────────────────────────────────────
    svc_desc = (
        f"Wynajem pokoju w lokalu mieszkalnym na cele mieszkaniowe"
        f" — {inv['service_address']}"
    )
    svc = Table(
        [
            [Paragraph("<b>Opis usługi</b>", ps("bh", fontName=fnb)),
             Paragraph("<b>Ilość</b>", s_cb),
             Paragraph("<b>J.m.</b>", s_cb),
             Paragraph("<b>Cena jedn.</b>", s_rb),
             Paragraph("<b>Wartość brutto</b>", s_rb)],
            [Paragraph(svc_desc, s_n),
             Paragraph("1,00", s_c),
             Paragraph("szt.", s_c),
             Paragraph(amt_str, s_r),
             Paragraph(amt_str, s_r)],
        ],
        colWidths=[W * 0.44, W * 0.12, W * 0.08, W * 0.18, W * 0.18]
    )
    svc.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0), GREY),
        ("BOX",          (0, 0), (-1,-1), 0.5, BDR),
        ("INNERGRID",    (0, 0), (-1,-1), 0.25, BDR),
        ("VALIGN",       (0, 0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (-1,-1), 6),
        ("RIGHTPADDING", (0, 0), (-1,-1), 6),
        ("TOPPADDING",   (0, 0), (-1,-1), 5),
        ("BOTTOMPADDING",(0, 0), (-1,-1), 5),
    ]))
    story.append(svc)
    story.append(Spacer(1, 0.35 * cm))

    # ── 4. Tabela kwot ────────────────────────────────────────────────
    amts = Table(
        [
            [Paragraph("Wartość netto:", s_r),     Paragraph(amt_str, s_r)],
            [Paragraph("Stawka VAT:", s_r),         Paragraph("zwolniony", s_r)],
            [Paragraph("Kwota VAT:", s_r),          Paragraph(_format_pln(0), s_r)],
            [Paragraph("<b>Wartość brutto:</b>", s_rb),
             Paragraph(f"<b>{amt_str}</b>", s_rb)],
        ],
        colWidths=[W * 0.72, W * 0.28]
    )
    amts.setStyle(TableStyle([
        ("BOX",          (0, 0), (-1,-1), 0.5, BDR),
        ("LINEABOVE",    (0, 3), (-1, 3), 0.5, BDR),
        ("BACKGROUND",   (0, 3), (-1, 3), GREY),
        ("RIGHTPADDING", (0, 0), (-1,-1), 8),
        ("LEFTPADDING",  (0, 0), (-1,-1), 6),
        ("TOPPADDING",   (0, 0), (-1,-1), 4),
        ("BOTTOMPADDING",(0, 0), (-1,-1), 4),
    ]))
    story.append(amts)
    story.append(Spacer(1, 0.4 * cm))

    # ── 5. Info platnosci ─────────────────────────────────────────────
    story.append(Paragraph(
        f"Metoda płatności: <b>{inv['payment_method']}</b>"
        f"   |   Termin płatności: <b>{inv['payment_deadline'].strftime('%d.%m.%Y')}</b>",
        s_n
    ))
    story.append(Spacer(1, 0.3 * cm))

    # ── 6. Podstawa zwolnienia VAT ────────────────────────────────────
    story.append(Paragraph(
        "Podstawa zwolnienia: Zwolnienie z VAT na podstawie art. 43 ust. 1 pkt 36 "
        "ustawy z dnia 11 marca 2004 r. o podatku od towarów i usług — "
        "wynajem lokali mieszkalnych na cele mieszkaniowe.",
        s_sm
    ))
    story.append(Spacer(1, 0.5 * cm))

    # ── 7. DO ZAPLATY ─────────────────────────────────────────────────
    words = _amount_words_pl(amt)
    zapł  = Table(
        [[Paragraph(f"DO ZAPŁATY: {amt_str}", s_wb)],
         [Paragraph(f"słownie: {words}", s_wn)]],
        colWidths=[W]
    )
    zapł.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1,-1), DARK),
        ("LEFTPADDING",  (0, 0), (-1,-1), 12),
        ("RIGHTPADDING", (0, 0), (-1,-1), 12),
        ("TOPPADDING",   (0, 0), ( 0, 0), 10),
        ("BOTTOMPADDING",(0, 0), ( 0, 0), 3),
        ("TOPPADDING",   (0, 1), ( 0, 1), 3),
        ("BOTTOMPADDING",(0, 1), ( 0, 1), 10),
    ]))
    story.append(zapł)

    doc.build(story)
    return buf.getvalue()


def merge_pdf_bytes(pdf_bytes_list):
    """Scala liste bajtow PDF w jeden plik PDF."""
    from pypdf import PdfWriter, PdfReader
    writer = PdfWriter()
    for pdf_b in pdf_bytes_list:
        reader = PdfReader(io.BytesIO(pdf_b))
        for page in reader.pages:
            writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def get_or_create_subfolder(service, parent_id, folder_name):
    """Znajduje lub tworzy podfolder na Drive. Zwraca ID folderu."""
    folder = find_subfolder(service, parent_id, folder_name)
    if folder:
        return folder["id"]
    metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    return service.files().create(body=metadata, fields="id").execute()["id"]


def upload_file_to_drive(service, folder_id, filename, content_bytes,
                         mimetype="application/pdf"):
    """Wgrywa bajty jako plik na Drive. Zastepuje istniejacy plik o tej samej nazwie."""
    q = (f"'{folder_id}' in parents and name = '{filename}' "
         "and trashed = false")
    existing = (service.files().list(q=q, fields="files(id)")
                .execute().get("files", []))
    for f in existing:
        service.files().delete(fileId=f["id"]).execute()

    media = MediaIoBaseUpload(io.BytesIO(content_bytes), mimetype=mimetype)
    service.files().create(
        body={"name": filename, "parents": [folder_id]},
        media_body=media, fields="id"
    ).execute()


def _get_user_drive_service():
    """
    Zwraca Drive service uzywajac user OAuth2 credentials ze secrets.
    Wymaga sekcji [google_drive_oauth] z client_id, client_secret, refresh_token.
    Zwraca None jesli credentials nie sa skonfigurowane.
    """
    oauth = dict(st.secrets.get("google_drive_oauth", {}))
    if not oauth.get("refresh_token"):
        return None
    try:
        from google.oauth2.credentials import Credentials
        creds = Credentials(
            token=None,
            refresh_token=oauth["refresh_token"],
            client_id=oauth["client_id"],
            client_secret=oauth["client_secret"],
            token_uri="https://oauth2.googleapis.com/token",
        )
        return build("drive", "v3", credentials=creds)
    except Exception:
        return None


def _invoices_summary(invoices):
    """Zwraca (ilosc, suma_int) na podstawie kwot z nazw plikow fvs_..._KWOTA.pdf."""
    total = 0
    for fname, _ in invoices:
        parts = fname.replace(".pdf", "").split("_")
        try:
            total += int(parts[-1])
        except (ValueError, IndexError):
            pass
    return len(invoices), total


def _merged_filename(subfolder_name, invoices):
    count, total = _invoices_summary(invoices)
    today = date.today().strftime("%d%m%Y")
    return f"Fs_najemcy_{subfolder_name}_{count}szt_{total}zl_{today}.pdf"


def upload_invoices_to_drive(user_drive_service, invoices, subfolder_name):
    """
    Wgrywa faktury PDF na Drive uzywajac user OAuth credentials.
    Tworzy folder 'Faktury sprzedazy MMRRRR' wewnatrz 'Faktury-sprzedazy'.
    """
    folder_name = f"{subfolder_name} {FAKTURY_SPRZEDAZY_SUFFIX}"
    parent_id   = get_or_create_subfolder(user_drive_service, FOLDER_ID, "Faktury-sprzedazy")
    folder_id   = get_or_create_subfolder(user_drive_service, parent_id, folder_name)
    for filename, pdf_b in invoices:
        upload_file_to_drive(user_drive_service, folder_id, filename, pdf_b)
    if invoices:
        merged = merge_pdf_bytes([b for _, b in invoices])
        upload_file_to_drive(
            user_drive_service, folder_id,
            _merged_filename(subfolder_name, invoices), merged
        )
    return folder_name


def generate_invoice_pdfs(drive_service, worksheet, subfolder_name, credentials=None):
    """
    Generuje PDF faktur sprzedazy dla wszystkich wierszy sekcji SPRZEDAZ.
    Zwraca liste (filename, pdf_bytes) — bez wgrywania na Drive.
    """
    month = int(subfolder_name[:2])
    year  = int(subfolder_name[2:])
    last_day         = calendar.monthrange(year, month)[1]
    sale_date        = date(year, month, last_day)
    default_issue    = date(year, month, 1)
    payment_deadline = sale_date

    # Dane najemcow (adres, data umowy) z Abido najemcy
    najemcy_lookup = read_najemcy_lookup(credentials) if credentials else {}

    # Wiersze sekcji SPRZEDAZ
    sections = read_all_sections(worksheet)
    rows = sections[SEP_SPRZEDAZ]
    if not rows:
        return []

    _get_pdf_fonts()  # rejestracja czcionek przed generowaniem

    results = []  # lista (filename, pdf_bytes)

    num = 0
    for row in rows:
        name       = row[0] if len(row) > 0 else ""
        amount_str = row[1] if len(row) > 1 else ""
        klucz      = row[5] if len(row) > 5 else ""

        if not name:
            continue

        # Kwota 0 = najemca zaplacil z gory za ten miesiac — faktura nie jest wystawiana
        try:
            _chk = abs(float(re.sub(r"[^\d,.]", "", str(amount_str)).replace(",", ".")))
        except (ValueError, TypeError):
            _chk = 0.0
        if _chk == 0.0:
            continue

        num += 1
        nj        = najemcy_lookup.get(name, {})
        address   = nj.get("address", "")
        dates_str = nj.get("dates",   "")

        # Kwota (zawsze dodatnia na fakturze)
        try:
            amount = abs(float(
                re.sub(r"[^\d,.]", "", str(amount_str)).replace(",", ".")
            ))
        except (ValueError, TypeError):
            amount = 0.0

        # Data wystawienia: jezeli najemca wszedl w srodku miesiaca, faktura
        # ma date wystawienia = data startu umowy (nie 1. dnia miesiaca).
        # Kwota NIE jest redukowana — uzytkownik wpisuje juz poprawna kwote w kol. B.
        issue_date = default_issue
        contract_start = _parse_contract_start(dates_str)
        if (contract_start
                and contract_start.year == year
                and contract_start.month == month
                and contract_start.day > 1):
            issue_date = contract_start

        payment_method = "Przelew" if "pr_in" in klucz else "Gotówka"
        invoice_nr     = f"FVS {year} {month:02d} {num:02d} T"
        name_norm      = _normalize_name_for_filename(name)
        amt_str_fn     = f"{int(round(amount))}"
        filename       = f"fvs_{year}_{month:02d}_{num:02d}_t_{name_norm}_{amt_str_fn}.pdf"

        try:
            pdf_b = build_invoice_pdf_bytes({
                "invoice_nr":       invoice_nr,
                "issue_date":       issue_date,
                "sale_date":        sale_date,
                "buyer_name":       name,
                "service_address":  address,
                "amount":           amount,
                "payment_method":   payment_method,
                "payment_deadline": payment_deadline,
            })
            results.append((filename, pdf_b))
        except Exception:
            pass

    return results


SECTION_ORDER = [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC, SEP_INNE_RK, SEP_NIEZNANE]

# Liczba pustych wierszy dodawanych po kazdej sekcji (w rebuild_sheet i create_month_template)
SECTION_BLANK_ROWS = {
    SEP_KOSZTOWE: 15,
    SEP_SPRZEDAZ: 15,
    SEP_WLASC:    44,
    SEP_INNE_RK:  15,
    SEP_NIEZNANE: 15,
}


def get_or_create_worksheet(spreadsheet, sheet_name):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=6)


def create_month_template(spreadsheet, sheet_name):
    """
    Tworzy szablon miesiaca w arkuszu sheet_name.
    - Jesli arkusz nie istnieje: tworzy go z wszystkimi sekcjami i 15 pustymi wierszami pod kazdą.
    - Jesli istnieje: sprawdza brakujace sekcje i dodaje je w odpowiednim miejscu z 15 pustymi wierszami.
    Zwraca (status, added_sections):
      status = "created"         — nowy arkusz
               "exists_partial"  — arkusz istnial, dodano brakujace sekcje
               "exists_complete" — arkusz istnial, wszystkie sekcje juz byly
    """
    _BLANK_ROW = [""] * len(HEADER_ROW)

    def _blank_rows(sep):
        return [_BLANK_ROW] * SECTION_BLANK_ROWS.get(sep, 15)

    def _color_sep(ws, row_num, sep):
        _api(ws.format, f"A{row_num}:N{row_num}", {
            "backgroundColor": SEP_COLORS[sep],
            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
        })

    try:
        ws = spreadsheet.worksheet(sheet_name)
        existed = True
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=200, cols=20)
        existed = False

    if not existed:
        # Nowy arkusz — wstaw naglowek + wszystkie sekcje z 15 pustymi wierszami
        all_rows = [HEADER_ROW]
        sep_row_nums = {}
        for sep in SECTION_ORDER:
            all_rows.append([sep] + [""] * (len(HEADER_ROW) - 1))
            sep_row_nums[sep] = len(all_rows)
            all_rows.extend(_blank_rows(sep))
        _api(ws.update, "A1", all_rows, value_input_option="USER_ENTERED")
        for sep, rn in sep_row_nums.items():
            _color_sep(ws, rn, sep)
        return "created", list(SECTION_ORDER)

    # Istniejacy arkusz — sprawdz ktore sekcje sa obecne
    all_vals = _api(ws.get_all_values)
    present_seps = {}
    for i, row in enumerate(all_vals):
        val = row[0] if row else ""
        m = _match_separator(val)
        if m:
            present_seps[m] = i + 1  # 1-based

    missing = [sep for sep in SECTION_ORDER if sep not in present_seps]
    if not missing:
        return "exists_complete", []

    # Wstaw brakujace sekcje w odpowiednim miejscu (od gory do dolu z kumulowanym offsetem)
    row_shift = 0
    added = []
    for sep in SECTION_ORDER:
        if sep not in missing:
            continue
        sep_idx = SECTION_ORDER.index(sep)
        insert_at = None
        for next_sep in SECTION_ORDER[sep_idx + 1:]:
            if next_sep in present_seps:
                insert_at = present_seps[next_sep] + row_shift
                break
        rows_to_insert = [[sep] + [""] * (len(HEADER_ROW) - 1)] + _blank_rows(sep)
        if insert_at is not None:
            _api(ws.insert_rows, rows_to_insert, insert_at, value_input_option="USER_ENTERED")
        else:
            _api(ws.append_rows, rows_to_insert, value_input_option="USER_ENTERED")
        row_shift += len(rows_to_insert)
        added.append(sep)

    # Pokoloruj nowo wstawione separatory
    all_vals2 = _api(ws.get_all_values)
    for i, row in enumerate(all_vals2):
        val = row[0] if row else ""
        m = _match_separator(val)
        if m and m in added:
            _color_sep(ws, i + 1, m)

    return "exists_partial", added


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


def _migrate_row(row):
    """Konwertuje stary wiersz 17-kolumnowy (D/E/F = raport_kasowy/Adres/Data_umowy)
    do nowego 14-kolumnowego (A,B,C + Klucz..Uwagi).
    Wywolywane automatycznie przy odczycie — jednorazowa migracja bez akcji uzytkownika.
    """
    if len(row) > 14:
        row = list(row[:3]) + list(row[6:])
    else:
        row = list(row)
    return row


def read_all_sections(worksheet):
    """
    Czyta caly arkusz i zwraca slownik sekcji:
    { SEP_KOSZTOWE: [rows...], SEP_SPRZEDAZ: [rows...], ... }
    Wiersze separatorow sa pomijane.
    """
    all_rows = worksheet.get_all_values()
    sections   = {sep: [] for sep in SECTION_ORDER}
    current    = None
    in_summary = False   # pomijamy wiersze tabeli podsumowania (add_section_summary)
    for row in all_rows:
        val = row[0] if row else ""
        matched = _match_separator(val)
        if matched:
            current    = matched
            in_summary = False
            continue
        # Wykryj naglowek tabeli podsumowania: col A="Segment", col B zawiera "pozycji"
        if (str(val).strip() == "Segment"
                and len(row) > 1 and "pozycji" in str(row[1]).lower()):
            in_summary = True
        if in_summary:
            continue
        if not current:
            continue
        row = _migrate_row(row)   # migracja 17-kol → 14-kol (jesli stary format)
        if current == SEP_NIEZNANE:
            # SEP_NIEZNANE: brak faktury w A — wlaczamy jesli cokolwiek wypelnione
            if any(c for c in row):
                sections[current].append(row)
        elif val or str(row[2]).strip() == "3":
            # Normalny wiersz — klucz w kol A, lub status=3 (beton) bez nazwy pliku
            sections[current].append(row)
        elif (len(row) > 4 and row[4]) or (len(row) > 5 and row[5]):
            # Dodatkowy wiersz parowania: kontrahent w kol E LUB wyciag_Kwota w kol F
            sections[current].append(row)
    return sections


_PURPLE_MARKER = "_purple"
_PURPLE_BG     = {"red": 0.87, "green": 0.78, "blue": 0.97}
_ORANGE_MARKER = "_orange"
_ORANGE_BG     = {"red": 1.0,  "green": 0.88, "blue": 0.70}
_KP_BG         = {"red": 0.82, "green": 0.96, "blue": 0.77}   # kasa przyjela (gotowka wplyn)
_KW_BG         = {"red": 0.97, "green": 0.82, "blue": 0.82}   # kasa wyplacila (gotowka wydat)
_MULTI_BG      = {"red": 0.91, "green": 0.88, "blue": 0.98}   # multi-parowanie (1 faktura → kilka TX)
_ZERO_SPRZEDAZ_BG = {"red": 0.55, "green": 0.55, "blue": 0.55}  # sprzedaz kwota=0 (zaplacone z gory)
_FROZEN3_BG    = {"red": 0.78, "green": 0.78, "blue": 0.78}   # status=3 beton (szary)

SEP_COLORS = {
    SEP_KOSZTOWE: {"red": 0.90, "green": 0.22, "blue": 0.22},  # czerwony
    SEP_SPRZEDAZ: {"red": 0.18, "green": 0.65, "blue": 0.32},  # zielony
    SEP_WLASC:    {"red": 0.95, "green": 0.55, "blue": 0.10},  # pomaranczowy
    SEP_INNE_RK:  {"red": 0.18, "green": 0.55, "blue": 0.65},  # ciemny turkus
    SEP_NIEZNANE: {"red": 0.50, "green": 0.50, "blue": 0.50},  # szary
}


def rebuild_sheet(worksheet, sections, blank_rows=None):
    """
    Zapisuje caly arkusz w poprawnej kolejnosci sekcji:
    Header → Kosztowe → Sprzedaz → Wlasciciele.
    Pomija puste sekcje. Koloruje separatory i wiersze kp/kw.
    blank_rows: dict {sep: n} ile pustych wierszy pod sekcja. None = domyslne SECTION_BLANK_ROWS.
                Przekaz {} zeby nie dodawac zadnych pustych wierszy.
    """
    if blank_rows is None:
        blank_rows = SECTION_BLANK_ROWS
    all_new = [HEADER_ROW]
    sep_row_nums  = {}   # sep -> numer wiersza (1-based)
    kp_rows       = []   # wiersze z _rk_kp w kluczu (col D)
    kw_rows       = []   # wiersze z _rk_kw w kluczu (col D)
    multi_rows    = []   # wiersze nalezace do multi-parowania (1 faktura → kilka TX)
    purple_rows   = []   # sparowane po nazwie bez zgodnosci kwoty (col B ≠ col F)
    frozen3_rows  = []   # wiersze ze statusem=3 (nietykalny kolor i pozycja)
    zero_sprzedaz_rows = []  # sprzedaz z kwota=0 w col B
    last_main_num = None # numer wiersza ostatniego "glownego" wiersza (col A niepuste)
    for sep in SECTION_ORDER:
        all_new.append([sep, "", "", ""])
        sep_row_nums[sep] = len(all_new)
        last_main_num = None
        for row in sections[sep]:
            all_new.append(row)
            row_num = len(all_new)
            klucz = str(row[3]).strip() if len(row) > 3 else ""
            col_a = str(row[0]).strip() if row else ""
            col_c = str(row[2]).strip() if len(row) > 2 else ""
            col_h = str(row[4]).strip() if len(row) > 4 else ""
            if col_c == "3":
                frozen3_rows.append(row_num)
                last_main_num = None
                continue   # status=3: nie dotykaj koloru ani pozycji
            if col_a:
                last_main_num = row_num   # zapamietaj glowny wiersz
                # Fioletowy: glowny wiersz sparowany po nazwie, kwota niezgodna (col B ≠ col F)
                # status=2 NIE chroni koloru — tylko status=3 jest beton
                if col_h and col_c not in ("3",):
                    try:
                        _inv = abs(round(float(
                            re.sub(r"[^\d,.\-]", "", str(row[1] if len(row) > 1 else "")).replace(",", ".")
                        ), 2))
                        _tx  = abs(round(float(
                            re.sub(r"[^\d,.\-]", "", str(row[5] if len(row) > 5 else "")).replace(",", ".")
                        ), 2))
                        if _inv and _tx and _inv != _tx:
                            purple_rows.append(row_num)
                    except (ValueError, TypeError):
                        pass
            elif not col_a and col_h:
                # Sub-wiersz (puste A, jest kontrahent) = multi-parowanie
                if last_main_num is not None:
                    if last_main_num not in multi_rows:
                        multi_rows.append(last_main_num)
                multi_rows.append(row_num)
            else:
                last_main_num = None
            if "_rk_kp" in klucz:
                kp_rows.append(row_num)
            elif "_rk_kw" in klucz:
                kw_rows.append(row_num)
            if sep == SEP_SPRZEDAZ and col_a:
                try:
                    _amt = abs(float(
                        re.sub(r"[^\d,.\-]", "", str(row[1] if len(row) > 1 else "")).replace(",", ".")
                    ))
                    if _amt == 0.0:
                        zero_sprzedaz_rows.append(row_num)
                except (ValueError, TypeError):
                    pass
        _blank = [""] * len(HEADER_ROW)
        for _ in range(blank_rows.get(sep, 0)):
            all_new.append(_blank)
    _col_b_notes = _read_col_b_notes(worksheet)
    _api(worksheet.clear)
    # Reset formatowania — white dla wszystkich POZA status=3 (te zachowuja kolor uzytkownika)
    _white_fmt = {
        "backgroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
        "textFormat": {"bold": False},
        "horizontalAlignment": "CENTER",
    }
    _frozen3_set = set(frozen3_rows)
    if not _frozen3_set:
        _api(worksheet.format, "A1:N500", _white_fmt)
    else:
        _sheet_id   = worksheet._properties["sheetId"]
        _fields     = ("userEnteredFormat.backgroundColor,"
                       "userEnteredFormat.textFormat,"
                       "userEnteredFormat.horizontalAlignment")
        _w_reqs = []
        _start  = 1
        for _r in sorted(_frozen3_set) + [len(all_new) + 1]:
            if _r > _start:
                _w_reqs.append({"repeatCell": {
                    "range": {"sheetId": _sheet_id,
                              "startRowIndex": _start - 1, "endRowIndex": _r - 1,
                              "startColumnIndex": 0,     "endColumnIndex": 14},
                    "cell": {"userEnteredFormat": _white_fmt},
                    "fields": _fields,
                }})
            _start = _r + 1
        if _w_reqs:
            _api(worksheet.spreadsheet.batch_update, {"requests": _w_reqs})
    if all_new:
        _api(worksheet.update, "A1", all_new, value_input_option="USER_ENTERED")
    _write_col_b_notes(worksheet, _col_b_notes, all_new)
    row_fmts = []
    for sep, row_num in sep_row_nums.items():
        row_fmts.append((row_num, {
            "backgroundColor": SEP_COLORS[sep],
            "textFormat": {"bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
        }))
    for row_num in multi_rows:
        row_fmts.append((row_num, {"backgroundColor": _MULTI_BG}))
    for row_num in kp_rows:
        row_fmts.append((row_num, {"backgroundColor": _KP_BG}))
    for row_num in kw_rows:
        row_fmts.append((row_num, {"backgroundColor": _KW_BG}))
    for row_num in purple_rows:
        row_fmts.append((row_num, {"backgroundColor": _PURPLE_BG}))
    for row_num in zero_sprzedaz_rows:
        row_fmts.append((row_num, {"backgroundColor": _ZERO_SPRZEDAZ_BG}))
    # status=3: kolor nie jest nadpisywany — uzytkownik ustawia go recznie i jest zachowany
    _batch_format_rows(worksheet, row_fmts)
    # Jawnie ustaw format liczbowy dla wyciag_Kwota (kol F) — bez tego kol. dziedziczy
    # format DATE po starej kolumnie i liczby wyswietlaja sie jako daty.
    _api(worksheet.format, "F2:F500", {"numberFormat": {"type": "NUMBER", "pattern": "0.00"}})
    # Przytnij tekst w wyciag_Imie_Nazwisko (kol M) — nie wylewa sie do pustej kol Uwagi
    _api(worksheet.format, "M1:M500", {"wrapStrategy": "CLIP"})


def apply_sync_logic(existing_rows, new_data, has_address=False, default_status="0"):
    """
    Laczy istniejace zweryfikowane wiersze (C=1) z nowymi danymi.
    Wiersze z C=1 sa zachowane nawet jesli plik zostal usuniety z Drive.
    Sub-wiersze (puste A, dane wyciagu w H) sa zawsze zachowywane razem z rodzicem.
    Zwraca (nowe_wiersze, skipped, added).
    """
    # Kotwice status=3 — wyjmij przed przetwarzaniem, przywroc na oryginalne pozycje po
    frozen3  = [(i, row) for i, row in enumerate(existing_rows)
                if str(row[2] if len(row) > 2 else "").strip() == "3"]
    working  = [row for row in existing_rows
                if str(row[2] if len(row) > 2 else "").strip() != "3"]
    # Sub-wiersze: puste col A, ale maja dane wyciagu (col E) — zachowujemy osobno
    sub_rows = [row for row in working
                if not (row[0] if row else "") and len(row) > 4 and row[4]]
    main_rows = [row for row in working
                 if (row[0] if row else "")]

    verified = {
        row[0]: row
        for row in main_rows
        if len(row) > 2 and str(row[2]).strip() in ("1", "2", "9")
    }
    new_keys = {item["key"] for item in new_data}
    result = []
    for item in new_data:
        key = item["key"]
        if key in verified:
            result.append(verified[key])
        else:
            addr  = item.get("address", "") if has_address else ""
            dates = item.get("dates",   "") if has_address else ""
            result.append([key, item.get("brutto", ""), item.get("status", default_status), "", addr, dates])
    # Zachowaj zweryfikowane wiersze ktorych plik zostal usuniety z Drive
    for key, row in verified.items():
        if key not in new_keys:
            result.append(row)
    # Dolacz sub-wiersze na koniec (beda po rebuild w sekcji, nie ma lepszego miejsca)
    result.extend(sub_rows)
    # Przywroc kotwice status=3 na ich oryginalne pozycje
    for pos, row in sorted(frozen3, key=lambda x: x[0]):
        result.insert(min(pos, len(result)), row)
    return result, len(verified), len(new_data) - len(verified)


def add_empty_rows_to_segment(worksheet, sep, n):
    """Dodaje n pustych wierszy na koniec wybranego segmentu."""
    sections = read_all_sections(worksheet)
    blank = [""] * len(HEADER_ROW)
    sections[sep].extend([blank[:] for _ in range(n)])
    rebuild_sheet(worksheet, sections, blank_rows={})


def sort_kosztowe(worksheet):
    """
    Sortuje SEP_KOSZTOWE: normalne → _rk_kw → _rk_kp na dole.
    Status=3 to kotwice — zostają na swoich pozycjach.
    Zwraca liczbę wierszy bez kotwic.
    """
    sections = read_all_sections(worksheet)
    rows = sections[SEP_KOSZTOWE]

    anchors     = [(i, r) for i, r in enumerate(rows)
                   if str(r[2] if len(r) > 2 else "").strip() == "3"]
    non_anchors = [r for r in rows
                   if str(r[2] if len(r) > 2 else "").strip() != "3"]

    normal = [r for r in non_anchors if "_rk_kw" not in str(r[3] if len(r) > 3 else "")
                                     and "_rk_kp" not in str(r[3] if len(r) > 3 else "")]
    rk_kw  = [r for r in non_anchors if "_rk_kw" in str(r[3] if len(r) > 3 else "")]
    rk_kp  = [r for r in non_anchors if "_rk_kp" in str(r[3] if len(r) > 3 else "")]

    result = normal + rk_kw + rk_kp
    for pos, row in sorted(anchors, key=lambda x: x[0]):
        result.insert(pos, row)

    sections[SEP_KOSZTOWE] = result
    rebuild_sheet(worksheet, sections, blank_rows={})
    return len(non_anchors)


def sort_inne_rk_nieznane(worksheet):
    """
    Sortuje SEP_INNE_RK i SEP_NIEZNANE.
    Status=3 to kotwice — pozostaja na swoich pozycjach, reszta sortuje sie wokol nich.
    Zwraca (n_inne_rk, n_nieznane) — liczby posortowanych wierszy bez kotwic.
    """
    sections = read_all_sections(worksheet)

    def _with_anchors(rows, key_fn):
        anchors     = [(i, r) for i, r in enumerate(rows)
                       if str(r[2] if len(r) > 2 else "").strip() == "3"]
        non_anchors = [r for r in rows
                       if str(r[2] if len(r) > 2 else "").strip() != "3"]
        result = sorted(non_anchors, key=key_fn)
        for pos, row in sorted(anchors, key=lambda x: x[0]):
            result.insert(pos, row)
        return result, len(non_anchors)

    def _date_key(row):
        date_s = str(row[6]).strip() if len(row) > 6 else ""
        m = re.match(r'^(\d{4}-\d{2}-\d{2})', date_s)
        return m.group(1) if m else ("9999" if not date_s else date_s)

    def _inne_rk_key(row):
        k = str(row[3]).strip().lower() if len(row) > 3 else ""
        if   "depo" in k and ("pr_in" in k or "_in" in k): p = 0  # depo otrzymane
        elif "depo" in k:                                   p = 1  # depo wydane
        elif "bankomat" in k and "_kp" in k:               p = 2  # wpłata bankomatowa
        elif "bankomat" in k and "_kw" in k:               p = 3  # wypłata bankomatowa
        else:                                               p = 4  # pozostałe
        return (p, _date_key(row), str(row[0]).strip().lower() if row else "")

    def _nieznane_key(row):
        k = str(row[3]).strip().lower() if len(row) > 3 else ""
        if   "depo" in k and "pr_in"  in k:               p = 0  # depo pr_in
        elif "depo" in k and "pr_out" in k:               p = 1  # depo pr_out
        elif "depo" in k:                                  p = 2  # inne depo
        elif "_in"  in k or "_kp" in k:                   p = 3  # wszystkie wpływy
        elif "_out" in k or "_kw" in k:                   p = 4  # wszystkie wypływy
        else:                                              p = 5  # brak klucza
        return (p, _date_key(row))

    sections[SEP_INNE_RK],  n_inne  = _with_anchors(sections[SEP_INNE_RK],  _inne_rk_key)
    sections[SEP_NIEZNANE], n_niezn = _with_anchors(sections[SEP_NIEZNANE], _nieznane_key)
    rebuild_sheet(worksheet, sections, blank_rows={})
    return n_inne, n_niezn


def sync_kosztowe(worksheet, files_data):
    sections                  = read_all_sections(worksheet)
    new_rows, skipped, added  = apply_sync_logic(sections[SEP_KOSZTOWE], files_data)
    # Faktury 'cash' na koniec — kotwice status=3 nie sa sortowane, zostaja na pozycji
    anchors     = [(i, r) for i, r in enumerate(new_rows)
                   if str(r[2] if len(r) > 2 else "").strip() == "3"]
    non_anchors = [r for r in new_rows
                   if str(r[2] if len(r) > 2 else "").strip() != "3"]
    non_anchors.sort(key=lambda r: _is_cash(r[0] if r else ""))
    for pos, row in sorted(anchors, key=lambda x: x[0]):
        non_anchors.insert(min(pos, len(non_anchors)), row)
    sections[SEP_KOSZTOWE]    = non_anchors
    rebuild_sheet(worksheet, sections, blank_rows={})
    return skipped, added


def sync_sprzedaz(worksheet, tenants_data):
    sections                  = read_all_sections(worksheet)
    new_rows, skipped, added  = apply_sync_logic(
        sections[SEP_SPRZEDAZ], tenants_data, has_address=True, default_status="1"
    )
    sections[SEP_SPRZEDAZ]    = new_rows
    rebuild_sheet(worksheet, sections)
    return skipped, added


def check_sprzedaz_status(drive_service, credentials, spreadsheet_id, sheet_name):
    """Sprawdza stan faktur sprzedazy: plik zbiorczy na Drive vs suma/liczba w Sheets.
    Zwraca dict z kluczami: drive_filename, drive_szt, drive_kwota, sheet_szt, sheet_kwota."""
    # --- Drive: szukaj pliku Fs_najemcy_* w folderze MMYYYY Faktury sprzedazy ---
    sprzedaz_root = find_subfolder(drive_service, FOLDER_ID, "Faktury-sprzedazy")
    drive_filename = None
    drive_szt      = None
    drive_kwota    = None
    if sprzedaz_root:
        month_folder = find_subfolder(
            drive_service, sprzedaz_root["id"],
            f"{sheet_name} {FAKTURY_SPRZEDAZY_SUFFIX}"
        )
        if month_folder:
            query = (
                f"'{month_folder['id']}' in parents "
                "and mimeType='application/pdf' "
                "and name contains 'Fs_najemcy_' "
                "and trashed=false"
            )
            results = drive_service.files().list(
                q=query, fields="files(name)", orderBy="name"
            ).execute()
            files = results.get("files", [])
            if files:
                drive_filename = files[-1]["name"]   # najnowszy jeśli kilka
                # Wyciągnij szt i kwotę z nazwy: Fs_najemcy_042026_54szt_89305zl_07052026.pdf
                m = re.search(r"_(\d+)szt_(\d+)zl_", drive_filename)
                if m:
                    drive_szt   = int(m.group(1))
                    drive_kwota = int(m.group(2))

    # --- Sheets: suma kolumny B i liczba wierszy SEP_SPRZEDAZ ---
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return None
    rows = read_all_sections(worksheet)[SEP_SPRZEDAZ]
    sheet_szt   = 0
    sheet_kwota = 0.0
    for row in rows:
        if not row or not row[0]:
            continue
        b = row[1] if len(row) > 1 else None
        val = _parse_amount(b) or 0.0
        if val > 0:
            sheet_szt   += 1
            sheet_kwota += val

    return {
        "drive_filename": drive_filename,
        "drive_szt":      drive_szt,
        "drive_kwota":    drive_kwota,
        "sheet_szt":      sheet_szt,
        "sheet_kwota":    sheet_kwota,
    }


def diff_kosztowe(credentials, spreadsheet_id, sheet_name, drive_file_names):
    """Porownuje pliki na Drive z wierszami w sekcji KOSZTOWE arkusza.
    Zwraca (tylko_na_drive, tylko_w_sheets) — posortowane listy nazw."""
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(spreadsheet_id)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return None, None
    rows = read_all_sections(worksheet)[SEP_KOSZTOWE]
    sheet_names = set()
    for row in rows:
        if row and row[0]:
            sheet_names.add(str(row[0]).strip())
    drive_set = set(drive_file_names)
    only_drive  = sorted(drive_set  - sheet_names)
    only_sheets = sorted(sheet_names - drive_set)
    return only_drive, only_sheets


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
            has_pair = len(row) > 4 and str(row[4]).strip() != ""
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


def find_wyciag_file(service, subfolder_name):
    """Szuka pliku wyciag_MMYYYY.pdf w folderze Listy_operacji_abido."""
    filename = f"wyciag_{subfolder_name}.pdf"
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


def parse_wyciag_summary(pdf_bytes):
    """
    Parsuje wyciąg bankowy PDF.
    Sumy z nagłówka strony 1, liczniki z transakcji na stronach 2+.
    Zwraca (wplywy_sum, wyplywy_sum, wplywy_count, wyplywy_count)
    lub (None, None, None, None) gdy nie znaleziono.
    """
    _pat_in  = re.compile(r"Suma\s+wp[łl]yw[óo]w\s+([\d\s]+,\d{2})", re.IGNORECASE)
    _pat_out = re.compile(r"Suma\s+wyp[łl]yw[óo]w\s+(-?[\d\s]+,\d{2})", re.IGNORECASE)
    _pat_amt = re.compile(r"(-?[\d\s]{1,15},\d{2})\s+PLN")
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page1_text = pdf.pages[0].extract_text() or ""
            m_in  = _pat_in.search(page1_text)
            m_out = _pat_out.search(page1_text)
            wplywy_sum  = float(m_in.group(1).replace(" ", "").replace(",", "."))  if m_in  else None
            wyplywy_sum = float(m_out.group(1).replace(" ", "").replace(",", ".")) if m_out else None

            in_count = 0; out_count = 0
            for page in pdf.pages[1:]:
                text = page.extract_text() or ""
                for m in _pat_amt.finditer(text):
                    try:
                        val = float(m.group(1).replace(" ", "").replace(",", "."))
                        if val > 0:
                            in_count += 1
                        elif val < 0:
                            out_count += 1
                    except ValueError:
                        pass

        return wplywy_sum, wyplywy_sum, in_count, out_count
    except Exception:
        return None, None, None, None


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
    """Parsuje kwote z komorki arkusza (obsluguje przecinek, minus, suffix zl)."""
    try:
        return abs(float(re.sub(r"[^\d,.]", "", str(s)).replace(",", ".")))
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


def _is_cash(filename):
    return "cash" in str(filename).lower()


def assign_klucz_ksiegowy(section, tx, amount_b_str, filename=""):
    """Wyznacza Klucz_Ksiegowy na podstawie sekcji i transakcji."""
    # Faktura gotowkowa (nazwa zawiera 'cash') → zawsze kos_rk_kw
    if section == SEP_KOSZTOWE and _is_cash(filename):
        return "kos_rk_kw"

    if tx is None:
        if section == SEP_SPRZEDAZ:
            return "prz_naj_rk_kp"
        try:
            val = float(str(amount_b_str).replace(",", "."))
        except (ValueError, TypeError):
            val = -1
        if val > 0:
            return "prz_pr_in" if section == SEP_KOSZTOWE else "kos_rk_kp"
        return "kos_rk_kw"

    kwota = tx["kwota"]
    if section == SEP_KOSZTOWE:
        if kwota > 0:
            return "prz_pr_in"
        return "kos_med_pr_out" if _is_media(tx) else "kos_pr_out"
    if section == SEP_SPRZEDAZ:
        return "prz_naj_pr_in" if kwota > 0 else "prz_naj_rk_kp"
    if section == SEP_WLASC:
        return ("kos_wla_med_pr_out" if _is_media(tx) else "kos_wla_pr_out") if kwota < 0 else "kos_wla_pr_in"
    return "nieznany_out" if kwota < 0 else "nieznany_in"


def _frozen_tx_pre_used(sections, transactions, statuses=("2",)):
    """
    Zwraca zbior indeksow transakcji z wyciagu juz uzytych przez wiersze o podanych statusach.
    Domyslnie status=2. Moze tez obslugiwac status=3 (beton).
    Status=9 NIE blokuje TX — TX wraca do puli i moze sie ponownie sparowac.
    Sygnatura: (kwota, data_ks, nr_rachunku, tytul[:20]) — wystarczajaco unikalna.
    """
    frozen_sigs = set()
    for sep in SECTION_ORDER:
        for row in sections[sep]:
            if str(row[2] if len(row) > 2 else "").strip() not in statuses:
                continue
            if len(row) <= 6 or not str(row[6]).strip():
                continue  # brak daty_ks = nigdy nie bylo sparowania
            try:
                kwota = round(float(
                    re.sub(r"[^\d,.\-]", "", str(row[5])).replace(",", ".")
                ), 2)
            except (ValueError, TypeError):
                continue
            sig = (
                kwota,
                _norm_date(row[6]),
                str(row[11]).strip() if len(row) > 11 else "",
                str(row[7]).strip()[:20] if len(row) > 7 else "",
            )
            frozen_sigs.add(sig)

    if not frozen_sigs:
        return set()

    pre_used = set()
    for i, tx in enumerate(transactions):
        sig = (
            round(tx["kwota"], 2),
            _norm_date(tx["data_ks"]),
            str(tx["nr_rachunku"]).strip(),
            str(tx["tytul"]).strip()[:20],
        )
        if sig in frozen_sigs:
            pre_used.add(i)
    return pre_used


def pair_transactions(candidates, transactions, pre_used=None, blocked=None, multi_eligible=None):
    """
    Paruje kandydatow (wiersze arkusza) z transakcjami bankowymi w 6 przebiegach.
    candidates: lista (idx, name, amount_float, direction)
                direction: 1 = wpływ (sprzedaz), -1 = wydatek (kosztowe, wlasciciele)
    transactions: lista slownikow transakcji
    pre_used: zbior indeksow transakcji juz uzytych (status=2)
    blocked:  slownik {cand_idx: tx_idx} — kandydat NIE moze wrocic do starego TX (status=9)
    multi_eligible: zbior flat_idx dozwolonych w przejsciu 6 (tylko SPRZEDAZ i WLASC)
    Zwraca: (matched, name_only, used_tx)
      matched:   {cand_idx: tx_idx} — wszystkie udane parowania
      name_only: set(cand_idx)      — sparowane po nazwie bez zgodnosci kwoty (fioletowe)
      used_tx:   set(tx_idx)
    """
    matched          = {}
    name_only        = set()
    name_amt_matched = set()   # przebiegi 1 i 2: nazwisko/imie + kwota
    extras           = {}   # {flat_idx: [tx_idx, ...]} — dodatkowe TX do sub-wierszy (multi-parowanie)
    used_tx          = set(pre_used) if pre_used else set()
    blocked          = blocked or {}

    def ok(cand_idx, tx_idx):
        """True jesli TX nie jest zablokowana dla tego kandydata."""
        return blocked.get(cand_idx) != tx_idx

    def free_by_amount(cand_idx, amount, direction):
        return [i for i, tx in enumerate(transactions)
                if i not in used_tx
                and ok(cand_idx, i)
                and _parse_amount(tx["kwota"]) == amount
                and tx["kwota"] * direction > 0]

    def free_by_direction(cand_idx, direction):
        return [i for i, tx in enumerate(transactions)
                if i not in used_tx
                and ok(cand_idx, i)
                and tx["kwota"] * direction > 0]

    def assign(cand_idx, tx_idx):
        matched[cand_idx] = tx_idx
        used_tx.add(tx_idx)

    # Przebieg 1: nazwisko (ostatni token) + kwota
    for idx, name, amount, direction in candidates:
        if idx in matched or amount is None:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        hits = [i for i in free_by_amount(idx, amount, direction) if _search_token(transactions[i], tokens[-1])]
        if hits:
            assign(idx, hits[0])
            name_amt_matched.add(idx)

    # Przebieg 2: imie (pierwszy token) + kwota
    for idx, name, amount, direction in candidates:
        if idx in matched or amount is None:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        hits = [i for i in free_by_amount(idx, amount, direction) if _search_token(transactions[i], tokens[0])]
        if hits:
            assign(idx, hits[0])
            name_amt_matched.add(idx)

    # Przebieg 3: nazwisko (bez kwoty) — WSZYSTKIE pasujace TX → multi-parowanie (fioletowe)
    # Multi-parowanie (extras + used_tx) tylko dla SPRZEDAZ/WLASC — KOSZTOWE bierze jeden TX.
    for idx, name, amount, direction in candidates:
        if idx in matched:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        hits = [i for i in free_by_direction(idx, direction) if _search_token(transactions[i], tokens[-1])]
        if hits:
            assign(idx, hits[0])
            name_only.add(idx)
            if len(hits) > 1 and (multi_eligible is None or idx in multi_eligible):
                extras[idx] = hits[1:]
                for tx_i in hits[1:]:
                    used_tx.add(tx_i)

    # Przebieg 4: imie (pierwszy token, bez kwoty) — WSZYSTKIE pasujace TX → multi-parowanie (fioletowe)
    # Multi-parowanie (extras + used_tx) tylko dla SPRZEDAZ/WLASC — KOSZTOWE bierze jeden TX.
    for idx, name, amount, direction in candidates:
        if idx in matched:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        hits = [i for i in free_by_direction(idx, direction) if _search_token(transactions[i], tokens[0])]
        if hits:
            assign(idx, hits[0])
            name_only.add(idx)
            if len(hits) > 1 and (multi_eligible is None or idx in multi_eligible):
                extras[idx] = hits[1:]
                for tx_i in hits[1:]:
                    used_tx.add(tx_i)

    # Przebieg 5: sama kwota (ostatnia szansa, dokladnie 1 tx)
    for idx, name, amount, direction in candidates:
        if idx in matched or amount is None:
            continue
        pool = free_by_amount(idx, amount, direction)
        if len(pool) == 1:
            assign(idx, pool[0])

    # Przebieg 6: dla juz sparowanych kandydatow z SPRZEDAZ/WLASC — poszukaj dodatkowych
    # TX z tym samym nazwiskiem. Trafi tam np. platnosc za media od tego samego najemcy.
    # KOSZTOWE wykluczone — inaczej faktura "Mafika" pochlanialaby inne TX Mafiki.
    for idx, name, amount, direction in candidates:
        if idx not in matched:
            continue
        if multi_eligible is not None and idx not in multi_eligible:
            continue
        tokens = _extract_name_tokens(name)
        if not tokens:
            continue
        extra_hits = [
            i for i in free_by_direction(idx, direction)
            if _search_token(transactions[i], tokens[-1])
            or (len(tokens) > 1 and _search_token(transactions[i], tokens[0]))
        ]
        if extra_hits:
            prev = list(extras.get(idx, []))
            extras[idx] = prev + extra_hits
            for tx_i in extra_hits:
                used_tx.add(tx_i)

    return matched, name_only, name_amt_matched, extras, used_tx


def _build_paired_row(existing_row, tx, klucz, uwagi=""):
    """Uzupelnia wiersz arkusza danymi z transakcji bankowej."""
    row = list(existing_row) + [""] * max(0, 14 - len(existing_row))
    row[3]  = klucz
    row[4]  = tx["kontrahent"].split("|")[0]
    row[5]  = tx["kwota"]
    row[6]  = tx["data_ks"]
    row[7]  = tx["tytul"][:100]
    row[8]  = tx["data_op"]
    row[9]  = tx["rodzaj"]
    row[10] = tx["waluta"]
    row[11] = tx["nr_rachunku"]
    row[12] = _extract_name_from_tx(tx)
    row[13] = uwagi
    return row


def _build_sub_row(tx, klucz):
    """Sub-wiersz: puste A i B, status=1 (do zatwierdzenia), dane TX w H-P."""
    return [
        "", "", "1",
        klucz,
        tx["kontrahent"].split("|")[0],
        tx["kwota"],
        tx["data_ks"],
        tx["tytul"][:100],
        tx["data_op"],
        tx["rodzaj"],
        tx["waluta"],
        tx["nr_rachunku"],
        _extract_name_from_tx(tx),
        "",
    ]


def _build_unmatched_row(tx):
    """Buduje wiersz dla niesparowanej transakcji z wyciagu (A i B puste)."""
    kontrahent_low = str(tx.get("kontrahent", "")).lower()
    combined       = kontrahent_low + " " + str(tx.get("tytul", "")).lower()
    tytul_low = str(tx.get("tytul", "")).lower()
    if "bankomat" in kontrahent_low:
        klucz = "roz_bankomat_rk_kw" if tx["kwota"] > 0 else "roz_bankomat_rk_kp"
    elif ("urz" in combined and "skarbowy" in combined
          or re.search(r'(?<![a-z])(?:cit|pit)(?![a-z])', tytul_low)):
        klucz = "kos_pod_pr_out"
    elif "zakład ubezpieczeń" in combined or "zus" in combined:
        klucz = "kos_zus_pr_out"
    else:
        klucz = "nieznany_out" if tx["kwota"] < 0 else "nieznany_in"
    return [
        "", "", "",   # A=nazwa, B=kwota, C=status
        klucz,
        tx["kontrahent"].split("|")[0],
        tx["kwota"],
        tx["data_ks"],
        tx["tytul"][:100],
        tx["data_op"],
        tx["rodzaj"],
        tx["waluta"],
        tx["nr_rachunku"],
        _extract_name_from_tx(tx),
        "",                      # uwagi
    ]


_SEP_LABELS = {
    SEP_KOSZTOWE: "Faktury kosztowe",
    SEP_SPRZEDAZ: "Faktury sprzedazy najemcom",
    SEP_WLASC:    "Wlasciciele i spoldzielnie",
    SEP_INNE_RK:  "Inne raporty kasowe",
    SEP_NIEZNANE: "Nieznane / niesparowane",
}


def add_section_summary(worksheet, service=None, subfolder_name=None):
    """Dodaje tabelkę podsumowania na dole arkusza (trzy sub-tabele).
    Parowanie i sortowanie usuwają ją automatycznie (rebuild_sheet → clear()).
    """
    sections = read_all_sections(worksheet)

    def _parse_b(row):
        try:
            return float(re.sub(r"[^\d,.\-]", "", str(row[1] if len(row) > 1 else "")).replace(",", "."))
        except (ValueError, TypeError):
            return 0.0

    def _parse_f(row):
        try:
            return float(re.sub(r"[^\d,.\-]", "", str(row[5] if len(row) > 5 else "")).replace(",", "."))
        except (ValueError, TypeError):
            return 0.0

    # ── TABELA 1: KOSZTOWE, SPRZEDAZ, WLASC ──────────────────────────────────
    # Kolumny: Segment | Ilość | Suma kol.B | Suma wyciąg | Suma RK | Bilans
    T1_SEPS = [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC]
    t1_stats = {}
    for sep in T1_SEPS:
        count = 0; sum_b = 0.0; sum_f = 0.0; sum_rk = 0.0
        for row in sections[sep]:
            col_a = str(row[0]).strip() if row else ""
            klucz = str(row[3]).strip() if len(row) > 3 else ""
            if col_a:
                count += 1
                b = _parse_b(row)
                sum_b += b
                if "_rk_" in klucz:
                    sum_rk += b
            if len(row) > 5 and str(row[5]).strip():
                sum_f += _parse_f(row)
        t1_stats[sep] = (count, round(sum_b, 2), round(sum_f, 2), round(sum_rk, 2),
                         round(sum_b - sum_f - sum_rk, 2))

    t1_rc  = sum(s[0] for s in t1_stats.values())
    t1_rb  = round(sum(s[1] for s in t1_stats.values()), 2)
    t1_rf  = round(sum(s[2] for s in t1_stats.values()), 2)
    t1_rrk = round(sum(s[3] for s in t1_stats.values()), 2)
    t1_rbil = round(t1_rb - t1_rf - t1_rrk, 2)

    # ── TABELA 2: INNE_RK, NIEZNANE ──────────────────────────────────────────
    # Kolumny: Segment | Ilość | kos_ sum | prz_ sum | roz_ sum
    # INNE_RK: col B per prefiks klucza
    # NIEZNANE: wyciag_Kwota (ujemne→kos_, dodatnie→prz_)
    t2_stats = {}

    count = 0; sk = 0.0; sp = 0.0; sr = 0.0
    for row in sections[SEP_INNE_RK]:
        col_a = str(row[0]).strip() if row else ""
        klucz = str(row[3]).strip() if len(row) > 3 else ""
        if col_a:
            count += 1
            b = _parse_b(row)
            if klucz.startswith("kos_"):
                sk += b
            elif klucz.startswith("prz_"):
                sp += b
            elif klucz.startswith("roz_"):
                sr += b
    t2_stats[SEP_INNE_RK] = (count, round(sk, 2), round(sp, 2), round(sr, 2))

    count = 0; sk = 0.0; sp = 0.0; sr = 0.0
    for row in sections[SEP_NIEZNANE]:
        col_e = str(row[4]).strip() if len(row) > 4 else ""
        klucz = str(row[3]).strip() if len(row) > 3 else ""
        if col_e:
            count += 1
        f = _parse_f(row)
        if klucz.startswith("kos_"):
            sk += f
        elif klucz.startswith("prz_"):
            sp += f
        elif klucz.startswith("roz_"):
            sr += f
        elif f < 0:
            sk += f
        elif f > 0:
            sp += f
    t2_stats[SEP_NIEZNANE] = (count, round(sk, 2), round(sp, 2), round(sr, 2))

    t2_rc  = sum(s[0] for s in t2_stats.values())
    t2_rk  = round(sum(s[1] for s in t2_stats.values()), 2)
    t2_rp  = round(sum(s[2] for s in t2_stats.values()), 2)
    t2_rr  = round(sum(s[3] for s in t2_stats.values()), 2)

    # ── MATRYCA: kos_/prz_/roz_ × wszystkie 5 sekcji ────────────────────────
    # Kolumny: Segment | KOSZTOWE | SPRZEDAZ | WLASC | INNE_RK | NIEZNANE | Bilans
    ALL_SEPS = [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC, SEP_INNE_RK, SEP_NIEZNANE]
    matrix = {"kos_": {s: 0.0 for s in ALL_SEPS},
              "prz_": {s: 0.0 for s in ALL_SEPS},
              "roz_": {s: 0.0 for s in ALL_SEPS}}
    for sep in ALL_SEPS:
        for row in sections[sep]:
            klucz = str(row[3]).strip() if len(row) > 3 else ""
            if sep == SEP_NIEZNANE:
                f = _parse_f(row)
                if klucz.startswith("kos_"):
                    matrix["kos_"][sep] += f
                elif klucz.startswith("prz_"):
                    matrix["prz_"][sep] += f
                elif klucz.startswith("roz_"):
                    matrix["roz_"][sep] += f
                elif f < 0:
                    matrix["kos_"][sep] += f
                elif f > 0:
                    matrix["prz_"][sep] += f
            else:
                col_a = str(row[0]).strip() if row else ""
                if col_a:
                    b = _parse_b(row)
                    for pfx in ("kos_", "prz_", "roz_"):
                        if klucz.startswith(pfx):
                            matrix[pfx][sep] += b
                            break
    for pfx in matrix:
        for sep in ALL_SEPS:
            matrix[pfx][sep] = round(matrix[pfx][sep], 2)
    mat_bil = {pfx: round(sum(matrix[pfx].values()), 2) for pfx in matrix}

    # ── STATUS 0/1 ───────────────────────────────────────────────────────────
    status0_count = 0; status1_count = 0
    status_null_count = 0; klucz_null_count = 0; klucz_nieznane_count = 0
    for sep, sec_rows in sections.items():
        for row in sec_rows:
            status = str(row[2] if len(row) > 2 else "").strip()
            klucz  = str(row[3] if len(row) > 3 else "").strip()
            # Klucz Nieznane: liczymy wszędzie gdzie klucz = nieznany_*, niezależnie od col A
            if klucz.startswith("nieznany_"):
                klucz_nieznane_count += 1
                continue
            if sep == SEP_NIEZNANE:
                has_content = (str(row[5]).strip() if len(row) > 5
                               else str(row[4]).strip() if len(row) > 4 else "")
            else:
                has_content = str(row[0]).strip() if row else ""
            if not has_content:
                continue
            if status == "0":
                status0_count += 1
            elif status == "1":
                status1_count += 1
            elif not status:
                status_null_count += 1
            if not klucz:
                klucz_null_count += 1

    # ── WYCIĄG W ARKUSZU ─────────────────────────────────────────────────────
    wyciag_count = 0; wyciag_sum = 0.0
    for sec_rows in sections.values():
        for row in sec_rows:
            if len(row) > 5 and str(row[5]).strip():
                wyciag_count += 1
                wyciag_sum += _parse_f(row)
    wyciag_sum = round(wyciag_sum, 2)

    # ── DANE Z PLIKU LISTA_OPERACJI (opcjonalnie) ─────────────────────────────
    bank_tx_count = None; bank_tx_sum = None; bank_diag = None
    bank_tx_in_count = None; bank_tx_in_sum = None
    bank_tx_out_count = None; bank_tx_out_sum = None
    wyciag_in = None; wyciag_out = None; wyciag_found = False
    wyciag_in_count = None; wyciag_out_count = None
    if service and subfolder_name:
        bank_file = find_bank_file(service, subfolder_name)
        if bank_file:
            xls_bytes = download_pdf(service, bank_file["id"])
            transactions = parse_bank_statement(xls_bytes)
            bank_tx_count     = len(transactions)
            bank_tx_sum       = round(sum(t["kwota"] for t in transactions), 2)
            bank_tx_in_count  = sum(1 for t in transactions if t["kwota"] > 0)
            bank_tx_in_sum    = round(sum(t["kwota"] for t in transactions if t["kwota"] > 0), 2)
            bank_tx_out_count = sum(1 for t in transactions if t["kwota"] < 0)
            bank_tx_out_sum   = round(sum(t["kwota"] for t in transactions if t["kwota"] < 0), 2)

            def _tx_sig_local(tx):
                return (round(float(tx["kwota"]), 2), _norm_date(tx["data_ks"]),
                        str(tx["nr_rachunku"]).strip(), str(tx["tytul"]).strip()[:30])

            sheet_sig_counter = Counter()
            for _sec_rows in sections.values():
                for _r in _sec_rows:
                    if len(_r) > 4 and str(_r[4]).strip():
                        try:
                            _kw = round(float(re.sub(r"[^\d,.\-]", "", str(_r[5])).replace(",", ".")), 2) \
                                  if len(_r) > 5 and str(_r[5]).strip() else 0.0
                        except (ValueError, TypeError):
                            _kw = 0.0
                        _sig = (_kw,
                                _norm_date(_r[6]) if len(_r) > 6 else "",
                                str(_r[11]).strip() if len(_r) > 11 else "",
                                str(_r[7]).strip()[:30] if len(_r) > 7 else "")
                        sheet_sig_counter[_sig] += 1

            file_sig_counter = Counter(_tx_sig_local(tx) for tx in transactions)
            missing_sigs = file_sig_counter - sheet_sig_counter
            dupe_sigs    = sheet_sig_counter - file_sig_counter
            missing_txs = []; duplicate_txs = []
            used_m = Counter(); used_d = Counter()
            for tx in transactions:
                sig = _tx_sig_local(tx)
                if missing_sigs[sig] > used_m[sig]:
                    missing_txs.append(tx); used_m[sig] += 1
                elif dupe_sigs[sig] > used_d[sig]:
                    duplicate_txs.append(tx); used_d[sig] += 1
            bank_diag = {"missing": missing_txs, "duplicates": duplicate_txs}

        wyciag_file = find_wyciag_file(service, subfolder_name)
        if wyciag_file:
            wyciag_found = True
            wyciag_bytes = download_pdf(service, wyciag_file["id"])
            wyciag_in, wyciag_out, wyciag_in_count, wyciag_out_count = parse_wyciag_summary(wyciag_bytes)
        else:
            wyciag_found = False

    # ── POZYCJA STARTOWA ─────────────────────────────────────────────────────
    all_vals = _api(worksheet.get_all_values)
    old_summary_row = None
    # Nowy format: szukaj "Status 0" jako kotwicy
    for _i, _r in enumerate(all_vals):
        if str(_r[0]).strip() == "Status 0":
            old_summary_row = _i + 1
            break
    # Fallback: stary format zaczyna się od "Segment" / "pozycji"
    if not old_summary_row:
        for _i, _r in enumerate(all_vals):
            if (str(_r[0]).strip() == "Segment"
                    and len(_r) > 1 and "pozycji" in str(_r[1]).lower()):
                old_summary_row = _i + 1
                break

    if old_summary_row:
        _blank_rows = [[""] * 7] * (len(all_vals) - old_summary_row + 1)
        _api(worksheet.update, f"A{old_summary_row}", _blank_rows,
             value_input_option="USER_ENTERED")
        start = old_summary_row
    else:
        last_row = len(all_vals)
        while last_row > 0 and not any(c for c in all_vals[last_row - 1]):
            last_row -= 1
        start = last_row + 3

    # ── BUDUJ WIERSZE ────────────────────────────────────────────────────────
    E = ""
    rows = []

    # Status 0/1/Null + Klucz Null/Nieznane
    rows.append(["Status 0",        status0_count,        E, E, E, E, E])
    rows.append(["Status 1",        status1_count,        E, E, E, E, E])
    rows.append(["Status Null",     status_null_count,    E, E, E, E, E])
    rows.append(["Klucz Null",      klucz_null_count,     E, E, E, E, E])
    rows.append(["Klucz Nieznane",  klucz_nieznane_count, E, E, E, E, E])
    rows.append([E] * 7)

    # Wyciąg kontrola
    rows.append(["wyciag_Kwota w arkuszu", wyciag_count, wyciag_sum, E, E, E, E])
    if bank_tx_count is not None:
        rows.append(["Lista operacji xlsx \u2014 liczba TX", bank_tx_count, bank_tx_sum, E, E, E, E])
        rows.append(["Lista operacji wp\u0142ywy",  bank_tx_in_count,  bank_tx_in_sum,  E, E, E, E])
        rows.append(["Lista operacji wyp\u0142ywy", bank_tx_out_count, bank_tx_out_sum, E, E, E, E])
    if service and subfolder_name:
        if wyciag_found:
            rows.append(["Wyci\u0105g bankowy wp\u0142ywy",
                         wyciag_in_count  if wyciag_in_count  is not None else E,
                         wyciag_in        if wyciag_in        is not None else E,
                         E, E, E, E])
            rows.append(["Wyci\u0105g bankowy wyp\u0142ywy",
                         wyciag_out_count if wyciag_out_count is not None else E,
                         wyciag_out       if wyciag_out       is not None else E,
                         E, E, E, E])
        else:
            rows.append(["Wyci\u0105g bankowy wp\u0142ywy",  E, "brak wyci\u0105gu w folderze", E, E, E, E])
            rows.append(["Wyci\u0105g bankowy wyp\u0142ywy", E, "brak wyci\u0105gu w folderze", E, E, E, E])
    rows.append([E] * 7)
    rows.append([E] * 7)

    # Tabela 1 — KOSZTOWE, SPRZEDAZ, WLASC
    rows.append(["Segment", "Ilo\u015b\u0107 pozycji", "Suma kol. B (faktura)",
                 "Suma wyci\u0105g_Kwota", "Suma RK", "Bilans", E])
    for sep in T1_SEPS:
        c, sb, sf, srk, bil = t1_stats[sep]
        rows.append([_SEP_LABELS[sep], c, sb, sf, srk, bil, E])
    rows.append(["RAZEM", t1_rc, t1_rb, t1_rf, t1_rrk, t1_rbil, E])
    rows.append([E] * 7)
    rows.append([E] * 7)

    # Tabela 2 — INNE_RK, NIEZNANE
    rows.append(["Segment", "Ilo\u015b\u0107 pozycji", "klucz_kos_kolB",
                 "klucz_prz_kolB", "klucz_roz_kolB", E, E])
    for sep in [SEP_INNE_RK, SEP_NIEZNANE]:
        c, sk_, sp_, sr_ = t2_stats[sep]
        rows.append([_SEP_LABELS[sep], c, sk_, sp_, sr_, E, E])
    rows.append(["RAZEM", t2_rc, t2_rk, t2_rp, t2_rr, E, E])
    rows.append([E] * 7)
    rows.append([E] * 7)

    # Matryca — kos_/prz_/roz_ × segmenty
    seg_labels = [_SEP_LABELS[s] for s in ALL_SEPS]
    rows.append(["Segment"] + seg_labels + ["Bilans"])
    for pfx, label in [("kos_", "Koszty (kos_)"),
                       ("prz_", "Przychody (prz_)"),
                       ("roz_", "Rozrachunkowe (roz_)")]:
        rows.append([label] + [matrix[pfx][s] for s in ALL_SEPS] + [mat_bil[pfx]])
    rows.append([E] * 7)
    rows.append([E] * 7)

    # Bilans miesiąca (prz_ + kos_)
    miesiac_label = f"Bilans miesi\u0105ca prz-kos {subfolder_name}" if subfolder_name else "Bilans miesi\u0105ca prz-kos"
    miesiac_bil   = round(mat_bil["prz_"] + mat_bil["kos_"], 2)
    rows.append([miesiac_label, miesiac_bil, E, E, E, E, E])

    # Timestamp generowania
    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    rows.append([E] * 7)
    rows.append([f"Wygenerowano: {ts}", E, E, E, E, E, E])

    _api(worksheet.update, f"A{start}", rows, value_input_option="USER_ENTERED")

    # ── FORMATOWANIE ─────────────────────────────────────────────────────────
    _green  = {"red": 0.20, "green": 0.78, "blue": 0.35}
    _orange = {"red": 1.0,  "green": 0.76, "blue": 0.30}
    _white  = {"red": 1.0,  "green": 1.0,  "blue": 1.0}
    _bold_white = {"bold": True, "foregroundColor": _white}
    _hdr_style = {"backgroundColor": {"red": 0.85, "green": 0.85, "blue": 0.85},
                  "textFormat": {"bold": True}}
    _razem_style = {"backgroundColor": {"red": 0.2, "green": 0.2, "blue": 0.2},
                    "textFormat": _bold_white}
    _bilans_style = {"backgroundColor": {"red": 0.25, "green": 0.25, "blue": 0.50},
                     "textFormat": _bold_white}
    _wyciag_style = {"backgroundColor": {"red": 0.18, "green": 0.45, "blue": 0.75},
                     "textFormat": _bold_white}

    row_fmts = []
    cur = start

    # Status 0/1/Null + Klucz Null/Nieznane
    _yellow = {"red": 0.99, "green": 0.90, "blue": 0.40}
    _bold_dark = {"bold": True, "foregroundColor": {"red": 0.1, "green": 0.1, "blue": 0.1}}
    row_fmts.append((cur,     {"backgroundColor": _green  if status0_count == 0        else _orange, "textFormat": _bold_white}))
    row_fmts.append((cur + 1, {"backgroundColor": _green  if status1_count == 0        else _orange, "textFormat": _bold_white}))
    row_fmts.append((cur + 2, {"backgroundColor": _yellow if status_null_count == 0    else _orange, "textFormat": _bold_dark}))
    row_fmts.append((cur + 3, {"backgroundColor": _yellow if klucz_null_count == 0     else _orange, "textFormat": _bold_dark}))
    row_fmts.append((cur + 4, {"backgroundColor": _yellow if klucz_nieznane_count == 0 else _orange, "textFormat": _bold_dark}))
    cur += 6  # 5 wierszy + 1 separator

    # Wyciąg kontrola
    row_fmts.append((cur, _wyciag_style))
    if bank_tx_count is not None:
        row_fmts.append((cur + 1, _wyciag_style))
        cur += 4  # 2 wyciąg + 2 separatory
    else:
        cur += 3  # 1 wyciąg + 2 separatory

    # Tabela 1
    row_fmts.append((cur, _hdr_style)); cur += 1
    for sep in T1_SEPS:
        row_fmts.append((cur, {"backgroundColor": SEP_COLORS[sep], "textFormat": _bold_white}))
        cur += 1
    row_fmts.append((cur, _razem_style)); cur += 3  # RAZEM + 2 separatory

    # Tabela 2
    row_fmts.append((cur, _hdr_style)); cur += 1
    for sep in [SEP_INNE_RK, SEP_NIEZNANE]:
        row_fmts.append((cur, {"backgroundColor": SEP_COLORS[sep], "textFormat": _bold_white}))
        cur += 1
    row_fmts.append((cur, _razem_style)); cur += 3  # RAZEM + 2 separatory

    # Matryca
    row_fmts.append((cur, _hdr_style)); cur += 1
    for _ in range(3):
        row_fmts.append((cur, _bilans_style)); cur += 1
    cur += 2  # 2 separatory

    # Bilans miesiąca
    _lime = {"red": 0.20, "green": 0.93, "blue": 0.20}
    _bold_black = {"bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}
    row_fmts.append((cur, {"backgroundColor": _lime, "textFormat": _bold_black}))
    cur += 2  # bilans + pusty separator

    # Timestamp
    _gray_txt = {"red": 0.45, "green": 0.45, "blue": 0.45}
    row_fmts.append((cur, {"textFormat": {"italic": True, "foregroundColor": _gray_txt}}))

    _batch_format_rows(worksheet, row_fmts)

    # Kolumna B w arkuszu ma format walutowy (to kolumna "Kwota brutto").
    # Resetujemy format liczbowy kol B w całym zakresie podsumowania,
    # żeby "Ilość pozycji" wyświetlała się jako gołe liczby, nie "27,00zł".
    _sum_rows = len(rows)
    _sheet_id = worksheet._properties["sheetId"]
    _api(worksheet.spreadsheet.batch_update, {"requests": [{
        "repeatCell": {
            "range": {
                "sheetId": _sheet_id,
                "startRowIndex": start - 1,
                "endRowIndex": start - 1 + _sum_rows,
                "startColumnIndex": 1,
                "endColumnIndex": 2,
            },
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0"}}},
            "fields": "userEnteredFormat.numberFormat",
        }
    }]})

    return bank_diag


def sync_parowanie(worksheet, transactions):
    """
    Paruje wiersze ze statusem 1 z transakcjami bankowymi.
    Status 2 — nietykalny (fizycznie wyjety z sekcji przed parowaniem, wracaja po).
    Status 0 — pomijany.
    Niesparowane transakcje bankowe laduja na dole w sekcji SEP_NIEZNANE.
    Wiersze sparowane tylko po nazwie (bez kwoty) dostaja kolor fioletowy.
    """
    sections = read_all_sections(worksheet)

    # Sygnatury TX już obecnych w INNE_RK — zapobiegają podwójnemu dodaniu bankomatów.
    # Counter zamiast set: wiele identycznych TX (np. 8x -200 zł z tego samego bankomatu
    # tego samego dnia) ma tę samą sygnaturę — Counter pozwala dodać tyle ile jest w pliku.
    _inne_rk_sigs = Counter()
    for _r in sections[SEP_INNE_RK]:
        if len(_r) > 6 and str(_r[6]).strip():
            try:
                _kw = round(float(re.sub(r"[^\d,.\-]", "", str(_r[5])).replace(",", ".")), 2)
            except (ValueError, TypeError):
                continue
            _inne_rk_sigs[(_kw, str(_r[6]).strip(), str(_r[11]).strip() if len(_r) > 11 else "")] += 1

    # Zachowaj wiersze status=3 z NIEZNANE zanim KROK 0 zmodyfikuje sekcje.
    # Sekcja NIEZNANE jest pozniej calkowicie nadpisywana — bez tego wiersze gina.
    _frozen3_nieznane = [
        (i, r) for i, r in enumerate(sections[SEP_NIEZNANE])
        if str(r[2] if len(r) > 2 else "").strip() == "3"
    ]

    # TX uzywane przez wiersze status=3 (we wszystkich sekcjach).
    # Musza byc w pre_used (blokada re-parowania) i w used_tx (blokada unmatched_rows).
    _status3_pre_used = _frozen_tx_pre_used(sections, transactions, statuses=("3",))

    # ── KROK 0: Snapshot status=2 ──────────────────────────────────────────────
    # Fizycznie wyjmij zamrozone wiersze ze wszystkich sekcji zanim cokolwiek
    # zostanie dotkniete. Zadna petla parowania ich nie zobaczy.
    # Sub-wiersze (A='') deduplikujemy po sygnaturze TX — zapobiega narastaniu
    # kopii z kolejnych uruchomien parowania.
    def _sub_row_sig(r):
        try:
            kw = round(float(re.sub(r"[^\d,.\-]", "", str(r[5])).replace(",", ".")), 2)
        except (ValueError, TypeError):
            kw = 0.0
        return (kw, str(r[6]).strip() if len(r) > 6 else "",
                str(r[11]).strip() if len(r) > 11 else "")

    frozen_backup = {}
    frozen_order  = {}   # {sep: ['f'|'a', ...]} — kolejnosc glownych wierszy dla SPRZEDAZ/WLASC
    for sep in SECTION_ORDER:
        if sep == SEP_INNE_RK:
            frozen_backup[sep] = []  # INNE_RK: sekcja nienaruszona, parowanie jej nie dotyka
            continue
        frozen, active = [], []
        order = [] if sep in (SEP_SPRZEDAZ, SEP_WLASC) else None
        seen_sub_sigs = set()
        last_main_frozen  = False   # ostatni glowny to status=2 (frozen)
        last_main_status3 = False   # ostatni glowny to status=3 (beton)
        for row in sections[sep]:
            if sep == SEP_NIEZNANE:
                # NIEZNANE nie ma sub-wierszy — stara prosta logika
                if len(row) > 2 and str(row[2]).strip() == "2":
                    frozen.append(row)
                else:
                    active.append(row)
                continue
            is_main = bool(row[0] if row else "")
            if is_main:
                status_c = str(row[2]).strip() if len(row) > 2 else ""
                if status_c == "2":
                    frozen.append(row)
                    last_main_frozen  = True
                    last_main_status3 = False
                    if order is not None:
                        order.append('f')
                else:
                    active.append(row)
                    last_main_frozen  = False
                    last_main_status3 = (status_c == "3")
                    if order is not None:
                        order.append('a')
            else:
                # Sub-wiersz (A='')
                sub_status = str(row[2]).strip() if len(row) > 2 else ""
                if sub_status == "3":
                    # Sam jest betonem — zawsze zachowaj w frozen niezaleznie od rodzica
                    frozen.append(row)
                elif last_main_frozen:
                    # Rodzic status=2 → zachowaj sub-wiersz w frozen, deduplikuj
                    sig = _sub_row_sig(row)
                    if sig not in seen_sub_sigs:
                        seen_sub_sigs.add(sig)
                        frozen.append(row)
                elif last_main_status3:
                    # Rodzic status=3 (beton) → zachowaj sub-wiersz w active bez zmian
                    active.append(row)
                # else: Rodzic aktywny → wyrzuc; pass 6 odtworzy sub-wiersz
        frozen_backup[sep] = frozen
        sections[sep] = active
        if order is not None:
            frozen_order[sep] = order

    # ── SPRZEDAZ B=0: auto-status=2 (miesiac juz oplacony z gory — "sprawdzony") ─
    for _i, _row in enumerate(sections[SEP_SPRZEDAZ]):
        if not (_row[0] if _row else ""):
            continue  # sub-wiersz
        if str(_row[2]).strip() not in ("1", "9"):
            continue  # juz zamrozony lub status=0 — nie ruszaj
        _amt = _parse_amount(_row[1] if len(_row) > 1 else "")
        if _amt is None or _amt == 0.0:
            _row = list(_row) + [""] * max(0, 3 - len(_row))
            _row[2] = "2"
            sections[SEP_SPRZEDAZ][_i] = _row

    # ── Kandydaci do parowania ─────────────────────────────────────────────────
    _DIRECTION = {SEP_KOSZTOWE: -1, SEP_SPRZEDAZ: 1, SEP_WLASC: -1}
    candidates = []   # (flat_idx, section, row_idx_in_section, name, amount, direction)
    for sep in [SEP_KOSZTOWE, SEP_SPRZEDAZ, SEP_WLASC]:
        for i, row in enumerate(sections[sep]):
            if not (row[0] if row else ""):
                continue   # sub-wiersz (puste A) — nie jest kandydatem
            if str(row[2]).strip() in ("1", "9"):
                amount = _parse_amount(row[1] if len(row) > 1 else "")
                # SPRZEDAZ z kwota=0: najemca zaplacil z gory — nieaktywny, pomijamy
                if sep == SEP_SPRZEDAZ and (amount is None or amount == 0.0):
                    continue
                # Refaktura/zwrot w KOSZTOWE: kwota faktury ze znakiem (+) → kierunek +1 (wpływ)
                # Uwaga: _parse_amount zwraca abs() — sprawdzamy znak z surowej wartosci
                try:
                    b_signed = float(re.sub(r"[^\d,.\-]", "", str(row[1] if len(row) > 1 else "")).replace(",", "."))
                except (ValueError, TypeError):
                    b_signed = -1
                direction = 1 if (sep == SEP_KOSZTOWE and b_signed > 0) else _DIRECTION[sep]
                candidates.append((len(candidates), sep, i, row[0], amount, direction))

    # TX juz uzyte przez zamrozone wiersze — skanujemy frozen_backup
    # statuses=("1","2"): sub-wiersze zamrozonych rodzicow maja teraz status=1 (nie status=2)
    real_pre_used = _frozen_tx_pre_used(frozen_backup, transactions, statuses=("1", "2"))

    # Bankomat — wyklucz z parowania (tylko "bankomat" w kontrahencie)
    # Nie uzywamy "blik" w tytule/rodzaju — platnosci BLIK przez PayU/OLX
    # to koszty, nie transakcje bankomatowe.
    bankomat_indices = {
        i for i, tx in enumerate(transactions)
        if "bankomat" in str(tx.get("kontrahent", "")).lower()
    }
    # Tylko bankomaty ktore NIE sa juz zamrozone — te wylaczamy z parowania
    bankomat_excluded = bankomat_indices - real_pre_used
    pre_used = real_pre_used | bankomat_excluded | _status3_pre_used

    # ── Status=9: zablokuj ponowne sparowanie ze starym TX ────────────────────
    def _tx_sig(tx):
        return (round(tx["kwota"], 2), str(tx["data_ks"]).strip(),
                str(tx["nr_rachunku"]).strip(), str(tx["tytul"]).strip()[:20])

    blocked = {}
    for flat_idx, sep, row_idx, name, amount, direction in candidates:
        row = sections[sep][row_idx]
        if str(row[2]).strip() != "9":
            continue
        if len(row) <= 6 or not str(row[6]).strip():
            continue
        try:
            old_kwota = round(float(re.sub(r"[^\d,.\-]", "", str(row[5])).replace(",", ".")), 2)
        except (ValueError, TypeError):
            continue
        old_sig = (old_kwota, str(row[6]).strip(),
                   str(row[11]).strip() if len(row) > 11 else "",
                   str(row[7]).strip()[:20] if len(row) > 7 else "")
        for i, tx in enumerate(transactions):
            if _tx_sig(tx) == old_sig:
                blocked[flat_idx] = i
                break

    flat = [(c[0], c[3], c[4], c[5]) for c in candidates]
    multi_eligible = {c[0] for c in candidates if c[1] in (SEP_SPRZEDAZ, SEP_WLASC)}
    matched, name_only, name_amt_matched, extras, used_tx = pair_transactions(flat, transactions, pre_used=pre_used, blocked=blocked, multi_eligible=multi_eligible)
    # Bankomat TX (nie-zamrozone) ida do SEP_INNE_RK — usun z used_tx
    used_tx -= bankomat_excluded
    # TX wierszy status=3 — juz sa w arkuszu, nie trafiaja do unmatched_rows
    used_tx |= _status3_pre_used

    unmatched_count = 0

    # ── Zapisz wyniki parowania do wierszy ────────────────────────────────────
    for flat_idx, sep, row_idx, name, amount, direction in candidates:
        row = sections[sep][row_idx]
        tx_idx = matched.get(flat_idx)
        if tx_idx is not None:
            tx = transactions[tx_idx]
            klucz = assign_klucz_ksiegowy(sep, tx, row[1] if len(row) > 1 else "", row[0] if row else "")
            r = _build_paired_row(row, tx, klucz)
            if str(r[2]).strip() == "9":
                r[2] = "1"
            tx_amount = abs(round(tx["kwota"], 2))
            inv_amount = _parse_amount(row[1] if len(row) > 1 else "")
            if flat_idx in name_only or (inv_amount is not None and round(inv_amount, 2) != tx_amount):
                r[13] = _PURPLE_MARKER
            elif flat_idx in name_amt_matched and sep in (SEP_SPRZEDAZ, SEP_WLASC):
                r[2] = "2"
            sections[sep][row_idx] = r
        else:
            klucz = assign_klucz_ksiegowy(sep, None, row[1] if len(row) > 1 else "", row[0] if row else "")
            r = list(row) + [""] * max(0, 14 - len(row))
            if str(r[2]).strip() == "9":
                r[2] = "1"
            r[3] = klucz
            for col in range(4, 13):
                r[col] = ""
            # rk_kp / rk_kw = platnosc gotowkowa — brak TX bankowej to norma, nie "brak pary"
            # kolor zielony/rozowy nadaje rebuild_sheet; orange tu byloby mylace
            if "_rk_kp" not in klucz and "_rk_kw" not in klucz:
                r[13] = _ORANGE_MARKER
            sections[sep][row_idx] = r
            unmatched_count += 1

    # ── Sub-wiersze dla multi-parowan (extras) ────────────────────────────────
    if extras:
        flat_to_pos = {c[0]: (c[1], c[2]) for c in candidates}
        for sep in [SEP_SPRZEDAZ, SEP_WLASC]:
            inserts = []
            for flat_idx, extra_tx_idxs in extras.items():
                pos_sep, pos_row_idx = flat_to_pos[flat_idx]
                if pos_sep != sep:
                    continue
                sub_rows_list = []
                for tx_i in extra_tx_idxs:
                    tx = transactions[tx_i]
                    klucz = assign_klucz_ksiegowy(sep, tx, "", "")
                    sub_rows_list.append(_build_sub_row(tx, klucz))
                if sub_rows_list:
                    inserts.append((pos_row_idx, sub_rows_list))
            inserts.sort(key=lambda x: x[0], reverse=True)
            for row_idx, sub_rows_list in inserts:
                for i, sr in enumerate(sub_rows_list):
                    sections[sep].insert(row_idx + 1 + i, sr)

    # ── KROK N: Przywroc zamrozone wiersze do sekcji ──────────────────────────
    # Wiersze status=2 SA nietykalane — nie dostaja nowych sub-wierszy.
    # KOSZTOWE: zamrozone przed aktywnymi (kolejnosc nieistotna dla uzytkownika).
    # SPRZEDAZ / WLASC: odtwarzamy ORYGINALNA kolejnosc wierszy — zakazane sortowanie.
    def _row_blocks(rows):
        """Dzieli liste wierszy na bloki: kazdy blok zaczyna sie od glownego wiersza (col A)."""
        blocks, cur = [], None
        for row in rows:
            if row[0] if row else "":   # glowny wiersz
                if cur is not None:
                    blocks.append(cur)
                cur = [row]
            else:
                if cur is None:
                    cur = []            # osierocone sub-wiersze na poczatku
                cur.append(row)
        if cur is not None:
            blocks.append(cur)
        return blocks

    sections[SEP_KOSZTOWE] = list(frozen_backup[SEP_KOSZTOWE]) + sections[SEP_KOSZTOWE]
    for sep in (SEP_SPRZEDAZ, SEP_WLASC):
        order  = frozen_order.get(sep, [])
        if not order:
            # Brak sledzonej kolejnosci — fallback: zamrozone przed aktywnymi
            sections[sep] = list(frozen_backup[sep]) + sections[sep]
            continue
        f_blocks = _row_blocks(frozen_backup[sep])
        a_blocks = _row_blocks(sections[sep])
        fi = ai = 0
        merged = []
        for kind in order:
            if kind == 'f' and fi < len(f_blocks):
                merged.extend(f_blocks[fi]); fi += 1
            elif kind == 'a' and ai < len(a_blocks):
                merged.extend(a_blocks[ai]); ai += 1
        while fi < len(f_blocks): merged.extend(f_blocks[fi]); fi += 1
        while ai < len(a_blocks): merged.extend(a_blocks[ai]); ai += 1
        sections[sep] = merged

    # ── SEP_NIEZNANE: zamrozone z backupu + niesparowane ─────────────────────
    # Deduplikuj po sygnaturze TX (na wypadek gdyby poprzedni bug zostawil kopie)
    _seen_frozen_sigs = set()
    frozen_nieznane = []
    for r in frozen_backup.get(SEP_NIEZNANE, []):
        if len(r) > 6 and str(r[6]).strip():
            try:
                _kw = round(float(re.sub(r"[^\d,.\-]", "", str(r[5])).replace(",", ".")), 2)
            except (ValueError, TypeError):
                _kw = 0.0
            _sig = (_kw, str(r[6]).strip(), str(r[11]).strip() if len(r) > 11 else "")
            if _sig in _seen_frozen_sigs:
                continue
            _seen_frozen_sigs.add(_sig)
        frozen_nieznane.append(r)

    # Rozdziel niesparowane: bankomaty → SEP_INNE_RK, reszta → SEP_NIEZNANE
    _new_inne_rk = []
    unmatched_rows = []
    for i in range(len(transactions)):
        if i in used_tx:
            continue
        tx = transactions[i]
        kontrahent_low = str(tx.get("kontrahent", "")).lower()
        if "bankomat" in kontrahent_low:
            klucz = "roz_bankomat_rk_kw" if tx["kwota"] > 0 else "roz_bankomat_rk_kp"
            try:
                _kw_sig = round(float(tx["kwota"]), 2)
            except (ValueError, TypeError):
                _kw_sig = 0.0
            _sig = (_kw_sig, str(tx["data_ks"]).strip(), str(tx.get("nr_rachunku", "")).strip())
            if _inne_rk_sigs[_sig] > 0:
                _inne_rk_sigs[_sig] -= 1
                continue
            _new_inne_rk.append([
                tx["kontrahent"].split("|")[0],  # A
                tx["kwota"],                     # B
                "0",                             # C
                klucz,                           # D
                tx["kontrahent"].split("|")[0],  # E
                tx["kwota"],                     # F
                tx["data_ks"],                   # G
                tx["tytul"][:100],               # H
                tx["data_op"],                   # I
                tx["rodzaj"],                    # J
                tx["waluta"],                    # K
                tx["nr_rachunku"],               # L
                _extract_name_from_tx(tx),       # M
                "",                              # N
            ])
        else:
            unmatched_rows.append(_build_unmatched_row(tx))
    sections[SEP_INNE_RK] = sections[SEP_INNE_RK] + _new_inne_rk

    def _nieznane_sort_key(row):
        k = str(row[3]).strip().lower() if len(row) > 3 else ""
        if "depo" in k and "pr_in"  in k: return 0
        if "depo" in k and "pr_out" in k: return 1
        if "depo" in k and "kp"     in k: return 2
        if "depo" in k and "kw"     in k: return 3
        if "bankomat" in k and "kp" in k: return 4
        if "bankomat" in k and "kw" in k: return 5
        if "pod" in k:                    return 6
        if "zus" in k:                    return 7
        return 8

    sections[SEP_NIEZNANE] = sorted(
        frozen_nieznane + unmatched_rows,
        key=_nieznane_sort_key,
    )

    # Przywroc wiersze status=3 z NIEZNANE na ich oryginalne pozycje (zabetonowane)
    for orig_pos, row in sorted(_frozen3_nieznane, key=lambda x: x[0]):
        sections[SEP_NIEZNANE].insert(min(orig_pos, len(sections[SEP_NIEZNANE])), row)

    # Weryfikacja: policz i zsumuj wiersze z danymi wyciagu (col H niepusta)
    # Obejmuje: sparowane (w sekcjach) + niesparowane (SEP_NIEZNANE) + zamrozone (status=2)

    sheet_tx_count = 0
    sheet_tx_sum   = 0.0
    sheet_rows_with_bank = []   # (sig, row) — do diagnostyki
    for sec_rows in sections.values():
        for r in sec_rows:
            if len(r) > 4 and str(r[4]).strip():   # col E = wyciag_Kontrahent
                sheet_tx_count += 1
                try:
                    kwota_r = float(re.sub(r"[^\d,.\-]", "", str(r[5])).replace(",", ".")) if len(r) > 5 and str(r[5]).strip() else 0.0
                except (ValueError, TypeError):
                    kwota_r = 0.0
                if kwota_r:
                    sheet_tx_sum += kwota_r
                sig = (
                    round(kwota_r, 2),
                    _norm_date(r[6]) if len(r) > 6 else "",
                    str(r[11]).strip() if len(r) > 11 else "",
                    str(r[7]).strip()[:30] if len(r) > 7 else "",
                )
                sheet_rows_with_bank.append((sig, r))

    # Diagnostyka: ktore TX sa w pliku ale nie w arkuszu i odwrotnie
    def _tx_sig(tx):
        return (round(float(tx["kwota"]), 2), _norm_date(tx["data_ks"]),
                str(tx["nr_rachunku"]).strip(), str(tx["tytul"]).strip()[:30])

    file_sig_counter  = Counter(_tx_sig(tx) for tx in transactions)
    sheet_sig_counter = Counter(sig for sig, _ in sheet_rows_with_bank)

    missing_sigs = file_sig_counter - sheet_sig_counter
    extra_sigs   = sheet_sig_counter - file_sig_counter

    missing_txs = []
    used_m = Counter()
    for tx in transactions:
        sig = _tx_sig(tx)
        if missing_sigs[sig] > used_m[sig]:
            missing_txs.append(tx)
            used_m[sig] += 1

    extra_rows = []
    used_e = Counter()
    for sig, r in sheet_rows_with_bank:
        if extra_sigs[sig] > used_e[sig]:
            extra_rows.append(r)
            used_e[sig] += 1

    # ── FINAŁ: usuń nadmiarowe + dodaj brakujące (sygnatura-based) ───────────────
    # Chronimy: status=3 (beton) oraz wszystkie główne wiersze faktur (col A niepuste).
    # Reconcile może usuwać TYLKO sub-wiersze (col A puste) — nigdy główne faktury.
    _extra_to_remove = Counter(extra_sigs)
    for _sec in SECTION_ORDER:
        _new_rows = []
        for r in sections[_sec]:
            if len(r) > 4 and str(r[4]).strip():   # wiersz ma dane TX (col E)
                try:
                    _kw = round(float(
                        re.sub(r"[^\d,.\-]", "", str(r[5])).replace(",", ".")
                    ), 2) if len(r) > 5 and str(r[5]).strip() else 0.0
                except (ValueError, TypeError):
                    _kw = 0.0
                _rsig = (
                    _kw,
                    _norm_date(r[6]) if len(r) > 6 else "",
                    str(r[11]).strip() if len(r) > 11 else "",
                    str(r[7]).strip()[:30] if len(r) > 7 else "",
                )
                if _extra_to_remove[_rsig] > 0:
                    _is_main = bool(r[0] if r else "")
                    _status  = str(r[2] if len(r) > 2 else "").strip()
                    if _status == "3" or _is_main:
                        _new_rows.append(r)   # główna faktura lub beton — nietykalny
                    else:
                        _extra_to_remove[_rsig] -= 1   # usuń sub-wiersz
                    continue
            _new_rows.append(r)
        sections[_sec] = _new_rows

    # Brakujące TX → dodaj na dół NIEZNANE jako niesparowane
    for tx in missing_txs:
        sections[SEP_NIEZNANE].append(_build_unmatched_row(tx))

    diff_info = {"missing": missing_txs, "extra": extra_rows,
                 "reconciled": True}   # naprawiono automatycznie

    # Przelicz sheet_tx_count/sum po reconcile (zmodyfikowane sekcje)
    sheet_tx_count = 0
    sheet_tx_sum   = 0.0
    for _sec_rows in sections.values():
        for _r in _sec_rows:
            if len(_r) > 4 and str(_r[4]).strip():
                sheet_tx_count += 1
                try:
                    _v = float(re.sub(r"[^\d,.\-]", "", str(_r[5])).replace(",", ".")) if len(_r) > 5 and str(_r[5]).strip() else 0.0
                except (ValueError, TypeError):
                    _v = 0.0
                sheet_tx_sum += _v

    rebuild_sheet(worksheet, sections, blank_rows={})

    # Kolorowanie wierszy po markerach w kolumnie N (Uwagi, poz. 13)
    all_vals = worksheet.get_all_values()
    clear_updates = []
    purple_rows = []
    orange_rows = []
    for row_i, row_vals in enumerate(all_vals):
        if len(row_vals) <= 13:
            continue
        marker = row_vals[13]
        row_num = row_i + 1
        if marker == _PURPLE_MARKER:
            purple_rows.append(row_num)
            clear_updates.append({"range": f"N{row_num}", "values": [[""]]})
        elif marker == _ORANGE_MARKER:
            orange_rows.append(row_num)
            clear_updates.append({"range": f"N{row_num}", "values": [[""]]})
    color_fmts = (
        [(r, {"backgroundColor": _PURPLE_BG}) for r in purple_rows] +
        [(r, {"backgroundColor": _ORANGE_BG}) for r in orange_rows]
    )
    _batch_format_rows(worksheet, color_fmts)
    if clear_updates:
        _api(worksheet.batch_update, clear_updates)

    tx_total = len(transactions)
    tx_sum   = sum(tx["kwota"] for tx in transactions)
    return (
        len(matched), len(sections[SEP_NIEZNANE]),
        len(purple_rows), unmatched_count,
        tx_total, round(tx_sum, 2),
        sheet_tx_count, round(sheet_tx_sum, 2),
        diff_info,
    )


# ================================================================
# KP i KW — zakładka raportu kasowego
# ================================================================

_KP_KW_SHEET  = "Kp i Kw"
_KP_KW_MARKER = "=== {} ==="

_MONTHS_PL = {
    1: "STYCZEŃ", 2: "LUTY", 3: "MARZEC", 4: "KWIECIEŃ",
    5: "MAJ", 6: "CZERWIEC", 7: "LIPIEC", 8: "SIERPIEŃ",
    9: "WRZESIEŃ", 10: "PAŹDZIERNIK", 11: "LISTOPAD", 12: "GRUDZIEŃ",
}

_CAT_LABELS_KP = {0: "Przychody najemców", 1: "Kaucje wpłacone",
                  2: "Bankomat / BLIK", 3: "Inne KP"}
_CAT_LABELS_KW = {0: "Koszty gotówkowe", 1: "Zwroty kaucji",
                  2: "Bankomat / BLIK", 3: "Inne KW"}


def _month_label_pl(subfolder_name):
    try:
        m, y = int(subfolder_name[:2]), int(subfolder_name[2:])
        return f"{_MONTHS_PL.get(m, '?')} {y}"
    except Exception:
        return subfolder_name


def _kp_kw_opis(klucz, col_a):
    col_a = col_a.strip()
    if "prz_naj" in klucz and "_rk_kp" in klucz:
        return f"{col_a} — wynajem pokoju" if col_a else "Wynajem pokoju"
    if "roz_depo_part" in klucz and "_rk_kp" in klucz:
        return f"Zapłata części kaucji — {col_a}" if col_a else "Zapłata części kaucji"
    if "roz_depo_all" in klucz and "_rk_kp" in klucz:
        return f"Zapłata kaucji — {col_a}" if col_a else "Zapłata kaucji"
    if "roz_depo_part" in klucz and "_rk_kw" in klucz:
        return f"Zwrot części kaucji — {col_a}" if col_a else "Zwrot części kaucji"
    if "roz_depo_all" in klucz and "_rk_kw" in klucz:
        return f"Zwrot kaucji — {col_a}" if col_a else "Zwrot kaucji"
    if "bankomat" in klucz and "_rk_kp" in klucz:
        return "Wypłata z banku do kasy"
    if "bankomat" in klucz and "_rk_kw" in klucz:
        return "Wpłata do banku z kasy"
    return col_a if col_a else klucz


def _kp_kw_cat(klucz):
    if "prz_naj"  in klucz: return 0
    if "roz_depo" in klucz: return 1
    if "bankomat" in klucz: return 2
    return 3


def _extract_rk_entries(sections):
    """Zwraca (kp_entries, kw_entries) — listy (klucz, col_a, kwota, data)."""
    kp, kw = [], []
    for sec_rows in sections.values():
        for row in sec_rows:
            klucz = str(row[3]).strip() if len(row) > 3 else ""
            if "_rk_" not in klucz:
                continue
            col_a = str(row[0]).strip() if row else ""
            def _pq(s):
                try:
                    return abs(float(re.sub(r"[^\d,.\-]", "", s).replace(",", ".")))
                except Exception:
                    return 0.0
            kwota_b = _pq(str(row[1])) if len(row) > 1 and str(row[1]).strip() else 0.0
            kwota_i = _pq(str(row[5])) if len(row) > 5 and str(row[5]).strip() else 0.0
            kwota   = kwota_b if kwota_b else kwota_i
            data    = str(row[6]).strip() if len(row) > 6 else ""
            entry   = (klucz, col_a, kwota, data)
            if   "_rk_kp" in klucz: kp.append(entry)
            elif "_rk_kw" in klucz: kw.append(entry)
    return kp, kw


_KPKW_HEADER_BG = {"red": 0.20, "green": 0.44, "blue": 0.69}   # ciemny niebieski
_KPKW_CAT_KP_BG = {"red": 0.78, "green": 0.92, "blue": 0.78}  # zielony nagłówek kategorii
_KPKW_CAT_KW_BG = {"red": 0.96, "green": 0.78, "blue": 0.78}  # czerwony nagłówek kategorii
_KPKW_DATA_KP   = {"red": 0.92, "green": 0.98, "blue": 0.90}  # bardzo jasny zielony
_KPKW_DATA_KW   = {"red": 0.99, "green": 0.92, "blue": 0.90}  # bardzo jasny różowy
_KPKW_STAN_BG   = {"red": 1.0,  "green": 0.95, "blue": 0.77}  # złoty — stan kasy
_WHITE          = {"red": 1.0,  "green": 1.0,  "blue": 1.0}


def _build_kp_kw_block(subfolder_name, kp_entries, kw_entries):
    """Buduje (rows, row_types) dla bloku miesięcznego.
    row_types: 'marker' | 'header' | 'cat_header' | 'data' | 'separator'
    KP kwoty: dodatnie (kolumna C). KW kwoty: ujemne (kolumna G).
    Nagłówek: C = suma KP, D = bilans (KP+KW), G = suma KW (ujemna).
    """
    kp_total = sum(e[2] for e in kp_entries)
    kw_total = sum(e[2] for e in kw_entries)
    bilans   = kp_total - kw_total

    def group(entries):
        d = {}
        for e in sorted(entries, key=lambda x: (_kp_kw_cat(x[0]), x[3])):
            d.setdefault(_kp_kw_cat(e[0]), []).append(e)
        return d

    kp_by_cat = group(kp_entries)
    kw_by_cat = group(kw_entries)

    rows, types = [], []

    def add(row, rtype):
        rows.append(row)
        types.append(rtype)

    add([_KP_KW_MARKER.format(subfolder_name), "", "", "", "", "", ""], "marker")
    # C = suma KP, D = bilans, G = suma KW (ujemna)
    add([
        _month_label_pl(subfolder_name),
        "",
        round(kp_total, 2),
        round(bilans, 2),
        "",
        "",
        round(-kw_total, 2),
    ], "header")

    for cat in sorted(set(list(kp_by_cat) + list(kw_by_cat))):
        kp_cat = kp_by_cat.get(cat, [])
        kw_cat = kw_by_cat.get(cat, [])
        add([_CAT_LABELS_KP.get(cat, ""), "", "", "",
             _CAT_LABELS_KW.get(cat, ""), "", ""], "cat_header")
        for i in range(max(len(kp_cat), len(kw_cat))):
            kp_part = ["", "", ""]
            kw_part = ["", "", ""]
            if i < len(kp_cat):
                k, a, q, d = kp_cat[i]
                kp_part = [d, _kp_kw_opis(k, a), round(q, 2)]
            if i < len(kw_cat):
                k, a, q, d = kw_cat[i]
                kw_part = [d, _kp_kw_opis(k, a), round(-q, 2)]  # KW ujemne
            add(kp_part + [""] + kw_part, "data")

    add(["", "", "", "", "", "", ""], "separator")
    return rows, types


def preview_kp_kw_html(subfolder_name, sections):
    """Generuje HTML podglądu KP/KW z sekcji arkusza (zero zapisu do Sheets)."""
    kp_entries, kw_entries = _extract_rk_entries(sections)
    block, row_types = _build_kp_kw_block(subfolder_name, kp_entries, kw_entries)

    HDR  = "#3370B0"
    CHKP = "#C7EBC7"
    CHKW = "#F5C7C7"
    DKPP = "#EBFAE6"
    DKWP = "#FCEBE6"

    def td(val, bg="#ffffff", bold=False, align="left", fg="#333333"):
        styles = f"padding:3px 7px;border:1px solid #ccc;background:{bg};color:{fg};"
        if bold: styles += "font-weight:bold;"
        v = str(val) if str(val).strip() else "&nbsp;"
        return f'<td style="{styles}text-align:{align}">{v}</td>'

    lines = ['<table style="width:100%;border-collapse:collapse;font-size:12px;font-family:sans-serif;background:#ffffff">']
    for row, rtype in zip(block, row_types):
        if rtype == "separator":
            continue
        r = [str(x) if str(x).strip() else "" for x in row]
        if rtype == "marker":
            lines.append(
                f'<tr><td colspan="7" style="background:#EDEDED;color:#333333;padding:3px 7px;'
                f'border:1px solid #ccc;font-weight:bold">{r[0]}</td></tr>'
            )
        elif rtype == "header":
            lines.append("<tr>")
            lines.append(td(r[0], HDR, bold=True, fg="white"))
            lines.append(td("",   HDR, fg="white"))
            lines.append(td(r[2], HDR, bold=True, align="right", fg="white"))
            lines.append(td(r[3], HDR, bold=True, align="right", fg="white"))
            lines.append(td("",   HDR, fg="white"))
            lines.append(td("",   HDR, fg="white"))
            lines.append(td(r[6], HDR, bold=True, align="right", fg="white"))
            lines.append("</tr>")
        elif rtype == "cat_header":
            lines.append("<tr>")
            lines.append(
                f'<td colspan="3" style="background:{CHKP};color:#333333;padding:3px 7px;'
                f'border:1px solid #ccc;font-weight:bold">{r[0] or "&nbsp;"}</td>'
            )
            lines.append(td(""))
            lines.append(
                f'<td colspan="3" style="background:{CHKW};color:#333333;padding:3px 7px;'
                f'border:1px solid #ccc;font-weight:bold">{r[4] or "&nbsp;"}</td>'
            )
            lines.append("</tr>")
        elif rtype == "data":
            has_kp = bool(r[0] or r[1] or r[2])
            has_kw = bool(r[4] or r[5] or r[6])
            lines.append("<tr>")
            lines.append(td(r[0], DKPP if has_kp else ""))
            lines.append(td(r[1], DKPP if has_kp else ""))
            lines.append(td(r[2], DKPP if has_kp else "", align="right"))
            lines.append(td(""))
            lines.append(td(r[4], DKWP if has_kw else ""))
            lines.append(td(r[5], DKWP if has_kw else ""))
            lines.append(td(r[6], DKWP if has_kw else "", align="right"))
            lines.append("</tr>")
    lines.append("</table>")
    return "\n".join(lines)


def _fmt_ranges(row_nums):
    """Zamienia listę numerów wierszy na listę (start, end) ciągłych zakresów."""
    if not row_nums:
        return []
    segs, start, prev = [], row_nums[0], row_nums[0]
    for r in row_nums[1:]:
        if r == prev + 1:
            prev = r
        else:
            segs.append((start, prev))
            start = prev = r
    segs.append((start, prev))
    return segs


def refresh_kp_kw(spreadsheet, subfolder_name, sections):
    """Nadpisuje blok biezacego miesiaca w zakładce 'Kp i Kw'.

    Logika stanu kasy (zakładka 'Kp i Kw'):
    ─────────────────────────────────────────
    • A1 = zawsze stan kasy na koniec NAJNOWSZEGO miesiąca w nowym systemie.
    • Tuż po separatorze każdego bloku (=== MMYYYY ===) stoi złoty wiersz
      zamknięcia (liczba w kolumnie A) = stan kasy na koniec POPRZEDNIEGO okresu
      (czyli stan sprzed tego bloku). Wiersz ten jest wstawiany raz przy
      pierwszym dodaniu miesiąca i NIE jest kasowany przy kolejnych odświeżeniach.
    • Przy PIERWSZYM uruchomieniu (np. 042026):
        - A1 przed wstawieniem = stary stan kasy wpisany ręcznie (np. 78 112,40)
        - Ten stan staje się wierszem zamknięcia pod blokiem (np. A29 = 78 112,40)
        - A1 = 78 112,40 + bilans_042026
    • Przy ODŚWIEŻENIU najnowszego miesiąca (blok na górze):
        - Stary stan kasy odczytywany z wiersza zamknięcia poniżej separatora
        - A1 = stary_stan + nowy_bilans_biezacego_miesiaca
    • Przy ODŚWIEŻENIU starszego miesiąca (bloki nowsze leżą powyżej):
        - Liczymy delta = nowy_bilans - stary_bilans (sprzed odświeżenia)
        - A1 += delta (stan kasy najnowszego miesiąca przesuwa się o deltę)
        - Złote wiersze zamknięcia wszystkich nowszych bloków += delta
          (każdy z nich przechowuje stan kasy końca poprzedniego okresu —
           muszą być skorygowane, żeby łańcuch sum był spójny)
    • Przy DODANIU kolejnego miesiąca (052026 nad 042026):
        - Obecna wartość A1 (= 67 882,72 = koniec kwietnia) staje się wierszem
          zamknięcia pod nowym blokiem
        - A1 = 67 882,72 + bilans_052026

    Przygotowanie arkusza (raz, przed pierwszym odświeżeniem):
        1. W zakładce 'Kp i Kw' wpisz ręcznie w A1 stan kasy na koniec
           ostatniego miesiąca starego systemu (np. 78 112,40).
        2. Poniżej (od wiersza 2) możesz mieć dowolne stare dane KP/KW.
        3. Kliknij 'Odśwież KP / KW' — reszta dzieje się automatycznie.
    """
    try:
        ws = spreadsheet.worksheet(_KP_KW_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=_KP_KW_SHEET, rows=600, cols=10)

    kp_entries, kw_entries = _extract_rk_entries(sections)
    bilans_current = sum(e[2] for e in kp_entries) - sum(e[2] for e in kw_entries)
    block, row_types = _build_kp_kw_block(subfolder_name, kp_entries, kw_entries)

    all_vals = ws.get_all_values()
    marker   = _KP_KW_MARKER.format(subfolder_name)

    def _parse_num(s):
        try:
            return float(re.sub(r"[^\d\-.,]", "", str(s)).replace(",", "."))
        except Exception:
            return 0.0

    # ── Wyszukaj marker i separator końca bloku ───────────────────
    start_row = None
    sep_idx   = None   # 0-based index separatora (pustego wiersza)
    end_row   = None   # 1-based: kasuj wiersze start_row..end_row-1 (separator włącznie)

    for i, row in enumerate(all_vals):
        val = str(row[0]).strip() if row else ""
        if val == marker:
            start_row = i + 1
        elif start_row is not None:
            if not any(str(c).strip() for c in row):
                sep_idx = i
                end_row = i + 2   # kasuj do separatora; wiersz zamknięcia (i+2) zostaje
                break
            elif re.match(r'^=== \d{6} ===$', val):
                end_row = i + 1   # awaryjnie (brak separatora)
                break

    # Zmienne do kaskadowej aktualizacji złotych wierszy przy odświeżeniu starszego miesiąca
    is_newest     = True   # True gdy odświeżany blok jest najnowszym (góra arkusza)
    old_bilans    = 0.0    # bilans bloku sprzed odświeżenia (do wyliczenia delta)
    newer_closing = []     # lista (1-based row, old_value) złotych wierszy nowszych bloków

    if start_row is None:
        # ── Nowy miesiąc ─────────────────────────────────────────
        # Obecna A1 = poprzedni stan kasy (stary system lub poprzedni miesiąc)
        old_base  = _parse_num(all_vals[0][0] if all_vals and all_vals[0] else "")
        insert_at = 2 if all_vals and any(c for c in all_vals[0]) else 1
        _api(ws.insert_rows, block, row=insert_at)
        start_row = insert_at

        # Wstaw wiersz zamknięcia (poprzedni stan kasy) tuż po separatorze bloku
        if old_base:
            closing_num = insert_at + len(block)
            _api(ws.insert_rows, [[round(old_base, 2), "", "", "", "", "", ""]], row=closing_num)
            _api(ws.format, f"A{closing_num}:G{closing_num}",
                 {"backgroundColor": _KPKW_STAN_BG, "textFormat": {"bold": True}})
    else:
        # ── Odświeżenie istniejącego bloku ───────────────────────
        # Sprawdź czy to najnowszy blok (brak innych markerów powyżej)
        is_newest = not any(
            re.match(r'^=== \d{6} ===$', str(r[0]).strip() if r else "")
            for r in all_vals[:start_row - 1]
        )

        # Stary bilans z nagłówka bloku (kolumna D = index 3), przed usunięciem
        if start_row < len(all_vals):
            old_bilans = _parse_num(all_vals[start_row][3])

        # Złote wiersze zamknięcia nowszych bloków — zbieramy przed modyfikacją arkusza
        if not is_newest:
            j = 0
            while j < start_row - 1:
                val = str(all_vals[j][0]).strip() if all_vals[j] else ""
                if re.match(r'^=== \d{6} ===$', val):
                    k = j + 1
                    while k < start_row - 1:
                        if not any(str(c).strip() for c in all_vals[k]):
                            cr = k + 2  # 1-based row number złotego wiersza
                            if k + 1 < len(all_vals):
                                newer_closing.append((cr, _parse_num(all_vals[k + 1][0])))
                            break
                        k += 1
                j += 1

        # Stary stan kasy = wiersz zamknięcia poniżej separatora
        old_base = 0.0
        if sep_idx is not None and sep_idx + 1 < len(all_vals):
            old_base = _parse_num(all_vals[sep_idx + 1][0])

        if end_row is None:
            for j in range(start_row - 1, len(all_vals)):
                if not any(str(c).strip() for c in all_vals[j]):
                    sep_idx  = j
                    end_row  = j + 2
                    if j + 1 < len(all_vals):
                        old_base = _parse_num(all_vals[j + 1][0])
                    break
            if end_row is None:
                end_row = start_row + len(block)

        _api(ws.delete_rows, start_row, end_row - 1)
        _api(ws.insert_rows, block, row=start_row)

    # ── A1 = stan kasy bieżącego miesiąca ────────────────────────
    # Przy odświeżeniu starszego miesiąca: A1 i złote wiersze nowszych bloków
    # przesuwamy o deltę — nie tracimy bilansu miesięcy leżących powyżej.
    if start_row is not None and not is_newest:
        delta    = round(bilans_current - old_bilans, 2)
        stan_net = round(_parse_num(all_vals[0][0]) + delta, 2)
        for cr, old_val in newer_closing:
            _api(ws.update, f"A{cr}", [[round(old_val + delta, 2)]])
            _api(ws.format, f"A{cr}:G{cr}",
                 {"backgroundColor": _KPKW_STAN_BG, "textFormat": {"bold": True}})
    else:
        stan_net = round(old_base + bilans_current, 2)
    _api(ws.update, "A1:G1", [[stan_net, "", "", "", "", "", ""]])
    _api(ws.format, "A1", {"backgroundColor": _KPKW_STAN_BG, "textFormat": {"bold": True}})

    # ── Formatowanie bloku ────────────────────────────────────────
    type_rows = {"header": [], "cat_header": [], "data": [], "marker": [], "separator": []}
    for i, rtype in enumerate(row_types):
        type_rows[rtype].append(start_row + i)

    block_end = start_row + len(block) - 1
    _api(ws.format, f"A{start_row}:G{block_end}", {"backgroundColor": _WHITE})

    for r in type_rows["marker"]:
        _api(ws.format, f"A{r}:G{r}", {"backgroundColor": {"red": 0.93, "green": 0.93, "blue": 0.93}})

    for r in type_rows["header"]:
        _api(ws.format, f"A{r}:G{r}", {
            "backgroundColor": _KPKW_HEADER_BG,
            "textFormat": {"bold": True, "foregroundColor": _WHITE},
        })

    for r in type_rows["cat_header"]:
        _api(ws.format, f"A{r}:C{r}", {"backgroundColor": _KPKW_CAT_KP_BG,
                                        "textFormat": {"bold": True}})
        _api(ws.format, f"E{r}:G{r}", {"backgroundColor": _KPKW_CAT_KW_BG,
                                        "textFormat": {"bold": True}})

    for r in type_rows["data"]:
        _api(ws.format, f"A{r}:C{r}", {"backgroundColor": _KPKW_DATA_KP})
        _api(ws.format, f"E{r}:G{r}", {"backgroundColor": _KPKW_DATA_KW})


# ----------------------------------------------------------------
# SEARCH — wyszukiwanie na Google Drive
# ----------------------------------------------------------------

def _get_item_path(service, item_id, cache):
    """Buduje pełną ścieżkę elementu na Drive przez przechodzenie po rodzicach."""
    if item_id in cache:
        return cache[item_id]
    try:
        meta = service.files().get(
            fileId=item_id,
            fields="name,parents",
            supportsAllDrives=True,
        ).execute()
        name = meta.get("name", "?")
        parents = meta.get("parents", [])
        if not parents:
            result = name
        else:
            parent_path = _get_item_path(service, parents[0], cache)
            result = f"{parent_path}/{name}"
    except Exception:
        result = "?"
    cache[item_id] = result
    return result


def search_drive_items(service, query_text, search_type):
    """
    Szuka plików lub folderów na Drive zawierających query_text w nazwie.
    search_type: 'Pliki' lub 'Foldery'
    Zwraca listę słowników {"Nazwa": ..., "Ścieżka": ...}
    """
    safe = query_text.replace("'", "\\'")
    if search_type == "Foldery":
        mime = "mimeType = 'application/vnd.google-apps.folder'"
    else:
        mime = "mimeType != 'application/vnd.google-apps.folder'"
    q = f"name contains '{safe}' and {mime} and trashed = false"

    resp = service.files().list(
        q=q,
        fields="files(id, name, parents, webViewLink)",
        pageSize=50,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    cache = {}
    results = []
    for item in resp.get("files", []):
        name    = item["name"]
        parents = item.get("parents", [])
        link    = item.get("webViewLink", "")
        if search_type == "Pliki":
            path = _get_item_path(service, parents[0], cache) if parents else "/"
        else:
            path = _get_item_path(service, item["id"], cache)
        results.append({"Link": link, "Nazwa": name, "Ścieżka": path})

    return results


_KLUCZ_IDX = HEADER_ROW.index("Klucz_Ksiegowy")

# Dopasowanie tagów do kolumny Klucz_Ksiegowy
_SHEET_TAG_MATCHERS = {
    "kos":      lambda k: k.startswith("kos_"),
    "prz_naj":  lambda k: k.startswith("prz_"),
    "wla":      lambda k: k.startswith("wla_"),
    "nieznany": lambda k: k.startswith("nieznany_"),
    "rk_kp":    lambda k: "rk_kp" in k,
    "rk_kw":    lambda k: "rk_kw" in k,
    "pr_in":    lambda k: "pr_in" in k and "pr_out" not in k,
    "pr_out":   lambda k: "pr_out" in k,
    "depo":     lambda k: "depo" in k,
    "roz":      lambda k: "_roz_" in k and "depo" not in k,
}

# Aliasy wpisywane w polu tekstowym (np. "kos", "kosztowe", "netia*kos")
_TAG_ALIASES = {
    "kos": "kos", "kosztowe": "kos",
    "prz": "prz_naj", "prz_naj": "prz_naj", "najemcy": "prz_naj",
    "wla": "wla", "wlasciciel": "wla",
    "nieznany": "nieznany",
    "rk_kp": "rk_kp", "kp": "rk_kp",
    "rk_kw": "rk_kw", "kw": "rk_kw",
    "pr_in": "pr_in",
    "pr_out": "pr_out",
    "depo": "depo", "kaucja": "depo",
    "roz": "roz",
}


def _parse_sh_query(raw):
    """Parsuje zapytanie z pola tekstowego.
    'netia*kos'  -> ('netia', ['kos'], 'AND')
    'kos'        -> ('',      ['kos'], 'AND')
    'netia'      -> ('netia', [],      'OR')
    """
    raw = raw.strip()
    if "*" in raw:
        parts = [p.strip() for p in raw.split("*")]
        text_parts = [p for p in parts if p.lower() not in _TAG_ALIASES]
        tag_parts  = [_TAG_ALIASES[p.lower()] for p in parts if p.lower() in _TAG_ALIASES]
        return (" ".join(text_parts).strip(), tag_parts, "AND")
    low = raw.lower()
    if low in _TAG_ALIASES:
        return ("", [_TAG_ALIASES[low]], "AND")
    return (raw, [], "OR")


_MMYYYY_PAT = re.compile(r"^(0[1-9]|1[0-2])\d{4}$")


def search_sheet_rows(spreadsheet, query_text, sheet_filter=None, tags=None, mode="OR"):
    """
    Szuka wierszy w Google Sheets.
    query_text: tekst do szukania w dowolnej kolumnie (lub '').
    tags:       lista tagów z _SHEET_TAG_MATCHERS (lub None/[]).
    mode:       'OR' — tekst LUB tag; 'AND' — tekst I tag (oba muszą pasowac).
    sheet_filter: None lub '' = zakładki MMYYYY; inaczej = konkretna zakladka.
    Zwraca (wyniki: list[dict], nazwy_zakladek: list[str]).
    """
    q            = (query_text or "").strip().lower()
    active_tags  = tags or []
    all_worksheets = spreadsheet.worksheets()
    sheet_names    = [ws.title for ws in all_worksheets]

    if sheet_filter:
        worksheets = [ws for ws in all_worksheets if ws.title == sheet_filter]
    else:
        worksheets = [ws for ws in all_worksheets if _MMYYYY_PAT.match(ws.title)]

    results = []
    for ws in worksheets:
        for row in ws.get_all_values():
            if not any(cell for cell in row):
                continue
            if row and _match_separator(row[0]):
                continue
            if row and row[0] == HEADER_ROW[0]:
                continue
            padded = row + [""] * max(0, 13 - len(row))
            klucz  = padded[_KLUCZ_IDX].lower()

            text_match = bool(q) and any(q in str(cell).lower() for cell in row)
            tag_match  = bool(active_tags) and all(
                _SHEET_TAG_MATCHERS[t](klucz)
                for t in active_tags
                if t in _SHEET_TAG_MATCHERS
            )

            if q and active_tags:
                hit = text_match and tag_match
            elif q:
                hit = text_match
            else:
                hit = tag_match

            if hit:
                entry = {"Zakladka": ws.title}
                for j, col in enumerate(HEADER_ROW):
                    entry[col] = padded[j] if j < len(padded) else ""
                results.append(entry)

    return results, sheet_names


# ----------------------------------------------------------------
# WIDOK NAJEMCY — wyszukiwanie transakcji i PDF po imieniu/nazwisku
# ----------------------------------------------------------------

def _month_tab_range(od_str, do_str):
    """Zwraca listę nazw zakładek MMYYYY dla zakresu Od (MM/YYYY) Do (MM/YYYY)."""
    m1, y1 = int(od_str[:2]), int(od_str[3:])
    m2, y2 = int(do_str[:2]), int(do_str[3:])
    tabs = []
    y, m = y1, m1
    while (y, m) <= (y2, m2):
        tabs.append(f"{m:02d}{y}")
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return tabs


def search_najemca_sheets(spreadsheet, imie, nazwisko, tabs, mode="AND"):
    """
    Szuka wierszy pasujących do najemcy w zakładkach arkusza.
    Sprawdza kolumnę A (Nazwa/Plik), G (Klucz_Ksiegowy) i P (wyciag_Imie_Nazwisko).
    mode='AND': oba człony muszą pasować (gdy podano oba).
    mode='OR':  wystarczy jeden człon.
    Jedno puste pole → szuka tylko po podanym.
    Zwraca listę słowników.
    """
    imie_n = _normalize_name_for_filename(imie) if imie.strip() else ""
    nazw_n = _normalize_name_for_filename(nazwisko) if nazwisko.strip() else ""

    def _col_hit(col_val):
        if imie_n and nazw_n:
            if mode == "OR":
                return imie_n in col_val or nazw_n in col_val
            return imie_n in col_val and nazw_n in col_val
        if imie_n:
            return imie_n in col_val
        if nazw_n:
            return nazw_n in col_val
        return False

    all_ws = {ws.title: ws for ws in spreadsheet.worksheets()}
    results = []

    for tab in tabs:
        if tab not in all_ws:
            continue
        for row in all_ws[tab].get_all_values():
            if not any(row):
                continue
            if row[0] == HEADER_ROW[0] or _match_separator(row[0]):
                continue
            padded = row + [""] * max(0, 14 - len(row))
            col_a          = _normalize_name_for_filename(padded[0])
            col_klucz      = _normalize_name_for_filename(padded[3])
            col_kontrahent = _normalize_name_for_filename(padded[4])
            col_wyc        = _normalize_name_for_filename(padded[12])
            if _col_hit(col_a) or _col_hit(col_klucz) or _col_hit(col_kontrahent) or _col_hit(col_wyc):
                entry = {"Zakladka": tab}
                for _ci, _col_name in enumerate(HEADER_ROW):
                    entry[_col_name] = padded[_ci] if _ci < len(padded) else ""
                results.append(entry)
    return results


def search_najemca_pdfs(service, imie, nazwisko, tabs, mode="AND"):
    """
    Szuka plików PDF pasujących do najemcy w podfolderach Drive (Faktury sprzedazy MMYYYY).
    Zwraca listę słowników z nazwą pliku i linkiem do Drive.
    """
    imie_n = _normalize_name_for_filename(imie) if imie.strip() else ""
    nazw_n = _normalize_name_for_filename(nazwisko) if nazwisko.strip() else ""

    def _fname_hit(fname):
        if imie_n and nazw_n:
            if mode == "OR":
                return imie_n in fname or nazw_n in fname
            return imie_n in fname and nazw_n in fname
        if imie_n:
            return imie_n in fname
        if nazw_n:
            return nazw_n in fname
        return False

    sprzedaz_folder = find_subfolder(service, FOLDER_ID, "Faktury-sprzedazy")
    if not sprzedaz_folder:
        return []

    results = []
    for tab in tabs:
        folder_name  = f"{tab} {FAKTURY_SPRZEDAZY_SUFFIX}"
        month_folder = find_subfolder(service, sprzedaz_folder["id"], folder_name)
        if not month_folder:
            continue
        for pdf in list_pdfs_from_drive(service, month_folder["id"]):
            fname = _normalize_name_for_filename(pdf["name"])
            if _fname_hit(fname):
                try:
                    meta = service.files().get(
                        fileId=pdf["id"], fields="webViewLink"
                    ).execute()
                    link = meta.get("webViewLink", "")
                except Exception:
                    link = ""
                # kwota z nazwy pliku: ostatni segment numeryczny przed .pdf
                amt_match = re.search(r"_(\d+)\.pdf$", pdf["name"])
                kwota_pdf = int(amt_match.group(1)) if amt_match else None
                results.append({
                    "Zakladka":    tab,
                    "Nazwa pliku": pdf["name"],
                    "Kwota":       kwota_pdf,
                    "Link":        link,
                })
    return results


def find_drive_folders_by_name(service, imie, nazwisko, mode="AND"):
    """
    Szuka folderów na Drive zawierających imię/nazwisko najemcy oraz słowo 'fvs'.
    Zwraca listę słowników z nazwą folderu i linkiem.
    """
    imie_n = imie.strip().lower()
    nazw_n = nazwisko.strip().lower()

    name_parts = []
    if imie_n:
        name_parts.append(f"name contains '{imie_n}'")
    if nazw_n:
        name_parts.append(f"name contains '{nazw_n}'")
    if not name_parts:
        return []

    join_op = " and " if mode == "AND" else " or "
    name_query = join_op.join(name_parts)
    full_query = (
        f"mimeType = 'application/vnd.google-apps.folder' "
        f"and name contains 'fvs' "
        f"and ({name_query}) "
        f"and trashed = false"
    )
    try:
        res = service.files().list(
            q=full_query,
            fields="files(id, name, webViewLink)",
            pageSize=20,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        return res.get("files", [])
    except Exception:
        return []


# ----------------------------------------------------------------
# INTERFEJS STREAMLIT
# ----------------------------------------------------------------

st.set_page_config(
    page_title="System Fakturowania",
    page_icon=":page_facing_up:",
    layout="wide",
)

# ================================================================
# AUTENTYKACJA
# ================================================================
_role = "admin"   # domyslnie pelny dostep (gdy brak konfiguracji [auth] w secrets)
_AUTH_KEY = "abido_auth_user"

if "auth" in st.secrets:
    if _AUTH_KEY not in st.session_state:
        st.session_state[_AUTH_KEY] = None

    if st.session_state[_AUTH_KEY] is None:
        st.title("System Fakturowania \u2014 Logowanie")
        with st.form("login_form"):
            _lu = st.text_input("Login")
            _lp = st.text_input("Has\u0142o", type="password")
            _lb = st.form_submit_button("Zaloguj", use_container_width=True)
        if _lb:
            _ph = hashlib.sha256(_lp.encode()).hexdigest()
            _ah = st.secrets["auth"].get("admin_hash", "")
            _kh = st.secrets["auth"].get("ksiegowa_hash", "")
            if _lu == "admin" and _ph == str(_ah):
                st.session_state[_AUTH_KEY] = "admin"
                st.rerun()
            elif _lu == "ksiegowa" and _ph == str(_kh):
                st.session_state[_AUTH_KEY] = "ksiegowa"
                st.rerun()
            else:
                st.error("\u274c Nieprawid\u0142owy login lub has\u0142o.")
        st.stop()

    _role = st.session_state[_AUTH_KEY]

st.markdown("""
<style>
/* Czerwona ramka dla wewnetrznego kontenera w 3. kolumnie segmentu miesiac */
div[data-testid="stVerticalBlock"]
  > div[data-testid="stHorizontalBlock"]
  > div[data-testid="stColumn"]:nth-child(3)
  > div
  > div[data-testid="stVerticalBlockBorderWrapper"] {
    border-color: #cc2222 !important;
    border-width: 2px !important;
}
/* Tlo segmentu miesiac — fioletowe */
details[data-testid="stExpander"]:has(.abido-month-bg),
div[data-testid="stVerticalBlockBorderWrapper"]:has(.abido-month-bg),
div[data-testid="stVerticalBlockBorderWrapper"]:has(.abido-month-bg) > div {
    background-color: rgba(130, 60, 220, 0.22) !important;
    border-color: rgba(160, 90, 255, 0.6) !important;
}
/* Tlo bilansu najemcy — niebieskawa */
details[data-testid="stExpander"]:has(.abido-bilans-bg) {
    background-color: rgba(40, 130, 220, 0.18) !important;
    border-color: rgba(60, 160, 255, 0.6) !important;
}
/* Tlo wyswietlonego arkusza (ex) — zielone */
div[data-testid="stVerticalBlockBorderWrapper"]:has(.abido-ex-bg),
div[data-testid="stVerticalBlockBorderWrapper"]:has(.abido-ex-bg) > div {
    background-color: rgba(40, 180, 90, 0.18) !important;
    border-color: rgba(60, 200, 100, 0.6) !important;
}
/* Przycisk Szukaj bilansu — zielony po zakonczeniu wyszukiwania */
div[data-testid="stColumn"]:has(.abido-nj-search-done) button {
    background-color: #1a7a38 !important;
    border-color: #1a7a38 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

_title_col, _user_col = st.columns([7, 1])
with _title_col:
    st.title("System Fakturowania")
with _user_col:
    if "auth" in st.secrets and st.session_state.get(_AUTH_KEY):
        _label = "Admin" if _role == "admin" else "Ksi\u0119gowa"
        st.markdown(f"<div style='text-align:right;padding-top:6px'>\U0001f464 <b>{_label}</b></div>", unsafe_allow_html=True)
        if st.button("Wyloguj", use_container_width=True):
            st.session_state[_AUTH_KEY] = None
            st.rerun()

# ── Szukanie Google Drive ────────────────────────────────────────────
with st.expander("Szukanie Google Drive", expanded=False, key="exp_szukanie_drive"):
    srch_input_col, srch_type_col, srch_btn_col = st.columns([5, 1.2, 0.5])
    with srch_input_col:
        search_query = st.text_input(
            "Szukaj na Drive",
            placeholder="Szukaj na Google Drive (nazwa pliku lub folderu)...",
            label_visibility="collapsed",
        )
    with srch_type_col:
        search_type = st.radio(
            "Typ",
            ["Pliki", "Foldery"],
            horizontal=True,
            label_visibility="collapsed",
        )
    with srch_btn_col:
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
        btn_search = st.button("🔍 Szukaj", use_container_width=True)

# ── Szukanie Google Sheets ───────────────────────────────────────────
with st.expander("Szukanie Google Sheets", expanded=False, key="exp_szukanie_sheets"):
    sh_r1c1, sh_r1c2, sh_r1c3 = st.columns([4.5, 1.5, 0.8])
    with sh_r1c1:
        sh_query = st.text_input(
            "Szukaj w Sheets",
            placeholder="słowo, tag (kos/prz/roz...) lub słowo*tag",
            label_visibility="collapsed",
        )
    with sh_r1c2:
        sh_tab_options = ["Wszystkie"] + st.session_state.get("sheet_tab_names", [])
        sh_tab_selected = st.selectbox(
            "Zakladka",
            sh_tab_options,
            label_visibility="collapsed",
        )
    with sh_r1c3:
        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
        btn_sh_search = st.button("🔍 Szukaj", use_container_width=True, key="btn_sh_search")

    # Checkboxy tagów
    _SH_TAG_KEYS = [
        "sh_tag_kos", "sh_tag_prz", "sh_tag_wla", "sh_tag_nieznany", "sh_tag_roz",
        "sh_tag_rk_kp", "sh_tag_rk_kw", "sh_tag_pr_in", "sh_tag_pr_out", "sh_tag_depo",
    ]
    _sh_th1, _sh_th2, _sh_th3 = st.columns([4, 0.9, 0.9])
    with _sh_th1:
        st.markdown("**Tagi (Klucz_Ksiegowy):**")
    with _sh_th2:
        if st.button("✓ wszystkie", key="sh_tags_all", use_container_width=True):
            for _k in _SH_TAG_KEYS:
                st.session_state[_k] = True
            st.rerun()
    with _sh_th3:
        if st.button("✗ odznacz", key="sh_tags_none", use_container_width=True):
            for _k in _SH_TAG_KEYS:
                st.session_state[_k] = False
            st.rerun()
    _sh_tc = st.columns(5)
    with _sh_tc[0]:
        sh_tag_kos      = st.checkbox("kos — kosztowe",     key="sh_tag_kos")
    with _sh_tc[1]:
        sh_tag_prz      = st.checkbox("prz — najemcy",      key="sh_tag_prz")
    with _sh_tc[2]:
        sh_tag_wla      = st.checkbox("wla — właściciele",  key="sh_tag_wla")
    with _sh_tc[3]:
        sh_tag_nieznany = st.checkbox("nieznany",            key="sh_tag_nieznany")
    with _sh_tc[4]:
        sh_tag_roz      = st.checkbox("roz",                 key="sh_tag_roz")

    _sh_tc2 = st.columns(5)
    with _sh_tc2[0]:
        sh_tag_rk_kp  = st.checkbox("rk_kp — got. +",   key="sh_tag_rk_kp")
    with _sh_tc2[1]:
        sh_tag_rk_kw  = st.checkbox("rk_kw — got. -",   key="sh_tag_rk_kw")
    with _sh_tc2[2]:
        sh_tag_pr_in  = st.checkbox("pr_in — przelew +", key="sh_tag_pr_in")
    with _sh_tc2[3]:
        sh_tag_pr_out = st.checkbox("pr_out — przelew -",key="sh_tag_pr_out")
    with _sh_tc2[4]:
        sh_tag_depo   = st.checkbox("depo — kaucja",     key="sh_tag_depo")

    sh_logic = st.radio(
        "Logika wielu warunków",
        ["OR — dowolny pasuje", "AND — wszystkie muszą pasować"],
        horizontal=True,
        key="sh_logic",
        label_visibility="visible",
    )

    # ── Wyniki dynamiczne (filtrowane przez tagi bez ponownego Szukaj) ──
    if "sh_results" in st.session_state:
        import pandas as pd
        _sh_stored = st.session_state["sh_results"]
        _sh_all_rows = _sh_stored["rows"]
        _sh_qlabel   = _sh_stored["label"]

        # Zbierz aktywne tagi
        _sh_active = []
        if sh_tag_kos:      _sh_active.append("kos")
        if sh_tag_prz:      _sh_active.append("prz_naj")
        if sh_tag_wla:      _sh_active.append("wla")
        if sh_tag_nieznany: _sh_active.append("nieznany")
        if sh_tag_roz:      _sh_active.append("roz")
        if sh_tag_rk_kp:    _sh_active.append("rk_kp")
        if sh_tag_rk_kw:    _sh_active.append("rk_kw")
        if sh_tag_pr_in:    _sh_active.append("pr_in")
        if sh_tag_pr_out:   _sh_active.append("pr_out")
        if sh_tag_depo:     _sh_active.append("depo")

        _sh_mode   = "AND" if sh_logic.startswith("AND") else "OR"
        _sh_tag_fn = all if _sh_mode == "AND" else any

        def _sh_row_matches(row):
            if not _sh_active:
                return True
            klucz = row.get("Klucz_Ksiegowy", "").lower()
            return _sh_tag_fn(
                _SHEET_TAG_MATCHERS[t](klucz)
                for t in _sh_active
                if t in _SHEET_TAG_MATCHERS
            )

        _sh_filtered = [r for r in _sh_all_rows if _sh_row_matches(r)]
        _tag_str = f" [{_sh_mode}: {', '.join(_sh_active)}]" if _sh_active else ""
        st.markdown(
            f"**Wyniki: {_sh_qlabel}{_tag_str} — {len(_sh_filtered)} z {len(_sh_all_rows)} wierszy**"
        )
        if _sh_filtered:
            _sh_df = pd.DataFrame(_sh_filtered)
            _sh_sum_b = sum(_parse_amount(r.get("Kwota brutto", "")) or 0.0 for r in _sh_filtered)
            _sh_sum_f = sum(_parse_amount(r.get("wyciag_Kwota", "")) or 0.0 for r in _sh_filtered)
            _sh_cnt_a = sum(1 for r in _sh_filtered if str(r.get("Nazwa / Plik", "")).strip())
            _sh_cnt_s = sum(1 for r in _sh_filtered if str(r.get("Status", "")).strip())
            _sh_sum_row = {c: "" for c in _sh_df.columns}
            _sh_sum_row["Nazwa / Plik"]  = f"Ilość A: {_sh_cnt_a}"
            _sh_sum_row["Kwota brutto"]  = f"{_sh_sum_b:,.2f}".replace(",", " ")
            _sh_sum_row["Status"]        = f"Ilość: {_sh_cnt_s}"
            _sh_sum_row["wyciag_Kwota"]  = f"{_sh_sum_f:,.2f}".replace(",", " ")
            _sh_df_sum = pd.concat([_sh_df, pd.DataFrame([_sh_sum_row])], ignore_index=True)
            _sh_last = len(_sh_df_sum) - 1
            def _sh_highlight(row):
                if row.name == _sh_last:
                    return ["background-color:#FFF3CD;font-weight:bold"] * len(row)
                return [""] * len(row)
            _sh_height = min((len(_sh_df_sum) + 1) * 35 + 3, 600)
            st.dataframe(
                _sh_df_sum.style.apply(_sh_highlight, axis=1),
                use_container_width=True,
                hide_index=True,
                height=_sh_height,
            )
            _sh_csv = _sh_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
            st.download_button(
                "⬇ Pobierz CSV",
                data=_sh_csv,
                file_name=f"szukanie_{_sh_qlabel.replace(' ', '_')}.csv",
                mime="text/csv",
                key="btn_sh_csv",
            )
        else:
            st.info("Brak wyników dla wybranych tagów.")

# ================================================================
# BILANS NAJEMCY
# ================================================================
with st.expander("Bilans najemcy", expanded=False, key="exp_bilans_najemcy"):
    st.markdown('<span class="abido-bilans-bg"></span>', unsafe_allow_html=True)

    nj_r1c1, nj_r1c2, nj_r1c3 = st.columns([2, 2, 1.5])
    with nj_r1c1:
        nj_imie = st.text_input(
            "Imię", placeholder="np. Fuzi (samo imię wystarczy)", key="nj_imie"
        )
    with nj_r1c2:
        nj_nazwisko = st.text_input(
            "Nazwisko", placeholder="np. Yang (samo nazwisko wystarczy)", key="nj_nazwisko"
        )
    with nj_r1c3:
        nj_name_mode = st.radio(
            "Logika imię+nazwisko",
            ["AND", "OR"],
            horizontal=True,
            key="nj_name_mode",
            help="AND: oba muszą pasować  |  OR: wystarczy jedno",
        )

    _nj_month_opts = [f"{m:02d}/{y}" for y in range(2000, 2100) for m in range(1, 13)]
    _today = date.today()
    _od_year  = _today.year - 1 if _today.month > 1 else _today.year - 2
    _od_month = _today.month - 1 if _today.month > 1 else 12
    _idx_od = (_od_year - 2000) * 12 + (_od_month - 1)
    _idx_do = (_today.year - 2000) * 12 + (_today.month - 1)
    nj_dc1, nj_dc2, nj_bc = st.columns([2, 2, 1])
    with nj_dc1:
        nj_od = st.selectbox("Od", _nj_month_opts, index=_idx_od, key="nj_od")
    with nj_dc2:
        nj_do = st.selectbox("Do", _nj_month_opts, index=_idx_do, key="nj_do")
    with nj_bc:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if "nj_results" in st.session_state:
            st.markdown('<span class="abido-nj-search-done"></span>', unsafe_allow_html=True)
        btn_nj_search = st.button("Szukaj", use_container_width=True, key="btn_nj_search")

    _NJ_FILTER_KEYS = ["nj_kp", "nj_kw", "nj_pr_in", "nj_pr_out", "nj_roz", "nj_depo", "nj_prz"]
    _nj_fc1, _nj_fc2, _nj_fc3, _nj_fc4, _nj_fc5 = st.columns([2.5, 0.9, 0.9, 1.2, 2.5])
    with _nj_fc1:
        st.markdown("**Filtry transakcji:**")
    with _nj_fc2:
        if st.button("✓ wszystkie", key="nj_filter_all", use_container_width=True):
            for _k in _NJ_FILTER_KEYS:
                st.session_state[_k] = True
            st.rerun()
    with _nj_fc3:
        if st.button("✗ odznacz", key="nj_filter_none", use_container_width=True):
            for _k in _NJ_FILTER_KEYS:
                st.session_state[_k] = False
            st.rerun()
    with _nj_fc4:
        nj_filter_mode = st.radio(
            "Logika filtrów",
            ["OR", "AND"],
            horizontal=True,
            key="nj_filter_mode",
            help="OR: pasuje dowolny zaznaczony typ  |  AND: klucz musi spełniać WSZYSTKIE zaznaczone",
        )
    with _nj_fc5:
        if "nj_results" in st.session_state:
            _nj_folders = st.session_state["nj_results"].get("drive_folders", [])
            if _nj_folders:
                _folder_links = " &nbsp;|&nbsp; ".join(
                    f"[📁 {_f['name']}]({_f['webViewLink']})" for _f in _nj_folders
                )
                st.markdown(f"**Foldery fvs:** {_folder_links}")
            else:
                st.caption("Brak folderów 'fvs' dla tej nazwy")
    nj_cc1, nj_cc2, nj_cc3, nj_cc4 = st.columns(4)
    with nj_cc1:
        st.markdown("**Gotówka (rk)**")
        nj_kp  = st.checkbox("kp — wpływ",  value=True, key="nj_kp")
        nj_kw  = st.checkbox("kw — wypływ", value=True, key="nj_kw")
    with nj_cc2:
        st.markdown("**Przelew**")
        nj_pr_in  = st.checkbox("pr_in — wpływ",   value=True, key="nj_pr_in")
        nj_pr_out = st.checkbox("pr_out — wypływ",  value=True, key="nj_pr_out")
    with nj_cc3:
        st.markdown("**Rozrachunkowe**")
        nj_roz  = st.checkbox("roz",            value=True, key="nj_roz")
        nj_depo = st.checkbox("depo — kaucja",  value=True, key="nj_depo")
    with nj_cc4:
        st.markdown("**Przychodowe**")
        nj_prz = st.checkbox("prz — przychody", value=True, key="nj_prz")

    # ── Wyniki ──────────────────────────────────────────────────
    if "nj_results" in st.session_state:
        import pandas as pd
        _nj = st.session_state["nj_results"]
        st.markdown(
            f"#### Wyniki: {_nj['imie']} {_nj['nazwisko']}  "
            f"({_nj['od']} – {_nj['do']})"
        )

        # Faktury PDF
        with st.expander(f"Faktury PDF ({len(_nj['pdfs'])})", expanded=True, key="exp_nj_pdfs"):
            if _nj["pdfs"]:
                for _p in _nj["pdfs"]:
                    _link  = _p.get("Link", "")
                    _fname = _p["Nazwa pliku"]
                    _kwota = _p.get("Kwota")
                    _kwota_str = f" — **{_kwota} zł**" if _kwota else ""
                    _tab = f"`{_p['Zakladka']}`"
                    if _link:
                        st.markdown(f"{_tab} &nbsp; [📄 {_fname}]({_link}){_kwota_str}")
                    else:
                        st.markdown(f"{_tab} &nbsp; 📄 {_fname}{_kwota_str}")
            else:
                st.caption("Brak faktur PDF dla tego najemcy w podanym zakresie.")

        # Filtr checkboxów
        def _row_matches_filters(klucz):
            k = klucz.lower()
            checks = []
            if nj_kp:     checks.append("rk_kp" in k)
            if nj_kw:     checks.append("rk_kw" in k)
            if nj_pr_in:  checks.append("pr_in" in k)
            if nj_pr_out: checks.append("pr_out" in k)
            if nj_depo:   checks.append("depo" in k)
            if nj_roz:    checks.append("_roz_" in k and "depo" not in k)
            if nj_prz:    checks.append(k.startswith("prz_"))
            if not checks:
                return True
            return all(checks) if nj_filter_mode == "AND" else any(checks)

        filtered_rows = [r for r in _nj["rows"] if _row_matches_filters(r["Klucz_Ksiegowy"])]

        # Transakcje
        with st.expander(
            f"Transakcje w arkuszu ({len(filtered_rows)} z {len(_nj['rows'])})",
            expanded=True,
            key="exp_nj_transakcje",
        ):
            if filtered_rows:
                _nj_df = pd.DataFrame(filtered_rows)
                st.dataframe(_nj_df, use_container_width=True, hide_index=True)
                _nj_csv = _nj_df.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
                st.download_button(
                    "⬇ Pobierz CSV",
                    data=_nj_csv,
                    file_name=f"bilans_{_nj['imie']}_{_nj['nazwisko']}.csv",
                    mime="text/csv",
                    key="btn_nj_csv",
                )
            else:
                st.caption("Brak transakcji dla wybranych filtrów.")

        # Bilans
        with st.expander("Bilans", expanded=True, key="exp_nj_bilans"):
            _prz_rows  = [r for r in _nj["rows"] if r["Klucz_Ksiegowy"].lower().startswith("prz_")]
            _depo_in   = [r for r in _nj["rows"]
                          if "depo" in r["Klucz_Ksiegowy"].lower()
                          and ("_in" in r["Klucz_Ksiegowy"].lower() or "_kp" in r["Klucz_Ksiegowy"].lower())]
            _depo_out  = [r for r in _nj["rows"]
                          if "depo" in r["Klucz_Ksiegowy"].lower()
                          and ("_out" in r["Klucz_Ksiegowy"].lower() or "_kw" in r["Klucz_Ksiegowy"].lower())]
            _bilans_rows = [r for r in _nj["rows"] if "depo" not in r["Klucz_Ksiegowy"].lower()]

            def _parse_signed(s):
                """Parsuje kwotę zachowując znak (ujemne koszty zostają ujemne)."""
                try:
                    return float(re.sub(r"[^\d,.\-]", "", str(s)).replace(",", "."))
                except (ValueError, TypeError):
                    return None

            def _sum_kwota_bil(rows):
                """Sumuje col B tylko z głównych wierszy (col A niepuste = faktura).
                Dla roz_depo: gdy col B pusta, bierze wyciag_Kwota (col F)."""
                total = 0.0
                for r in rows:
                    if not str(r.get("Nazwa / Plik", "")).strip():
                        continue  # sub-wiersz — pomiń
                    v = _parse_amount(r["Kwota brutto"])
                    if (v is None or v == 0.0) and "roz_depo" in r.get("Klucz_Ksiegowy", "").lower():
                        v = abs(_parse_signed(r.get("wyciag_Kwota", "")) or 0.0)
                    total += abs(v or 0.0)
                return total

            def _sum_bilans(rows):
                """Bilans (ze znakiem): wyciag_Kwota (bank) lub col B ze znakiem."""
                total = 0.0
                for r in rows:
                    v = _parse_signed(r.get("wyciag_Kwota", ""))
                    if v is None or v == 0.0:
                        v = _parse_signed(r["Kwota brutto"])
                    total += (v or 0.0)
                return total

            _prz_main     = [r for r in _prz_rows if str(r.get("Nazwa / Plik", "")).strip()]
            prz_naleznosc = _sum_kwota_bil(_prz_main)
            bilans_sum    = _sum_bilans(_bilans_rows)
            depo_in_sum   = _sum_kwota_bil(_depo_in)
            depo_out_sum  = _sum_kwota_bil(_depo_out)
            depo_saldo    = depo_in_sum - depo_out_sum

            _inne_rows = [
                r for r in _nj["rows"]
                if not r["Klucz_Ksiegowy"].lower().startswith("prz_")
                and "depo" not in r["Klucz_Ksiegowy"].lower()
            ]

            pdf_kwoty = [p["Kwota"] for p in _nj["pdfs"] if p.get("Kwota") is not None]
            pdf_sum   = sum(pdf_kwoty)

            def _fmt(v):
                return f"{v:,.2f} zł".replace(",", " ")

            b1, b2, b3, b4, b5 = st.columns(5)
            with b1:
                st.markdown("**Faktury arkusz**")
                st.metric("Pozycji prz", len(_prz_main))
                st.metric("Suma", _fmt(prz_naleznosc))
            with b2:
                st.markdown("**Bilans arkusz bez depo**")
                st.metric("Pozycji", len(_bilans_rows))
                st.metric("Bilans", _fmt(bilans_sum))
                _bil_bank = sum((_parse_signed(r.get("wyciag_Kwota", "")) or 0.0) for r in _bilans_rows)
                _bil_rk   = sum(
                    (_parse_signed(r["Kwota brutto"]) or 0.0)
                    for r in _bilans_rows
                    if not (_parse_signed(r.get("wyciag_Kwota", "")) or 0.0)
                )
                st.caption(f"wyciąg: {_fmt(_bil_bank)}  |  RK: {_fmt(_bil_rk)}")
            with b3:
                st.markdown("**Faktury PDF**")
                st.metric("Szt.", len(_nj["pdfs"]))
                st.metric("Suma", _fmt(pdf_sum))
            with b4:
                st.markdown("**Kaucja (depo)**")
                st.metric("Wpłacona", _fmt(depo_in_sum))
                st.metric("Zwrócona", _fmt(depo_out_sum))
                st.metric("Saldo", _fmt(depo_saldo))
            with b5:
                st.markdown("**Inne transakcje**")
                if _inne_rows:
                    for r in _inne_rows:
                        _kw = _parse_signed(r.get("wyciag_Kwota", "")) or _parse_signed(r["Kwota brutto"]) or 0.0
                        _klucz = r["Klucz_Ksiegowy"] or "(brak klucza)"
                        st.caption(f"{r['Zakladka']} | {_klucz} | {_fmt(_kw)}")
                else:
                    st.caption("—")

# ── Segment: miesiac + akcje ────────────────────────────────────────
# ── Domyslne wartosci przyciskow admina (ksiegowa ich nie widzi) ──
subfolder_name = ""
btn_wyswietl = btn_szablon = btn_czytaj = btn_sprawdz = btn_upload_ksef = False
btn_sprzedaz = btn_generuj_pdf = btn_sprawdz_sprzedaz = btn_paruj = btn_kolory_sprzedaz = False
btn_status_parowania = btn_refresh_kpkw = btn_show_kpkw = False
btn_sortuj_inne_rk = btn_usun_puste = btn_podsumowanie = btn_sort_kosztowe = btn_dodaj_puste = False

if _role == "admin":
    with st.expander("Miesiac — tworzenie faktur i parowanie", expanded=True, key="exp_miesiac"):
        st.markdown('<span class="abido-month-bg"></span>', unsafe_allow_html=True)
    
        # Pole miesiaca + Wyswietl ex
        _, input_col, btn_ex_col, _ = st.columns([1, 1.6, 0.4, 1])
        with input_col:
            subfolder_name = st.text_input(
                "Miesiac (np. 032026)",
                placeholder="wpisz nazwe podfolderu miesiacowego...",
            )
        with btn_ex_col:
            st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
            btn_wyswietl = st.button("Wyswietl ex", use_container_width=True)
    
        # Trzy kolumny akcji
        left_col, mid_col, right_col, extra_col = st.columns(4)
    
        @st.dialog("Podgląd KP / KW", width="large")
        def _dialog_podglad_kp_kw(html):
            st.markdown(html, unsafe_allow_html=True)
    
        @st.dialog("Stan faktur kosztowych", width="large")
        def _dialog_kosztowe_status(ks):
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Google Drive**")
                with st.container(border=True):
                    if not ks["subfolder_found"]:
                        st.warning("Nie znaleziono folderu na Drive.")
                    else:
                        st.metric("Pliki PDF", len(ks["drive_file_names"]))
            with col_b:
                st.markdown("**Google Sheets (sekcja kosztowa)**")
                with st.container(border=True):
                    if ks["sheet_counts"] is None:
                        st.warning("Brak arkusza o tej nazwie.")
                    else:
                        st.metric("Wierszy lacznie", sum(ks["sheet_counts"].values()))
                        st.markdown(
                            f"- Niezweryfikowane (0): **{ks['sheet_counts']['0']}**  \n"
                            f"- Zweryfikowane (1): **{ks['sheet_counts']['1']}**  \n"
                            f"- Inne: **{ks['sheet_counts']['inne']}**"
                        )
            only_drive  = ks["only_drive"]
            only_sheets = ks["only_sheets"]
            if only_drive is not None and only_sheets is not None:
                if only_drive or only_sheets:
                    st.markdown("---")
                    st.markdown("**Roznice:**")
                    d_col_a, d_col_b = st.columns(2)
                    with d_col_a:
                        st.markdown(f"**Na Drive, brak w Sheets ({len(only_drive)})**")
                        for fname in (only_drive or ["*(brak)*"]):
                            st.markdown(f"- {fname}")
                    with d_col_b:
                        st.markdown(f"**W Sheets, brak na Drive ({len(only_sheets)})**")
                        for fname in (only_sheets or ["*(brak)*"]):
                            st.markdown(f"- {fname}")
                else:
                    st.success("Drive i Sheets sa zgodne — brak roznic.")
    
        @st.dialog("Stan faktur sprzedazy", width="large")
        def _dialog_sprzedaz_status(ss):
            data = ss["data"]
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**Google Drive (plik zbiorczy)**")
                with st.container(border=True):
                    if data["drive_filename"]:
                        st.markdown(f"`{data['drive_filename']}`")
                        st.metric("Sztuk (z nazwy pliku)",
                                  data["drive_szt"] if data["drive_szt"] is not None else "?")
                        st.metric("Kwota zł (z nazwy pliku)",
                                  data["drive_kwota"] if data["drive_kwota"] is not None else "?")
                    else:
                        st.warning("Brak pliku Fs_najemcy_* na Drive.")
            with col_b:
                st.markdown("**Google Sheets (sekcja sprzedazy)**")
                with st.container(border=True):
                    st.metric("Wierszy z kwotą > 0", data["sheet_szt"])
                    st.metric("Suma kolumny B (zł)", f"{data['sheet_kwota']:,.0f}".replace(",", " "))
            if data["drive_filename"] and data["drive_szt"] is not None:
                szt_ok   = data["drive_szt"]  == data["sheet_szt"]
                kwota_ok = data["drive_kwota"] == int(round(data["sheet_kwota"]))
                if szt_ok and kwota_ok:
                    st.success("Drive i Sheets sa zgodne — sztuki i kwota sie zgadzaja.")
                else:
                    if not szt_ok:
                        st.warning(
                            f"Roznica sztuk: Drive={data['drive_szt']}, "
                            f"Sheets={data['sheet_szt']}"
                        )
                    if not kwota_ok:
                        st.warning(
                            f"Roznica kwoty: Drive={data['drive_kwota']} zl, "
                            f"Sheets={int(round(data['sheet_kwota']))} zl"
                        )
    
        @st.dialog("Wynik parowania", width="large")
        def _dialog_wynik_parowania(data):
            sparowane   = data["sparowane"]
            niesparowane = data["niesparowane"]
            fioletowe   = data["fioletowe"]
            pomaranczowe = data["pomaranczowe"]
            subfolder_name_d = data.get("subfolder_name", "")
            tx_total    = data["tx_total"]
            tx_sum      = data["tx_sum"]
            sheet_tx_count = data["sheet_tx_count"]
            sheet_tx_sum   = data["sheet_tx_sum"]
            diff_info   = data["diff_info"]
    
            parts = [f"**Sparowano: {sparowane}**"]
            if fioletowe:
                parts.append(f"🟣 Fioletowe (niezgodna kwota): {fioletowe}")
            if pomaranczowe:
                parts.append(f"🟠 Pomarańczowe (brak pary): {pomaranczowe}")
            parts.append(f"Niesparowane z wyciągu: {niesparowane}")
            st.success(" | ".join(parts))
    
            ok_count = (tx_total == sheet_tx_count)
            ok_sum   = (abs(tx_sum - sheet_tx_sum) < 0.02)
            count_icon = "✅" if ok_count else "⚠️"
            sum_icon   = "✅" if ok_sum   else "⚠️"
            st.info(
                f"{count_icon} Pozycje: Plik {tx_total} / Arkusz {sheet_tx_count}"
                f"{'  ✓' if ok_count else f'  ← RÓŻNICA: {sheet_tx_count - tx_total:+d}'}"
                f"   |   "
                f"{sum_icon} Kwoty: Plik {tx_sum:,.2f} PLN / Arkusz {sheet_tx_sum:,.2f} PLN"
                f"{'  ✓' if ok_sum else f'  ← RÓŻNICA: {sheet_tx_sum - tx_sum:+.2f} PLN'}"
            )
    
            missing = diff_info.get("missing", [])
            extra   = diff_info.get("extra", [])
            reconciled = diff_info.get("reconciled", False)
            label = "Naprawiono automatycznie" if reconciled else "Wykryto"
            if missing:
                st.info(f"ℹ️ **{label} — brakujące** ({len(missing)} poz.) dodane do NIEZNANE:")
                st.dataframe(
                    [{"Kwota": tx["kwota"], "Data KS": tx["data_ks"],
                      "Kontrahent": tx["kontrahent"].split("|")[0],
                      "Nr rachunku": tx["nr_rachunku"], "Tytuł": tx["tytul"][:60]}
                     for tx in missing],
                    use_container_width=True,
                )
            if extra:
                st.info(f"ℹ️ **{label} — nadmiarowe** ({len(extra)} poz.) usunięte z arkusza:")
                rows_e = []
                for r in extra:
                    try:
                        kwota_e = float(re.sub(r"[^\d,.\-]", "", str(r[5])).replace(",", ".")) if len(r) > 5 else 0.0
                    except (ValueError, TypeError):
                        kwota_e = 0.0
                    rows_e.append({"Kwota": kwota_e, "Data KS": r[6] if len(r) > 6 else "",
                                   "Kontrahent": r[4] if len(r) > 4 else "",
                                   "Nr rachunku": r[11] if len(r) > 11 else "",
                                   "Tytuł": str(r[7])[:60] if len(r) > 7 else ""})
                st.dataframe(rows_e, use_container_width=True)
            if not ok_count or not ok_sum:
                st.warning("Nadal są różnice po reconcile — sprawdź arkusz ręcznie.")
    
        with left_col:
            with st.container(border=True):
                st.markdown("#### Szablon miesiaca")
                btn_szablon = st.button(
                    "Utwórz szablon miesiaca",
                    use_container_width=True,
                    type="primary",
                )
    
            with st.container(border=True):
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
                btn_sort_kosztowe = st.button(
                    "Sortuj i nadaj kolory — Kosztowe",
                    use_container_width=True,
                )
                ksef_zip_files = st.file_uploader(
                    "Faktury KSeF (ZIP)",
                    type=["zip"],
                    accept_multiple_files=True,
                    key="ksef_zip_uploader",
                )
                btn_upload_ksef = st.button(
                    "Wgraj faktury KSeF z zip",
                    use_container_width=True,
                )
    
        with mid_col:
            with st.container(border=True):
                st.markdown("#### Faktury sprzedazy")
                btn_sprzedaz = st.button(
                    "Tworz wstepne wiersze faktur sprzedazy",
                    use_container_width=True,
                    type="primary",
                )
                btn_generuj_pdf = st.button(
                    "Generuj faktury sprzedazy PDF",
                    use_container_width=True,
                )
                btn_sprawdz_sprzedaz = st.button(
                    "Sprawdz stan faktur sprzedazy",
                    use_container_width=True,
                )
                btn_kolory_sprzedaz = st.button(
                    "Nadaj kolory — Sprzedaz",
                    use_container_width=True,
                )
    
        with right_col:
            with st.container(border=True):
                st.markdown("#### Parowanie")
                btn_paruj = st.button(
                    "Paruj wyciag bankowy z arkuszem",
                    use_container_width=True,
                    type="primary",
                )
                btn_status_parowania = st.button(
                    "Status parowania",
                    use_container_width=True,
                )
                btn_refresh_kpkw = st.button(
                    "Odśwież KP / KW",
                    use_container_width=True,
                )
                btn_show_kpkw = st.button(
                    "Pokaż KP / KW",
                    use_container_width=True,
                )
                btn_sortuj_inne_rk = st.button(
                    "Sortuj Inne RK oraz Nieznane",
                    use_container_width=True,
                )
                btn_usun_puste = st.button(
                    "Usuń puste wiersze",
                    use_container_width=True,
                )
                btn_podsumowanie = st.button(
                    "Dodaj podsumowanie segmentów",
                    use_container_width=True,
                )

        with extra_col:
            with st.container(border=True):
                st.markdown("#### Inne funkcje")
                _SEG_OPTIONS = {
                    "Faktury kosztowe":            SEP_KOSZTOWE,
                    "Faktury sprzedazy":           SEP_SPRZEDAZ,
                    "Wlasciciele i spoldzielnie":  SEP_WLASC,
                    "Inne raporty kasowe":         SEP_INNE_RK,
                    "Nieznane / niesparowane":     SEP_NIEZNANE,
                }
                sel_segment_label = st.selectbox(
                    "Segment",
                    options=list(_SEG_OPTIONS.keys()),
                    key="sel_segment_puste",
                )
                n_puste = st.number_input(
                    "Liczba pustych wierszy",
                    min_value=1, max_value=100, value=5, step=1,
                    key="n_puste_wiersze",
                )
                btn_dodaj_puste = st.button(
                    "Dodaj puste wiersze do segmentu",
                    use_container_width=True,
                )

# ----------------------------------------------------------------
# AKCJA: Dodaj puste wiersze do segmentu
# ----------------------------------------------------------------
if btn_dodaj_puste:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        _seg_options = {
            "Faktury kosztowe":            SEP_KOSZTOWE,
            "Faktury sprzedazy":           SEP_SPRZEDAZ,
            "Wlasciciele i spoldzielnie":  SEP_WLASC,
            "Inne raporty kasowe":         SEP_INNE_RK,
            "Nieznane / niesparowane":     SEP_NIEZNANE,
        }
        _sep = _seg_options.get(st.session_state.get("sel_segment_puste", ""), SEP_KOSZTOWE)
        _n   = int(st.session_state.get("n_puste_wiersze", 5))
        try:
            creds = get_credentials()
            worksheet = get_or_create_worksheet(
                gspread.authorize(creds).open_by_key(SPREADSHEET_ID), name
            )
            with st.spinner(f"Dodaje {_n} pustych wierszy..."):
                add_empty_rows_to_segment(worksheet, _sep, _n)
            st.success(f"Dodano {_n} pustych wierszy do segmentu '{st.session_state.get('sel_segment_puste')}'.")
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Sortuj Inne RK
# ----------------------------------------------------------------
if btn_sortuj_inne_rk:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        st.session_state["confirm_sortuj"] = subfolder_name.strip()

if st.session_state.get("confirm_sortuj"):
    _confirm_name = st.session_state["confirm_sortuj"]
    st.warning(
        f"⚠️ Czy na pewno chcesz posortować Inne RK oraz Nieznane w arkuszu **{_confirm_name}**?"
    )
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("✅ Tak, sortuj", key="confirm_sortuj_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_sortuj", None)
            st.session_state["run_sortuj"] = _confirm_name
            st.rerun()
    with _c2:
        if st.button("❌ Anuluj", key="confirm_sortuj_nie", use_container_width=True):
            st.session_state.pop("confirm_sortuj", None)
            st.rerun()

if st.session_state.get("run_sortuj"):
    name = st.session_state.pop("run_sortuj")
    try:
        creds  = get_credentials()
        client = gspread.authorize(creds)
        worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(name)
        with st.spinner("Sortuję Inne RK oraz Nieznane..."):
            n_inne, n_niezn = sort_inne_rk_nieznane(worksheet)
        st.success(
            f"Posortowano: Inne RK — {n_inne} wierszy, "
            f"Nieznane — {n_niezn} wierszy (kotwice status=3 zachowane)."
        )
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Arkusz '{name}' nie istnieje.")
    except Exception as e:
        st.error(f"Wystąpił błąd: {e}")

# ----------------------------------------------------------------
# AKCJA: Usuń puste wiersze
# ----------------------------------------------------------------
if btn_usun_puste:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds     = get_credentials()
            client    = gspread.authorize(creds)
            worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(name)
            with st.spinner("Usuwam puste wiersze..."):
                sections = read_all_sections(worksheet)
                rebuild_sheet(worksheet, sections, blank_rows={})
            st.success("Puste wiersze usunięte.")
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"Arkusz '{name}' nie istnieje.")
        except Exception as e:
            st.error(f"Wystąpił błąd: {e}")

# ----------------------------------------------------------------
# AKCJA: Dodaj podsumowanie segmentów
# ----------------------------------------------------------------
if btn_podsumowanie:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds  = get_credentials()
            client = gspread.authorize(creds)
            worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(name)
            drive_service = build("drive", "v3", credentials=creds)
            with st.spinner("Dodaję podsumowanie..."):
                diag = add_section_summary(worksheet, service=drive_service, subfolder_name=name)
            st.success("Podsumowanie segmentów dodane na dole arkusza.")
            if diag is not None:
                missing   = diag.get("missing", [])
                duplicates = diag.get("duplicates", [])
                if not missing and not duplicates:
                    st.success("✅ Wszystkie TX z listy operacji są w arkuszu (bez duplikatów).")
                if missing:
                    st.warning(f"⚠️ Brak w arkuszu — {len(missing)} TX z listy operacji:")
                    st.dataframe(
                        [{"Kwota": tx["kwota"], "Data KS": tx["data_ks"],
                          "Kontrahent": tx["kontrahent"], "Tytuł": tx["tytul"][:60]}
                         for tx in missing],
                        use_container_width=True,
                    )
                if duplicates:
                    st.warning(f"⚠️ Duplikaty w arkuszu — {len(duplicates)} TX z listy operacji:")
                    st.dataframe(
                        [{"Kwota": tx["kwota"], "Data KS": tx["data_ks"],
                          "Kontrahent": tx["kontrahent"], "Tytuł": tx["tytul"][:60],
                          "Status": "DUPLIKAT"}
                         for tx in duplicates],
                        use_container_width=True,
                    )
        except gspread.exceptions.WorksheetNotFound:
            st.error(f"Arkusz '{name}' nie istnieje.")
        except Exception as e:
            st.error(f"Wystąpił błąd: {e}")

# ----------------------------------------------------------------
# AKCJA: Utwórz szablon miesiaca
# ----------------------------------------------------------------
if btn_szablon:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu (np. 052026).")
    else:
        st.session_state["confirm_szablon"] = subfolder_name.strip()

if st.session_state.get("confirm_szablon"):
    _confirm_name = st.session_state["confirm_szablon"]
    st.warning(
        f"⚠️ Czy na pewno chcesz utworzyć szablon miesiąca **{_confirm_name}**?"
    )
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("✅ Tak, utwórz", key="confirm_szablon_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_szablon", None)
            st.session_state["run_szablon"] = _confirm_name
            st.rerun()
    with _c2:
        if st.button("❌ Anuluj", key="confirm_szablon_nie", use_container_width=True):
            st.session_state.pop("confirm_szablon", None)
            st.rerun()

if st.session_state.get("run_szablon"):
    name = st.session_state.pop("run_szablon")
    try:
        creds  = get_credentials()
        client = gspread.authorize(creds)
        sp     = client.open_by_key(SPREADSHEET_ID)
        with st.spinner(f"Tworzę szablon dla {name}..."):
            status, added = create_month_template(sp, name)
        if status == "created":
            st.success(f"Szablon {name} utworzony. Dodano {len(added)} sekcji.")
        elif status == "exists_partial":
            names = ", ".join(added)
            st.warning(f"{name} już istnieje. Dodano brakujące sekcje: {names}")
        else:
            st.info(f"{name} już istnieje — wszystkie sekcje są kompletne.")
    except Exception as e:
        st.error(f"Wystąpił błąd: {e}")

# ----------------------------------------------------------------
# AKCJA: Search Drive
# ----------------------------------------------------------------
if btn_search:
    q = search_query.strip()
    if not q:
        st.warning("Wpisz frazę do wyszukania.")
    else:
        try:
            creds        = get_credentials()
            drv          = build("drive", "v3", credentials=creds)
            with st.spinner(f"Szukam '{q}' ({search_type.lower()})..."):
                results = search_drive_items(drv, q, search_type)
            if results:
                import pandas as pd
                st.markdown(f"**Wyniki dla '{q}' ({len(results)} {search_type.lower()}):**")
                st.dataframe(
                    pd.DataFrame(results),
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Link": st.column_config.LinkColumn("Link", display_text="otwórz", width="small"),
                    },
                )
            else:
                st.info(f"Brak wynikow dla '{q}' ({search_type.lower()}).")
        except Exception as e:
            st.error(f"Błąd wyszukiwania: {e}")

# ----------------------------------------------------------------
# AKCJA: Search Sheets
# ----------------------------------------------------------------
if btn_sh_search:
    # Zbierz tagi z checkboxów (używane gdy brak tekstu)
    _cb_tags = []
    if sh_tag_kos:      _cb_tags.append("kos")
    if sh_tag_prz:      _cb_tags.append("prz_naj")
    if sh_tag_wla:      _cb_tags.append("wla")
    if sh_tag_nieznany: _cb_tags.append("nieznany")
    if sh_tag_roz:      _cb_tags.append("roz")
    if sh_tag_rk_kp:    _cb_tags.append("rk_kp")
    if sh_tag_rk_kw:    _cb_tags.append("rk_kw")
    if sh_tag_pr_in:    _cb_tags.append("pr_in")
    if sh_tag_pr_out:   _cb_tags.append("pr_out")
    if sh_tag_depo:     _cb_tags.append("depo")

    # Parsuj pole tekstowe
    raw_q = sh_query.strip()
    text_q, query_tags, _ = _parse_sh_query(raw_q) if raw_q else ("", [], "OR")

    # Gdy tekst: szukaj tylko po tekście (tagi będą UI-filtrem)
    # Gdy brak tekstu: szukaj po tagach z checkboxów jako baza
    if text_q:
        search_tags, search_mode = None, "OR"
    else:
        search_tags = _cb_tags if _cb_tags else None
        search_mode = "OR"  # dla bazy zawsze OR (AND to UI-filtr)

    _label = raw_q if raw_q else "/".join(_cb_tags)
    if not text_q and not _cb_tags:
        st.warning("Wpisz frazę lub zaznacz przynajmniej jeden tag.")
    else:
        try:
            creds  = get_credentials()
            client = gspread.authorize(creds)
            sp     = client.open_by_key(SPREADSHEET_ID)
            sheet_filter = "" if sh_tab_selected == "Wszystkie" else sh_tab_selected
            with st.spinner(f"Szukam '{_label}' w arkuszu..."):
                results, sheet_names = search_sheet_rows(
                    sp, text_q, sheet_filter, tags=search_tags, mode=search_mode
                )
            st.session_state["sheet_tab_names"] = sorted(sheet_names)
            st.session_state["sh_results"] = {"rows": results, "label": _label}
            # Pre-select checkboxów z *-składni
            _tag_to_key = {
                "kos": "sh_tag_kos", "prz_naj": "sh_tag_prz", "wla": "sh_tag_wla",
                "nieznany": "sh_tag_nieznany", "roz": "sh_tag_roz",
                "rk_kp": "sh_tag_rk_kp", "rk_kw": "sh_tag_rk_kw",
                "pr_in": "sh_tag_pr_in", "pr_out": "sh_tag_pr_out", "depo": "sh_tag_depo",
            }
            for _qt in query_tags:
                if _qt in _tag_to_key:
                    st.session_state[_tag_to_key[_qt]] = True
            st.rerun()
        except Exception as e:
            st.error(f"Błąd wyszukiwania w arkuszu: {e}")

# ----------------------------------------------------------------
# AKCJA: Paruj wyciag bankowy
# ----------------------------------------------------------------
if btn_paruj:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed parowaniem.")
    else:
        st.session_state["confirm_paruj"] = subfolder_name.strip()
        st.rerun()

if st.session_state.get("confirm_paruj"):
    _confirm_name = st.session_state["confirm_paruj"]
    st.warning(
        f"⚠️ Czy na pewno chcesz parować wyciąg bankowy z arkuszem **{_confirm_name}**?"
    )
    _col_tak, _col_nie = st.columns(2)
    with _col_tak:
        if st.button("✅ Tak, paruj", key="confirm_paruj_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_paruj", None)
            st.session_state["run_paruj"] = _confirm_name
            st.rerun()
    with _col_nie:
        if st.button("❌ Anuluj", key="confirm_paruj_nie", use_container_width=True):
            st.session_state.pop("confirm_paruj", None)
            st.rerun()

if st.session_state.get("run_paruj"):
    name = st.session_state.pop("run_paruj")
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
                (sparowane, niesparowane, fioletowe, pomaranczowe,
                 tx_total, tx_sum, sheet_tx_count, sheet_tx_sum,
                 diff_info) = sync_parowanie(worksheet, transactions)

            _dialog_wynik_parowania({
                "sparowane":      sparowane,
                "niesparowane":   niesparowane,
                "fioletowe":      fioletowe,
                "pomaranczowe":   pomaranczowe,
                "tx_total":       tx_total,
                "tx_sum":         tx_sum,
                "sheet_tx_count": sheet_tx_count,
                "sheet_tx_sum":   sheet_tx_sum,
                "diff_info":      diff_info,
                "subfolder_name": name,
            })

    except Exception as e:
        st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Odswierz KP / KW
# ----------------------------------------------------------------
if btn_refresh_kpkw:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        st.session_state["confirm_kpkw"] = subfolder_name.strip()

if st.session_state.get("confirm_kpkw"):
    _confirm_name = st.session_state["confirm_kpkw"]
    st.warning(
        f"⚠️ Czy na pewno chcesz odświeżyć KP / KW dla arkusza **{_confirm_name}**?"
    )
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("✅ Tak, odśwież", key="confirm_kpkw_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_kpkw", None)
            st.session_state["run_kpkw"] = _confirm_name
            st.rerun()
    with _c2:
        if st.button("❌ Anuluj", key="confirm_kpkw_nie", use_container_width=True):
            st.session_state.pop("confirm_kpkw", None)
            st.rerun()

if st.session_state.get("run_kpkw"):
    name = st.session_state.pop("run_kpkw")
    try:
        creds = get_credentials()
        client = gspread.authorize(creds)
        with st.spinner("Odświeżam KP / KW..."):
            worksheet = get_or_create_worksheet(
                client.open_by_key(SPREADSHEET_ID), name
            )
            sections_kpkw = read_all_sections(worksheet)
            refresh_kp_kw(client.open_by_key(SPREADSHEET_ID), name, sections_kpkw)
        st.success(f"KP / KW zaktualizowane dla {name}.")
    except Exception as e:
        st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Pokaz KP / KW (podglad bez zapisu)
# ----------------------------------------------------------------
if btn_show_kpkw:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            client = gspread.authorize(creds)
            with st.spinner("Czytam arkusz..."):
                worksheet = get_or_create_worksheet(
                    client.open_by_key(SPREADSHEET_ID), name
                )
                sections_podglad = read_all_sections(worksheet)
            html = preview_kp_kw_html(name, sections_podglad)
            _dialog_podglad_kp_kw(html)
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
# AKCJA: Sortuj i nadaj kolory — Kosztowe
# ----------------------------------------------------------------
if btn_sort_kosztowe:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            worksheet = get_or_create_worksheet(
                gspread.authorize(creds).open_by_key(SPREADSHEET_ID), name
            )
            with st.spinner("Sortuję i koloruję Kosztowe..."):
                n = sort_kosztowe(worksheet)
            st.success(f"Posortowano {n} wierszy Kosztowych (rk_kw → rk_kp na dole).")
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

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
                subfolder    = find_subfolder(drive_service, FAKTURY_KOSZTOWE_ID, f"{name} {FAKTURY_KOSZTOWE_SUFFIX}")
                if subfolder:
                    drive_files = list_pdfs_from_drive(drive_service, subfolder["id"], include_images=True)
                    ksef_sub = find_subfolder(drive_service, subfolder["id"], f"ksef{name}")
                    if ksef_sub:
                        drive_files += list_pdfs_from_drive(drive_service, ksef_sub["id"], include_images=True)
                    drive_file_names = [f["name"] for f in drive_files]
                else:
                    drive_file_names = []
                sheet_counts = count_kosztowe_statuses(creds, SPREADSHEET_ID, name)
                only_drive, only_sheets = diff_kosztowe(creds, SPREADSHEET_ID, name, drive_file_names)
            _dialog_kosztowe_status({
                "name": name,
                "subfolder_found": bool(subfolder),
                "drive_file_names": drive_file_names,
                "sheet_counts": sheet_counts,
                "only_drive": only_drive,
                "only_sheets": only_sheets,
            })
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Zaczytaj faktury kosztowe
# ----------------------------------------------------------------
if btn_czytaj:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed uruchomieniem.")
    else:
        st.session_state["confirm_czytaj"] = subfolder_name.strip()

if st.session_state.get("confirm_czytaj"):
    _confirm_name = st.session_state["confirm_czytaj"]
    st.warning(
        f"⚠️ Czy na pewno chcesz zaczytać faktury kosztowe dla **{_confirm_name}**? "
        f"Istniejące wiersze bez statusu 1 zostaną nadpisane."
    )
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("✅ Tak, zaczytaj", key="confirm_czytaj_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_czytaj", None)
            st.session_state["run_czytaj"] = _confirm_name
            st.rerun()
    with _c2:
        if st.button("❌ Anuluj", key="confirm_czytaj_nie", use_container_width=True):
            st.session_state.pop("confirm_czytaj", None)
            st.rerun()

if st.session_state.get("run_czytaj"):
    name = st.session_state.pop("run_czytaj")
    try:
        creds = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)

        with st.spinner(f"Szukam podfolderu '{name}'..."):
            subfolder = find_subfolder(drive_service, FAKTURY_KOSZTOWE_ID, f"{name} {FAKTURY_KOSZTOWE_SUFFIX}")

        if subfolder is None:
            st.error(f"Nie znaleziono podfolderu '{name}' w Faktury-kosztowe.")
        else:
            with st.spinner("Pobieram liste plikow PDF..."):
                files = list_pdfs_from_drive(drive_service, subfolder["id"])
                ksef_sub = find_subfolder(drive_service, subfolder["id"], f"ksef{name}")
                ksef_files = list_pdfs_from_drive(drive_service, ksef_sub["id"]) if ksef_sub else []
                all_files = files + ksef_files

            if not all_files:
                st.warning(f"Brak plikow PDF w podfolderze '{name}'.")
            else:
                if ksef_sub:
                    st.info(f"Znaleziono podfolder ksef{name} — dołączono {len(ksef_files)} faktur KSeF.")
                progress = st.progress(0, text="Analizuje faktury...")
                ksef_ids   = {f["id"] for f in ksef_files}
                files_data = []
                for i, f in enumerate(all_files):
                    progress.progress((i + 1) / len(all_files), text=f"Analizuje: {f['name']}")
                    brutto = extract_gross_amount(download_pdf(drive_service, f["id"]))
                    item   = {"key": f["name"], "brutto": brutto}
                    if f["id"] in ksef_ids:
                        item["status"] = "1"
                    files_data.append(item)
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
# AKCJA: Wgraj faktury KSeF z zip
# ----------------------------------------------------------------
if btn_upload_ksef:
    if not subfolder_name.strip():
        st.error("Wpisz nazwę podfolderu.")
    elif not ksef_zip_files:
        st.error("Dodaj przynajmniej jeden plik ZIP.")
    else:
        name       = subfolder_name.strip()
        user_drive = _get_user_drive_service()
        if user_drive is None:
            st.error("Brak OAuth credentials. Skonfiguruj [google_drive_oauth] w Secrets.")
        else:
            try:
                creds         = get_credentials()
                drive_service = build("drive", "v3", credentials=creds)
                with st.spinner(f"Szukam folderu miesiąca '{name}'..."):
                    month_folder = find_subfolder(
                        drive_service, FAKTURY_KOSZTOWE_ID,
                        f"{name} {FAKTURY_KOSZTOWE_SUFFIX}"
                    )
                if month_folder is None:
                    st.error(f"Nie znaleziono folderu miesiąca '{name}' w Faktury-kosztowe.")
                else:
                    all_uploaded, all_skipped, all_renamed, all_errors = [], [], [], []
                    for zf in ksef_zip_files:
                        with st.spinner(f"Przetwarzam {zf.name}..."):
                            up, sk, ren, err = upload_ksef_from_zip_bytes(
                                zf.read(), drive_service, user_drive, month_folder["id"]
                            )
                        all_uploaded.extend(up)
                        all_skipped.extend(sk)
                        all_renamed.extend(ren)
                        all_errors.extend(err)
                    st.success(
                        f"Wgrano: {len(all_uploaded)} | "
                        f"Przemianowano: {len(all_renamed)} | "
                        f"Pominięto (duplikaty): {len(all_skipped)} | "
                        f"Błędy: {len(all_errors)}"
                    )
                    if all_renamed:
                        st.info("Przemianowano (stara → nowa konwencja):\n" +
                                "\n".join(f"• {n}" for n in all_renamed))
                    if all_skipped:
                        st.info("Pominięto (już w folderze):\n" +
                                "\n".join(f"• {n}" for n in all_skipped))
                    for e in all_errors:
                        st.error(e)
            except Exception as e:
                st.error(f"Wystąpił błąd: {e}")

# ----------------------------------------------------------------
# AKCJA: Tworz wiersze faktur sprzedazy najemcom
# ----------------------------------------------------------------
if btn_sprzedaz:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed uruchomieniem.")
    else:
        st.session_state["confirm_sprzedaz"] = subfolder_name.strip()

if st.session_state.get("confirm_sprzedaz"):
    _confirm_name = st.session_state["confirm_sprzedaz"]
    st.warning(
        f"⚠️ Czy na pewno chcesz utworzyć wstępne wiersze faktur sprzedaży dla **{_confirm_name}**? "
        f"Istniejące wiersze bez statusu 1 zostaną nadpisane."
    )
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("✅ Tak, utwórz", key="confirm_sprzedaz_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_sprzedaz", None)
            st.session_state["run_sprzedaz"] = _confirm_name
            st.rerun()
    with _c2:
        if st.button("❌ Anuluj", key="confirm_sprzedaz_nie", use_container_width=True):
            st.session_state.pop("confirm_sprzedaz", None)
            st.rerun()

if st.session_state.get("run_sprzedaz"):
    name = st.session_state.pop("run_sprzedaz")
    try:
        creds = get_credentials()

        with st.spinner("Czytam najemcow z arkusza Abido najemcy..."):
            tenants_data = read_najemcy_for_invoices(creds)

        if not tenants_data:
            st.warning("Brak najemcow ze Status=1 w arkuszu Abido najemcy.")
        else:
            cur_month, cur_year = int(name[:2]), int(name[2:])
            def _sort_key(t):
                d = _parse_contract_start(t.get("_dates", ""))
                mid = d is not None and d.day > 1 and d.month == cur_month and d.year == cur_year
                return (1 if mid else 0, d.day if mid and d else 0)
            tenants_data.sort(key=_sort_key)

            with st.spinner("Zapisuje do Google Sheets..."):
                client = gspread.authorize(creds)
                worksheet = get_or_create_worksheet(
                    client.open_by_key(SPREADSHEET_ID), name
                )
                skipped, added = sync_sprzedaz(worksheet, tenants_data)

            st.session_state["msg_sprzedaz"] = {
                "text": f"Gotowe! Dodano: {added} najemcow | Zachowano (C=1): {skipped}",
                "tenants": [{"Najemca": t["key"], "Kwota": t["brutto"], "Adres": t["address"]}
                            for t in tenants_data],
            }
            st.rerun()
    except Exception as e:
        st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Generuj faktury sprzedazy PDF
# ----------------------------------------------------------------
if btn_generuj_pdf:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed generowaniem.")
    else:
        st.session_state["confirm_generuj"] = subfolder_name.strip()

if st.session_state.get("confirm_generuj"):
    _confirm_name = st.session_state["confirm_generuj"]
    st.warning(
        f"⚠️ Czy na pewno chcesz wygenerować faktury sprzedaży PDF dla **{_confirm_name}**?"
    )
    _c1, _c2 = st.columns(2)
    with _c1:
        if st.button("✅ Tak, generuj", key="confirm_generuj_tak", use_container_width=True, type="primary"):
            st.session_state.pop("confirm_generuj", None)
            st.session_state["run_generuj"] = _confirm_name
            st.rerun()
    with _c2:
        if st.button("❌ Anuluj", key="confirm_generuj_nie", use_container_width=True):
            st.session_state.pop("confirm_generuj", None)
            st.rerun()

if st.session_state.get("run_generuj"):
    name = st.session_state.pop("run_generuj")
    try:
        creds         = get_credentials()
        drive_service = build("drive", "v3", credentials=creds)
        client        = gspread.authorize(creds)

        with st.spinner("Otwieram arkusz..."):
            worksheet = get_or_create_worksheet(
                client.open_by_key(SPREADSHEET_ID), name
            )

        with st.spinner("Generuje faktury PDF..."):
            invoices = generate_invoice_pdfs(drive_service, worksheet, name, credentials=creds)

        if not invoices:
            st.warning("Brak wierszy w sekcji FAKTURY SPRZEDAZY NAJEMCOM.")
        else:
            def _show_download_buttons(invoices, name):
                merged_bytes = merge_pdf_bytes([b for _, b in invoices])
                st.download_button(
                    f"Pobierz scalony PDF ({len(invoices)} faktur)",
                    data=merged_bytes,
                    file_name=_merged_filename(name, invoices),
                    mime="application/pdf",
                    use_container_width=True,
                    type="primary",
                )
                with st.expander("Poszczegolne faktury PDF"):
                    for filename, pdf_bytes in invoices:
                        st.download_button(
                            filename,
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            key=f"dl_{filename}",
                        )

            user_drive = _get_user_drive_service()
            if user_drive:
                try:
                    with st.spinner("Wgrywam na Google Drive..."):
                        folder_name = upload_invoices_to_drive(user_drive, invoices, name)
                    st.success(
                        f"Gotowe! Wygenerowano {len(invoices)} faktur. "
                        f"Folder na Drive: '{folder_name}'"
                    )
                except Exception as drive_err:
                    st.warning(
                        f"Nie udalo sie wgrac na Drive ({drive_err}). "
                        f"Pobierz faktury ponizej."
                    )
                    _show_download_buttons(invoices, name)
            else:
                st.success(f"Wygenerowano {len(invoices)} faktur. Pobierz ponizej.")
                _show_download_buttons(invoices, name)
    except Exception as e:
        st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Nadaj kolory — Sprzedaz
# ----------------------------------------------------------------
if btn_kolory_sprzedaz:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds = get_credentials()
            worksheet = get_or_create_worksheet(
                gspread.authorize(creds).open_by_key(SPREADSHEET_ID), name
            )
            with st.spinner("Nadaje kolory — Sprzedaz..."):
                sections = read_all_sections(worksheet)
                rebuild_sheet(worksheet, sections, blank_rows={})
            st.success("Kolory nadane.")
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# AKCJA: Sprawdz stan faktur sprzedazy
# ----------------------------------------------------------------
if btn_sprawdz_sprzedaz:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed sprawdzeniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds         = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)
            with st.spinner("Sprawdzam..."):
                data = check_sprzedaz_status(drive_service, creds, SPREADSHEET_ID, name)
            if data is None:
                st.error(f"Brak arkusza '{name}'.")
            else:
                _dialog_sprzedaz_status({"name": name, "data": data})
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

# ----------------------------------------------------------------
# AKCJA: Wyswietl ex
# ----------------------------------------------------------------
if btn_wyswietl:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu.")
    else:
        name = subfolder_name.strip()
        try:
            creds         = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)
            client        = gspread.authorize(creds)
            spreadsheet   = client.open_by_key(SPREADSHEET_ID)
            try:
                worksheet = spreadsheet.worksheet(name)
            except gspread.exceptions.WorksheetNotFound:
                st.error(f"Arkusz '{name}' nie istnieje.")
                worksheet = None
            if worksheet:
                sections = read_all_sections(worksheet)
                st.session_state["ex_name"]     = name
                st.session_state["ex_sections"] = sections
                # Pobierz linki do plikow na Drive (kosztowe + sprzedazy)
                file_links = {}
                try:
                    def _fetch_links(folder_id):
                        resp = drive_service.files().list(
                            q=(f"'{folder_id}' in parents and "
                               "mimeType='application/pdf' and trashed=false"),
                            fields="files(name, webViewLink)",
                        ).execute()
                        for f in resp.get("files", []):
                            file_links[f["name"]] = f.get("webViewLink", "")
                    kos_folder = find_subfolder(drive_service, FAKTURY_KOSZTOWE_ID, f"{name} {FAKTURY_KOSZTOWE_SUFFIX}")
                    if kos_folder:
                        _fetch_links(kos_folder["id"])
                    sprzedaz_root = find_subfolder(drive_service, FOLDER_ID, "Faktury-sprzedazy")
                    if sprzedaz_root:
                        sprzedaz_sub = find_subfolder(
                            drive_service, sprzedaz_root["id"],
                            f"{name} {FAKTURY_SPRZEDAZY_SUFFIX}"
                        )
                        if sprzedaz_sub:
                            _fetch_links(sprzedaz_sub["id"])
                except Exception:
                    pass
                st.session_state["ex_file_links"] = file_links
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

EX_COL_NAMES = [
    "Nazwa / Plik", "Kwota brutto", "Status",
    "Klucz_Ksiegowy", "wyciag_Kontrahent", "wyciag_Kwota",
    "Data_ksiegowania", "wyciag_Tytul",
    "wyciag_Data_op", "wyciag_Rodzaj", "wyciag_Waluta",
    "wyciag_Nr_rachunku", "wyciag_Imie_Nazwisko", "Uwagi",
]
EX_READONLY = [c for c in EX_COL_NAMES if c not in ("Status", "Kwota brutto", "Uwagi")]
EX_LABELS = {
    SEP_KOSZTOWE: "Faktury kosztowe",
    SEP_SPRZEDAZ: "Faktury sprzedazy najemcom",
    SEP_WLASC:    "Wlasciciele i spoldzielnie",
    SEP_NIEZNANE: "Nieznane / niesparowane",
}

if "msg_sprzedaz" in st.session_state:
    _msg = st.session_state["msg_sprzedaz"]
    _mc, _xc = st.columns([11, 1])
    with _mc:
        st.success(_msg["text"])
    with _xc:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("✕", key="close_msg_sprzedaz", use_container_width=True):
            del st.session_state["msg_sprzedaz"]
            st.rerun()
    st.dataframe(_msg["tenants"], use_container_width=True, hide_index=True)

@st.fragment
def _show_ex():
    if "ex_sections" not in st.session_state:
        return
    ex_name     = st.session_state["ex_name"]
    ex_sections = st.session_state["ex_sections"]
    _ex_links   = st.session_state.get("ex_file_links", {})

    with st.container(border=True):
        st.markdown('<span class="abido-ex-bg"></span>', unsafe_allow_html=True)
        _ex_title_col, _ex_close_col = st.columns([8, 1])
        with _ex_title_col:
            st.markdown(f"### Arkusz: {ex_name}")
        with _ex_close_col:
            if st.button("✕ Zamknij", key="btn_ex_close", use_container_width=True):
                for _k in ("ex_name", "ex_sections", "ex_file_links"):
                    st.session_state.pop(_k, None)
                st.rerun(scope="app")

        import pandas as pd

        _all_rows = []
        for sep in SECTION_ORDER:
            for row in ex_sections.get(sep, []):
                _all_rows.append((sep, row))

        if _all_rows:
            padded = [r + [""] * (14 - len(r)) for _, r in _all_rows]
            df_all = pd.DataFrame([dict(zip(EX_COL_NAMES, r[:14])) for r in padded])
            df_all["Status"] = pd.to_numeric(df_all["Status"], errors="coerce").fillna(0).astype(int)
            df_all.insert(1, "Link", df_all["Nazwa / Plik"].map(
                lambda n: _ex_links.get(str(n), "")
            ))
            st.markdown(f"**Łącznie {len(_all_rows)} wierszy**")
            result_df = st.data_editor(
                df_all,
                key="editor_all",
                use_container_width=True,
                disabled=EX_READONLY + ["Link"],
                hide_index=True,
                height=700,
                column_config={
                    "Status": st.column_config.NumberColumn(min_value=0, max_value=2, step=1),
                    "Link": st.column_config.LinkColumn(
                        "Link", display_text="otwórz", width="small"
                    ),
                },
            )

            if st.button("Zapisz zmiany do Google Sheets", type="primary"):
                try:
                    new_sections = {sep: [] for sep in SECTION_ORDER}
                    for i, (sep, _) in enumerate(_all_rows):
                        row_vals = (
                            result_df.iloc[i]
                            .drop(labels=["Link"])
                            .astype(str)
                            .replace("nan", "")
                            .tolist()
                        )
                        new_sections[sep].append(row_vals)
                    creds = get_credentials()
                    client = gspread.authorize(creds)
                    worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(ex_name)
                    rebuild_sheet(worksheet, new_sections)
                    st.toast("Zapisano zmiany!")
                    del st.session_state["ex_sections"]
                    st.rerun(scope="app")
                except Exception as e:
                    st.error(f"Blad zapisu: {e}")
        else:
            st.caption("Arkusz jest pusty.")

_show_ex()

# ----------------------------------------------------------------
# AKCJA: Widok najemcy
# ----------------------------------------------------------------
if btn_nj_search:
    _nj_imie     = nj_imie.strip()
    _nj_nazwisko = nj_nazwisko.strip()
    _nj_mode     = nj_name_mode  # "AND" lub "OR"
    if not _nj_imie and not _nj_nazwisko:
        st.error("Wpisz przynajmniej imię lub nazwisko najemcy.")
    else:
        try:
            creds         = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)
            client        = gspread.authorize(creds)
            sp            = client.open_by_key(SPREADSHEET_ID)
            _all_tabs     = {ws.title for ws in sp.worksheets()}
            _range_tabs   = _month_tab_range(nj_od, nj_do)
            tabs          = [t for t in _range_tabs if t in _all_tabs]
            _nj_label     = " ".join(filter(None, [_nj_imie, _nj_nazwisko]))

            with st.spinner(
                f"Szukam '{_nj_label}' [{_nj_mode}] w {len(tabs)} miesiącach..."
            ):
                pdfs          = search_najemca_pdfs(drive_service, _nj_imie, _nj_nazwisko, tabs, mode=_nj_mode)
                rows          = search_najemca_sheets(sp, _nj_imie, _nj_nazwisko, tabs, mode=_nj_mode)
                drive_folders = find_drive_folders_by_name(drive_service, _nj_imie, _nj_nazwisko, mode=_nj_mode)

            st.session_state["nj_results"] = {
                "imie":          _nj_imie,
                "nazwisko":      _nj_nazwisko,
                "od":            nj_od,
                "do":            nj_do,
                "pdfs":          pdfs,
                "rows":          rows,
                "drive_folders": drive_folders,
            }
            st.rerun()
        except Exception as e:
            st.error(f"Błąd wyszukiwania najemcy: {e}")
