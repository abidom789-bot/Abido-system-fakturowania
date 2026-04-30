import io
import os
import re
import calendar
import unicodedata
import xlrd
import streamlit as st
import gspread
import pdfplumber
from datetime import date
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ----------------------------------------------------------------
# KONFIGURACJA
# ----------------------------------------------------------------
FOLDER_ID            = "1kwY6tOalKS2jnidABw6uUV23ykMj1iR2"
FAKTURY_KOSZTOWE_ID  = "12RxQDakB6y9pxURM_Z73sS0fLNQyGtm1"
MIESZKANIA_FOLDER_ID = "1mvVZN6y2vaKyWGV6SIWd7FuK38T2DHAI"
SPREADSHEET_ID       = "1oFgjTnx6JwjD6j1pvRhcfSmd-1IwP2l3nLsw-8qv8a0"
# ----------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

HEADER_ROW = [
    "Nazwa / Plik", "Kwota brutto", "Status", "Kwota_raport_kasowy",
    "Adres", "Klucz_Ksiegowy", "wyciag_Kontrahent", "wyciag_Kwota",
    "Data_ksiegowania", "wyciag_Tytul",
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
    _NUM = r"([\d ]+[,.][\d]{2})"
    patterns = [
        # "Do zapłaty" / "Pozostaje do zapłaty" / "Razem do zapłaty"
        #   — pomijamy 0,00 (Allegro: zapłacone przy zakupie)
        r"(?:razem\s+|pozostaje\s+)?do\s+zap[lł]aty[^\d]*?" + _NUM,
        # "Wartość brutto X,XX" jako osobna linia (np. Allegro)
        r"warto[śs][ćc]\s+brutto\s+" + _NUM,
        # "Należność X,XX" — np. E.ON
        r"nale[żz]no[śs][ćc]\s+" + _NUM,
        # "Razem brutto" / "Suma brutto" / "Ogółem brutto"
        r"(?:razem|suma|og[oó][lł]em)\s+brutto[^\d]*?" + _NUM,
        # "Kwota brutto: X,XX"
        r"kwota\s+brutto\s*[:\-]\s*" + _NUM,
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
                return "-" + val
        # Ostatnia szansa: linia "Razem netto VAT brutto" — ostatnia liczba w linii
        for line in tl.splitlines():
            if re.match(r"\s*(?:\d+\.\s+)?razem\b", line):
                nums = re.findall(r"\d+[,.]\d{2}", line)
                if len(nums) >= 2:
                    return "-" + nums[-1]
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
FAKTURY_SPRZEDAZY_PREFIX = "Faktury sprzedazy"

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


def upload_invoices_to_drive(user_drive_service, invoices, subfolder_name):
    """
    Wgrywa faktury PDF na Drive uzywajac user OAuth credentials.
    Tworzy folder 'Faktury sprzedazy MMRRRR' wewnatrz 'Faktury-sprzedazy'.
    """
    folder_name = f"{FAKTURY_SPRZEDAZY_PREFIX} {subfolder_name}"
    parent_id   = get_or_create_subfolder(user_drive_service, FOLDER_ID, "Faktury-sprzedazy")
    folder_id   = get_or_create_subfolder(user_drive_service, parent_id, folder_name)
    for filename, pdf_b in invoices:
        upload_file_to_drive(user_drive_service, folder_id, filename, pdf_b)
    if invoices:
        merged = merge_pdf_bytes([b for _, b in invoices])
        upload_file_to_drive(
            user_drive_service, folder_id,
            f"Fs_najemcy_{subfolder_name}.pdf", merged
        )
    return folder_name


def generate_invoice_pdfs(drive_service, worksheet, subfolder_name):
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

    # Dane folderow FVS — do odczytu daty poczatku umowy
    fvs_folders = list_fvs_folders(drive_service)
    fvs_by_name = {}
    for f in fvs_folders:
        d = parse_fvs_folder(f["name"])
        fvs_by_name[d["name"]] = d

    # Wiersze sekcji SPRZEDAZ
    sections = read_all_sections(worksheet)
    rows = sections[SEP_SPRZEDAZ]
    if not rows:
        return []

    _get_pdf_fonts()  # rejestracja czcionek przed generowaniem

    results = []  # lista (filename, pdf_bytes)

    for num, row in enumerate(rows, 1):
        name       = row[0] if len(row) > 0 else ""
        amount_str = row[1] if len(row) > 1 else ""
        address    = row[4] if len(row) > 4 else ""
        klucz      = row[6] if len(row) > 6 else ""

        if not name:
            continue

        # Kwota (zawsze dodatnia na fakturze)
        try:
            amount = abs(float(
                str(amount_str).replace(",", ".").replace(" ", "")
            ))
        except (ValueError, TypeError):
            amount = 0.0

        # Data wystawienia i ewentualna korekta proporcjonalna
        issue_date = default_issue
        fvs = fvs_by_name.get(name)
        if fvs and fvs.get("dates"):
            contract_start = _parse_contract_start(fvs["dates"])
            if (contract_start
                    and contract_start.year == year
                    and contract_start.month == month
                    and contract_start.day > 1):
                issue_date = contract_start
                days_remaining = last_day - contract_start.day + 1
                amount = round(amount * days_remaining / last_day, 2)

        payment_method = "Gotówka" if klucz == "prz_naj_rk_kp" else "Przelew"
        invoice_nr     = f"FVS {year} {month:02d} {num:02d} T"
        name_norm      = _normalize_name_for_filename(name)
        filename       = f"fvs_{year}_{month:02d}_{num:02d}_t_{name_norm}.pdf"

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
        "horizontalAlignment": "CENTER",
    })
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


def apply_sync_logic(existing_rows, new_data, has_address=False, default_status="0"):
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
            addr  = item.get("address", "") if has_address else ""
            dates = item.get("dates",   "") if has_address else ""
            result.append([key, item.get("brutto", ""), default_status, "", addr, dates])
    # Zachowaj zweryfikowane wiersze ktorych plik zostal usuniety z Drive
    for key, row in verified.items():
        if key not in new_keys:
            result.append(row)
    return result, len(verified), len(new_data) - len(verified)


def sync_kosztowe(worksheet, files_data):
    sections                  = read_all_sections(worksheet)
    new_rows, skipped, added  = apply_sync_logic(sections[SEP_KOSZTOWE], files_data)
    # Faktury 'cash' na koniec sekcji kosztowej
    new_rows.sort(key=lambda r: _is_cash(r[0] if r else ""))
    sections[SEP_KOSZTOWE]    = new_rows
    rebuild_sheet(worksheet, sections)
    return skipped, added


def sync_sprzedaz(worksheet, tenants_data):
    sections                  = read_all_sections(worksheet)
    new_rows, skipped, added  = apply_sync_logic(
        sections[SEP_SPRZEDAZ], tenants_data, has_address=True, default_status="1"
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
            has_pair = len(row) > 7 and str(row[7]).strip() != ""
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
    row = list(existing_row) + [""] * max(0, 17 - len(existing_row))
    row[3]  = ""   # raport kasowy — puste dla transakcji bankowych
    row[6]  = klucz
    row[7]  = tx["kontrahent"].split("|")[0]
    row[8]  = tx["kwota"]
    row[9]  = tx["data_ks"]
    row[10] = tx["tytul"][:100]
    row[11] = tx["data_op"]
    row[12] = tx["rodzaj"]
    row[13] = tx["waluta"]
    row[14] = tx["nr_rachunku"]
    row[15] = _extract_name_from_tx(tx)
    row[16] = uwagi
    return row


def _build_unmatched_row(tx):
    """Buduje wiersz dla niesparowanej transakcji z wyciagu (A i B puste)."""
    klucz = "nieznany_out" if tx["kwota"] < 0 else "nieznany_in"
    return [
        "", "", "", "", "", "",   # A=nazwa, B=kwota, C=status, D=raport_kasowy, E=adres, F=data_umowy
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
            klucz = assign_klucz_ksiegowy(sep, tx, row[1] if len(row) > 1 else "", row[0] if row else "")
            sections[sep][row_idx] = _build_paired_row(row, tx, klucz)
        else:
            # Brak pary — klucz ksiegowy + wyczysc kolumny wyciagu
            klucz = assign_klucz_ksiegowy(sep, None, row[1] if len(row) > 1 else "", row[0] if row else "")
            r = list(row) + [""] * max(0, 17 - len(row))
            r[6] = klucz
            for col in range(7, 17):
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
        fields="files(id, name, parents)",
        pageSize=50,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()

    cache = {}
    results = []
    for item in resp.get("files", []):
        name    = item["name"]
        parents = item.get("parents", [])
        if search_type == "Pliki":
            path = _get_item_path(service, parents[0], cache) if parents else "/"
        else:
            path = _get_item_path(service, item["id"], cache)
        results.append({"Nazwa": name, "Ścieżka": path})

    return results


def search_sheet_rows(spreadsheet, query_text, sheet_filter=None):
    """
    Szuka query_text we wszystkich wierszach arkusza Google Sheets.
    sheet_filter: None lub '' = wszystkie zakladki; inaczej = konkretna zakladka.
    Zwraca (wyniki: list[dict], nazwy_zakladek: list[str]).
    """
    q = query_text.strip().lower()
    all_worksheets = spreadsheet.worksheets()
    sheet_names = [ws.title for ws in all_worksheets]

    if sheet_filter:
        worksheets = [ws for ws in all_worksheets if ws.title == sheet_filter]
    else:
        worksheets = all_worksheets

    results = []
    for ws in worksheets:
        for row in ws.get_all_values():
            if not any(cell for cell in row):
                continue
            if row and _match_separator(row[0]):
                continue
            if row and row[0] == HEADER_ROW[0]:   # naglowek
                continue
            if any(q in str(cell).lower() for cell in row):
                padded = row + [""] * max(0, 16 - len(row))
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


def search_najemca_sheets(spreadsheet, imie, nazwisko, tabs):
    """
    Szuka wierszy pasujących do najemcy w zakładkach arkusza.
    Sprawdza kolumnę A (Nazwa/Plik), G (Klucz_Ksiegowy) i P (wyciag_Imie_Nazwisko).
    Zwraca listę słowników.
    """
    imie_n = _normalize_name_for_filename(imie)
    nazw_n = _normalize_name_for_filename(nazwisko)
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
            padded = row + [""] * max(0, 17 - len(row))
            col_a    = _normalize_name_for_filename(padded[0])
            col_klucz = _normalize_name_for_filename(padded[6])
            col_wyc  = _normalize_name_for_filename(padded[15])
            hit = (
                (imie_n in col_a     and nazw_n in col_a)     or
                (imie_n in col_klucz and nazw_n in col_klucz) or
                (imie_n in col_wyc   and nazw_n in col_wyc)
            )
            if hit:
                results.append({
                    "Zakladka":      tab,
                    "Nazwa":         padded[0],
                    "Kwota":         padded[1],
                    "Status":        padded[2],
                    "Raport_kasowy": padded[3],
                    "Klucz":         padded[6],
                    "wyciag_Kwota":  padded[8],
                    "Data":          padded[9],
                })
    return results


def search_najemca_pdfs(service, imie, nazwisko, tabs):
    """
    Szuka plików PDF pasujących do najemcy w podfolderach Drive (Faktury sprzedazy MMYYYY).
    Zwraca listę słowników z nazwą pliku i linkiem do Drive.
    """
    imie_n = _normalize_name_for_filename(imie)
    nazw_n = _normalize_name_for_filename(nazwisko)
    sprzedaz_folder = find_subfolder(service, FOLDER_ID, "Faktury-sprzedazy")
    if not sprzedaz_folder:
        return []

    results = []
    for tab in tabs:
        folder_name  = f"{FAKTURY_SPRZEDAZY_PREFIX} {tab}"
        month_folder = find_subfolder(service, sprzedaz_folder["id"], folder_name)
        if not month_folder:
            continue
        for pdf in list_pdfs_from_drive(service, month_folder["id"]):
            fname = _normalize_name_for_filename(pdf["name"])
            if imie_n in fname and nazw_n in fname:
                try:
                    meta = service.files().get(
                        fileId=pdf["id"], fields="webViewLink"
                    ).execute()
                    link = meta.get("webViewLink", "")
                except Exception:
                    link = ""
                results.append({
                    "Zakladka":    tab,
                    "Nazwa pliku": pdf["name"],
                    "Link":        link,
                })
    return results


# ----------------------------------------------------------------
# INTERFEJS STREAMLIT
# ----------------------------------------------------------------

st.set_page_config(
    page_title="System Fakturowania",
    page_icon=":page_facing_up:",
    layout="wide",
)

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
</style>
""", unsafe_allow_html=True)

st.title("System Fakturowania")

# ── Search — na samej gorze ─────────────────────────────────────────
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

st.markdown("")

# ── Wyszukiwanie w Google Sheets ─────────────────────────────────────
sh_q_col, sh_tab_col, sh_btn_col = st.columns([5, 1.2, 0.5])
with sh_q_col:
    sh_query = st.text_input(
        "Szukaj w Sheets",
        placeholder="Szukaj w Google Sheets (adres, nazwa, kwota...)...",
        label_visibility="collapsed",
    )
with sh_tab_col:
    sh_tab_options = ["Wszystkie"] + st.session_state.get("sheet_tab_names", [])
    sh_tab_selected = st.selectbox(
        "Zakladka",
        sh_tab_options,
        label_visibility="collapsed",
    )
with sh_btn_col:
    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    btn_sh_search = st.button("🔍 Szukaj", use_container_width=True, key="btn_sh_search")

st.markdown("")

# ── Segment: miesiac + akcje ────────────────────────────────────────
with st.container(border=True):

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
    left_col, mid_col, right_col = st.columns(3)

    with left_col:
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

st.markdown("---")

# ================================================================
# WIDOK NAJEMCY
# ================================================================
with st.container(border=True):
    st.markdown("### Widok najemcy")

    nj_r1c1, nj_r1c2 = st.columns(2)
    with nj_r1c1:
        nj_imie = st.text_input(
            "Imię najemcy", placeholder="np. Mehdi", key="nj_imie"
        )
    with nj_r1c2:
        nj_nazwisko = st.text_input(
            "Nazwisko najemcy", placeholder="np. Edbouche", key="nj_nazwisko"
        )

    _nj_month_opts = [f"{m:02d}/{y}" for y in range(2024, 2028) for m in range(1, 13)]
    nj_dc1, nj_dc2, nj_bc = st.columns([2, 2, 1])
    with nj_dc1:
        nj_od = st.selectbox("Od", _nj_month_opts, index=15, key="nj_od")
    with nj_dc2:
        nj_do = st.selectbox("Do", _nj_month_opts, index=27, key="nj_do")
    with nj_bc:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        btn_nj_search = st.button("Szukaj", use_container_width=True, key="btn_nj_search")

    st.markdown("**Filtry transakcji:**")
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
        with st.expander(f"Faktury PDF ({len(_nj['pdfs'])})", expanded=True):
            if _nj["pdfs"]:
                _pdf_df = pd.DataFrame(_nj["pdfs"])[["Zakladka", "Nazwa pliku", "Link"]]
                st.dataframe(
                    _pdf_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={"Link": st.column_config.LinkColumn("Link do Drive")},
                )
            else:
                st.caption("Brak faktur PDF dla tego najemcy w podanym zakresie.")

        # Filtr checkboxów
        def _row_matches_filters(klucz):
            k = klucz.lower()
            if nj_prz    and k.startswith("prz_"):             return True
            if nj_kp     and "rk_kp" in k:                    return True
            if nj_kw     and "rk_kw" in k:                    return True
            if nj_pr_in  and "pr_in" in k:                    return True
            if nj_pr_out and "pr_out" in k:                   return True
            if nj_depo   and "depo" in k:                     return True
            if nj_roz    and "_roz_" in k and "depo" not in k: return True
            return False

        filtered_rows = [r for r in _nj["rows"] if _row_matches_filters(r["Klucz"])]

        # Transakcje
        with st.expander(
            f"Transakcje w arkuszu ({len(filtered_rows)} z {len(_nj['rows'])})",
            expanded=True,
        ):
            if filtered_rows:
                st.dataframe(
                    pd.DataFrame(filtered_rows),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.caption("Brak transakcji dla wybranych filtrów.")

        # Bilans
        with st.expander("Bilans", expanded=True):
            _prz_rows  = [r for r in _nj["rows"] if r["Klucz"].lower().startswith("prz_")]
            _depo_in   = [r for r in _nj["rows"]
                          if "depo" in r["Klucz"].lower()
                          and ("_in" in r["Klucz"].lower() or "_kp" in r["Klucz"].lower())]
            _depo_out  = [r for r in _nj["rows"]
                          if "depo" in r["Klucz"].lower()
                          and ("_out" in r["Klucz"].lower() or "_kw" in r["Klucz"].lower())]

            def _sum_kwota_bil(rows):
                return sum(abs(_parse_amount(r["Kwota"]) or 0.0) for r in rows)

            prz_sum      = _sum_kwota_bil(_prz_rows)
            depo_in_sum  = _sum_kwota_bil(_depo_in)
            depo_out_sum = _sum_kwota_bil(_depo_out)
            depo_saldo   = depo_in_sum - depo_out_sum

            bil_c1, bil_c2 = st.columns(2)
            with bil_c1:
                st.markdown("**Przychody z wynajmu (prz)**")
                st.metric("Faktury PDF", len(_nj["pdfs"]))
                st.metric("Wierszy prz w arkuszu", len(_prz_rows))
                st.metric(
                    "Suma kwot prz",
                    f"{prz_sum:,.2f} zł".replace(",", " "),
                )
            with bil_c2:
                st.markdown("**Kaucja (depo)**")
                st.metric("Wpłacona",  f"{depo_in_sum:,.2f} zł".replace(",", " "))
                st.metric("Zwrócona",  f"{depo_out_sum:,.2f} zł".replace(",", " "))
                _saldo_icon = "+" if depo_saldo >= 0 else "-"
                st.metric(
                    "Saldo kaucji",
                    f"{depo_saldo:,.2f} zł".replace(",", " "),
                )

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
                )
            else:
                st.info(f"Brak wynikow dla '{q}' ({search_type.lower()}).")
        except Exception as e:
            st.error(f"Błąd wyszukiwania: {e}")

# ----------------------------------------------------------------
# AKCJA: Search Sheets
# ----------------------------------------------------------------
if btn_sh_search:
    q = sh_query.strip()
    if not q:
        st.warning("Wpisz frazę do wyszukania w arkuszu.")
    else:
        try:
            creds  = get_credentials()
            client = gspread.authorize(creds)
            sp     = client.open_by_key(SPREADSHEET_ID)
            sheet_filter = "" if sh_tab_selected == "Wszystkie" else sh_tab_selected
            with st.spinner(f"Szukam '{q}' w arkuszu..."):
                results, sheet_names = search_sheet_rows(sp, q, sheet_filter)
            st.session_state["sheet_tab_names"] = sorted(sheet_names)
            if results:
                import pandas as pd
                label = sh_tab_selected if sh_tab_selected != "Wszystkie" else "wszystkich zakladkach"
                st.markdown(f"**Wyniki dla '{q}' w {label} ({len(results)} wierszy):**")
                df = pd.DataFrame(results)
                st.dataframe(df, use_container_width=True, hide_index=True)
            else:
                st.info(f"Brak wynikow dla '{q}'.")
        except Exception as e:
            st.error(f"Błąd wyszukiwania w arkuszu: {e}")

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
                cur_month, cur_year = int(name[:2]), int(name[2:])

                def _tenant_sort_key(t):
                    d = _parse_contract_start(t["dates"])
                    mid_this_month = (
                        d is not None
                        and d.day > 1
                        and d.month == cur_month
                        and d.year  == cur_year
                    )
                    return (
                        1 if mid_this_month else 0,   # mid-month bieżącego mies. → koniec
                        t.get("address", ""),          # alfabetycznie po adresie
                        d.day if mid_this_month and d else 0,  # w grupie końcowej wg dnia
                    )

                tenants.sort(key=_tenant_sort_key)
                tenants_data = [
                    {"key": t["name"], "brutto": t["price"], "address": t["address"], "dates": t["dates"]}
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

# ----------------------------------------------------------------
# AKCJA: Generuj faktury sprzedazy PDF
# ----------------------------------------------------------------
if btn_generuj_pdf:
    if not subfolder_name.strip():
        st.error("Wpisz nazwe podfolderu przed generowaniem.")
    else:
        name = subfolder_name.strip()
        try:
            creds         = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)
            client        = gspread.authorize(creds)

            with st.spinner("Otwieram arkusz..."):
                worksheet = get_or_create_worksheet(
                    client.open_by_key(SPREADSHEET_ID), name
                )

            with st.spinner("Generuje faktury PDF..."):
                invoices = generate_invoice_pdfs(drive_service, worksheet, name)

            if not invoices:
                st.warning("Brak wierszy w sekcji FAKTURY SPRZEDAZY NAJEMCOM.")
            else:
                # Próbuj wgrać na Drive przez user OAuth credentials
                user_drive = _get_user_drive_service()
                if user_drive:
                    with st.spinner("Wgrywam na Google Drive..."):
                        folder_name = upload_invoices_to_drive(user_drive, invoices, name)
                    st.success(
                        f"Gotowe! Wygenerowano {len(invoices)} faktur. "
                        f"Folder na Drive: '{folder_name}'"
                    )
                else:
                    # Fallback: przyciski download (brak google_drive_oauth w secrets)
                    st.success(f"Wygenerowano {len(invoices)} faktur. Pobierz ponizej.")
                    st.info(
                        "Aby wgrywac automatycznie na Drive, dodaj sekcje "
                        "[google_drive_oauth] do Streamlit secrets "
                        "(uruchom get_refresh_token.py — instrukcja w pliku)."
                    )
                    merged_bytes = merge_pdf_bytes([b for _, b in invoices])
                    st.download_button(
                        f"Pobierz scalony PDF ({len(invoices)} faktur)",
                        data=merged_bytes,
                        file_name=f"Fs_najemcy_{name}.pdf",
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
            creds = get_credentials()
            client = gspread.authorize(creds)
            spreadsheet = client.open_by_key(SPREADSHEET_ID)
            try:
                worksheet = spreadsheet.worksheet(name)
            except gspread.exceptions.WorksheetNotFound:
                st.error(f"Arkusz '{name}' nie istnieje.")
                worksheet = None
            if worksheet:
                sections = read_all_sections(worksheet)
                st.session_state["ex_name"]     = name
                st.session_state["ex_sections"] = sections
        except Exception as e:
            st.error(f"Wystapil blad: {e}")

EX_COL_NAMES = [
    "Nazwa / Plik", "Kwota brutto", "Status", "Kwota_raport_kasowy",
    "Adres", "Data_umowy",
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

if "ex_sections" in st.session_state:
    ex_name     = st.session_state["ex_name"]
    ex_sections = st.session_state["ex_sections"]

    with st.container(border=True):
        st.markdown(f"### Arkusz: {ex_name}")

        edited = {}
        for sep in SECTION_ORDER:
            rows  = ex_sections.get(sep, [])
            label = EX_LABELS.get(sep, sep)
            with st.expander(f"{label} ({len(rows)} wierszy)", expanded=True):
                if rows:
                    import pandas as pd
                    padded = [r + [""] * (17 - len(r)) for r in rows]
                    df = pd.DataFrame([dict(zip(EX_COL_NAMES, r[:17])) for r in padded])
                    df["Status"] = pd.to_numeric(df["Status"], errors="coerce").fillna(0).astype(int)
                    result_df = st.data_editor(
                        df,
                        key=f"editor_{sep}",
                        use_container_width=True,
                        disabled=EX_READONLY,
                        hide_index=True,
                        column_config={
                            "Status": st.column_config.NumberColumn(min_value=0, max_value=2, step=1),
                        },
                    )
                    edited[sep] = result_df
                else:
                    st.caption("(brak wierszy)")
                    edited[sep] = None

        if st.button("Zapisz zmiany do Google Sheets", type="primary"):
            try:
                new_sections = {}
                for sep in SECTION_ORDER:
                    df_ed = edited.get(sep)
                    if df_ed is not None:
                        new_sections[sep] = df_ed.astype(str).replace("nan", "").values.tolist()
                    else:
                        new_sections[sep] = []
                creds = get_credentials()
                client = gspread.authorize(creds)
                worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(ex_name)
                rebuild_sheet(worksheet, new_sections)
                st.success("Zapisano zmiany!")
                del st.session_state["ex_sections"]
            except Exception as e:
                st.error(f"Blad zapisu: {e}")

# ----------------------------------------------------------------
# AKCJA: Widok najemcy
# ----------------------------------------------------------------
if btn_nj_search:
    _nj_imie     = nj_imie.strip()
    _nj_nazwisko = nj_nazwisko.strip()
    if not _nj_imie or not _nj_nazwisko:
        st.error("Wpisz imię i nazwisko najemcy.")
    else:
        try:
            tabs          = _month_tab_range(nj_od, nj_do)
            creds         = get_credentials()
            drive_service = build("drive", "v3", credentials=creds)
            client        = gspread.authorize(creds)
            sp            = client.open_by_key(SPREADSHEET_ID)

            with st.spinner(
                f"Szukam '{_nj_imie} {_nj_nazwisko}' w {len(tabs)} miesiącach..."
            ):
                pdfs = search_najemca_pdfs(drive_service, _nj_imie, _nj_nazwisko, tabs)
                rows = search_najemca_sheets(sp, _nj_imie, _nj_nazwisko, tabs)

            st.session_state["nj_results"] = {
                "imie":     _nj_imie,
                "nazwisko": _nj_nazwisko,
                "od":       nj_od,
                "do":       nj_do,
                "pdfs":     pdfs,
                "rows":     rows,
            }
            st.rerun()
        except Exception as e:
            st.error(f"Błąd wyszukiwania najemcy: {e}")
