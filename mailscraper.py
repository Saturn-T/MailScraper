import imaplib
import email
from email.header import decode_header
import time
import os
import io
from datetime import datetime, timedelta, timezone

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from pypdf import PdfWriter, PdfReader
from bs4 import BeautifulSoup

from dotenv import load_dotenv
import os

# Laad de .env variabelen
load_dotenv()

# ─────────────────────────────────────────
# MEERDERE ACCOUNTS INLEZEN
# ─────────────────────────────────────────
BEKENDE_SERVERS = {
    "gmail.com":      "imap.gmail.com",
    "outlook.com":    "outlook.office365.com",
    "hotmail.com":    "outlook.office365.com",
    "live.com":       "outlook.office365.com",
    "yahoo.com":      "imap.mail.yahoo.com",
    "icloud.com":     "imap.mail.me.com",
    "me.com":         "imap.mail.me.com",
    "mac.com":        "imap.mail.me.com",
}

def server_raden(email_adres):
    domein = email_adres.split("@")[-1].lower()
    return BEKENDE_SERVERS.get(domein)

def accounts_inlezen():
    accounts = []
    basis_email = os.getenv("EMAIL_ADRES")
    basis_wachtwoord = os.getenv("PASSWORD")
    if basis_email and basis_wachtwoord:
        server = os.getenv("IMAP_SERVER") or server_raden(basis_email) or "imap.gmail.com"
        accounts.append({
            "naam": basis_email,
            "email": basis_email,
            "wachtwoord": basis_wachtwoord,
            "server": server,
        })
    i = 1
    while True:
        email_adres = os.getenv(f"EMAIL_ADRES_{i}")
        wachtwoord  = os.getenv(f"PASSWORD_{i}")
        if not email_adres or not wachtwoord:
            break
        server = os.getenv(f"IMAP_SERVER_{i}") or server_raden(email_adres) or "imap.gmail.com"
        accounts.append({
            "naam": email_adres,
            "email": email_adres,
            "wachtwoord": wachtwoord,
            "server": server,
        })
        i += 1
    return accounts

CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL", 30))
UITVOER_MAP    = os.getenv("UITVOER_MAP", "mail_pdfs")

# ─────────────────────────────────────────

def verbinding_maken(account):
    verbinding = imaplib.IMAP4_SSL(account["server"], 993)
    verbinding.login(account["email"], account["wachtwoord"])
    verbinding.select("INBOX")
    return verbinding

def alle_ids_ophalen(verbinding):
    _, data = verbinding.search(None, "ALL")
    if not data or not data[0]:
        return set()
    return set(data[0].split())

def header_lezen(waarde):
    if not waarde:
        return ""
    stuk, codering = decode_header(waarde)[0]
    if isinstance(stuk, bytes):
        return stuk.decode(codering or "utf-8", errors="replace")
    return stuk

def vind_grootste_bytes(obj):
    gevonden = []
    def _zoek(item):
        if isinstance(item, bytes):
            gevonden.append(item)
        elif isinstance(item, (list, tuple)):
            for sub in item:
                _zoek(sub)
    _zoek(obj)
    if not gevonden:
        return None
    return max(gevonden, key=len)

def is_reclame(bericht):
    if bericht.get("List-Unsubscribe"):
        return True
    afzender = (bericht.get("From") or "").lower()
    verdachte_woorden = [
        "newsletter", "nieuwsbrief",
        "marketing", "promo", "promotie",
        "notifications@", "notification@",
        "do-not-reply", "donotreply",
    ]
    for woord in verdachte_woorden:
        if woord in afzender:
            return True
    return False

def mail_lezen(verbinding, mail_id):
    _, data = verbinding.fetch(mail_id, "(BODY.PEEK[])")
    ruwe_mail = vind_grootste_bytes(data)
    if not ruwe_mail or len(ruwe_mail) < 10:
        print("    ⚠️  Kon maildata niet lezen (onverwachte IMAP-response)")
        return None, None, None, None, None, None

    bericht = email.message_from_bytes(ruwe_mail)
    onderwerp = header_lezen(bericht["Subject"]) or "(geen onderwerp)"
    afzender  = header_lezen(bericht["From"])    or "Onbekend"
    datum     = bericht["Date"]                  or "Onbekend"

    print(f"    Van:       {afzender}")
    print(f"    Onderwerp: {onderwerp}")
    print(f"    Datum:     {datum}")

    if is_reclame(bericht):
        print(f"    ⏭️  Overgeslagen (reclame/nieuwsbrief)")
        verbinding.store(mail_id, '+FLAGS', '\\Seen')
        return None, None, None, None, None, None

    tekst    = ""
    html     = ""
    bijlagen = []

    if bericht.is_multipart():
        for deel in bericht.walk():
            soort        = deel.get_content_type()
            bestandsnaam = deel.get_filename()
            if soort == "text/plain" and not bestandsnaam:
                ruwe_tekst = deel.get_payload(decode=True)
                codering   = deel.get_content_charset() or "utf-8"
                tekst     += ruwe_tekst.decode(codering, errors="replace")
            elif soort == "text/html" and not bestandsnaam:
                ruwe_html = deel.get_payload(decode=True)
                codering  = deel.get_content_charset() or "utf-8"
                html      += ruwe_html.decode(codering, errors="replace")
            elif bestandsnaam:
                naam = header_lezen(bestandsnaam)
                bijlagen.append({
                    "naam": naam,
                    "type": soort,
                    "data": deel.get_payload(decode=True)
                })
                print(f"    Bijlage:   {naam} ({soort})")
    else:
        ruwe_tekst = bericht.get_payload(decode=True)
        tekst      = ruwe_tekst.decode("utf-8", errors="replace")

    import re

    def html_strippen(invoer):
        invoer = re.sub(r"<br\s*/?>", "\n", invoer)
        invoer = re.sub(r"</p>|</div>", "\n", invoer)
        invoer = re.sub(r"&nbsp;", " ", invoer)
        invoer = re.sub(r"&amp;", "&", invoer)
        invoer = re.sub(r"&lt;", "<", invoer)
        invoer = re.sub(r"&gt;", ">", invoer)
        invoer = re.sub(r"<[^>]+>", "", invoer)
        invoer = re.sub(r"\n{3,}", "\n\n", invoer).strip()
        return invoer

    # Controleer of tekst eigenlijk HTML is
    if not tekst.strip() and html:
        print(f"    (HTML mail → WeasyPrint)")
    elif re.search(r"<[a-zA-Z]+[\s>]", tekst):
        # tekst/plain claimt maar stuurt HTML → gebruik dat als html
        html = tekst
        tekst = ""
        print(f"    (HTML in tekstveld → WeasyPrint)")

    return onderwerp, afzender, datum, tekst, html, bijlagen

def veilige_naam(tekst):
    for teken in r'\/:*?"<>|':
        tekst = tekst.replace(teken, "_")
    tekst = tekst.strip()
    tekst = tekst.rstrip(". ")
    tekst = tekst[:60].strip()
    tekst = tekst.rstrip(". ")
    if not tekst:
        tekst = "geen_onderwerp"
    return tekst

def html_naar_alineas(html):
    """
    Zet HTML om naar een gestructureerde lijst van alinea-objecten voor ReportLab.
    Behoudt: vetgedrukt, cursief, kopjes, lijsten.
    Voorkomt dubbele regels — HTML-mails bevatten dezelfde tekst vaak 2-3x
    (desktop/mobiel/fallback versie). We onthouden welke teksten al gezien zijn.
    """
    soup = BeautifulSoup(html, "html.parser")

    # Verwijder onzichtbare elementen en Microsoft conditional comments
    for tag in soup.find_all(["script", "style", "head", "meta", "link"]):
        tag.decompose()

    # Verwijder verborgen elementen (display:none) — die bevatten de fallback-kopieën
    for tag in soup.find_all(style=True):
        stijl = tag.get("style", "").lower().replace(" ", "")
        if "display:none" in stijl or "visibility:hidden" in stijl or "mso-hide:all" in stijl:
            tag.decompose()

    # Verwijder elementen met aria-hidden of hidden attribuut
    for tag in soup.find_all(attrs={"aria-hidden": "true"}):
        tag.decompose()
    for tag in soup.find_all(attrs={"hidden": True}):
        tag.decompose()

    alineas = []
    geziene_teksten = set()  # bijhouden wat we al hebben toegevoegd

    def esc(t):
        return (t or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def tag_naar_tekst(tag):
        if isinstance(tag, str):
            return esc(tag)
        naam = tag.name or ""
        inhoud_delen = [tag_naar_tekst(k) for k in tag.children]
        tekst = "".join(inhoud_delen).strip()
        if not tekst:
            return ""
        if naam in ("b", "strong"):
            return f"<b>{tekst}</b>"
        if naam in ("i", "em"):
            return f"<i>{tekst}</i>"
        if naam == "a":
            href = tag.get("href", "").strip()
            # Alleen echte http-links klikbaar maken, geen mailto/javascript/lege links
            if href.startswith("http") and tekst:
                return f'<a href="{href}" color="#3498DB">{tekst}</a>'
            return tekst
        if naam == "br":
            return "<br/>"
        return tekst

    stijlen = getSampleStyleSheet()
    kop_st  = ParagraphStyle("KH", parent=stijlen["Normal"],
                  fontSize=12, fontName="Helvetica-Bold",
                  textColor=colors.HexColor("#2C3E50"), spaceBefore=8, spaceAfter=2)
    body_st = ParagraphStyle("KB", parent=stijlen["Normal"],
                  fontSize=10, leading=16,
                  textColor=colors.HexColor("#34495E"), spaceBefore=2)
    li_st   = ParagraphStyle("KL", parent=stijlen["Normal"],
                  fontSize=10, leading=16, leftIndent=16,
                  textColor=colors.HexColor("#34495E"), spaceBefore=1)

    # Alleen directe inhouds-tags ophalen, geen geneste containers
    # We gebruiken find_all met recursive=False per blok zodat we
    # geen parent én child allebei pakken
    for tag in soup.find_all(["h1","h2","h3","h4","h5","h6","p","li"]):
        naam  = tag.name
        tekst = tag.get_text(separator=" ", strip=True)

        # Sla leeg over
        if not tekst or len(tekst) < 2:
            continue

        # Sla dubbele teksten over (case-insensitief, whitespace genormaliseerd)
        sleutel = " ".join(tekst.lower().split())
        if sleutel in geziene_teksten:
            continue
        geziene_teksten.add(sleutel)

        # Sla technische rommel over (MSO conditional comments etc.)
        if tekst.startswith("[if ") or tekst.startswith("<!--"):
            continue

        opgemaakte_tekst = tag_naar_tekst(tag).strip()
        if not opgemaakte_tekst:
            continue

        try:
            if naam in ("h1","h2","h3","h4","h5","h6"):
                alineas.append(Paragraph(opgemaakte_tekst, kop_st))
            elif naam == "li":
                alineas.append(Paragraph(f"• {opgemaakte_tekst}", li_st))
            else:
                alineas.append(Paragraph(opgemaakte_tekst, body_st))
            alineas.append(Spacer(1, 0.1*cm))
        except Exception:
            pass

    return alineas


def pdf_maken(account, onderwerp, afzender, datum, tekst, html, bijlagen):
    tijdstempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    account_map = os.path.join(UITVOER_MAP, veilige_naam(account["naam"]))
    mail_map    = os.path.join(account_map, f"{tijdstempel}_{veilige_naam(onderwerp)}")
    os.makedirs(mail_map, exist_ok=True)

    pdf_pad   = os.path.join(mail_map, "mail.pdf")
    tijdelijk = os.path.join(mail_map, "_tijdelijk.pdf")

    stijlen   = getSampleStyleSheet()
    titel_st  = ParagraphStyle("T", parent=stijlen["Normal"],
                    fontSize=17, fontName="Helvetica-Bold",
                    textColor=colors.HexColor("#2C3E50"), spaceAfter=4)
    label_st  = ParagraphStyle("L", parent=stijlen["Normal"],
                    fontSize=9, textColor=colors.HexColor("#7F8C8D"), spaceBefore=6)
    waarde_st = ParagraphStyle("W", parent=stijlen["Normal"],
                    fontSize=10, fontName="Helvetica-Bold",
                    textColor=colors.HexColor("#2C3E50"))
    body_st   = ParagraphStyle("B", parent=stijlen["Normal"],
                    fontSize=10, leading=16,
                    textColor=colors.HexColor("#34495E"), spaceBefore=2)

    def veilig(t):
        return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # ── Header (zelfde voor HTML en tekst mails) ──────────────────────
    flow = []
    flow.append(Paragraph("Mail Export", titel_st))
    flow.append(Spacer(1, 0.3*cm))
    flow.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#3498DB")))
    flow.append(Spacer(1, 0.3*cm))

    for label, waarde in [("VAN", afzender), ("ONDERWERP", onderwerp), ("DATUM", datum)]:
        flow.append(Paragraph(label, label_st))
        flow.append(Paragraph(veilig(waarde), waarde_st))

    flow.append(Spacer(1, 0.4*cm))
    flow.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#BDC3C7")))
    flow.append(Spacer(1, 0.3*cm))
    flow.append(Paragraph("Inhoud", ParagraphStyle("K", parent=stijlen["Normal"],
                    fontSize=11, fontName="Helvetica-Bold",
                    textColor=colors.HexColor("#2C3E50"))))
    flow.append(Spacer(1, 0.2*cm))

    if html:
        # ── HTML mail → BeautifulSoup → ReportLab ────────────────────
        flow.extend(html_naar_alineas(html))
    else:
        # ── Tekst mail → ReportLab ────────────────────────────────────
        for regel in tekst.splitlines():
            if regel.strip():
                flow.append(Paragraph(veilig(regel), body_st))
            else:
                flow.append(Spacer(1, 0.15*cm))

    # ── PDF bouwen ────────────────────────────────────────────────────
    doc = SimpleDocTemplate(tijdelijk, pagesize=A4,
                            leftMargin=2.5*cm, rightMargin=2.5*cm,
                            topMargin=2*cm,    bottomMargin=2*cm)
    doc.build(flow)

    # ── PDF bijlagen samenvoegen ──────────────────────────────────────
    schrijver = PdfWriter()
    for pagina in PdfReader(tijdelijk).pages:
        schrijver.add_page(pagina)

    for bijlage in bijlagen:
        if bijlage["type"] == "application/pdf":
            print(f"    PDF bijlage toegevoegd: {bijlage['naam']}")
            for pagina in PdfReader(io.BytesIO(bijlage["data"])).pages:
                schrijver.add_page(pagina)
        else:
            bijlage_pad = os.path.join(mail_map, veilige_naam(bijlage["naam"]))
            with open(bijlage_pad, "wb") as f:
                f.write(bijlage["data"])
            print(f"    Bijlage opgeslagen: {bijlage_pad}")

    with open(pdf_pad, "wb") as f:
        schrijver.write(f)
    os.remove(tijdelijk)

    grootte = os.path.getsize(pdf_pad) // 1024
    print(f"    ✅ Map aangemaakt: {mail_map}/")
    print(f"       mail.pdf ({grootte} KB)")

def verwerk_recente_mails(account, verbinding):
    nu      = datetime.now(timezone.utc)
    grens   = nu - timedelta(hours=12)
    gisteren = (nu - timedelta(days=1)).strftime("%d-%b-%Y")
    zoekquery = f'(UNSEEN SINCE "{gisteren}")'

    _, data = verbinding.search(None, zoekquery)
    if not data or not data[0]:
        print("  Geen ongelezen mails gevonden in de afgelopen 12 uur.")
        return set()
    ids = data[0].split()
    if not ids:
        print("  Geen ongelezen mails gevonden in de afgelopen 12 uur.")
        return set()

    verwerkte_ids = set()

    for mail_id in ids:
        _, datum_data = verbinding.fetch(mail_id, "(BODY.PEEK[HEADER.FIELDS (DATE)])")
        datum_waarde = ""
        try:
            ruwe_data = vind_grootste_bytes(datum_data)
            if ruwe_data:
                datum_tekst  = ruwe_data.decode(errors="replace")
                datum_waarde = datum_tekst.replace("Date:", "").replace("DATE:", "").strip()
        except Exception:
            pass

        try:
            from email.utils import parsedate_to_datetime
            if datum_waarde:
                mail_tijd = parsedate_to_datetime(datum_waarde)
                if mail_tijd.tzinfo is None:
                    mail_tijd = mail_tijd.replace(tzinfo=timezone.utc)
                if mail_tijd < grens:
                    continue
        except Exception:
            pass

        print(f"\n  ── Recente ongelezen mail ─────────────────")
        onderwerp, afzender, datum, tekst, html, bijlagen = mail_lezen(verbinding, mail_id)
        verwerkte_ids.add(mail_id)

        if onderwerp is None:
            continue

        pdf_maken(account, onderwerp, afzender, datum, tekst, html, bijlagen)
        verbinding.store(mail_id, '+FLAGS', '\\Seen')

    return verwerkte_ids

# ── Starten ──────────────────────────────
accounts = accounts_inlezen()

if not accounts:
    print("❌ Geen accounts gevonden in .env")
    print("   Zet minimaal EMAIL_ADRES en PASSWORD in je .env bestand.")
    input("\nDruk op Enter om af te sluiten...")
    exit()

print(f"=== Mail Watcher ({len(accounts)} account(en)) ===\n")

bekende_ids_per_account = {}

for account in accounts:
    print(f"Verbinding maken met {account['naam']} ({account['server']})...")
    try:
        verbinding = verbinding_maken(account)
    except imaplib.IMAP4.error:
        print(f"❌ Inloggen mislukt voor {account['naam']}. Controleer e-mailadres/wachtwoord/server.\n")
        continue

    print("Verbonden!\n")

    print(f"── Ongelezen mails afgelopen 12 uur ({account['naam']}) ──────────")
    verwerkte_ids = verwerk_recente_mails(account, verbinding)
    print(f"\n{len(verwerkte_ids)} recente mail(s) verwerkt voor {account['naam']}.")

    bekende_ids_per_account[account["naam"]] = alle_ids_ophalen(verbinding)
    verbinding.logout()
    print()

if not bekende_ids_per_account:
    print("❌ Geen enkel account kon verbinden. Script stopt.")
    input("\nDruk op Enter om af te sluiten...")
    exit()

print(f"Wachten op nieuwe mail in {len(bekende_ids_per_account)} mailbox(en)... (check elke {CHECK_INTERVAL} seconden)")
print("Stop met Ctrl+C\n")

# ── Hoofdlus ─────────────────────────────
while True:
    time.sleep(CHECK_INTERVAL)

    print("Controleren...", end=" ")

    totaal_nieuw = 0

    for account in accounts:
        naam = account["naam"]
        if naam not in bekende_ids_per_account:
            continue

        try:
            verbinding = verbinding_maken(account)
        except imaplib.IMAP4.error:
            print(f"\n  ⚠️  Kan niet verbinden met {naam}, sla deze ronde over.")
            continue

        huidige_ids = alle_ids_ophalen(verbinding)
        nieuwe_ids  = huidige_ids - bekende_ids_per_account[naam]

        if nieuwe_ids:
            totaal_nieuw += len(nieuwe_ids)
            print(f"\n  📬 {len(nieuwe_ids)} nieuwe mail(s) op {naam}!")

            for mail_id in nieuwe_ids:
                print(f"\n  ── Nieuwe mail ({naam}) ────────────────────────")
                onderwerp, afzender, datum, tekst, html, bijlagen = mail_lezen(verbinding, mail_id)
                bekende_ids_per_account[naam].add(mail_id)

                if onderwerp is None:
                    continue

                pdf_maken(account, onderwerp, afzender, datum, tekst, html, bijlagen)
                verbinding.store(mail_id, '+FLAGS', '\\Seen')

        verbinding.logout()

    if totaal_nieuw == 0:
        print("geen nieuwe mail.")