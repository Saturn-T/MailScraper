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

from dotenv import load_dotenv
import os

# Laad de .env variabelen
load_dotenv()

# ─────────────────────────────────────────
# MEERDERE ACCOUNTS INLEZEN
# ─────────────────────────────────────────
# In je .env zet je nu per account een genummerde set variabelen:
#   EMAIL_ADRES_1 / PASSWORD_1 / IMAP_SERVER_1
#   EMAIL_ADRES_2 / PASSWORD_2 / IMAP_SERVER_2
#   ... enzovoort
#
# Je oude losse EMAIL_ADRES / PASSWORD blijft ook werken (wordt account 0).
#
# Bekende servers, zodat je IMAP_SERVER_x niet altijd hoeft op te zoeken:
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
    """Probeert de IMAP-server te laden op basis van het mailadres."""
    domein = email_adres.split("@")[-1].lower()
    return BEKENDE_SERVERS.get(domein)

def accounts_inlezen():
    """
    Bouwt een lijst van accounts op uit de .env.
    Elk account is een dict met: naam, email, wachtwoord, server.
    """
    accounts = []

    # Eerst: het oorspronkelijke enkele account (EMAIL_ADRES / PASSWORD)
    # blijft gewoon werken, zodat bestaande .env bestanden niet kapot gaan.
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

    # Daarna: genummerde accounts EMAIL_ADRES_1, EMAIL_ADRES_2, ...
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

CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL", 30))  # elke 30 seconden controleren
UITVOER_MAP    = os.getenv("UITVOER_MAP", "mail_pdfs")  # map waar de PDF's worden opgeslagen

# ─────────────────────────────────────────

def verbinding_maken(account):
    """Maakt verbinding met de mailbox van één account."""
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
    """
    Zoekt recursief naar alle bytes-objecten en retourneert het grootste.
    Bij IMAP fetch responses is de grootste bytes altijd de daadwerkelijke 
    maildata, niet de metadata-string.
    """
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
    
    # Retourneer het grootste bytes object
    return max(gevonden, key=len)

def is_reclame(bericht):
    """
    Controleert of een mail waarschijnlijk reclame/nieuwsbrief is.
    Werkt voor elke mailprovider omdat het kijkt naar kenmerken
    van de mail zelf, niet naar Gmail-specifieke labels.

    Twee signalen:
    1. List-Unsubscribe header → bijna elke nieuwsbrief/marketing-mail
       heeft deze header. Normale persoonlijke mail heeft die niet.
    2. Afzenderadres bevat woorden als "no-reply", "newsletter", etc.
    """
    # Signaal 1: List-Unsubscribe header aanwezig
    if bericht.get("List-Unsubscribe"):
        return True

    # Signaal 2: verdachte woorden in het afzenderadres
    afzender = (bericht.get("From") or "").lower()
    verdachte_woorden = [
        "no-reply", "noreply", "no.reply",
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

    if not ruwe_mail or len(ruwe_mail) < 10:  # minder dan 10 bytes = geen echte mail
        print("    ⚠️  Kon maildata niet lezen (onverwachte IMAP-response)")
        return None, None, None, None, None
    
    bericht = email.message_from_bytes(ruwe_mail)

    onderwerp = header_lezen(bericht["Subject"]) or "(geen onderwerp)"
    afzender  = header_lezen(bericht["From"])    or "Onbekend"
    datum     = bericht["Date"]                  or "Onbekend"

    print(f"    Van:       {afzender}")
    print(f"    Onderwerp: {onderwerp}")
    print(f"    Datum:     {datum}")

    # Reclame/nieuwsbrief? Dan overslaan.
    if is_reclame(bericht):
        print(f"    ⏭️  Overgeslagen (reclame/nieuwsbrief)")
        return None, None, None, None, None

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

    if not tekst.strip() and html:
        tekst = html_strippen(html)
        print(f"    (HTML mail → tekst geëxtraheerd)")
    elif re.search(r"<[a-zA-Z]+[\s>]", tekst):
        tekst = html_strippen(tekst)
        print(f"    (HTML in tekstveld → tags gestript)")

    return onderwerp, afzender, datum, tekst, bijlagen

def veilige_naam(tekst):
    """
    Maakt een string geschikt als Windows-bestandsnaam/mapnaam.
    Windows accepteert geen mappen/bestanden die eindigen op een spatie of punt.
    """
    for teken in r'\/:*?"<>|':
        tekst = tekst.replace(teken, "_")
    tekst = tekst.strip()           # spaties aan begin/eind weghalen
    tekst = tekst.rstrip(". ")      # Windows accepteert geen punt/spatie aan het eind
    tekst = tekst[:60].strip()      # inkorten en nogmaals trimmen
    tekst = tekst.rstrip(". ")      # voor het geval het afkappen weer op spatie/punt eindigt
    if not tekst:
        tekst = "geen_onderwerp"    # als er niks overblijft
    return tekst

def pdf_maken(account, onderwerp, afzender, datum, tekst, bijlagen):
    tijdstempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Elk account krijgt zijn eigen submap binnen UITVOER_MAP,
    # zodat je meteen ziet via welke mailbox een mail binnenkwam.
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

    for regel in tekst.splitlines():
        if regel.strip():
            flow.append(Paragraph(veilig(regel), body_st))
        else:
            flow.append(Spacer(1, 0.15*cm))

    doc = SimpleDocTemplate(tijdelijk, pagesize=A4,
                            leftMargin=2.5*cm, rightMargin=2.5*cm,
                            topMargin=2*cm,    bottomMargin=2*cm)
    doc.build(flow)

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
    """
    Zoekt bij opstarten alle ongelezen mails van de afgelopen 12 uur
    en verwerkt die meteen naar PDF.
    """
    nu          = datetime.now(timezone.utc)
    grens       = nu - timedelta(hours=12)

    gisteren    = (nu - timedelta(days=1)).strftime("%d-%b-%Y")
    zoekquery   = f'(UNSEEN SINCE "{gisteren}")'

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
        # Haal alleen de datum header op (sneller dan de hele mail)
        _, datum_data = verbinding.fetch(mail_id, "(BODY.PEEK[HEADER.FIELDS (DATE)])")
        
        datum_waarde = ""
        try:
            ruwe_data = vind_grootste_bytes(datum_data)
            if ruwe_data:
                datum_tekst = ruwe_data.decode(errors="replace")
                datum_waarde = datum_tekst.replace("Date:", "").replace("DATE:", "").strip()
        except Exception:
            pass  # Als we de datum niet kunnen lezen, toch verwerken

        try:
            from email.utils import parsedate_to_datetime
            if datum_waarde:
                mail_tijd = parsedate_to_datetime(datum_waarde)

                if mail_tijd.tzinfo is None:
                    mail_tijd = mail_tijd.replace(tzinfo=timezone.utc)

                if mail_tijd < grens:
                    continue  # Te oud, overslaan
        except Exception:
            pass  # Als we de datum niet kunnen parsen, toch verwerken

        print(f"\n  ── Recente ongelezen mail ─────────────────")
        onderwerp, afzender, datum, tekst, bijlagen = mail_lezen(verbinding, mail_id)
        verwerkte_ids.add(mail_id)

        if onderwerp is None:
            continue  # was reclame, overslaan

        pdf_maken(account, onderwerp, afzender, datum, tekst, bijlagen)

    return verwerkte_ids

# ── Starten ──────────────────────────────
accounts = accounts_inlezen()

if not accounts:
    print("❌ Geen accounts gevonden in .env")
    print("   Zet minimaal EMAIL_ADRES en PASSWORD in je .env bestand,")
    print("   of gebruik EMAIL_ADRES_1 / PASSWORD_1 voor meerdere accounts.")
    input("\nDruk op Enter om af te sluiten...")
    exit()

print(f"=== Mail Watcher ({len(accounts)} account(en)) ===\n")

# bekende_ids_per_account onthoudt, per account, welke mail-IDs we al kennen.
# Zo houden we de mailboxen los van elkaar.
bekende_ids_per_account = {}

for account in accounts:
    print(f"Verbinding maken met {account['naam']} ({account['server']})...")
    try:
        verbinding = verbinding_maken(account)
    except imaplib.IMAP4.error:
        print(f"❌ Inloggen mislukt voor {account['naam']}. Controleer e-mailadres/wachtwoord/server.\n")
        continue

    print("Verbonden!\n")

    # Stap 1: verwerk ongelezen mails van afgelopen 12 uur
    print(f"── Ongelezen mails afgelopen 12 uur ({account['naam']}) ──────────")
    verwerkte_ids = verwerk_recente_mails(account, verbinding)
    print(f"\n{len(verwerkte_ids)} recente mail(s) verwerkt voor {account['naam']}.")

    # Stap 2: onthoud alle huidige mail IDs voor live monitoring
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

    # Loop door elk account heen en check op nieuwe mail
    for account in accounts:
        naam = account["naam"]
        if naam not in bekende_ids_per_account:
            continue  # dit account kon niet inloggen bij het opstarten

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
                onderwerp, afzender, datum, tekst, bijlagen = mail_lezen(verbinding, mail_id)
                bekende_ids_per_account[naam].add(mail_id)

                if onderwerp is None:
                    continue  # was reclame, overslaan

                pdf_maken(account, onderwerp, afzender, datum, tekst, bijlagen)
                verbinding.store(mail_id, '+FLAGS', '\\Seen')


        verbinding.logout()

    if totaal_nieuw == 0:
        print("geen nieuwe mail.")
