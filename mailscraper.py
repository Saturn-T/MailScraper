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

# Jouw instellingen
EMAIL = os.getenv("EMAIL_ADRES") # verander bij ander email
WACHTWOORD = os.getenv("PASSWORD") # verander bij ander email

CHECK_INTERVAL = 30  # elke 30 seconden controleren
UITVOER_MAP    = "mail_pdfs"  # map waar de PDF's worden opgeslagen

def verbinding_maken():
    verbinding = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    verbinding.login(EMAIL, WACHTWOORD)
    verbinding.select("INBOX")
    return verbinding
 
def alle_ids_ophalen(verbinding):
    _, data = verbinding.search(None, "ALL")
    return set(data[0].split())
 
def header_lezen(waarde):
    if not waarde:
        return ""
    stuk, codering = decode_header(waarde)[0]
    if isinstance(stuk, bytes):
        return stuk.decode(codering or "utf-8", errors="replace")
    return stuk
 
def mail_lezen(verbinding, mail_id):
    _, data = verbinding.fetch(mail_id, "(RFC822)")
    bericht = email.message_from_bytes(data[0][1])
 
    onderwerp = header_lezen(bericht["Subject"]) or "(geen onderwerp)"
    afzender  = header_lezen(bericht["From"])    or "Onbekend"
    datum     = bericht["Date"]                  or "Onbekend"
 
    print(f"    Van:       {afzender}")
    print(f"    Onderwerp: {onderwerp}")
    print(f"    Datum:     {datum}")
 
    tekst    = ""
    bijlagen = []
 
    if bericht.is_multipart():
        for deel in bericht.walk():
            soort        = deel.get_content_type()
            bestandsnaam = deel.get_filename()
 
            if soort == "text/plain" and not bestandsnaam:
                ruwe_tekst = deel.get_payload(decode=True)
                codering   = deel.get_content_charset() or "utf-8"
                tekst     += ruwe_tekst.decode(codering, errors="replace")
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
 
    return onderwerp, afzender, datum, tekst, bijlagen
 
def veilige_naam(tekst):
    for teken in r'\/:*?"<>|':
        tekst = tekst.replace(teken, "_")
    return tekst[:60]
 
def pdf_maken(onderwerp, afzender, datum, tekst, bijlagen):
    tijdstempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    mail_map    = os.path.join(UITVOER_MAP, f"{tijdstempel}_{veilige_naam(onderwerp)}")
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
 
def verwerk_recente_mails(verbinding):
    """
    Zoekt bij opstarten alle ongelezen mails van de afgelopen 12 uur
    en verwerkt die meteen naar PDF.
 
    IMAP heeft geen filter op uur, alleen op datum.
    Dus we zoeken op UNSEEN + datum van vandaag (en gisteren als het
    minder dan 12 uur geleden middernacht was), en filteren daarna
    zelf op timestamp of de mail binnen 12 uur valt.
    """
    nu          = datetime.now(timezone.utc)
    grens       = nu - timedelta(hours=12)
 
    # IMAP datum formaat: "01-Jan-2026"
    # We zoeken op SINCE gisteren om zeker te zijn dat we niets missen
    gisteren    = (nu - timedelta(days=1)).strftime("%d-%b-%Y")
    zoekquery   = f'(UNSEEN SINCE "{gisteren}")'
 
    _, data = verbinding.search(None, zoekquery)
    ids     = data[0].split()
 
    if not ids:
        print("  Geen ongelezen mails gevonden in de afgelopen 12 uur.")
        return set()
 
    verwerkte_ids = set()
 
    for mail_id in ids:
        # Haal alleen de datum header op (sneller dan de hele mail)
        _, datum_data = verbinding.fetch(mail_id, "(BODY[HEADER.FIELDS (DATE)])")
        datum_header  = datum_data[0][1].decode(errors="replace").strip()
        datum_waarde  = datum_header.replace("Date:", "").replace("DATE:", "").strip()
 
        try:
            # Zet de datum om naar een datetime object
            from email.utils import parsedate_to_datetime
            mail_tijd = parsedate_to_datetime(datum_waarde)
 
            # Maak timezone-aware als dat nog niet zo is
            if mail_tijd.tzinfo is None:
                mail_tijd = mail_tijd.replace(tzinfo=timezone.utc)
 
            # Controleer of de mail binnen de afgelopen 12 uur valt
            if mail_tijd < grens:
                continue  # Te oud, overslaan
 
        except Exception:
            pass  # Als we de datum niet kunnen lezen, toch verwerken
 
        print(f"\n  ── Recente ongelezen mail ─────────────────")
        onderwerp, afzender, datum, tekst, bijlagen = mail_lezen(verbinding, mail_id)
        pdf_maken(onderwerp, afzender, datum, tekst, bijlagen)
        verwerkte_ids.add(mail_id)
 
    return verwerkte_ids
 
# ── Starten ──────────────────────────────
print("\nVerbinding maken...")
try:
    verbinding = verbinding_maken()
except imaplib.IMAP4.error:
    print("❌ Inloggen mislukt. Controleer je e-mailadres en app-wachtwoord.")
    input("\nDruk op Enter om af te sluiten...")
    exit()
 
print("Verbonden!\n")
 
# Stap 1: verwerk ongelezen mails van afgelopen 12 uur
print("── Ongelezen mails afgelopen 12 uur ──────────")
verwerkte_ids = verwerk_recente_mails(verbinding)
print(f"\n{len(verwerkte_ids)} recente mail(s) verwerkt.")
 
# Stap 2: onthoud alle huidige mail IDs voor live monitoring
bekende_ids = alle_ids_ophalen(verbinding)
verbinding.logout()
 
print(f"\nWachten op nieuwe mail... (check elke {CHECK_INTERVAL} seconden)")
print("Stop met Ctrl+C\n")
 
# ── Hoofdlus ─────────────────────────────
while True:
    time.sleep(CHECK_INTERVAL)
 
    print("Controleren...", end=" ")
 
    verbinding  = verbinding_maken()
    huidige_ids = alle_ids_ophalen(verbinding)
    nieuwe_ids  = huidige_ids - bekende_ids
 
    if not nieuwe_ids:
        print("geen nieuwe mail.")
        verbinding.logout()
        continue
 
    print(f"{len(nieuwe_ids)} nieuwe mail(s)!")
 
    for mail_id in nieuwe_ids:
        print(f"\n  ── Nieuwe mail ────────────────────────")
        onderwerp, afzender, datum, tekst, bijlagen = mail_lezen(verbinding, mail_id)
        pdf_maken(onderwerp, afzender, datum, tekst, bijlagen)
        bekende_ids.add(mail_id)
 
    verbinding.logout()