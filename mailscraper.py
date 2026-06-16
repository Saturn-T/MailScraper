import imaplib
import email
from email.header import decode_header
import time
import os
import io
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from pypdf import PdfWriter, PdfReader

# Jouw instellingen
EMAIL = os.environ.get("EMAIL_ADRES") # verander bij ander email
WACHTWOORD = os.environ.get("PASSWORD") # verander bij ander email

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
    # ── Elke mail krijgt zijn eigen map ──────────────────────────────
    # Structuur: mail_pdfs / 20260616_201828_Onderwerp /
    #                          mail.pdf
    #                          bijlage1.txt
    #                          bijlage2.pdf
    tijdstempel = datetime.now().strftime("%Y%m%d_%H%M%S")
    map_naam    = f"{tijdstempel}_{veilige_naam(onderwerp)}"
    mail_map    = os.path.join(UITVOER_MAP, map_naam)
    os.makedirs(mail_map, exist_ok=True)

    pdf_pad   = os.path.join(mail_map, "mail.pdf")
    tijdelijk = os.path.join(mail_map, "_tijdelijk.pdf")

    # ── Stijlen ───────────────────────────────────────────────────────
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

    # ── Pagina opbouwen ───────────────────────────────────────────────
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

    # ── PDF bouwen ────────────────────────────────────────────────────
    doc = SimpleDocTemplate(tijdelijk, pagesize=A4,
                            leftMargin=2.5*cm, rightMargin=2.5*cm,
                            topMargin=2*cm,    bottomMargin=2*cm)
    doc.build(flow)

    # ── Bijlagen verwerken ────────────────────────────────────────────
    schrijver = PdfWriter()
    for pagina in PdfReader(tijdelijk).pages:
        schrijver.add_page(pagina)

    for bijlage in bijlagen:
        if bijlage["type"] == "application/pdf":
            # PDF bijlage → plakken achter de mail in mail.pdf
            print(f"    PDF bijlage toegevoegd aan mail.pdf: {bijlage['naam']}")
            for pagina in PdfReader(io.BytesIO(bijlage["data"])).pages:
                schrijver.add_page(pagina)
        else:
            # Andere bijlagen → gewoon opslaan in dezelfde map
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

# ── Starten ──────────────────────────────
print("Verbinding maken...")
verbinding = verbinding_maken()
print("Verbonden!")

bekende_ids = alle_ids_ophalen(verbinding)
verbinding.logout()
print(f"{len(bekende_ids)} bestaande mail(s) onthouden.")
print(f"Wachten op nieuwe mail... (check elke {CHECK_INTERVAL} seconden)")
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
