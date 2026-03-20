import requests
import pandas as pd
import time
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

API_KEY = "1d02aff99134b687a088bd877d4c43d63962253b"

steden = ["Arnhem", "Apeldoorn", "Zutphen", "Deventer", "Doetinchem"]
zoektermen = ["duurzaamheid", "circulariteit", "advies energietransitie", "circulair", "warmtetransitie", "procesadvies duurzaamheid", "klimaatadvies", "gebiedsinnovatie"]

headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

filter_woorden = [
    "wat wij doen", "home", "nieuws", "over ons", "contact",
    "blog", "vacatures", "producten", "diensten", "welkom",
    "bedrijven", "zoeken", "resultaten", "pagina", "artikel",
    "wikipedia", "linkedin.com/posts", "youtube", "facebook"
]

filter_urls = [
    ".nl/gemeente", "gemeente.nl", "provincie.nl",
    "klimaatbeheersing", "klimaattechniek", "installatie",
    "bouwbedrijf", "acel.nl", "despaan.nl", "ktd-", "kso-",
    "bergevoet.nl", "oude-ijsselstreek.nl", "/gebruikte-bouwmaterialen",
    "/voorbeelden-van-", "/duurzame-ondernemers",
    "arnhem.nl/alle-onderwerpen", "arnhem.nl/",
    "deventer.nl/", "apeldoorn.nl/", "zutphen.nl/",
    "doetinchem.nl/", "gelderland.nl/"
]

filter_titels = [
    "installatie", "klimaattechniek", "klimaatbeheersing",
    "bouwbedrijf", "gemeente", "provincie", "techniek b.v",
    "loodgieter", "elektricien", "aannemer",
    "gebruikte bouwmaterialen", "voorbeelden van"
]

def is_geen_bedrijf(titel, link):
    titel_lower = titel.lower()
    link_lower = link.lower()
    for woord in filter_woorden:
        if woord in titel_lower or woord in link_lower:
            return True
    for woord in filter_urls:
        if woord in link_lower:
            return True
    for woord in filter_titels:
        if woord in titel_lower:
            return True
    return False

def schoon_naam_op(titel):
    for scheidingsteken in [" | ", " - ", " :: ", " — "]:
        if scheidingsteken in titel:
            delen = titel.split(scheidingsteken)
            delen = [d.strip() for d in delen if len(d.strip()) > 3]
            if delen:
                titel = min(delen, key=len)
    return titel.strip()

def haal_omschrijving_op(url):
    try:
        response = requests.get(url, headers=headers, timeout=8)
        content_type = response.headers.get("Content-Type", "")
        if "text/html" not in content_type:
            return ""
        soup = BeautifulSoup(response.text, "html.parser")
        for tag in soup(["nav", "footer", "script", "style", "header"]):
            tag.decompose()
        teksten = []
        for p in soup.find_all("p"):
            tekst = p.get_text(strip=True)
            if len(tekst) > 60:
                teksten.append(tekst)
            if len(teksten) >= 2:
                break
        if teksten:
            omschrijving = " ".join(teksten)
            if len(omschrijving) > 300:
                afgekapt = omschrijving[:300]
                laatste_punt = max(afgekapt.rfind("."), afgekapt.rfind("!"), afgekapt.rfind("?"))
                if laatste_punt > 100:
                    omschrijving = afgekapt[:laatste_punt + 1]
                else:
                    omschrijving = afgekapt + "..."
            return omschrijving
        return ""
    except:
        return ""

def zoek_bedrijven(zoekterm, stad):
    query = f"{zoekterm} bedrijf {stad}"
    url = "https://google.serper.dev/search"
    headers_api = {
        "X-API-KEY": API_KEY,
        "Content-Type": "application/json"
    }
    payload = {
        "q": query,
        "hl": "nl",
        "gl": "nl",
        "num": 10
    }
    try:
        response = requests.post(url, headers=headers_api, json=payload, timeout=10)
        data = response.json()
        resultaten = []
        for item in data.get("organic", []):
            titel = schoon_naam_op(item.get("title", ""))
            link = item.get("link", "")
            snippet = item.get("snippet", "")
            if titel and link and not is_geen_bedrijf(titel, link):
                omschrijving = haal_omschrijving_op(link)
                if not omschrijving:
                    omschrijving = snippet
                resultaten.append({
                    "Bedrijfsnaam": titel,
                    "Website": link,
                    "Omschrijving": omschrijving,
                    "Zoekterm": zoekterm,
                    "Stad": stad
                })
        return resultaten
    except Exception as e:
        print(f"Fout bij '{query}': {e}")
        return []

alle_resultaten = []

for stad in steden:
    for zoekterm in zoektermen:
        print(f"Zoeken: {zoekterm} in {stad}...")
        resultaten = zoek_bedrijven(zoekterm, stad)
        print(f"  {len(resultaten)} resultaten gevonden")
        alle_resultaten.extend(resultaten)
        time.sleep(1)

df = pd.DataFrame(alle_resultaten)
df = df.drop_duplicates(subset=["Website"])
df = df.reset_index(drop=True)

bestandsnaam = "bedrijven_resultaten.xlsx"
df.to_excel(bestandsnaam, index=False)

wb = load_workbook(bestandsnaam)
ws = wb.active

# Klikbare links
for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    cel = row[1]
    if cel.value and str(cel.value).startswith("http"):
        cel.hyperlink = cel.value
        cel.font = Font(color="0000FF", underline="single")

# Kopteksten
header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=11)
for cel in ws[1]:
    cel.fill = header_fill
    cel.font = header_font
    cel.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 25

# Rijen opmaak
lichtgrijs = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
wit = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
for i, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
    for cel in row:
        cel.fill = lichtgrijs if i % 2 == 0 else wit
        cel.alignment = Alignment(vertical="center", wrap_text=False)
    ws.row_dimensions[i].height = 18

# Kolombreedte
for column in ws.columns:
    max_breedte = 0
    for cel in column:
        if cel.value:
            max_breedte = max(max_breedte, len(str(cel.value)))
    ws.column_dimensions[column[0].column_letter].width = min(max_breedte + 2, 50)

wb.save(bestandsnaam)
print(f"Klaar! {len(df)} bedrijven opgeslagen in '{bestandsnaam}'")