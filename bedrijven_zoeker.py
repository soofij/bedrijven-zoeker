import streamlit as st
import requests
import pandas as pd
import time
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
import io

st.set_page_config(page_title="Bedrijven Zoeker", page_icon="🔍")
st.title("Bedrijven Zoeker")
st.write("Zoek bedrijven op thema en regio en download de resultaten als Excel.")

api_key = st.text_input("Serper API key", type="password")
steden_input = st.text_input("Steden (kommagescheiden)", "Arnhem, Apeldoorn, Zutphen, Deventer, Doetinchem")
zoektermen_input = st.text_input("Zoektermen (kommagescheiden)", "duurzaamheid, circulariteit, warmtetransitie, klimaatadvies")
max_resultaten = st.slider("Max resultaten per zoekopdracht", 5, 50, 10)

filter_urls = [
    ".nl/gemeente", "gemeente.nl", "provincie.nl",
    "klimaatbeheersing", "klimaattechniek", "installatie",
    "bouwbedrijf", "acel.nl", "despaan.nl", "ktd-", "kso-",
    "bergevoet.nl", "oude-ijsselstreek.nl", "/gebruikte-bouwmaterialen",
    "/voorbeelden-van-", "/duurzame-ondernemers",
    "arnhem.nl/alle-onderwerpen", "deventer.nl/", "apeldoorn.nl/",
    "zutphen.nl/", "doetinchem.nl/", "gelderland.nl/"
]

filter_woorden = [
    "wat wij doen", "home", "nieuws", "over ons", "contact",
    "blog", "vacatures", "producten", "diensten", "welkom",
    "bedrijven", "zoeken", "resultaten", "pagina", "artikel",
    "wikipedia", "linkedin.com/posts", "youtube", "facebook"
]

filter_titels = [
    "installatie", "klimaattechniek", "klimaatbeheersing",
    "bouwbedrijf", "gemeente", "provincie", "techniek b.v",
    "loodgieter", "elektricien", "aannemer",
    "gebruikte bouwmaterialen", "voorbeelden van"
]

headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

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

def zoek_bedrijven(zoekterm, stad, api_key, max_resultaten):
    query = f"{zoekterm} bedrijf {stad}"
    url = "https://google.serper.dev/search"
    headers_api = {
        "X-API-KEY": api_key,
        "Content-Type": "application/json"
    }
    payload = {"q": query, "hl": "nl", "gl": "nl", "num": max_resultaten}
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
        return []

if st.button("Zoeken"):
    if not api_key:
        st.error("Vul je Serper API key in!")
    else:
        steden = [s.strip() for s in steden_input.split(",")]
        zoektermen = [z.strip() for z in zoektermen_input.split(",")]
        alle_resultaten = []
        totaal = len(steden) * len(zoektermen)
        voortgang = st.progress(0)
        status = st.empty()
        stap = 0

        for stad in steden:
            for zoekterm in zoektermen:
                status.text(f"Zoeken: {zoekterm} in {stad}...")
                resultaten = zoek_bedrijven(zoekterm, stad, api_key, max_resultaten)
                alle_resultaten.extend(resultaten)
                stap += 1
                voortgang.progress(stap / totaal)
                time.sleep(1)

        df = pd.DataFrame(alle_resultaten)
        df = df.drop_duplicates(subset=["Website"])
        df = df.reset_index(drop=True)

        status.text(f"Klaar! {len(df)} bedrijven gevonden.")
        voortgang.progress(1.0)

        st.success(f"{len(df)} bedrijven gevonden!")
        st.dataframe(
            df[["Bedrijfsnaam", "Website", "Omschrijving", "Stad"]],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Website": st.column_config.LinkColumn("Website")
            }
        )
