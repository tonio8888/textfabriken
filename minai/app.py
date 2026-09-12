import streamlit as st
import requests
import time
import sqlite3
import json
import uuid
import csv
from pypdf import PdfReader
import io
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font

st.set_page_config(page_title="TextFabriken AI", page_icon="🏭", layout="centered", initial_sidebar_state="expanded")

# --- SÄKRAD API-NYCKEL (Hämtas från ditt Streamlit-valv) ---
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]

# --- SPRÅKINSTÄLLNINGAR ---
SPRAK_ALTERNATIV = ["Svenska", "Norsk", "Dansk", "Suomi", "English"]

UI_TEXTS = {
    "Svenska": {
        "subheader": "Nordens smartaste löpande band for produktbeskrivningar",
        "intro": "Ladda upp din rådata i bottenmenyn. TextFabriken transformerar den till säljande SEO-texter och skapar en färdig Word-fil åt dig!",
        "lang_label": "🌐 Språk för texter och gränssnitt",
        "workspace_info": "🔑 Din arbetsyta-kod: **{kod}**\n\nSpara den här sidans webbadress (bokmärk fliken) för att komma tillbaka till dina sparade listor senare.",
        "saved_header": "📁 Sparade produktlistor",
        "no_saved": "Inga sparade listor ännu.",
        "active_now": "Aktiv nu:",
        "chat_placeholder": "Skriv instruktion till fabriken...",
        "expander_label": "📎 Klicka här för att mata in produktfakta (PDF/TXT/WORD)",
        "uploader_label": "Välj dokument från din dator",
        "rocket_button": "🚀 Starta Massgenerering",
        "error_no_file": "⚠️ Du måste välja en fil i rutan till vänster först!",
        "download_header": "### 📥 Din färdiga Word-fil från TextFabriken är klar!",
        "download_info": "💡 Innan du publicerar: dubbelkolla alla siffror (mått, kapacitet, batteritid, dB-nivåer m.m.) mot leverantörens originaldata. TextFabriken skriver säljande texter, men ansvarar inte för att specifikationerna är korrekta.",
        "download_button": "📝 Ladda ner produkttexter (.docx)",
        "download_button_csv": "📊 Ladda ner som CSV (.csv)",
        "download_button_xlsx": "📊 Ladda ner som Excel (.xlsx)",
        "warning_text": "⚠️ **Kontrollera alltid siffror och specifikationer** (t.ex. batteritid, mått, prestanda) mot din egen produktdata innan du publicerar texterna.",
        "processing_batch": "*Bearbetar del {i} av {n} i TextFabrikens maskiner...*",
        "processing_single": "*TextFabriken bearbetar dina ord i molnet...*",
        "processing_mass": "*Maskinerna i TextFabriken startar upp i molnet...*",
        "user_label": "Du",
        "assistant_label": "TextFabriken",
        "doc_heading": "SEO Produktbeskrivningar - TextFabriken AI",
        "mass_prompt_text": "Massgenerera SEO-produktbeskrivningar för alla produkter i listan.",
        "system_prompt_free_chat": (
            "Du är TextFabriken, en AI-assistent som hjälper e-handlare att skriva SEO-produktbeskrivningar. Svara alltid på svenska. "
            "Du är EN specifik app, inte en generell e-handelsplattform. Du kan ENDAST: "
            "1) transformera en uppladdad produktlista (PDF/TXT/Word) till säljande SEO-texter, och "
            "2) svara på allmänna frågor om copywriting, SEO eller produkttexter. "
            "Du kan INTE hantera ordrar, kundkonton, betalningar, leveranser, returer eller lagerstatus – det är inte vad den här appen gör, och du ska aldrig hitta på att du kan det. "
            "Om användaren vill ladda ner något du skrivit, säg åt dem att använda den blå nedladdningsknappen som automatiskt dyker upp under chatten efter varje svar du ger. "
            "Om en fråga ligger utanför det du faktiskt kan göra, säg det ärligt istället för att gissa eller hitta på en lösning. Svara naturligt och hjälpsamt inom dessa gränser."
        ),
    },
    "Norsk": {
        "subheader": "Nordens smarteste samlebånd for produktbeskrivelser",
        "intro": "Last opp rådataene dine i menyen nederst. TextFabriken forvandler dem til salgsfremmende SEO-tekster og lager en ferdig Word-fil for deg!",
        "lang_label": "🌐 Språk for tekster og grensesnitt",
        "workspace_info": "🔑 Din arbeidsflate-kode: **{kod}**\n\nLagre nettadressen til denne siden (bokmerk fanen) for å komme tilbake til dine lagrede lister senere.",
        "saved_header": "📁 Lagrede produktlister",
        "no_saved": "Ingen lagrede lister ennå.",
        "active_now": "Aktiv nå:",
        "chat_placeholder": "Skriv en instruksjon til fabrikken...",
        "expander_label": "📎 Klikk her for å legge inn produktdata (PDF/TXT/WORD)",
        "uploader_label": "Velg dokument fra datamaskinen din",
        "rocket_button": "🚀 Start Massegenerering",
        "error_no_file": "⚠️ Du må velge en fil i boksen til venstre først!",
        "download_header": "### 📥 Din ferdige Word-fil fra TextFabriken er klar!",
        "download_info": "💡 Før du publiserer: dobbeltsjekk alle tall (mål, kapasitet, batteritid, dB-nivåer osv.) mot leverandørens originaldata. TextFabriken skriver salgstekster, men er ikke ansvarlig for at spesifikasjonene er korrekte.",
        "download_button": "📝 Last ned produkttekster (.docx)",
        "download_button_csv": "📊 Last ned som CSV (.csv)",
        "download_button_xlsx": "📊 Last ned som Excel (.xlsx)",
        "warning_text": "⚠️ **Kontroller alltid tall og spesifikasjoner** (f.eks. batteritid, mål, ytelse) mot din egen produktdata før du publiserer tekstene.",
        "processing_batch": "*Behandler del {i} av {n} i TextFabrikens maskiner...*",
        "processing_single": "*TextFabriken behandler ordene dine i skyen...*",
        "processing_mass": "*Maskinene i TextFabriken starter opp i skyen...*",
        "user_label": "Du",
        "assistant_label": "TextFabriken",
        "doc_heading": "SEO Produktbeskrivelser - TextFabriken AI",
        "mass_prompt_text": "Massegenerer SEO-produktbeskrivelser for alle produktene i listen.",
        "system_prompt_free_chat": (
            "Du er TextFabriken, en AI-assistent som hjelper nettbutikker med å skrive SEO-produktbeskrivelser. Svar alltid på norsk. "
            "Du er ÉN spesifikk app, ikke en generell e-handelsplattform. Du kan KUN: "
            "1) omdanne en opplastet produktliste (PDF/TXT/Word) til salgsfremmende SEO-tekster, og "
            "2) svare på generelle spørsmål om copywriting, SEO eller produkttekster. "
            "Du kan IKKE håndtere bestillinger, kundekontoer, betalinger, levering, retur eller lagerstatus – det er ikke hva denne appen gjør, og du skal aldri late som om du kan det. "
            "Hvis brukeren vil laste ned noe du har skrevet, be dem bruke den blå nedlastingsknappen som automatisk dukker opp under chatten etter hvert svar. "
            "Hvis et spørsmål ligger utenfor det du faktisk kan gjøre, si det ærlig i stedet for å gjette eller finne på en løsning. Svar naturlig og hjelpsomt innenfor disse grensene."
        ),
    },
    "Dansk": {
        "subheader": "Nordens smarteste samlebånd til produktbeskrivelser",
        "intro": "Upload dine rådata i menuen nedenfor. TextFabriken forvandler dem til salgsfremmende SEO-tekster og genererer en færdig Word-fil til dig!",
        "lang_label": "🌐 Sprog til tekster og brugerflade",
        "workspace_info": "🔑 Din arbejdsområde-kode: **{kod}**\n\nGem denne sides webadresse (sæt bogmærke i fanen) for at komme tilbage til dine gemte lister senere.",
        "saved_header": "📁 Gemte produktlister",
        "no_saved": "Ingen gemte lister endnu.",
        "active_now": "Aktiv nu:",
        "chat_placeholder": "Skriv en instruktion til fabrikken...",
        "expander_label": "📎 Klik her for at tilføje produktdata (PDF/TXT/WORD)",
        "uploader_label": "Vælg dokument fra din computer",
        "rocket_button": "🚀 Start Masseproduktion",
        "error_no_file": "⚠️ Du skal vælge en fil i boksen til venstre først!",
        "download_header": "### 📥 Din færdige Word-fil fra TextFabriken er klar!",
        "download_info": "💡 Inden du publicerer: dobbelttjek alle tal (mål, kapacitet, batteritid, dB-niveauer m.m.) mod leverandørens originaldata. TextFabriken skriver salgsfremmende tekster, men er ikke ansvarlig for at specifikationerne er korrekte.",
        "download_button": "📝 Download produkttekster (.docx)",
        "download_button_csv": "📊 Download som CSV (.csv)",
        "download_button_xlsx": "📊 Download som Excel (.xlsx)",
        "warning_text": "⚠️ **Kontroller altid tal og specifikationer** (f.eks. batteritid, mål, ydeevne) mod dine egne produktdata, inden du publicerer teksterne.",
        "processing_batch": "*Behandler del {i} af {n} i TextFabrikens maskiner...*",
        "processing_single": "*TextFabriken behandler dine ord i skyen...*",
        "processing_mass": "*Maskinerne i TextFabriken starter op i skyen...*",
        "user_label": "Dig",
        "assistant_label": "TextFabriken",
        "doc_heading": "SEO Produktbeskrivelser - TextFabriken AI",
        "mass_prompt_text": "Masseproducér SEO-produktbeskrivelser for alle produkter i listen.",
        "system_prompt_free_chat": (
            "Du er TextFabriken, en AI-assistent, der hjælper netbutikker med at skrive SEO-produktbeskrivelser. Svar altid på dansk. "
            "Du er ÉN specifik app, ikke en generel e-handelsplatform. Du kan KUN: "
            "1) omdanne en uploadet produktliste (PDF/TXT/Word) til salgsfremmende SEO-tekster, og "
            "2) svare på generelle spørgsmål om copywriting, SEO eller produkttekster. "
            "Du kan IKKE håndtere ordrer, kundekonti, betalinger, levering, returnering eller lagerstatus – det er ikke, hvad denne app gør, og du må aldrig lade som om du kan det. "
            "Hvis brugeren vil downloade noget, du har skrevet, så bed dem bruge den blå downloadknap, der automatisk dukker op under chatten efter hvert svar. "
            "Hvis et spørgsmål ligger uden for det, du faktisk kan gøre, så sig det ærligt i stedet for at gætte eller finde på en løsning. Svar naturligt og hjælpsomt inden for disse grænser."
        ),
    },
    "Suomi": {
        "subheader": "Pohjolan fiksuin tuotantolinja tuotekuvauksille",
        "intro": "Lataa raakadatasi alavalikossa. TextFabriken muuttaa sen myyväksi SEO-tekstiksi ja luo sinulle valmiin Word-tiedoston!",
        "lang_label": "🌐 Tekstien ja käyttöliittymän kieli",
        "workspace_info": "🔑 Työtilasi koodi: **{kod}**\n\nTallenna tämän sivun verkko-osoite (lisää kirjanmerkki) päästäksesi takaisin tallennettuihin listoihisi myöhemmin.",
        "saved_header": "📁 Tallennetut tuotelistat",
        "no_saved": "Ei vielä tallennettuja listoja.",
        "active_now": "Aktiivinen nyt:",
        "chat_placeholder": "Kirjoita ohje tehtaalle...",
        "expander_label": "📎 Klikkaa tästä syöttääksesi tuotetiedot (PDF/TXT/WORD)",
        "uploader_label": "Valitse tiedosto tietokoneeltasi",
        "rocket_button": "🚀 Aloita Massatuotanto",
        "error_no_file": "⚠️ Sinun täytyy ensin valita tiedosto vasemmalla olevasta ruudusta!",
        "download_header": "### 📥 Valmis Word-tiedostosi TextFabrikenilta on nyt valmis!",
        "download_info": "💡 Ennen julkaisua: tarkista aina kaikki luvut (mitat, kapasiteetti, akun kesto, dB-tasot jne.) valmistajan alkuperäisistä tiedoista. TextFabriken kirjoittaa myyviä tekstejä, mutta ei vastaa spesifikaatioiden oikeellisuudesta.",
        "download_button": "📝 Lataa tuotetekstit (.docx)",
        "download_button_csv": "📊 Lataa CSV-tiedostona (.csv)",
        "download_button_xlsx": "📊 Lataa Excel-tiedostona (.xlsx)",
        "warning_text": "⚠️ **Tarkista aina luvut ja spesifikaatiot** (esim. akun kesto, mitat, suorituskyky) omista tuotetiedoistasi ennen tekstien julkaisua.",
        "processing_batch": "*Käsitellään osaa {i}/{n} TextFabrikenin koneissa...*",
        "processing_single": "*TextFabriken käsittelee sanojasi pilvessä...*",
        "processing_mass": "*TextFabrikenin koneet käynnistyvät pilvessä...*",
        "user_label": "Sinä",
        "assistant_label": "TextFabriken",
        "doc_heading": "SEO-tuotekuvaukset - TextFabriken AI",
        "mass_prompt_text": "Luo massana SEO-tuotekuvaukset kaikille listan tuotteille.",
        "system_prompt_free_chat": (
            "CRITICAL RULE: You must respond ONLY in Finnish (suomi). NEVER respond in Swedish, Danish, Norwegian, or English, "
            "regardless of what language the user writes in. This rule overrides everything else. "
            "Olet TextFabriken, tekoälyavustaja, joka auttaa verkkokauppiaita kirjoittamaan SEO-tuotekuvauksia. Vastaa AINA JA VAIN suomeksi. "
            "Olet YKSI tietty sovellus, et yleinen verkkokauppa-alusta. Voit AINOASTAAN: "
            "1) muuttaa ladatun tuotelistan (PDF/TXT/Word) myyviksi SEO-teksteiksi, ja "
            "2) vastata yleisiin kysymyksiin copywritingista, SEO:sta tai tuoteteksteistä. "
            "Et voi käsitellä tilauksia, asiakastilejä, maksuja, toimituksia, palautuksia tai varastotilannetta – tämä sovellus ei tee sitä, äläkä koskaan väitä pystyväsi siihen. "
            "Jos käyttäjä haluaa ladata jotain kirjoittamaasi, kehota häntä käyttämään sinistä latauspainiketta, joka ilmestyy automaattisesti keskustelun alle jokaisen vastauksesi jälkeen. "
            "Jos kysymys on jotain, mitä et oikeasti osaa tehdä, kerro se rehellisesti sen sijaan, että arvailisit tai keksisit ratkaisun. Vastaa luonnollisesti ja avuliaasti näiden rajojen sisällä."
        ),
    },
    "English": {
        "subheader": "The Nordics' smartest assembly line for product descriptions",
        "intro": "Upload your raw data in the menu below. TextFabriken transforms it into compelling SEO copy and creates a ready-made Word file for you!",
        "lang_label": "🌐 Language for texts and interface",
        "workspace_info": "🔑 Your workspace code: **{kod}**\n\nSave this page's URL (bookmark the tab) to return to your saved lists later.",
        "saved_header": "📁 Saved product lists",
        "no_saved": "No saved lists yet.",
        "active_now": "Currently active:",
        "chat_placeholder": "Write an instruction to the factory...",
        "expander_label": "📎 Click here to add product data (PDF/TXT/WORD)",
        "uploader_label": "Choose a document from your computer",
        "rocket_button": "🚀 Start Mass Generation",
        "error_no_file": "⚠️ You need to select a file in the box on the left first!",
        "download_header": "### 📥 Your finished Word file from TextFabriken is ready!",
        "download_info": "💡 Before publishing: double-check all figures (dimensions, capacity, battery life, dB levels, etc.) against the manufacturer's original data. TextFabriken writes compelling copy, but is not responsible for the accuracy of the specifications.",
        "download_button": "📝 Download product texts (.docx)",
        "download_button_csv": "📊 Download as CSV (.csv)",
        "download_button_xlsx": "📊 Download as Excel (.xlsx)",
        "warning_text": "⚠️ **Always verify figures and specifications** (e.g. battery life, dimensions, performance) against your own product data before publishing the texts.",
        "processing_batch": "*Processing part {i} of {n} in TextFabriken's machines...*",
        "processing_single": "*TextFabriken is processing your words in the cloud...*",
        "processing_mass": "*TextFabriken's machines are starting up in the cloud...*",
        "user_label": "You",
        "assistant_label": "TextFabriken",
        "doc_heading": "SEO Product Descriptions - TextFabriken AI",
        "mass_prompt_text": "Mass-generate SEO product descriptions for all products in the list.",
        "system_prompt_free_chat": (
            "You are TextFabriken, an AI assistant that helps online retailers write SEO product descriptions. Always respond in English. "
            "You are ONE specific app, not a general e-commerce platform. You can ONLY: "
            "1) transform an uploaded product list (PDF/TXT/Word) into compelling SEO copy, and "
            "2) answer general questions about copywriting, SEO, or product texts. "
            "You CANNOT handle orders, customer accounts, payments, shipping, returns, or stock status – that is not what this app does, and you must never pretend that you can. "
            "If the user wants to download something you wrote, tell them to use the blue download button that automatically appears below the chat after every response. "
            "If a question is outside what you can actually do, say so honestly instead of guessing or making up a solution. Respond naturally and helpfully within these boundaries."
        ),
    },
}

SEO_DIREKTIV = {
    "Svenska": (
        "Du är TextFabriken, en absolut världsmästare på e-handel, digital marknadsföring och SEO-copywriting för den nordiska marknaden. "
        "Skriv ALLTID på svenska. "
        "Du har fått en fil med rådata eller en lista på produkter. Din uppgift är att transformera denna lista till supersäljande, kaxiga och moderna produktbeskrivningar på svenska. "
        "VIKTIGT: Hitta aldrig på exakta siffror, mått, tekniska specifikationer eller prestandavärden (t.ex. batteritid, dB-nivåer, DPI, kapacitet i ml/liter) som inte finns i den rådata du fått. "
        "Om en specifik siffra saknas i underlaget, skriv istället kvalitativt (t.ex. 'lång batteritid' eller 'kraftfull brusreducering') utan att gissa ett exakt tal. "
        "VIKTIGT: Skapa aldrig fler produkter än de som faktiskt finns i rådatan. Om ett produktnamn består av flera ord, behandla alltid hela namnet som EN enda produkt – dela aldrig upp det i flera separata produkter. "
        "Varje produkt ska delas upp enligt följande strikta struktur:\n"
        "PRODUKTNAMN (Använd fetstil)\n"
        "SÄLJANDE BESKRIVNING: Skriv cirka 100 ord som skapar ett extremt starkt 'ha-begär' hos kunden.\n"
        "NYCKELFÖRDELAR:\n- Punkt 1\n- Punkt 2\n- Punkt 3\n"
        "SEO-TAGGAR: Lägg till 5 relevanta sökord för Google.\n"
        "Använd absolut inga emojier eller färgade prickar. Skriv ut texterna direkt efter varandra, separation med ett streck (---) mellan varje produkt."
    ),
    "Norsk": (
        "Du er TextFabriken, en absolutt verdensmester innen e-handel, digital markedsføring og SEO-copywriting for det nordiske markedet. "
        "Skriv ALLTID på norsk (bokmål). "
        "Du har fått en fil med rådata eller en liste over produkter. Din oppgave er å omdanne denne listen til superselgende, frekke og moderne produktbeskrivelser på norsk. "
        "VIKTIG: Finn aldri på eksakte tall, mål, tekniske spesifikasjoner eller ytelsesverdier (f.eks. batteritid, dB-nivåer, DPI, kapasitet i ml/liter) som ikke finnes i rådataen du har fått. "
        "Hvis et spesifikt tall mangler i grunnlaget, skriv i stedet kvalitativt (f.eks. 'lang batteritid' eller 'kraftig støyreduksjon') uten å gjette et eksakt tall. "
        "VIKTIG: Lag aldri flere produkter enn de som faktisk finnes i rådataen. Hvis et produktnavn består av flere ord, behandle alltid hele navnet som ÉTT produkt – del det aldri opp i flere separate produkter. "
        "Hvert produkt skal deles opp etter følgende strikte struktur:\n"
        "PRODUKTNAVN (Bruk fet skrift)\n"
        "SALGSBESKRIVELSE: Skriv rundt 100 ord som skaper et ekstremt sterkt 'må ha'-behov hos kunden.\n"
        "NØKKELFORDELER:\n- Punkt 1\n- Punkt 2\n- Punkt 3\n"
        "SEO-STIKKORD: Legg til 5 relevante søkeord for Google.\n"
        "Bruk absolutt ingen emojier eller fargede prikker. Skriv ut tekstene rett etter hverandre, separert med en strek (---) mellom hvert produkt."
    ),
    "Dansk": (
        "Du er TextFabriken, en absolut verdensmester inden for e-handel, digital markedsføring og SEO-copywriting til det nordiske marked. "
        "Skriv ALTID på dansk. "
        "Du har fået en fil med rådata eller en liste over produkter. Din opgave er at omdanne denne liste til superoverbevisende, freaky og moderne produktbeskrivelser på dansk. "
        "VIGTIGT: Find aldrig på eksakte tal, mål, tekniske specifikationer eller ydeevneværdier (f.eks. batteritid, dB-niveauer, DPI, kapacitet i ml/liter), som ikke findes i den rådata, du har fået. "
        "Hvis et specifikt tal mangler i grundlaget, skriv i stedet kvalitativt (f.eks. 'lang batteritid' eller 'kraftig støjreduktion') uden at gætte et eksakt tal. "
        "VIGTIGT: Skab aldrig flere produkter end dem, der faktisk findes i rådataen. Hvis et produktnavn består af flere ord, behandl altid hele navnet som ÉT produkt – del det aldrig op i flere separate produkter. "
        "Hvert produkt skal opdeles efter følgende strikte struktur:\n"
        "PRODUKTNAVN (Brug fed skrift)\n"
        "SALGSBESKRIVELSE: Skriv omkring 100 ord, der skaber et ekstremt stærkt 'må have'-behov hos kunden.\n"
        "NØGLEFORDELE:\n- Punkt 1\n- Punkt 2\n- Punkt 3\n"
        "SEO-TAGS: Tilføj 5 relevante søgeord til Google.\n"
        "Brug absolut ingen emojis eller farvede prikker. Skriv teksterne direkte efter hinanden, adskilt med en streg (---) mellem hvert produkt."
    ),
    "Suomi": (
        "CRITICAL RULE: You must write ONLY in Finnish (suomi). NEVER write in Swedish, Danish, Norwegian, or English. This rule overrides everything else. "
        "Olet TextFabriken, ehdoton maailmanmestari verkkokaupassa, digitaalisessa markkinoinnissa ja SEO-copywritingissa Pohjoismaiden markkinoille. "
        "Kirjoita AINA JA VAIN suomeksi. "
        "Olet saanut tiedoston, jossa on raakadataa tai lista tuotteista. Tehtäväsi on muuttaa tämä lista supermyyviksi, rohkeiksi ja moderneiksi tuotekuvauksiksi suomeksi. "
        "TÄRKEÄÄ: Älä koskaan keksi tarkkoja lukuja, mittoja, teknisiä tietoja tai suorituskykyarvoja (esim. akun kesto, dB-tasot, DPI, kapasiteetti ml/l), joita ei ole annetussa raakadatassa. "
        "Jos tarkka luku puuttuu lähdemateriaalista, kirjoita sen sijaan laadullisesti (esim. 'pitkä akunkesto' tai 'tehokas melunvaimennus') arvaamatta tarkkaa lukua. "
        "TÄRKEÄÄ: Älä koskaan luo enempää tuotteita kuin raakadatassa todella on. Jos tuotenimi koostuu useasta sanasta, käsittele koko nimeä aina YHTENÄ tuotteena – älä koskaan jaa sitä useiksi erillisiksi tuotteiksi. "
        "Jokainen tuote tulee jäsentää seuraavan tiukan rakenteen mukaisesti:\n"
        "TUOTTEEN NIMI (Käytä lihavointia)\n"
        "MYYVÄ KUVAUS: Kirjoita noin 100 sanaa, jotka luovat asiakkaalle vahvan ostohalun.\n"
        "AVAINEDUT:\n- Kohta 1\n- Kohta 2\n- Kohta 3\n"
        "SEO-AVAINSANAT: Lisää 5 relevanttia hakusanaa Googlea varten.\n"
        "Älä käytä emojeita tai värillisiä merkkejä. Erota tuotteet toisistaan viivalla (---)."
    ),
    "English": (
        "You are TextFabriken, an absolute world champion in e-commerce, digital marketing, and SEO copywriting for the Nordic and international market. "
        "ALWAYS write in English. "
        "You have been given a file with raw data or a list of products. Your task is to transform this list into extremely compelling, bold, and modern product descriptions in English. "
        "IMPORTANT: Never invent exact figures, dimensions, technical specifications, or performance values (e.g. battery life, dB levels, DPI, capacity in ml/liters) that are not present in the raw data you received. "
        "If a specific figure is missing from the source material, write qualitatively instead (e.g. 'long battery life' or 'powerful noise cancellation') without guessing an exact number. "
        "IMPORTANT: Never create more products than actually exist in the raw data. If a product name consists of several words, always treat the entire name as ONE single product – never split it into several separate products. "
        "Each product must be structured according to the following strict format:\n"
        "PRODUCT NAME (Use bold text)\n"
        "SELLING DESCRIPTION: Write around 100 words that create an extremely strong 'must-have' feeling in the customer.\n"
        "KEY BENEFITS:\n- Point 1\n- Point 2\n- Point 3\n"
        "SEO TAGS: Add 5 relevant keywords for Google.\n"
        "Do not use any emojis or colored bullet points whatsoever. Print the texts directly one after another, separated by a dash (---) between each product."
    ),
}

# --- RUBRIKORD PER SPRÅK (används för att dela upp den genererade texten i kolumner vid CSV/Excel-export) ---
PRODUKT_RUBRIKER = {
    "Svenska": {"desc": "SÄLJANDE BESKRIVNING:", "fordelar": "NYCKELFÖRDELAR:", "taggar": "SEO-TAGGAR:"},
    "Norsk":   {"desc": "SALGSBESKRIVELSE:", "fordelar": "NØKKELFORDELER:", "taggar": "SEO-STIKKORD:"},
    "Dansk":   {"desc": "SALGSBESKRIVELSE:", "fordelar": "NØGLEFORDELE:", "taggar": "SEO-TAGS:"},
    "Suomi":   {"desc": "MYYVÄ KUVAUS:", "fordelar": "AVAINEDUT:", "taggar": "SEO-AVAINSANAT:"},
    "English": {"desc": "SELLING DESCRIPTION:", "fordelar": "KEY BENEFITS:", "taggar": "SEO TAGS:"},
}

# Kolumnrubriker för CSV/Excel-exporten, per språk
EXPORT_KOLUMNER = {
    "Svenska": ["Produktnamn", "Säljande beskrivning", "Nyckelfördelar", "SEO-taggar"],
    "Norsk":   ["Produktnavn", "Salgsbeskrivelse", "Nøkkelfordeler", "SEO-stikkord"],
    "Dansk":   ["Produktnavn", "Salgsbeskrivelse", "Nøglefordele", "SEO-tags"],
    "Suomi":   ["Tuotenimi", "Myyvä kuvaus", "Avainedut", "SEO-avainsanat"],
    "English": ["Product name", "Selling description", "Key benefits", "SEO tags"],
}

def parsa_produkter(text, sprak):
    """Delar upp den genererade produkttexten i strukturerade rader (namn/beskrivning/fördelar/taggar)
       baserat på rubrikorden för det aktuella språket. Block som saknar en giltig beskrivningsrubrik
       (t.ex. varningstexten på slutet) hoppas över."""
    h = PRODUKT_RUBRIKER[sprak]
    block_lista = [b.strip() for b in text.split("---") if b.strip()]
    produkter = []
    for block in block_lista:
        if h["desc"] not in block:
            continue
        rader = [r for r in block.splitlines() if r.strip() != ""]
        namn = rader[0].strip().strip("*").strip() if rader else ""

        def extrahera(start_etikett, slut_etiketter):
            start_idx = block.find(start_etikett)
            if start_idx == -1:
                return ""
            start_pos = start_idx + len(start_etikett)
            slut_pos = len(block)
            for slut_etikett in slut_etiketter:
                p = block.find(slut_etikett, start_pos)
                if p != -1 and p < slut_pos:
                    slut_pos = p
            return block[start_pos:slut_pos].strip(" \n-*")

        beskrivning = extrahera(h["desc"], [h["fordelar"], h["taggar"]])
        fordelar = extrahera(h["fordelar"], [h["taggar"]])
        taggar = extrahera(h["taggar"], [])
        produkter.append([namn, beskrivning, fordelar, taggar])
    return produkter

def skapa_csv(produkter, kolumner):
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(kolumner)
    for rad in produkter:
        writer.writerow(rad)
    # utf-8-sig (BOM) så att Excel läser nordiska tecken (å,ä,ö,ø,æ) korrekt vid dubbelklick
    return output.getvalue().encode("utf-8-sig")

def skapa_xlsx(produkter, kolumner):
    wb = Workbook()
    ws = wb.active
    ws.title = "Produkter"
    fet_stil = Font(name="Arial", bold=True)
    normal_stil = Font(name="Arial")
    for col_idx, rubrik in enumerate(kolumner, start=1):
        cell = ws.cell(row=1, column=col_idx, value=rubrik)
        cell.font = fet_stil
    for rad_idx, rad in enumerate(produkter, start=2):
        for col_idx, varde in enumerate(rad, start=1):
            cell = ws.cell(row=rad_idx, column=col_idx, value=varde)
            cell.font = normal_stil
    bredder = [25, 60, 40, 35]
    for i, bredd in enumerate(bredder, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = bredd
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# --- DATABAS FÖR PERMANENT LAGRING AV SPARADE PRODUKTLISTOR (uppdelat per kund) ---
DB_FIL = "textfabriken.db"

def db_init():
    conn = sqlite3.connect(DB_FIL)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sessioner (
            kund_id TEXT NOT NULL,
            namn TEXT NOT NULL,
            messages TEXT,
            file_content TEXT,
            show_download INTEGER,
            PRIMARY KEY (kund_id, namn)
        )
    """)
    conn.commit()
    # Migrering: om tabellen finns sedan tidigare UTAN kund_id-kolumnen (gammal version av appen),
    # återskapas den med rätt struktur. Gammal osparad testdata utan kund_id kan inte återanvändas säkert.
    cursor.execute("PRAGMA table_info(sessioner)")
    kolumner = [rad[1] for rad in cursor.fetchall()]
    if "kund_id" not in kolumner:
        cursor.execute("DROP TABLE sessioner")
        cursor.execute("""
            CREATE TABLE sessioner (
                kund_id TEXT NOT NULL,
                namn TEXT NOT NULL,
                messages TEXT,
                file_content TEXT,
                show_download INTEGER,
                PRIMARY KEY (kund_id, namn)
            )
        """)
        conn.commit()
    conn.close()

def db_spara_session(kund_id, namn, messages, file_content, show_download):
    conn = sqlite3.connect(DB_FIL)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO sessioner (kund_id, namn, messages, file_content, show_download)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(kund_id, namn) DO UPDATE SET
            messages=excluded.messages,
            file_content=excluded.file_content,
            show_download=excluded.show_download
    """, (kund_id, namn, json.dumps(messages), file_content, int(show_download)))
    conn.commit()
    conn.close()

def db_hamta_alla_sessioner(kund_id):
    conn = sqlite3.connect(DB_FIL)
    cursor = conn.cursor()
    cursor.execute("SELECT namn, messages, file_content, show_download FROM sessioner WHERE kund_id = ?", (kund_id,))
    rader = cursor.fetchall()
    conn.close()
    sessioner = {}
    for namn, messages_json, file_content, show_download in rader:
        sessioner[namn] = {
            "messages": json.loads(messages_json),
            "file_content": file_content,
            "show_download": bool(show_download)
        }
    return sessioner

db_init()

# --- TILLDELA EN UNIK ARBETSYTEKOD PER BESÖKARE (håller sparade listor separata mellan kunder) ---
if "kod" in st.query_params:
    kund_id = st.query_params["kod"]
else:
    kund_id = uuid.uuid4().hex[:10]
    st.query_params["kod"] = kund_id
if "kund_id" not in st.session_state:
    st.session_state.kund_id = kund_id

# --- INSTÄLLNINGAR & MINNE ---
if "saved_sessions" not in st.session_state:
    st.session_state.saved_sessions = db_hamta_alla_sessioner(st.session_state.kund_id)  # Läs in ENDAST denna kunds sparade listor
if "current_session_name" not in st.session_state: st.session_state.current_session_name = "Aktuell produktlista"
if "messages" not in st.session_state: st.session_state.messages = []
if "generated_file_content" not in st.session_state: st.session_state.generated_file_content = ""
if "show_download" not in st.session_state: st.session_state.show_download = False
if "fil_bearbetad" not in st.session_state: st.session_state.fil_bearbetad = False
if "senast_uppladdad_fil" not in st.session_state: st.session_state.senast_uppladdad_fil = None
if "sprak" not in st.session_state: st.session_state.sprak = "Svenska"

t = UI_TEXTS[st.session_state.sprak]  # Genväg till aktuellt gränssnittsspråks texter

# Huvudsida (appnamnet TEXTFABRIKEN översätts ALDRIG)
st.title("🏭 TEXTFABRIKEN AI")
st.subheader(t["subheader"])
st.write(t["intro"])

# --- SIDOMENY ---
with st.sidebar:
    st.markdown("# 🏭 TEXTFABRIKEN")
    st.write("---")

    vald_sprak = st.selectbox(t["lang_label"], SPRAK_ALTERNATIV, index=SPRAK_ALTERNATIV.index(st.session_state.sprak))
    if vald_sprak != st.session_state.sprak:
        st.session_state.sprak = vald_sprak
        st.rerun()

    st.write("---")
    st.caption(t["workspace_info"].format(kod=st.session_state.kund_id))

    st.write("---")
    st.markdown(f"### {t['saved_header']}")
    if st.session_state.saved_sessions:
        for session_title in st.session_state.saved_sessions.keys():
            if st.button(f"📄 {session_title}"):
                st.session_state.messages = st.session_state.saved_sessions[session_title]["messages"]
                st.session_state.generated_file_content = st.session_state.saved_sessions[session_title]["file_content"]
                st.session_state.show_download = st.session_state.saved_sessions[session_title]["show_download"]
                st.session_state.current_session_name = session_title
                st.rerun()
    else:
        st.caption(t["no_saved"])
    st.write("---")
    st.caption(f"{t['active_now']} {st.session_state.current_session_name}")

# Visa historik
for message in st.session_state.messages:
    if message["role"] == "user":
        with st.chat_message("user", avatar="👤"): st.markdown(f"**{t['user_label']}:** {message['content']}")
    else:
        with st.chat_message("assistant", avatar="🏭"): st.markdown(f"**{t['assistant_label']}:** {message['content']}")

# --- FASTKLISTRAD INPUT & FILER I BOTTEN ---
uploaded_file = None
extratext = ""
copy_klick = False

with st.container():
    prompt = st.chat_input(t["chat_placeholder"])
    with st.expander(t["expander_label"]):
        col1, col2 = st.columns(2)
        with col1:
            uploaded_file = st.file_uploader(t["uploader_label"], type=["pdf", "txt", "docx"], label_visibility="collapsed")
        with col2:
            copy_klick = st.button(t["rocket_button"], use_container_width=True)

# Processa filen
if uploaded_file is not None:
    st.session_state.current_session_name = uploaded_file.name
    # Om det är en NY fil (annat namn än senast), återställ "bearbetad"-flaggan
    if uploaded_file.name != st.session_state.senast_uppladdad_fil:
        st.session_state.fil_bearbetad = False
        st.session_state.senast_uppladdad_fil = uploaded_file.name

    if uploaded_file.name.endswith(".pdf"):
        try:
            pdf_reader = PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                text_content = page.extract_text()
                if text_content: extratext += text_content + "\n"
        except Exception as e: st.error("Kunde inte läsa PDF.")
    elif uploaded_file.name.endswith(".docx"):
        try:
            doc_reader = Document(uploaded_file)
            for paragraph in doc_reader.paragraphs:
                if paragraph.text: extratext += paragraph.text + "\n"
        except Exception as e: st.error("Kunde inte läsa Word.")
    else:
        extratext = uploaded_file.read().decode("utf-8")

# STENSÄKRAD SAMMANKOPPLING MED GROQ (med automatiska omförsök vid rate limit)
def fraga_groq(system_prompt, user_prompt, forsok=5):
    url = "https://api.groq.com/openai/v1/chat/completions"
    headers = {
        "Authorization": "Bearer " + GROQ_API_KEY,
        "Content-Type": "application/json"
    }
    data = {
        "model": "openai/gpt-oss-120b",
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.3,
        "max_tokens": 8000
    }
    for i in range(forsok):
        try:
            respons = requests.post(url, json=data, headers=headers)
            if respons.status_code == 200:
                return respons.json()["choices"][0]["message"]["content"]
            elif respons.status_code == 429:
                vantetid = int(float(respons.headers.get("Retry-After", 10))) + 1
                time.sleep(vantetid)
                continue
            else:
                return "Anslutningsfel (Status " + str(respons.status_code) + ")"
        except Exception as e:
            return "Kunde inte skicka förfrågan."
    return "Groq-servern är överbelastad just nu (Status 429). Försök igen om en liten stund."

# --- BATCH-HANTERING FÖR STORA PRODUKTLISTOR ---
RADER_PER_BATCH = 15  # Sänkt från 40 för att undvika avkapade svar och för stora anrop (413)

def dela_upp_i_batchar(text, rader_per_batch=RADER_PER_BATCH):
    rader = [r for r in text.strip().split('\n') if r.strip() != ""]
    batchar = []
    for i in range(0, len(rader), rader_per_batch):
        batch = '\n'.join(rader[i:i + rader_per_batch])
        if batch.strip():
            batchar.append(batch)
    return batchar if batchar else [text]

def kor_massgenerering(extra_instruktion=""):
    """Delar upp produktdatan i batchar och kör Groq-anrop med progressbar.
       Används av både raketknappen och chattrutan för att undvika för stora anrop (413)."""
    direktiv = SEO_DIREKTIV[st.session_state.sprak]
    progress_bar = st.progress(0)
    status_text = st.empty()

    batchar = dela_upp_i_batchar(extratext)
    alla_svar = []
    for i, batch in enumerate(batchar):
        status_text.markdown(t["processing_batch"].format(i=i + 1, n=len(batchar)))
        anvandarprompt = "Produktlista/Rådata:\n" + batch
        if extra_instruktion:
            anvandarprompt = "Användarens extra instruktion: " + extra_instruktion + "\n\n" + anvandarprompt
        svar = fraga_groq(direktiv, anvandarprompt)
        alla_svar.append(svar)
        progress_bar.progress((i + 1) / len(batchar))
        if i < len(batchar) - 1:
            time.sleep(2)  # Kort paus mellan batchar för att undvika rate limit

    status_text.empty()
    progress_bar.empty()

    ai_svar = "\n\n---\n\n".join(alla_svar)
    return ai_svar + "\n\n---\n" + t["warning_text"]

# LOGIK FÖR RAKET-KNAPPEN
if copy_klick:
    if not extratext:
        st.error(t["error_no_file"])
    else:
        prompt_text = t["mass_prompt_text"]
        with st.chat_message("user", avatar="👤"): st.markdown(f"**{t['user_label']}:** {prompt_text}")
        st.session_state.messages.append({"role": "user", "content": prompt_text})
        with st.chat_message("assistant", avatar="🏭"):
            ai_svar_med_varning = kor_massgenerering()
            st.markdown(f"**{t['assistant_label']}:**\n\n{ai_svar_med_varning}")
            st.session_state.messages.append({"role": "assistant", "content": ai_svar_med_varning})
            st.session_state.generated_file_content = ai_svar_med_varning
            st.session_state.show_download = True
            st.session_state.fil_bearbetad = True  # Markera filen som klar - chatten blir nu fri
            st.session_state.saved_sessions[st.session_state.current_session_name] = {"messages": st.session_state.messages, "file_content": st.session_state.generated_file_content, "show_download": st.session_state.show_download}
            db_spara_session(st.session_state.kund_id, st.session_state.current_session_name, st.session_state.messages, st.session_state.generated_file_content, st.session_state.show_download)
            st.rerun()

# LOGIK FÖR VANLIGA CHATTRUTAN
if prompt:
    with st.chat_message("user", avatar="👤"): st.markdown(f"**{t['user_label']}:** {prompt}")
    st.session_state.messages.append({"role": "user", "content": prompt})

    # Skicka bara med hela produktlistan om filen INTE redan bearbetats via raketknappen
    anvand_produktdata = extratext and not st.session_state.fil_bearbetad

    with st.chat_message("assistant", avatar="🏭"):
        if anvand_produktdata:
            # Använd samma säkra batch-funktion som raketknappen, för att undvika för stora anrop (413)
            ai_svar = kor_massgenerering(extra_instruktion=prompt)
            st.markdown(f"**{t['assistant_label']}:**\n\n{ai_svar}")
            st.session_state.messages.append({"role": "assistant", "content": ai_svar})
            st.session_state.generated_file_content = ai_svar
            st.session_state.show_download = True
            st.session_state.fil_bearbetad = True
        else:
            message_placeholder = st.empty()
            message_placeholder.markdown(t["processing_single"])
            system_d = t["system_prompt_free_chat"]
            ai_svar = fraga_groq(system_d, prompt)
            message_placeholder.markdown(f"**{t['assistant_label']}:**\n\n{ai_svar}")
            st.session_state.messages.append({"role": "assistant", "content": ai_svar})
            # Gör svaret nedladdningsbart som Word-fil, precis som vid massgenerering
            st.session_state.generated_file_content = ai_svar
            st.session_state.show_download = True

        if uploaded_file is not None:
            st.session_state.saved_sessions[st.session_state.current_session_name] = {"messages": st.session_state.messages, "file_content": st.session_state.generated_file_content, "show_download": st.session_state.show_download}
            db_spara_session(st.session_state.kund_id, st.session_state.current_session_name, st.session_state.messages, st.session_state.generated_file_content, st.session_state.show_download)
        st.rerun()

# VISA NEDLADDNINGSKNAPPAR (Word, CSV, Excel)
if st.session_state.show_download and st.session_state.generated_file_content:
    st.write("---")
    st.markdown(t["download_header"])
    st.info(t["download_info"])
    rensad_text = st.session_state.generated_file_content.replace(f"**{t['assistant_label']}:**\n\n", "")

    col_word, col_csv, col_xlsx = st.columns(3)

    with col_word:
        doc = Document()
        doc.add_heading(t["doc_heading"], level=1)
        for rad in rensad_text.split('\n'): doc.add_paragraph(rad)
        bio_docx = io.BytesIO()
        doc.save(bio_docx)
        st.download_button(label=t["download_button"], data=bio_docx.getvalue(), file_name="textfabriken_produkter.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

    produkter = parsa_produkter(rensad_text, st.session_state.sprak)
    kolumner = EXPORT_KOLUMNER[st.session_state.sprak]

    if produkter:
        with col_csv:
            csv_data = skapa_csv(produkter, kolumner)
            st.download_button(label=t["download_button_csv"], data=csv_data, file_name="textfabriken_produkter.csv", mime="text/csv", use_container_width=True)

        with col_xlsx:
            xlsx_data = skapa_xlsx(produkter, kolumner)
            st.download_button(label=t["download_button_xlsx"], data=xlsx_data, file_name="textfabriken_produkter.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
