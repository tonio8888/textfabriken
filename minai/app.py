import streamlit as st
import requests
from pypdf import PdfReader
import io
from docx import Document

st.set_page_config(page_title="TextFabriken AI", page_icon="🏭", layout="centered", initial_sidebar_state="expanded")

# --- SÄKRAD API-NYCKEL (Hämtas från ditt Streamlit-valv) ---
GROQ_API_KEY = st.secrets["GROQ_API_KEY"]

# --- INSTÄLLNINGAR & MINNE ---
if "saved_sessions" not in st.session_state: st.session_state.saved_sessions = {}
if "current_session_name" not in st.session_state: st.session_state.current_session_name = "Aktuell produktlista"
if "messages" not in st.session_state: st.session_state.messages = []
if "generated_file_content" not in st.session_state: st.session_state.generated_file_content = ""
if "show_download" not in st.session_state: st.session_state.show_download = False

# Huvudsida
st.title("🏭 TEXTFABRIKEN AI")
st.subheader("Nordens smartaste löpande band for produktbeskrivningar")
st.write("Ladda upp din rådata i bottenmenyn. TextFabriken transformerar den till säljande SEO-texter och skapar en färdig Word-fil åt dig!")

# --- SIDOMENY ---
with st.sidebar:
    st.markdown("# 🏭 TEXTFABRIKEN")
    st.write("---")
    st.markdown("### 📁 Sparade produktlistor")
    if st.session_state.saved_sessions:
        for session_title in st.session_state.saved_sessions.keys():
            if st.button(f"📄 {session_title}"):
                st.session_state.messages = st.session_state.saved_sessions[session_title]["messages"]
                st.session_state.generated_file_content = st.session_state.saved_sessions[session_title]["file_content"]
                st.session_state.show_download = st.session_state.saved_sessions[session_title]["show_download"]
                st.session_state.current_session_name = session_title
                st.rerun()
    else:
        st.caption("Inga sparade listor ännu.")
    st.write("---")
    st.caption(f"Aktiv nu: {st.session_state.current_session_name}")

# Visa historik
for message in st.session_state.messages:
    if message["role"] == "user":
        with st.chat_message("user", avatar="👤"): st.markdown(f"**Du:** {message['content']}")
    else:
        with st.chat_message("assistant", avatar="🏭"): st.markdown(f"**TextFabriken:** {message['content']}")

# DET STENHÅRDA NORDISKA MARKNADSFÖRINGS-DIREKTIVET
seo_direktiv = (
    "Du är TextFabriken, en absolut världsmästare på e-handel, digital marknadsföring och SEO-copywriting för den nordiska marknaden. "
    "Du har fått en fil med rådata eller en lista på produkter. Din uppgift är att transformera denna lista till supersäljande, kaxiga och moderna produktbeskrivningar på svenska. "
    "VIKTIGT: Hitta aldrig på exakta siffror, mått, tekniska specifikationer eller prestandavärden (t.ex. batteritid, dB-nivåer, DPI, kapacitet i ml/liter) som inte finns i den rådata du fått. "
    "Om en specifik siffra saknas i underlaget, skriv istället kvalitativt (t.ex. 'lång batteritid' eller 'kraftfull brusreducering') utan att gissa ett exakt tal. "
    "Varje produkt ska delas upp enligt följande strikta struktur:\n"
    "PRODUKTNAMN (Använd fetstil)\n"
    "SÄLJANDE BESKRIVNING: Skriv cirka 100 ord som skapar ett extremt starkt 'ha-begär' hos kunden.\n"
    "NYCKELFÖRDELAR:\n- Punkt 1\n- Punkt 2\n- Punkt 3\n"
    "SEO-TAGGAR: Lägg till 5 relevanta sökord för Google.\n"
    "Använd absolut inga emojier eller färgade prickar. Skriv ut texterna direkt efter varandra, separation med ett streck (---) mellan varje produkt."
)

# --- FASTKLISTRAD INPUT & FILER I BOTTEN ---
uploaded_file = None
extratext = ""
copy_klick = False

with st.container():
    prompt = st.chat_input("Skriv instruktion till fabriken...")
    with st.expander("📎 Klicka här för att mata in produktfakta (PDF/TXT/WORD)"):
        col1, col2 = st.columns(2)
        with col1:
            uploaded_file = st.file_uploader("Välj dokument från din dator", type=["pdf", "txt", "docx"], label_visibility="collapsed")
        with col2:
            copy_klick = st.button("🚀 Starta Massgenerering", use_container_width=True)

# Processa filen
if uploaded_file is not None:
    st.session_state.current_session_name = uploaded_file.name
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

# STENSÄKRAD SAMMANKOPPLING MED GROQ
def fraga_groq(system_prompt, user_prompt):
    try:
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
            "temperature": 0.3
        }
        respons = requests.post(url, json=data, headers=headers)
        if respons.status_code == 200:
            return respons.json()["choices"][0]["message"]["content"]
        else:
            return "Anslutningsfel (Status " + str(respons.status_code) + ")"
    except Exception as e:
        return "Kunde inte skicka förfrågan."

# LOGIK FÖR RAKET-KNAPPEN
if copy_klick:
    if not extratext:
        st.error("⚠️ Du måste välja en fil i rutan till vänster först!")
    else:
        prompt_text = "Massgenerera SEO-produktbeskrivningar för alla produkter i listan."
        with st.chat_message("user", avatar="👤"): st.markdown(f"**Du:** {prompt_text}")
        st.session_state.messages.append({"role": "user", "content": prompt_text})
        with st.chat_message("assistant", avatar="🏭"):
            message_placeholder = st.empty()
            message_placeholder.markdown("*Maskinerna i TextFabriken startar upp i molnet...*")
            
            ai_svar = fraga_groq(seo_direktiv, "Produktlista/Rådata:\n" + extratext)
            
            ai_svar_med_varning = ai_svar + "\n\n---\n⚠️ **Kontrollera alltid siffror och specifikationer** (t.ex. batteritid, mått, prestanda) mot din egen produktdata innan du publicerar texterna."
            message_placeholder.markdown(f"**TextFabriken:**\n\n{ai_svar_med_varning}")
            st.session_state.messages.append({"role": "assistant", "content": ai_svar_med_varning})
            st.session_state.generated_file_content = ai_svar_med_varning
            st.session_state.show_download = True
            st.session_state.saved_sessions[st.session_state.current_session_name] = {"messages": st.session_state.messages, "file_content": st.session_state.generated_file_content, "show_download": st.session_state.show_download}
            st.rerun()

# LOGIK FÖR VANLIGA CHATTRUTAN
if prompt:
    with st.chat_message("user", avatar="👤"): st.markdown(f"**Du:** {prompt}")
    st.session_state.messages.append({"role": "user", "content": prompt})

    if not extratext:
        system_d = "Du är TextFabriken, en glad, vis och effektiv e-handelsassistent på svenska. Hälsa användaren välkommen till TextFabriken och be dem klicka på gemet (📎) längst ner för att mata in sin produktfakta."
        user_d = prompt
    else:
        system_d = seo_direktiv
        user_d = "Användarens extra instruktion: " + prompt + "\n\nProduktdata:\n" + extratext

    with st.chat_message("assistant", avatar="🏭"):
        message_placeholder = st.empty()
        message_placeholder.markdown("*TextFabriken bearbetar dina ord i molnet...*")
        
        ai_svar = fraga_groq(system_d, user_d)
        
        if extratext:
            ai_svar = ai_svar + "\n\n---\n⚠️ **Kontrollera alltid siffror och specifikationer** mot din egen produktdata innan du publicerar texterna."
        
        message_placeholder.markdown(f"**TextFabriken:**\n\n{ai_svar}")
        st.session_state.messages.append({"role": "assistant", "content": ai_svar})
        
        if extratext:
            st.session_state.generated_file_content = ai_svar
            st.session_state.show_download = True
        else:
            st.session_state.show_download = False
        
        if uploaded_file is not None:
            st.session_state.saved_sessions[st.session_state.current_session_name] = {"messages": st.session_state.messages, "file_content": st.session_state.generated_file_content, "show_download": st.session_state.show_download}
        st.rerun()

# VISA WORD-KNAPP
if st.session_state.show_download and st.session_state.generated_file_content:
    st.write("---")
    st.markdown("### 📥 Din färdiga Word-fil från TextFabriken är klar!")
    st.info("💡 Innan du publicerar: dubbelkolla alla siffror (mått, kapacitet, batteritid, dB-nivåer m.m.) mot leverantörens originaldata. TextFabriken skriver säljande texter, men ansvarar inte för att specifikationerna är korrekta.")
    doc = Document()
    doc.add_heading("SEO Produktbeskrivningar - TextFabriken AI", level=1)
    rensad_text = st.session_state.generated_file_content.replace("**TextFabriken:**\n\n", "")
    for rad in rensad_text.split('\n'): doc.add_paragraph(rad)
    bio = io.BytesIO()
    doc.save(bio)
    st.download_button(label="📝 Ladda ner produkttexter (.docx)", data=bio.getvalue(), file_name="textfabriken_produkter.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
