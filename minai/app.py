import streamlit as st
import requests
from pypdf import PdfReader
import io
from docx import Document

st.set_page_config(page_title="TextFabriken AI", page_icon="🏭", layout="centered", initial_sidebar_state="expanded")

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

# --- SIDOMENY (Nu helt renrakad på knappar!) ---
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
    
    # HÄR SKAPAR VI POPUP-MENYN I BOTTEN
    with st.expander("📎 Klicka här för att mata in produktfakta (PDF/TXT/WORD)"):
        # Vi delar upp rutan i två kolumner bredvid varandra!
        col1, col2 = st.columns([2, 1])
        with col1:
            uploaded_file = st.file_uploader("Välj dokument från din dator", type=["pdf", "txt", "docx"], label_visibility="collapsed")
        with col2:
            # HÄR LIGGER DEN NYA SMIDIGA STARTKNAPPEN BREVID UPLOAD!
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

# LOGIK FÖR KNAPPEN (KÖR I CHATTEN DIREKT)
if copy_klick:
    if not extratext:
        st.error("⚠️ Du måste välja en fil i rutan till vänster först!")
    else:
        prompt_text = "Massgenerera SEO-produktbeskrivningar för alla produkter i listan."
        with st.chat_message("user", avatar="👤"): st.markdown(f"**Du:** {prompt_text}")
        st.session_state.messages.append({"role": "user", "content": prompt_text})
        with st.chat_message("assistant", avatar="🏭"):
            message_placeholder = st.empty()
            message_placeholder.markdown("*Maskinerna i TextFabriken startar upp...*")
            try:
                respons = requests.post("http://localhost:11434/api/generate", json={"model": "llama3.1", "prompt": f"{seo_direktiv}\n\nProduktlista/Rådata:\n{extratext}", "stream": False})
                if respons.status_code == 200:
                    ai_svar = respons.json()['response']
                    message_placeholder.markdown(f"**TextFabriken:**\n\n{ai_svar}")
                    st.session_state.messages.append({"role": "assistant", "content": ai_svar})
                    st.session_state.generated_file_content = ai_svar
                    st.session_state.show_download = True
                    st.session_state.saved_sessions[st.session_state.current_session_name] = {"messages": st.session_state.messages, "file_content": st.session_state.generated_file_content, "show_download": st.session_state.show_download}
                    st.rerun()
            except Exception as e: message_placeholder.markdown("Kunde inte nå Ollama.")

# LOGIK FÖR VANLIGA CHATTRUTAN
if prompt:
    with st.chat_message("user", avatar="👤"): st.markdown(f"**Du:** {prompt}")
    st.session_state.messages.append({"role": "user", "content": prompt})

    if not extratext:
        system_direktiv = "Du är TextFabriken, en glad, vis och effektiv e-handelsassistent på svenska. Hälsa användaren välkommen och be dem klicka på gemet (📎) längst ner för att mata in sin produktfakta."
        full_prompt = f"{system_direktiv}\nAnvändare: {prompt}"
    else:
        full_prompt = f"{seo_direktiv}\n\nAnvändarens extra instruktion: {prompt}\n\nProduktdata:\n{extratext}"

    with st.chat_message("assistant", avatar="🏭"):
        message_placeholder = st.empty()
        message_placeholder.markdown("*TextFabriken bearbetar dina ord...*")
        try:
            respons = requests.post("http://localhost:11434/api/generate", json={"model": "llama3.1", "prompt": full_prompt, "stream": False})
            if respons.status_code == 200:
                ai_svar = respons.json()['response']
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
        except Exception as e: message_placeholder.markdown("Kunde inte nå Ollama.")

# VISA WORD-KNAPP
if st.session_state.show_download and st.session_state.generated_file_content:
    st.write("---")
    st.markdown("### 📥 Din färdiga Word-fil från TextFabriken är klar!")
    doc = Document()
    doc.add_heading("SEO Produktbeskrivningar - TextFabriken AI", level=1)
    rensad_text = st.session_state.generated_file_content.replace("**TextFabriken:**\n\n", "")
    for rad in rensad_text.split('\n'): doc.add_paragraph(rad)
    bio = io.BytesIO()
    doc.save(bio)
    st.download_button(label="📝 Ladda ner produkttexter (.docx)", data=bio.getvalue(), file_name="textfabriken_produkter.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
