import streamlit as st
import whisper

##########################################################
st.set_page_config(
    page_title="Audio Testing with Whisper",
    page_icon="🎧",
    layout="centered",
    initial_sidebar_state="expanded",
)

##########################################################
# Load Whisper model
@st.cache_resource
def load_whisper_model():
    return whisper.load_model("base")

whisper_model = load_whisper_model()

# Record audio
audio_data = st.audio_input("Speak something to transcribe")

if audio_data:
    # Save audio
    with open("temp.wav", "wb") as f:
        f.write(audio_data.getvalue())

    # Transcribe
    st.write("Transcribing...")
    result = whisper_model.transcribe("temp.wav")
    text = result['text']
    st.markdown("**Transcription:**")
    st.text_area("Verify and edit transcription", text)