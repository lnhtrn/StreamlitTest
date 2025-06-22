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

st.markdown("**Transcription:**")
editable_trans = st.text_area("Verify and edit transcription", "")

if audio_data:
    # Save audio
    with open("temp.wav", "wb") as f:
        f.write(audio_data.getvalue())

    # Transcribe
    with st.spinner("Transcribing...", show_time=True):
        result = whisper_model.transcribe("temp.wav")
    editable_trans = result['text']
    
if st.button("Show final text"):
    st.markdown("**Finalized text:**")
    st.write(editable_trans)