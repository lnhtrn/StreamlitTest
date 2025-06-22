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

# Use text_area key to modify it 
def on_upper_updated():
    st.session_state.transcript = st.session_state.transcript.upper()

def on_upper_reload(new_text):
    st.session_state.transcript = new_text

whisper_model = load_whisper_model()

# Record audio
audio_data = st.audio_input("Speak something to transcribe")

if audio_data:
    # Save audio
    with open("temp.wav", "wb") as f:
        f.write(audio_data.getvalue())

    # Transcribe
    with st.spinner("Transcribing...", show_time=True):
        result = whisper_model.transcribe("temp.wav")

    st.markdown("**Transcription:**")
    editable_trans = st.text_area(
        "Verify and edit transcription", 
        result['text'],
        key="transcript"
    )
    
if st.button("Show final text"):
    st.markdown("**Finalized text:**")
    on_upper_updated()
    st.write(editable_trans)