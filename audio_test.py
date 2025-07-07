import streamlit as st
import whisper
from openai import OpenAI
import yaml

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
# def on_upper_updated():
#     st.session_state.transcript = st.session_state.transcript.upper()

# def on_upper_reload(new_text):
#     st.session_state.transcript = new_text

# Session state keys: 'openai_output', 'final_text'
if 'openai_output' not in st.session_state:
    st.session_state.openai_output = ""
if 'final_text' not in st.session_state:
    st.session_state.final_text = ""
data = {}

whisper_model = load_whisper_model()

# Load OpenAI client 
client = OpenAI(api_key=st.secrets["openai_key"])

# Record audio
audio_data = st.audio_input("Speak something to transcribe")
transcript_data = None
editable_trans = ""

if st.button("Transcribe"):
    if audio_data:
        # Save audio
        with open("temp.wav", "wb") as f:
            f.write(audio_data.getvalue())

        # Transcribe
        with st.spinner("Transcribing...", show_time=True):
            result = whisper_model.transcribe("temp.wav")

        st.markdown("## Transcription:")
        st.write(result['text'])
        # editable_trans = st.text_area(
        #     "Verify and edit transcription", 
        #     result['text'],
        #     key="transcript"
        # )

        response = client.responses.create(
            prompt={
                "id": "pmpt_685c1d7f4f4c819083a45722b78921830b7eea852e8a39f1",
                "version": "1",
                "variables": {
                "transcription": result['text']
                }
            }       
        )
        # with open("temp.txt", "w") as file:
        #     file.write(response.output_text)
        st.session_state.openai_output = response.output_text

with st.form('EditResponse'):
    st.header("Edit OpenAI Response")

    st.markdown("## OpenAI Response:")
    editable_trans = st.text_area(
        "Edit OpenAI response before submitting the form", 
        st.session_state.openai_output,
        height=200,
    )

    
    data['{{Residence City/State}}'] = st.text_input("Residence City/State")
    # st.selectbox(
    #     "Residence City/State", states, index=None,
    # )

    data['{{Narrative}}'] = st.text_area('Narrative to finish \"Patient lives with...\"')
    
    submit = st.form_submit_button('Submit')
    
if submit:
    st.session_state.final_text = editable_trans
    # Display data 
    yaml_string = yaml.dump(data, sort_keys=False)
    yaml_data = st.code(yaml_string, language=None)

# Step 4: Show final output
if st.session_state.final_text:
    st.header("3. Final Output")
    st.write(st.session_state.final_text)