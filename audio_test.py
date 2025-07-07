import streamlit as st
import whisper
from openai import OpenAI
import yaml
from docx import Document
import io

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

# Session state keys: 
if 'behavior_observation' not in st.session_state:
    st.session_state.behavior_observation = ""
if 'development_history' not in st.session_state:
    st.session_state.development_history = ""
if 'final_text' not in st.session_state:
    st.session_state.final_text = ""
data = {}

whisper_model = load_whisper_model()

# Load OpenAI client 
client = OpenAI(api_key=st.secrets["openai_key"])

##################################################################
def transcribe_audio(audio_file):
    if audio_file:
        with open("temp.wav", "wb") as f:
            f.write(audio_behavior.getvalue())

        # Transcribe
        with st.spinner("Transcribing...", show_time=True):
            result = whisper_model.transcribe("temp.wav")

        st.markdown("## Transcription:")
        st.write(result['text'])
        
        return result['text']

# Record audio
# audio_data = st.audio_input("Speak something to transcribe")

audio_behavior = st.audio_input("Behavioral Observation")
audio_development = st.audio_input("Developmental History")

if st.button("Transcribe"):
    if audio_behavior and audio_development:
        transcript_behavior = transcribe_audio(audio_behavior)
        transcript_development = transcribe_audio(audio_development)
        
        response = client.responses.create(
            prompt={
                "id": "pmpt_685c1d7f4f4c819083a45722b78921830b7eea852e8a39f1",
                "version": "1",
                "variables": {
                "transcription": transcript_behavior
                }
            }       
        )
        st.session_state.behavior_observation = response.output_text

        response = client.responses.create(
            prompt={
                "id": "pmpt_685c1d7f4f4c819083a45722b78921830b7eea852e8a39f1",
                "version": "1",
                "variables": {
                "transcription": transcript_development
                }
            }       
        )
        st.session_state.development_history = response.output_text

with st.form('EditResponse'):
    st.header("Edit OpenAI Response")

    # st.markdown("**Behavioral Observation:**")
    data['behavior_observation'] = st.text_area(
        "Behavioral Observation: Edit the response before submitting the form", 
        st.session_state.behavior_observation,
        height=200,
    )

    # st.markdown("**Developmental History:**")
    data['development_history'] = st.text_area(
        "Developmental History: Edit the response before submitting the form", 
        st.session_state.development_history,
        height=200,
    )
    
    data['{{Residence City/State}}'] = st.text_input("Residence City/State")
    # st.selectbox(
    #     "Residence City/State", states, index=None,
    # )

    data['{{Narrative}}'] = st.text_area('Narrative to finish \"Patient lives with...\"')
    
    submit = st.form_submit_button('Submit')
    
if submit:
    st.session_state.final_text = data['Transcription']
    # Display data 
    yaml_string = yaml.dump(data, sort_keys=False)
    yaml_data = st.code(yaml_string, language=None)

    #### Edit document 
    doc = Document('templates/template_mod_12.docx')
    if doc:
        doc.add_paragraph(data['behavior_observation'])
        doc.add_paragraph(data['development_history'])

        # Save content to file
        filename = "Test_audio.docx"
        doc.save(filename)

        # Download 
        bio = io.BytesIO()
        document = Document(filename)
        document.save(bio)

        st.download_button(
            label="Click here to download",
            key="report_download",
            data=bio.getvalue(),
            file_name=filename,
            mime="docx"
        )

# # Step 4: Show final output
# if st.session_state.final_text:
#     st.header("3. Final Output")
#     st.write(st.session_state.final_text)