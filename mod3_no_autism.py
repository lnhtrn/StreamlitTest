import streamlit as st
from docx import Document
import docx
import yaml
import io
import docxedit
import datetime
from docx.enum.style import WD_STYLE_TYPE
from streamlit_gsheets import GSheetsConnection
from docxtpl import DocxTemplate
from docx.shared import Inches, Pt
from docx.oxml.shared import OxmlElement, qn
from openai import OpenAI

##########################################################
st.set_page_config(
    page_title="Module 3 No Autism",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="expanded",
)

##########################################################
# Set up OpenAI 
if 'behavior_observation_mod3_no_autism' not in st.session_state:
    st.session_state.behavior_observation_mod3_no_autism = ""
if 'development_history_mod3_no_autism' not in st.session_state:
    st.session_state.development_history_mod3_no_autism = ""

# Load OpenAI client 
client = OpenAI(api_key=st.secrets["openai_key"])

##################################################################
def transcribe_audio(audio_file, name='temp'):
    if audio_file:
        # Transcribe
        with st.spinner("Transcribing...", show_time=True):
            # result = whisper_model.transcribe(f"{name}.wav")
            result = client.audio.transcriptions.create(
                model="whisper-1", 
                file=audio_file, 
                response_format="text"
            )
        return result 

##########################################################
# Access Google Sheets

dropdowns = {}
connections = {}

# Create a connection object.
connections['All'] = st.connection(f"mod3_all", type=GSheetsConnection)
# Read object
df = connections['All'].read(
    ttl="30m",
    usecols=list(range(6)),
    nrows=30,
) 
for col_name in df.columns:
    dropdowns[col_name] = df[col_name].tolist()
    dropdowns[col_name] = [x for x in dropdowns[col_name] if str(x) != 'nan']

# DSM dropdowns
connections['DSM'] = st.connection(f"dsm", type=GSheetsConnection)
# Read object
df = connections['DSM'].read(
    ttl="30m",
    usecols=list(range(7)),
    nrows=15,
) 
for col_name in df.columns:
    dropdowns[col_name] = df[col_name].tolist()
    dropdowns[col_name] = [x for x in dropdowns[col_name] if str(x) != 'nan']
    dropdowns[col_name].append("None")

##################################################
# Set up side bar
def clear_my_cache():
    st.cache_data.clear()

with st.sidebar:
    st.markdown("**After editing dropdown options, please reload data using the button below to update within the form.**")
    st.link_button("Edit Dropdown Options", st.secrets['mod3_spreadsheet'])
    st.button('Reload Dropdown Data', on_click=clear_my_cache)

    # Display data 
    # yaml_dropdown = yaml.dump(dropdowns, sort_keys=False)
    # st.code(yaml_dropdown, language=None)
    
    ####################################################
    st.markdown("**Check to include score in the form:** Scores to report:")
    scq_result = st.checkbox("Social Communication Questionnaire (SCQ) - Lifetime Form")
    teacher_srs_eval = st.checkbox("Teacher's SRS Scores")
    caregiver_srs_eval = st.checkbox("Teacher's SRS Scores")
    caregiver_vineland_eval = st.checkbox("Caregiver's Vineland Adaptive Behavior Scales")
    teacher_vineland_eval = st.checkbox("Teacher's Vineland Adaptive Behavior Scales")
    wppsi_score = st.checkbox("Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed. (WPPSI) Score")
    dppr_score = st.checkbox("Developmental Profile – Fourth Edition - Parent Report (DPPR)")
    pls_score = st.checkbox("Preschool Language Scale - Fifth Edition (PLS)")
    pdms_score = st.checkbox("Peabody Developmental Motor Scales - Second Edition")
    peshv_score = st.checkbox("Preschool Evaluation Scale Home Version - Second Edition")
    reelt_score = st.checkbox("Receptive Expressive Emergent Language Test - Fourth Edition")
    abas_score = st.checkbox("Adaptive Behavior Assessment System - Third Edition")


col1,col2 = st.columns(2)
col1.title('Module 3 No Autism Report Builder')

def format_date_with_ordinal(date_obj):
    day = date_obj.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return date_obj.strftime(f"%B {day}{suffix}, %Y")

# Set up dictionary to store data 
data = {}
optional = {}
teacher_score = {}
bullet = {}
lines = {}
comma = {}

####################################################
st.header("Appointment Summary")

data['{{Patient First Name}}'] = st.text_input('Patient First Name')

data['{{Patient Last Name}}'] = st.text_input('Patient Last Name')

preferred = st.selectbox(
    "Patient's Preferred Pronoun",
    ("They/them", "He/him", "She/her"),
)

data['{{Location of the evaluation}}'] = st.radio(
    "Location of the evaluation",
    ['home', 'school', 'the office'],
    index=None,
)

# Audio section 
st.markdown(f"**Behavioral Observation:** Things to mention: eye contact, attention to task, social affect and restricted and repetitive behavior.")
audio_behavior = st.audio_input("Behavioral Observation")
if audio_behavior:
    # 3. Create a download button
    st.download_button(
        label="Download Behavioral Observation Recording",
        key="audio_behavior",
        data=audio_behavior,
        file_name=f"{data['{{Patient First Name}}']} {data['{{Patient Last Name}}']} - Behavioral Observation.wav",
        mime="audio/wav",
    )

st.markdown(f"**Developmental History:** Things to mention: social communication skills, repetitive behavior and related behavioral concerns.")
audio_development = st.audio_input("Developmental History")
if audio_development:
    # 3. Create a download button
    st.download_button(
        label="Download Developmental History Recording",
        key="audio_development",
        data=audio_development,
        file_name=f"{data['{{Patient First Name}}']} {data['{{Patient Last Name}}']} - Developmental History.wav",
        mime="audio/wav",
    )

if st.button("Transcribe"):
    if audio_behavior and audio_development:
        transcript_behavior = transcribe_audio(audio_behavior, name='behavior')
        st.markdown(f"**Transcription:** {transcript_behavior}")

        transcript_development = transcribe_audio(audio_development, name='development')
        st.markdown(f"**Transcription:** {transcript_development}")
        
        response = client.responses.create(
            prompt={
                "id": st.secrets["behavior_prompt_mod3_id"],
                # "version": "3",
                "variables": {
                    "first_name": data['{{Patient First Name}}'],
                    "pronouns": preferred,
                    "evaluation_location": data['{{Location of the evaluation}}'],
                    "transcription": transcript_behavior
                }
            }
        )
        st.session_state.behavior_observation_mod3_no_autism = response.output_text

        response = client.responses.create(
            prompt={
                "id": st.secrets["development_prompt_mod3_id"],
                # "version": "5",
                "variables": {
                    "first_name": data['{{Patient First Name}}'],
                    "pronouns": preferred,
                    "transcription": transcript_development
                }
            }
        )
        st.session_state.development_history_mod3_no_autism = response.output_text
   
####################################################
with st.form('BasicInfo'):
    st.header("Patient's data")

    data["{{Patient Age}}"] = st.number_input("Patient's Age", 0, 100)
    data['{{Patient age unit}}'] = st.radio(
        "Year/month?",
        ("year", "month")
    )

    data['{{Caregiver type}}'] = st.selectbox(
        "Patient's Caregiver",
        ("mother", "father", "parent", "grandparent", "legal custodian", "foster parent"),
        placeholder="Select from the choices or enter a new one",
        index=None,
        accept_new_options=True,
    )

    bullet['CaregiverPrimaryConcerns'] = st.multiselect(
        "Caregiver\'s Primary Concerns",
        dropdowns["Caregiver\'s Primary Concerns"],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )
    
    data['{{Residence City/State}}'] = st.text_input("Residence City/State")

    data['{{Narrative}}'] = st.text_area('Narrative to finish \"Patient lives with...\"')

    ##########################################################
    st.header("BRH Evaluation Details")

    data['{{Evaluation Date}}'] = format_date_with_ordinal(st.date_input("Evaluation Date"))

    data['{{Results Shared Date}}'] = format_date_with_ordinal(st.date_input("Results Shared Date"))
    
    data['{{Date Report Sent to Patient}}'] = format_date_with_ordinal(st.date_input("Date Report Sent to Patient"))
    
    #########################################################
    st.header("Medical/Developmental History")
    
    data['{{Developmental History}}'] = st.text_input(
        "Developmental History")

    data['{{Medical Diagnoses}}'] = st.text_input(
        "Medical Diagnoses")

    lines['{{Medications}}'] = st.multiselect(
        "Medications",
        ['None noted or reported.'],
        placeholder="Can input multiple options",
        accept_new_options=True
    )

    ##########################################################
    if scq_result:
        st.header("(SCQ) - Lifetime Form")
        data["Results (SCQ) Lifetime Form"] = st.text_input(
            "Results (SCQ) - Lifetime Form"
        )

    ##########################################################
    if caregiver_srs_eval:
        st.header("SRS-2 Caregiver Score")

        data["{{SRS-2 Score Caregiver}}"] = st.text_input("Caregiver's SRS-2 Score")
        
        data["{{Social Communication and Interaction Score Caregiver}}"] = st.text_input("Social Communication and Interaction Score Caregiver")
        
        data["{{Restricted Interests and Repetitive Behavior Score Caregiver}}"] = st.text_input("Restricted Interests and Repetitive Behavior Score Caregiver")

        data["{{Caregiver's level of concern}}"] = st.radio(
            "Caregiver's level of concern",
            ['no', 'mild', 'moderate', 'severe']
        )

        data["{{Evaluator's level of concern}}"] = st.radio(
            "Evaluator's level of concern",
            ['no', 'mild', 'moderate', 'severe']
        )

    ##########################################################
    if teacher_srs_eval:
        st.header("Teacher SRS Score")
        st.markdown("*Skip this section if teacher did not give SRS Score*")

        teacher_score['{{SRS-2 Score Teacher}}'] = st.text_input("Teacher's SRS-2 Score")

        teacher_score['{{Social Communication and Interaction Score Teacher}}'] = st.text_input("Social Communication and Interaction Score Teacher")

        teacher_score['{{Restricted Interests and Repetitive Behavior Score Teacher}}'] = st.text_input("Restricted Interests and Repetitive Behavior Score Teacher")

        teacher_score["{{Teacher level of concern}}"] = st.radio(
            "Teacher's level of concern",
            ['no', 'mild', 'moderate', 'severe']
        )
    
    ##########################################################
    if caregiver_vineland_eval:
        st.header("Vineland Adaptive Behavior Scales - Caregiver's Scores")

        data["{{Vineland Score Caregiver}}"] = st.text_input("Vineland Score Caregiver")
        
        data["{{Communication Score Caregiver}}"] = st.text_input("Communication Score Caregiver")
        
        data["{{Daily Living Skills Score Caregiver}}"] = st.text_input("Daily Living Skills Score Caregiver")
        
        data["{{Socialization Score Caregiver}}"] = st.text_input("Socialization Score Caregiver")

    if teacher_vineland_eval:
        st.header("Vineland Adaptive Behavior Scales - Teacher's Scores")

        data["{{Vineland Score Teacher}}"] = st.text_input("Vineland Score Teacher")
        
        data["{{Communication Score Teacher}}"] = st.text_input("Communication Score Teacher")
        
        data["{{Daily Living Skills Score Teacher}}"] = st.text_input("Daily Living Skills Score Teacher")
        
        data["{{Socialization Score Teacher}}"] = st.text_input("Socialization Score Teacher")

    ###############################################
    st.header("Educational Background")

    data['{{School District}}'] = st.selectbox(
        "School District",
        ['Rochester City'],
        index=None,
        placeholder="Select a school district or enter a new one",
        accept_new_options=True,
    )

    # data['{{School Name}}'] = st.text_input("School Name")

    if teacher_srs_eval or teacher_vineland_eval:
        teacher_score['{{Teacher name, title}}'] = st.text_input("Teacher name, title")

    data['{{Grade}}'] = st.text_input(
        "Grade",
    )

    data['School Year'] = st.text_input(
        "School Year",
    )

    data['{{Education Setting}}'] = st.selectbox(
        "Education Setting",
        ["General Education", "Integrated Co-Taught", "12:1:1", "8:1:1", "6:1:1"],
        index=None,
        placeholder="Select a grade or enter a new one",
        accept_new_options=True,
    )

    comma['{{Services}}'] = st.multiselect(
        "Services",
        dropdowns['Services'],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    comma['{{Classification}}'] = st.multiselect(
        "Classification",
        dropdowns['Classification'],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    ##########################################################
    if wppsi_score:
        st.header("Wechsler Preschool & Primary Scales of Intelligence - Fourth Ed. (WPPSI)")
        st.markdown("*Skip this section if there is no WPPSI Score*")
        optional["wppsi"] = {}

        optional["wppsi"]["Test Date"] = st.date_input("WPPSI Test Date").strftime("%m/%Y")
        optional["wppsi"]['WPPSI Full Scale IQ Score'] = st.text_input("WPPSI Full Scale IQ Score")

        optional["wppsi"]['WPPSI Verbal Comprehension Score'] = st.text_input("WPPSI Verbal Comprehension Score")

        optional["wppsi"]['WPPSI Visual Spatial Score'] = st.text_input("WPPSI Visual Spatial Score")
    
    if dppr_score:
        st.header("Developmental Profile - Fourth Edition - Parent Report (DPPR)")
        st.markdown("*Skip this section if there is no DPPR Score*")
        optional["dppr"] = {}

        optional["dppr"]["Test Date"] = st.date_input("DPPR Test Date").strftime("%m/%Y")
        optional["dppr"]['DPPR Cognitive Score'] = st.text_input("DPPR Cognitive Score")
        optional["dppr"]['DPPR Social-Emotional Score'] = st.text_input("DPPR Social-Emotional Score")
        optional["dppr"]['DPPR Adaptive Score'] = st.text_input("DPPR Adaptive Score")
        optional["dppr"]['DPPR Physical Score'] = st.text_input("DPPR Physical Score")
    
    if pls_score:
        st.header("Preschool Language Scale - Fifth Edition (PLS)")
        st.markdown("*Skip this section if there is no PLS Score*")
        optional["pls"] = {}
        optional["pls"]["Test Date"] = st.date_input("PLS Test Date").strftime("%m/%Y")
        optional["pls"]['PLS Total Language Score'] = st.text_input("PLS Total Language Score")
        optional["pls"]['PLS Auditory Comprehension Score'] = st.text_input("PLS Auditory Comprehension Score")
        optional["pls"]['PLS Expressive Communication Score'] = st.text_input("PLS Expressive Communication Score")

    if pdms_score:
        st.header("Peabody Developmental Motor Scales - Second Edition (PDMS)")
        st.markdown("*Skip this section if there is no PDMS Score*")
        optional["pdms"] = {}
        optional["pdms"]["Test Date"] = st.date_input("Test Date").strftime("%m/%Y")
        optional["pdms"]['PDMS Gross Motor Score'] = st.text_input("PDMS Gross Motor Score")
        optional["pdms"]['PDMS Fine Motor Score'] = st.text_input("PDMS Fine Motor Score")

    if peshv_score:
        st.header("Preschool Evaluation Scale Home Version - Second Edition (PESHV)")
        st.markdown("*Skip this section if there is no PESHV Score*")
        optional["peshv"] = {}
        optional["peshv"]["Test Date"] = st.date_input("PESHV Test Date").strftime("%m/%Y")
        optional["peshv"]['PESHV Cognitive Score'] = st.text_input("PESHV Cognitive Score")
        optional["peshv"]['PESHV Social Emotional Score'] = st.text_input("PESHV Social Emotional Score")
    
    if peshv_score:
        st.header("Preschool Evaluation Scale Home Version – Second Edition (PESHV)")
        st.markdown("*Skip this section if there is no PESHV Score*")
        optional[""] = {}
        optional["peshv"]["Test Date"] = st.date_input("PESHV Test Date").strftime("%m/%Y")
        optional["peshv"]['Total Language'] = st.text_input("Total Language")
        optional["peshv"]['PESHV Social Emotional Score'] = st.text_input("PESHV Social Emotional Score")

    if reelt_score:
        st.header("Receptive Expressive Emergent Language Test - Fourth Edition (REELT)")
        st.markdown("*Skip this section if there is no REELT Score*")
        optional["reelt"] = {}
        optional["reelt"]["Test Date"] = st.date_input("REELT Test Date").strftime("%m/%Y")
        optional["reelt"]['REELT Total Language Score'] = st.text_input("REELT Total Language Score")
        optional["reelt"]['REELT Auditory Comprehension Score'] = st.text_input("REELT Auditory Comprehension Score")
        optional["reelt"]['REELT Expressive Communication Score'] = st.text_input("REELT Expressive Communication Score")

    if abas_score:
        st.header("Adaptive Behavior Assessment System - Third Edition (ABAS)")
        st.markdown("*Skip this section if there is no ABAS Score*")
        optional["abas"] = {}
        optional["abas"]["Test Date"] = st.date_input("ABAS Test Date").strftime("%m/%Y")
        optional["abas"]['ABAS General Adaptive Composite'] = st.text_input("ABAS General Adaptive Composite")
        optional["abas"]['ABAS Conceptual'] = st.text_input("ABAS Conceptual")
        optional["abas"]['ABAS Social'] = st.text_input("ABAS Social")
        optional["abas"]['ABAS Practical'] = st.text_input("ABAS Practical")

    ########################################################
    st.header("Behavioral Presentation")
    data['behavior_observation'] = st.text_area(
        "Behavioral Observation: Edit the response before submitting the form", 
        # behavior_observation,
        st.session_state.behavior_observation_mod3_no_autism,
        height=400,
    )

    ########################################################
    st.header("Developmental History")
    data['development_history'] = st.text_area(
        "Developmental History: Edit the response before submitting the form", 
        # development_history,
        st.session_state.development_history_mod3_no_autism,
        height=400,
    )

    ##########################################################
    st.header("Diagnostic Formulation")

    st.markdown("Elaboration on Diagnostic Formulation. Goes after this: *Based on observation, history, and clinical measures, Norah does not meet the criteria for autism spectrum disorder.*")
    data['{{Diagnostic Formulation}}'] = st.text_area(
        "Elaboration on Diagnostic Formulation",
        f"{data['{{Patient First Name}}']} constellation of concerns and developmental history suggest that their difficulties relating to others and self-care derive from mental health difficulties...",
        height=300,
    )

    lines["{{Result of the evaluation}}"] = st.multiselect(
        "Result of the evaluation",
        dropdowns['Result of the evaluation'],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    ##########################################################
    st.header("Recommendations")

    check_accomodations = st.checkbox("Accommodations")
    check_check_in_anxiety = st.checkbox("Check-in Support for Anxiety")
    check_hearing_assistive = st.checkbox("Hearing Assistive Technology")
    check_additional_reading_suggestions = st.checkbox("Additional Suggestions")
    check_home_community = st.checkbox("Home/Community Based Supports")
    check_coordinated_care = st.checkbox("Coordinated Care")
    check_therapeutic_supports = st.checkbox("Therapeutic Supports")
    check_coordinated_healthcare = st.checkbox("Coordinated Healthcare")
    check_executive_functioning = st.checkbox("Executive Functioning Accommodations")
    check_support = st.checkbox("Support")
    check_adult_career = st.checkbox("Adult Career and Continuing Education Services-Vocational Rehabilitation (ACCES-VR)")
    check_elopement_plan = st.checkbox("Elopement Plan")
    check_occupational_therapy = st.checkbox("Occupational Therapy")

    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")

    submit = st.form_submit_button('Submit')

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def add_behavior_presentation(paragraph, transcript):
    # separate transcript
    small_para = transcript.split('\n\n')

    st.write(small_para)

    paragraph.insert_paragraph_before().add_run(small_para[0], style='CustomStyle')
    paragraph.insert_paragraph_before()

    for sub_para in small_para[1:]:
        sub_para = sub_para.split(":")
        p = paragraph.insert_paragraph_before()
        p.add_run(sub_para[0], style='CustomStyle').italic = True
        p.add_run(f":{sub_para[1]}\n", style='CustomStyle')
        
    delete_paragraph(paragraph)

def add_developmental_history(paragraph, transcript):
    # separate transcript
    small_para = transcript.split('\n\n')
    st.write(small_para)

    for sub_para in small_para:
        sub_para = sub_para.split(":")
        p = paragraph.insert_paragraph_before()
        p.add_run(sub_para[0], style='CustomStyle').italic = True
        p.add_run(f":{sub_para[1]}\n", style='CustomStyle')
        
    delete_paragraph(paragraph)

def add_school(paragraph):
    p = paragraph.insert_paragraph_before()
    tab_stops = p.paragraph_format.tab_stops
    # tab_stops.clear()  # Start fresh for this paragraph only
    tab_stops.add_tab_stop(Inches(2.5))
    # Add data
    p.add_run("District", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{School District}}']}\t", style='CustomStyle')
    p.add_run("Classification", style='CustomStyle').font.underline = True
    p.add_run(": {{Classification}}\n\n", style='CustomStyle')

    p.add_run("Setting", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{Education Setting}}']}\t", style='CustomStyle')
    p.add_run("Grade", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{Grade}}']} (", style='CustomStyle')
    p.add_run({data['School Year']}, style='CustomStyle').italic = True
    p.add_run(")", style='CustomStyle')
    delete_paragraph(paragraph)

def add_scq_form(paragraph):
    r = paragraph.insert_paragraph_before().add_run('Social Communication Questionnaire (SCQ) – Lifetime Form', style='CustomStyle')
    r.italic = True
    r.font.underline = True
    p = paragraph.insert_paragraph_before()
    p.add_run("The SCQ evaluates for symptoms of autism spectrum disorder across developmental history. Scores above 15 are suggestive of an autism diagnosis. Based on the {{Caregiver type}}’s report, ", style='CustomStyle')
    p.add_run("{{Patient First Name}}’s score was ", style='CustomStyle').bold = True
    p.add_run(f"{data['Results (SCQ) Lifetime Form']}. ", style='CustomStyle').bold = True
    r = p.add_run("This score is clearly consistent with autism at present.\n", style='CustomStyle')
    r.bold = True
    r.italic = True
    delete_paragraph(paragraph)

def add_vineland_no_teacher(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Vineland Adaptive Behavior Scales – 3rd Ed. (VABS-3) – Parent Report', style='CustomStyle')
    r.italic = True
    r.font.underline = True

    paragraph.insert_paragraph_before().add_run("The VABS-3 yields information about an individual’s adaptive functioning, which is the ability to independently perform daily activities for personal and social sufficiency. The Adaptive Behavior Composite measures overall adaptive functioning, while separate scores provide more details about communication, daily living skills, and socialization.\n\nStandard scores on the VABS-3 have a mean of 100 and a standard deviation of 15.  Scores between 85 and 115 are within the average range for this test, scores between 70 and 84 are considered moderately low, and scores below 70 are considered very low.\n", style='CustomStyle')
    
    paragraph.insert_paragraph_before().add_run('\tAdaptive Behavior Composite: {{Vineland Score Caregiver}} ({{Caregiver type}})', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run('\t\tCommunication: {{Communication Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('\t\tDaily Living Skills: {{Daily Living Skills Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('\t\tSocialization: {{Socialization Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')

def add_vineland_yes_teacher(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Vineland Adaptive Behavior Scales – 3rd Ed. (VABS-3) – Parent & Teacher Report', style='CustomStyle')
    r.italic = True
    r.font.underline = True

    paragraph.insert_paragraph_before().add_run("The VABS-3 yields information about an individual’s adaptive functioning, which is the ability to independently perform daily activities for personal and social sufficiency. The Adaptive Behavior Composite measures overall adaptive functioning, while separate scores provide more details about communication, daily living skills, and socialization.\n\nStandard scores on the VABS-3 have a mean of 100 and a standard deviation of 15.  Scores between 85 and 115 are within the average range for this test, scores between 70 and 84 are considered moderately low, and scores below 70 are considered very low.\n", style='CustomStyle')
    
    paragraph.insert_paragraph_before().add_run('\tAdaptive Behavior Composite: {{Vineland Score Caregiver}} ({{Caregiver type}}), {{Vineland Score Teacher}} (teacher)', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run('\t\tCommunication: {{Communication Score Caregiver}} ({{Caregiver type}}), {{Communication Score Teacher}} (teacher)', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('\t\tDaily Living Skills: {{Daily Living Skills Score Caregiver}} ({{Caregiver type}}), {{Daily Living Skills Score Teacher}} (teacher)', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('\t\tSocialization: {{Socialization Score Caregiver}} ({{Caregiver type}}), {{Socialization Score Teacher}} (teacher)', style='CustomStyle')

def add_srs_no_teacher(paragraph):
    r = paragraph.insert_paragraph_before().add_run('Social Responsiveness Scale – Second Edition (SRS-2) – Parent Report', style='CustomStyle')
    r.italic = True
    r.font.underline = True
    p = paragraph.insert_paragraph_before()
    p.add_run('The SRS-2 is an objective measure that identifies social impairments associated with autism spectrum disorder and quantifies ASD-related severity throughout the lifespan.\n\nStudies show that the SRS-2 discriminates both ', style='CustomStyle')
    p.add_run("within ", style='CustomStyle').italic = True
    p.add_run("the autism spectrum and between ASD and other disorder, which makes the test useful for differential diagnosis. Raters evaluate symptoms using a scale representing a range of severity. Although not used for diagnosis, subscale scores are helpful in designing and evaluating treatment.\n\nSRS-2 scores are reported here as T-scores with a mean of 50 and a standard deviation of 10 with higher scores indicating greater levels of concern for how social behavior impacts or interferes with everyday interactions. The following interpretative guidelines are offered here for the benefit of the reader: Less than 59 indicates within normal limits, between 60 and 65 as mild concern, between 65 and 75 as moderate concern, and greater than 76 as severe. {{Patient First Name}}’s {{Caregiver type}} and teacher reported the following:", style='CustomStyle')
    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before()
    r = p.add_run('\tSRS-2 Total Score', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run(': {{SRS-2 Score Caregiver}} ({{Caregiver type}})', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run('\tSocial Communication and Interaction: {{Social Communication and Interaction Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('\tRestricted Interests and Repetitive Behavior: {{Restricted Interests and Repetitive Behavior Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before()
    observe = paragraph.insert_paragraph_before()
    observe.add_run("Based on the report provided by {{Preferred Pronouns 2}} {{Caregiver type}}, ", style='CustomStyle')
    observe.add_run("{{Patient First Name}}’s social communication and related behaviors indicated {{Caregiver's level of concern}} concerns. ", style='CustomStyle').italic = True
    observe.add_run("My observation aligned with a {{Evaluator's level of concern}} concern.\n", style='CustomStyle').bold = True
    delete_paragraph(paragraph)

def add_srs_yes_teacher(paragraph, score_data):
    r = paragraph.insert_paragraph_before().add_run('Social Responsiveness Scale – Second Edition (SRS-2) – Parent Report', style='CustomStyle')
    r.italic = True
    r.font.underline = True
    p = paragraph.insert_paragraph_before()
    p.add_run('The SRS-2 is an objective measure that identifies social impairments associated with autism spectrum disorder and quantifies ASD-related severity throughout the lifespan.\n\nStudies show that the SRS-2 discriminates both ', style='CustomStyle')
    p.add_run("within ", style='CustomStyle').italic = True
    p.add_run("the autism spectrum and between ASD and other disorder, which makes the test useful for differential diagnosis. Raters evaluate symptoms using a scale representing a range of severity. Although not used for diagnosis, subscale scores are helpful in designing and evaluating treatment.\n\nSRS-2 scores are reported here as T-scores with a mean of 50 and a standard deviation of 10 with higher scores indicating greater levels of concern for how social behavior impacts or interferes with everyday interactions. The following interpretative guidelines are offered here for the benefit of the reader: Less than 59 indicates within normal limits, between 60 and 65 as mild concern, between 65 and 75 as moderate concern, and greater than 76 as severe. {{Patient First Name}}’s {{Caregiver type}} and teacher reported the following:", style='CustomStyle')
    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before()
    p.add_run('\tSRS-2 Total Score: {{SRS-2 Score Caregiver}} ({{Caregiver type}}), ', style='CustomStyle').bold = True
    p.add_run(f"{score_data['{{SRS-2 Score Teacher}}']} (teacher)", style='CustomStyle').bold = True
    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before()
    p.add_run('\tSocial Communication and Interaction: {{Social Communication and Interaction Score Caregiver}} ({{Caregiver type}}), ', style='CustomStyle')
    p.add_run(f"{score_data['{{Social Communication and Interaction Score Teacher}}']} (teacher)", style='CustomStyle')
    p = paragraph.insert_paragraph_before()
    p.add_run('\tRestricted Interests and Repetitive Behavior: {{Restricted Interests and Repetitive Behavior Score Caregiver}} ({{Caregiver type}}), ', style='CustomStyle')
    p.add_run(f'{score_data["{{Restricted Interests and Repetitive Behavior Score Teacher}}"]} (teacher)', style='CustomStyle')
    paragraph.insert_paragraph_before()
    observe = paragraph.insert_paragraph_before()
    observe.add_run("Based on the report provided by {{Preferred Pronouns 2}} {{Caregiver type}}, ", style='CustomStyle')
    observe.add_run("{{Patient First Name}}’s social communication and related behaviors indicated {{Caregiver's level of concern}} concerns. ", style='CustomStyle').italic = True
    observe.add_run("{{Patient First Name}}’s teacher reported a ", style='CustomStyle')
    observe.add_run(f"{score_data['{{Teacher level of concern}}']} concern, and ", style='CustomStyle')
    observe.add_run("my observation aligned with a {{Evaluator's level of concern}} concern.\n", style='CustomStyle').bold = True
    delete_paragraph(paragraph)

def add_wppsi(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed.', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tFull Scale IQ: {score_data["WPPSI Full Scale IQ Score"]}', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run(f'\tVerbal Comprehension: {score_data["WPPSI Verbal Comprehension Score"]}\t\t\tVisual Spatial: {score_data["WPPSI Visual Spatial Score"]}', style='CustomStyle')
    
def add_dppr(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Developmental Profile – Fourth Edition – Parent Report', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tCognitive: {score_data["DPPR Cognitive Score"]}\t\t\t\t\tSocial-Emotional: {score_data["DPPR Social-Emotional Score"]}', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run(f'\tAdaptive: {score_data["DPPR Adaptive Score"]}\t\t\t\t\tPhysical: {score_data["DPPR Physical Score"]}', style='CustomStyle')

def add_pls(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Preschool Language Scale – Fifth Edition', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tTotal Language Score: {score_data["PLS Total Language Score"]}', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run(f'\tAuditory Comprehension: {score_data["PLS Auditory Comprehension Score"]} \t\tExpressive Communication: {score_data["PLS Expressive Communication Score"]}', style='CustomStyle')

def add_pdms(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Peabody Developmental Motor Scales – Second Edition', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tGross Motor: {score_data["PDMS Gross Motor Score"]}\t\t\t\tFine Motor: {score_data["PDMS Fine Motor Score"]}', style='CustomStyle')
    
def add_peshv(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Preschool Evaluation Scale Home Version – Second Edition', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tCognitive: {score_data["PESHV Cognitive Score"]} \t\t\t\t\tSocial Emotional: {score_data["PESHV Social Emotional Score"]}', style='CustomStyle')

def add_reelt(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Receptive Expressive Emergent Language Test – Fourth Edition', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tTotal Language: {score_data["REELT Total Language Score"]}', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run(f'\tAuditory Comprehension: {score_data["REELT Auditory Comprehension Score"]}', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run(f'\tExpressive Communication: {score_data["REELT Expressive Communication Score"]}', style='CustomStyle')
    
def add_abas(paragraph, score_data):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["Test Date"]}) – Adaptive Behavior Assessment System – Third Edition', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tGeneral Adaptive Composite: {score_data["ABAS General Adaptive Composite"]}', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run(f'\tConceptual: {score_data["ABAS Conceptual"]}', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run(f'\tSocial: {score_data["ABAS Social"]}\t\t\tPractical: {score_data["ABAS Practical"]}', style='CustomStyle')
    
def add_bullet(paragraph, list_data):
    paragraph.insert_paragraph_before()
    for item in list_data:
        paragraph.insert_paragraph_before().add_run(item, style='ListStyle')
    delete_paragraph(paragraph)

   
###############################################################
# Recommendations

def add_hyperlink(paragraph, url, size=24):
    """
    A function that places a hyperlink within a paragraph object with custom font and size.

    :param paragraph: The paragraph we are adding the hyperlink to.
    :param url: A string containing the required url
    :param text: The text displayed for the url
    :param color: Hex color string (e.g., '0000FF')
    :param underline: Bool indicating whether the link is underlined
    :return: The hyperlink object
    """

    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    # Set font to Georgia
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Georgia')
    rFonts.set(qn('w:hAnsi'), 'Georgia')
    rPr.append(rFonts)

    # Set font size to 11.5pt (23 half-points)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), f'{size}')
    rPr.append(sz)

    c = OxmlElement('w:color')
    c.set(qn('w:val'), '1155cc')
    rPr.append(c)

    # Set underline
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)

    # Add text
    text_elem = OxmlElement('w:t')
    text_elem.text = url
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

    return hyperlink

def add_accomodations(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Accommodations. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I would recommend that a designated location be determined for assessments, especially those with a timed component. There is also value providing {{Patient First Name}} with additional time and supports to remember to ask for accommodations.\n', style='CustomStyle')

def add_check_in_anxiety(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Check-in Support for Anxiety. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('{{Patient First Name}}’s parents reported that {{Preferred Pronouns 1}} has an anxiety diagnosis, and we would recommend adding counseling supports to help determine if anxiety is contributing to difficulties with auditory filtering or attention on assessments.\n', style='CustomStyle')

def add_hearing_assistive(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Hearing Assistive Technology. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run("Given {{Patient First Name}}'s presentation, results of assessments, and behavioral observations across settings, we recommend trialing the use of a hearing assistive technology to address concerns related to auditory filtering and attention.\n", style='CustomStyle')

def add_additional_reading_suggestions(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Additional Suggestions. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run("Given {{Patient First Name}}'s reading difficulties, I recommend the following:\n", style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Enlarged font with strategic spacing between lines of text for assignments\n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Hearing assistive technology to rule-out potential auditory or attention concerns\n', style='CustomStyle')

    p = paragraph.insert_paragraph_before()
    p.add_run("In so far as {{Patient First Name}}'s difficulties could be attributed to, in part, to auditory processing or filtering concerns, I think there’s value in trialing the use of a hearing assistive tech.", style='CustomStyle').italic = True
    
def add_home_community(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Home/Community Based Supports. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('Based on the report by {{Preferred Pronouns 2}} mother and our observations today, we recommend considering therapeutic supports through organizations that provide specialized care. To that end, we recommend the following organizations: \n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('URMC Pediatric Behavioral Health and Wellness:\n', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https://www.urmc.rochester.edu/childrens-hospital/behavioral-health-wellness/outpatient')

def add_coordinated_care(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Coordinated Care. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('We encourage collaboration between home, school, and medical providers to ensure a unified approach to supporting {{Patient First Name}}’s needs across settings.\n', style='CustomStyle')

def add_therapeutic_supports(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Therapeutic Supports. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('Based on the report by {{Preferred Pronouns 2}} mother and our observations today, we recommend considering therapeutic supports through organizations that provide specialized care. To that end, we recommend the following organizations: \n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('URMC Pediatric Behavioral Health and Wellness:\n', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https://www.urmc.rochester.edu/childrens-hospital/behavioral-health-wellness/outpatient')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Genesee Valley Psychology: ', style='CustomStyle')
    add_hyperlink(p, 'https://gviproc.org/about')
    paragraph.insert_paragraph_before()

def add_coordinated_healthcare(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Therapeutic Supports. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run("Based on {{Patient First Name}}'s history of mental health and inter-related social concerns, we think there is value in pursuing coordinated healthcare such as:\n", style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https://www.rochesterregional.org/services/adult-mental-health/pros')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    add_hyperlink(p, 'https://www.hhuny.org/')
    paragraph.insert_paragraph_before()

def add_executive_functioning(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Executive Functioning Accommodations. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run("{{Patient First Name}} will need frequent check-ins to review {{Preferred Pronouns 2}} daily and weekly activities to make sure {{Preferred Pronouns 1}} remains on target with longer-term assignments. Organization of materials is a critical component of this process, and, as such, we recommend daily adult support to initiate and complete work for {{Patient First Name}}.\n\nOne approach that might be worth pursuing at school would be:\n", style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run("Unstuck & On Target - ", style='CustomStyle')
    add_hyperlink(p, 'https://www.rochesterregional.org/services/adult-mental-health/pros')

def add_support(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Support. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('We think that {{Patient First Name}} would benefit from the school’s behavioral specialist to accommodate {{Preferred Pronouns 2}} day-to-day management needs at school. We encourage opportunities to engage with {{Preferred Pronouns 2}} peers as well as executive functioning suggestions outlined above.\n', style='CustomStyle')

def add_adult_career(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Adult Career and Continuing Education Services-Vocational Rehabilitation (ACCES-VR)', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('{{Patient First Name}} may benefit from vocational support to help maintain employment and supported {{Preferred Pronouns 2}} independent living. Information can be found on the NYSD Department’s website: ', style='CustomStyle')
    add_hyperlink(p, 'https://www.acces.nysed.gov/vr')
    paragraph.insert_paragraph_before()

def add_elopement_plan(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Elopement Plan. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('Given {{Patient First Name}}’s predisposition to wander and bolt if not closely monitored, I think that it is medically necessary for {{Preferred Pronouns 2}} team to have in place a series of preventative and responsive procedures related to {{Preferred Pronouns 2}} elopement. This could be done in consultation with the school team (teacher, social worker) and a behavior specialist.\nResources to consider include:\n', style='CustomStyle')
    
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Big Red Safety Toolkit - ', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https://nationalautismassociation.org/docs/BigRedSafetyToolkit.pdf')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Angel Sense - ', style='CustomStyle')
    add_hyperlink(p, 'https://www.angelsense.com/gps-tracker-lifesaving-features/')

    paragraph.insert_paragraph_before(style='Normal')

def add_occupational_therapy(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Occupational Therapy. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('Based on {{Preferred Pronouns 2}} parents’ report, there is a clear pattern of concerns related to auditory aversions that set to the occasion for potential elopement. As such, I would recommend reaching out to URMC Pediatric Occupational Therapy:\n', style='CustomStyle')
    
    p = paragraph.insert_paragraph_before(style='Bullet New')
    add_hyperlink(p, 'https://www.urmc.rochester.edu/locations/pediatric-neurology')

if submit:
    # Update session state 
    st.session_state.behavior_observation_mod3_no_autism = data['behavior_observation']
    st.session_state.development_history_mod3_no_autism = data['development_history'] 

    # handle word to replace 
    # pronouns
    with open("misc_data/pronouns.yaml", "r") as file:
        pronoun = yaml.safe_load(file)

    replace_word = {
        "{{Preferred Pronouns 1}}": pronoun[preferred]['pronoun1'],
        "{{Preferred Pronouns 1 CAP}}": pronoun[preferred]['pronoun1cap'],
        "{{Preferred Pronouns 2}}": pronoun[preferred]['pronoun2'],
        "{{Preferred Pronouns 2 CAP}}": pronoun[preferred]['pronoun2cap'],
        "{{Gender}}": pronoun[preferred]['gender'],
    }

    replace_word.update(data)

    # Add optional data 
    if not wppsi_score and 'wppsi' in optional:
        del optional['wppsi']
    if not dppr_score and 'dppr' in optional:
        del optional['dppr']
    if not pls_score and 'pls' in optional:
        del optional['pls']
    if not pdms_score and 'pdms' in optional:
        del optional['pdms']
    if not peshv_score and 'peshv' in optional:
        del optional['peshv']
    if not reelt_score and 'reelt' in optional:
        del optional['reelt']
    if not abas_score and 'abas' in optional:
        del optional['abas']

    # Display data 
    yaml_string = yaml.dump(replace_word, sort_keys=False)
    yaml_string = yaml_string + '\n' + yaml.dump(optional, sort_keys=False)
    yaml_string = yaml_string + '\n' + yaml.dump(bullet, sort_keys=False)
    yaml_data = st.code(yaml_string, language=None)
    

    #### Edit document 
    doc = Document('templates/template_mod_3_no_autism.docx')
    if doc:
        # Get file name
        today_date = format_date_with_ordinal(datetime.date.today())
        filename = f"{data['{{Patient First Name}}']} {data['{{Patient Last Name}}']} {today_date}.docx"
        
        ### create document 
        norm_style = doc.styles['Normal']
        norm_style.paragraph_format.line_spacing = 1

        custom_style = doc.styles.add_style('CustomStyle', WD_STYLE_TYPE.CHARACTER)
        custom_style.font.size = Pt(12)
        custom_style.font.name = 'Georgia'

        custom_style_2 = doc.styles.add_style('CustomStyle2', WD_STYLE_TYPE.CHARACTER)
        custom_style_2.font.size = Pt(11.5)
        custom_style_2.font.name = 'Georgia'

        list_style = doc.styles['Bullet New']
        list_style.paragraph_format.line_spacing = 1

        # Add scores 
        for i, paragraph in enumerate(doc.paragraphs):
            if len(optional) > 0:
                if "Scores are reported here as standard scores" in paragraph.text:
                    if 'wppsi' in optional:
                        add_wppsi(paragraph, optional['wppsi'])
                    if 'dppr' in optional:
                        add_dppr(paragraph, optional["dppr"])
                    if 'pls' in optional:
                        add_pls(paragraph, optional["pls"])
                    if 'pdms' in optional:
                        add_pdms(paragraph, optional["pdms"])
                    if 'peshv' in optional:
                        add_peshv(paragraph, optional['peshv'])
                    if 'reelt' in optional:
                        add_reelt(paragraph, optional['reelt'])
                    if 'abas' in optional:
                        add_abas(paragraph, optional['abas'])
            
            if "[[Behavioral Presentation]]" in paragraph.text:
                add_behavior_presentation(paragraph, st.session_state.behavior_observation_mod3_no_autism)
            
            if "[[Developmental History]]" in paragraph.text:
                add_developmental_history(paragraph, st.session_state.development_history_mod3_no_autism)

            if "[[Recommendations]]" in paragraph.text:
                if check_accomodations:
                    add_accomodations(paragraph)
                if check_check_in_anxiety:
                    add_check_in_anxiety(paragraph)
                if check_hearing_assistive:
                    add_hearing_assistive(paragraph)
                if check_additional_reading_suggestions:
                    add_additional_reading_suggestions(paragraph)
                if check_home_community:
                    add_home_community(paragraph)
                if check_coordinated_care:
                    add_coordinated_care(paragraph)
                if check_therapeutic_supports:
                    add_therapeutic_supports(paragraph)
                if check_coordinated_healthcare:
                    add_coordinated_healthcare(paragraph)
                if check_executive_functioning:
                    add_executive_functioning(paragraph)
                if check_support:
                    add_support(paragraph)
                if check_adult_career:
                    add_adult_career(paragraph)
                if check_elopement_plan:
                    add_elopement_plan(paragraph)
                if check_occupational_therapy:
                    add_occupational_therapy(paragraph)

            if "[[SCQ Report Information]]" in paragraph.text:
                # Add SCQ
                if scq_result:
                    add_scq_form(paragraph)
                else:
                    delete_paragraph(paragraph)
                
            if "[[SRS Report Information]]" in paragraph.text:
                # Add SRS
                if caregiver_srs_eval:
                    if teacher_srs_eval:
                        add_srs_yes_teacher(paragraph, teacher_score)
                    else:
                        add_srs_no_teacher(paragraph)
                # if nothing 
                delete_paragraph(paragraph)
            
            if "Social Responsiveness Scale" in paragraph.text:
                if teacher_srs_eval:
                    paragraph.add_run(" & teacher", style='CustomStyle')

            if "[[Assessment Procedures]]" in paragraph.text:
                if scq_result:
                    paragraph.insert_paragraph_before().add_run("Social Communication Questionnaire (SCQ): Completed by {{Preferred Pronouns 2}} {{Caregiver type}}", style='CustomStyle')

                if caregiver_srs_eval:
                    if teacher_srs_eval:
                        paragraph.insert_paragraph_before().add_run("Social Responsiveness Scale – 2nd Edition: Completed by {{Preferred Pronouns 2}} {{Caregiver type}} & teacher", style='CustomStyle')
                    else:
                        paragraph.insert_paragraph_before().add_run("Social Responsiveness Scale – 2nd Edition: Completed by {{Preferred Pronouns 2}} {{Caregiver type}}", style='CustomStyle')

                if caregiver_vineland_eval:
                    if teacher_vineland_eval:
                        paragraph.insert_paragraph_before().add_run("Vineland Adaptive Behavior Scale 3rd Edition: Completed by {{Preferred Pronouns 2}} {{Caregiver type}} & teacher", style='CustomStyle')
                    else:
                        paragraph.insert_paragraph_before().add_run("Vineland Adaptive Behavior Scale 3rd Edition: Completed by {{Preferred Pronouns 2}} {{Caregiver type}}", style='CustomStyle')
                        
                delete_paragraph(paragraph)

            if "[[Vineland Scale Information]]" in paragraph.text:
                if caregiver_vineland_eval:
                    if teacher_vineland_eval:
                        add_vineland_yes_teacher(paragraph)
                    else:
                        add_vineland_no_teacher(paragraph)
                # if nothing 
                delete_paragraph(paragraph)
            
            if "Developmental History & Review of Records" in paragraph.text:
                if teacher_srs_eval:
                    paragraph.add_run(f"\nSchool Report on SRS-2 provided by {teacher_score['{{Teacher name, title}}']}", style='CustomStyle')
                if teacher_vineland_eval:
                    paragraph.add_run(f"\nReport on Vineland Adaptive Behavior Scale provided by {teacher_score['{{Teacher name, title}}']}", style='CustomStyle')

            if "[[District Grade School Setting]]" in paragraph.text:
                add_school(paragraph)
        
        # Edit document
        for word in replace_word:
            docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

        # Replace for lists separated by comma:
        for word in comma:
            new_word = ", ".join(comma[word])
            docxedit.replace_string(doc, old_string=word, new_string=new_word)

        # Replace for lists separated by new line:
        for word in lines:
            new_word = "\n".join(lines[word])
            docxedit.replace_string(doc, old_string=word, new_string=new_word)

        # Save content to file
        doc.save(filename)

        # Replace for lists separated by bullet points
        tpl=DocxTemplate(filename)
        print("Load template!")

        tpl.render(bullet)
        print("Bullet rendered!")

        tpl.save(filename)
        print("File saved at", filename)

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