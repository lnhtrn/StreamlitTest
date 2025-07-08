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
    page_title="Module 1&2",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="expanded",
)

##########################################################
# Set up OpenAI 
if 'behavior_observation' not in st.session_state:
    st.session_state.behavior_observation = ""
if 'development_history' not in st.session_state:
    st.session_state.development_history = ""

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
connections['All'] = st.connection(f"mod12_all", type=GSheetsConnection)
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
    st.link_button("Edit Dropdown Options", st.secrets['mod12_spreadsheet'])
    st.button('Reload Dropdown Data', on_click=clear_my_cache)

    # Display data 
    # yaml_dropdown = yaml.dump(dropdowns, sort_keys=False)
    # st.code(yaml_dropdown, language=None)
    
    ####################################################
    st.markdown("**Check to include score in the form:** Scores to report:")
    scq_result = st.checkbox("Social Communication Questionnaire (SCQ) - Lifetime Form")
    teacher_eval = st.checkbox("Teacher's SSR Scores")
    wppsi_score = st.checkbox("Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed. (WPPSI) Score")
    dppr_score = st.checkbox("Developmental Profile – Fourth Edition - Parent Report (DPPR)")
    pls_score = st.checkbox("Preschool Language Scale - Fifth Edition (PLS)")
    pdms_score = st.checkbox("Peabody Developmental Motor Scales - Second Edition")
    peshv_score = st.checkbox("Preschool Evaluation Scale Home Version - Second Edition")
    reelt_score = st.checkbox("Receptive Expressive Emergent Language Test - Fourth Edition")
    abas_score = st.checkbox("Adaptive Behavior Assessment System - Third Edition")


col1,col2 = st.columns(2)
col1.title('Module 1&2 Report Builder')

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

#############################################################
# Start of form 
st.header("Appointment Summary")
data['{{Patient First Name}}'] = st.text_input('Patient First Name')
data['{{Patient Last Name}}'] = st.text_input('Patient Last Name')
preferred = st.selectbox(
    "Patient's Preferred Pronoun",
    ("They/them", "He/him", "She/her"),
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
                "id": st.secrets["behavior_prompt_mod12_id"],
                # "version": "3",
                "variables": {
                    "first_name": data['{{Patient First Name}}'],
                    "pronouns": preferred,
                    "diagnosis": "having autism",
                    "transcription": transcript_behavior
                }
            }
        )
        st.session_state.behavior_observation = response.output_text

        response = client.responses.create(
            prompt={
                "id": st.secrets["development_prompt_mod12_id"],
                # "version": "5",
                "variables": {
                    "first_name": data['{{Patient First Name}}'],
                    "pronouns": preferred,
                    "diagnosis": "having autism",
                    "transcription": transcript_development
                }
            }
        )
        st.session_state.development_history = response.output_text
        
################################################################
# Start form
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
        # [
        #     "Speech delays impacting social opportunities.",
        #     "Clarifying diagnostic presentation.",
        #     "Determining service eligibility.",
        #     "Language delays and difficulties.",
        #     "Elopement and related safety concerns.",
        #     "Determining appropriate supports."
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )
    
    data['{{Residence City/State}}'] = st.text_input("Residence City/State")
    # st.selectbox(
    #     "Residence City/State", states, index=None,
    # )

    data['{{Narrative}}'] = st.text_area('Narrative to finish \"Patient lives with...\"')

    ##########################################################
    st.header("BRH Evaluation Details")

    data['{{Evaluation Date}}'] = format_date_with_ordinal(st.date_input("Evaluation Date"))

    data['{{Module used}}'] = st.radio("Module used", ["Module 1", "Module 2"])
    if data['{{Module used}}'] == "Module 1":
        data['{{Module Description}}'] = "Module 1 is designed for children with single words"
    else:
        data['{{Module Description}}'] = "Module 2 is designed for children with phrase speech"

    data['{{Location of the evaluation}}'] = st.radio(
        "Location of the evaluation",
        ['home', 'school', 'the office'],
        index=None,
    )

    data['{{Results Shared Date}}'] = format_date_with_ordinal(st.date_input("Results Shared Date"))
    
    data['{{Date Report Sent to Patient}}'] = format_date_with_ordinal(st.date_input("Date Report Sent to Patient"))

    lines["{{Result of the evaluation}}"] = st.multiselect(
        "Result of the evaluation",
        dropdowns["Result of the evaluation"],
        # [
        #     "F84.0 - Autism Spectrum Disorder (per the above referenced evaluation)",
        #     "F88.0 - Global Developmental Delay (per behavioral presentation)",
        #     "F80.2 - Mixed Receptive-Expressive Language Disorder",
        #     "F90.2 - Attention Deficit Hyperactivity Disorder - Combined-Type",
        #     "F50.82 Avoidant/Restrictive Food Intake Disorder",
        #     "None"
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    if scq_result:
        data["{{Results (SCQ) - Lifetime Form}}"] = st.text_input(
            "Results (SCQ) - Lifetime Form"
        )

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
    if teacher_eval:
        st.header("Teacher SSR Score")
        st.markdown("*Skip this section if teacher did not give SSR Score*")

        teacher_score['{{Teacher name, title}}'] = st.text_input("Teacher name, title")

        teacher_score['{{SRS-2 Score Teacher}}'] = st.text_input("Teacher's SRS-2 Score")

        teacher_score['{{Social Communication and Interaction Score Teacher}}'] = st.text_input("Social Communication and Interaction Score Teacher")

        teacher_score['{{Restricted Interests and Repetitive Behavior Score Teacher}}'] = st.text_input("Restricted Interests and Repetitive Behavior Score Teacher")

        teacher_score["{{Teacher level of concern}}"] = st.radio(
            "Teacher's level of concern",
            ['no', 'mild', 'moderate', 'severe']
        )

    ######################################################
    st.header("Medical/Developmental History")
    
    lines['{{Diagnosis History}}'] = st.multiselect(
        "Diagnosis History",
        dropdowns['Diagnosis History'],
        # ['History of language and social communication delays.'],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    lines['{{Medications}}'] = st.multiselect(
        "Medications",
        ['None noted or reported.'],
        placeholder="Can input multiple options",
        accept_new_options=True
    )

    ###############################################
    st.header("Educational Background")

    data['{{School District}}'] = st.selectbox(
        "School District",
        ['Rochester City'],
        index=None,
        placeholder="Select a school district or enter a new one",
        accept_new_options=True,
    )

    data['{{School Name}}'] = st.text_input("School Name")

    data['{{Grade}}'] = st.text_input(
        "Grade",
        # dropdowns['Grade'],
        # index=None,
        # placeholder="Select a grade or enter a new one",
        # accept_new_options=True,
    )

    data['School Year'] = st.text_input(
        "School Year",
        # dropdowns["School Year"],
        # index=None,
        # placeholder="Select a school year or enter a new one",
        # accept_new_options=True,
    )

    data['{{Education Setting}}'] = st.selectbox(
        "Education Setting",
        ["General Education", "Integrated Co-Taught", "12:1:1", "8:1:1", "6:1:1"],
        index=None,
        placeholder="Select an education setting or enter a new one",
        accept_new_options=True,
    )

    comma['{{Services}}'] = st.multiselect(
        "Services",
        dropdowns['Services'],
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
        st.session_state.behavior_observation,
        height=400,
    )

    ########################################################
    st.header("Developmental History")
    data['development_history'] = st.text_area(
        "Developmental History: Edit the response before submitting the form", 
        # development_history,
        st.session_state.development_history,
        height=400,
    )

    ########################################################
    st.header("DSM Criteria")
    
    bullet['SocialReciprocity'] = st.multiselect(    
        "Deficits in social emotional reciprocity",
        dropdowns['SocialReciprocity'],
        # [
        #     "None",
        #     "Awkward social initiation and response",
        #     "Difficulties with chit-chat",
        #     "Difficulty interpreting figurative language",
        #     "Limited social approach or greetings",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    bullet['NonverbalComm'] = st.multiselect(
        "Deficits in nonverbal communicative behaviors used for social interaction",
        dropdowns['NonverbalComm'],
        # [
        #     "None",
        #     "Limited well-directed eye contact",
        #     "Difficulty reading facial expressions",
        #     "Absence of joint attention",
        #     "Lack of well-integrated gestures",
        #     "Limited range of facial expression",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    bullet['Relationships'] = st.multiselect(
        "Deficits in developing, maintaining, and understanding relationships",
        dropdowns['Relationships'],
        # [
        #     "None",
        #     "Limited engagement with same age peers",
        #     "Difficulties adjusting behavior to social context",
        #     "Difficulties forming friendships",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    bullet['RepetitiveBehaviors'] = st.multiselect(
        "Stereotyped or repetitive motor movements, use of objects, or speech",
        dropdowns['RepetitiveBehaviors'],
        # [
        #     "None",
        #     "Repetitive whole-body movements",
        #     "Repetitive hand movements",
        #     "Echolalia of sounds",
        #     "Echolalia of words",
        #     "Stereotyped speech",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    bullet['SamenessRoutines'] = st.multiselect(
        "Insistence on sameness, inflexible adherence to routines or ritualized behavior",
        dropdowns['SamenessRoutines'],
        # [
        #     "None",
        #     "Difficulties with changes in routine across developmental course",
        #     "Notable difficulties with transitions",
        #     "Insistence on following very specific routines",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    bullet['RestrictedInterests'] = st.multiselect(
        "Highly restricted, fixated interests that are abnormal in intensity or focus",
        dropdowns['RestrictedInterests'],
        # [
        #     "None",
        #     "Persistent pattern of perseverative interests",
        #     "Notable interest in topics others may find odd",
        #     "Very restricted pattern of eating and sleep time behavior",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    bullet['SensoryReactivity'] = st.multiselect(
        "Hyper- or hypo-reactivity to sensory aspects of the environment",
        dropdowns['SensoryReactivity'],
        # [
        #     "None",
        #     "Auditory sensitivities",
        #     "Tactile defensiveness",
        #     "Proprioceptive-seeking behavior",
        # ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    comma['{{Symptoms present in the early developmental period}}'] = st.multiselect(
        "Symptoms present in the early developmental period",
        [
            "Confirmed by record review",
            "None",
        ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    comma['{{Symptoms cause clinically significant impairment}}'] = st.multiselect(
        "Symptoms cause clinically significant impairment",
        [
            "Confirmed by record review",
            "None",
        ],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    ########################################################################
    st.header("Recommendations")

    check_developmental_pediatrics = st.checkbox("Developmental Pediatrics Appointment")
    check_feeding_treatment = st.checkbox("Feeding Treatment & Support")
    check_levine_clinic = st.checkbox("Levine Autism Clinic")
    check_parent_parent = st.checkbox("Parent to Parent")
    check_100_days = st.checkbox("Autism Speaks 100 Days 100 Kit")
    check_caregiver_support = st.checkbox("Caregiver Support")
    check_edu_placement = st.checkbox("Educational Placement")
    check_effective_treatments = st.checkbox("Components of Effective Treatment")
    check_elopement_plan = st.checkbox("Elopement Plan")
    check_develop_disability_office = st.checkbox("Developmental Disabilities Regional Office (DDRO)")
    check_evidence_therapy = st.checkbox("Evidence-Based Therapies")

    # data['{{}}'] = st.text_input("")
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
    tab_stops.add_tab_stop(Inches(3))
    # Add data
    p.add_run("District", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{School District}}']}\t", style='CustomStyle')
    p.add_run("Grade", style='CustomStyle').font.underline = True
    ### italics for school year
    p.add_run(f": {data['{{Grade}}']} (", style='CustomStyle')
    p.add_run({data['School Year']}, style='CustomStyle').italic = True
    p.add_run(")\n\n", style='CustomStyle')
    p.add_run("School", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{School Name}}']}\t", style='CustomStyle')
    p.add_run("Setting", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{Education Setting}}']}", style='CustomStyle')
    delete_paragraph(paragraph)

def add_scq_form(paragraph):
    r = paragraph.insert_paragraph_before().add_run('Social Communication Questionnaire (SCQ) – Lifetime Form', style='CustomStyle')
    r.italic = True
    r.font.underline = True
    p = paragraph.insert_paragraph_before()
    p.add_run("The SCQ evaluates for symptoms of autism spectrum disorder across developmental history. Scores above 15 are suggestive of an autism diagnosis. Based on {{Preferred Pronouns 2}} {{Caregiver type}}’s report, {{Patient First Name}}’s score was {{Results (SCQ) - Lifetime Form}}. ", style='CustomStyle')
    p.add_run("This score is clearly consistent with autism at present.\n", style='CustomStyle').italic = True

def add_srs_no_teacher(paragraph):
    r = paragraph.insert_paragraph_before().add_run('Social Responsiveness Scale – Second Edition (SRS-2) – Parent Report', style='CustomStyle')
    r.italic = True
    r.font.underline = True
    paragraph.insert_paragraph_before().add_run('The SRS-2 is an objective measure that identifies social impairments associated with autism spectrum disorder and quantifies ASD-related severity throughout the lifespan. \nThe following interpretative guidelines are offered here for the benefit of the reader: Less than 59 indicates within normal limits, between 60 and 65 as mild concern, between 65 and 75 as moderate concern, and greater than 76 as severe concern. ', style='CustomStyle')
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run('\tSRS-2 Total Score: {{SRS-2 Score Caregiver}} ({{Caregiver type}})', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run('\tSocial Communication and Interaction: {{Social Communication and Interaction Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('\tRestricted Interests and Repetitive Behavior: {{Restricted Interests and Repetitive Behavior Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before()
    observe = paragraph.insert_paragraph_before()
    observe.add_run("Based on the report provided by {{Preferred Pronouns 2}} {{Caregiver type}}, ", style='CustomStyle')
    observe.add_run("{{Patient First Name}}’s social communication and related behaviors indicated {{Caregiver's level of concern}} concerns. ", style='CustomStyle').italic = True
    observe.add_run("My observation aligned with a {{Evaluator's level of concern}} concern.", style='CustomStyle').bold = True
    delete_paragraph(paragraph)

def add_srs_yes_teacher(paragraph, score_data):
    r = paragraph.insert_paragraph_before().add_run('Social Responsiveness Scale – Second Edition (SRS-2) – Parent & Teacher Report', style='CustomStyle')
    r.italic = True
    r.font.underline = True
    paragraph.insert_paragraph_before().add_run('The SRS-2 is an objective measure that identifies social impairments associated with autism spectrum disorder and quantifies ASD-related severity throughout the lifespan. \nThe following interpretative guidelines are offered here for the benefit of the reader: Less than 59 indicates within normal limits, between 60 and 65 as mild concern, between 65 and 75 as moderate concern, and greater than 76 as severe concern. ', style='CustomStyle')
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
    observe.add_run("my observation aligned with a {{Evaluator's level of concern}} concern.", style='CustomStyle').bold = True
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

def add_developmental_pediatrics(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Developmental Pediatrics Appointment. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I believe that {{Patient First Name}} would benefit from being seen by a developmental medical provider as part of comprehensive care related to the diagnosis described here. An appointment can be made by calling one of the following local specialty clinics or at URMC and Rochester Regional Health Center:\n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('University of Rochester Medical Center, Levine Autism Clinic at 585-275-2986,', style='CustomStyle2')
    p = paragraph.insert_paragraph_before(style='Normal')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https:/www.urmc.rochester.edu/childrens-hospital/developmental-disabilities/services/levine.aspx', size=23)

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Rochester Regional Health Center, Developmental Behavioral Pediatrics Program at 585-922-4698, ', style='CustomStyle2')
    add_hyperlink(p, 'https://www.rochesterregional.org/services/pediatrics/developmental-behavioral-pediatrics-program', size=23)
    paragraph.insert_paragraph_before()

def add_feeding_treatment(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Feeding Treatment & Support. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('{{Patient First Name}} presents with a range of concerns related to mealtime behavior and food variety, so I recommend that {{Preferred Pronouns 2}} parents seek out support from one of the following local agencies. I am happy to discuss this in detail.\n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('University of Rochester Medical Center - ', style='CustomStyle2')
    p = paragraph.insert_paragraph_before(style='Normal')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https://www.urmc.rochester.edu/childrens-hospital/developmental-disabilities/services/feeding-disorders.aspx')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Step-by-Step - ', style='CustomStyle')
    add_hyperlink(p, 'https://www.sbstherapycenter.com/feeding-therapy')
    
    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Mealtime Rediscovered - ', style='CustomStyle')
    add_hyperlink(p, 'https://mealtimerediscovered.com/')
    paragraph.insert_paragraph_before()

def add_levine_clinic(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Levine Autism Clinic. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I recommend that {{Patient First Name}}’s {{Caregiver type}} refer to the Levine Autism Clinic Facebook page for information about services, supports, events, and information that may be of help: ', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')
    add_hyperlink(p, 'https://www.facebook.com/DBPeds.GCH/')
    paragraph.insert_paragraph_before()

def add_parent_parent(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Parent to Parent. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('(', style='CustomStyle')
    add_hyperlink(p, 'http://parenttoparentnys.org/offices/Finger-Lakes/')
    p.add_run(') This group could help to connect {{Patient First Name}}’s family with another family in their area who knows more about local resources and supports related to {{Patient First Name}}’s age-level and interests.', style='CustomStyle')
    paragraph.insert_paragraph_before()

def add_100_days(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Autism Speaks 100 Days 100 Kit. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I would recommend that {{Patient First Name}}’s {{Caregiver type}} refer to this kit to help structure their next steps in determining {{Patient First Name}}’s care. The kit contains information and advice collected from trusted and respected experts. ', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')
    add_hyperlink(p, 'http://www.autismspeaks.org/community/family_services/100_day_kit.php')
    paragraph.insert_paragraph_before()

def add_caregiver_support(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Caregiver Support.  ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I encourage {{Patient First Name}}’s {{Caregiver type}} to review these resources:\n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('AutismUp - ', style='CustomStyle')
    add_hyperlink(p, 'https://autismup.org/support/family-navigator')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Autism Council of Rochester - ', style='CustomStyle')
    add_hyperlink(p, 'https://www.theautismcouncil.org/')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Camp Puzzle Peace - ', style='CustomStyle')
    add_hyperlink(p, 'www.familyautismcenter.com/')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Rochester Regional Center for Autism Spectrum Disorders - \n', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')
    p.paragraph_format.left_indent = Inches(0.5)
    add_hyperlink(p, 'https://www.urmc.rochester.edu/strong-center-developmental-disabilities/programs/rochester-regional-ctr-autism-spectrum-disorder.aspx')

    paragraph.insert_paragraph_before()

def add_edu_placement(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Educational Placement. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('The matter of which setting {{Patient First Name}} is educated in feels of paramount concern given {{Preferred Pronouns 2}} current skills and areas of need. I encourage {{Preferred Pronouns 2}} {{Caregiver type}} and school team to engage in ongoing conversations about placement options available for next year. I recommend that discussions about educational placement and programming be held within the CPSE meeting process.', style='CustomStyle')
    p = paragraph.insert_paragraph_before(style='Normal')

def add_effective_treatments(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Components of Effective Treatment. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('A body of research has accumulated about effective treatment for children with autism. A list of components of this presented below. How these are implemented is best determined by those who work with {{Patient First Name}}. \n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Comprehensive curriculum focusing on teaching a wide range of skills, including attention to the environment, imitation, comprehension and production of language, functional communication, toy play, and peer interaction.', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Supportive teaching environments structured to maximize attention to tasks.', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Emphasis on providing children with predictability and routine.', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Functional behavior analytic approach to assessing and treating behaviors.', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Systematic intervention for facilitating transitions from home to school setting.', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Consultation with a professional with expertise in autism-related interventions.', style='CustomStyle')

    paragraph.insert_paragraph_before(style=norm_style)

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

def add_develop_disability_office(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Developmental Disabilities Regional Office (DDRO). ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I discussed DDRO case management and Medicaid Waiver services with {{Patient First Name}}’s {{Caregiver type}}. To qualify for services, a person must have a diagnosis of a developmental disability along with documentation of cognitive and/or adaptive deficits. Based on {{Preferred Pronouns 2}} presentation and chart review, I believe that {{Patient First Name}} ought to quality for OPWDD waiver services due to {{Preferred Pronouns 2}} adaptive and cognitive delays. More information on Front Door Sessions can be found online at: ', style='CustomStyle')
    add_hyperlink(p, 'https://opwdd.ny.gov/get-started/information-sessions')
    paragraph.insert_paragraph_before()
    
    p = paragraph.insert_paragraph_before()
    p.add_run('Information can be obtained through the Office of Persons with Developmental Disabilities (OPWDD), ', style='CustomStyle')
    p.add_run('Front Door Office Finger Lakes', style='CustomStyle').bold = True
    p.add_run(' at 855-679-3335', style='CustomStyle')
    paragraph.insert_paragraph_before()

def add_evidence_therapy(paragraph):
    p = paragraph.insert_paragraph_before()
    r = p.add_run('Evidence-Based Therapies. ', style='CustomStyle')
    r.bold = True
    r.italic = True
    p.add_run('I would encourage {{Patient First Name}}’s family to consider seeking services that are informed by the principles of applied behavior analysis (ABA). In particular, I would recommend that {{Patient First Name}} receive intensive intervention under the supervision of a licensed professional or board-certified behavioral analyst.\n\nResources to consider include:\n', style='CustomStyle')

    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Autism Learning Partners - ', style='CustomStyle')
    add_hyperlink(p, 'https://www.autismlearningpartners.com/')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Living Soul - ', style='CustomStyle')
    add_hyperlink(p, 'https://livingsoulllc.com/')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('Proud Moments - ', style='CustomStyle')
    add_hyperlink(p, 'https://discover.proudmomentsaba.com/rochester.html')

    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before(style='Bullet New')
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run('TruNorth Autism Services - \n', style='CustomStyle')
    add_hyperlink(p, 'https://www.trunorthautism.com/')

    paragraph.insert_paragraph_before()

def add_bullet(paragraph, list_data):
    paragraph.insert_paragraph_before()
    for item in list_data:
        paragraph.insert_paragraph_before().add_run(item, style='ListStyle')
    delete_paragraph(paragraph)


if submit:
    # Update session state 
    st.session_state.behavior_observation = data['behavior_observation']
    st.session_state.development_history = data['development_history'] 

    # handle word to replace 
    # pronouns
    with open("misc_data/pronouns.yaml", "r") as file:
        pronoun = yaml.safe_load(file)

    replace_word = {
        "{{Preferred Pronouns 1}}": pronoun[preferred]['pronoun1'],
        "{{Preferred Pronouns 1 CAP}}": pronoun[preferred]['pronoun1cap'],
        "{{Preferred Pronouns 2}}": pronoun[preferred]['pronoun2'],
        "{{Preferred Pronouns 2 CAP}}": pronoun[preferred]['pronoun2cap'],
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
    doc = Document('templates/template_mod_12.docx')
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
                add_behavior_presentation(paragraph, st.session_state.behavior_observation)
            
            if "[[Developmental History]]" in paragraph.text:
                add_developmental_history(paragraph, st.session_state.development_history)

            if "[[Recommendations]]" in paragraph.text:
                if check_developmental_pediatrics:
                    add_developmental_pediatrics(paragraph)
                if check_feeding_treatment:
                    add_feeding_treatment(paragraph)
                if check_levine_clinic:
                    add_levine_clinic(paragraph)
                if check_parent_parent:
                    add_parent_parent(paragraph)
                if check_100_days:
                    add_100_days(paragraph)
                if check_caregiver_support:
                    add_caregiver_support(paragraph)
                if check_edu_placement:
                    add_edu_placement(paragraph)
                if check_effective_treatments:
                    add_effective_treatments(paragraph)
                if check_elopement_plan:
                    add_elopement_plan(paragraph)
                if check_develop_disability_office:
                    add_develop_disability_office(paragraph)
                if check_evidence_therapy:
                    add_evidence_therapy(paragraph)
                
                delete_paragraph(paragraph)
                
            if "SRS Report Information" in paragraph.text:
                # Add SCQ
                if scq_result:
                    add_scq_form(paragraph)
                # Add SRS
                if len(teacher_score) == 0:
                    add_srs_no_teacher(paragraph)
                else:
                    add_srs_yes_teacher(paragraph, teacher_score)
            
            if "Social Responsiveness Scale" in paragraph.text:
                if teacher_eval:
                    paragraph.add_run(" & teacher\nDevelopmental History & Review of Records\n", style='CustomStyle')
                    paragraph.add_run(f"School Report on SRS-2 provided by {teacher_score['{{Teacher name, title}}']}", style='CustomStyle')
                else:
                    paragraph.add_run("\nDevelopmental History & Review of Records", style='CustomStyle')

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