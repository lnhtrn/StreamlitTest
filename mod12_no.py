import streamlit as st
from docx import Document
import yaml
import io
import docxedit
import datetime
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from streamlit_gsheets import GSheetsConnection
from docxtpl import DocxTemplate
from docx.shared import Inches
from openai import OpenAI

##########################################################
st.set_page_config(
    page_title="Module 1&2 No Autism",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="expanded",
)

##########################################################
# Set up OpenAI 
if 'behavior_observation' not in st.session_state:
    st.session_state.behavior_observation = ""

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

# Create a connection object.
connections['noAutism'] = st.connection(f"mod12_noAutism", type=GSheetsConnection)
# Read object
df = connections['noAutism'].read(
    ttl="30m",
    usecols=list(range(3)),
    nrows=30,
) 
for col_name in df.columns:
    dropdowns[col_name] = df[col_name].tolist()
    dropdowns[col_name] = [x for x in dropdowns[col_name] if str(x) != 'nan']

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
    # scq_result = st.checkbox("Social Communication Questionnaire (SCQ) - Lifetime Form")
    # teacher_eval = st.checkbox("Teacher's SSR Scores")
    wppsi_score = st.checkbox("Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed. (WPPSI) Score")
    dppr_score = st.checkbox("Developmental Profile – Fourth Edition - Parent Report (DPPR)")
    pls_score = st.checkbox("Preschool Language Scale - Fifth Edition (PLS)")
    pdms_score = st.checkbox("Peabody Developmental Motor Scales - Second Edition")
    peshv_score = st.checkbox("Preschool Evaluation Scale Home Version - Second Edition")
    reelt_score = st.checkbox("Receptive Expressive Emergent Language Test - Fourth Edition")
    abas_score = st.checkbox("Adaptive Behavior Assessment System - Third Edition")


col1,col2 = st.columns(2)
col1.title('Module 1&2 No Autism Report Builder')

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

if st.button("Transcribe"):
    if audio_behavior:
        transcript_behavior = transcribe_audio(audio_behavior, name='behavior')
        st.markdown(f"**Transcription:** {transcript_behavior}")
        
        response = client.responses.create(
            prompt={
                "id": st.secrets["behavior_prompt_mod12no_id"],
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

with st.form('BasicInfo'):
    ####################################################
    st.header("Patient's data")
    # Dict to store data
    preferred = st.selectbox(
        "Patient's Preferred Pronoun",
        ("They/them", "He/him", "She/her"),
    )
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
        dropdowns['Caregiver\'s Primary Concerns'],
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

    ######################################################
    st.header("Medical/Developmental History")
    
    lines['{{Developmental Concerns}}'] = st.multiselect(
        "Developmental Concerns",
        dropdowns['Developmental Concerns'],
        placeholder="Select multiple options from the list or enter a new one",
        accept_new_options=True
    )

    lines['{{Medical Concerns}}'] = st.multiselect(
        "Medical Concerns",
        dropdowns['Medical Concerns'],
        placeholder="Can input multiple options",
        accept_new_options=True
    )

    bullet['CaregiverDevelopmentalConcerns'] = st.multiselect(
        "Caregiver\'s Developmental Concerns",
        dropdowns['Caregiver\'s Developmental Concerns'],
        placeholder="Select multiple options from the list or enter a new one",
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
        # dropdowns["Grade"],
        # index=None,
        # placeholder="Select a grade or enter a new one",
        # accept_new_options=True,
    )

    data['School Year'] = st.text_input(
        "School Year",
        # dropdowns["School Year"],
        # index=None,
        # placeholder="Select a grade or enter a new one",
        # accept_new_options=True,
    )

    data['{{Education Setting}}'] = st.selectbox(
        "Education Setting",
        [
            "General Education", 
            "Integrated Co-Taught", 
            "12:1:1", 
            "8:1:1", 
            "6:1:1"
        ],
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
        st.header("Receptive Expressive Emergent Language Test - Fourth Edition (REELT)")
        st.markdown("*Skip this section if there is no REELT Score*")
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

def add_school(paragraph):
    p = paragraph.insert_paragraph_before()
    tab_stops = p.paragraph_format.tab_stops
    # tab_stops.clear()  # Start fresh for this paragraph only
    tab_stops.add_tab_stop(Inches(3.5))
    # Add data
    p.add_run("District", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{School District}}']}\t", style='CustomStyle')
    p.add_run("Grade", style='CustomStyle').font.underline = True
    ### italics for school year
    p.add_run(f": {data['{{Grade}}']} (", style='CustomStyle')
    p.add_run(f"{data['School Year']})\n\n", style='CustomStyle').italic = True
    p.add_run("School", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{School Name}}']}\t", style='CustomStyle')
    p.add_run("Setting", style='CustomStyle').font.underline = True
    p.add_run(f": {data['{{Education Setting}}']}", style='CustomStyle')
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


if submit:
    # Update session state 
    st.session_state.behavior_observation = data['behavior_observation']

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
    doc = Document('templates/template_mod_12_no_autism.docx')
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
        # custom_style.paragraph_format.line_spacing = 1 

        # list_style = doc.styles.add_style('ListStyle', WD_STYLE.LIST_BULLET)
        # list_style.font.size = Pt(12)
        # list_style.font.name = 'Georgia'

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
                
            # if "SRS Report Information" in paragraph.text:
            #     # Add SCQ
            #     if scq_result:
            #         add_scq_form(paragraph)
            #     # Add SRS
            #     if len(teacher_score) == 0:
            #         add_srs_no_teacher(paragraph)
            #     else:
            #         add_srs_yes_teacher(paragraph, teacher_score)
            
            # if "Social Responsiveness Scale" in paragraph.text:
                # if teacher_eval:
                #     paragraph.add_run(" & teacher\nDevelopmental History & Review of Records\n", style='CustomStyle')
                #     paragraph.add_run(f"School Report on SRS-2 provided by {teacher_score['{{Teacher name, title}}']}", style='CustomStyle')
                # else:
                #     paragraph.add_run("\nDevelopmental History & Review of Records", style='CustomStyle')

            if "[[Behavioral Presentation]]" in paragraph.text:
                add_behavior_presentation(paragraph, st.session_state.behavior_observation)
            
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