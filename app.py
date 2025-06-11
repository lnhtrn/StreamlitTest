import streamlit as st
from docx import Document
import yaml
import io
import docxedit
import datetime
from docx.shared import Pt


col1,col2 = st.columns(2)
col1.title('Report Builder')

with open('misc_data/states.txt', 'r') as file:
    states = [x.strip() for x in file]

def format_date_with_ordinal(date_obj):
    day = date_obj.day
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return date_obj.strftime(f"%B {day}{suffix}, %Y")

# Set up dictionary to store data 
data = {}
optional = {}

####################################################
st.header("Appointment Summary")
audio_summary = st.audio_input("Summarize the appointment details")
st.markdown("***Check before proceeding with form:*** Scores to report:")
teacher_eval = st.checkbox("Teacher's SSR Scores")
wppsi_score = st.checkbox("Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed. (WPPSI) Score")
dppr_score = st.checkbox("Developmental Profile – Fourth Edition – Parent Report (DPPR)")
pls_score = st.checkbox("Preschool Language Scale – Fifth Edition (PLS)")
pdms_score = st.checkbox("Peabody Developmental Motor Scales – Second Edition")
peshv_score = st.checkbox("Preschool Evaluation Scale Home Version – Second Edition")
reelt_score = st.checkbox("Receptive Expressive Emergent Language Test – Fourth Edition")
abas_score = st.checkbox("Adaptive Behavior Assessment System – Third Edition")

with st.form('BasicInfo'):
    ####################################################
    st.header("Patient's data")
    # Dict to store data
    data['{{Patient First Name}}'] = st.text_input('Patient First Name')
    data['{{Patient Last Name}}'] = st.text_input('Patient Last Name')
    preferred = st.selectbox(
        "Patient's Preferred Pronoun",
        ("They/them", "He/him", "She/her"),
    )
    data["{{Patient Age}}"] = st.number_input("Patient's Age", 0, 100)
    data['age_unit'] = st.radio(
        "Year/month?",
        ("Year", "Month")
    )

    data['{{Caregiver type}}'] = st.selectbox(
        "Patient's Caregiver",
        ("mother", "father", "parent", "grandparent", "legal custodian", "Foster Parent"),
    )

    data['{{Caregiver Primary Concerns}}'] = st.multiselect(
        "Caregiver\'s Primary Concerns",
        (
            "Speech delays impacting social opportunities.",
            "Clarifying diagnostic presentation.",
            "Determining service eligibility.",
            "Language delays and difficulties.",
            "Elopement and related safety concerns.",
            "Determining appropriate supports."
        ),
        placeholder="Select from the choices or enter a new one",
        accept_new_options=True
    )
    data['{{Caregiver Primary Concerns}}'] = "\n".join(data['{{Caregiver Primary Concerns}}'])

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

    data["{{Result of the evaluation}}"] = st.multiselect(
        "Result of the evaluation",
        [
            "F84.0 - Autism Spectrum Disorder (per the above referenced evaluation)",
            "F88.0 - Global Developmental Delay (per behavioral presentation)",
            "F80.2 - Mixed Receptive-Expressive Language Disorder",
            "F90.2 - Attention Deficit Hyperactivity Disorder - Combined-Type",
            "F50.82 Avoidant/Restrictive Food Intake Disorder",
            "None"
        ],
        placeholder="Select from the choices or enter a new one",
        accept_new_options=True
    )

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
        optional["teacher"] = {}

        optional["teacher"]['{{SRS-2 Score Teacher}}'] = st.text_input("Teacher's SRS-2 Score")

        optional["teacher"]['{{Social Communication and Interaction Score Teacher}}'] = st.text_input("Social Communication and Interaction Score Teacher")

        optional["teacher"]['{{Restricted Interests and Repetitive Behavior Score Teacher}}'] = st.text_input("Restricted Interests and Repetitive Behavior Score Teacher")

        optional["teacher"]["{{Teacher's level of concern}}"] = st.radio(
            "Teacher's level of concern",
            ['no', 'mild', 'moderate', 'severe']
        )

    ######################################################
    st.header("Medical/Developmental History")
    
    data['{{Diagnosis History}}'] = st.multiselect(
        "Diagnosis History (Select or add your own)",
        ['History of language and social communication delays.'],
        accept_new_options=True
    )

    data['{{Medications}}'] = st.multiselect(
        "Medications (Select or add your own)",
        ['None noted or reported.'],
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

    data['{{Grade}}'] = st.selectbox(
        "Grade",
        ['EPK (2023-24 school year)', 'UPK (2023-24 school year)', 'Kindergarten (2023-24 school year)'],
        index=None,
        placeholder="Select a grade or enter a new one",
        accept_new_options=True,
    )

    data['{{Teacher name, title}}'] = st.text_input("Teacher name, title")

    data['{{Education Setting}}'] = st.selectbox(
        "Education Setting",
        ["General Education", "Integrated Co-Taught", "12:1:1", "8:1:1", "6:1:1"],
        index=None,
        placeholder="Select a grade or enter a new one",
        accept_new_options=True,
    )

    data['{{Services}}'] = st.multiselect(
        "Services",
        [
            "None",
            "Speech therapy",
            "Occupational therapy",
            "Physical therapy",
            "Extended school year services",
            "Testing accommodations"
        ],
        placeholder="Select from the choices or enter a new one",
        accept_new_options=True
    )

    ##########################################################
    if wppsi_score:
        st.header("Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed. (WPPSI)")
        st.markdown("*Skip this section if there is no WPPSI Score*")
        optional["wppsi"] = {}

        optional["wppsi"]["WPPSI Test Date"] = format_date_with_ordinal(st.date_input("WPPSI Test Date"))
        optional["wppsi"]['{{WPPSI Full Scale IQ Score}}'] = st.text_input("WPPSI Full Scale IQ Score")

        optional["wppsi"]['{{WPPSI Verbal Comprehension Score}}'] = st.text_input("WPPSI Verbal Comprehension Score")

        optional["wppsi"]['{{WPPSI Visual Spatial Score}}'] = st.text_input("WPPSI Visual Spatial Score")
    
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")
    # data['{{}}'] = st.text_input("")

    submit = st.form_submit_button('Submit')

def add_wppsi(paragraph, score_data, style):
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run(f'\t({score_data["WPPSI Test Date"]}) – Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed.', style=style).italic = True
    paragraph.insert_paragraph_before().add_run(f'\tFull Scale IQ: {score_data["WPPSI Full Scale IQ Score"]}', style=style).bold = True
    paragraph.insert_paragraph_before().add_run(f'\tVerbal Comprehension: {score_data["WPPSI Verbal Comprehension Score"]}\t\tVisual Spatial: {score_data["WPPSI Visual Spatial Score"]}', style=style)
    paragraph.insert_paragraph_before()

if submit:
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

    # format date time to something like May 23rd, 2025
    # for key, value in replace_word.items():
    #     if isinstance(value, datetime.date):
    #         replace_word[key] = value.strftime("%B %d, %Y")

    # Display data 
    yaml_string = yaml.dump(replace_word, sort_keys=False)
    yaml_data = st.code(yaml_string, language=None)
    

    #### Edit document 
    doc = Document('templates/template_mod_12.docx')
    if doc:
        ### create document style
        doc_style = doc.styles['Normal']
        font = doc_style.font
        font.name = 'Georgia'
        font.size = Pt(12)

        # Edit document
        for word in replace_word:
            docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

        # Add scores 
        if len(optional) > 0:
            for i, paragraph in enumerate(doc.paragraphs):
                if "Scores are reported here as standard scores" in paragraph.text:
                    if 'wppsi' in optional:
                        add_wppsi(paragraph, optional['wppsi'], doc_style)

        # Save content to file
        bio = io.BytesIO()
        doc.save(bio)

        today_date = format_date_with_ordinal(datetime.date.today())
        
        # Download 
        st.download_button(
            label="Click here to download",
            data=bio.getvalue(),
            file_name=f"{data['{{Patient First Name}}']} {data['{{Patient Last Name}}']} {today_date}.docx",
            mime="docx"
        )