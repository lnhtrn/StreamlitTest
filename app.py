import streamlit as st
from docx import Document
import yaml
import io
import docxedit


col1,col2 = st.columns(2)
col1.title('Report Builder')

with open('misc_data/states.txt', 'r') as file:
    states = [x.strip() for x in file]

# Set up dictionary to store data 
data = {}

with st.form('BasicInfo'):
    ####################################################
    st.header("Appointment Summary")
    audio_summary = st.audio_input("Summarize the appointment details")

    ####################################################
    st.header("Patient's data")
    # Dict to store data
    data['{{Patient First Name}}'] = st.text_input('Patient First Name')
    data['{{Patient Last Name}}'] = st.text_input('Patient Last Name')
    preferred = st.selectbox(
        "Patient's Preferred Pronoun",
        ("They/them", "He/him", "She/her"),
    )
    data["{{Patient's Age}}"] = st.number_input("Patient's Age", 0, 100)
    data['age_unit'] = st.selectbox(
        "Year/month?",
        ("Year", "Month"),
    )

    data['{{Caregiver type}}'] = st.selectbox(
        "Patient's Caregiver",
        ("mother", "father", "parent", "grandparent", "legal custodian", "Foster Parent"),
    )

    data['{{Caregiver\'s Primary Concerns}}'] = st.multiselect(
        "Caregiver\'s Primary Concerns (select or add your own)",
        (
            "Speech delays impacting social opportunities.",
            "Clarifying diagnostic presentation.",
            "Determining service eligibility.",
            "Language delays and difficulties.",
            "Elopement and related safety concerns.",
            "Determining appropriate supports."
        ),
        accept_new_options=True
    )
    data['{{Caregiver\'s Primary Concerns}}'] = "\n".join(data['{{Caregiver\'s Primary Concerns}}'])

    data['{{Residence City/State}}'] = st.selectbox(
        "Residence City/State", states
    )

    data['{{Narrative}}'] = st.text_area('Narrative to finish \"Patient lives with...\"')

    ##########################################################
    st.header("BRH Evaluation Details")

    data['{{Evaluation Date}}'] = st.date_input("Evaluation Date")

    data['{{Module used}}'] = st.radio("Module used", ["Module 1", "Module 2"])
    if data['{{Module used}}'] == "Module 1":
        data['{{Module Description}}'] = "Module 1 is designed for children with single words"
    else:
        data['{{Module Description}}'] = "Module 2 is designed for children with phrase speech"

    data['{{Location of the evaluation}}'] = st.radio(
        "Location of the evaluation",
        ['home', 'school', 'the office']
    )

    data['{{Results Shared Date}}'] = st.date_input("Results Shared Date")
    
    data['{{Date Report Sent to Patient}}'] = st.date_input("Date Report Sent to Patient")

    data["{{Result of the evaluation}}"] = st.multiselect(
        "Result of the evaluation",
        (
            "F84.0 - Autism Spectrum Disorder (per the above referenced evaluation)",
            "F88.0 - Global Developmental Delay (per behavioral presentation)",
            "F80.2 - Mixed Receptive-Expressive Language Disorder",
            "F90.2 - Attention Deficit Hyperactivity Disorder - Combined-Type",
            "F50.82 Avoidant/Restrictive Food Intake Disorder",
            "None",
        ),
        accept_new_options=True
    )

    data["{{Results (SCQ) - Lifetime Form}}"] = st.text_input(
        "Results (SCQ) - Lifetime Form"
    )

    data["{{SRS-2 Score Caregiver}}"] = st.text_input("SRS-2 Score Caregiver")
    
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

    teacher_eval = st.checkbox("Teacher gave SSR Scores")

    ######################################################

    submit = st.form_submit_button('Submit')



if submit:
    doc = Document('templates/template_mod_12.docx')

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

    # Display data 
    yaml_string = yaml.dump(replace_word, sort_keys=False)
    yaml_data = st.code(yaml_string, language=None)
    
    for word in replace_word:
        docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

    bio = io.BytesIO()
    doc.save(bio)

    if doc:
        st.download_button(
            label="Click here to download",
            data=bio.getvalue(),
            file_name=f"Report_{data['{{Patient First Name}}']}_{data['{{Patient Last Name}}']}.docx",
            mime="docx"
        )