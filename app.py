import streamlit as st
from docx import Document
import yaml
import io


col1,col2 = st.columns(2)
col1.title('Report Builder')

with open('misc_data\states.txt', 'r') as file:
    states = [x.strip() for x in file]

# Set up dictionary to store data 
data = {}

with st.form('addition'):
    st.title("Patient's Demographic")
    # Dict to store data
    demographic = {}
    demographic['firstname'] = st.text_input('Patient First Name')
    demographic['lastname'] = st.text_input('Patient Last Name')
    demographic['pronoun'] = st.selectbox(
        "Patient's Preferred Pronoun",
        ("They/them", "He/him", "She/her"),
    )
    age1,age2 = st.columns(2)
    with age1:
        demographic['age_amount'] = st.number_input("Patient's Age", 0, 100)
    with age2:
        demographic['age_unit'] = st.selectbox(
            "Year/month?",
            ("Year", "Month"),
        )

    demographic['caregiver'] = st.selectbox(
        "Patient's Caregiver",
        ("Mother", "Father", "Parent", "Grandparent", "Legal Custodian", "Foster Parent"),
    )

    demographic['primary_concern'] = st.multiselect(
        "Caregiver\'s Primary Concerns",
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

    demographic['state'] = st.selectbox(
        "Residence City/State", states
    )

    demographic['narrative'] = st.text_area('Narrative to finish \"Patient lives with...\"')
    data['Demographic'] = demographic
    
    submit = st.form_submit_button('Submit')
    

if submit:
    doc = Document('templates\\template_mod_12.docx')

    # handle word to replace 
    # pronouns
    with open("misc_data\pronouns.yaml", "r") as file:
        pronoun = yaml.safe_load(file)
    preferred = data['Demographic']['pronoun']

    replace_word = {
        "{{Preferred Pronouns 1}}": pronoun[preferred]['pronoun1'],
        "{{Preferred Pronouns 1 CAP}}": pronoun[preferred]['pronoun1cap'],
        "{{Preferred Pronouns 2}}": pronoun[preferred]['pronoun2'],
        "{{Preferred Pronouns 2 CAP}}": pronoun[preferred]['pronoun2cap'],
    }

    for word in replace_word:
        for p in doc.paragraphs:
            print(type(p.text))
            if p.text.find(word) >= 0:
                p.text = p.text.replace(word, replace_word[word])
    
    doc.save(f'output/Report_{data['Demographic']['firstname']}_{data['Demographic']['lastname']}.docx')

    bio = io.BytesIO()
    doc.save(bio)

    if doc:
        st.download_button(
            label="Click here to download",
            data=bio.getvalue(),
            file_name=f"Report_{data['Demographic']['firstname']}_{data['Demographic']['lastname']}.docx.docx",
            mime="docx"
        )