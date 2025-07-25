import streamlit as st
import yaml
from docx import Document
import io
import docxedit
import datetime
import re 
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE, WD_STYLE
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
import pandas as pd
from st_aggrid import AgGrid

# Sample data
df = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'San Francisco', 'London']
})

#########################################################
# Load OpenAI client 
client = OpenAI(api_key=st.secrets["openai_key"])

data = {}
wais_data = {}

#########################################################
with st.form("BasicInfo"):  
    st.header("Patient's data")
    
    data['{{Patient First Name}}'] = st.text_input('Patient First Name')

    data['{{Patient Last Name}}'] = st.text_input('Patient Last Name')

    data["{{Patient Age}}"] = st.number_input("Patient's Age", 0, 100)

    data['{{Patient age unit}}'] = st.radio(
        "Year/month?",
        ("year", "month")
    )

    preferred = st.selectbox(
        "Patient's Preferred Pronoun",
        ("They/them", "He/him", "She/her"),
    )

    ################################################# 
    st.header("WAIS-5 Score Report")

    wais_data['overall'] = st.text_area(
        "WAIS-5 Overall Score - Input percentile without any suffix. For example: Percentile: 42", 
        """Full Scale IQ:
  Standard Score: 
  Confidence Interval: 
  Percentile: 

Verbal Comprehension Index:
  Standard Score: 
  Confidence Interval: 
  Percentile: 

Visual Spatial Index:
  Standard Score: 
  Confidence Interval: 
  Percentile:

Fluid Reasoning Index:
  Standard Score: 
  Confidence Interval:
  Percentile: 

Working Memory Index:
  Standard Score: 
  Confidence Interval: 
  Percentile: 

Processing Speed Index:
  Standard Score: 
  Confidence Interval: 
  Percentile: 
""",
        height=350,
    )

    wais_data['subtest'] = st.text_area(
        "WAIS-5 Subtest Score", 
        """Verbal Comprehension Index:
  Similarities: 
  Vocabulary: 

Working Memory Index:
  Digit Sequencing: 
  Running Digits: 

Visual Spatial Index:
  Block Design: 
  Visual Puzzles: 

Processing Speed Index:
  Symbol Search: 
  Coding: 

Fluid Reasoning Index:
  Matrix Reasoning: 
  Figure Weights: 
""",
        height=350,
    )

    #############################################
    # First table
    st.header("Editable Table")
    grid_return = AgGrid(df, editable=True)

    submit = st.form_submit_button('Submit')

###########################################################


def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def get_ordinal(number):
    suffix = 'th' if 11 <= int(number) <= 13 else {"1": 'st', "2": 'nd', "3": 'rd'}.get(number[-1], 'th')
    return suffix

def replace_with_superscript(para, old_text, number_part):
    superscript_part = get_ordinal(number_part)
    if old_text in para.text:
        # Save everything before, match, and after
        before, match, after = para.text.partition(old_text)
        
        # Clear existing runs
        para.clear()
        
        # Add before text
        para.add_run(before, style='CustomStyle')
        
        # Add number part
        para.add_run(number_part, style='CustomStyle')
        
        # Add superscript part
        sup_run = para.add_run(superscript_part, style='CustomStyle')
        sup_run.font.superscript = True
        
        # Add after text
        para.add_run(after, style='CustomStyle')


def replace_ordinal_with_superscript(para, full_text):
    # Regex to find ordinal numbers like 1st, 2nd, 77th, 103rd, etc.
    pattern = re.compile(r'(\d+)(st|nd|rd|th)')
    paragraph = para.insert_paragraph_before()  # Remove all existing runs

    last_index = 0
    for match in pattern.finditer(full_text):
        # Add text before the match
        paragraph.add_run(full_text[last_index:match.start()], style='CustomStyle')

        # Add number part (e.g., "77")
        paragraph.add_run(match.group(1), style='CustomStyle')

        # Add superscript suffix (e.g., "th")
        sup_run = paragraph.add_run(match.group(2), style='CustomStyle')
        sup_run.font.superscript = True

        last_index = match.end()

    # Add remaining text after last match
    paragraph.add_run(full_text[last_index:], style='CustomStyle')
    delete_paragraph(para)


###########################################################

if submit:
    # Display the newly edited dataframe
    st.subheader("Updated Data")
    st.dataframe(grid_return['data'])

    
    wais_analysis = ""
    # if wais_data["subtest"] and wais_data['overall']:
    #     response = client.responses.create(
    #         prompt={
    #             "id": st.secrets["wais_analysis_id"],
    #             "variables": {
    #                 "first_name": data['{{Patient First Name}}'],
    #                 "pronouns": preferred,
    #                 "wais_subtest": wais_data['subtest'],
    #                 "wais_overall": wais_data['overall'],
    #             }
    #         }
    #     )
    #     wais_analysis = response.output_text

    # Edit Table Score
    replace_word = {}
    replace_percent = {}
        
    wais_subtest_score = yaml.safe_load(wais_data['subtest'])
    wais_score = yaml.safe_load(wais_data['overall'])

    info_list = ['IQ', 'Verbal Comp', 'Visual Spatial']

    for info in wais_score:
        # Info meaning type of score, i.e. "Full Scale IQ", "Verbal Comprehension Index, etc"
        replace_word[f"[[{info} Standard]]"] = str(wais_score[info]['Standard Score'])
        replace_word[f"[[{info} CI]]"] = str(wais_score[info]['Confidence Interval'])
        replace_percent[f"[[{info} Percent]]"] = str(wais_score[info]['Percentile'])
    
    for info in wais_subtest_score:
        for subtest in wais_subtest_score[info]:
            replace_word[f"[[{subtest}]]"] = str(wais_subtest_score[info][subtest])

    # Display data 
    yaml_string = yaml.dump(data, sort_keys=False)
    yaml_string += "\nOverall:\n" + yaml.dump(replace_word) + yaml.dump(replace_percent)
    yaml_string += "\n\n" + wais_analysis
    yaml_data = st.code(yaml_string, language=None)

    # Edit document 
    doc = Document('templates/template_mod_4.docx')
    if doc:
        ### create document style
        doc_style = doc.styles['Normal']
        font = doc_style.font
        font.name = 'Georgia'
        font.size = Pt(12)

        custom_style = doc.styles.add_style('CustomStyle', WD_STYLE_TYPE.CHARACTER)
        custom_style.font.size = Pt(12)
        custom_style.font.name = 'Georgia'

        # Replace words even in a table 
        for word in replace_word:
            docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

        # Replace percent in table 
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        # Loop through all percentage needed to be replaced in the table
                        for key in replace_percent:
                            if key in paragraph.text:
                                p = paragraph.insert_paragraph_before()
                                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                p.add_run(replace_percent[key], style='CustomStyle')
                                suffix = get_ordinal(replace_percent[key])
                                p.add_run(suffix, style='CustomStyle').font.superscript = True
                                delete_paragraph(paragraph)

        # Now do the analysis replacement
        for paragraph in doc.paragraphs:
            if "[[WAIS-Analysis]]" in paragraph.text:
                replace_ordinal_with_superscript(paragraph, wais_analysis)

        # Get file name
        filename = f"{data['{{Patient First Name}}']} {data['{{Patient Last Name}}']} test Mod 4 WAIS Score.docx"

        # Save content to file
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


    
    
