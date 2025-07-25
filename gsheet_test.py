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
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from modules.recommendations import *

#########################################################
# Load OpenAI client 
client = OpenAI(api_key=st.secrets["openai_key"])

data = {}
wais_data = {}

#########################################################
def st_normal():
    _, col, _ = st.columns([1, 2, 1])
    return col

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
    st.header("Informant's Report - Vineland Adaptive Behavior Scales")

    # Load data
    df = pd.read_csv("misc_data/vineland_informant.csv")

    # JavaScript code to apply bold styling if "bold" column is True
    row_style_jscode = JsCode("""
    function(params) {
        if (params.data.bold === true) {
            return {
                'font-weight': 'bold',
                'font-size': 16,
            }
        } else {
            return {
                'font-size': 16,
            }
        }
        return {};
    }
    """)

    # Build Grid Options
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_grid_options(getRowStyle=row_style_jscode)
    gb.configure_column("data", editable=True)
    gridOptions = gb.build()

    # Display grid
    # with st_normal():
    grid_return = AgGrid(
        df,
        gridOptions=gridOptions,
        editable=True,
        height=800,
        theme="balham",
        # custom_css=custom_css,
        allow_unsafe_jscode=True
    )

    #############################################
    submit = st.form_submit_button('Submit')


###########################################################

if submit:
    # Display the newly edited dataframe
    st.subheader("Updated Data")
    st.dataframe(grid_return['data'])

    # WAIS Score
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

    # WAIS Score 
        
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


    # Vineland Informant Score
    vineland_info_dict = {
        "[[{}]]".format(row["col1"]): row["col2"] 
        for _, row in grid_return['data'].iterrows()
    }

    vineland_perc_dict = {
        "[[{}]]".format(row["col1"]): row["col2"] 
        for _, row in grid_return['data'].iterrows()
        if "Percentile" in row["col1"]
    }

    replace_percent.update(vineland_perc_dict)
    replace_word.update(vineland_info_dict)

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

        # Replace percent in table 
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        # Loop through all percentage needed to be replaced in the table
                        for key in replace_percent:
                            if key in paragraph.text:
                                write_ordinal(paragraph, replace_percent[key])

        # Replace words even in a table 
        for word in replace_word:
            docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

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


    
    
