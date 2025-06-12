import streamlit as st
from docx import Document
import yaml
import io
import docxedit
import datetime
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_BREAK


#### Edit document 
def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def add_srs_no_teacher(paragraph, score_data):
    paragraph.insert_paragraph_before().add_run('Social Responsiveness Scale – Second Edition (SRS-2) – Parent', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run('The SRS-2 is an objective measure that identifies social impairments associated with autism spectrum disorder and quantifies ASD-related severity throughout the lifespan. \nThe following interpretative guidelines are offered here for the benefit of the reader: Less than 59 indicates within normal limits, between 60 and 65 as mild concern, between 65 and 75 as moderate concern, and greater than 76 as severe concern. ', style='CustomStyle')
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run('SRS-2 Total Score: {{SRS-2 Score Caregiver}} ({{Caregiver type}})', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run('Social Communication and Interaction: {{Social Communication and Interaction Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before().add_run('Restricted Interests and Repetitive Behavior: {{Restricted Interests and Repetitive Behavior Score Caregiver}} ({{Caregiver type}})', style='CustomStyle')
    paragraph.insert_paragraph_before()
    observe = paragraph.insert_paragraph_before()
    observe.add_run("Based on the report provided by {{Preferred Pronouns 2}} {{Caregiver type}}, ", style='CustomStyle')
    observe.add_run("{{Patient First Name}}’s social communication and related behaviors indicated {{Caregiver's level of concern}} concerns. ", style='CustomStyle').italic = True
    observe.add_run("My observation aligned with a {{Evaluator's level of concern}} level of concern", style='CustomStyle').bold = True
    delete_paragraph(paragraph)
    observe.add_run().add_break(WD_BREAK.PAGE)

def add_srs_yes_teacher(paragraph, score_data):
    paragraph.insert_paragraph_before().add_run('Social Responsiveness Scale – Second Edition (SRS-2) – Parent', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run('The SRS-2 is an objective measure that identifies social impairments associated with autism spectrum disorder and quantifies ASD-related severity throughout the lifespan. \nThe following interpretative guidelines are offered here for the benefit of the reader: Less than 59 indicates within normal limits, between 60 and 65 as mild concern, between 65 and 75 as moderate concern, and greater than 76 as severe concern. ', style='CustomStyle')
    paragraph.insert_paragraph_before()
    paragraph.insert_paragraph_before().add_run('SRS-2 Total Score: {{SRS-2 Score Caregiver}} ({{Caregiver type}}), {{SRS-2 Score Teacher}} (teacher)', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before()
    p = paragraph.insert_paragraph_before()
    p.add_run('Social Communication and Interaction: {{Social Communication and Interaction Score Caregiver}} ({{Caregiver type}}), ', style='CustomStyle')
    p.add_run(f'{92} (teacher)', style='CustomStyle')
    p = paragraph.insert_paragraph_before()
    p.add_run('Restricted Interests and Repetitive Behavior: {{Restricted Interests and Repetitive Behavior Score Caregiver}} ({{Caregiver type}}), ', style='CustomStyle')
    p.add_run(f'{92} (teacher)', style='CustomStyle')
    paragraph.insert_paragraph_before()
    observe = paragraph.insert_paragraph_before()
    observe.add_run("Based on the report provided by {{Preferred Pronouns 2}} {{Caregiver type}}, ", style='CustomStyle')
    observe.add_run("{{Patient First Name}}’s social communication and related behaviors indicated {{Caregiver's level of concern}} concerns. ", style='CustomStyle').italic = True
    observe.add_run("{{Patient First Name}}’s teacher reported a ", style='CustomStyle')
    observe.add_run(f"{86} level of concern, and ", style='CustomStyle')
    observe.add_run("my observation aligned with a {{Evaluator's level of concern}} level of concern.", style='CustomStyle').bold = True
    delete_paragraph(paragraph)
    observe.add_run().add_break(WD_BREAK.PAGE)

def add_wppsi(paragraph, score_data):
    paragraph.insert_paragraph_before().add_run(f'\t(23/5) – Wechsler Preschool & Primary Scales of Intelligence – Fourth Ed.', style='CustomStyle').italic = True
    paragraph.insert_paragraph_before().add_run(f'\tFull Scale IQ: 32', style='CustomStyle').bold = True
    paragraph.insert_paragraph_before().add_run(f'\tVerbal Comprehension: 12\t\t\tVisual Spatial: 12', style='CustomStyle')
    paragraph.insert_paragraph_before()


'''
({{DPPR Test Date}}) – Developmental Profile – Fourth Edition – Parent Report
Cognitive: {{DPPR Cognitive Score}} 				Social-Emotional: {{DPPR Social-Emotional Score}}
Adaptive: {{DPPR Adaptive Score}} 				Physical: {{DPPR Physical Score}}
'''

doc = Document('templates/template_mod_12_noScore.docx')
if doc:
    ### create document style
    doc_style = doc.styles['Normal']
    font = doc_style.font
    font.name = 'Georgia'
    font.size = Pt(12)

    custom_style = doc.styles.add_style('CustomStyle', WD_STYLE_TYPE.CHARACTER)
    custom_style.font.size = Pt(12)
    custom_style.font.name = 'Georgia'

    # Edit document
    # for word in replace_word:
    #     docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

    # Add scores 
    # if len(optional) > 0:
    for i, paragraph in enumerate(doc.paragraphs):
        if "Scores are reported here as standard scores" in paragraph.text:
            # if 'wppsi' in optional:
            add_wppsi(paragraph, dict())

        if "SRS Report Information" in paragraph.text:
            add_srs_yes_teacher(paragraph, dict())

doc.save("New_file.docx")