import streamlit as st
from docx import Document
import yaml
import io
import docxedit
import datetime
import re 
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE, WD_STYLE
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_ALIGN_PARAGRAPH


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

    # Edit document
    # for word in replace_word:
    #     docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])
    # docxedit.replace_string(doc, old_string='[[Patient First Name]]', new_string='Linh')
    # docxedit.replace_string(doc, old_string='[[Patient Last Name]]', new_string='Tran')

#     wais_score = """99,94-104,47
# 111,103-117,77 
# 100,93-107,50 
# """

#     wais_subtest_score = """Verbal Comprehension Index:
#   Similarities: 13
#   Vocabulary: 11

# Working Memory Index:
#   Digit Sequencing: 7
#   Running Digits: 8

# Visual Spatial Index:
#   Block Design: 9
#   Visual Puzzles: 11

# Processing Speed Index:
#   Symbol Search: 9
#   Coding: 10

# Fluid Reasoning Index:
#   Matrix Reasoning: 12
#   Figure Weights: 8
# """
#     wais_score = """Full Scale IQ:
#   Standard Score: 99
#   Confidence Interval: 94-104
#   Percentile: 47

# Verbal Comprehension Index:
#   Standard Score: 111
#   Confidence Interval: 103-117
#   Percentile: 77

# Visual Spatial Index:
#   Standard Score: 100
#   Confidence Interval: 93-107
#   Percentile: 50

# Fluid Reasoning Index:
#   Standard Score: 100
#   Confidence Interval: 93-107
#   Percentile: 50

# Working Memory Index:
#   Standard Score: 85
#   Confidence Interval: 79-93
#   Percentile: 16

# Processing Speed Index:
#   Standard Score: 97
#   Confidence Interval: 89-106
#   Percentile: 42
# """

#     wais_subtest_score = yaml.safe_load(wais_subtest_score)
#     print(wais_subtest_score)

#     wais_score = yaml.safe_load(wais_score)
#     print(wais_score)


    # info_list = ['IQ', 'Verbal Comp', 'Visual Spatial']

    # replace_word = {}
    # replace_percent = {}

    # for line, info in zip(wais_score.split("\n"), info_list):
    #     line_items = line.split(",")
    #     replace_word[f"[[{info} Standard]]"] = line_items[0].strip()
    #     replace_word[f"[[{info} CI]]"] = line_items[1].strip()
    #     replace_percent[f"[[{info} Percent]]"] = line_items[2].strip()

    # # Edit document
    # for word in replace_word:
    #     docxedit.replace_string(doc, old_string=word, new_string=replace_word[word])

    # # Replace percent 
    # for table in doc.tables:
    #     for row in table.rows:
    #         for cell in row.cells:
    #             for paragraph in cell.paragraphs:
    #                 for key in replace_percent:
    #                     if key in paragraph.text:
    #                         p = paragraph.insert_paragraph_before()
    #                         p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    #                         p.add_run(replace_percent[key], style='CustomStyle')
    #                         suffix = get_ordinal(replace_percent[key])
    #                         p.add_run(suffix, style='CustomStyle').font.superscript = True
    #                         delete_paragraph(paragraph)

    wais_analysis = """\tDeanna’s Full-Scale IQ on the WAIS-5 was 99, which is at the 47th percentile. This score indicates that Deanna’s overall cognitive functioning is within the Average Range, with a 95 percent chance of falling between 94 and 104.

\n\n\tThe Verbal Comprehension Index (VCI) measures language development, vocabulary, and other verbal skills. Deanna received a score of 111, which places her in the 77th percentile compared to her same age peers. Her score on this index fell within the Above Average Range. She had similar scores on the two subtests (Similarities and Vocabulary) within this index measuring her ability to use verbal and crystallized abilities, word knowledge, and abstract and conceptual reasoning. This index reflects a relative strength within her cognitive profile.

\n\n\tThe Visual Spatial Index (VSI) measures visual processing and perceptual reasoning skills. Deanna received a score of 100, which places her in the 50th percentile. Her score on this index fell within the Average Range. The subtests within this domain (Block Design and Visual Puzzles) indicate her ability to analyze and synthesize abstract visual stimuli, reflecting average visual-spatial reasoning and perceptual skills."""

    # Test paragraph 
    for paragraph in doc.paragraphs:
        if "[[WAIS-Analysis]]" in paragraph.text:
            replace_ordinal_with_superscript(paragraph, wais_analysis)

        # if "[[Test Percentile]]" in paragraph.text:
        #     replace_with_superscript(paragraph, "[[Test Percentile]]", "50")

doc.save("Test_percent_mod4.docx")