from docxtpl import DocxTemplate
import yaml


tpl=DocxTemplate('output\Linh Tran June 20th, 2025.docx')

with open('data.yaml', 'r') as file:
    context = yaml.safe_load(file)
# context = {
#     'bullets': [
#         "Awkward social initiation and response",
#         "Difficulties with chit-chat",
#         "Difficulty interpreting figurative language",
#     ],
# }

tpl.render(context)
tpl.save("output.docx")
