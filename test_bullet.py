from docxtpl import DocxTemplate

tpl=DocxTemplate('templates/template_mod_12_noBrackets.docx')

context = {
    'bullets': [
        "Awkward social initiation and response",
        "Difficulties with chit-chat",
        "Difficulty interpreting figurative language",
    ],
}

tpl.render(context)
tpl.save("output.docx")
