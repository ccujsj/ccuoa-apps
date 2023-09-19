from apps import Excel,Word

template = Word("ori_templates.docx")
data = Excel("ori_sheet.xlsx")
maps = data.get_header_mapping()
for insert in data.get_templates_render_dicts():
    template.render_word(insert)
    template.save(str(insert.get("id"))+".docx")

