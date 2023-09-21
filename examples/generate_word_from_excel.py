"""
examples/generate_word_from_excel.py
generate_word_from_excel.py is a example that can render a word file with a dictionary and save it.
@Author: Fei Dongxu
@Date: 2023-9-21
@Version: 1.0
@License: Apache License 2.0
"""

from apps import Excel,Word

template = Word("ori_templates.docx")
data = Excel("ori_sheet.xlsx")
maps = data.get_header_mapping()
for insert in data.get_templates_render_dicts():
    template.render_word(insert)
    template.save(str(insert.get("id"))+".docx")

