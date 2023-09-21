"""
example/fulfill_docx.py
fulfill_docx.py is a example that can render a docx file with a dictionary and save it.
It can also convert a docx file to a pdf file.
@Author: Fei Dongxu
@Date: 2023-9-21
@Version: 1.0
@License: Apache License 2.0
"""

from apps import Word

word = Word("templates.docx")
word.render_word({"name": "this is name", "obj": "this is a unused variable"})
word.save("result.docx")
