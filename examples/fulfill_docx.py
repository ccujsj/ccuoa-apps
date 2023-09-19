from apps import Word

word = Word("templates.docx")
word.render_word({"name": "this is name", "obj": "this is a unused variable"})
word.save("result.docx")
