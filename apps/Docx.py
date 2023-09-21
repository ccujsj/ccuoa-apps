"""
app/Docx.py
Docx.py is a module that can render a docx file with a dictionary and save it.
It can also convert a docx file to a pdf file.
@Author: Fei Dongxu
@Date: 2023-9-21
@Version: 1.0
@License: Apache License 2.0
"""


from docxtpl import DocxTemplate
from .Logger import logger
import os
import win32com.client as win32
from docx import Document

class Word:
    def __init__(self, filename):
        self.word = DocxTemplate(filename)
        logger.debug(filename + " loaded success")
        self.filename = filename
        pass

    def render_word(self, context: dict):
        self.word.render(context)

    def set_filename(self, filename):
        self.filename = filename

    def save(self, filepath=None):
        if filepath:
            self.word.save(filepath)
            logger.debug(filepath + " saved at " + str(os.getcwd()))
            return
        self.word.save(self.filename)
        logger.debug(self.filename + " saved at default position: " + str(os.getcwd()))

    def convert_to_pdf(self,input_path,output_path):
        # create a Microsoft Word application object
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        # set visibility to 0 (hidden)
        word_app.Visible = False
        try:
            # convert docx file 1 to pdf file 1
            doc = word_app.Documents.Open(input_path)
            doc.SaveAs(output_path, FileFormat=17)
            doc.Close()
            logger.debug("Convert to pdf success")
        except Exception as e:
            logger.error("Convert to pdf failed")
            logger.error(e)
        finally:
            word_app.Quit()
        