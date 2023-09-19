from docxtpl import DocxTemplate
from .Logger import logger
import os


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
