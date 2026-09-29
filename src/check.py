from pathlib import Path

from docx import Document
import os
from src.documentProcessors.paragraphProcessor import ParagraphProcessor

currentDirectory = Path(__file__).resolve().parent

filePath = 'samples/or_submission_2_4_1.docx'
document = Document(os.path.join(currentDirectory, filePath))

paragraphs = ParagraphProcessor(document.paragraphs)

title = paragraphs.getMainTitle()
print('tk main', title)
