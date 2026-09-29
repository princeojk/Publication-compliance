from docx.text.paragraph import Paragraph
from sections.mainTitle import MainTitle

class ParagraphProcessor:
    def __init__(self, paragraph: list[Paragraph]):
        self.paragraph = paragraph

    def getMainTitle(self):
        mainTitle =  MainTitle(self.paragraph[0])
        return mainTitle.serializer()


    
