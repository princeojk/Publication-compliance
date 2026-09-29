from xmlrpc.client import boolean

from docx.styles.style import ParagraphStyle
from docx.styles.style import BaseStyle
from docx.text.font import Font

class FontStyle:
    def __init__(self, object: ParagraphStyle):
        self.object = object

    def isBold(self):
        return self.noneCheck('bold')

    def noneCheck(self, attribute: str) -> boolean | None:
        style: boolean | None = getattr(self.object.font, attribute)
        baseLine = None

        while (style == None and baseLine != None):
            baseLine: BaseStyle | None = self.object.base_style
            if hasattr(baseLine, 'font'):
                # print('tk checker if if ', baseLine)
                currentFont: Font = getattr(baseLine, 'font')
                # print('tk checker if currentfonet ', currentFont)

                style = getattr(currentFont, attribute)
                # print(style)

        return style
