import math
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docx.styles.style import CharacterStyle, ParagraphStyle
from docx.shared import Length
from docx.parts.document import DocumentPart
from typing import Literal, cast
from docx.enum.text import WD_ALIGN_PARAGRAPH
from src.errors import ComplianceChecker
from utils.spaceUtils import SpaceAttributes
from utils.expected.expectedTitleAttrs import ExpectedTitle
from utils.fontEnum import FontNames

StyleType = bool | float | None | str | WD_ALIGN_PARAGRAPH
CustomAttr = Literal['bold', 'size', 'name', 'alignment', 'space_after', 'space_before']
StyleAttr = Literal['font', 'paragraph_format']
    
class MainTitle:
    DEFAULT_FONT_SIZE = 10.0

    def __init__(self, title: Paragraph) -> None:
        self.title = title

    def isBold(self) -> bool:
        bold = None
        for run in self.title.runs:
            if run.bold is not None:
                if run.bold is False:
                    continue
                else:
                    return True
            else:
                bold =  self.getFromRunStyle(run, 'bold')

            if bold is True:
                return True

        if bold is None:
            if self.title.style is not None:
                bold =  self.title.style.font.bold

        if isinstance(bold, bool):
            return bold
        
        return False

    def getFontSize(self) -> float:
        size = None
        titleStyle = None
        sizeInPoints = self.DEFAULT_FONT_SIZE
        for run in self.title.runs:
            size = run.font.size
            if size is None:
                size = self.getFromRunStyle(run, 'size')
                if isinstance(size, Length):
                    sizeInPoints = size.pt
                
        if size is None:
            titleStyle = self.title.style

        if titleStyle is not None:
            size = titleStyle.font.size
            if isinstance(size, Length):
                sizeInPoints = size.pt

        return sizeInPoints

    def getFontName(self) -> str | None:
        name = None
        titleStyle = None
        for run in self.title.runs:
            name = run.font.name
            if name is None:
                name = self.getFromRunStyle(run, 'name')
                
        if name is None:
            titleStyle = self.title.style

        if titleStyle is not None:
            name = titleStyle.font.name

        if name is None:
            name = self.getFromParaStyle('font', 'name')

        if name is None:
            name = self.getFromDefaults()

        if isinstance(name, str):
            return name

    # set. default later
    def getAlignment(self) -> WD_ALIGN_PARAGRAPH | ComplianceChecker:
        if self.title.alignment:
            return self.title.alignment

        style: ParagraphStyle | None  = self.title.style
        value = None
        if style is None:
            return ComplianceChecker('alignment', 'center', 'Style Not found')
        
        value = getattr(style.paragraph_format, "alignment", None)

        if isinstance(value, WD_ALIGN_PARAGRAPH):
            return value

        value = self.getFromBaseStyle(style, 'paragraph_format', 'alignment')
        if isinstance(value, WD_ALIGN_PARAGRAPH):
            return value

        # wrong get the default. value later!!!!
        return ComplianceChecker('alignment', 'center', 'value Not found')

    # set. default later
    def getSpace(self) -> SpaceAttributes | None | ComplianceChecker:
        after = None
        before = None

        style: ParagraphStyle | None  = self.title.style

        if style is None:
            return ComplianceChecker('space', 'center', 'Style Not found')
        
        if self.title.paragraph_format:
            after =  getattr(self.title.paragraph_format, "space_after", None)
            before = getattr(self.title.paragraph_format, "space_before", None)

        if after is None or before is None:
            after = getattr(style.paragraph_format, "space_after")
            before = getattr(style.paragraph_format, "space_before")

        if after is None or before is None:
            after = self.getFromBaseStyle(style, 'paragraph_format', "space_after")
            before = self.getFromBaseStyle(style, 'paragraph_format', "space_before")

        if isinstance(after, Length) and isinstance(before, Length):
            return SpaceAttributes(before.pt, after.pt)

    def getFromRunStyle(self, run: Run, attribute: CustomAttr) -> StyleType:
        style: CharacterStyle = run.style
        attr =  getattr(style.font, attribute)
        if attr is None:
            attr = self.getFromBaseStyle(style, 'font', attribute)

        return attr

    def getFromParaStyle(self, styleAttribute: StyleAttr, attribute: CustomAttr) -> StyleType:
        style: ParagraphStyle | None = self.title.style
        attr = None

        if style is not None:
                styleAttr = getattr(style, styleAttribute)
                attr =  getattr(styleAttr, attribute)

        if attr is None and style is not None:
            attr = self.getFromBaseStyle(style, styleAttribute, attribute)

        return attr
        
    def getFromBaseStyle(self, style: CharacterStyle | ParagraphStyle, styleAttribute: StyleAttr, attribute: CustomAttr) -> StyleType:
        baseStyle = style.base_style
        print('tk base', baseStyle)
        while baseStyle is not None:

            baseStyle = cast(ParagraphStyle, baseStyle)
            styleAttr = getattr(baseStyle, styleAttribute)
            print('tk base two', styleAttr, styleAttribute)

            attr =  getattr(styleAttr, attribute)
            print('tk base iner', attr)
            if attr is not None:
                return attr
            baseStyle =  baseStyle.base_style
            print('tk basestyle inner', attr)

        return None

    def getFromDefaults(self) -> StyleType:
        part = self.title.part
        docPart = cast(DocumentPart, part)

        names = docPart.styles.element.xpath(
            "w:docDefaults/w:rPrDefault/w:rPr/w:rFonts/@w:ascii"
        )
        name = names[0] if names else None
        return name

    def serializer(self) -> dict[str, dict[str, str|bool|float|None|dict[str, float]]]:
        result: dict[str, dict[str, str|bool|float|None|dict[str, float]]] = {}

        if not math.isclose(self.getFontSize(), ExpectedTitle.EXPECTED_FONT_SIZE):
            checker = ComplianceChecker('FontSize failed', ExpectedTitle.EXPECTED_FONT_SIZE, self.getFontSize())
            result['fontSize'] = checker.serilizer()

        if self.isBold():
            checker = ComplianceChecker('Title should Not Bold', ExpectedTitle.EXPECTED_BOLD, self.isBold())
            result['bold'] = checker.serilizer()  

        if self.getFontName() != FontNames.Roman.value:
            checker = ComplianceChecker('Incorrect Font Name', ExpectedTitle.EXPECTED_FONT_NAME, self.getFontName())
            result['fontName'] = checker.serilizer() 

        value = self.getAlignment()
        if value != WD_ALIGN_PARAGRAPH.CENTER:
            if isinstance(value, WD_ALIGN_PARAGRAPH):
                checker = ComplianceChecker('Incorrect ALignment', ExpectedTitle.EXPECTED_ALIGNMENT, value.name.lower())
            else:
                checker = value
            result['alignmnet'] = checker.serilizer() 

        space = self.getSpace()

        if isinstance(space, SpaceAttributes):
            if not math.isclose(space.before, ExpectedTitle.EXPECTED_SPACE.SPACE_BEFORE) or not math.isclose(space.after, ExpectedTitle.EXPECTED_SPACE.SPACE_AFTER):
                expected = ExpectedTitle()
                checker = ComplianceChecker('Incorrect ALignment', expected.spaceSerilizer(), space.serializer())
                result['space'] = checker.serilizer() 

        return result