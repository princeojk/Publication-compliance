from docx.document import Document as DocumentClass
import pytest
from docx.text.run import Run
from typing import Literal, cast
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Length
from docx.styles.style import CharacterStyle, ParagraphStyle
from src.utils.expected.expectedTitleAttrs import ExpectedTitle
from src.errors import ComplianceChecker

StyleType = bool | float | None | str | WD_ALIGN_PARAGRAPH
CustomAttr = Literal['bold', 'size', 'name', 'alignment', 'space_after', 'space_before']
StyleAttr = Literal['font', 'paragraph_format']

class TestTitle:
    
    @pytest.fixture(autouse=True)
    def setup_method(self, loaded_document: DocumentClass):
        self.title = loaded_document.paragraphs[0]

    def test_isBold_is_not_true(self):
        isBold = False
        for run in self.title.runs:
            if run.bold is True:
                isBold = True
                break
        assert isBold == False

    def test_font_size_is_valid(self):
        size = None
        titleStyle = None
        sizeInPoints = 0.0
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

        assert sizeInPoints == ExpectedTitle.EXPECTED_FONT_SIZE

    def test_alignment_is_valid(self):
        if self.title.alignment:
            return self.title.alignment

        style: ParagraphStyle | None  = self.title.style
        value = None
        if style is None:
            return ComplianceChecker('alignment', 'center', 'Style Not found')
        
        value = getattr(style.paragraph_format, "alignment", None)

        if not isinstance(value, WD_ALIGN_PARAGRAPH):
            value = self.getFromBaseStyle(style, 'paragraph_format', 'alignment')

            assert value == ExpectedTitle.EXPECTED_ALIGNMENT

    
    def getFromRunStyle(self, run: Run, attribute: CustomAttr) -> StyleType:
        style: CharacterStyle = run.style
        attr =  getattr(style.font, attribute)
        if attr is None:
            attr = self.getFromBaseStyle(style, 'font', attribute)

        return attr

    def getFromBaseStyle(self, style: CharacterStyle | ParagraphStyle, styleAttribute: StyleAttr, attribute: CustomAttr) -> StyleType:
        baseStyle = style.base_style
        while baseStyle is not None:

            baseStyle = cast(ParagraphStyle, baseStyle)
            styleAttr = getattr(baseStyle, styleAttribute)

            attr =  getattr(styleAttr, attribute)
            if attr is not None:
                return attr
            baseStyle =  baseStyle.base_style

        return None
