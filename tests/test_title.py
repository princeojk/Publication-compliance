from docx.document import Document as DocumentClass
import pytest
from docx.text.run import Run
from typing import Literal, cast
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.styles.style import CharacterStyle, ParagraphStyle
from src.utils.expected.expectedTitleAttrs import ExpectedTitle
from src.sections.mainTitle import MainTitle

StyleType = bool | float | None | str | WD_ALIGN_PARAGRAPH
CustomAttr = Literal['bold', 'size', 'name', 'alignment', 'space_after', 'space_before']
StyleAttr = Literal['font', 'paragraph_format']

# add test for fails later!
class TestTitle:
    @pytest.fixture(autouse=True)
    def setup_method(self, loaded_document: DocumentClass):
        self.title = MainTitle(loaded_document.paragraphs[0])

    def test_isBold_is_not_true(self):
        assert self.title.isBold() == False

    def test_font_size_is_valid(self):
        assert self.title.getFontSize() == ExpectedTitle.EXPECTED_FONT_SIZE

    def test_alignment_is_valid(self):
        assert self.title.getAlignment() == ExpectedTitle.EXPECTED_ALIGNMENT

    
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
    
