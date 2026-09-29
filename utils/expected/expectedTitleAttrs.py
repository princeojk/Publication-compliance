from utils.fontEnum import FontNames
from docx.enum.text import WD_ALIGN_PARAGRAPH

class ExpectedTitle:
    EXPECTED_FONT_SIZE = 20.0
    EXPECTED_FONT_NAME = FontNames.Roman.value
    EXPECTED_FONT_ALIGNMENT = WD_ALIGN_PARAGRAPH.CENTER
    EXPECTED_BOLD = False

    class EXPECTED_SPACE:
        SPACE_BEFORE = 24.0
        SPACE_AFTER = 16.0

    def spaceSerilizer(self) -> dict[str,float]:

        return {
            'before': self.EXPECTED_SPACE.SPACE_BEFORE,
            'after' : self.EXPECTED_SPACE.SPACE_AFTER
        }