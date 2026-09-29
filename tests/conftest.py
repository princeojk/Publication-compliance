from pathlib import Path
import pytest
from docx import Document

@pytest.fixture(scope="session")
def loaded_document():
    currentDirectory = Path(__file__).resolve().parent.parent

    filePath = currentDirectory / 'samples/or_submission_2_4_1.docx'
    return Document(str(filePath))
