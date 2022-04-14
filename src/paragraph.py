from __future__ import annotations
from typing import TYPE_CHECKING
from docx import Document
from loguru import logger

if TYPE_CHECKING:
    from docx.document import Document as DocumentType

if __name__ == "__main__":
    doc: DocumentType = Document("./samples/python-docx.docx")

    """
    can define run.font set as subscript or superscript.
    """
    for paragraph in doc.paragraphs:
        logger.debug(paragraph.text)
        for run in paragraph.runs:
            logger.debug(f"{run.text} {run.font.subscript} {run.font.superscript}")
