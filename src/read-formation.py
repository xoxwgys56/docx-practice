from __future__ import annotations
from typing import TYPE_CHECKING
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml import OxmlElement
from loguru import logger

if TYPE_CHECKING:
    from docx.document import Document as DocumentType


def main(file_path: str):
    # read file
    try:
        doc: DocumentType = Document(file_path)
    except PackageNotFoundError as err:
        raise err(f"Failed read file from {file_path}")
    else:
        logger.info(f"Finished read file from {file_path}")

    # logging read paragraphs
    for paragraph in doc.paragraphs:
        logger.debug(paragraph.text)

        for i in paragraph.runs:
            logger.debug(f"{i.text} {i.font.subscript}")
            i.text = "1"

    doc.save("./export/python-docx.docx")


if __name__ == "__main__":
    file_path = "./samples/python-docx.docx"
    main(file_path)
