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
    """
    2022-04-15 16:28:19.059 | INFO     | __main__:main:19 - Finished read file from ./samples/python-docx.docx
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:23 - Hello This is test doc
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:26 - Hello This is test doc None
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:23 - 
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:26 -  None
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:23 - x1 + y2 = z3
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:26 - x None
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:26 - 1 True
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:26 -  + y None
    2022-04-15 16:28:19.059 | DEBUG    | __main__:main:26 - 2 False
    2022-04-15 16:28:19.060 | DEBUG    | __main__:main:26 -  = z None
    2022-04-15 16:28:19.060 | DEBUG    | __main__:main:26 - 3 None
    """
    main(file_path)
