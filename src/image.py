from __future__ import annotations
from typing import TYPE_CHECKING
import io

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from loguru import logger
from PIL import Image, ImageDraw

if TYPE_CHECKING:
    from docx.document import Document as DocumentType
    from docx.parts.document import DocumentPart
    from docx.shape import InlineShape, InlineShapes


def main(file_path: str):
    """
    Read image from docx and write image as a file.

    NOTE Refs
        - [python-docx.picture](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/shapes/picture.html)
    """
    try:
        doc: DocumentType = Document(file_path)
    except PackageNotFoundError as err:
        logger.error(f"Failed read file from {file_path}")
        raise err
    else:
        logger.info(f"Succeed read file from {file_path}")

    part: DocumentPart = doc.part

    for idx, image in enumerate(part.package._image_parts):
        logger.debug(
            f"{image} image.blob {type(image.image.blob)} {image.content_type}"
        )
        # filename = image.filename
        """Read file meta info and convert to byteIO type"""
        content_type = image.content_type.split("/")[1]
        r_data = io.BytesIO(image.blob)

        """Create image instance"""
        img = Image.open(r_data)
        ImageDraw.Draw(img)
        """Write image file to `export`"""
        img.save(f"./export/image-{idx}.{content_type}")

    for shape in part.inline_shapes:
        shape: InlineShape
        shape.height = int(shape.height / 2)

        """
        NOTE picture is 3
            this values are referenced http://msdn.microsoft.com/en-us/library/office/ff192587.aspx
        """
        # logger.debug(f"image type is {shape.type}")
        # logger.debug(shape._inline)

    """export file"""
    result_file_path = "./export/text-with-image.docx"
    try:
        doc.save(result_file_path)
    except PackageNotFoundError as err:
        logger.error(f"Failed write file to {result_file_path}")
        raise err
    else:
        logger.info(f"Succeed write file to {result_file_path}")


if __name__ == "__main__":
    file_path = "./samples/text-with-image.docx"
    main(file_path)
