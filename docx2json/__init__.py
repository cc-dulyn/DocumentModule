from .docx2json import DocumentStructureExtractor, extract_document
from .docx2json_img import DocumentStructureExtractor as ImageDocumentStructureExtractor

__all__ = [
    "DocumentStructureExtractor",
    "ImageDocumentStructureExtractor",
    "extract_document"
]