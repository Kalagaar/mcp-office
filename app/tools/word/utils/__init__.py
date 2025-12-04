"""Utility functions for Word tools."""

from .file_utils import check_file_writeable, create_document_copy, ensure_docx_extension
from .document_utils import (
    get_document_properties,
    extract_document_text,
    get_document_structure,
    find_paragraph_by_text,
    find_and_replace_text,
    get_document_xml,
    insert_header_near_text,
    insert_numbered_list_near_text,
    insert_line_or_paragraph_near_text,
    replace_paragraph_block_below_header,
    replace_block_between_manual_anchors,
)
from .extended_document_utils import get_paragraph_text, find_text

__all__ = [name for name in globals().keys() if not name.startswith("_")]
