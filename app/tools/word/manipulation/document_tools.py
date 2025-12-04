"""Document tools for Word."""

import os
import json
from typing import Dict, List, Optional
from docx import Document

from app.tools.word.utils import (
    check_file_writeable,
    ensure_docx_extension,
    create_document_copy,
    get_document_properties,
    extract_document_text,
    get_document_structure,
    get_document_xml,
    insert_header_near_text,
    insert_line_or_paragraph_near_text,
)
from app.tools.word.core import ensure_heading_style, ensure_table_style, copy_table
