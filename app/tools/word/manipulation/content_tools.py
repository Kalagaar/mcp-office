"""Content tools for Word manipulation."""

import os
from typing import List, Optional
from docx import Document
from docx.shared import Inches, Pt, RGBColor

from app.tools.word.utils import (
    check_file_writeable,
    ensure_docx_extension,
    find_and_replace_text,
    insert_header_near_text,
    insert_numbered_list_near_text,
    insert_line_or_paragraph_near_text,
    replace_paragraph_block_below_header,
    replace_block_between_manual_anchors,
)
from app.tools.word.core import ensure_heading_style, ensure_table_style
