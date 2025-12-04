"""Formatting tools for Word documents."""

import os
from typing import List, Optional
from docx import Document
from docx.shared import Pt, RGBColor

from app.tools.word.utils import check_file_writeable, ensure_docx_extension
from app.tools.word.core import (
    create_style,
    set_cell_border,
    apply_table_style,
    copy_table,
)
