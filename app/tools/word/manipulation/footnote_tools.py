"""Footnote tools for Word documents."""

import os
from typing import Optional, Dict, Any
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE

from app.tools.word.utils import check_file_writeable, ensure_docx_extension
from app.tools.word.core.footnotes import (
    add_footnote,
    add_endnote,
    convert_footnotes_to_endnotes,
)
