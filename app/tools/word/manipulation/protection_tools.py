"""Protection tools for Word."""

import os
import hashlib
import datetime
import io
from typing import List, Optional
from docx import Document
import msoffcrypto

from app.tools.word.utils import check_file_writeable, ensure_docx_extension
from app.tools.word.core.protection import (
    add_protection_info,
    verify_document_protection,
    create_signature_info,
)
