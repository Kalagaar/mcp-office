"""Core functionality for Word tools."""

from .styles import ensure_heading_style, ensure_table_style, create_style
from .protection import add_protection_info, verify_document_protection, is_section_editable, create_signature_info, verify_signature
from .footnotes import add_footnote, add_endnote, convert_footnotes_to_endnotes, find_footnote_references, get_format_symbols, customize_footnote_formatting
from .tables import set_cell_border, apply_table_style, copy_table

__all__ = [name for name in globals().keys() if not name.startswith("_")]
