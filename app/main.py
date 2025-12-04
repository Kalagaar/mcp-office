from fastmcp import FastMCP
from pydantic import BaseModel, Field
from typing import Annotated, List, Dict, Optional, Literal
from pathlib import Path
import logging

from .config import get_config
from .tools.excel import markdown_to_excel
from .tools.word import markdown_to_word
from .tools.word import manipulation as word_ops
from .tools.pptx import create_presentation
from .tools.email import create_eml
from .tools.email.dynamic_email_tools import register_email_template_tools_from_yaml
mcp = FastMCP("MCP Office Documents")

# Initialize config and logging
config = get_config()
logger = logging.getLogger(__name__)

# Look for dynamic email templates in production and local locations.
# Production (container): /app/config/email_templates.yaml
# Local development: <project_root>/config/email_templates.yaml
APP_CONFIG_PATH = Path("/app/config") / "email_templates.yaml"
LOCAL_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "email_templates.yaml"

# Prefer the production path when present, otherwise fall back to local config.
_primary_yaml = None
for candidate in (APP_CONFIG_PATH, LOCAL_CONFIG_PATH):
    if candidate.exists():
        _primary_yaml = candidate
        logger.info("[dynamic-email] Found email templates file: %s", candidate)
        break

if _primary_yaml:
    try:
        register_email_template_tools_from_yaml(mcp, _primary_yaml)
    except Exception as e:
        logger.exception("[dynamic-email] Failed to register email templates from %s: %s", _primary_yaml, e)
else:
    logger.info(
        "[dynamic-email] No dynamic email templates file found at /app/config/email_templates.yaml or config/email_templates.yaml - skipping"
    )

class PowerPointSlide(BaseModel):
    """PowerPoint slide - can be title, section, or content slide based on slide_type."""
    slide_type: Literal["title", "section", "content"] = Field(description="Type of slide: 'title' for presentation opening, 'section' for dividers, 'content' for slide with bullet points")
    slide_title: str = Field(description="Title text for the slide")

    # Optional fields based on slide type
    author: Optional[str] = Field(default="", description="Author name for title slides - appears in subtitle placeholder. Leave empty for section/content slides.")
    slide_text: Optional[List[Dict]] = Field(
        default=None,
        description="Array of bullet points for content slides. Each bullet point must have 'text' (string) and 'indentation_level' (integer 1-5). Leave empty/null for title and section slides."
    )

@mcp.tool(
    name="create_excel_from_markdown",
    description="Converts markdown content with tables and formulas to Excel (.xlsx) format.",
    tags={"excel", "spreadsheet", "data"},
    annotations={"title": "Markdown to Excel Converter"}
)
async def create_excel_document(
    markdown_content: Annotated[str, Field(description="Markdown content containing tables, headers, and formulas. Use T1.B[0] for cross-table references and B[0] for current row references. ALWAYS use [0], [1], [2] notation, NEVER use absolute row numbers like B2, B3. Do NOT count table header as first row, first row has index [0]. Supports cell formatting: **bold**, *italic*.")]
) -> str:
    """
    Converts markdown to Excel with advanced formula support.
    """

    logger.info("Converting markdown to Excel document")

    try:
        result = markdown_to_excel(markdown_content)
        logger.info("Excel document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Excel document: {e}")
        return f"Error creating Excel document: {str(e)}"

@mcp.tool(
    name="create_word_from_markdown",
    description="Converts markdown content to Word (.docx) format. Supports headers, tables, lists, formatting, hyperlinks, and block quotes.",
    tags={"word", "document", "text", "legal", "contract"},
    annotations={"title": "Markdown to Word Converter"}
)
async def create_word_document(
    markdown_content: Annotated[str, Field(description="Markdown content. For LEGAL CONTRACTS use numbered lists (1., 2., 3.) for sections and nested lists for provisions - DO NOT use headers (except for contract title). For other documents use headers (# ## ###).")]
) -> str:
    """
    Converts markdown to professionally formatted Word document.

    """

    logger.info("Converting markdown to Word document")

    try:
        result = markdown_to_word(markdown_content)
        logger.info("Word document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Word document: {e}")
        return f"Error creating Word document: {str(e)}"


def register_word_manipulation_tools():
    """Register advanced Word manipulation tools from word_ops module."""

    @mcp.tool(name="word_create_document", description="Create a new Word document with optional metadata.")
    async def word_create_document(filename: str, title: Optional[str] = None, author: Optional[str] = None) -> str:
        return await word_ops.create_document(filename, title, author)

    @mcp.tool(name="word_list_documents", description="List Word documents in a directory.")
    async def word_list_documents(directory: str = ".") -> str:
        return await word_ops.list_available_documents(directory)

    @mcp.tool(name="word_get_info", description="Get metadata information from a Word document.")
    async def word_get_info(filename: str) -> str:
        return await word_ops.get_document_info(filename)

    @mcp.tool(name="word_get_outline", description="Get paragraph and table outline of a Word document.")
    async def word_get_outline(filename: str) -> str:
        return await word_ops.get_document_outline(filename)

    @mcp.tool(name="word_get_text", description="Extract all text from a Word document.")
    async def word_get_text(filename: str) -> str:
        return await word_ops.get_document_text(filename)

    @mcp.tool(name="word_copy_document", description="Copy a Word document.")
    async def word_copy_document(source_filename: str, destination_filename: Optional[str] = None) -> str:
        return await word_ops.copy_document(source_filename, destination_filename)

    @mcp.tool(name="word_merge_documents", description="Merge multiple Word documents into one.")
    async def word_merge_documents(target_filename: str, source_filenames: List[str], add_page_breaks: bool = True) -> str:
        return await word_ops.merge_documents(target_filename, source_filenames, add_page_breaks)

    @mcp.tool(name="word_add_paragraph", description="Add a paragraph with optional styling.")
    async def word_add_paragraph(filename: str, text: str, style: Optional[str] = None,
                                 font_name: Optional[str] = None, font_size: Optional[int] = None,
                                 bold: Optional[bool] = None, italic: Optional[bool] = None, color: Optional[str] = None) -> str:
        return await word_ops.add_paragraph(filename, text, style, font_name, font_size, bold, italic, color)

    @mcp.tool(name="word_add_heading", description="Add a heading to a document.")
    async def word_add_heading(filename: str, text: str, level: int = 1,
                               font_name: Optional[str] = None, font_size: Optional[int] = None,
                               bold: Optional[bool] = None, italic: Optional[bool] = None,
                               border_bottom: bool = False) -> str:
        return await word_ops.add_heading(filename, text, level, font_name, font_size, bold, italic, border_bottom)

    @mcp.tool(name="word_add_table", description="Add a table to a document.")
    async def word_add_table(filename: str, rows: int, cols: int, data: Optional[List[List[str]]] = None) -> str:
        return await word_ops.add_table(filename, rows, cols, data)

    @mcp.tool(name="word_search_replace", description="Search and replace text across paragraphs and tables.")
    async def word_search_replace(filename: str, find_text: str, replace_text: str) -> str:
        return await word_ops.search_and_replace(filename, find_text, replace_text)

    @mcp.tool(name="word_insert_header_near_text", description="Insert a header before/after target text or paragraph index.")
    async def word_insert_header_near_text(filename: str, target_text: Optional[str] = None, header_title: str = "",
                                           position: str = 'after', header_style: str = 'Heading 1',
                                           target_paragraph_index: Optional[int] = None) -> str:
        return await word_ops.insert_header_near_text_tool(filename, target_text, header_title, position, header_style, target_paragraph_index)

    @mcp.tool(name="word_insert_paragraph_near_text", description="Insert paragraph near target text or index.")
    async def word_insert_paragraph_near_text(filename: str, target_text: Optional[str] = None, line_text: str = "",
                                              position: str = 'after', line_style: Optional[str] = None,
                                              target_paragraph_index: Optional[int] = None) -> str:
        return await word_ops.insert_line_or_paragraph_near_text_tool(filename, target_text, line_text, position, line_style, target_paragraph_index)

    @mcp.tool(name="word_insert_list_near_text", description="Insert bulleted/numbered list near target text or index.")
    async def word_insert_list_near_text(filename: str, target_text: Optional[str] = None, list_items: Optional[List[str]] = None,
                                         position: str = 'after', target_paragraph_index: Optional[int] = None,
                                         bullet_type: str = 'bullet') -> str:
        return await word_ops.insert_numbered_list_near_text_tool(filename, target_text, list_items, position, target_paragraph_index, bullet_type)

    @mcp.tool(name="word_format_text", description="Format a specific range in a paragraph.")
    async def word_format_text(filename: str, paragraph_index: int, start_pos: int, end_pos: int,
                               bold: Optional[bool] = None, italic: Optional[bool] = None,
                               underline: Optional[bool] = None, color: Optional[str] = None,
                               font_size: Optional[int] = None, font_name: Optional[str] = None) -> str:
        return await word_ops.format_text(filename, paragraph_index, start_pos, end_pos, bold, italic, underline, color, font_size, font_name)

    @mcp.tool(name="word_create_custom_style", description="Create a custom text style in the document.")
    async def word_create_custom_style(filename: str, style_name: str, bold: Optional[bool] = None,
                                       italic: Optional[bool] = None, font_size: Optional[int] = None,
                                       font_name: Optional[str] = None, color: Optional[str] = None,
                                       base_style: Optional[str] = None) -> str:
        return await word_ops.create_custom_style(filename, style_name, bold, italic, font_size, font_name, color, base_style)

    @mcp.tool(name="word_protect_document", description="Protect a Word document with password or restrictions.")
    async def word_protect_document(filename: str, password: str) -> str:
        return await word_ops.protect_document(filename, password)

    @mcp.tool(name="word_unprotect_document", description="Remove password protection from a Word document.")
    async def word_unprotect_document(filename: str, password: str) -> str:
        return await word_ops.unprotect_document(filename, password)

    @mcp.tool(name="word_add_footnote", description="Add a footnote to a specific paragraph.")
    async def word_add_footnote(filename: str, paragraph_index: int, footnote_text: str) -> str:
        return await word_ops.add_footnote_to_document(filename, paragraph_index, footnote_text)

    @mcp.tool(name="word_add_footnote_after_text", description="Add a footnote after a specific text.")
    async def word_add_footnote_after_text(filename: str, search_text: str, footnote_text: str,
                                           output_filename: Optional[str] = None) -> str:
        return await word_ops.add_footnote_after_text(filename, search_text, footnote_text, output_filename)

    @mcp.tool(name="word_add_footnote_before_text", description="Add a footnote before a specific text.")
    async def word_add_footnote_before_text(filename: str, search_text: str, footnote_text: str,
                                            output_filename: Optional[str] = None) -> str:
        return await word_ops.add_footnote_before_text(filename, search_text, footnote_text, output_filename)

    @mcp.tool(name="word_add_comment", description="Retrieve all comments from the document.")
    async def word_get_all_comments(filename: str) -> str:
        return await word_ops.get_all_comments(filename)

    @mcp.tool(name="word_get_comments_by_author", description="Retrieve comments filtered by author.")
    async def word_get_comments_by_author(filename: str, author: str) -> str:
        return await word_ops.get_comments_by_author(filename, author)

    @mcp.tool(name="word_get_comments_for_paragraph", description="Retrieve comments for a paragraph index.")
    async def word_get_comments_for_paragraph(filename: str, paragraph_index: int) -> str:
        return await word_ops.get_comments_for_paragraph(filename, paragraph_index)

    @mcp.tool(name="word_find_text", description="Find occurrences of text with options.")
    async def word_find_text(filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False) -> str:
        return await word_ops.find_text_in_document(filename, text_to_find, match_case, whole_word)

    @mcp.tool(name="word_convert_to_pdf", description="Convert Word document to PDF.")
    async def word_convert_to_pdf(filename: str, output_filename: Optional[str] = None) -> str:
        return await word_ops.convert_to_pdf(filename, output_filename)


register_word_manipulation_tools()

@mcp.tool(
    name="create_powerpoint_presentation",
    description="Creates PowerPoint presentations with professional templates using structured slide models.",
    tags={"powerpoint", "presentation", "slides"},
    annotations={"title": "PowerPoint Presentation Creator"}
)
async def create_powerpoint_presentation(
    slides: List[PowerPointSlide],
    format: Annotated[Literal["4:3", "16:9"], Field(
        default="4:3",
        description="Presentation formating: '4:3' for traditional or '16:9' for widescreen"
    )]
) -> str:
    """Creates PowerPoint presentations with structured slide models and professional templates."""

    logger.info(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        slides_data = [slide.model_dump() for slide in slides]
        result = create_presentation(slides_data, format)
        logger.info(f"PowerPoint presentation created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating PowerPoint presentation: {e}")
        return f"Error creating PowerPoint presentation: {str(e)}"

@mcp.tool(
    name="create_email_draft",
    description="Creates an email draft in EML format with HTML content using preset professional styling.",
    tags={"email", "eml", "communication"},
    annotations={"title": "Email Draft Creator"}
)
async def create_email_draft(
    content: Annotated[str, Field(description="BODY CONTENT ONLY - Do NOT include HTML structure tags like <html>, <head>, <body>, or <style>. Do NOT include any CSS styling. Use <p> for greetings and for signatures, never headers. Use <h2> for section headers (will be bold), <h3> for subsection headers (will be underlined). HTML tags allowed: <p>, <h2>, <h3>, <ul>, <li>, <strong>, <em>, <div>.")],
    subject: Annotated[str, Field(description="Email subject line")],
    to: Annotated[Optional[List[str]], Field(description="List of recipient email addresses", default=None)],
    cc: Annotated[Optional[List[str]], Field(description="List of CC recipient email addresses", default=None)],
    bcc: Annotated[Optional[List[str]], Field(description="List of BCC recipient email addresses", default=None)],
    priority: Annotated[str, Field(description="Email priority: 'low', 'normal', or 'high'", default="normal")],
    language: Annotated[str, Field(description="Language code for proofreading in Outlook (e.g., 'cs-CZ' for Czech, 'en-US' for English, 'de-DE' for German, 'sk-SK' for Slovak)", default="cs-CZ")]
) -> str:
    """
    Creates professional email drafts in EML format with preset styling and language settings.
    """

    logger.info(f"Creating email draft with subject: {subject}")

    try:
        result = create_eml(
            to=to,
            cc=cc,
            bcc=bcc,
            re=subject,
            content=content,
            priority=priority,
            language=language
        )
        logger.info(f"Email draft created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating email draft: {e}")
        return f"Error creating email draft: {str(e)}"

if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8958,
        log_level=config.logging.mcp_level_str,
        path="/mcp"
    )
