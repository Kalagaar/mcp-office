from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Annotated, Dict, List, Literal, Optional

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, Field

from app.config import get_config
from app.tools.email import create_eml
from app.tools.email.dynamic_email_tools import register_email_template_tools_from_yaml
from app.tools.excel import markdown_to_excel
from app.tools.pptx import create_presentation
from app.tools.word import markdown_to_word
from app.tools.word import manipulation as word_ops

logger = logging.getLogger(__name__)
config = get_config()


class PowerPointSlide(BaseModel):
    slide_type: Literal["title", "section", "content"] = Field(
        description="'title' for opening, 'section' for dividers, 'content' for bullet slides"
    )
    slide_title: str = Field(description="Title text for the slide")
    author: Optional[str] = Field(default="", description="Author shown on title slides")
    slide_text: Optional[List[Dict]] = Field(
        default=None,
        description="Bullet array for content slides. Each entry must provide 'text' and 'indentation_level'",
    )


def build_mcp() -> FastMCP:
    app = FastMCP("MCP Office Documents")
    register_word_tools(app)
    register_excel_tools(app)
    register_powerpoint_tools(app)
    register_email_tools(app)
    register_dynamic_email_templates(app)
    return app


def register_excel_tools(app: FastMCP) -> None:
    @app.tool(
        name="create_excel_from_markdown",
        description="Convert markdown tables/formulas to Excel (.xlsx)",
    )
    async def create_excel_document(
        markdown_content: Annotated[
            str,
            Field(description="Markdown containing tables (use B[0] style references only)")
        ]
    ) -> str:
        logger.info("Converting markdown to Excel")
        try:
            return markdown_to_excel(markdown_content)
        except Exception as exc:  # pragma: no cover
            logger.exception("Excel conversion failed")
            return f"Error creating Excel document: {exc}"


def register_word_tools(app: FastMCP) -> None:
    @app.tool(
        name="create_word_from_markdown",
        description="Convert markdown content to Word (.docx)",
    )
    async def create_word_document(
        markdown_content: Annotated[
            str,
            Field(description="Markdown content. Use numbered lists for contracts; headers elsewhere."),
        ]
    ) -> str:
        logger.info("Converting markdown to Word document")
        try:
            return markdown_to_word(markdown_content)
        except Exception as exc:  # pragma: no cover
            logger.exception("Word conversion failed")
            return f"Error creating Word document: {exc}"

    register_word_manipulation_tools(app)


def register_word_manipulation_tools(app: FastMCP) -> None:
    @app.tool(name="word_create_document", description="Create a new Word document with optional metadata.")
    async def word_create_document(filename: str, title: Optional[str] = None, author: Optional[str] = None) -> str:
        return await word_ops.create_document(filename, title, author)

    @app.tool(name="word_list_documents", description="List Word documents in a directory.")
    async def word_list_documents(directory: str = ".") -> str:
        return await word_ops.list_available_documents(directory)

    @app.tool(name="word_get_info", description="Get metadata information from a Word document.")
    async def word_get_info(filename: str) -> str:
        return await word_ops.get_document_info(filename)

    @app.tool(name="word_get_outline", description="Get paragraph and table outline of a Word document.")
    async def word_get_outline(filename: str) -> str:
        return await word_ops.get_document_outline(filename)

    @app.tool(name="word_get_text", description="Extract all text from a Word document.")
    async def word_get_text(filename: str) -> str:
        return await word_ops.get_document_text(filename)

    @app.tool(name="word_copy_document", description="Copy a Word document.")
    async def word_copy_document(source_filename: str, destination_filename: Optional[str] = None) -> str:
        return await word_ops.copy_document(source_filename, destination_filename)

    @app.tool(name="word_merge_documents", description="Merge multiple Word documents into one.")
    async def word_merge_documents(target_filename: str, source_filenames: List[str], add_page_breaks: bool = True) -> str:
        return await word_ops.merge_documents(target_filename, source_filenames, add_page_breaks)

    @app.tool(name="word_add_paragraph", description="Add a paragraph with optional styling.")
    async def word_add_paragraph(
        filename: str,
        text: str,
        style: Optional[str] = None,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None,
    ) -> str:
        return await word_ops.add_paragraph(filename, text, style, font_name, font_size, bold, italic, color)

    @app.tool(name="word_add_heading", description="Add a heading to a document.")
    async def word_add_heading(
        filename: str,
        text: str,
        level: int = 1,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        border_bottom: bool = False,
    ) -> str:
        return await word_ops.add_heading(filename, text, level, font_name, font_size, bold, italic, border_bottom)

    @app.tool(name="word_add_table", description="Add a table to a document.")
    async def word_add_table(filename: str, rows: int, cols: int, data: Optional[List[List[str]]] = None) -> str:
        return await word_ops.add_table(filename, rows, cols, data)

    @app.tool(name="word_search_replace", description="Search and replace text across paragraphs and tables.")
    async def word_search_replace(filename: str, find_text: str, replace_text: str) -> str:
        return await word_ops.search_and_replace(filename, find_text, replace_text)

    @app.tool(name="word_insert_header_near_text", description="Insert a header before/after target text or paragraph index.")
    async def word_insert_header_near_text(
        filename: str,
        target_text: Optional[str] = None,
        header_title: str = "",
        position: str = "after",
        header_style: str = "Heading 1",
        target_paragraph_index: Optional[int] = None,
    ) -> str:
        return await word_ops.insert_header_near_text_tool(
            filename, target_text, header_title, position, header_style, target_paragraph_index
        )

    @app.tool(name="word_insert_paragraph_near_text", description="Insert paragraph near target text or index.")
    async def word_insert_paragraph_near_text(
        filename: str,
        target_text: Optional[str] = None,
        line_text: str = "",
        position: str = "after",
        line_style: Optional[str] = None,
        target_paragraph_index: Optional[int] = None,
    ) -> str:
        return await word_ops.insert_line_or_paragraph_near_text_tool(
            filename, target_text, line_text, position, line_style, target_paragraph_index
        )

    @app.tool(name="word_insert_list_near_text", description="Insert bulleted/numbered list near target text or index.")
    async def word_insert_list_near_text(
        filename: str,
        target_text: Optional[str] = None,
        list_items: Optional[List[str]] = None,
        position: str = "after",
        target_paragraph_index: Optional[int] = None,
        bullet_type: str = "bullet",
    ) -> str:
        return await word_ops.insert_numbered_list_near_text_tool(
            filename, target_text, list_items, position, target_paragraph_index, bullet_type
        )

    @app.tool(name="word_format_text", description="Format a specific range in a paragraph.")
    async def word_format_text(
        filename: str,
        paragraph_index: int,
        start_pos: int,
        end_pos: int,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        color: Optional[str] = None,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
    ) -> str:
        return await word_ops.format_text(
            filename, paragraph_index, start_pos, end_pos, bold, italic, underline, color, font_size, font_name
        )

    @app.tool(name="word_create_custom_style", description="Create a custom text style in the document.")
    async def word_create_custom_style(
        filename: str,
        style_name: str,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        color: Optional[str] = None,
        base_style: Optional[str] = None,
    ) -> str:
        return await word_ops.create_custom_style(filename, style_name, bold, italic, font_size, font_name, color, base_style)

    @app.tool(name="word_protect_document", description="Protect a Word document with password or restrictions.")
    async def word_protect_document(filename: str, password: str) -> str:
        return await word_ops.protect_document(filename, password)

    @app.tool(name="word_unprotect_document", description="Remove password protection from a Word document.")
    async def word_unprotect_document(filename: str, password: str) -> str:
        return await word_ops.unprotect_document(filename, password)

    @app.tool(name="word_add_footnote", description="Add a footnote to a specific paragraph.")
    async def word_add_footnote(filename: str, paragraph_index: int, footnote_text: str) -> str:
        return await word_ops.add_footnote_to_document(filename, paragraph_index, footnote_text)

    @app.tool(name="word_add_footnote_after_text", description="Add a footnote after a specific text.")
    async def word_add_footnote_after_text(
        filename: str, search_text: str, footnote_text: str, output_filename: Optional[str] = None
    ) -> str:
        return await word_ops.add_footnote_after_text(filename, search_text, footnote_text, output_filename)

    @app.tool(name="word_add_footnote_before_text", description="Add a footnote before a specific text.")
    async def word_add_footnote_before_text(
        filename: str, search_text: str, footnote_text: str, output_filename: Optional[str] = None
    ) -> str:
        return await word_ops.add_footnote_before_text(filename, search_text, footnote_text, output_filename)

    @app.tool(name="word_add_comment", description="Retrieve all comments from the document.")
    async def word_get_all_comments(filename: str) -> str:
        return await word_ops.get_all_comments(filename)

    @app.tool(name="word_get_comments_by_author", description="Retrieve comments filtered by author.")
    async def word_get_comments_by_author(filename: str, author: str) -> str:
        return await word_ops.get_comments_by_author(filename, author)

    @app.tool(name="word_get_comments_for_paragraph", description="Retrieve comments for a paragraph index.")
    async def word_get_comments_for_paragraph(filename: str, paragraph_index: int) -> str:
        return await word_ops.get_comments_for_paragraph(filename, paragraph_index)

    @app.tool(name="word_find_text", description="Find occurrences of text with options.")
    async def word_find_text(
        filename: str,
        text_to_find: str,
        match_case: bool = True,
        whole_word: bool = False,
    ) -> str:
        return await word_ops.find_text_in_document(filename, text_to_find, match_case, whole_word)

    @app.tool(name="word_convert_to_pdf", description="Convert Word document to PDF.")
    async def word_convert_to_pdf(filename: str, output_filename: Optional[str] = None) -> str:
        return await word_ops.convert_to_pdf(filename, output_filename)


def register_powerpoint_tools(app: FastMCP) -> None:
    @app.tool(
        name="create_powerpoint_presentation",
        description="Create PowerPoint presentations from structured slide input",
    )
    async def create_powerpoint_presentation(
        slides: List[PowerPointSlide],
        format: Annotated[
            Literal["4:3", "16:9"],
            Field(default="4:3", description="'4:3' traditional or '16:9' widescreen"),
        ],
    ) -> str:
        logger.info("Creating PowerPoint with %d slides (%s)", len(slides), format)
        try:
            payload = [slide.model_dump() for slide in slides]
            return create_presentation(payload, format)
        except Exception as exc:  # pragma: no cover
            logger.exception("PowerPoint creation failed")
            return f"Error creating PowerPoint presentation: {exc}"


def register_email_tools(app: FastMCP) -> None:
    @app.tool(
        name="create_email_draft",
        description="Create an EML draft with preset styling",
    )
    async def create_email_draft(
        content: Annotated[
            str,
            Field(description="BODY CONTENT ONLY – allowed tags: <p>, <h2>, <h3>, <ul>, <li>, <strong>, <em>, <div>"),
        ],
        subject: Annotated[str, Field(description="Email subject line")],
        to: Annotated[Optional[List[str]], Field(default=None, description="Recipients")],
        cc: Annotated[Optional[List[str]], Field(default=None, description="CC recipients")],
        bcc: Annotated[Optional[List[str]], Field(default=None, description="BCC recipients")],
        priority: Annotated[str, Field(default="normal", description="low | normal | high")],
        language: Annotated[str, Field(default="cs-CZ", description="Proofing language (e.g., en-US)")],
    ) -> str:
        logger.info("Creating email draft: %s", subject)
        try:
            return create_eml(to=to, cc=cc, bcc=bcc, re=subject, content=content, priority=priority, language=language)
        except Exception as exc:  # pragma: no cover
            logger.exception("Email draft creation failed")
            return f"Error creating email draft: {exc}"


def register_dynamic_email_templates(app: FastMCP) -> None:
    app_config = Path("/app/config/email_templates.yaml")
    local_config = Path(__file__).resolve().parent / "config" / "email_templates.yaml"

    for candidate in (app_config, local_config):
        if candidate.exists():
            try:
                register_email_template_tools_from_yaml(app, candidate)
                logger.info("[dynamic-email] Registered templates from %s", candidate)
            except Exception as exc:  # pragma: no cover
                logger.exception("[dynamic-email] Failed to register templates from %s: %s", candidate, exc)
            break
    else:
        logger.info("[dynamic-email] No dynamic templates found – skipping")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run MCP Office server")
    parser.add_argument("--transport", choices=["stdio", "streamable-http"], default="streamable-http")
    parser.add_argument("--port", type=int, default=8900)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=config.logging.level_no)
    app = build_mcp()

    app.settings.port = args.port
    app.settings.log_level = config.logging.mcp_level_str

    app.run(transport=args.transport)


if __name__ == "__main__":
    main()
