"""MCP server for agentdocx - Claude Code-integrable docx editing.

Provides comprehensive Word document editing via MCP tools:
- Track changes (insert/delete with revision marks)
- Comments (add, list, delete)
- Font and paragraph formatting
- Find and replace
- Tables, images, headings, lists
"""

from __future__ import annotations

import json
import sys
import logging
from pathlib import Path
from typing import Any, Optional

from mcp.server import Server
from mcp.server.stdio import stdio_server
import mcp.types as types

from .document import DocxDocument
from .track_changes import insert_text, delete_text
from .comments import add_comment, delete_comment, list_comments
from .formatting import set_font, set_paragraph_format, set_style
from .batch import execute_batch
from .operations import (
    find_and_replace, add_table, insert_table_after_paragraph,
    insert_image, set_heading, set_list_format,
)

# ── Logging ────────────────────────────────────────────────────

logger = logging.getLogger("agentdocx")
logger.setLevel(logging.INFO)
handler = logging.StreamHandler(sys.stderr)
handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(handler)

# ── Server ─────────────────────────────────────────────────────

server = Server("agentdocx")

# Document store: path -> DocxDocument
_docs: dict[str, DocxDocument] = {}


def _get_doc(doc_id: str) -> DocxDocument:
    if doc_id not in _docs:
        raise ValueError(f"Document '{doc_id}' is not open. Use docx_open first.")
    return _docs[doc_id]


# ── Tool definitions ────────────────────────────────────────────

TOOLS = [
    {
        "name": "docx_open",
        "description": "Open a .docx file for editing. Returns document structure overview.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Absolute or relative path to the .docx file."},
            },
            "required": ["path"],
        },
    },
    {
        "name": "docx_save",
        "description": "Save the document. Optionally save to a new path (Save As).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID (path) returned by docx_open."},
                "path": {"type": "string", "description": "Optional new path for Save As."},
            },
            "required": ["doc_id"],
        },
    },
    {
        "name": "docx_close",
        "description": "Close a document. Optionally save before closing.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID (path) returned by docx_open."},
                "save": {"type": "boolean", "description": "Save before closing (default true)."},
            },
            "required": ["doc_id"],
        },
    },
    {
        "name": "docx_get_info",
        "description": "Get document structure: paragraph count, styles, headings overview.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID (path) returned by docx_open."},
            },
            "required": ["doc_id"],
        },
    },
    {
        "name": "docx_get_text",
        "description": "Get text content. Specify paragraph_index for a single paragraph, or omit for full text.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID (path) returned by docx_open."},
                "paragraph_index": {"type": "integer", "description": "Optional paragraph index (0-based)."},
            },
            "required": ["doc_id"],
        },
    },
    {
        "name": "docx_search",
        "description": "Search for text in the document. Returns matches with paragraph index, offset, and context.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID (path) returned by docx_open."},
                "query": {"type": "string", "description": "Text to search for."},
            },
            "required": ["doc_id", "query"],
        },
    },
    {
        "name": "docx_insert_text",
        "description": "Insert text at a specific position. Use track_changes=True for revision marks.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "offset": {"type": "integer", "description": "Character offset within the paragraph."},
                "text": {"type": "string", "description": "Text to insert."},
                "track_changes": {"type": "boolean", "description": "Use revision tracking (default true)."},
                "author": {"type": "string", "description": "Author name for revision marks (default 'Claude')."},
            },
            "required": ["doc_id", "paragraph_index", "offset", "text"],
        },
    },
    {
        "name": "docx_delete_text",
        "description": "Delete text from a paragraph range. Use track_changes=True for revision marks.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "start_offset": {"type": "integer", "description": "Start character offset (inclusive)."},
                "end_offset": {"type": "integer", "description": "End character offset (exclusive)."},
                "track_changes": {"type": "boolean", "description": "Use revision tracking (default true)."},
                "author": {"type": "string", "description": "Author name for revision marks."},
            },
            "required": ["doc_id", "paragraph_index", "start_offset", "end_offset"],
        },
    },
    {
        "name": "docx_add_comment",
        "description": "Add a comment on a range of text in a paragraph.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "start_offset": {"type": "integer", "description": "Start character offset of commented range."},
                "end_offset": {"type": "integer", "description": "End character offset (exclusive)."},
                "text": {"type": "string", "description": "Comment text."},
                "author": {"type": "string", "description": "Comment author (default 'Claude')."},
            },
            "required": ["doc_id", "paragraph_index", "start_offset", "end_offset", "text"],
        },
    },
    {
        "name": "docx_list_comments",
        "description": "List all comments in the document.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
            },
            "required": ["doc_id"],
        },
    },
    {
        "name": "docx_delete_comment",
        "description": "Delete a comment by ID.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "comment_id": {"type": "integer", "description": "The comment ID to delete."},
                "paragraph_index": {"type": "integer", "description": "The paragraph containing the comment markers."},
            },
            "required": ["doc_id", "comment_id", "paragraph_index"],
        },
    },
    {
        "name": "docx_set_font",
        "description": "Set font properties on text. Common sizes: 24=12pt, 28=14pt, 32=16pt, 36=18pt.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "start_offset": {"type": "integer", "description": "Start offset. Omit for entire paragraph."},
                "end_offset": {"type": "integer", "description": "End offset. Omit for entire paragraph."},
                "name": {"type": "string", "description": "Font name (e.g., 'Arial', 'SimSun', 'Times New Roman')."},
                "size": {"type": "integer", "description": "Font size in half-points (24=12pt, 32=16pt)."},
                "bold": {"type": "boolean", "description": "Bold on/off."},
                "italic": {"type": "boolean", "description": "Italic on/off."},
                "underline": {"type": "string", "description": "'single', 'double', or 'none'."},
                "color": {"type": "string", "description": "Hex color (e.g., 'FF0000' for red)."},
                "highlight": {"type": "string", "description": "Highlight color ('yellow', 'cyan', 'none')."},
                "strikethrough": {"type": "boolean", "description": "Strikethrough on/off."},
                "superscript": {"type": "boolean", "description": "Superscript."},
                "subscript": {"type": "boolean", "description": "Subscript."},
            },
            "required": ["doc_id", "paragraph_index"],
        },
    },
    {
        "name": "docx_set_paragraph_format",
        "description": "Set paragraph formatting: alignment, indentation, spacing.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "alignment": {"type": "string", "description": "'left', 'right', 'center', or 'both' (justify)."},
                "left_indent": {"type": "integer", "description": "Left indent in twips (1 inch = 1440 twips)."},
                "right_indent": {"type": "integer", "description": "Right indent in twips."},
                "first_line_indent": {"type": "integer", "description": "First line indent in twips."},
                "space_before": {"type": "integer", "description": "Space before in twips."},
                "space_after": {"type": "integer", "description": "Space after in twips."},
                "line_spacing": {"type": "number", "description": "Line spacing multiplier (1.0, 1.5, 2.0)."},
                "outline_level": {"type": "integer", "description": "Outline level 0-9."},
                "page_break_before": {"type": "boolean", "description": "Page break before paragraph."},
            },
            "required": ["doc_id", "paragraph_index"],
        },
    },
    {
        "name": "docx_set_style",
        "description": "Apply a named paragraph style (e.g., 'Normal', 'Heading 1', 'Title', 'Quote').",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "style_name": {"type": "string", "description": "Style name (e.g., 'Heading 1', 'Normal', 'Quote')."},
            },
            "required": ["doc_id", "paragraph_index", "style_name"],
        },
    },
    {
        "name": "docx_set_heading",
        "description": "Set a paragraph as a heading with a specific level (1-9).",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index (0-based)."},
                "level": {"type": "integer", "description": "Heading level 1-9 (default 1)."},
                "text": {"type": "string", "description": "Optional new heading text."},
            },
            "required": ["doc_id", "paragraph_index"],
        },
    },
    {
        "name": "docx_add_paragraph",
        "description": "Add a new paragraph to the end of the document. Supports track_changes for revision marking.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "text": {"type": "string", "description": "Paragraph text."},
                "style": {"type": "string", "description": "Optional style name (e.g., 'Heading 1')."},
                "track_changes": {"type": "boolean", "description": "Wrap the new paragraph in w:ins for revision tracking."},
                "author": {"type": "string", "description": "Author name for revision marks."},
            },
            "required": ["doc_id"],
        },
    },
    {
        "name": "docx_insert_paragraph",
        "description": "Insert a paragraph at a specific index. Supports track_changes for revision marking.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "index": {"type": "integer", "description": "Insert position (0=beginning)."},
                "text": {"type": "string", "description": "Paragraph text."},
                "style": {"type": "string", "description": "Optional style name."},
                "track_changes": {"type": "boolean", "description": "Wrap the new paragraph in w:ins for revision tracking."},
                "author": {"type": "string", "description": "Author name for revision marks."},
            },
            "required": ["doc_id", "index"],
        },
    },
    {
        "name": "docx_delete_paragraph",
        "description": "Delete a paragraph by index. Use track_changes=True to wrap in w:del instead of removing.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index to delete."},
                "track_changes": {"type": "boolean", "description": "Wrap the paragraph in w:del for revision tracking."},
                "author": {"type": "string", "description": "Author name for revision marks."},
            },
            "required": ["doc_id", "paragraph_index"],
        },
    },
    {
        "name": "docx_find_and_replace",
        "description": "Find and replace text throughout the document. Optionally use track changes.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "search_text": {"type": "string", "description": "Text to search for."},
                "replace_text": {"type": "string", "description": "Replacement text."},
                "track_changes": {"type": "boolean", "description": "Use track changes for replacements."},
                "case_sensitive": {"type": "boolean", "description": "Case-sensitive search (default true)."},
                "author": {"type": "string", "description": "Author name for revision marks."},
            },
            "required": ["doc_id", "search_text", "replace_text"],
        },
    },
    {
        "name": "docx_add_table",
        "description": "Add a table. headers and data are JSON arrays of arrays.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "rows": {"type": "integer", "description": "Number of rows."},
                "cols": {"type": "integer", "description": "Number of columns."},
                "headers": {"type": "string", "description": "JSON array of header texts (e.g., '[\"Name\",\"Age\"]')."},
                "data": {"type": "string", "description": "JSON array of row arrays."},
                "after_paragraph": {"type": "integer", "description": "Insert after this paragraph index."},
            },
            "required": ["doc_id", "rows", "cols"],
        },
    },
    {
        "name": "docx_insert_image",
        "description": "Insert an image as a drawing paragraph.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "image_path": {"type": "string", "description": "Path to image file (.png, .jpg, .gif, .bmp)."},
                "width": {"type": "integer", "description": "Width in EMUs (1mm ~ 36000 EMU, default 3000000)."},
                "height": {"type": "integer", "description": "Height in EMUs (default 2000000)."},
                "after_paragraph": {"type": "integer", "description": "Insert after this paragraph index."},
            },
            "required": ["doc_id", "image_path"],
        },
    },
    {
        "name": "docx_set_list_format",
        "description": "Apply bullet or numbered list formatting to a paragraph.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "paragraph_index": {"type": "integer", "description": "Paragraph index."},
                "list_type": {"type": "string", "description": "'bullet' or 'number'."},
                "level": {"type": "integer", "description": "List nesting level (0-based)."},
            },
            "required": ["doc_id", "paragraph_index"],
        },
    },
    {
        "name": "docx_batch",
        "description": (
            "Execute multiple operations in a single call. Two modes:\n"
            "1. SEMANTIC MODE (recommended): replace_text with find/new; "
            "insert_text/delete_text/add_comment with find+position. "
            "No offset calculation needed.\n"
            "2. OFFSET MODE: insert_text/delete_text with paragraph_index+offset.\n"
            "Other ops: set_font, set_paragraph_format, set_style, set_heading, "
            "find_and_replace, add_paragraph, insert_paragraph, delete_paragraph."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_id": {"type": "string", "description": "The document ID."},
                "operations": {
                    "type": "string",
                    "description": "JSON array of operations. Use SEMANTIC MODE (find-based) — no offset calculation. "
                                   "replace_text: paragraph_index, find, new. "
                                   "insert_text: paragraph_index, find, position(before/after), text. "
                                   "delete_text: paragraph_index, find, occurrence(optional). "
                                   "add_comment: paragraph_index, find, text. "
                                   "find can be a string or {text, occurrence, context_before, context_after}. "
                                   "Use long specific find text to avoid ambiguous matches. "
                                   "set_font/set_paragraph_format/set_style/set_heading: paragraph_index, props.",
                },
            },
            "required": ["doc_id", "operations"],
        },
    },
]


# ── Tool handlers ───────────────────────────────────────────────


@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    """Return the list of available tools."""
    return [types.Tool(**t) for t in TOOLS]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict[str, Any]) -> list[types.TextContent]:
    """Dispatch tool calls."""
    try:
        if name == "docx_open":
            return await _docx_open(**arguments)
        elif name == "docx_save":
            return await _docx_save(**arguments)
        elif name == "docx_close":
            return await _docx_close(**arguments)
        elif name == "docx_get_info":
            return await _docx_get_info(**arguments)
        elif name == "docx_get_text":
            return await _docx_get_text(**arguments)
        elif name == "docx_search":
            return await _docx_search(**arguments)
        elif name == "docx_insert_text":
            return await _docx_insert_text(**arguments)
        elif name == "docx_delete_text":
            return await _docx_delete_text(**arguments)
        elif name == "docx_add_comment":
            return await _docx_add_comment(**arguments)
        elif name == "docx_list_comments":
            return await _docx_list_comments(**arguments)
        elif name == "docx_delete_comment":
            return await _docx_delete_comment(**arguments)
        elif name == "docx_set_font":
            return await _docx_set_font(**arguments)
        elif name == "docx_set_paragraph_format":
            return await _docx_set_paragraph_format(**arguments)
        elif name == "docx_set_style":
            return await _docx_set_style(**arguments)
        elif name == "docx_set_heading":
            return await _docx_set_heading(**arguments)
        elif name == "docx_add_paragraph":
            return await _docx_add_paragraph(**arguments)
        elif name == "docx_insert_paragraph":
            return await _docx_insert_paragraph(**arguments)
        elif name == "docx_delete_paragraph":
            return await _docx_delete_paragraph(**arguments)
        elif name == "docx_find_and_replace":
            return await _docx_find_and_replace(**arguments)
        elif name == "docx_add_table":
            return await _docx_add_table(**arguments)
        elif name == "docx_insert_image":
            return await _docx_insert_image(**arguments)
        elif name == "docx_set_list_format":
            return await _docx_set_list_format(**arguments)
        elif name == "docx_batch":
            return await _docx_batch(**arguments)
        else:
            return [types.TextContent(type="text", text=json.dumps({"status": "error", "message": f"Unknown tool: {name}"}, ensure_ascii=False))]
    except Exception as e:
        logger.exception(f"Error in tool {name}")
        return [types.TextContent(type="text", text=json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False))]


# ── Tool implementations ────────────────────────────────────────

def _ok(data: Any) -> list[types.TextContent]:
    return [types.TextContent(type="text", text=json.dumps(data, ensure_ascii=False, indent=2))]


# ── Document lifecycle ──────────────────────────────────────────

async def _docx_open(path: str) -> list[types.TextContent]:
    doc_path = Path(path).resolve()
    if not doc_path.exists():
        return _ok({"status": "error", "message": f"File not found: {path}"})

    doc = DocxDocument(str(doc_path))
    doc_id = str(doc_path)
    _docs[doc_id] = doc

    info = doc.get_info()
    return _ok({
        "status": "ok",
        "message": f"Opened: {doc_path.name}",
        "doc_id": doc_id,
        "paragraph_count": info.paragraph_count,
        "paragraphs": [
            {"index": p.index, "style": p.style, "is_heading": p.is_heading,
             "heading_level": p.heading_level, "text_preview": p.text}
            for p in info.paragraphs
        ],
    })


async def _docx_save(doc_id: str, path: Optional[str] = None) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    save_path = doc.save(path)
    return _ok({"status": "ok", "message": f"Saved to: {save_path}", "path": str(save_path)})


async def _docx_close(doc_id: str, save: bool = True) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    if save:
        doc.save()
    del _docs[doc_id]
    return _ok({"status": "ok", "message": f"Closed: {doc.path.name}"})


# ── Document info ───────────────────────────────────────────────

async def _docx_get_info(doc_id: str) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    info = doc.get_info()
    return _ok({
        "status": "ok",
        "paragraph_count": info.paragraph_count,
        "paragraphs": [
            {"index": p.index, "style": p.style, "is_heading": p.is_heading,
             "heading_level": p.heading_level, "text_preview": p.text}
            for p in info.paragraphs
        ],
    })


async def _docx_get_text(doc_id: str, paragraph_index: Optional[int] = None) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    if paragraph_index is not None:
        text = doc.get_paragraph_text(paragraph_index)
        return _ok({"status": "ok", "paragraph_index": paragraph_index, "text": text})
    else:
        text = doc.get_full_text()
        return _ok({"status": "ok", "text": text, "paragraph_count": doc.paragraph_count})


async def _docx_search(doc_id: str, query: str) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    results = doc.search_text(query)
    return _ok({"status": "ok", "query": query, "match_count": len(results), "matches": results})


# ── Track changes ───────────────────────────────────────────────

async def _docx_insert_text(
    doc_id: str, paragraph_index: int, offset: int, text: str,
    track_changes: bool = True, author: str = "Claude",
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = insert_text(paras[paragraph_index], doc.body, offset, text, author, track_changes)
    return _ok(result)


async def _docx_delete_text(
    doc_id: str, paragraph_index: int, start_offset: int, end_offset: int,
    track_changes: bool = True, author: str = "Claude",
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = delete_text(paras[paragraph_index], doc.body, start_offset, end_offset, author, track_changes)
    return _ok(result)


# ── Comments ────────────────────────────────────────────────────

async def _docx_add_comment(
    doc_id: str, paragraph_index: int, start_offset: int, end_offset: int,
    text: str, author: str = "Claude",
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = add_comment(paras[paragraph_index], doc, start_offset, end_offset, text, author)
    return _ok(result)


async def _docx_list_comments(doc_id: str) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    comments = list_comments(doc.comments_element)
    return _ok({"status": "ok", "count": len(comments), "comments": comments})


async def _docx_delete_comment(doc_id: str, comment_id: int, paragraph_index: int) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range"})
    result = delete_comment(paras[paragraph_index], doc.comments_element, comment_id)
    return _ok(result)


# ── Formatting ──────────────────────────────────────────────────

async def _docx_set_font(
    doc_id: str, paragraph_index: int,
    start_offset: Optional[int] = None, end_offset: Optional[int] = None,
    name: Optional[str] = None, size: Optional[int] = None,
    bold: Optional[bool] = None, italic: Optional[bool] = None,
    underline: Optional[str] = None, color: Optional[str] = None,
    highlight: Optional[str] = None, strikethrough: Optional[bool] = None,
    superscript: Optional[bool] = None, subscript: Optional[bool] = None,
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = set_font(
        paras[paragraph_index], start_offset, end_offset,
        name=name, size=size, bold=bold, italic=italic,
        underline=underline, color=color, highlight=highlight,
        strikethrough=strikethrough,
        superscript=superscript, subscript=subscript,
    )
    return _ok(result)


async def _docx_set_paragraph_format(
    doc_id: str, paragraph_index: int,
    alignment: Optional[str] = None,
    left_indent: Optional[int] = None, right_indent: Optional[int] = None,
    first_line_indent: Optional[int] = None,
    space_before: Optional[int] = None, space_after: Optional[int] = None,
    line_spacing: Optional[float] = None,
    outline_level: Optional[int] = None,
    page_break_before: Optional[bool] = None,
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = set_paragraph_format(
        paras[paragraph_index],
        alignment=alignment, left_indent=left_indent, right_indent=right_indent,
        first_line_indent=first_line_indent,
        space_before=space_before, space_after=space_after,
        line_spacing=line_spacing, outline_level=outline_level,
        page_break_before=page_break_before,
    )
    return _ok(result)


async def _docx_set_style(doc_id: str, paragraph_index: int, style_name: str) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = set_style(paras[paragraph_index], style_name)
    return _ok(result)


async def _docx_set_heading(
    doc_id: str, paragraph_index: int, level: int = 1, text: Optional[str] = None,
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = set_heading(paras[paragraph_index], level, text)
    return _ok(result)


# ── Paragraph ops ───────────────────────────────────────────────

async def _docx_add_paragraph(doc_id: str, text: str = "", style: Optional[str] = None,
                               track_changes: bool = False, author: str = "Claude") -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    idx = doc.add_paragraph(text, style, track_changes, author)
    return _ok({"status": "ok", "message": f"Added paragraph at index {idx}", "paragraph_index": idx})


async def _docx_insert_paragraph(
    doc_id: str, index: int, text: str = "", style: Optional[str] = None,
    track_changes: bool = False, author: str = "Claude",
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    idx = doc.insert_paragraph(index, text, style, track_changes, author)
    return _ok({"status": "ok", "message": f"Inserted paragraph at index {idx}", "paragraph_index": idx})


async def _docx_delete_paragraph(doc_id: str, paragraph_index: int,
                                  track_changes: bool = False, author: str = "Claude") -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    doc.delete_paragraph(paragraph_index, track_changes, author)
    return _ok({"status": "ok", "message": f"Deleted paragraph {paragraph_index}"})


# ── Find & replace ─────────────────────────────────────────────

async def _docx_find_and_replace(
    doc_id: str, search_text: str, replace_text: str,
    track_changes: bool = False, case_sensitive: bool = True,
    author: str = "Claude",
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    result = find_and_replace(doc.body, search_text, replace_text, track_changes, author, case_sensitive=case_sensitive)
    return _ok(result)


# ── Table ──────────────────────────────────────────────────────

async def _docx_add_table(
    doc_id: str, rows: int, cols: int,
    headers: Optional[str] = None, data: Optional[str] = None,
    after_paragraph: Optional[int] = None,
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    header_row = json.loads(headers) if headers else None
    data_rows = json.loads(data) if data else None
    if after_paragraph is not None:
        result = insert_table_after_paragraph(doc.body, after_paragraph, rows, cols, header_row, data_rows)
    else:
        result = add_table(doc.body, rows, cols, header_row, data_rows)
    return _ok(result)


# ── Image ──────────────────────────────────────────────────────

async def _docx_insert_image(
    doc_id: str, image_path: str, width: int = 3000000, height: int = 2000000,
    after_paragraph: Optional[int] = None,
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    result = insert_image(doc.body, image_path, width, height, after_paragraph)
    return _ok(result)


# ── List ───────────────────────────────────────────────────────

async def _docx_set_list_format(
    doc_id: str, paragraph_index: int, list_type: str = "bullet", level: int = 0,
) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    paras = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paras):
        return _ok({"status": "error", "message": f"Paragraph {paragraph_index} out of range (0-{len(paras)-1})"})
    result = set_list_format(paras[paragraph_index], list_type, level)
    return _ok(result)


# ── Batch operations ────────────────────────────────────────────

async def _docx_batch(doc_id: str, operations: str) -> list[types.TextContent]:
    doc = _get_doc(doc_id)
    result = execute_batch(doc, operations)
    return _ok(result)


# ── Entry point ─────────────────────────────────────────────────

async def main():
    logger.info("Starting agentdocx MCP server")
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


def run():
    import asyncio
    asyncio.run(main())


if __name__ == "__main__":
    run()
