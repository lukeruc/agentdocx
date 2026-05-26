# agentdocx

Claude Code-integrable Word (.docx) document editing toolkit with track changes, comments, and formatting.

## Overview

agentdocx is an MCP (Model Context Protocol) server that gives Claude Code the ability to edit .docx files directly. Unlike existing tools that wrap python-docx's high-level API, agentdocx operates at the OOXML level via lxml, enabling features that python-docx does not expose natively: tracked changes (w:ins/w:del), comments (w:comment), and fine-grained run/paragraph property manipulation.

## Architecture

```
Claude Code ── MCP stdio ── agentdocx server ── python-docx (.docx packaging)
                                         └── lxml (OOXML element manipulation)
```

### Module Map

| Module | Lines | Purpose |
|--------|-------|---------|
| `oxml_helpers.py` | 352 | OOXML namespace constants, element factory functions, text/run/paragraph traversal utilities |
| `document.py` | 431 | Document wrapper: open/save, text search, paragraph access, comments part ZIP injection |
| `track_changes.py` | 513 | Revision tracking: `w:ins` (insertion) and `w:del` (deletion) with author/date metadata |
| `comments.py` | 226 | Comment management: `w:commentRangeStart/End`, `w:commentReference`, and comments part CRUD |
| `formatting.py` | 389 | Font properties (name, size, bold, italic, underline, color, highlight, strikethrough, superscript/subscript) and paragraph properties (alignment, indentation, spacing, outline level) |
| `operations.py` | 465 | Find-and-replace, tables, images, headings, bullet/numbered lists |
| `batch.py` | 350+ | Batch executor with semantic (find-based) coordinate resolution and automatic offset recalculation |
| `server.py` | 700+ | MCP server: 22 individual tools + 1 batch tool, tool dispatch, session management |

## Installation

```bash
cd /path/to/agentdocx
pip install -e .
```

### Dependencies

- `python-docx >= 0.8.11` — .docx packaging (OPC/ZIP level)
- `lxml >= 4.9.0` — direct OOXML element tree manipulation
- `mcp >= 1.0.0` — MCP server framework (stdio transport)

### Claude Code Integration

Add to `.claude/settings.local.json`:

```json
{
  "mcpServers": {
    "agentdocx": {
      "command": "python3",
      "args": ["-m", "agentdocx.server"],
      "cwd": "/path/to/agentdocx"
    }
  }
}
```

## Editing Model: Three Layers

agentdocx provides three layers of abstraction for document editing:

### Layer 1: Semantic (find-based) — Recommended for LLM/Agent Use

The caller specifies **what text** to find, and the server resolves exact positions internally.

```json
{"op": "replace_text", "paragraph_index": 92, "find": "2天", "new": "7个工作日"}
{"op": "insert_text",  "paragraph_index": 92, "find": "不得拖延或延误", "position": "after", "text": "新条款内容。"}
{"op": "delete_text",  "paragraph_index": 125, "find": "人民币人民币"}
{"op": "add_comment",  "paragraph_index": 59, "find": "42个月", "text": "建议修改为无限TOC挂钩"}
```

**Find disambiguation**: When `find` text appears multiple times in a paragraph, the server returns all match positions with context. Use `occurrence` to pick a specific one, or `context_before`/`context_after` to narrow:

```json
{"find": {"text": "重新计算", "context_before": "修理或更换合格后"}, "new": "重新计算，无上限"}
```

**Quote normalization**: ASCII `"` and Unicode curly quotes `“` `”` are automatically normalized, so you don't need to match the document's exact quote style.

### Layer 2: Offset-based — For Exact Character-Level Edits

The caller specifies exact `paragraph_index` and character `offset`.

```json
{"op": "insert_text", "paragraph_index": 0, "offset": 5, "text": "Hello"}
{"op": "delete_text", "paragraph_index": 0, "start_offset": 10, "end_offset": 15}
```

### Layer 3: Batch — Multiple Operations, Single Call

All operations in a batch are executed sequentially. **All `find` fields are resolved against the original document state before any modifications are applied.** Offsets are automatically recalculated as text is inserted or deleted.

```json
{
  "doc_id": "/path/to/document.docx",
  "operations": "[...]"
}
```

For offset recalculation, the server tracks a per-paragraph edit history of `(position, delta)` tuples. When operation N+1 references a paragraph, its offsets are adjusted through all prior edits on that paragraph.

## MCP Tools

### Document Lifecycle
| Tool | Description |
|------|-------------|
| `docx_open` | Open a .docx file, returns document structure |
| `docx_save` | Save (optionally to new path) |
| `docx_close` | Close with optional save |

### Query
| Tool | Description |
|------|-------------|
| `docx_get_info` | Paragraph count, styles, heading structure |
| `docx_get_text` | Get text of a paragraph or the whole document |
| `docx_search` | Search for text, returns paragraph index + offset + context |

### Track Changes (Revision Mode)
| Tool | Description |
|------|-------------|
| `docx_insert_text` | Insert with optional `w:ins` revision marks |
| `docx_delete_text` | Delete with optional `w:del` revision marks |

### Comments
| Tool | Description |
|------|-------------|
| `docx_add_comment` | Add a comment on a text range |
| `docx_list_comments` | List all comments |
| `docx_delete_comment` | Delete a comment by ID |

### Formatting
| Tool | Description |
|------|-------------|
| `docx_set_font` | Font name, size (half-points), bold, italic, underline, color (hex), highlight, strikethrough, superscript/subscript |
| `docx_set_paragraph_format` | Alignment, indentation (twips), spacing, line spacing, outline level, page break |
| `docx_set_style` | Apply named style (e.g., "Heading 1", "Normal") |
| `docx_set_heading` | Set paragraph as heading level 1-9 |

### Paragraph Manipulation
| Tool | Description |
|------|-------------|
| `docx_add_paragraph` | Add paragraph at end |
| `docx_insert_paragraph` | Insert at specific index |
| `docx_delete_paragraph` | Delete by index |

### Other
| Tool | Description |
|------|-------------|
| `docx_find_and_replace` | Find and replace across entire document |
| `docx_add_table` | Add table with headers and data rows |
| `docx_insert_image` | Insert image as drawing paragraph |
| `docx_set_list_format` | Apply bullet or numbered list |
| `docx_batch` | **Execute multiple operations in one call** (supports both semantic and offset modes) |

## OOXML Implementation Details

### Track Changes

- **Insertion**: New text is wrapped in `w:ins` elements with `w:id`, `w:author`, and `w:date` attributes. The `w:ins` contains a `w:r` (run) with the inserted text.
- **Deletion**: Original text is moved from `w:r/w:t` into `w:del/w:r/w:delText` elements. The run is split at deletion boundaries, preserving run properties on remaining text.
- **Edge cases handled**: Inserting inside an existing tracked insertion, deleting text that spans multiple runs, deleting text that overlaps with existing `w:ins` or `w:del` regions.

### Comments

Comments in OOXML have two parts:
1. **Inline markers** in the document body: `w:commentRangeStart`, `w:commentRangeEnd`, and a `w:r/w:commentReference` run
2. **Comment content** in a separate `word/comments.xml` part

Since python-docx's OPC layer does not expose a clean API for creating new parts, agentdocx handles this by:
1. Building the `w:comments` element tree in memory
2. Saving the document via python-docx
3. Re-opening the .docx as a ZIP, injecting `word/comments.xml`, updating `[Content_Types].xml` and `word/_rels/document.xml.rels`

### Formatting

- **Font**: Properties are set on `w:rPr` (run properties) children of `w:r` elements. Font size is in half-points (e.g., 24 = 12pt).
- **Paragraph**: Properties are set on `w:pPr` (paragraph properties). Indentation and spacing are in twips (1/20 of a point; 1 inch = 1440 twips).

## Limitations

- **Image embedding**: Creates the OOXML drawing structure but does not handle image binary part injection in the OPC package. For full image support, use python-docx's `add_picture()` on the underlying Document object.
- **Comments via bytes**: The `save_to_bytes()` method does not support comments injection. Use file-based `save()` for full comment support.
- **Style existence**: `set_style` applies a style name reference but does not verify the style exists in the document's styles part. Applying a non-existent style may cause Word to display the paragraph with default formatting.

## License

MIT
