---
name: agentdocx
description: Use when editing Word (.docx) documents — reviewing contracts, making tracked changes, adding comments, formatting text. Triggers: "edit this docx", "revise the contract", "add tracked changes", "修改合同", "修订文档", "批注".
---

# agentdocx — Word Document Editing

You have access to the agentdocx MCP server for editing .docx files. All edits support track changes (w:ins/w:del) visible in Microsoft Word.

## Core Rule

**Always use `docx_batch` for all modifications.** One call, all operations. Use semantic mode (find-based) — specify text to find, server resolves the position.

## Workflow

```
docx_open → docx_search/docx_get_text → docx_batch → docx_save
```

## Tools

| Tool | Purpose |
|------|---------|
| `docx_open` | Open .docx, returns paragraph structure |
| `docx_save` | Save (optionally to new path) |
| `docx_close` | Close document |
| `docx_get_info` | Paragraph count, styles, headings |
| `docx_get_text` | Get text of paragraph or full document |
| `docx_search` | Search text, returns paragraph_index + offset |
| `docx_insert_text` | Insert with optional track changes |
| `docx_delete_text` | Delete with optional track changes |
| `docx_add_comment` | Add comment on text range |
| `docx_list_comments` | List all comments |
| `docx_delete_comment` | Delete comment by ID |
| `docx_set_font` | Font: name, size, bold, italic, underline, color, highlight, strikethrough, superscript/subscript |
| `docx_set_paragraph_format` | Paragraph: alignment, indentation, spacing, line spacing |
| `docx_set_style` | Apply named style (e.g. "Heading 1") |
| `docx_set_heading` | Set as heading level 1-9 |
| `docx_add_paragraph` | Add paragraph (supports track_changes) |
| `docx_insert_paragraph` | Insert paragraph at index (supports track_changes) |
| `docx_delete_paragraph` | Delete paragraph by index (supports track_changes) |
| `docx_find_and_replace` | Find and replace across document (supports track_changes) |
| `docx_add_table` | Add table with headers and data |
| `docx_insert_image` | Insert image |
| `docx_set_list_format` | Bullet or numbered list |
| `docx_batch` | **Execute multiple operations in one call** |

## Batch Operations

**Prefer semantic mode (find-based).** Avoid offset calculation.

```json
{"op":"replace_text","paragraph_index":92,"find":"2天","new":"7个工作日","track_changes":true,"author":"吴家坤"}
{"op":"insert_text","paragraph_index":92,"find":"不得拖延","position":"after","text":"新条款。","track_changes":true,"author":"吴家坤"}
{"op":"delete_text","paragraph_index":125,"find":"人民币人民币","track_changes":true,"author":"吴家坤"}
{"op":"add_comment","paragraph_index":59,"find":"42个月","text":"建议修改","author":"吴家坤"}
{"op":"add_paragraph","text":"新增段落内容","track_changes":true,"author":"吴家坤"}
{"op":"insert_paragraph","index":0,"text":"插入首段","track_changes":true,"author":"吴家坤"}
{"op":"delete_paragraph","paragraph_index":5,"track_changes":true,"author":"吴家坤"}
{"op":"set_font","paragraph_index":88,"bold":true}
{"op":"set_heading","paragraph_index":41,"level":1}
{"op":"find_and_replace","search_text":"CMEC","replace_text":"CMEC-GROUP","track_changes":true,"author":"吴家坤"}
```

## Key Rules

1. **Use long, specific `find` text** to avoid ambiguous matches. If multiple matches, use `{"text":"...","occurrence":N}` or `{"text":"...","context_before":"..."}`.
2. **Always `track_changes:true`** for text modifications, and **set `"author"`** to identify who made each change.
3. **Save to new file** when reviewing — don't overwrite original.
4. Font size in half-points (12pt=24). Indentation in twips (1inch=1440).
5. **When editing Chinese documents, use Chinese curly quotes** `""` `''` not ASCII straight quotes `""` `''`. The document text almost always uses curly quotes, and the find-based locator will auto-normalize them, but the inserted/replacement text should also use proper Chinese punctuation.

For detailed reference: read `skill/reference.md`.
