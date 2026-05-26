# agentdocx Reference

## Editing Model

agentdocx provides three layers of editing:

### Layer 1: Semantic Mode (find-based) — Primary Interface

Specify text to find; server resolves to exact positions. **Always prefer this.**

```
{"op": "replace_text", "paragraph_index": 92, "find": "2天", "new": "7个工作日"}
{"op": "insert_text",  "paragraph_index": 92, "find": "不得拖延或延误", "position": "after", "text": "默示同意条款。"}
{"op": "delete_text",  "paragraph_index": 125, "find": "人民币人民币"}
{"op": "add_comment",  "paragraph_index": 59, "find": "42个月", "text": "建议修改"}
```

**Find disambiguation**: when the same text appears multiple times:

```json
// By occurrence number
{"find": {"text": "供方", "occurrence": 3}, "new": "卖方"}

// By surrounding context
{"find": {"text": "重新计算", "context_before": "修理或更换合格后"}, "new": "重新计算，无上限"}
```

**Position parameter** (for insert_text):
- `"before"` — insert before found text
- `"after"` — insert after found text

**Quote normalization**: ASCII `"` and Unicode curly `"` `"` are auto-normalized. No need to match document quote style exactly.

### Layer 2: Offset Mode

For precise character-level edits when you have exact positions from `docx_search`:

```json
{"op": "insert_text", "paragraph_index": 0, "offset": 5, "text": "Hello"}
{"op": "delete_text", "paragraph_index": 0, "start_offset": 10, "end_offset": 15}
```

### Layer 3: Batch Execution

All operations in one `docx_batch` call. **All `find` fields resolve against original document state before any modifications execute.** Offsets auto-recalculate as text changes.

## Full Batch Operation Reference

### replace_text

Replace found text with new text.

```json
{
  "op": "replace_text",
  "paragraph_index": 92,
  "find": "2天",
  "new": "7个工作日",
  "track_changes": true,
  "author": "Claude"
}
```

`find` may be a string or `{"text":"...","occurrence":N,"context_before":"...","context_after":"..."}`.

### insert_text

Insert text at a position.

```json
// Semantic
{"op":"insert_text","paragraph_index":92,"find":"不得拖延","position":"after","text":"新条款。","track_changes":true,"author":"Claude"}

// Offset-based
{"op":"insert_text","paragraph_index":0,"offset":5,"text":"Hello","track_changes":true,"author":"Claude"}
```

### delete_text

Delete text.

```json
// Semantic
{"op":"delete_text","paragraph_index":125,"find":"人民币人民币","track_changes":true,"author":"Claude"}

// Offset-based
{"op":"delete_text","paragraph_index":0,"start_offset":10,"end_offset":15,"track_changes":true,"author":"Claude"}
```

### add_comment

Add a Word comment on a text range.

```json
// Semantic
{"op":"add_comment","paragraph_index":59,"find":"42个月","text":"建议修改为无限TOC挂钩","author":"Claude"}

// Offset-based
{"op":"add_comment","paragraph_index":0,"start_offset":0,"end_offset":10,"text":"审核意见","author":"Claude"}
```

### set_font

Set font properties. Applies to entire paragraph unless offsets given.

```json
{
  "op": "set_font",
  "paragraph_index": 88,
  "name": "SimHei",
  "size": 32,
  "bold": true,
  "italic": false,
  "underline": "single",
  "color": "FF0000",
  "highlight": "yellow",
  "strikethrough": false,
  "superscript": false,
  "subscript": false
}
```

- `name`: Font name ("SimSun", "SimHei", "Arial", "Times New Roman")
- `size`: Half-points (24=12pt, 28=14pt, 32=16pt, 36=18pt)
- `underline`: "single", "double", "none"
- `color`: Hex RGB ("FF0000" = red, "0000FF" = blue)
- `highlight`: "yellow", "cyan", "red", "green", "none"

### set_paragraph_format

Set paragraph formatting.

```json
{
  "op": "set_paragraph_format",
  "paragraph_index": 8,
  "alignment": "center",
  "left_indent": 720,
  "right_indent": 0,
  "first_line_indent": 360,
  "space_before": 200,
  "space_after": 100,
  "line_spacing": 1.5,
  "outline_level": 0,
  "page_break_before": false
}
```

- `alignment`: "left", "right", "center", "both" (justify)
- Indent/spacing values in **twips** (1/20 point; 1 inch = 1440 twips, 0.5 inch = 720)
- `line_spacing`: multiplier (1.0, 1.5, 2.0)

### set_style / set_heading

```json
{"op": "set_style", "paragraph_index": 10, "style_name": "Heading 2"}
{"op": "set_heading", "paragraph_index": 41, "level": 1}
```

### find_and_replace

Document-wide find and replace.

```json
{
  "op": "find_and_replace",
  "search_text": "CMEC",
  "replace_text": "CMEC-GROUP",
  "track_changes": true,
  "author": "Claude",
  "case_sensitive": true
}
```

### Paragraph Management

```json
{"op": "add_paragraph", "text": "New paragraph text", "style": "Normal"}
{"op": "insert_paragraph", "index": 5, "text": "Insert at position 5"}
{"op": "delete_paragraph", "paragraph_index": 12}
```

### Table

```json
{
  "op": "add_table",
  "rows": 3,
  "cols": 2,
  "headers": "[\"项目\", \"内容\"]",
  "data": "[[\"编号\", \"001\"], [\"日期\", \"2026-05-26\"]]",
  "after_paragraph": 10
}
```

## Common Workflow: Contract Audit

```
1. docx_open("contract.docx")
   → Get paragraph_count and structure overview

2. docx_search(doc_id, "关键条款名称")
   → Find paragraph_index for content you need to modify

3. docx_get_text(doc_id, paragraph_index=X)
   → Read exact text to verify before editing

4. docx_batch(doc_id, json.dumps([
     {"op":"replace_text","paragraph_index":61,"find":"潜在缺陷","new":"潜在缺陷是指","track_changes":true,"author":"Claude"},
     {"op":"replace_text","paragraph_index":92,"find":"2天","new":"7个工作日","track_changes":true,"author":"Claude"},
     {"op":"set_font","paragraph_index":88,"bold":true},
     {"op":"add_comment","paragraph_index":63,"find":"延长质保期","text":"建议删除总上限","author":"Claude"},
     {"op":"find_and_replace","search_text":"CMEC","replace_text":"CMEC-GROUP","track_changes":true,"author":"Claude"},
   ]))
   → All modifications in a single call

5. docx_get_text on key paragraphs
   → Verify changes look correct

6. docx_save(doc_id, "contract-reviewed.docx")
   → Save to new file
```

## Troubleshooting

### "在段落 N 中未找到 'text'"

- Check that the find text exists **exactly** in the paragraph (use `docx_get_text` to verify)
- Try longer or more specific text
- Quotes may differ (ASCII vs Unicode) — auto-normalization handles most cases

### "出现 N 次。请指定 occurrence"

- Use `{"text":"...","occurrence":N}` to pick which occurrence
- Use `{"text":"...","context_before":"...","context_after":"..."}` to filter by nearby text
- Use longer find text to make it unique

### Comment range invalid after offset adjustment

- The text the comment references was modified by an earlier operation in the batch
- The comment is skipped gracefully with a warning
- Solution: if a comment's target text will also be modified, use a separate batch after the modifications
