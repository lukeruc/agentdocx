"""Batch executor for multiple docx operations in a single MCP call.

Automatically recalculates offsets and paragraph indices as operations
shift the document structure.

Supports two modes:
1. Offset-based: {"op":"insert_text","paragraph_index":0,"offset":5,"text":"X"}
2. Find-based:  {"op":"replace_text","paragraph_index":0,"old":"X","new":"Y"}
                {"op":"insert_text","paragraph_index":0,"find":"X","position":"after","text":"Y"}
"""

from __future__ import annotations

import json
from typing import Any, Optional

from .document import DocxDocument
from .oxml_helpers import text_content
from .track_changes import insert_text, delete_text
from .comments import add_comment
from .formatting import set_font, set_paragraph_format, set_style
from .operations import find_and_replace, set_heading


def _normalize_quotes(text: str) -> str:
    """Swap all Unicode curly quotes to ASCII straight quotes and vice versa.

    Since the LLM might use "X" while the document uses “X”,
    this creates a variant that catches either case.
    """
    swap = {
        '“': '"', '”': '"',  # curly double → straight
        '‘': "'", '’': "'",  # curly single → straight
        '＂': '"',                 # fullwidth double → straight
        '「': '"', '」': '"',  # corner brackets
        '"': '“',                 # straight → left curly (open)
        "'": '‘',                 # straight → left single curly
    }
    # For straight quotes, we use position-agnostic approach:
    # first occurrence gets opening quote, second gets closing
    result = []
    straight_double_count = 0
    straight_single_count = 0
    for ch in text:
        if ch == '"':
            straight_double_count += 1
            result.append('“' if straight_double_count % 2 == 1 else '”')
        elif ch == "'":
            straight_single_count += 1
            result.append('‘' if straight_single_count % 2 == 1 else '’')
        elif ch in swap:
            result.append(swap[ch])
        else:
            result.append(ch)
    return ''.join(result)


def _resolve_find(doc: DocxDocument, op: dict) -> dict:
    """Resolve a 'find' field into exact paragraph_index + offsets.

    Supports:
      find: "text"           - search in paragraph_index, error if not found
      find: {"text":"...", "paragraph_index": N, "occurrence": N}
      position: "replace"    - delete old text, insert new (for replace_text)
      position: "before"     - insert/delete at start of found text
      position: "after"      - insert/delete at end of found text
      position: "insert_at"  - for add_comment, marks the found range

    If the op already has explicit offsets, find is ignored.
    """
    find = op.get("find")
    if find is None:
        return op

    # Normalize find to dict
    if isinstance(find, str):
        find = {"text": find}
    search_text = find.get("text", find) if isinstance(find, dict) else str(find)
    if isinstance(find, dict):
        search_text = find.get("text", "")

    pi = find.get("paragraph_index", op.get("paragraph_index"))
    if pi is None:
        raise ValueError("find requires paragraph_index")

    occurrence = find.get("occurrence", 1)

    para = doc.paragraphs[pi]
    para_text = text_content(para)

    # Find ALL occurrences first (for dedup warnings)
    all_positions = []
    pos = -1
    while True:
        pos = para_text.find(search_text, pos + 1)
        if pos == -1:
            break
        all_positions.append(pos)

    # If not found with original, try quote normalization
    normalized_search = search_text
    if not all_positions:
        normalized = _normalize_quotes(search_text)
        if normalized != search_text:
            normalized_search = normalized
            pos = -1
            while True:
                pos = para_text.find(normalized, pos + 1)
                if pos == -1:
                    break
                all_positions.append(pos)

    # If still not found, error with helpful context
    if not all_positions:
        snippet = para_text[:100] + "..." if len(para_text) > 100 else para_text
        raise ValueError(
            f"在段落 {pi} 中未找到 '{search_text}'"
            + f"。段落文本: \"{snippet}\""
        )

    # Context-based disambiguation
    context_before = find.get("context_before")
    context_after = find.get("context_after")
    if context_before or context_after:
        filtered = []
        for p in all_positions:
            before_ok = True
            after_ok = True
            if context_before:
                start = max(0, p - len(context_before) - 5)
                snippet = para_text[start:p]
                before_ok = context_before in snippet or context_before in para_text[max(0, p - len(context_before) - 2):p]
            if context_after:
                end = min(len(para_text), p + len(search_text) + len(context_after) + 5)
                snippet = para_text[p + len(search_text):end]
                after_ok = context_after in snippet
            if before_ok and after_ok:
                filtered.append(p)
        if len(filtered) == 1:
            all_positions = filtered
        elif len(filtered) > 1:
            all_positions = filtered  # narrowed but still multiple

    # Multiple matches: require explicit occurrence or unique context
    if len(all_positions) > 1 and occurrence == 1 and not (context_before or context_after):
        # Show all matches with context so caller can disambiguate
        details = []
        for i, p in enumerate(all_positions[:5], 1):  # max 5 shown
            start = max(0, p - 10)
            end = min(len(para_text), p + len(search_text) + 10)
            details.append(f"  #{i} at offset {p}: …{para_text[start:end]}…")
        raise ValueError(
            f"段落 {pi} 中 '{search_text}' 出现 {len(all_positions)} 次。"
            + f"请指定 occurrence 或用更长的 find 文本。匹配位置:\n"
            + "\n".join(details)
        )

    # Use the right occurrence
    idx = all_positions[min(occurrence - 1, len(all_positions) - 1)]
    search_text = normalized_search  # use the version that matched

    position = op.get("position", "replace")

    if position == "before":
        op["offset"] = idx
    elif position == "after":
        op["offset"] = idx + len(search_text)
    elif position == "replace":
        op["start_offset"] = idx
        op["end_offset"] = idx + len(search_text)
        if op.get("op") == "replace_text":
            op["offset"] = idx  # insert at start of deleted range
    elif position == "insert_at":
        op["start_offset"] = idx
        op["end_offset"] = idx + len(search_text)

    op["paragraph_index"] = pi
    # Store resolved offsets for debugging
    op["_resolved"] = {
        "found_at": idx,
        "found_text": search_text,
        "paragraph_index": pi,
    }
    return op


def _op_replace_text(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Semantic replace: find old text, delete it, insert new text.

    Uses either explicit offsets or 'find' field for location.
    """
    pi = op["paragraph_index"]
    start = op["start_offset"]
    end = op["end_offset"]
    new_text = op.get("text", op.get("new", ""))
    paras = doc.paragraphs

    # Delete old text
    del_result = delete_text(
        paras[pi], doc.body, start, end,
        op.get("author", "Claude"), op.get("track_changes", True),
    )
    deleted = del_result.get("deleted_chars", end - start)

    # Insert new text at the deletion point
    ins_result = insert_text(
        paras[pi], doc.body, start, new_text,
        op.get("author", "Claude"), op.get("track_changes", True),
    )

    net_shift = len(new_text) - deleted
    resolved = op.get("_resolved", {})
    return {
        "status": "ok",
        "message": f"Replaced '{resolved.get('found_text', '?')}' -> '{new_text[:50]}' in paragraph {pi}",
        "deleted_chars": deleted,
        "inserted_chars": len(new_text),
        "net_shift": net_shift,
        "resolved": resolved,
    }, net_shift


def _op_insert_text(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute insert_text. Returns (result, offset_shift)."""
    pi = op["paragraph_index"]
    offset = op["offset"]
    paras = doc.paragraphs
    result = insert_text(
        paras[pi], doc.body, offset, op.get("text", ""),
        op.get("author", "Claude"), op.get("track_changes", True),
    )
    if result.get("status") == "ok":
        return result, len(op.get("text", ""))
    return result, 0


def _op_delete_text(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute delete_text. Returns (result, offset_shift)."""
    pi = op["paragraph_index"]
    start = op["start_offset"]
    end = op["end_offset"]
    paras = doc.paragraphs
    result = delete_text(
        paras[pi], doc.body, start, end,
        op.get("author", "Claude"), op.get("track_changes", True),
    )
    deleted = result.get("deleted_chars", 0)
    return result, -deleted


def _op_set_font(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute set_font. No offset impact."""
    pi = op["paragraph_index"]
    paras = doc.paragraphs
    kwargs = {k: v for k, v in op.items() if k not in ("op", "paragraph_index", "find", "position", "_resolved")}
    result = set_font(paras[pi], **kwargs)
    return result, 0


def _op_set_paragraph_format(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute set_paragraph_format. No offset impact."""
    pi = op["paragraph_index"]
    paras = doc.paragraphs
    kwargs = {k: v for k, v in op.items() if k not in ("op", "paragraph_index", "find", "position", "_resolved")}
    result = set_paragraph_format(paras[pi], **kwargs)
    return result, 0


def _op_set_style(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute set_style. No offset impact."""
    pi = op["paragraph_index"]
    paras = doc.paragraphs
    result = set_style(paras[pi], op.get("style_name", op.get("style", "Normal")))
    return result, 0


def _op_set_heading(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute set_heading. No offset impact."""
    pi = op["paragraph_index"]
    paras = doc.paragraphs
    result = set_heading(paras[pi], op.get("level", 1), op.get("text"))
    return result, 0


def _op_add_comment(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute add_comment. No offset impact on paragraph text."""
    pi = op["paragraph_index"]
    paras = doc.paragraphs
    start = op["start_offset"]
    end = op["end_offset"]
    if start >= end:
        return {
            "status": "skipped",
            "message": f"Comment range invalid after offset adjustment (start={start}, end={end}). Target text was likely modified by a prior operation.",
        }, 0
    result = add_comment(
        paras[pi], doc,
        start, end,
        op.get("text", ""), op.get("author", "Claude"),
    )
    return result, 0


def _op_find_and_replace(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute find_and_replace across the whole document."""
    result = find_and_replace(
        doc.body, op["search_text"], op["replace_text"],
        op.get("track_changes", False),
        op.get("author", "Claude"),
        op.get("case_sensitive", True),
    )
    return result, 0


def _op_add_paragraph(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute add_paragraph at end. Returns (result, offset_shift)."""
    idx = doc.add_paragraph(
        op.get("text", ""),
        op.get("style"),
    )
    return {"status": "ok", "paragraph_index": idx, "message": f"Added paragraph at {idx}"}, 0


def _op_insert_paragraph(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute insert_paragraph at index."""
    idx = doc.insert_paragraph(
        op["index"],
        op.get("text", ""),
        op.get("style"),
    )
    return {"status": "ok", "paragraph_index": idx, "message": f"Inserted paragraph at {idx}"}, 0


def _op_delete_paragraph(doc: DocxDocument, op: dict, para_shift: int) -> tuple[dict, int]:
    """Execute delete_paragraph by index."""
    pi = op["paragraph_index"]
    doc.delete_paragraph(pi)
    return {"status": "ok", "message": f"Deleted paragraph {pi}"}, 0


# Operation registry
OPERATIONS = {
    "replace_text": _op_replace_text,
    "insert_text": _op_insert_text,
    "delete_text": _op_delete_text,
    "set_font": _op_set_font,
    "set_paragraph_format": _op_set_paragraph_format,
    "set_style": _op_set_style,
    "set_heading": _op_set_heading,
    "add_comment": _op_add_comment,
    "find_and_replace": _op_find_and_replace,
    "add_paragraph": _op_add_paragraph,
    "insert_paragraph": _op_insert_paragraph,
    "delete_paragraph": _op_delete_paragraph,
}

FIND_RESOLVE_OPS = {"replace_text", "insert_text", "delete_text", "add_comment"}


def _adjust_offset(offset: int, edits: list[tuple[int, int]]) -> int:
    """Adjust an offset through a list of prior edits.

    edits: list of (position, delta) where delta > 0 for insertions,
           delta < 0 for deletions. position is the offset WHERE the edit
           occurred (in the original coordinate system).

    An insertion at position P shifts all offsets >= P right by delta.
    A deletion from P1 to P2 shifts offsets >= P2 left by (P2-P1), and
    offsets inside [P1, P2) are clamped.
    """
    for pos, delta in edits:
        if offset >= pos:
            if delta > 0:
                # Insertion: shift right
                offset += delta
            else:
                # Deletion (delta is negative): shift left
                # The deletion removed text of length abs(delta)
                # Offsets within the deleted range get clamped to pos
                del_len = abs(delta)
                if offset < pos + del_len:
                    offset = pos  # Clamp to deletion start
                else:
                    offset += delta  # shift left
    return max(0, offset)


def execute_batch(doc: DocxDocument, operations_json: str) -> dict:
    """Execute a batch of operations on a document.

    Two modes:
    1. Offset-based (precise):
       {"op":"insert_text","paragraph_index":0,"offset":5,"text":"X"}

    2. Find-based (semantic, recommended for LLM use):
       {"op":"replace_text","paragraph_index":92,"find":"2天","new":"7个工作日"}
       {"op":"insert_text","paragraph_index":92,"find":"不得拖延","position":"after","text":"新条款"}
       {"op":"delete_text","paragraph_index":125,"find":"人民币","occurrence":2}
       {"op":"add_comment","paragraph_index":59,"find":"42个月","text":"建议修改"}

    All offsets in mode 1 are relative to the document state BEFORE any ops.
    In mode 2, the server resolves find→offset before execution.
    """
    try:
        operations = json.loads(operations_json)
    except json.JSONDecodeError as e:
        return {"status": "error", "message": f"Invalid JSON: {e}"}

    if not isinstance(operations, list):
        return {"status": "error", "message": "Operations must be a JSON array"}

    # Step 1: Resolve all 'find' fields to concrete offsets
    resolved_ops = []
    for i, op in enumerate(operations):
        op_type = op.get("op", "")
        if op_type not in OPERATIONS:
            return {
                "status": "error",
                "message": f"Unknown operation at index {i}: {op_type}",
                "known_ops": list(OPERATIONS.keys()),
            }

        if op_type in FIND_RESOLVE_OPS and op.get("find"):
            try:
                op = _resolve_find(doc, op)
            except ValueError as e:
                return {
                    "status": "error",
                    "message": f"Operation [{i}] {op_type}: {e}",
                    "op": op,
                }
        resolved_ops.append(op)

    # Step 2: Execute resolved operations with offset tracking
    para_edits: dict[int, list[tuple[int, int]]] = {}
    results = []
    success_count = 0
    error_count = 0

    for i, op in enumerate(resolved_ops):
        op_type = op["op"]

        try:
            handler = OPERATIONS[op_type]
            pi = op.get("paragraph_index")

            # Save original offsets for edit tracking (pre-adjustment)
            orig_offset = op.get("offset")
            orig_start = op.get("start_offset")
            orig_end = op.get("end_offset")

            # Adjust offsets through prior edits on this paragraph
            if pi is not None and pi in para_edits:
                for key in ("offset", "start_offset", "end_offset"):
                    if key in op and op[key] is not None:
                        op[key] = _adjust_offset(op[key], para_edits[pi])

            # Execute the handler
            result, offset_shift = handler(doc, op, 0)

            # Track edit for future offset adjustments
            if offset_shift != 0 and pi is not None:
                if pi not in para_edits:
                    para_edits[pi] = []

                if op_type in ("insert_text", "replace_text") and orig_offset is not None:
                    para_edits[pi].append((orig_offset, offset_shift))
                elif op_type == "delete_text" and orig_start is not None:
                    para_edits[pi].append((orig_start, offset_shift))

            results.append({"index": i, "op": op_type, **result})

            if result.get("status") == "ok":
                success_count += 1
            else:
                error_count += 1

        except Exception as e:
            results.append({
                "index": i, "op": op_type,
                "status": "error", "message": str(e),
            })
            error_count += 1

    return {
        "status": "ok",
        "total": len(operations),
        "success": success_count,
        "errors": error_count,
        "results": results,
    }
