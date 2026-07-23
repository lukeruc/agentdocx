"""Track changes (修订模式) operations for docx documents.

Implements insert and delete with revision marks using OOXML w:ins and w:del elements.
"""

from __future__ import annotations

import datetime
from typing import Optional

from lxml import etree

from .oxml_helpers import (
    NSMAP, W,
    tag, make_element, child, children,
    text_content, get_next_revision_id,
    split_run_element_at, insert_after,
    ensure_rPr, clone_element,
)


def _make_revision_attrs(author: str, rev_id: int) -> dict:
    """Create attributes for a revision element (w:ins or w:del)."""
    now = datetime.datetime.now().isoformat()
    return {
        W + "id": str(rev_id),
        W + "author": author,
        W + "date": now,
    }


def insert_text(
    para: etree._Element,
    body: etree._Element,
    offset: int,
    text: str,
    author: str = "Claude",
    track_changes: bool = True,
) -> dict:
    """Insert text into a paragraph at the given offset.

    Args:
        para: The w:p paragraph element.
        body: The w:body element (needed for revision ID allocation).
        offset: Character offset within the paragraph's text.
        text: Text to insert.
        author: Author name for revision marks.
        track_changes: If True, wraps insertion in w:ins for revision tracking.

    Returns:
        dict with status info.
    """
    if not text:
        return {"status": "noop", "message": "No text to insert"}

    para_text = text_content(para)

    # Handle insertion at end of paragraph
    if offset >= len(para_text):
        return _insert_at_end(para, body, text, author, track_changes)

    # Walk paragraph children to find insertion point
    accumulated = 0
    for child_el in list(para):
        child_tag = child_el.tag

        if child_tag == tag("r"):
            t = text_content(child_el)
            if accumulated + len(t) > offset:
                # Found the run to split
                local_offset = offset - accumulated

                if local_offset == 0:
                    # Insert right before this run
                    if track_changes:
                        ins = _create_ins_element(body, text, author, child_el)
                        _insert_before_in_parent(para, child_el, ins)
                    else:
                        _insert_text_direct(child_el, text, at_start=True)
                    return {"status": "ok", "message": f"Inserted '{text}' at paragraph offset {offset}"}
                elif local_offset == len(t):
                    # Insert right after this run
                    if track_changes:
                        ins = _create_ins_element(body, text, author, child_el)
                        insert_after(child_el, ins)
                    else:
                        _insert_text_direct(child_el, text, at_start=False)
                    return {"status": "ok", "message": f"Inserted '{text}' at paragraph offset {offset}"}
                else:
                    # Split the run and insert between
                    left, right = split_run_element_at(child_el, local_offset)
                    # Replace original with left, insert new content, then right
                    para.replace(child_el, left)

                    if track_changes:
                        ins = _create_ins_element(body, text, author, left)
                        insert_after(left, ins)
                    else:
                        # Append the text to left run
                        left_t = child(left, "t")
                        if left_t is not None:
                            left_t.text = (left_t.text or "") + text

                    insert_after(left, right)
                    return {"status": "ok", "message": f"Inserted '{text}' at paragraph offset {offset}"}

            accumulated += len(t)

        elif child_tag == tag("ins"):
            for r in children(child_el, "r"):
                t = text_content(r)
                if accumulated + len(t) > offset:
                    # Inserting inside a tracked insertion
                    local_offset = offset - accumulated
                    if local_offset == 0:
                        new_r = _make_text_run(text, r)
                        child_el.insert(list(child_el).index(r), new_r)
                    elif local_offset >= len(t):
                        new_r = _make_text_run(text, r)
                        idx = list(child_el).index(r)
                        child_el.insert(idx + 1, new_r)
                    else:
                        # Split the run inside the insertion
                        left_r, right_r = split_run_element_at(r, local_offset)
                        child_el.replace(r, left_r)
                        new_r = _make_text_run(text, left_r)
                        insert_after(left_r, new_r)
                        insert_after(new_r, right_r)
                    return {"status": "ok", "message": f"Inserted '{text}' at paragraph offset {offset}"}
                accumulated += len(t)

        elif child_tag == tag("del"):
            for r in children(child_el, "r"):
                del_texts = []
                for dt in r.iter(tag("delText")):
                    if dt.text:
                        del_texts.append(dt.text)
                t = "".join(del_texts)
                if accumulated + len(t) > offset:
                    # Find neighboring run for format inheritance
                    ref = _find_neighbor_text_run(para, child_el)
                    if track_changes:
                        ins = _create_ins_element(body, text, author, ref)
                        insert_after(child_el, ins)
                    else:
                        new_r = _make_text_run(text, ref)
                        insert_after(child_el, new_r)
                    return {"status": "ok", "message": f"Inserted '{text}' at paragraph offset {offset}"}
                accumulated += len(t)

        # Don't accumulate for non-text elements (comment markers, bookmarks, etc.)

    # If we get here, insert at end
    return _insert_at_end(para, body, text, author, track_changes)


def _find_last_text_run(para: etree._Element) -> Optional[etree._Element]:
    """Find the last text-bearing run in a paragraph for format inheritance."""
    for child_el in reversed(list(para)):
        if child_el.tag == tag("r"):
            t = child(child_el, "t")
            if t is not None and t.text:
                return child_el
        elif child_el.tag == tag("ins"):
            runs = children(child_el, "r")
            if runs:
                return runs[-1]
    return None


def _find_neighbor_text_run(para: etree._Element, target: etree._Element) -> Optional[etree._Element]:
    """Find the nearest text run before or after target for format inheritance."""
    children_list = list(para)
    try:
        idx = children_list.index(target)
    except ValueError:
        return None

    # Look backwards
    for i in range(idx - 1, -1, -1):
        el = children_list[i]
        if el.tag == tag("r"):
            return el
        if el.tag == tag("ins"):
            runs = children(el, "r")
            if runs:
                return runs[-1]

    # Look forwards
    for i in range(idx + 1, len(children_list)):
        el = children_list[i]
        if el.tag == tag("r"):
            return el
        if el.tag == tag("ins"):
            runs = children(el, "r")
            if runs:
                return runs[0]

    return None


def _insert_at_end(
    para: etree._Element,
    body: etree._Element,
    text: str,
    author: str,
    track_changes: bool,
) -> dict:
    """Insert text at the end of a paragraph."""
    ref_run = _find_last_text_run(para)
    if track_changes:
        ins = _create_ins_element(body, text, author, ref_run)
        para.append(ins)
    else:
        r = _make_text_run(text, ref_run)
        para.append(r)
    return {"status": "ok", "message": f"Inserted '{text}' at end of paragraph"}


def delete_text(
    para: etree._Element,
    body: etree._Element,
    start_offset: int,
    end_offset: int,
    author: str = "Claude",
    track_changes: bool = True,
) -> dict:
    """Delete text from a paragraph between start_offset and end_offset.

    Args:
        para: The w:p paragraph element.
        body: The w:body element (needed for revision ID).
        start_offset: Start character offset (inclusive).
        end_offset: End character offset (exclusive).
        author: Author name for revision marks.
        track_changes: If True, wraps deleted text in w:del for revision tracking.

    Returns:
        dict with status info.
    """
    if start_offset >= end_offset:
        return {"status": "noop", "message": "Empty delete range"}

    para_text = text_content(para)
    if start_offset >= len(para_text):
        return {"status": "noop", "message": "Start offset beyond paragraph text"}

    end_offset = min(end_offset, len(para_text))

    if not track_changes:
        return _delete_text_direct(para, start_offset, end_offset)

    return _delete_text_with_revision(para, body, start_offset, end_offset, author)


def _delete_text_with_revision(
    para: etree._Element,
    body: etree._Element,
    start: int,
    end: int,
    author: str,
) -> dict:
    """Delete text with track changes (w:del elements)."""
    rev_id = get_next_revision_id(body)

    # Collect text spans from paragraph children
    spans = _get_text_spans(para)
    if not spans:
        return {"status": "noop", "message": "No text in paragraph"}

    deleted_chars = 0

    # Process spans in reverse order so indices don't shift as we modify
    for child_el, span_start, span_end, span_text in reversed(spans):
        # Calculate overlap
        overlap_start = max(start, span_start)
        overlap_end = min(end, span_end)

        if overlap_start >= overlap_end:
            continue  # No overlap

        local_start = overlap_start - span_start
        local_end = overlap_end - span_start

        if child_el.tag == tag("r"):
            # Regular run - move text to w:del
            _move_to_del(para, child_el, local_start, local_end, author, rev_id)
            deleted_chars += (local_end - local_start)

        elif child_el.tag == tag("ins"):
            # Text is in a tracked insertion - just remove it
            _remove_from_ins(child_el, local_start, local_end)
            deleted_chars += (local_end - local_start)

        elif child_el.tag == tag("del"):
            # Text is already deleted - skip
            pass

    return {
        "status": "ok",
        "message": f"Deleted {deleted_chars} characters from offset {start} to {end}",
        "deleted_chars": deleted_chars,
    }


def _delete_text_direct(
    para: etree._Element,
    start: int,
    end: int,
) -> dict:
    """Delete text directly without revision marks."""
    spans = _get_text_spans(para)
    deleted_chars = 0

    for child_el, span_start, span_end, span_text in reversed(spans):
        overlap_start = max(start, span_start)
        overlap_end = min(end, span_end)

        if overlap_start >= overlap_end:
            continue

        local_start = overlap_start - span_start
        local_end = overlap_end - span_start

        if child_el.tag == tag("r"):
            _remove_text_from_run(child_el, local_start, local_end)
            deleted_chars += (local_end - local_start)
        elif child_el.tag == tag("ins"):
            _remove_from_ins(child_el, local_start, local_end)
            deleted_chars += (local_end - local_start)

    return {
        "status": "ok",
        "message": f"Deleted {deleted_chars} characters",
        "deleted_chars": deleted_chars,
    }


# ── Resolve (accept/reject) revisions ───────────────────────────


def resolve_revisions(
    paras: list[etree._Element],
    action: str = "accept",
    author: Optional[str] = None,
) -> dict:
    """Accept or reject tracked changes (w:ins / w:del) in given paragraphs.

    Args:
        paras: list of w:p elements to process.
        action: "accept" keeps insertions and applies deletions;
                "reject" removes insertions and restores deleted text.
        author: if given, only resolve revisions by this author.

    Returns:
        dict with counts of resolved insertions/deletions.
    """
    if action not in ("accept", "reject"):
        return {"status": "error", "message": "action must be 'accept' or 'reject'"}

    ins_resolved = 0
    del_resolved = 0

    for para in paras:
        i_count, d_count = _resolve_in_para(para, action, author)
        ins_resolved += i_count
        del_resolved += d_count

    return {
        "status": "ok",
        "action": action,
        "ins_resolved": ins_resolved,
        "del_resolved": del_resolved,
        "message": f"{action}ed {ins_resolved} insertion(s) and {del_resolved} deletion(s)",
    }


def _author_matches(el: etree._Element, author: Optional[str]) -> bool:
    """Check if a revision element's w:author matches the given author."""
    if author is None:
        return True
    return el.get(W + "author", "") == author


def _resolve_in_para(para: etree._Element, action: str, author: Optional[str]) -> tuple[int, int]:
    """Resolve all w:ins/w:del in a single paragraph. Returns (ins_count, del_count)."""
    ins_count = 0
    del_count = 0

    # Collect revision elements first (modifying tree during iteration breaks it)
    ins_elements = [c for c in para if c.tag == tag("ins") and _author_matches(c, author)]
    del_elements = [c for c in para if c.tag == tag("del") and _author_matches(c, author)]

    # Also handle paragraph-level ins/del wrappers (from add_paragraph/insert_paragraph
    # with track_changes=True). These are w:ins/w:del containing w:p, at body level.
    # But within a paragraph context, we handle inline ins/del here.

    for ins_el in ins_elements:
        if action == "accept":
            # Keep inserted runs: unwrap w:ins, replace with its child runs
            runs = children(ins_el, "r")
            parent = ins_el.getparent()
            idx = list(parent).index(ins_el)
            for j, r in enumerate(runs):
                parent.insert(idx + j, r)
            parent.remove(ins_el)
        else:
            # reject: remove inserted runs entirely
            ins_el.getparent().remove(ins_el)
        ins_count += 1

    for del_el in del_elements:
        if action == "accept":
            # accept deletion: remove the deleted text entirely
            del_el.getparent().remove(del_el)
        else:
            # reject deletion: restore deleted text as normal runs
            runs = children(del_el, "r")
            parent = del_el.getparent()
            idx = list(parent).index(del_el)
            for j, r in enumerate(runs):
                # Convert w:delText back to w:t
                for dt in r.findall(tag("delText")):
                    dt.tag = tag("t")
                parent.insert(idx + j, r)
            parent.remove(del_el)
        del_count += 1

    return ins_count, del_count


# ── Internal helpers ────────────────────────────────────────────


def _get_text_spans(para: etree._Element) -> list[tuple[etree._Element, int, int, str]]:
    """Get text spans for each child element of a paragraph.

    Returns list of (child_element, start_offset, end_offset, text).
    """
    spans = []
    offset = 0
    for child_el in para:
        if child_el.tag == tag("r"):
            t = text_content(child_el)
            if t:
                spans.append((child_el, offset, offset + len(t), t))
                offset += len(t)
        elif child_el.tag == tag("ins"):
            for r in children(child_el, "r"):
                t = text_content(r)
                if t:
                    spans.append((child_el, offset, offset + len(t), t))
                    offset += len(t)
        elif child_el.tag == tag("del"):
            for r in children(child_el, "r"):
                dt = "".join(d.text or "" for d in r.iter(tag("delText")))
                if dt:
                    spans.append((child_el, offset, offset + len(dt), dt))
                    offset += len(dt)
    return spans


def _create_ins_element(body: etree._Element, text: str, author: str,
                        ref_run: Optional[etree._Element] = None) -> etree._Element:
    """Create a w:ins element containing the text as a tracked insertion.

    If ref_run is provided, copies its w:rPr so inserted text matches
    the formatting of surrounding text.
    """
    rev_id = get_next_revision_id(body)
    ins = make_element("ins", attrib=_make_revision_attrs(author, rev_id))
    r = _make_text_run(text, ref_run)
    ins.append(r)
    return ins


def _copy_rpr(source_run: etree._Element) -> etree._Element:
    """Copy the w:rPr element from a source run, or return an empty one."""
    src_rpr = child(source_run, "rPr")
    if src_rpr is None:
        return make_element("rPr")
    return clone_element(src_rpr)


def _make_text_run(text: str, ref_run: Optional[etree._Element] = None) -> etree._Element:
    """Create a w:r element with a w:t child containing the text.

    If ref_run is provided, copies its w:rPr so the new run inherits
    formatting (font, size, bold, etc.) from the surrounding text.
    """
    r = make_element("r")
    rpr = _copy_rpr(ref_run) if ref_run is not None else make_element("rPr")
    r.append(rpr)
    t = make_element("t", text=text)
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(t)
    return r


def _make_del_text_run(text: str, author: str, rev_id: int,
                       ref_run: Optional[etree._Element] = None) -> etree._Element:
    """Create a w:del element containing w:r > w:delText for deleted text.

    If ref_run is provided, copies its w:rPr for formatting consistency.
    """
    del_el = make_element("del", attrib=_make_revision_attrs(author, rev_id))
    r = make_element("r")
    rpr = _copy_rpr(ref_run) if ref_run is not None else make_element("rPr")
    r.append(rpr)
    dt = make_element("delText", text=text)
    if text and (text[0].isspace() or text[-1].isspace()):
        dt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    r.append(dt)
    del_el.append(r)
    return del_el


def _insert_before_in_parent(parent: etree._Element, ref_child: etree._Element, new_child: etree._Element):
    """Insert new_child before ref_child in parent's children."""
    idx = list(parent).index(ref_child)
    parent.insert(idx, new_child)


def _insert_text_direct(run: etree._Element, text: str, at_start: bool):
    """Insert text directly into a run's w:t element without track changes."""
    t = child(run, "t")
    if t is None:
        t = make_element("t", text=text)
        run.append(t)
    else:
        if at_start:
            t.text = text + (t.text or "")
        else:
            t.text = (t.text or "") + text


def _move_to_del(
    para: etree._Element,
    run: etree._Element,
    local_start: int,
    local_end: int,
    author: str,
    rev_id: int,
):
    """Move text from a w:r element into a w:del element.

    local_start and local_end are offsets within the run's text.
    """
    t_el = child(run, "t")
    if t_el is None:
        return

    full_text = t_el.text or ""
    if local_end > len(full_text):
        local_end = len(full_text)

    deleted_text = full_text[local_start:local_end]
    remaining_left = full_text[:local_start]
    remaining_right = full_text[local_end:]

    if not deleted_text:
        return

    # Create a w:del element with the deleted text
    del_el = _make_del_text_run(deleted_text, author, rev_id)

    if remaining_left and remaining_right:
        # Split the run: left part stays, deleted goes to del, right part in new run
        t_el.text = remaining_left
        if remaining_left[0].isspace() or remaining_left[-1].isspace():
            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

        # Insert del after the left part
        insert_after(run, del_el)

        # Create right run
        right_run = _make_text_run(remaining_right)
        insert_after(del_el, right_run)

    elif remaining_left:
        # Text at end was deleted
        t_el.text = remaining_left
        insert_after(run, del_el)

    elif remaining_right:
        # Text at start was deleted
        t_el.text = remaining_right
        _insert_before_in_parent(para, run, del_el)

    else:
        # Entire run text is deleted
        # Replace the run with the del element
        para.replace(run, del_el)
        # But we need to keep the w:del, and remove the empty w:r... actually,
        # if the whole run is deleted, we just replace it with w:del
        # But check if run has properties we should preserve
        rpr = child(run, "rPr")
        if rpr is not None:
            # Copy run properties to the del element's run
            del_r = child(del_el, "r")
            if del_r is not None:
                existing_rpr = child(del_r, "rPr")
                if existing_rpr is not None:
                    del_r.replace(existing_rpr, rpr)


def _remove_from_ins(
    ins_el: etree._Element,
    local_start: int,
    local_end: int,
):
    """Remove text from within a w:ins element.

    local_start and local_end are offsets within the total text of the w:ins's runs.
    """
    runs = children(ins_el, "r")
    if not runs:
        return

    offset = 0
    runs_to_remove = []
    runs_to_modify = []

    for r in runs:
        t = text_content(r)
        run_start = offset
        run_end = offset + len(t)

        overlap_start = max(local_start, run_start)
        overlap_end = min(local_end, run_end)

        if overlap_start >= overlap_end:
            offset += len(t)
            continue

        if overlap_start <= run_start and overlap_end >= run_end:
            # Entire run is in deletion range
            runs_to_remove.append(r)
        else:
            # Partial overlap
            mod_start = overlap_start - run_start
            mod_end = overlap_end - run_start
            runs_to_modify.append((r, mod_start, mod_end))

        offset += len(t)

    for r in runs_to_remove:
        ins_el.remove(r)

    for r, ms, me in runs_to_modify:
        t_el = child(r, "t")
        if t_el is not None:
            full = t_el.text or ""
            t_el.text = full[:ms] + full[me:]

    # If ins is empty, remove it from the paragraph
    if len(children(ins_el, "r")) == 0:
        ins_el.getparent().remove(ins_el)


def _remove_text_from_run(
    run: etree._Element,
    local_start: int,
    local_end: int,
):
    """Remove text from a run directly (no track changes)."""
    t_el = child(run, "t")
    if t_el is None:
        return

    full = t_el.text or ""
    t_el.text = full[:local_start] + full[local_end:]
