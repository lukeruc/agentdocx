"""Comment operations for docx documents.

Handles w:commentRangeStart/End markers, w:commentReference, and the comments part.
"""

from __future__ import annotations

import datetime
from typing import Optional

from lxml import etree

from .oxml_helpers import (
    NSMAP, W,
    tag, make_element, child, children,
    text_content, get_next_comment_id,
    clone_element, insert_after,
)


def add_comment(
    para: etree._Element,
    doc,
    start_offset: int,
    end_offset: int,
    text: str,
    author: str = "Claude",
) -> dict:
    """Add a comment to a range of text in a paragraph.

    Args:
        para: The w:p paragraph element.
        doc: The DocxDocument instance (needed to access/create comments part).
        start_offset: Start character offset of the commented range.
        end_offset: End character offset of the commented range (exclusive).
        text: Comment text.
        author: Comment author name.

    Returns:
        dict with status info including comment_id.
    """
    if start_offset >= end_offset:
        return {"status": "error", "message": "Invalid comment range: start >= end"}

    para_text = text_content(para)
    if start_offset >= len(para_text):
        return {"status": "error", "message": "Start offset beyond paragraph text"}
    end_offset = min(end_offset, len(para_text))

    # Ensure comments element exists
    if doc._comments is None:
        doc._create_comments_element()

    comments_root = doc._comments
    comment_id = get_next_comment_id(comments_root)

    # Create comment range markers
    range_start = make_element("commentRangeStart", attrib={W + "id": str(comment_id)})
    range_end = make_element("commentRangeEnd", attrib={W + "id": str(comment_id)})

    # Create comment reference run
    ref_run = make_element("r")
    ref_run.append(make_element("commentReference", attrib={W + "id": str(comment_id)}))

    # Insert markers at the right positions
    _insert_comment_markers(para, start_offset, end_offset, range_start, range_end, ref_run)

    # Add comment to comments part
    comment_el = _make_comment_element(comment_id, author, text)
    comments_root.append(comment_el)

    return {
        "status": "ok",
        "message": f"Added comment #{comment_id}: '{text[:50]}...'",
        "comment_id": comment_id,
        "range": [start_offset, end_offset],
    }


def delete_comment(
    para: etree._Element,
    comments_root: Optional[etree._Element],
    comment_id: int,
) -> dict:
    """Delete a comment by ID.

    Removes comment markers from the paragraph and the comment from the comments part.
    """
    # Remove markers from paragraph
    markers_to_remove = []
    for child_el in list(para):
        if child_el.tag == tag("commentRangeStart"):
            if int(child_el.get(W + "id", "-1")) == comment_id:
                markers_to_remove.append(child_el)
        elif child_el.tag == tag("commentRangeEnd"):
            if int(child_el.get(W + "id", "-1")) == comment_id:
                markers_to_remove.append(child_el)
        elif child_el.tag == tag("r"):
            ref = child(child_el, "commentReference")
            if ref is not None and int(ref.get(W + "id", "-1")) == comment_id:
                markers_to_remove.append(child_el)

    for m in markers_to_remove:
        m.getparent().remove(m)

    # Remove comment from comments part
    if comments_root is not None:
        for c in list(comments_root):
            if c.tag == tag("comment") and int(c.get(W + "id", "-1")) == comment_id:
                comments_root.remove(c)
                break

    return {"status": "ok", "message": f"Deleted comment #{comment_id}"}


def list_comments(
    comments_root: Optional[etree._Element],
) -> list[dict]:
    """List all comments in the document."""
    if comments_root is None:
        return []

    result = []
    for c in comments_root:
        if c.tag != tag("comment"):
            continue
        cid = int(c.get(W + "id", "0"))
        author = c.get(W + "author", "Unknown")
        date = c.get(W + "date", "")
        # Extract comment text
        comment_text = ""
        for p in children(c, "p"):
            for r in children(p, "r"):
                t = child(r, "t")
                if t is not None and t.text:
                    comment_text += t.text
            comment_text += "\n"
        comment_text = comment_text.strip()

        result.append({
            "id": cid,
            "author": author,
            "date": date,
            "text": comment_text,
        })

    return result


# ── Internal helpers ────────────────────────────────────────────


def _make_comment_element(comment_id: int, author: str, text: str) -> etree._Element:
    """Create a w:comment element."""
    now = datetime.datetime.now().isoformat()
    attrib = {
        W + "id": str(comment_id),
        W + "author": author,
        W + "date": now,
    }
    comment = make_element("comment", attrib=attrib)

    # The comment text is in a w:p > w:r > w:t structure
    p = make_element("p")
    r = make_element("r")
    rpr = make_element("rPr")
    r.append(rpr)
    t = make_element("t", text=text)
    r.append(t)
    p.append(r)
    comment.append(p)

    return comment


def _insert_comment_markers(
    para: etree._Element,
    start_offset: int,
    end_offset: int,
    range_start: etree._Element,
    range_end: etree._Element,
    ref_run: etree._Element,
):
    """Insert comment range markers and reference run at the given offsets."""
    accumulated = 0
    start_inserted = False
    end_inserted = False

    for child_el in list(para):
        if child_el.tag == tag("r"):
            t = text_content(child_el)
            run_end = accumulated + len(t)

            if not start_inserted and run_end >= start_offset:
                from .oxml_helpers import insert_before
                insert_before(child_el, range_start)
                start_inserted = True

            if not end_inserted and run_end >= end_offset:
                insert_after(child_el, range_end)
                insert_after(range_end, ref_run)
                end_inserted = True

            accumulated += len(t)

        elif child_el.tag == tag("ins"):
            for r in children(child_el, "r"):
                t = text_content(r)
                run_end = accumulated + len(t)
                if not start_inserted and run_end >= start_offset:
                    from .oxml_helpers import insert_before
                    insert_before(child_el, range_start)
                    start_inserted = True
                if not end_inserted and run_end >= end_offset:
                    insert_after(child_el, range_end)
                    insert_after(range_end, ref_run)
                    end_inserted = True
                accumulated += len(t)

        elif child_el.tag == tag("del"):
            for r in children(child_el, "r"):
                dt = "".join(d.text or "" for d in r.iter(tag("delText")))
                accumulated += len(dt)

        if start_inserted and end_inserted:
            break
