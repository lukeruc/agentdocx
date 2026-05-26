"""OOXML utility functions and constants for direct XML manipulation."""

from lxml import etree
from copy import deepcopy
from typing import Optional, Sequence

# ── Namespaces ────────────────────────────────────────────────
NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
}

W = "{%s}" % NSMAP["w"]
R = "{%s}" % NSMAP["r"]
WP = "{%s}" % NSMAP["wp"]
A = "{%s}" % NSMAP["a"]
PIC = "{%s}" % NSMAP["pic"]

# ── Tag helpers ────────────────────────────────────────────────


def tag(name: str) -> str:
    """Return a fully qualified w: namespace tag."""
    return f"{W}{name}"


def make_element(name: str, attrib: dict | None = None, text: str | None = None) -> etree._Element:
    """Create a new element in the w: namespace."""
    el = etree.SubElement(etree.Element("root"), tag(name), attrib=attrib) if attrib else etree.Element(tag(name))
    # detach from dummy parent
    if attrib:
        p = el.getparent()
        if p is not None:
            p.remove(el)
    if text is not None:
        el.text = text
    return el


def find_all(element: etree._Element, name: str) -> list[etree._Element]:
    """Find all descendant elements matching the w: tag name."""
    return element.findall(f".//{tag(name)}")


def find_all_ns(element: etree._Element, name: str, ns: str = W) -> list[etree._Element]:
    """Find all descendant elements with explicit namespace."""
    return element.findall(f".//{{{ns}}}{name}")


def find(element: etree._Element, name: str) -> Optional[etree._Element]:
    """Find first descendant element matching the w: tag name."""
    return element.find(f".//{tag(name)}")


def child(element: etree._Element, name: str) -> Optional[etree._Element]:
    """Find first direct child matching the w: tag name."""
    return element.find(tag(name))


def children(element: etree._Element, name: str) -> list[etree._Element]:
    """Find all direct children matching the w: tag name."""
    return element.findall(tag(name))


def text_content(element: etree._Element) -> str:
    """Extract the text content from a paragraph or run element.

    Walks w:t elements, respecting xml:space='preserve'.
    """
    texts = []
    for t in element.iter(tag("t")):
        if t.text:
            texts.append(t.text)
    return "".join(texts)


# ── Run helpers ────────────────────────────────────────────────


def is_tracked_insert(run: etree._Element) -> bool:
    """Check if a w:r element is inside a w:ins (tracked insertion)."""
    p = run.getparent()
    return p is not None and p.tag == tag("ins")


def is_tracked_delete(run: etree._Element) -> bool:
    """Check if a w:r element is inside a w:del (tracked deletion)."""
    p = run.getparent()
    return p is not None and p.tag == tag("del")


def get_paragraph_runs(para: etree._Element) -> list[tuple[etree._Element, str]]:
    """Get all runs in a paragraph with their text, including tracked changes.

    Returns list of (run_element, text) where text is the run's text content.
    For regular runs, run_element is the w:r.
    For tracked insertions, run_element is the w:r inside w:ins.
    For tracked deletions, run_element is the w:r inside w:del (text from w:delText).
    """
    results = []
    for child_el in para:
        if child_el.tag == tag("r"):
            t = text_content(child_el)
            results.append((child_el, t))
        elif child_el.tag == tag("ins"):
            for r in children(child_el, "r"):
                t = text_content(r)
                results.append((r, t))
        elif child_el.tag == tag("del"):
            for r in children(child_el, "r"):
                # Deleted text is in w:delText elements
                del_texts = []
                for dt in r.iter(tag("delText")):
                    if dt.text:
                        del_texts.append(dt.text)
                results.append((r, "".join(del_texts)))
        elif child_el.tag == tag("r"):
            t = text_content(child_el)
            results.append((child_el, t))
    return results


def find_text_position(para: etree._Element, target_offset: int) -> tuple[Optional[etree._Element], int, int]:
    """Locate the run and offset within that run for a given text offset in a paragraph.

    Args:
        para: The w:p element.
        target_offset: Character offset in the paragraph's flat text.

    Returns:
        (run_element, offset_in_run, accumulated_offset_before_this_run)
        If target is beyond text length, run_element will be None.
    """
    accumulated = 0
    for child_el in para:
        if child_el.tag == tag("r"):
            t = text_content(child_el)
            if accumulated + len(t) >= target_offset:
                return (child_el, target_offset - accumulated, accumulated)
            accumulated += len(t)
        elif child_el.tag == tag("ins"):
            for r in children(child_el, "r"):
                t = text_content(r)
                if accumulated + len(t) >= target_offset:
                    return (r, target_offset - accumulated, accumulated)
                accumulated += len(t)
        elif child_el.tag == tag("del"):
            for r in children(child_el, "del"):
                del_texts = []
                for dt in r.iter(tag("delText")):
                    if dt.text:
                        del_texts.append(dt.text)
                t = "".join(del_texts)
                if accumulated + len(t) >= target_offset:
                    return (r, target_offset - accumulated, accumulated)
                accumulated += len(t)
    return (None, 0, accumulated)


def split_run(run: etree._Element, offset: int) -> tuple[etree._Element, etree._Element]:
    """Split a w:r element at the given character offset.

    The first returned run contains text up to (but not including) offset.
    The second contains text from offset to end.

    Args:
        run: The w:r element to split.
        offset: Character offset within the run's text content.

    Returns:
        (left_run, right_run) - both are standalone w:r elements.
    """
    t_el = child(run, "t")
    if t_el is None or not t_el.text:
        # No text to split - return original and empty
        right = deepcopy(run)
        return (run, right)

    full_text = t_el.text
    left_text = full_text[:offset]
    right_text = full_text[offset:]

    # Create left run (keep original, modify text)
    left_run = deepcopy(run)
    left_t = child(left_run, "t")
    if left_t is not None:
        left_t.text = left_text
        # Preserve xml:space attribute
        if not left_text or left_text[0].isspace() or left_text[-1].isspace():
            left_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Create right run (copy, modify text)
    right_run = deepcopy(run)
    right_t = child(right_run, "t")
    if right_t is not None:
        right_t.text = right_text
        if not right_text or right_text[0].isspace() or right_text[-1].isspace():
            right_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    return (left_run, right_run)


def split_run_element_at(run: etree._Element, offset: int) -> tuple[etree._Element, etree._Element]:
    """Split a standalone w:r element in-place at offset.

    Modifies the original run's w:t text to be the left part,
    and creates a new run with the right part inserted after.

    Returns (left_run, right_run).
    """
    t_el = child(run, "t")
    if t_el is None or not t_el.text:
        right = deepcopy(run)
        return (run, right)

    full_text = t_el.text
    left_text = full_text[:offset]
    right_text = full_text[offset:]

    # Modify original run to contain left text
    t_el.text = left_text
    if left_text and (left_text[0].isspace() or left_text[-1].isspace()):
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        t_el.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)

    # Create right run
    right_run = deepcopy(run)
    right_t = child(right_run, "t")
    if right_t is not None:
        right_t.text = right_text
        if right_text and (right_text[0].isspace() or right_text[-1].isspace()):
            right_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        else:
            right_t.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)

    return (run, right_run)


# ── Element insertion helpers ──────────────────────────────────


def insert_after(element: etree._Element, new_sibling: etree._Element) -> None:
    """Insert new_sibling after element in the parent's children."""
    parent = element.getparent()
    if parent is None:
        raise ValueError("Element has no parent")
    idx = list(parent).index(element)
    parent.insert(idx + 1, new_sibling)


def insert_before(element: etree._Element, new_sibling: etree._Element) -> None:
    """Insert new_sibling before element in the parent's children."""
    parent = element.getparent()
    if parent is None:
        raise ValueError("Element has no parent")
    idx = list(parent).index(element)
    parent.insert(idx, new_sibling)


def clone_element(el: etree._Element) -> etree._Element:
    """Deep copy an element."""
    return deepcopy(el)


# ── Comment helpers ─────────────────────────────────────────────


def get_comments_part(docx_element: etree._Element) -> Optional[etree._Element]:
    """Get the comments part from the document's package.

    This requires the python-docx Document object, not just the XML element.
    """
    # This will be called with the document, not here
    return None


def get_next_comment_id(comments_root: etree._Element) -> int:
    """Get the next available comment ID from a comments part."""
    max_id = -1
    for c in comments_root:
        if c.tag == tag("comment"):
            cid = int(c.get(W + "id", "0"))
            max_id = max(max_id, cid)
    return max_id + 1


# ── Revision ID helpers ─────────────────────────────────────────


def get_next_revision_id(body: etree._Element) -> int:
    """Get the next available revision ID for track changes.

    Scans w:ins and w:del elements in the document body.
    """
    max_id = -1
    for el in body.iter():
        if el.tag in (tag("ins"), tag("del")):
            rid = int(el.get(W + "id", "0"))
            max_id = max(max_id, rid)
    return max_id + 1


# ── Property helpers ────────────────────────────────────────────


def ensure_rPr(run: etree._Element) -> etree._Element:
    """Ensure a w:rPr element exists in the run, creating it if needed."""
    rpr = child(run, "rPr")
    if rpr is None:
        rpr = make_element("rPr")
        run.insert(0, rpr)
    return rpr


def ensure_pPr(para: etree._Element) -> etree._Element:
    """Ensure a w:pPr element exists in the paragraph, creating it if needed."""
    ppr = child(para, "pPr")
    if ppr is None:
        ppr = make_element("pPr")
        para.insert(0, ppr)
    return ppr


def set_or_replace(parent: etree._Element, tag_name: str, attrib: dict | None = None) -> etree._Element:
    """Set or replace a child element in the w: namespace.

    Removes any existing child with the same tag, then appends the new one.
    """
    for existing in children(parent, tag_name):
        parent.remove(existing)
    el = make_element(tag_name, attrib=attrib)
    parent.append(el)
    return el


def set_property(parent: etree._Element, tag_name: str, val: str | None = None, attrib: dict | None = None):
    """Set a simple property element. If val is None, the element is self-closing."""
    # Remove existing
    for existing in children(parent, tag_name):
        parent.remove(existing)
    el = make_element(tag_name, attrib=attrib)
    if val is not None:
        el.set(W + "val", str(val))
    parent.append(el)
