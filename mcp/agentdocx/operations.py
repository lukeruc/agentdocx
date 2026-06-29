"""Common document operations: find/replace, tables, images, headings."""

from __future__ import annotations

import uuid
from pathlib import Path
from typing import Optional, Union

from lxml import etree

from .oxml_helpers import (
    NSMAP, W, WP, A, PIC,
    tag, make_element, child, children,
    text_content, ensure_pPr, insert_after,
    set_or_replace, set_property,
)
from .track_changes import insert_text, delete_text
from .formatting import set_style


# ── Find and Replace ────────────────────────────────────────────


def find_and_replace(
    body: etree._Element,
    search_text: str,
    replace_text: str,
    track_changes: bool = False,
    author: str = "Claude",
    case_sensitive: bool = True,
) -> dict:
    """Find and replace text throughout the document.

    Args:
        body: The w:body element.
        search_text: Text to search for.
        replace_text: Replacement text.
        track_changes: If True, use track changes for the replacement.
        author: Author name for track changes.
        case_sensitive: Whether search is case-sensitive.

    Returns:
        dict with count of replacements.
    """
    count = 0
    for para in children(body, "p"):
        para_text = text_content(para)

        if case_sensitive:
            idx = para_text.find(search_text)
        else:
            idx = para_text.lower().find(search_text.lower())

        while idx != -1:
            start = idx
            end = idx + len(search_text)

            # Delete the old text
            delete_text(para, body, start, end, author, track_changes=track_changes)

            # Insert the new text
            insert_text(para, body, start, replace_text, author, track_changes=track_changes)

            count += 1

            # Re-read paragraph text and continue searching
            para_text = text_content(para)
            next_start = start + len(replace_text)
            if case_sensitive:
                idx = para_text.find(search_text, next_start)
            else:
                idx = para_text.lower().find(search_text.lower(), next_start)

    return {
        "status": "ok",
        "message": f"Replaced {count} occurrence(s) of '{search_text}' with '{replace_text}'",
        "count": count,
    }


# ── Table operations ────────────────────────────────────────────


def add_table(
    body: etree._Element,
    rows: int,
    cols: int,
    header_row: Optional[list[str]] = None,
    data_rows: Optional[list[list[str]]] = None,
) -> dict:
    """Add a table to the end of the document.

    Args:
        body: The w:body element.
        rows: Number of rows.
        cols: Number of columns.
        header_row: Optional header row text for each column.
        data_rows: Optional data rows, each a list of cell texts.

    Returns:
        dict with status info.
    """
    table = _make_table(rows, cols, header_row, data_rows)
    body.append(table)
    return {
        "status": "ok",
        "message": f"Added table with {rows} rows and {cols} columns",
    }


def _make_table(
    rows: int,
    cols: int,
    header_row: Optional[list[str]] = None,
    data_rows: Optional[list[list[str]]] = None,
) -> etree._Element:
    """Create a w:tbl element."""
    tbl = make_element("tbl")

    # Table properties
    tbl_pr = make_element("tblPr")
    tbl_width = make_element("tblW", attrib={
        W + "w": "5000",
        W + "type": "pct",
    })
    tbl_pr.append(tbl_width)
    tbl_borders = make_element("tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = make_element(border_name, attrib={
            W + "val": "single",
            W + "sz": "4",
            W + "space": "0",
            W + "color": "auto",
        })
        tbl_borders.append(border)
    tbl_pr.append(tbl_borders)
    tbl.append(tbl_pr)

    # Grid
    tbl_grid = make_element("tblGrid")
    for _ in range(cols):
        grid_col = make_element("gridCol", attrib={W + "w": str(5000 // cols)})
        tbl_grid.append(grid_col)
    tbl.append(tbl_grid)

    # Rows
    for row_idx in range(rows):
        tr = make_element("tr")
        for col_idx in range(cols):
            tc = make_element("tc")
            tc_pr = make_element("tcPr")
            tc.append(tc_pr)

            # Determine cell text
            cell_text = ""
            is_header = header_row and row_idx == 0
            if is_header and col_idx < len(header_row):
                cell_text = header_row[col_idx]
            elif data_rows and row_idx < len(data_rows):
                data_row = data_rows[row_idx]
                if col_idx < len(data_row):
                    cell_text = data_row[col_idx]
                else:
                    cell_text = ""

            # Create paragraph with text
            p = make_element("p")
            if cell_text:
                r = make_element("r")
                if is_header:
                    rpr = make_element("rPr")
                    rpr.append(make_element("b"))
                    r.append(rpr)
                t = make_element("t", text=cell_text)
                r.append(t)
                p.append(r)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)

    return tbl


def insert_table_after_paragraph(
    body: etree._Element,
    para_index: int,
    rows: int,
    cols: int,
    header_row: Optional[list[str]] = None,
    data_rows: Optional[list[list[str]]] = None,
) -> dict:
    """Insert a table after a specific paragraph.

    Args:
        body: The w:body element.
        para_index: Index of the paragraph to insert after.
        rows: Number of rows.
        cols: Number of columns.
        header_row: Optional header row texts.
        data_rows: Optional data rows.

    Returns:
        dict with status info.
    """
    paras = children(body, "p")
    if para_index < 0 or para_index >= len(paras):
        return {"status": "error", "message": f"Paragraph index {para_index} out of range"}

    table = _make_table(rows, cols, header_row, data_rows)
    target_para = paras[para_index]
    insert_after(target_para, table)

    return {
        "status": "ok",
        "message": f"Inserted {rows}x{cols} table after paragraph {para_index}",
    }


# ── Image operations ────────────────────────────────────────────


def insert_image(
    body: etree._Element,
    image_path: str,
    width: int = 300,  # in EMUs (English Metric Units)
    height: int = 200,
    after_paragraph: Optional[int] = None,
) -> dict:
    """Insert an image into the document.

    Note: Due to image format complexities, this creates a paragraph with
    a placeholder that indicates where the image should be. Full image
    embedding requires working with the OPC package for image parts.

    For production use, consider using python-docx's add_picture() on the
    python-docx Document object.

    Args:
        body: The w:body element.
        image_path: Path to the image file.
        width: Image width in EMUs (1px ≈ 9500 EMU).
        height: Image height in EMUs.
        after_paragraph: Insert after this paragraph index, or at end if None.

    Returns:
        dict with status info.
    """
    image_path_obj = Path(image_path)

    if not image_path_obj.exists():
        return {"status": "error", "message": f"Image not found: {image_path}"}

    # Determine image type
    ext = image_path_obj.suffix.lower()
    if ext in (".png",):
        content_type = "image/png"
    elif ext in (".jpg", ".jpeg"):
        content_type = "image/jpeg"
    elif ext in (".gif",):
        content_type = "image/gif"
    elif ext in (".bmp",):
        content_type = "image/bmp"
    elif ext in (".tiff", ".tif"):
        content_type = "image/tiff"
    elif ext in (".svg",):
        content_type = "image/svg+xml"
    else:
        return {"status": "error", "message": f"Unsupported image format: {ext}"}

    # Read image data
    with open(image_path_obj, "rb") as f:
        image_data = f.read()

    # For simple use, we add the image via python-docx's API which handles
    # the packaging properly. But since we're working at the XML level,
    # we create a drawing element manually.

    # Create a new paragraph with the image
    p = make_element("p")

    r = make_element("r")
    drawing = _make_image_drawing(
        image_data, content_type, width, height,
        f"rId{hash(image_path_obj.name) & 0x7FFFFFFF}"
    )
    r.append(drawing)
    p.append(r)

    if after_paragraph is not None:
        paras = children(body, "p")
        if after_paragraph < 0 or after_paragraph >= len(paras):
            return {"status": "error", "message": f"Paragraph index {after_paragraph} out of range"}
        insert_after(paras[after_paragraph], p)
    else:
        body.append(p)

    return {
        "status": "ok",
        "message": f"Inserted image '{image_path_obj.name}' ({width}x{height})",
    }


def _make_image_drawing(
    image_data: bytes,
    content_type: str,
    width: int,
    height: int,
    rId: str,
) -> etree._Element:
    """Create a w:drawing element for an inline image."""
    drawing = etree.SubElement(
        etree.Element("root"), tag("drawing")
    )

    # wp:inline
    inline = etree.SubElement(drawing, f"{WP}inline", attrib={
        "distT": "0",
        "distB": "0",
        "distL": "0",
        "distR": "0",
    })

    # wp:extent
    etree.SubElement(inline, f"{WP}extent", attrib={
        "cx": str(width),
        "cy": str(height),
    })

    # wp:effectExtent
    effect = etree.SubElement(inline, f"{WP}effectExtent", attrib={
        "l": "0", "t": "0", "r": "0", "b": "0",
    })

    # wp:docPr
    etree.SubElement(inline, f"{WP}docPr", attrib={
        "id": "1",
        "name": "Picture",
    })

    # wp:cNvGraphicFramePr
    etree.SubElement(inline, f"{WP}cNvGraphicFramePr")

    # a:graphic
    graphic = etree.SubElement(inline, f"{A}graphic")
    graphic_data = etree.SubElement(graphic, f"{A}graphicData", attrib={
        "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    })

    # pic:pic
    pic = etree.SubElement(graphic_data, f"{PIC}pic")
    nv_pic_pr = etree.SubElement(pic, f"{PIC}nvPicPr")
    c_nv_pr = etree.SubElement(nv_pic_pr, f"{PIC}cNvPr", attrib={
        "id": "0",
        "name": "Picture",
    })
    c_nv_pic_pr = etree.SubElement(nv_pic_pr, f"{PIC}cNvPicPr")
    blip_fill = etree.SubElement(pic, f"{PIC}blipFill")
    etree.SubElement(blip_fill, f"{A}blip", attrib={
        f"{R}embed": rId,
    })
    stretch = etree.SubElement(blip_fill, f"{A}stretch")
    etree.SubElement(stretch, f"{A}fillRect")

    sp_pr = etree.SubElement(pic, f"{PIC}spPr")
    xfrm = etree.SubElement(sp_pr, f"{A}xfrm")
    etree.SubElement(xfrm, f"{A}off", attrib={"x": "0", "y": "0"})
    etree.SubElement(xfrm, f"{A}ext", attrib={"cx": str(width), "cy": str(height)})
    etree.SubElement(sp_pr, f"{A}prstGeom", attrib={"prst": "rect"})
    etree.SubElement(sp_pr, f"{A}noFill")

    # Detach from dummy parent
    if drawing.getparent() is not None:
        drawing.getparent().remove(drawing)

    return drawing


# ── Heading operations ──────────────────────────────────────────


def set_heading(
    para: etree._Element,
    level: int = 1,
    text: Optional[str] = None,
) -> dict:
    """Set a paragraph as a heading.

    Args:
        para: The w:p paragraph element.
        level: Heading level (1-9).
        text: Optional new text for the heading.

    Returns:
        dict with status info.
    """
    if level < 1 or level > 9:
        return {"status": "error", "message": "Heading level must be 1-9"}

    set_style(para, f"Heading {level}")

    if text is not None:
        # Clear existing text and set new text
        runs = list(para.iter(tag("r")))
        for r in runs:
            r.getparent().remove(r)

        r = make_element("r")
        t = make_element("t", text=text)
        if text and (text[0].isspace() or text[-1].isspace()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        para.append(r)

    return {"status": "ok", "message": f"Set as Heading {level}"}


# ── Bullet and numbering ────────────────────────────────────────

_LIST_PREFIXES = {
    "bullet": "• ",
    "number": "1. ",
}
_LIST_INDENT = {"bullet": 720, "number": 720}  # twips


def set_list_format(
    para: etree._Element,
    list_type: str = "bullet",  # "bullet", "number", or "none"
    level: int = 0,
) -> dict:
    """Apply or remove bullet/numbered list formatting on a paragraph.

    list_type="none" removes any existing list indent and prefix.
    """
    ppr = ensure_pPr(para)

    if list_type == "none":
        return _remove_list_format(para, ppr)

    # Apply list indent
    base = _LIST_INDENT.get(list_type, 720)
    indent_attrib = {
        W + "left": str(base + level * 360),
        W + "hanging": "360",
    }
    set_or_replace(ppr, "ind", attrib=indent_attrib)

    # Prepend list prefix
    prefix = _LIST_PREFIXES.get(list_type, "• ")
    runs = children(para, "r")
    if runs:
        t = child(runs[0], "t")
        if t is not None and t.text and not t.text.startswith(prefix):
            t.text = prefix + t.text
    else:
        r = make_element("r")
        r.append(make_element("t", text=prefix))
        para.append(r)

    return {"status": "ok", "message": f"Applied {list_type} list at level {level}"}


def _remove_list_format(para: etree._Element, ppr: etree._Element) -> dict:
    """Strip list indentation and prefix from a paragraph."""
    changes = []

    ind_el = child(ppr, "ind")
    if ind_el is not None:
        ppr.remove(ind_el)
        changes.append("indentation")

    runs = children(para, "r")
    if runs:
        t = child(runs[0], "t")
        if t is not None and t.text:
            for prefix in _LIST_PREFIXES.values():
                if t.text.startswith(prefix):
                    t.text = t.text[len(prefix):]
                    changes.append("prefix removed")
                    break

    msg = ", ".join(changes) if changes else "no list formatting"
    return {"status": "ok", "message": f"Removed list: {msg}"}
