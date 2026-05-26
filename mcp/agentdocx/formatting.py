"""Font and paragraph formatting operations for docx documents.

Uses w:rPr (run properties) for font formatting and w:pPr for paragraph formatting.
"""

from __future__ import annotations

from typing import Optional

from lxml import etree

from .oxml_helpers import (
    NSMAP, W,
    tag, make_element, child, children,
    text_content, ensure_rPr, ensure_pPr,
    set_or_replace, set_property,
    split_run_element_at, insert_after,
)


# ── Font formatting ─────────────────────────────────────────────


def set_font(
    para: etree._Element,
    start_offset: Optional[int] = None,
    end_offset: Optional[int] = None,
    *,
    name: Optional[str] = None,
    name_ascii: Optional[str] = None,
    name_east_asia: Optional[str] = None,
    size: Optional[int] = None,  # in half-points (12pt = 24)
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[str] = None,  # "single", "double", "none"
    color: Optional[str] = None,  # hex color e.g., "FF0000"
    highlight: Optional[str] = None,  # "yellow", "red", etc.
    strikethrough: Optional[bool] = None,
    double_strikethrough: Optional[bool] = None,
    small_caps: Optional[bool] = None,
    all_caps: Optional[bool] = None,
    superscript: Optional[bool] = None,
    subscript: Optional[bool] = None,
) -> dict:
    """Set font properties on a range of text.

    If start_offset and end_offset are None, applies to the entire paragraph.
    Otherwise applies to the specified text range.

    Args:
        para: The w:p paragraph element.
        start_offset: Start character offset (inclusive).
        end_offset: End character offset (exclusive).
        name: Font name (sets ascii, hAnsi, and eastAsia).
        name_ascii: Western font name.
        name_east_asia: East Asian font name.
        size: Font size in half-points (24 = 12pt, 28 = 14pt, etc.).
        bold: Bold on/off.
        italic: Italic on/off.
        underline: Underline style ("single", "double", "none").
        color: Text color as hex RGB (e.g., "FF0000" for red).
        highlight: Highlight color name (e.g., "yellow", "cyan", "red").
        strikethrough: Strikethrough on/off.
        double_strikethrough: Double strikethrough on/off.
        small_caps: Small caps on/off.
        all_caps: All caps on/off.
        superscript: Superscript on/off.
        subscript: Subscript on/off.

    Returns:
        dict with status info.
    """
    runs = _get_runs_for_range(para, start_offset, end_offset)
    if not runs:
        return {"status": "noop", "message": "No runs found in range"}

    changes = []

    for run in runs:
        rpr = ensure_rPr(run)

        if name:
            # Set all font slots
            set_or_replace(rpr, "rFonts", attrib={
                W + "ascii": name,
                W + "hAnsi": name,
                W + "eastAsia": name,
            })
            changes.append(f"font={name}")
        else:
            if name_ascii:
                _set_rFonts_attribute(rpr, "ascii", name_ascii)
                changes.append(f"font_ascii={name_ascii}")
            if name_east_asia:
                _set_rFonts_attribute(rpr, "eastAsia", name_east_asia)
                changes.append(f"font_eastAsia={name_east_asia}")

        if size is not None:
            set_property(rpr, "sz", str(size))
            set_property(rpr, "szCs", str(size))
            changes.append(f"size={size}")

        if bold is not None:
            if bold:
                set_or_replace(rpr, "b")
            else:
                _remove_property(rpr, "b")
            changes.append(f"bold={bold}")

        if italic is not None:
            if italic:
                set_or_replace(rpr, "i")
            else:
                _remove_property(rpr, "i")
            changes.append(f"italic={italic}")

        if underline is not None:
            if underline == "none":
                _remove_property(rpr, "u")
            else:
                set_property(rpr, "u", underline)
            changes.append(f"underline={underline}")

        if color is not None:
            set_property(rpr, "color", color)
            changes.append(f"color={color}")

        if highlight is not None:
            if highlight == "none":
                _remove_property(rpr, "highlight")
            else:
                set_property(rpr, "highlight", highlight)
            changes.append(f"highlight={highlight}")

        if strikethrough is not None:
            if strikethrough:
                set_or_replace(rpr, "strike")
            else:
                _remove_property(rpr, "strike")
            changes.append(f"strikethrough={strikethrough}")

        if double_strikethrough is not None:
            if double_strikethrough:
                set_or_replace(rpr, "dstrike")
            else:
                _remove_property(rpr, "dstrike")
            changes.append(f"double_strikethrough={double_strikethrough}")

        if small_caps is not None:
            if small_caps:
                set_or_replace(rpr, "smallCaps")
            else:
                _remove_property(rpr, "smallCaps")
            changes.append(f"small_caps={small_caps}")

        if all_caps is not None:
            if all_caps:
                set_or_replace(rpr, "allCaps")
            else:
                _remove_property(rpr, "allCaps")
            changes.append(f"all_caps={all_caps}")

        if superscript:
            set_property(rpr, "vertAlign", "superscript")
            changes.append("superscript")
        elif subscript:
            set_property(rpr, "vertAlign", "subscript")
            changes.append("subscript")
        elif superscript is not None or subscript is not None:
            _remove_property(rpr, "vertAlign")

    return {
        "status": "ok",
        "message": f"Applied font formatting: {', '.join(changes)}",
        "runs_affected": len(runs),
    }


# ── Paragraph formatting ────────────────────────────────────────


def set_paragraph_format(
    para: etree._Element,
    *,
    alignment: Optional[str] = None,  # "left", "right", "center", "both"
    left_indent: Optional[int] = None,  # in twips (1/20 of a point)
    right_indent: Optional[int] = None,
    first_line_indent: Optional[int] = None,
    hanging_indent: Optional[int] = None,
    space_before: Optional[int] = None,  # in twips
    space_after: Optional[int] = None,
    line_spacing: Optional[float] = None,  # multiplier (1.0, 1.5, 2.0)
    line_spacing_exact: Optional[int] = None,  # exact line spacing in twips
    outline_level: Optional[int] = None,  # 0-9 (0 = body, 1-9 = heading levels)
    page_break_before: Optional[bool] = None,
    keep_lines_together: Optional[bool] = None,
    keep_with_next: Optional[bool] = None,
) -> dict:
    """Set paragraph formatting properties.

    Args:
        para: The w:p paragraph element.
        alignment: Text alignment ("left", "right", "center", "both").
        left_indent: Left indent in twips.
        right_indent: Right indent in twips.
        first_line_indent: First line indent in twips.
        hanging_indent: Hanging indent in twips.
        space_before: Space before paragraph in twips.
        space_after: Space after paragraph in twips.
        line_spacing: Line spacing multiplier (1.0, 1.5, 2.0, etc.).
        line_spacing_exact: Exact line spacing in twips (overrides line_spacing).
        outline_level: Outline level 0-9.
        page_break_before: Insert page break before paragraph.
        keep_lines_together: Keep all lines of paragraph on same page.
        keep_with_next: Keep paragraph with next paragraph.

    Returns:
        dict with status info.
    """
    ppr = ensure_pPr(para)
    changes = []

    if alignment:
        alignment_map = {
            "left": "left",
            "right": "right",
            "center": "center",
            "both": "both",
            "justify": "both",
            "distribute": "both",
        }
        val = alignment_map.get(alignment.lower(), alignment)
        set_property(ppr, "jc", val)
        changes.append(f"alignment={val}")

    if any(x is not None for x in [left_indent, right_indent, first_line_indent, hanging_indent]):
        ind_attrib = {}
        if left_indent is not None:
            ind_attrib[W + "left"] = str(left_indent)
        if right_indent is not None:
            ind_attrib[W + "right"] = str(right_indent)
        if first_line_indent is not None:
            ind_attrib[W + "firstLine"] = str(first_line_indent)
        if hanging_indent is not None:
            ind_attrib[W + "hanging"] = str(hanging_indent)
        set_or_replace(ppr, "ind", attrib=ind_attrib)
        changes.append("indentation")

    if any(x is not None for x in [space_before, space_after, line_spacing, line_spacing_exact]):
        spacing_attrib = {}
        if space_before is not None:
            spacing_attrib[W + "before"] = str(space_before)
        if space_after is not None:
            spacing_attrib[W + "after"] = str(space_after)

        if line_spacing_exact is not None:
            spacing_attrib[W + "line"] = str(line_spacing_exact)
            spacing_attrib[W + "lineRule"] = "exact"
        elif line_spacing is not None:
            # Line spacing multiplier * 240
            spacing_attrib[W + "line"] = str(int(line_spacing * 240))
            spacing_attrib[W + "lineRule"] = "auto"

        set_or_replace(ppr, "spacing", attrib=spacing_attrib)
        changes.append("spacing")

    if outline_level is not None:
        set_property(ppr, "outlineLvl", str(outline_level))
        changes.append(f"outline_level={outline_level}")

    if page_break_before is not None:
        if page_break_before:
            set_or_replace(ppr, "pageBreakBefore")
        else:
            _remove_property(ppr, "pageBreakBefore")
        changes.append(f"page_break_before={page_break_before}")

    if keep_lines_together is not None:
        if keep_lines_together:
            set_or_replace(ppr, "keepLines")
        else:
            _remove_property(ppr, "keepLines")
        changes.append(f"keep_lines_together={keep_lines_together}")

    if keep_with_next is not None:
        if keep_with_next:
            set_or_replace(ppr, "keepNext")
        else:
            _remove_property(ppr, "keepNext")
        changes.append(f"keep_with_next={keep_with_next}")

    return {
        "status": "ok",
        "message": f"Applied paragraph formatting: {', '.join(changes) if changes else 'none'}",
    }


def set_style(para: etree._Element, style_name: str) -> dict:
    """Apply a named paragraph style (e.g., 'Heading 1', 'Normal').

    Args:
        para: The w:p paragraph element.
        style_name: Name of the style to apply.

    Returns:
        dict with status info.
    """
    ppr = ensure_pPr(para)
    set_or_replace(ppr, "pStyle", attrib={W + "val": style_name})
    return {"status": "ok", "message": f"Applied style '{style_name}'"}


# ── Internal helpers ────────────────────────────────────────────


def _get_runs_for_range(
    para: etree._Element,
    start_offset: Optional[int],
    end_offset: Optional[int],
) -> list[etree._Element]:
    """Get the w:r elements that overlap with the given text range.

    If start_offset and end_offset are None, returns all runs in the paragraph.
    Note: this doesn't modify the paragraph structure for partial ranges.
    """
    if start_offset is None and end_offset is None:
        return _get_all_runs(para)

    if start_offset is None:
        start_offset = 0
    if end_offset is None:
        end_offset = len(text_content(para))

    runs = []
    accumulated = 0

    for child_el in para:
        if child_el.tag == tag("r"):
            t = text_content(child_el)
            run_end = accumulated + len(t)
            if accumulated < end_offset and run_end > start_offset:
                runs.append(child_el)
            accumulated += len(t)

        elif child_el.tag == tag("ins"):
            for r in children(child_el, "r"):
                t = text_content(r)
                run_end = accumulated + len(t)
                if accumulated < end_offset and run_end > start_offset:
                    runs.append(r)
                accumulated += len(t)

        elif child_el.tag == tag("del"):
            for r in children(child_el, "r"):
                dt = "".join(d.text or "" for d in r.iter(tag("delText")))
                run_end = accumulated + len(dt)
                if accumulated < end_offset and run_end > start_offset:
                    runs.append(r)
                accumulated += len(dt)

    return runs


def _get_all_runs(para: etree._Element) -> list[etree._Element]:
    """Get all w:r elements from a paragraph, including those in w:ins/w:del."""
    runs = []
    for child_el in para:
        if child_el.tag == tag("r"):
            runs.append(child_el)
        elif child_el.tag == tag("ins"):
            runs.extend(children(child_el, "r"))
        elif child_el.tag == tag("del"):
            runs.extend(children(child_el, "r"))
    return runs


def _set_rFonts_attribute(rpr: etree._Element, attr_name: str, value: str):
    """Set a specific attribute on the w:rFonts element, creating it if needed."""
    fonts = child(rpr, "rFonts")
    if fonts is None:
        fonts = make_element("rFonts")
        rpr.append(fonts)
    fonts.set(W + attr_name, value)


def _remove_property(parent: etree._Element, tag_name: str):
    """Remove a property element from parent."""
    for el in children(parent, tag_name):
        parent.remove(el)
