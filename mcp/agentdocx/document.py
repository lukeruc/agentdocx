"""Document wrapper combining python-docx with direct OOXML access."""

from __future__ import annotations

import copy
import uuid
import zipfile
import shutil
import tempfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Optional

from docx import Document as DocxDocument_
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from lxml import etree

from .oxml_helpers import (
    W, NSMAP,
    tag, make_element, find, child, children,
    text_content, get_paragraph_runs, find_text_position,
    get_next_comment_id, get_next_revision_id,
    ensure_pPr, ensure_rPr,
)


# ── Data classes ────────────────────────────────────────────────


@dataclass
class ParaInfo:
    """Information about a paragraph in the document."""
    index: int
    text: str
    style: str = ""
    is_heading: bool = False
    heading_level: int = 0


@dataclass
class DocInfo:
    """Document metadata returned by get_info."""
    path: str
    paragraph_count: int
    paragraphs: list[ParaInfo] = field(default_factory=list)


# ── Document class ──────────────────────────────────────────────


class DocxDocument:
    """Wraps a python-docx Document with direct OOXML access for advanced features.

    Provides:
    - Open/save .docx files
    - Access to paragraphs and runs via python-docx API
    - Direct lxml element access for OOXML manipulation
    - Comments part management
    - Track changes support
    """

    def __init__(self, path: str | Path):
        self._path = Path(path)
        if not self._path.exists():
            raise FileNotFoundError(f"Document not found: {path}")

        self._docx = DocxDocument_(str(self._path))
        self._body: etree._Element = self._docx.element.body

        # Comments part
        self._comments: Optional[etree._Element] = None
        self._comments_part = None
        self._load_comments()

    # ── Properties ──────────────────────────────────────────

    @property
    def path(self) -> Path:
        return self._path

    @property
    def docx(self) -> DocxDocument_:
        return self._docx

    @property
    def body(self) -> etree._Element:
        return self._body

    @property
    def paragraphs(self):
        """Return paragraph elements directly."""
        return children(self._body, "p")

    @property
    def paragraph_count(self) -> int:
        return len(self.paragraphs)

    # ── Comments ────────────────────────────────────────────

    def _load_comments(self):
        """Load the comments part if it exists."""
        try:
            comments_part = self._docx.part.part_related_by(RT.COMMENTS)
        except (KeyError, ValueError):
            self._comments = None
            self._comments_part = None
            return

        if comments_part is not None:
            self._comments_part = comments_part
            self._comments = etree.fromstring(comments_part.blob)

    def _create_comments_element(self):
        """Create a fresh w:comments element (without adding to package yet)."""
        comments_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:comments xmlns:w="{NSMAP["w"]}"'
            f' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
            f' mc:Ignorable="w14 w15 w16cex w16cid"'
            f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'</w:comments>'
        )
        self._comments = etree.fromstring(comments_xml.encode("utf-8"))

    @property
    def comments_element(self) -> Optional[etree._Element]:
        return self._comments

    def comments_count(self) -> int:
        if self._comments is None:
            return 0
        return len([c for c in self._comments if c.tag == tag("comment")])

    # ── Text access ─────────────────────────────────────────

    def get_paragraph_text(self, index: int) -> str:
        """Get the text of a paragraph by index."""
        paras = self.paragraphs
        if index < 0 or index >= len(paras):
            raise IndexError(f"Paragraph index {index} out of range (0-{len(paras) - 1})")
        return text_content(paras[index])

    def get_full_text(self) -> str:
        """Get the full text of the document with paragraph breaks."""
        texts = []
        for p in self.paragraphs:
            texts.append(text_content(p))
        return "\n".join(texts)

    def search_text(self, query: str) -> list[dict]:
        """Search for text in the document.

        Returns list of dicts with paragraph_index, offset, and context.
        """
        results = []
        for pi, para in enumerate(self.paragraphs):
            para_text = text_content(para)
            start = 0
            while True:
                idx = para_text.find(query, start)
                if idx == -1:
                    break
                # Get some context around the match
                ctx_start = max(0, idx - 20)
                ctx_end = min(len(para_text), idx + len(query) + 20)
                context = para_text[ctx_start:ctx_end]
                results.append({
                    "paragraph_index": pi,
                    "offset": idx,
                    "match_length": len(query),
                    "context": context,
                    "paragraph_text": para_text,
                })
                start = idx + 1
        return results

    def get_info(self) -> DocInfo:
        """Get document metadata and structure info."""
        paras_info = []
        for i, p in enumerate(self.paragraphs):
            txt = text_content(p)
            ppr = child(p, "pPr")
            style = ""
            is_heading = False
            heading_level = 0

            if ppr is not None:
                pstyle = child(ppr, "pStyle")
                if pstyle is not None:
                    style = pstyle.get(W + "val", "")

            if style and "Heading" in style or "heading" in style:
                is_heading = True
                # Extract level number
                for c in style:
                    if c.isdigit():
                        heading_level = int(c)
                        break

            paras_info.append(ParaInfo(
                index=i,
                text=txt[:200] + ("..." if len(txt) > 200 else ""),
                style=style,
                is_heading=is_heading,
                heading_level=heading_level,
            ))

        return DocInfo(
            path=str(self._path),
            paragraph_count=len(paras_info),
            paragraphs=paras_info,
        )

    # ── Save ────────────────────────────────────────────────

    def save(self, path: str | Path | None = None):
        """Save the document, optionally to a new path.

        If path is provided, saves a copy to that path.
        Otherwise, saves in place.
        """
        save_path = Path(path) if path else self._path

        # If we have comments, serialize and prepare for injection
        if self._comments is not None:
            self._save_comments()

        self._docx.save(str(save_path))

        # Inject comments part if needed (post-save ZIP manipulation)
        if hasattr(self, '_comments_pending'):
            self._inject_comments_after_save(save_path)

        # Update path reference if saved to new location
        if path:
            self._path = save_path

        return str(save_path)

    def save_to_bytes(self) -> bytes:
        """Save the document to bytes.

        Note: Comments injection is not supported in bytes mode.
        Save to a file path for full comment support.
        """
        if self._comments is not None:
            self._save_comments()
        buf = BytesIO()
        self._docx.save(buf)
        return buf.getvalue()

    def _save_comments(self):
        """Serialize comments into the document package."""
        if self._comments is None:
            return

        # Check if there are actual comment elements
        comment_count = len([c for c in self._comments if c.tag == tag("comment")])
        if comment_count == 0:
            return

        comments_xml = etree.tostring(
            self._comments, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # Try to update existing comments part
        try:
            comments_part = self._docx.part.part_related_by(RT.COMMENTS)
            comments_part._blob = comments_xml
            self._comments_saved = comments_xml
            return
        except (KeyError, ValueError):
            pass

        # Comments part doesn't exist in the package - it will be injected
        # after the initial save via _inject_comments_into_package
        self._comments_pending = comments_xml

    def _inject_comments_after_save(self, save_path: Path):
        """Post-save injection of comments part into the docx ZIP.

        This directly manipulates the docx ZIP archive to add the comments part,
        update [Content_Types].xml, and add the relationship to the document part.
        """
        if not hasattr(self, '_comments_pending'):
            return

        comments_xml = self._comments_pending
        del self._comments_pending

        import zipfile
        import os

        tmp_path = save_path.with_suffix(".tmp.docx")

        with zipfile.ZipFile(str(save_path), "r") as zin:
            with zipfile.ZipFile(str(tmp_path), "w", zipfile.ZIP_DEFLATED) as zout:
                # Add comments.xml to the word/ directory
                zout.writestr("word/comments.xml", comments_xml)

                # Copy all other entries
                for item in zin.infolist():
                    if item.filename == "word/comments.xml":
                        continue  # Already added

                    data = zin.read(item.filename)

                    # Update [Content_Types].xml to add comments content type
                    if item.filename == "[Content_Types].xml":
                        content_types = etree.fromstring(data)
                        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
                        ct_tag = f"{{{ct_ns}}}"

                        # Check if comments content type already exists
                        has_comments_ct = False
                        for ov in content_types:
                            if ov.tag == f"{ct_tag}Override" and ov.get("PartName") == "/word/comments.xml":
                                has_comments_ct = True
                                break

                        if not has_comments_ct:
                            ct_override = etree.SubElement(content_types, f"{ct_tag}Override")
                            ct_override.set("PartName", "/word/comments.xml")
                            ct_override.set(
                                "ContentType",
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
                            )
                            data = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone=True)

                    # Update word/_rels/document.xml.rels to add comments relationship
                    if item.filename == "word/_rels/document.xml.rels":
                        rels_xml = etree.fromstring(data)
                        rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

                        # Find next available rId
                        max_rid = 0
                        for rel in rels_xml:
                            rid = rel.get("Id", "")
                            if rid.startswith("rId"):
                                try:
                                    num = int(rid[3:])
                                    max_rid = max(max_rid, num)
                                except ValueError:
                                    pass

                        new_rid = f"rId{max_rid + 1}"

                        # Check if comments relationship already exists
                        has_comments_rel = False
                        for rel in rels_xml:
                            if rel.get("Type") == RT.COMMENTS:
                                has_comments_rel = True
                                break

                        if not has_comments_rel:
                            comments_rel = etree.SubElement(rels_xml, "Relationship")
                            comments_rel.set("Id", new_rid)
                            comments_rel.set("Type", RT.COMMENTS)
                            comments_rel.set("Target", "comments.xml")
                            data = etree.tostring(rels_xml, xml_declaration=True, encoding="UTF-8", standalone=True)

                    zout.writestr(item, data)

        # Replace original with modified version
        save_path.unlink()
        shutil.move(str(tmp_path), str(save_path))

    # ── Paragraph operations ────────────────────────────────

    def add_paragraph(self, text: str = "", style: str | None = None) -> int:
        """Add a new paragraph to the end of the document.

        Returns the index of the new paragraph.
        """
        p = make_element("p")
        ppr = make_element("pPr")
        p.append(ppr)

        if style:
            from .oxml_helpers import set_or_replace
            set_or_replace(ppr, "pStyle", attrib={W + "val": style})

        if text:
            r = make_element("r")
            t = make_element("t", text=text)
            r.append(t)
            p.append(r)

        self._body.append(p)
        return len(self.paragraphs) - 1

    def insert_paragraph(self, index: int, text: str = "", style: str | None = None) -> int:
        """Insert a paragraph at the given index.

        Returns the index of the new paragraph.
        """
        paras = self.paragraphs
        if index < 0 or index > len(paras):
            raise IndexError(f"Paragraph index {index} out of range (0-{len(paras)})")

        p = make_element("p")
        ppr = make_element("pPr")
        p.append(ppr)

        if style:
            from .oxml_helpers import set_or_replace
            set_or_replace(ppr, "pStyle", attrib={W + "val": style})

        if text:
            r = make_element("r")
            t = make_element("t", text=text)
            r.append(t)
            p.append(r)

        if index == len(paras):
            self._body.append(p)
        else:
            target = paras[index]
            from .oxml_helpers import insert_before
            insert_before(target, p)

        return index

    def delete_paragraph(self, index: int):
        """Delete a paragraph by index."""
        paras = self.paragraphs
        if index < 0 or index >= len(paras):
            raise IndexError(f"Paragraph index {index} out of range (0-{len(paras) - 1})")
        self._body.remove(paras[index])
