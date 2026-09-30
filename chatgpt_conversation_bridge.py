#!/usr/bin/env python3
"""
ChatGPT Conversation Bridge

Version: 2.0.0
Author: hummbugg
Copyright (c) 2026 hummbugg

Preserve a public ChatGPT shared conversation as a permanent ZIP archive
and a DOCX document for conversation continuation and archival.

Usage:
    python chatgpt_conversation_bridge.py "https://chatgpt.com/share/<share-id>"
    python chatgpt_conversation_bridge.py "archive/CONVERSATION_NAME.zip"

The script uses its own directory as the working directory.

For a public ChatGPT Share URL, the conversation JSON is retrieved from
ChatGPT and supported uploaded images are downloaded when available. A
permanent ZIP archive is created containing conversation.json and the
successfully retrieved images. The DOCX is then generated from that data.

For an existing ChatGPT Conversation Bridge ZIP archive, the DOCX is
regenerated directly from the archive without modifying the permanent ZIP.

Archives are written to:
    archive/CONVERSATION_NAME.zip

DOCX output is written to:
    docx/CONVERSATION_NAME.docx

An existing output DOCX is always overwritten. An existing permanent
conversation archive is not overwritten.

Requires:
    Python 3.10+
    curl_cffi

A network connection is required when processing a public ChatGPT Share URL.
Regenerating a DOCX from an existing archive is an offline operation.
"""

from __future__ import annotations

import argparse
import html as html_lib
import zipfile
import datetime as _dt
import time
from zoneinfo import ZoneInfo
from xml.etree import ElementTree as ET
import json
import re
import sys
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode
from curl_cffi import requests

class ExtractionError(RuntimeError):
    pass


def managed_directories(script_path: Path) -> tuple[Path, Path, Path]:
    """Return the working, archive, and DOCX directories."""
    working_dir = script_path.resolve().parent
    archive_dir = working_dir / "archive"
    docx_dir = working_dir / "docx"
    return working_dir, archive_dir, docx_dir


def png_dimensions(data: bytes | None) -> tuple[int, int] | None:
    """Validate PNG structure and return width and height from the IHDR chunk."""
    if data is None or len(data) < 24:
        return None

    if data[:8] != b"\x89PNG\r\n\x1a\n":
        return None

    if data[12:16] != b"IHDR":
        return None

    width = int.from_bytes(data[16:20], "big")
    height = int.from_bytes(data[20:24], "big")

    if width <= 0 or height <= 0:
        return None

    return width, height

def jpeg_dimensions(data: bytes | None) -> tuple[int, int] | None:
    """Return JPEG width and height from a valid Start Of Frame marker."""
    if data is None or len(data) < 4:
        return None

    if data[:2] != b"\xFF\xD8":
        return None

    i = 2

    while i < len(data):
        # JPEG markers begin with one or more 0xFF bytes.
        if data[i] != 0xFF:
            i += 1
            continue

        while i < len(data) and data[i] == 0xFF:
            i += 1

        if i >= len(data):
            return None

        marker = data[i]
        i += 1

        # Standalone markers have no length field.
        if marker in {0x01, *range(0xD0, 0xDA)}:
            continue

        # End Of Image or Start Of Scan before a SOF means no usable dimensions.
        if marker in {0xD9, 0xDA}:
            return None

        if i + 2 > len(data):
            return None

        segment_length = int.from_bytes(data[i:i + 2], "big")

        if segment_length < 2 or i + segment_length > len(data):
            return None

        # SOF markers that contain image dimensions.
        if marker in {
            0xC0, 0xC1, 0xC2, 0xC3,
            0xC5, 0xC6, 0xC7,
            0xC9, 0xCA, 0xCB,
            0xCD, 0xCE, 0xCF,
        }:
            if segment_length < 7:
                return None

            height = int.from_bytes(data[i + 3:i + 5], "big")
            width = int.from_bytes(data[i + 5:i + 7], "big")

            if width <= 0 or height <= 0:
                return None

            return width, height

        i += segment_length

    return None

def gif_dimensions(data: bytes | None) -> tuple[int, int] | None:
    """Validate GIF structure and return logical screen width and height."""
    if data is None or len(data) < 10:
        return None

    if data[:6] not in {b"GIF87a", b"GIF89a"}:
        return None

    width = int.from_bytes(data[6:8], "little")
    height = int.from_bytes(data[8:10], "little")

    if width <= 0 or height <= 0:
        return None

    return width, height

def webp_dimensions(data: bytes | None) -> tuple[int, int] | None:
    """Validate WebP structure and return image width and height."""
    if data is None or len(data) < 16:
        return None

    if data[:4] != b"RIFF" or data[8:12] != b"WEBP":
        return None

    chunk_type = data[12:16]

    if chunk_type == b"VP8X":
        if len(data) < 30:
            return None

        width = 1 + int.from_bytes(data[24:27], "little")
        height = 1 + int.from_bytes(data[27:30], "little")

    elif chunk_type == b"VP8L":
        if len(data) < 25 or data[20] != 0x2F:
            return None

        bits = int.from_bytes(data[21:25], "little")
        width = (bits & 0x3FFF) + 1
        height = ((bits >> 14) & 0x3FFF) + 1

    elif chunk_type == b"VP8 ":
        if len(data) < 30:
            return None

        if data[23:26] != b"\x9D\x01\x2A":
            return None

        width = int.from_bytes(data[26:28], "little") & 0x3FFF
        height = int.from_bytes(data[28:30], "little") & 0x3FFF

    else:
        return None

    if width <= 0 or height <= 0:
        return None

    return width, height

def image_info(
    data: bytes | None,
) -> dict[str, Any] | None:
    """Identify a supported saved image from its actual bytes."""
    if data is None:
        return None

    dimensions = png_dimensions(data)
    if dimensions is not None:
        return {
            "format": "png",
            "extension": "png",
            "content_type": "image/png",
            "dimensions": dimensions,
        }

    dimensions = jpeg_dimensions(data)
    if dimensions is not None:
        return {
            "format": "jpeg",
            "extension": "jpg",
            "content_type": "image/jpeg",
            "dimensions": dimensions,
        }

    dimensions = gif_dimensions(data)
    if dimensions is not None:
        return {
            "format": "gif",
            "extension": "gif",
            "content_type": "image/gif",
            "dimensions": dimensions,
        }

    dimensions = webp_dimensions(data)
    if dimensions is not None:
        return {
            "format": "webp",
            "extension": "webp",
            "content_type": "image/webp",
            "dimensions": dimensions,
        }        

    return None

def image_display_size_emu(dimensions: tuple[int, int] | None) -> tuple[int, int] | None:
    """Return image display size in EMU using 96 DPI, capped at 7 inches wide."""

    if dimensions is None:
        return None

    width_px, height_px = dimensions
    emu_per_inch = 914400
    pixels_per_inch = 96
    max_width_emu = 6400800  # 7.0 inches

    width_emu = round(width_px * emu_per_inch / pixels_per_inch)
    height_emu = round(height_px * emu_per_inch / pixels_per_inch)

    if width_emu > max_width_emu:
        scale = max_width_emu / width_emu
        width_emu = max_width_emu
        height_emu = round(height_emu * scale)

    return width_emu, height_emu

def validate_chain(nodes: list[dict[str, Any]], current_node: Any) -> None:
    if not nodes:
        raise ExtractionError("linear_conversation is empty.")

    ids: list[str] = []
    for n, node in enumerate(nodes):
        if not isinstance(node, dict):
            raise ExtractionError(f"Conversation node {n} is not an object.")
        node_id = node.get("id")
        if not isinstance(node_id, str) or not node_id:
            raise ExtractionError(f"Conversation node {n} has no valid id.")
        ids.append(node_id)

    if ids[-1] != current_node:
        raise ExtractionError("The final linear_conversation node is not current_node.")

    for i in range(1, len(nodes)):
        expected_parent = ids[i - 1]
        actual_parent = nodes[i].get("parent")
        if actual_parent != expected_parent:
            raise ExtractionError(
                f"Conversation chain is broken at node {i}: parent does not match."
            )
        children = nodes[i - 1].get("children")
        if isinstance(children, list) and ids[i] not in children:
            raise ExtractionError(
                f"Conversation chain is broken at node {i - 1}: next node is not a child."
            )


def iter_strings(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        yield value
    elif isinstance(value, list):
        for item in value:
            yield from iter_strings(item)
    elif isinstance(value, dict):
        for item in value.values():
            yield from iter_strings(item)


def contains_asset_pointer(value: Any) -> bool:
    for text in iter_strings(value):
        if text.startswith("sediment://"):
            return True
    if isinstance(value, dict):
        if value.get("content_type") in {"image_asset_pointer", "file_asset_pointer"}:
            return True
        if "asset_pointer" in value:
            return True
        return any(contains_asset_pointer(v) for v in value.values())
    if isinstance(value, list):
        return any(contains_asset_pointer(v) for v in value)
    return False


def upload_info(part: dict[str, Any], metadata: Any) -> dict[str, Any]:
    """Match an upload part to its same-message attachment metadata."""
    asset_pointer = part.get("asset_pointer")
    file_id = None

    if isinstance(asset_pointer, str) and asset_pointer.startswith("sediment://"):
        file_id = asset_pointer[len("sediment://"):].split("?", 1)[0]

    attachment = None
    if isinstance(metadata, dict):
        attachments = metadata.get("attachments")
        if isinstance(attachments, list):
            for item in attachments:
                if isinstance(item, dict) and item.get("id") == file_id:
                    attachment = item
                    break

    return {
        "content_type": part.get("content_type") or part.get("type"),
        "asset_pointer": asset_pointer,
        "filename": attachment.get("name") if attachment else None,
        "mime_type": attachment.get("mime_type") if attachment else None,
        "size": attachment.get("size") if attachment else part.get("size_bytes"),
        "width": attachment.get("width") if attachment else part.get("width"),
        "height": attachment.get("height") if attachment else part.get("height"),
    }


def metadata_attachment_info(attachment: dict[str, Any]) -> dict[str, Any]:
    """Build an upload record from attachment metadata alone."""
    return {
        "content_type": None,
        "asset_pointer": None,
        "filename": attachment.get("name"),
        "mime_type": attachment.get("mime_type"),
        "size": attachment.get("size"),
        "width": attachment.get("width"),
        "height": attachment.get("height"),
    }


def unavailable_upload_text(upload: dict[str, Any]) -> str:
    """Build placeholder text for an upload whose payload is unavailable."""
    is_image = upload.get("content_type") == "image_asset_pointer"
    label = "Image" if is_image else "Attachment"

    filename = upload.get("filename")
    if isinstance(filename, str) and filename.strip():
        return f"[{label} unavailable: {filename.strip()}]"

    return f"[{label} unavailable]"


def extract_text_and_uploads(content: Any, metadata: Any) -> tuple[str, list[dict[str, Any]]]:
    """Extract user-visible text and preserve upload attachment information."""
    if not isinstance(content, dict):
        return "", []

    ctype = content.get("content_type")
    parts = content.get("parts")
    texts: list[str] = []
    uploads: list[dict[str, Any]] = []

    if isinstance(parts, list):
        for part in parts:
            if isinstance(part, str):
                if part:
                    texts.append(part)
            elif isinstance(part, dict):
                ptype = part.get("content_type") or part.get("type")
                if ptype in {"image_asset_pointer", "file_asset_pointer", "image"} or contains_asset_pointer(part):
                    uploads.append(upload_info(part, metadata))
                else:
                    # Some content parts wrap their visible text in a text field.
                    ptext = part.get("text")
                    if isinstance(ptext, str) and ptext:
                        texts.append(ptext)
    else:
        text = content.get("text")
        if isinstance(text, str) and text:
            texts.append(text)

   
    # A pointer can occasionally live outside parts.
    if not uploads and contains_asset_pointer(content):
        uploads.append(upload_info(content, metadata))

    represented_file_ids: set[str] = set()
    for upload in uploads:
        asset_pointer = upload.get("asset_pointer")
        if isinstance(asset_pointer, str) and asset_pointer.startswith("sediment://"):
            represented_file_ids.add(
                asset_pointer[len("sediment://"):].split("?", 1)[0]
            )

    if isinstance(metadata, dict):
        attachments = metadata.get("attachments")
        if isinstance(attachments, list):
            for attachment in attachments:
                if not isinstance(attachment, dict):
                    continue
                attachment_id = attachment.get("id")
                if attachment_id not in represented_file_ids:
                    uploads.append(metadata_attachment_info(attachment))

    return "\n\n".join(texts).strip(), uploads


URL_RE = re.compile(r"https?://[^\s<>\]\[\"']+")


def sanitize_reference_url(url: str) -> str:
    """Return a public reference URL with ChatGPT referral attribution removed."""
    url = html_lib.unescape(url.strip())
    try:
        parts = urlsplit(url)
    except ValueError:
        return ""
    if parts.scheme.lower() not in {"http", "https"} or not parts.netloc:
        return ""

    host = (parts.hostname or "").lower()
    path = parts.path.lower()
    # Never expose ChatGPT's private/internal file and backend destinations.
    if (host == "chatgpt.com" and path.startswith("/backend-api/")) or \
       host.endswith(".oaiusercontent.com") or \
       host.endswith(".blob.core.windows.net"):
        return ""

    pairs = parse_qsl(parts.query, keep_blank_values=True)
    pairs = [(k, v) for k, v in pairs
             if not (k.lower() == "utm_source" and v.lower() == "chatgpt.com")]
    query = urlencode(pairs, doseq=True)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, query, parts.fragment))


def iter_http_urls(value: Any) -> Iterable[str]:
    if isinstance(value, str):
        if value.startswith(("http://", "https://")):
            yield value
    elif isinstance(value, list):
        for item in value:
            yield from iter_http_urls(item)
    elif isinstance(value, dict):
        for item in value.values():
            yield from iter_http_urls(item)


def reference_urls(message: dict[str, Any], visible_text: str) -> list[str]:
    """Extract actual web references attached to a visible message, without search-result debris."""
    metadata = message.get("metadata")
    refs = metadata.get("content_references") if isinstance(metadata, dict) else None
    if not isinstance(refs, list):
        return []

    candidates: list[str] = []
    for ref in refs:
        if not isinstance(ref, dict):
            continue
        # grouped_webpages is the citation actually attached to visible prose.
        # Prefer safe_urls because it commonly contains the canonical URL as
        # well as ChatGPT's referral-attributed variant.
        if ref.get("type") == "grouped_webpages":
            safe = ref.get("safe_urls")
            if isinstance(safe, list):
                candidates.extend(u for u in safe if isinstance(u, str))
            else:
                candidates.extend(iter_http_urls(ref.get("items")))

    visible_clean = {sanitize_reference_url(m.group(0).rstrip(".,;:!?)"))
                     for m in URL_RE.finditer(visible_text)}
    output: list[str] = []
    seen: set[str] = set()
    for raw in candidates:
        clean = sanitize_reference_url(raw)
        if not clean or clean in seen:
            continue
        seen.add(clean)
        if clean not in visible_clean:
            output.append(clean)
    return output


def visible_messages(nodes: list[dict[str, Any]]) -> list[tuple[str, str, list[dict[str, Any]], list[str], float | None]]:
    output: list[tuple[str, str, list[dict[str, Any]], list[str], float | None]] = []
    for node in nodes:
        message = node.get("message")
        if not isinstance(message, dict):
            continue
        author = message.get("author")
        role = author.get("role") if isinstance(author, dict) else None
        if role not in {"user", "assistant"}:
            continue

        # Newer assistant messages can contain hidden analysis and a visible final.
        # Include final-channel messages and legacy messages that have no channel.
        if role == "assistant":
            channel = message.get("channel")
            recipient = message.get("recipient")
            # Only normal conversation output belongs in the archive. Tool calls
            # can also use channel=None/final and otherwise look like replies.
            if recipient != "all" or channel not in (None, "final"):
                continue

        text, uploads = extract_text_and_uploads(
            message.get("content"),
            message.get("metadata"),
        )

        message_id = message.get("id")
        if isinstance(message_id, str) and message_id:
            for upload in uploads:
                upload["message_id"] = message_id

        if text or uploads:
            refs = reference_urls(message, text) if role == "assistant" else []
            create_time = message.get("create_time")
            timestamp = float(create_time) if isinstance(create_time, (int, float)) else None
            output.append((role, text, uploads, refs, timestamp))
    return output


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("pic", PIC_NS)

def wtag(name: str) -> str:
    return f"{{{W_NS}}}{name}"

def rtag(name: str) -> str:
    return f"{{{R_NS}}}{name}"


def wptag(name: str) -> str:
    return f"{{{WP_NS}}}{name}"


def atag(name: str) -> str:
    return f"{{{A_NS}}}{name}"


def pictag(name: str) -> str:
    return f"{{{PIC_NS}}}{name}"

def add_inline_image(
    body: ET.Element,
    relationship_id: str,
    width_emu: int,
    height_emu: int,
    drawing_id: int,
) -> None:
    """Add one inline image in its own Word paragraph."""
    p = _new_paragraph(body, after=120)
    r = ET.SubElement(p, wtag("r"))
    drawing = ET.SubElement(r, wtag("drawing"))

    inline = ET.SubElement(drawing, wptag("inline"))
    inline.set("distT", "0")
    inline.set("distB", "0")
    inline.set("distL", "0")
    inline.set("distR", "0")

    extent = ET.SubElement(inline, wptag("extent"))
    extent.set("cx", str(width_emu))
    extent.set("cy", str(height_emu))

    effect_extent = ET.SubElement(inline, wptag("effectExtent"))
    effect_extent.set("l", "0")
    effect_extent.set("t", "0")
    effect_extent.set("r", "0")
    effect_extent.set("b", "0")

    doc_pr = ET.SubElement(inline, wptag("docPr"))
    doc_pr.set("id", str(drawing_id))
    doc_pr.set("name", f"Picture {drawing_id}")
    doc_pr.set("descr", "Uploaded image")

    frame_pr = ET.SubElement(inline, wptag("cNvGraphicFramePr"))
    locks = ET.SubElement(frame_pr, atag("graphicFrameLocks"))
    locks.set("noChangeAspect", "1")

    graphic = ET.SubElement(inline, atag("graphic"))
    graphic_data = ET.SubElement(graphic, atag("graphicData"))
    graphic_data.set(
        "uri",
        "http://schemas.openxmlformats.org/drawingml/2006/picture",
    )

    pic = ET.SubElement(graphic_data, pictag("pic"))

    nv_pic_pr = ET.SubElement(pic, pictag("nvPicPr"))
    c_nv_pr = ET.SubElement(nv_pic_pr, pictag("cNvPr"))
    c_nv_pr.set("id", "0")
    c_nv_pr.set("name", f"Picture {drawing_id}")
    ET.SubElement(nv_pic_pr, pictag("cNvPicPr"))

    blip_fill = ET.SubElement(pic, pictag("blipFill"))
    blip = ET.SubElement(blip_fill, atag("blip"))
    blip.set(rtag("embed"), relationship_id)

    stretch = ET.SubElement(blip_fill, atag("stretch"))
    ET.SubElement(stretch, atag("fillRect"))

    sp_pr = ET.SubElement(pic, pictag("spPr"))

    xfrm = ET.SubElement(sp_pr, atag("xfrm"))
    off = ET.SubElement(xfrm, atag("off"))
    off.set("x", "0")
    off.set("y", "0")

    ext = ET.SubElement(xfrm, atag("ext"))
    ext.set("cx", str(width_emu))
    ext.set("cy", str(height_emu))

    geometry = ET.SubElement(sp_pr, atag("prstGeom"))
    geometry.set("prst", "rect")
    ET.SubElement(geometry, atag("avLst"))

def clean_xml_text(text: str) -> str:
    # XML 1.0 permits supplementary Unicode characters through U+10FFFF.
    # Keep emoji/symbols; remove only characters XML 1.0 cannot contain.
    return "".join(
        ch for ch in text
        if ch in "\t\n\r"
        or 0x20 <= ord(ch) <= 0xD7FF
        or 0xE000 <= ord(ch) <= 0xFFFD
        or 0x10000 <= ord(ch) <= 0x10FFFF
    )

def _set_run_font(rpr: ET.Element, font: str) -> None:
    fonts = ET.SubElement(rpr, wtag("rFonts"))
    fonts.set(wtag("ascii"), font); fonts.set(wtag("hAnsi"), font)


def add_run(p: ET.Element, text: str, *, bold=False, italic=False, size=21, font="Aptos", shade: str | None = None, color: str | None = None) -> None:
    if not text: return
    r=ET.SubElement(p,wtag("r")); rpr=ET.SubElement(r,wtag("rPr")); _set_run_font(rpr,font)
    if bold: ET.SubElement(rpr,wtag("b"))
    if italic: ET.SubElement(rpr,wtag("i"))
    if color:
        c=ET.SubElement(rpr,wtag("color")); c.set(wtag("val"),color)
    if shade:
        shd=ET.SubElement(rpr,wtag("shd")); shd.set(wtag("val"),"clear"); shd.set(wtag("fill"),shade)
    for tag in ("sz","szCs"):
        e=ET.SubElement(rpr,wtag(tag)); e.set(wtag("val"),str(size))
    t=ET.SubElement(r,wtag("t")); t.set("{http://www.w3.org/XML/1998/namespace}space","preserve"); t.text=clean_xml_text(text)


def _new_paragraph(body: ET.Element, *, before=0, after=120, left=0, right=0, shade=None, keep_next=False) -> ET.Element:
    p=ET.SubElement(body,wtag("p")); ppr=ET.SubElement(p,wtag("pPr")); sp=ET.SubElement(ppr,wtag("spacing")); sp.set(wtag("before"),str(before)); sp.set(wtag("after"),str(after))
    if left or right:
        ind=ET.SubElement(ppr,wtag("ind"))
        if left: ind.set(wtag("left"),str(left))
        if right: ind.set(wtag("right"),str(right))
    if shade:
        shd=ET.SubElement(ppr,wtag("shd")); shd.set(wtag("val"),"clear"); shd.set(wtag("fill"),shade)

    if keep_next: ET.SubElement(ppr,wtag("keepNext"))
    return p

def add_message_body_container(body: ET.Element, shade_color: str) -> ET.Element:
    tbl = ET.SubElement(body, wtag("tbl"))
    tblpr = ET.SubElement(tbl, wtag("tblPr"))

    width = ET.SubElement(tblpr, wtag("tblW"))
    width.set(wtag("w"), "0")
    width.set(wtag("type"), "auto")

    borders = ET.SubElement(tblpr, wtag("tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = ET.SubElement(borders, wtag(edge))
        border.set(wtag("val"), "nil")

    tr = ET.SubElement(tbl, wtag("tr"))
    tc = ET.SubElement(tr, wtag("tc"))
    tcpr = ET.SubElement(tc, wtag("tcPr"))

    shade = ET.SubElement(tcpr, wtag("shd"))
    shade.set(wtag("val"), "clear")
    shade.set(wtag("fill"), shade_color)

    return tc


def add_paragraph(body: ET.Element, text="", *, bold=False, italic=False, size=21, before=0, after=120, left=0, right=0, shade=None, font="Aptos") -> None:
    p=_new_paragraph(body,before=before,after=after,left=left,right=right,shade=shade)
    text=clean_xml_text(html_lib.unescape(text)); lines=text.split("\n")
    for i,line in enumerate(lines):
        if i:
            rr=ET.SubElement(p,wtag("r")); ET.SubElement(rr,wtag("br"))
        add_run(p,line,bold=bold,italic=italic,size=size,font=font)


INLINE_RE=re.compile(r"(\[[^]\n]+\]\([^\n)]+\)|`[^`\n]+`|\*\*\*.*?\*\*\*|___.*?___|\*\*.*?\*\*|__.*?__|(?<!\*)\*(?![\s*])[^*\n]*?(?<!\s)\*(?!\*)|(?<![\w_])_[^_\n]+_(?![\w_]))", re.DOTALL)
INTERNAL_CITATION_RE = re.compile(r"\s*(?:filecite|memcite|cite)(?:[^]*)?")
MESSAGE_REACTION_RE = re.compile(r"\s*\uE200message_reaction\uE202[^\uE201]*\uE201")
URL_REFERENCE_RE = re.compile(r"\uE200url\uE202([^\uE202\uE201]*)\uE202[^\uE201]*\uE201")
MARKDOWN_ESCAPABLE = "!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~"
_MARKDOWN_ESCAPE_BASE = 0xE100
_MARKDOWN_ESCAPE_TO_SENTINEL = {ch: chr(_MARKDOWN_ESCAPE_BASE + i) for i, ch in enumerate(MARKDOWN_ESCAPABLE)}
_MARKDOWN_SENTINEL_TO_LITERAL = {v: k for k, v in _MARKDOWN_ESCAPE_TO_SENTINEL.items()}


def protect_markdown_escapes(text: str) -> str:
    # CommonMark permits a backslash to escape any ASCII punctuation character.
    # Protect those pairs before block/inline parsing so escaped punctuation cannot
    # be mistaken for Markdown syntax. Backslash escapes are not interpreted inside
    # inline code spans, so those spans are copied verbatim. Ordinary paths such as
    # C:\Windows are unchanged because W is not punctuation.
    out=[]; i=0
    while i < len(text):
        if text[i] == "`":
            end=text.find("`", i + 1)
            if end != -1:
                out.append(text[i:end + 1]); i=end + 1; continue
        if text[i] == "\\" and i + 1 < len(text) and text[i + 1] in _MARKDOWN_ESCAPE_TO_SENTINEL:
            out.append(_MARKDOWN_ESCAPE_TO_SENTINEL[text[i + 1]])
            i += 2
        else:
            out.append(text[i]); i += 1
    return "".join(out)

def restore_markdown_escapes(text: str) -> str:
    return "".join(_MARKDOWN_SENTINEL_TO_LITERAL.get(ch, ch) for ch in text)

def strip_internal_citations(text: str) -> str:
    # ChatGPT UI tokens are transport/rendering artifacts, not prose.
    text = INTERNAL_CITATION_RE.sub("", text)
    text = MESSAGE_REACTION_RE.sub("", text)
    text = URL_REFERENCE_RE.sub(r"\1", text)
    return text

def _add_plain_inline(p: ET.Element, text: str, *, size=21, bold=False, italic=False) -> None:
    parts=text.split("\n")
    for j,part in enumerate(parts):
        if j:
            rr=ET.SubElement(p,wtag("r")); ET.SubElement(rr,wtag("br"))
        add_run(p,restore_markdown_escapes(part),bold=bold,italic=italic,size=size)


def add_inline_markdown(p: ET.Element, text: str, *, size=21, base_bold=False, base_italic=False) -> None:
    text=protect_markdown_escapes(html_lib.unescape(strip_internal_citations(text))); pos=0
    for m in INLINE_RE.finditer(text):
        if m.start()>pos:
            _add_plain_inline(p,text[pos:m.start()],size=size,bold=base_bold,italic=base_italic)
        token=m.group(0)
        if token.startswith("["):
            lm=re.match(r"^\[([^]]+)\]\((.*)\)$",token)
            if lm:
                label, raw_url = lm.group(1), lm.group(2)
                label = restore_markdown_escapes(label)
                raw_url = restore_markdown_escapes(raw_url)
                clean_url = sanitize_reference_url(raw_url)
                add_run(p,label,bold=base_bold,italic=base_italic,size=size)
                # URL text is intentionally an ordinary Word run, never a hyperlink.
                if clean_url and sanitize_reference_url(label) != clean_url:
                    add_run(p," (" + clean_url + ")",bold=base_bold,italic=base_italic,size=size)
            else:
                add_run(p,restore_markdown_escapes(token),bold=base_bold,italic=base_italic,size=size)
        elif token.startswith("`"):
            add_run(p,restore_markdown_escapes(token[1:-1]),bold=base_bold,italic=base_italic,size=max(16,size-1),font="Consolas",shade="EDEDED")
        elif token.startswith("***") or token.startswith("___"):
            add_inline_markdown(p,token[3:-3],size=size,base_bold=True,base_italic=True)
        elif token.startswith("**") or token.startswith("__"):
            # Parse the contents again so nested inline Markdown (for example
            # bold text containing `inline code`) is preserved instead of
            # leaving the inner Markdown delimiters visible. DOTALL above also
            # allows bold text to span a Markdown hard line break.
            add_inline_markdown(p,token[2:-2],size=size,base_bold=True,base_italic=base_italic)
        else:
            add_inline_markdown(p,token[1:-1],size=size,base_bold=base_bold,base_italic=True)
        pos=m.end()
    if pos<len(text):
        _add_plain_inline(p,text[pos:],size=size,bold=base_bold,italic=base_italic)


def add_code_block(body: ET.Element, lines: list[str], language: str="") -> None:
    if language:
        hp=_new_paragraph(body,before=80,after=0,left=288,right=288,shade="E3E3E3",keep_next=True); add_run(hp,language,size=17)
    p=_new_paragraph(body,before=80 if not language else 0,after=140,left=288,right=288,shade="F2F2F2")
    for i,line in enumerate(lines):
        if i:
            rr=ET.SubElement(p,wtag("r")); ET.SubElement(rr,wtag("br"))
        add_run(p,line if line else " ",size=18,font="Consolas")


def add_rule(body: ET.Element) -> None:
    p=_new_paragraph(body,before=80,after=100); ppr=p.find(wtag("pPr")); borders=ET.SubElement(ppr,wtag("pBdr")); bottom=ET.SubElement(borders,wtag("bottom"))
    bottom.set(wtag("val"),"single"); bottom.set(wtag("sz"),"4"); bottom.set(wtag("space"),"1"); bottom.set(wtag("color"),"D9D9D9")



def _split_markdown_table_row(line: str) -> list[str]:
    stripped=line.strip()
    if stripped.startswith("|"): stripped=stripped[1:]
    if stripped.endswith("|"): stripped=stripped[:-1]
    # Escaped pipes have already been protected. A pipe inside an inline-code
    # span is literal content too, so split only on pipes outside backticks.
    cells=[]; current=[]; in_code=False
    for ch in stripped:
        if ch == "`":
            in_code = not in_code
            current.append(ch)
        elif ch == "|" and not in_code:
            cells.append("".join(current).strip()); current=[]
        else:
            current.append(ch)
    cells.append("".join(current).strip())
    return cells


def _is_markdown_table_separator(line: str) -> bool:
    cells=_split_markdown_table_row(line)
    return len(cells) >= 1 and all(re.fullmatch(r":?-+:?",c.replace(" ","")) for c in cells)


def add_markdown_table(body: ET.Element, rows: list[list[str]]) -> None:
    if not rows: return
    # Empty one-cell tables occur as layout debris in pasted email content.
    # Preserve nonempty tables, but do not emit visible Word borders for tables
    # whose cells contain no actual content.
    if all(not cell.strip() for row in rows for cell in row):
        return
    cols=max(len(r) for r in rows)
    tbl=ET.SubElement(body,wtag("tbl")); tblpr=ET.SubElement(tbl,wtag("tblPr"))
    borders=ET.SubElement(tblpr,wtag("tblBorders"))
    for edge in ("top","left","bottom","right","insideH","insideV"):
        b=ET.SubElement(borders,wtag(edge)); b.set(wtag("val"),"single"); b.set(wtag("sz"),"4"); b.set(wtag("color"),"D9D9D9")
    width=ET.SubElement(tblpr,wtag("tblW")); width.set(wtag("w"),"0"); width.set(wtag("type"),"auto")
    for rix,row in enumerate(rows):
        tr=ET.SubElement(tbl,wtag("tr"))
        if rix==0:
            trpr=ET.SubElement(tr,wtag("trPr")); ET.SubElement(trpr,wtag("tblHeader"))
        for cix in range(cols):
            tc=ET.SubElement(tr,wtag("tc")); tcpr=ET.SubElement(tc,wtag("tcPr"))
            if rix==0:
                shd=ET.SubElement(tcpr,wtag("shd")); shd.set(wtag("val"),"clear"); shd.set(wtag("fill"),"F2F2F2")
            p=_new_paragraph(tc,after=40)
            add_inline_markdown(p,(row[cix] if cix<len(row) else "").replace("\uE000", "\n"),size=19,base_bold=(rix==0))
    spacer=_new_paragraph(body,after=80)

def _fence_match(line: str):
    # CommonMark permits fenced code blocks to be indented by up to 3 spaces.
    # ChatGPT commonly uses that form when a code block belongs to a list item.
    return re.match(r"^ {0,3}```(.*)$", line)


def add_markdown(body: ET.Element, text: str, preserve_soft_breaks: bool = False) -> None:
    text=strip_internal_citations(text)
    # Keep HTML break tags inside a single logical Markdown line until block
    # structures (especially pipe tables) have been recognized.  Converting
    # them to real line breaks too early can split a table row in the middle.
    text=re.sub(r"<br\s*/?>", "\uE000", text, flags=re.IGNORECASE)
    lines=clean_xml_text(text).replace("\r\n","\n").replace("\r","\n").split("\n"); i=0
    protected_lines=[protect_markdown_escapes(x) for x in lines]
    while i<len(lines):
        raw_line=lines[i]
        structural_line=protected_lines[i]
        line=structural_line.replace("\uE000", "\n")
        fence=_fence_match(structural_line)
        if fence:
            language=fence.group(1).strip(); code=[]; i+=1
            while i<len(lines) and not _fence_match(protected_lines[i]):
                code.append(lines[i]); i+=1
            if i<len(lines): i+=1
            add_code_block(body,code,language); continue
        if not line.strip(): i+=1; continue
        if i+1 < len(lines) and "|" in structural_line and _is_markdown_table_separator(protected_lines[i+1]):
            rows=[_split_markdown_table_row(structural_line)]; i+=2
            while i<len(lines) and lines[i].strip() and "|" in lines[i]:
                rows.append(_split_markdown_table_row(protected_lines[i])); i+=1
            add_markdown_table(body,rows); continue
        if re.fullmatch(r"\s*---+\s*",line): add_rule(body); i+=1; continue
        mh=re.match(r"^(#{1,6})\s+(.*)$",line)
        if mh:
            level=len(mh.group(1)); sizes={1:30,2:27,3:24,4:22,5:21,6:20}; p=_new_paragraph(body,before=180 if level<=2 else 120,after=70,keep_next=True); add_inline_markdown(p,mh.group(2),size=sizes[level],base_bold=True); i+=1; continue
        mq=re.match(r"^\s*>\s?(.*)$",line)
        if mq:
            qlines=[]
            while i<len(lines):
                q=re.match(r"^\s*>\s?(.*)$",lines[i])
                if not q: break
                qlines.append(q.group(1)); i+=1
            p=_new_paragraph(body,before=50,after=110,left=360,right=180,shade="F7F7F7")
            for j,qline in enumerate(qlines):
                if j:
                    rr=ET.SubElement(p,wtag("r")); ET.SubElement(rr,wtag("br"))
                add_inline_markdown(p,qline,size=20,base_italic=True)
            continue
        ml=re.match(r"^(\s*)[-*+]\s+(.*)$",line); mn=re.match(r"^(\s*)\d+[.)]\s+(.*)$",line)
        if ml or mn:
            m=ml or mn; indent_spaces=len(m.group(1)); content=m.group(2); marker="•" if ml else re.match(r"^\s*(\d+)[.)]",line).group(1)+"."
            p=_new_paragraph(body,after=35,left=360+indent_spaces*72); add_run(p,marker+"  ",size=21); add_inline_markdown(p,content,size=21); i+=1; continue
        para=[line]; i+=1
        while i<len(lines) and lines[i].strip():
            nxt=protected_lines[i]
            if _fence_match(nxt) or re.match(r"^(#{1,6})\s+",nxt) or re.match(r"^\s*>\s?",nxt) or re.match(r"^(\s*)[-*+]\s+",nxt) or re.match(r"^(\s*)\d+[.)]\s+",nxt) or re.fullmatch(r"\s*---+\s*",nxt): break
            para.append(nxt); i+=1
        # Markdown hard breaks are two trailing spaces or a trailing backslash.
        joined=""
        for j,x in enumerate(para):
            hard=x.endswith("  ") or (x.endswith("\\") and not x.endswith("\\\\"))
            piece=x[:-2] if x.endswith("  ") else (x[:-1] if hard and x.endswith("\\") else x)
            joined += piece if preserve_soft_breaks else piece.strip()
            if j < len(para)-1:
                joined += "\n" if (hard or preserve_soft_breaks) else " "

        p=_new_paragraph(body,after=120); add_inline_markdown(p,joined,size=21)


def format_message_timestamp(timestamp: float | None) -> str:
    if timestamp is None:
        return ""
    dt = _dt.datetime.fromtimestamp(timestamp, ZoneInfo("America/New_York"))
    zone_name = "Eastern Daylight Time" if dt.dst() and dt.dst() != _dt.timedelta(0) else "Eastern Standard Time"
    return f"{dt:%Y-%m-%d %I:%M:%S %p} {zone_name}"


def add_message_header(body: ET.Element, role: str, timestamp: float | None) -> None:
    p = _new_paragraph(body, before=160, after=60, keep_next=True)
    if role == "user":
        add_run(p, " YOU ", bold=True, size=20, color="000000", shade="FA7E33")
    else:
        add_run(p, " GPT ", bold=True, size=20, color="FFFFFF", shade="008000")
    stamp = format_message_timestamp(timestamp)
    if stamp:
        add_run(p, "  " + stamp, size=16, color="666666")


def create_docx(title: str, messages: list[tuple[str,str,list[dict[str, Any]],list[str],float | None]], output_path: Path) -> None:
    doc_relationships = [
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    ]

    image_relationships: list[tuple[dict[str, Any], str, str]] = []
    image_number = 0

    for _, _, uploads, _, _ in messages:
        for upload in uploads:
            if upload.get("image_info") is None:
                continue            

            image_number += 1
            relationship_id = f"rId{image_number + 2}"

            image_extension = upload["image_info"]["extension"]
            media_name = f"image{image_number}.{image_extension}"            

            upload["docx_relationship_id"] = relationship_id
            upload["docx_media_name"] = media_name            

            doc_relationships.append(
                f'<Relationship Id="{relationship_id}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="media/{media_name}"/>'
            )
            image_relationships.append((upload, relationship_id, media_name))    

    document=ET.Element(wtag("document")); body=ET.SubElement(document,wtag("body"))
    add_paragraph(body,title,bold=True,size=36,after=80); add_paragraph(body,"ChatGPT Conversation Archive",italic=True,size=20,after=240)
    drawing_id = 0

    for role,text,uploads,refs,timestamp in messages:
        add_message_header(body, role, timestamp)

        message_body = add_message_body_container(
            body,
            "FFFBF2" if role == "user" else "F7FFF5",
        )

        for upload in uploads:
            drawing_id += 1
            relationship_id = upload.get("docx_relationship_id")
            display_size = upload.get("display_size_emu")

            if relationship_id and display_size:
                width_emu, height_emu = display_size
                add_inline_image(
                    message_body,
                    relationship_id,
                    width_emu,
                    height_emu,
                    drawing_id,
                )
            else:
                add_paragraph(
                    body,
                    unavailable_upload_text(upload),
                    bold=True,
                    italic=True,
                    size=20,
                    after=100,
                    left=288,
                    shade="FFF2CC",
                )

        if text: add_markdown(message_body,text,preserve_soft_breaks=(role == "user"))

        if refs:
            add_paragraph(message_body,"Reference URLs:",bold=True,size=19,before=40,after=30)
            for url in refs:
                # Deliberately plain text: no w:hyperlink and no external relationship.
                add_paragraph(message_body,url,size=18,after=25,left=288,font="Consolas")

        _new_paragraph(message_body, after=0)

    sect=ET.SubElement(body,wtag("sectPr")); pgsz=ET.SubElement(sect,wtag("pgSz")); pgsz.set(wtag("w"),"12240"); pgsz.set(wtag("h"),"15840"); mar=ET.SubElement(sect,wtag("pgMar"))
    for k,v in {"top":"936","right":"1080","bottom":"936","left":"1080","header":"720","footer":"720","gutter":"0"}.items(): mar.set(wtag(k),v)
    content_types='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="webp" ContentType="image/webp"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'''
    root_rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'''
    
    styles='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style></w:styles>'''
    now=_dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00","Z")
    core=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{html_lib.escape(title)}</dc:title><dc:creator>ChatGPT Conversation Bridge</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>'''
    app='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>ChatGPT Conversation Bridge</Application></Properties>'''
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(doc_relationships)
        + '</Relationships>'
    )    
    if output_path.exists(): output_path.unlink()
    with zipfile.ZipFile(output_path,"w",compression=zipfile.ZIP_DEFLATED) as zf:

        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr(
            "word/document.xml",
            ET.tostring(document, encoding="utf-8", xml_declaration=True),
        )
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr(
            "word/settings.xml",
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:writeProtection w:recommended="true"/></w:settings>',
        )
        zf.writestr("word/styles.xml", styles)
        zf.writestr("docProps/core.xml", core)
        zf.writestr("docProps/app.xml", app)
        
        for upload, _, media_name in image_relationships:
            zf.writestr(f"word/media/{media_name}", upload["image_data"])

def load_public_share_conversation(
    share_url: str,
) -> tuple[dict[str, Any], Any, str]:
    """Load the conversation JSON for a public ChatGPT Share URL."""
    parsed = urlsplit(share_url)

    if parsed.scheme not in {"http", "https"}:
        raise ExtractionError("The Share URL must use http:// or https://.")

    if parsed.netloc.lower() != "chatgpt.com":
        raise ExtractionError("The Share URL must be on chatgpt.com.")

    parts = [part for part in parsed.path.split("/") if part]

    if len(parts) != 2 or parts[0] != "share":
        raise ExtractionError(
            "Expected a ChatGPT Share URL in this form:\n"
            "https://chatgpt.com/share/<share-id>"
        )

    share_id = parts[1]
    canonical_share_url = f"https://chatgpt.com/share/{share_id}"
    api_url = f"https://chatgpt.com/backend-api/share/{share_id}"

    try:
        session = requests.Session(impersonate="chrome")

        share_response = session.get(canonical_share_url)
        share_response.raise_for_status()

        api_response = session.get(api_url)
        api_response.raise_for_status()

        conversation = api_response.json()

    except Exception as exc:
        raise ExtractionError(
            f"Could not retrieve the public ChatGPT conversation: {exc}"
        ) from exc

    if not isinstance(conversation, dict):
        raise ExtractionError(
            "The ChatGPT Share response is not a conversation object."
        )

    return conversation, session, share_id

def load_public_share_image(
    session: Any,
    share_id: str,
    upload: dict[str, Any],
) -> bytes | None:
    """Retrieve one public-share image from its stable file ID."""
    asset_pointer = upload.get("asset_pointer")

    if not (
        isinstance(asset_pointer, str)
        and asset_pointer.startswith("sediment://")
    ):
        return None

    file_id = asset_pointer[len("sediment://"):].split("?", 1)[0]

    if not file_id.startswith("file_"):
        return None

    resolver_url = (
        f"https://chatgpt.com/backend-api/share/"
        f"{share_id}/file/{file_id}"
    )

    try:
        resolver_response = session.get(resolver_url)
        resolver_response.raise_for_status()
        resolver = resolver_response.json()
    except Exception as exc:
        raise ExtractionError(
            f"Could not resolve uploaded image {file_id}: {exc}"
        ) from exc

    if not isinstance(resolver, dict):
        return None

    if resolver.get("status") != "success":
        return None

    download_url = resolver.get("download_url")

    if not isinstance(download_url, str) or not download_url:
        return None

    try:
        image_response = session.get(download_url)
        image_response.raise_for_status()
    except Exception as exc:
        raise ExtractionError(
            f"Could not download uploaded image {file_id}: {exc}"
        ) from exc

    return image_response.content

def create_public_share_archive(
    conversation: dict[str, Any],
    messages: list[
        tuple[
            str,
            str,
            list[dict[str, Any]],
            list[str],
            float | None,
        ]
    ],
    archive_path: Path,
) -> None:
    """Create and verify the permanent archive for a public Share conversation."""
    if archive_path.exists():
        raise ExtractionError(
            f"Archive already exists: {archive_path}"
        )

    archive_path.parent.mkdir(exist_ok=True)

    expected_members = {"conversation.json"}

    try:
        with zipfile.ZipFile(
            archive_path,
            "x",
            compression=zipfile.ZIP_DEFLATED,
        ) as zf:
            conversation_json = json.dumps(
                conversation,
                ensure_ascii=False,
                indent=2,
            ).encode("utf-8")

            zf.writestr(
                "conversation.json",
                conversation_json,
            )

            archived_file_ids: set[str] = set()

            for _, _, uploads, _, _ in messages:
                for upload in uploads:
                    image_data = upload.get("image_data")
                    info = upload.get("image_info")

                    if image_data is None or info is None:
                        continue

                    asset_pointer = upload.get("asset_pointer")

                    if not (
                        isinstance(asset_pointer, str)
                        and asset_pointer.startswith("sediment://")
                    ):
                        raise ExtractionError(
                            "An embedded image has no valid "
                            "sediment asset pointer."
                        )

                    file_id = (
                        asset_pointer[len("sediment://"):]
                        .split("?", 1)[0]
                    )

                    if not file_id.startswith("file_"):
                        raise ExtractionError(
                            "An embedded image has no valid file ID."
                        )

                    if file_id in archived_file_ids:
                        continue

                    extension = info.get("extension")

                    if not (
                        isinstance(extension, str)
                        and extension
                    ):
                        raise ExtractionError(
                            f"Could not determine the image "
                            f"extension for {file_id}."
                        )

                    if not extension.startswith("."):
                        extension = f".{extension}"

                    archive_member = (
                        f"uploads/{file_id}{extension.lower()}"
                    )

                    zf.writestr(
                        archive_member,
                        image_data,
                    )

                    expected_members.add(archive_member)
                    archived_file_ids.add(file_id)

    except Exception:
        try:
            archive_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise

    try:
        with zipfile.ZipFile(archive_path, "r") as zf:
            bad_member = zf.testzip()

            if bad_member is not None:
                raise ExtractionError(
                    "Archive verification failed for member: "
                    f"{bad_member}"
                )

            actual_members = set(zf.namelist())

            if actual_members != expected_members:
                missing = sorted(
                    expected_members - actual_members
                )
                unexpected = sorted(
                    actual_members - expected_members
                )

                details = []

                if missing:
                    details.append(
                        "Missing: " + ", ".join(missing)
                    )

                if unexpected:
                    details.append(
                        "Unexpected: " + ", ".join(unexpected)
                    )

                raise ExtractionError(
                    "Archive membership verification failed. "
                    + " ".join(details)
                )

            archived_conversation = json.loads(
                zf.read("conversation.json").decode("utf-8")
            )

            if archived_conversation != conversation:
                raise ExtractionError(
                    "Archived conversation.json does not match "
                    "the retrieved conversation."
                )

    except Exception:
        try:
            archive_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise

def load_public_share_archive(
    archive_path: Path,
) -> tuple[dict[str, Any], dict[str, bytes]]:
    """
    Load and validate a permanent public-share archive.

    Return the archived conversation object and a mapping of stable
    file IDs to their archived image bytes.
    """
    if not archive_path.is_file():
        raise ExtractionError(
            f"Archive does not exist: {archive_path}"
        )

    try:
        with zipfile.ZipFile(archive_path, "r") as zf:
            bad_member = zf.testzip()

            if bad_member is not None:
                raise ExtractionError(
                    "Archive verification failed for member: "
                    f"{bad_member}"
                )

            members = zf.namelist()

            if "conversation.json" not in members:
                raise ExtractionError(
                    "Archive does not contain conversation.json."
                )

            try:
                conversation = json.loads(
                    zf.read("conversation.json").decode("utf-8")
                )
            except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                raise ExtractionError(
                    "Archive conversation.json is invalid."
                ) from exc

            if not isinstance(conversation, dict):
                raise ExtractionError(
                    "Archive conversation.json is not an object."
                )

            image_data_by_file_id: dict[str, bytes] = {}

            for member in members:
                if not member.startswith("uploads/"):
                    continue

                if member.endswith("/"):
                    continue

                filename = member[len("uploads/"):]

                if not filename or "/" in filename:
                    raise ExtractionError(
                        f"Invalid archived upload path: {member}"
                    )

                file_id = filename.rsplit(".", 1)[0]

                if not file_id.startswith("file_"):
                    raise ExtractionError(
                        f"Invalid archived upload file ID: {member}"
                    )

                if file_id in image_data_by_file_id:
                    raise ExtractionError(
                        "Archive contains more than one image for "
                        f"{file_id}."
                    )

                image_data = zf.read(member)

                if image_info(image_data) is None:
                    raise ExtractionError(
                        "Archive contains an unsupported or invalid "
                        f"image: {member}"
                    )

                image_data_by_file_id[file_id] = image_data

    except zipfile.BadZipFile as exc:
        raise ExtractionError(
            f"Archive is not a valid ZIP file: {archive_path}"
        ) from exc

    return conversation, image_data_by_file_id

def prepare_messages_from_public_share_archive(
    conversation: dict[str, Any],
    image_data_by_file_id: dict[str, bytes],
) -> tuple[
    list[dict[str, Any]],
    list[
        tuple[
            str,
            str,
            list[dict[str, Any]],
            list[str],
            float | None,
        ]
    ],
]:
    """
    Validate an archived conversation and prepare its visible messages
    for the existing DOCX renderer.
    """
    linear = conversation.get("linear_conversation")

    if not isinstance(linear, list):
        raise ExtractionError(
            "linear_conversation is missing or invalid."
        )

    nodes = [
        node
        for node in linear
        if isinstance(node, dict)
    ]

    if len(nodes) != len(linear):
        raise ExtractionError(
            "linear_conversation contains an invalid node."
        )

    validate_chain(
        nodes,
        conversation.get("current_node"),
    )

    messages = visible_messages(nodes)

    if not messages:
        raise ExtractionError(
            "No visible user/assistant messages were found."
        )

    used_image_file_ids: set[str] = set()

    for _, _, uploads, _, _ in messages:
        for upload in uploads:
            asset_pointer = upload.get("asset_pointer")
            file_id = None

            if (
                isinstance(asset_pointer, str)
                and asset_pointer.startswith("sediment://")
            ):
                file_id = (
                    asset_pointer[len("sediment://"):]
                    .split("?", 1)[0]
                )

            image_data = (
                image_data_by_file_id.get(file_id)
                if file_id is not None
                else None
            )

            upload["image_data"] = image_data
            upload["image_info"] = image_info(image_data)

            info = upload.get("image_info")

            upload["image_dimensions"] = (
                info.get("dimensions")
                if info is not None
                else None
            )

            upload["display_size_emu"] = image_display_size_emu(
                upload.get("image_dimensions")
            )

            if image_data is not None and file_id is not None:
                used_image_file_ids.add(file_id)

    unused_image_file_ids = (
        set(image_data_by_file_id) - used_image_file_ids
    )

    if unused_image_file_ids:
        raise ExtractionError(
            "Archive contains image data not referenced by the "
            "visible conversation: "
            + ", ".join(sorted(unused_image_file_ids))
        )

    return nodes, messages

def main() -> int:
    start_time = _dt.datetime.now()
    start_counter = time.perf_counter()
    print(
        f"Start Time: {start_time.strftime('%Y-%m-%d %I:%M:%S %p')}"
    )    

    parser = argparse.ArgumentParser(
        description=(
            "Create a DOCX from a public ChatGPT Share URL or "
            "regenerate one from a permanent archive."
        )
    )

    parser.add_argument(
        "source",
        help="Public ChatGPT Share URL or permanent archive ZIP",
    )

    args = parser.parse_args()

    working_dir, archive_dir, docx_dir = managed_directories(
        Path(__file__)
    )

    docx_dir.mkdir(exist_ok=True)

    source = args.source.strip()

    if source.startswith(("https://", "http://")):
        source_mode = "share"

        conversation, session, share_id = (
            load_public_share_conversation(
                source
            )
        )

        linear = conversation.get("linear_conversation")

        if not isinstance(linear, list):
            raise ExtractionError(
                "linear_conversation is missing or invalid."
            )

        nodes = [
            node
            for node in linear
            if isinstance(node, dict)
        ]

        if len(nodes) != len(linear):
            raise ExtractionError(
                "linear_conversation contains an invalid node."
            )

        validate_chain(
            nodes,
            conversation.get("current_node"),
        )

        share_title = conversation.get("title")

        if not isinstance(share_title, str) or not share_title.strip():
            share_title = "(No title)"

        share_archive_path = (
            archive_dir / f"{share_title.strip()}.zip"
        )

        if share_archive_path.exists():
            raise ExtractionError(
                f"Archive already exists: {share_archive_path}"
            )

        messages = visible_messages(nodes)

        total_uploads = sum(
            len(uploads)
            for _, _, uploads, _, _ in messages
        )

        processed_uploads = 0

        for _, _, uploads, _, _ in messages:
            for upload in uploads:
                upload["image_data"] = load_public_share_image(
                    session,
                    share_id,
                    upload,
                )

                upload["image_info"] = image_info(
                    upload.get("image_data")
                )

                info = upload.get("image_info")

                upload["image_dimensions"] = (
                    info.get("dimensions")
                    if info is not None
                    else None
                )

                upload["display_size_emu"] = image_display_size_emu(
                    upload.get("image_dimensions")
                )

                processed_uploads += 1
                progress_percent = round(
                    processed_uploads * 100 / total_uploads
                )

                print(
                    f"\rProcessing uploads: "
                    f"{processed_uploads}/{total_uploads} "
                    f"({progress_percent}%)",
                    end="",
                    flush=True,
                )

        if total_uploads:
            print()

        if not messages:
            raise ExtractionError(
                "No visible user/assistant messages were found."
            )

    else:
        source_mode = "archive"

        archive_input_path = Path(source)

        if not archive_input_path.is_absolute():
            if archive_input_path.parent == Path("."):
                archive_input_path = archive_dir / archive_input_path
            else:
                archive_input_path = working_dir / archive_input_path

        conversation, image_data_by_file_id = (
            load_public_share_archive(
                archive_input_path
            )
        )

        nodes, messages = (
            prepare_messages_from_public_share_archive(
                conversation,
                image_data_by_file_id,
            )
        )

    title = conversation.get("title")

    if not isinstance(title, str) or not title.strip():
        title = "(No title)"

    output_path = docx_dir / f"{title.strip()}.docx"

    create_docx(
        title.strip(),
        messages,
        output_path,
    )

    archive_path = None

    if source_mode == "share":
        archive_path = archive_dir / f"{title.strip()}.zip"

        create_public_share_archive(
            conversation,
            messages,
            archive_path,
        )

    user_count = sum(
        1
        for role, _, _, _, _ in messages
        if role == "user"
    )

    assistant_count = sum(
        1
        for role, _, _, _, _ in messages
        if role == "assistant"
    )

    upload_count = sum(
        len(uploads)
        for _, _, uploads, _, _ in messages
    )

    embedded_image_count = sum(
        1
        for _, _, uploads, _, _ in messages
        for upload in uploads
        if upload.get("image_info") is not None
    )

    unavailable_image_count = sum(
        1
        for _, _, uploads, _, _ in messages
        for upload in uploads
        if upload.get("image_info") is None
        and (
            upload.get("content_type") == "image_asset_pointer"
            or (
                isinstance(upload.get("mime_type"), str)
                and upload["mime_type"].lower().startswith("image/")
            )
        )
    )

    unavailable_attachment_count = sum(
        1
        for _, _, uploads, _, _ in messages
        for upload in uploads
        if upload.get("image_info") is None
        and not (
            upload.get("content_type") == "image_asset_pointer"
            or (
                isinstance(upload.get("mime_type"), str)
                and upload["mime_type"].lower().startswith("image/")
            )
        )
    )

    reference_count = sum(
        len(refs)
        for _, _, _, refs, _ in messages
    )

    timestamp_count = sum(
        1
        for *_, timestamp in messages
        if timestamp is not None
    )

    print(f"Output:     {output_path}")

    if archive_path is not None:
        print(f"Archive:    {archive_path}")
    else:
        print(f"Archive:    {archive_input_path}")

    print(f"Title:      {title.strip()}")
    print(f"Nodes:      {len(nodes)} (chain validated)")
    print(
        f"Messages:   {len(messages)} "
        f"({user_count} user, "
        f"{assistant_count} assistant)"
    )
    print(f"Uploads:    {upload_count} upload reference(s)")

    print(
        f"Images:     {embedded_image_count} embedded, "
        f"{unavailable_image_count} unavailable"
    )
    print(
        f"Attachments:{unavailable_attachment_count:3} unavailable"
    )

    print(
        f"Timestamps: {timestamp_count} "
        "message timestamp(s) preserved"
    )
    print(
        f"References: {reference_count} "
        "public reference URL(s) found"
    )

    end_time = _dt.datetime.now()
    duration_seconds = time.perf_counter() - start_counter

    hours, remainder = divmod(duration_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    print(
        f"End Time:   {end_time.strftime('%Y-%m-%d %I:%M:%S %p')}"
    )
    print(
        f"Duration:   {int(hours):02}:"
        f"{int(minutes):02}:"
        f"{seconds:05.2f}"
    )

    if source_mode == "share":
        print(
            "Status:     SHARE JSON + IMAGE + DOCX + ARCHIVE PASS"
        )
    else:
        print(
            "Status:     OFFLINE ARCHIVE + DOCX PASS"
        )

    return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except ExtractionError as exc: print(f"ERROR: {exc}",file=sys.stderr); raise SystemExit(1)
    except PermissionError as exc: print(f"ERROR: Cannot overwrite output file (is it open in Word?): {exc}",file=sys.stderr); raise SystemExit(1)
