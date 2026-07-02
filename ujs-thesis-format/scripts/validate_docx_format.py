#!/usr/bin/env python3
"""Validate key DOCX requirements for the UJS 2026 thesis format."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


@dataclass
class Finding:
    severity: str
    code: str
    message: str


def w_attr(name: str) -> str:
    return f"{{{NS['w']}}}{name}"


def mm_to_twips(mm: float) -> int:
    return round(mm / 25.4 * 1440)


def read_zip_text(zf: zipfile.ZipFile, name: str) -> str | None:
    try:
        return zf.read(name).decode("utf-8")
    except KeyError:
        return None


def parse_xml(text: str | None) -> ET.Element | None:
    if not text:
        return None
    return ET.fromstring(text.encode("utf-8"))


def element_text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.findall(".//w:t", NS))


def iter_paragraph_text(root: ET.Element) -> Iterable[str]:
    for p in root.findall(".//w:p", NS):
        txt = element_text(p).strip()
        if txt:
            yield txt


def check_page_setup(root: ET.Element) -> list[Finding]:
    findings: list[Finding] = []
    sects = root.findall(".//w:sectPr", NS)
    if not sects:
        return [Finding("ERROR", "page.section", "No section properties found in word/document.xml.")]

    sect = sects[-1]
    pg_sz = sect.find("w:pgSz", NS)
    pg_mar = sect.find("w:pgMar", NS)
    if pg_sz is None:
        findings.append(Finding("ERROR", "page.size.missing", "Missing page size. Expected A4 portrait."))
    else:
        width = int(pg_sz.get(w_attr("w"), "0"))
        height = int(pg_sz.get(w_attr("h"), "0"))
        expected_w = mm_to_twips(210)
        expected_h = mm_to_twips(297)
        if abs(width - expected_w) > 40 or abs(height - expected_h) > 40:
            findings.append(
                Finding(
                    "ERROR",
                    "page.size",
                    f"Page size is {width} x {height} twips; expected A4 portrait about {expected_w} x {expected_h}.",
                )
            )

    expected_margins = {
        "top": mm_to_twips(25),
        "left": mm_to_twips(25),
        "bottom": mm_to_twips(20),
        "right": mm_to_twips(20),
    }
    if pg_mar is None:
        findings.append(Finding("ERROR", "page.margin.missing", "Missing page margins."))
    else:
        for key, expected in expected_margins.items():
            actual = int(pg_mar.get(w_attr(key), "0"))
            if abs(actual - expected) > 60:
                findings.append(
                    Finding(
                        "ERROR",
                        f"page.margin.{key}",
                        f"{key} margin is {actual} twips; expected about {expected} twips.",
                    )
                )
    return findings


def sequence_warnings(label: str, numbers: list[int]) -> list[Finding]:
    if not numbers:
        return []
    findings: list[Finding] = []
    expected = list(range(1, max(numbers) + 1))
    unique_numbers = sorted(set(numbers))
    if unique_numbers != expected:
        findings.append(
            Finding(
                "WARN",
                f"{label}.sequence",
                f"{label} numbers are {unique_numbers}; expected a continuous sequence starting at 1.",
            )
        )
    if len(numbers) != len(set(numbers)):
        findings.append(Finding("WARN", f"{label}.duplicate", f"{label} captions contain duplicate numbers."))
    return findings


def is_black_font_color(color_el: ET.Element) -> bool:
    val = color_el.get(w_attr("val"))
    theme = color_el.get(w_attr("themeColor"))
    if val is None:
        return theme is None
    normalized = val.strip().lstrip("#").upper()
    return normalized in {"AUTO", "000", "000000"}


def describe_font_color(color_el: ET.Element) -> str:
    val = color_el.get(w_attr("val"))
    theme = color_el.get(w_attr("themeColor"))
    parts = []
    if val:
        parts.append(f"val={val}")
    if theme:
        parts.append(f"themeColor={theme}")
    return ", ".join(parts) if parts else "unspecified"


def collect_used_style_ids(story_roots: list[tuple[str, ET.Element]]) -> set[str]:
    style_ids: set[str] = set()
    for _, root in story_roots:
        for style in root.findall(".//w:pStyle", NS) + root.findall(".//w:rStyle", NS):
            style_id = style.get(w_attr("val"))
            if style_id:
                style_ids.add(style_id)
        for p in root.findall(".//w:p", NS):
            if p.find("w:pPr/w:pStyle", NS) is None:
                style_ids.add("Normal")
    return style_ids


def check_font_colors(zf: zipfile.ZipFile, story_roots: list[tuple[str, ET.Element]]) -> list[Finding]:
    bad_colors: list[str] = []

    for source, root in story_roots:
        for color in root.findall(".//w:rPr/w:color", NS):
            if not is_black_font_color(color):
                bad_colors.append(f"{source}: {describe_font_color(color)}")

    styles_root = parse_xml(read_zip_text(zf, "word/styles.xml"))
    if styles_root is not None:
        used_style_ids = collect_used_style_ids(story_roots)
        styles_by_id = {
            style.get(w_attr("styleId")): style
            for style in styles_root.findall(".//w:style", NS)
            if style.get(w_attr("styleId"))
        }
        pending = list(used_style_ids)
        seen: set[str] = set()
        while pending:
            style_id = pending.pop()
            if style_id in seen:
                continue
            seen.add(style_id)
            style = styles_by_id.get(style_id)
            if style is None:
                continue
            for color in style.findall("./w:rPr/w:color", NS):
                if not is_black_font_color(color):
                    bad_colors.append(f"style {style_id}: {describe_font_color(color)}")
            based_on = style.find("./w:basedOn", NS)
            if based_on is not None:
                base_id = based_on.get(w_attr("val"))
                if base_id and base_id not in seen:
                    pending.append(base_id)

    if not bad_colors:
        return []
    examples = "; ".join(dict.fromkeys(bad_colors[:8]))
    return [
        Finding(
            "WARN",
            "font.color",
            f"Found non-black font color(s): {examples}. Use #000000 unless a special requirement explicitly says otherwise.",
        )
    ]


def check_content(texts: list[str], image_count: int) -> list[Finding]:
    findings: list[Finding] = []
    joined = "\n".join(texts)

    required_terms = ["原创性声明", "中文摘要", "目录", "结论", "参考文献"]
    for term in required_terms:
        if term not in joined:
            findings.append(Finding("WARN", "section.missing", f"Could not find required section marker: {term}."))

    if "英文摘要" not in joined and "Abstract" not in joined and "ABSTRACT" not in joined:
        findings.append(Finding("WARN", "section.english-abstract", "Could not find 英文摘要 or Abstract section."))

    keyword_matches = re.findall(r"关键词\s*[:：]\s*([^\n\r]+)", joined)
    if not keyword_matches:
        findings.append(Finding("WARN", "keywords.missing", "Could not find a 关键词 line."))
    else:
        for kw_line in keyword_matches:
            keywords = [kw.strip() for kw in re.split(r"[；;]", kw_line) if kw.strip()]
            if not 3 <= len(keywords) <= 5:
                findings.append(
                    Finding("WARN", "keywords.count", f"关键词 count is {len(keywords)}; expected 3 to 5.")
                )
            if ";" in kw_line:
                findings.append(Finding("WARN", "keywords.separator", "Use Chinese semicolon `；` between keywords."))

    figure_numbers = [int(n) for n in re.findall(r"图\s*([0-9]+)", joined)]
    table_numbers = [int(n) for n in re.findall(r"表\s*([0-9]+)", joined)]
    findings.extend(sequence_warnings("figure", figure_numbers))
    findings.extend(sequence_warnings("table", table_numbers))

    if image_count and len(set(figure_numbers)) < image_count:
        findings.append(
            Finding(
                "WARN",
                "figure.caption.count",
                f"Document has {image_count} embedded image(s) but only {len(set(figure_numbers))} numbered figure caption(s).",
            )
        )

    for bad in ["同上", "同左"]:
        if bad in joined:
            findings.append(Finding("WARN", "table.placeholder", f"Do not use `{bad}` in tables; fill exact values."))

    if "参考文献" in joined:
        refs_text = joined.split("参考文献", 1)[1]
        ref_numbers = re.findall(r"(?:^|\n)\s*(?:\[[0-9]+\]|[0-9]+[.、])", refs_text)
        if len(ref_numbers) < 15:
            findings.append(
                Finding("WARN", "references.count", f"Found about {len(ref_numbers)} reference entries; expected at least 15.")
            )

    if re.search(r"(?<![\[\d])[0-9]{1,2}\](?!\])", joined):
        findings.append(Finding("WARN", "citation.bracket", "Possible citation missing left bracket. Use `[n]`."))

    return findings


def count_images(zf: zipfile.ZipFile, root: ET.Element) -> int:
    embedded_ids = set()
    for blip in root.findall(".//a:blip", NS):
        embed = blip.get(f"{{{NS['r']}}}embed")
        if embed:
            embedded_ids.add(embed)
    media_files = [n for n in zf.namelist() if n.startswith("word/media/")]
    return max(len(embedded_ids), len(media_files))


def validate_docx(path: Path) -> list[Finding]:
    if not path.exists():
        return [Finding("ERROR", "file.missing", f"File not found: {path}")]
    if path.suffix.lower() != ".docx":
        return [Finding("ERROR", "file.type", "Expected a .docx file.")]

    try:
        with zipfile.ZipFile(path) as zf:
            document_xml = read_zip_text(zf, "word/document.xml")
            root = parse_xml(document_xml)
            if root is None:
                return [Finding("ERROR", "docx.document", "Could not read word/document.xml.")]

            story_roots: list[tuple[str, ET.Element]] = []
            for name in zf.namelist():
                if re.fullmatch(r"word/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml", name):
                    story_root = parse_xml(read_zip_text(zf, name))
                    if story_root is not None:
                        story_roots.append((name, story_root))

            findings = check_page_setup(root)
            texts = list(iter_paragraph_text(root))
            findings.extend(check_content(texts, count_images(zf, root)))
            findings.extend(check_font_colors(zf, story_roots))
            return findings
    except zipfile.BadZipFile:
        return [Finding("ERROR", "file.zip", "File is not a valid DOCX zip package.")]


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a DOCX against key UJS thesis format rules.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    parser.add_argument("--strict", action="store_true", help="Exit non-zero when WARN findings are present.")
    args = parser.parse_args()

    findings = validate_docx(args.docx)
    if args.json:
        print(json.dumps([asdict(f) for f in findings], ensure_ascii=False, indent=2))
    else:
        if not findings:
            print("OK: no findings.")
        else:
            for finding in findings:
                print(f"{finding.severity} {finding.code}: {finding.message}")

    has_errors = any(f.severity == "ERROR" for f in findings)
    has_warnings = any(f.severity == "WARN" for f in findings)
    return 1 if has_errors or (args.strict and has_warnings) else 0


if __name__ == "__main__":
    sys.exit(main())
