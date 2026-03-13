#!/usr/bin/env python3
"""PPTX element overlap & boundary checker.

Parses slide XML to extract element bounding boxes, then detects:
  1. Overlapping elements (intersection area > threshold)
  2. Insufficient vertical gaps between adjacent elements
  3. Elements exceeding slide boundaries

Card-aware: groups elements inside the same card container to avoid
false positives from title/body text boxes overlapping their parent shape.

Usage:
  python check_overlaps.py input.pptx [--json report.json] [--min-overlap 0.01] [--min-gap 0.05]

Exit codes: 0 = clean, 1 = critical/major issues found.
"""
import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

try:
    from lxml import etree
except ImportError:
    print("ERROR: lxml required. Install: pip install lxml", file=sys.stderr)
    sys.exit(1)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

EMU_PER_INCH = 914400
SW, SH, MARGIN = 10.0, 5.625, 0.5


def emu_to_inch(emu: int) -> float:
    return emu / EMU_PER_INCH


@dataclass
class BBox:
    x: float
    y: float
    w: float
    h: float
    label: str = ""
    slide: int = 0
    has_text: bool = False
    has_fill: bool = False
    group_id: int = -1  # container group assignment

    @property
    def right(self):
        return self.x + self.w

    @property
    def bottom(self):
        return self.y + self.h

    @property
    def area(self):
        return self.w * self.h


def _extract_text(sp) -> str:
    parts = []
    for t in sp.iterfind(".//a:t", NS):
        if t.text:
            parts.append(t.text)
    text = "".join(parts).strip().replace("\n", " ")
    return text[:50] if text else ""


def _has_solid_fill(sp) -> bool:
    """Check if a shape has a solid fill (container rectangle)."""
    return sp.find(".//p:spPr/a:solidFill", NS) is not None


def _extract_bbox(sp, offset_x=0.0, offset_y=0.0) -> BBox | None:
    xfrm = sp.find(".//p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = sp.find(".//a:xfrm", NS)
    if xfrm is None:
        return None

    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None

    x = emu_to_inch(int(off.get("x", "0"))) + offset_x
    y = emu_to_inch(int(off.get("y", "0"))) + offset_y
    w = emu_to_inch(int(ext.get("cx", "0")))
    h = emu_to_inch(int(ext.get("cy", "0")))

    if w < 0.01 or h < 0.01:
        return None

    is_shape = sp.tag.endswith("}sp")
    label = _extract_text(sp) if is_shape else "[image]"
    has_text = bool(label) if is_shape else False
    has_fill = _has_solid_fill(sp) if is_shape else False

    return BBox(x=x, y=y, w=w, h=h, label=label, has_text=has_text, has_fill=has_fill)


def _is_background(b: BBox) -> bool:
    if b.y < 0.05 and b.w > 9.5 and b.h <= 0.85:
        return True
    if b.y > 4.5 and b.w > 9.5:
        return True
    return False


def _is_accent_bar(b: BBox) -> bool:
    return b.w <= 0.08


def _is_badge(b: BBox) -> bool:
    return b.w <= 0.4 and b.h <= 0.4 and b.area < 0.15


def _contains(outer: BBox, inner: BBox, tol=0.05) -> bool:
    return (inner.x >= outer.x - tol and
            inner.y >= outer.y - tol and
            inner.right <= outer.right + tol and
            inner.bottom <= outer.bottom + tol)


def _overlap_area(a: BBox, b: BBox) -> float:
    ix = max(0, min(a.right, b.right) - max(a.x, b.x))
    iy = max(0, min(a.bottom, b.bottom) - max(a.y, b.y))
    return ix * iy


def _assign_container_groups(elements: list[BBox]) -> None:
    """Group elements into card containers.

    A container is a filled shape (has_fill=True, no text) that is large enough
    to hold child elements. Elements fully inside a container get the same group_id.
    Elements sharing a group_id won't be checked against each other for overlaps.
    """
    # Find potential containers: filled shapes without text, area > 0.5 sq.in
    containers = []
    for i, e in enumerate(elements):
        if e.has_fill and not e.has_text and e.area > 0.5 and not _is_background(e):
            containers.append((i, e))

    group_counter = 0
    for ci, container in containers:
        members = [ci]
        for j, e in enumerate(elements):
            if j == ci:
                continue
            if _contains(container, e):
                members.append(j)
        if len(members) >= 2:  # container + at least one child
            for idx in members:
                elements[idx].group_id = group_counter
            group_counter += 1


def _is_full_dark_slide(slide_xml: bytes) -> bool:
    """Detect title/section/thank-you slides (full dark background, no title bar).

    These slides use full-bleed backgrounds and don't follow content slide boundary rules.
    Detected by: no content-area title bar (full-width rect at y=0 with h≤0.8 and dark fill).
    Instead they have a slide master background covering the whole slide.
    """
    tree = etree.fromstring(slide_xml)
    sp_tree = tree.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return False

    has_title_bar = False
    for sp in sp_tree.findall("p:sp", NS):
        bb = _extract_bbox(sp)
        if bb and bb.y < 0.05 and bb.w > 9.5 and bb.h <= 0.85 and bb.h > 0.3:
            has_title_bar = True
            break

    # If no title bar, it's likely a full-dark slide (title/section/thank-you)
    return not has_title_bar


def extract_elements(slide_xml: bytes, slide_num: int) -> list[BBox]:
    tree = etree.fromstring(slide_xml)
    elements = []

    def collect(parent_path, ox=0.0, oy=0.0):
        for sp in parent_path.findall("p:sp", NS):
            bb = _extract_bbox(sp, ox, oy)
            if bb:
                bb.slide = slide_num
                elements.append(bb)
        for pic in parent_path.findall("p:pic", NS):
            bb = _extract_bbox(pic, ox, oy)
            if bb:
                bb.slide = slide_num
                bb.label = bb.label or "[image]"
                elements.append(bb)
        for grp in parent_path.findall("p:grpSp", NS):
            grp_xfrm = grp.find("p:grpSpPr/a:xfrm", NS)
            gox, goy = ox, oy
            if grp_xfrm is not None:
                goff = grp_xfrm.find("a:off", NS)
                if goff is not None:
                    gox += emu_to_inch(int(goff.get("x", "0")))
                    goy += emu_to_inch(int(goff.get("y", "0")))
            collect(grp, gox, goy)

    sp_tree = tree.find("p:cSld/p:spTree", NS)
    if sp_tree is not None:
        collect(sp_tree)

    _assign_container_groups(elements)
    return elements


@dataclass
class Issue:
    slide: int
    severity: str
    kind: str
    message: str
    details: dict = field(default_factory=dict)


def check_slide(elements: list[BBox], slide_num: int,
                min_overlap: float, min_gap: float,
                is_dark_slide: bool = False) -> list[Issue]:
    issues = []

    content = [e for e in elements
               if not _is_background(e) and not _is_accent_bar(e)]

    # 1. Boundary check (skip for title/section/thank-you slides — full-bleed design)
    if not is_dark_slide:
        for e in content:
            if e.bottom > SH - MARGIN + 0.05:
                issues.append(Issue(
                    slide=slide_num, severity="CRITICAL", kind="boundary",
                    message=f"Bottom overflow: bottom={e.bottom:.2f}\" > limit={SH - MARGIN:.2f}\"",
                    details={"label": e.label, "y": round(e.y, 3), "h": round(e.h, 3),
                              "bottom": round(e.bottom, 3)}
                ))

    # Elements eligible for overlap/gap checks
    check_elems = [e for e in content if not _is_badge(e)]

    # 2. Overlap check — skip same-group pairs (intra-card)
    for i, a in enumerate(check_elems):
        for b in check_elems[i + 1:]:
            # Same container group → skip
            if a.group_id >= 0 and a.group_id == b.group_id:
                continue
            area = _overlap_area(a, b)
            if area < min_overlap:
                continue
            # One fully contains the other → container-content
            if _contains(a, b) or _contains(b, a):
                continue
            smaller_area = min(a.area, b.area)
            if smaller_area < 0.001:
                continue
            ratio = area / smaller_area
            if ratio < 0.05:
                continue

            sev = "CRITICAL" if ratio > 0.30 else "MAJOR"
            issues.append(Issue(
                slide=slide_num, severity=sev, kind="overlap",
                message=(f"Overlap ({ratio:.0%}): "
                         f"\"{a.label[:30]}\" vs \"{b.label[:30]}\""),
                details={
                    "elem_a": {"label": a.label, "y": round(a.y, 3), "bottom": round(a.bottom, 3)},
                    "elem_b": {"label": b.label, "y": round(b.y, 3), "bottom": round(b.bottom, 3)},
                    "overlap_ratio": round(ratio, 3),
                }
            ))

    # 3. Vertical gap check — only between top-level elements (different groups or ungrouped)
    top_level = [e for e in check_elems
                 if e.group_id < 0 or e.has_fill]  # containers represent their group
    sorted_elems = sorted(top_level, key=lambda e: (e.y, e.x))
    for i, a in enumerate(sorted_elems):
        for b in sorted_elems[i + 1:]:
            h_overlap = max(0, min(a.right, b.right) - max(a.x, b.x))
            if h_overlap < 0.3:
                continue
            gap_val = b.y - a.bottom
            if gap_val < 0 or gap_val >= min_gap:
                continue
            # Same group → skip
            if a.group_id >= 0 and a.group_id == b.group_id:
                continue
            issues.append(Issue(
                slide=slide_num, severity="WARNING", kind="gap",
                message=(f"Tight gap {gap_val:.3f}\": "
                         f"\"{a.label[:30]}\" (bot={a.bottom:.2f}\") → "
                         f"\"{b.label[:30]}\" (y={b.y:.2f}\")"),
                details={
                    "gap_inches": round(gap_val, 3),
                    "elem_above": {"label": a.label, "bottom": round(a.bottom, 3)},
                    "elem_below": {"label": b.label, "y": round(b.y, 3)},
                }
            ))

    return issues


def check_pptx(pptx_path: str, min_overlap=0.01, min_gap=0.05) -> list[Issue]:
    all_issues = []
    slide_re = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
    with zipfile.ZipFile(pptx_path, "r") as zf:
        slide_names = sorted(
            (n for n in zf.namelist() if slide_re.match(n)),
            key=lambda n: int(slide_re.match(n).group(1))
        )
        for name in slide_names:
            num = int(slide_re.match(name).group(1))
            xml_bytes = zf.read(name)
            is_dark = _is_full_dark_slide(xml_bytes)
            elements = extract_elements(xml_bytes, num)
            issues = check_slide(elements, num, min_overlap, min_gap, is_dark)
            all_issues.extend(issues)
    return all_issues


def print_report(issues: list[Issue]) -> None:
    if not issues:
        print("OK: no overlap or boundary issues found.")
        return

    by_slide = {}
    for iss in issues:
        by_slide.setdefault(iss.slide, []).append(iss)

    counts = {"CRITICAL": 0, "MAJOR": 0, "WARNING": 0}
    for iss in issues:
        counts[iss.severity] += 1

    print(f"\n{'='*60}")
    print(f"  OVERLAP CHECK: {counts['CRITICAL']} critical, "
          f"{counts['MAJOR']} major, {counts['WARNING']} warnings")
    print(f"{'='*60}\n")

    for slide_num in sorted(by_slide):
        print(f"  Slide {slide_num}:")
        for iss in by_slide[slide_num]:
            marker = {"CRITICAL": "!!!", "MAJOR": " ! ", "WARNING": " ~ "}[iss.severity]
            print(f"    [{marker}] {iss.severity}: {iss.message}")
        print()


def main():
    parser = argparse.ArgumentParser(description="PPTX overlap & boundary checker")
    parser.add_argument("pptx", help="Input .pptx file")
    parser.add_argument("--json", help="Write JSON report to file", default=None)
    parser.add_argument("--min-overlap", type=float, default=0.01,
                        help="Minimum overlap area (sq.in) to report (default: 0.01)")
    parser.add_argument("--min-gap", type=float, default=0.05,
                        help="Minimum vertical gap (in) before warning (default: 0.05)")
    args = parser.parse_args()

    if not Path(args.pptx).exists():
        print(f"ERROR: file not found: {args.pptx}", file=sys.stderr)
        sys.exit(1)

    issues = check_pptx(args.pptx, args.min_overlap, args.min_gap)
    print_report(issues)

    if args.json:
        report = [{"slide": i.slide, "severity": i.severity, "kind": i.kind,
                    "message": i.message, **i.details} for i in issues]
        Path(args.json).write_text(json.dumps(report, indent=2, ensure_ascii=False))
        print(f"  JSON report: {args.json}")

    has_serious = any(i.severity in ("CRITICAL", "MAJOR") for i in issues)
    sys.exit(1 if has_serious else 0)


if __name__ == "__main__":
    main()
