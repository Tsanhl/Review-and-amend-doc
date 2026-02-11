#!/usr/bin/env python3
"""
Refine a DOCX by copying *textual* amendments from an "amended" DOCX into the
original DOCX while preserving the original's formatting, footnotes, fonts,
sizes, spacing, and layout.

All inserted/replaced wording is marked as **bold + yellow highlight**.

This script intentionally does NOT use python-docx to write the output because
python-docx does not preserve certain Word features (notably footnotes) and can
normalize styles.
"""

from __future__ import annotations

import argparse
import difflib
import re
import sys
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS}


def w_tag(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _t(text: str) -> etree._Element:
    el = etree.Element(w_tag("t"))
    if text.startswith((" ", "\t", "\n")) or text.endswith((" ", "\t", "\n")):
        el.set(f"{{{XML_NS}}}space", "preserve")
    el.text = text
    return el


TOKEN_RE = re.compile(r"\s+|[^\s]+", re.UNICODE)


def _tokenize(s: str) -> List[str]:
    if not s:
        return []
    return TOKEN_RE.findall(s)


def _set_bold_and_highlight_yellow(rPr: etree._Element) -> None:
    b = rPr.find("w:b", namespaces=NS)
    if b is None:
        b = etree.Element(w_tag("b"))
        rPr.append(b)
    else:
        b.set(w_tag("val"), "1")

    highlight = rPr.find("w:highlight", namespaces=NS)
    if highlight is None:
        highlight = etree.Element(w_tag("highlight"))
        rPr.append(highlight)
    highlight.set(w_tag("val"), "yellow")


def _run_rPr(run: etree._Element) -> Optional[etree._Element]:
    return run.find("w:rPr", namespaces=NS)


def _clone_run_with_rPr(src_run: Optional[etree._Element]) -> etree._Element:
    r = etree.Element(w_tag("r"))
    if src_run is not None:
        for k, v in src_run.attrib.items():
            r.set(k, v)
        rPr = _run_rPr(src_run)
        if rPr is not None:
            r.append(deepcopy(rPr))
    return r


def _clone_run_for_changed_text(context_run: Optional[etree._Element], *, markup: bool) -> etree._Element:
    r = _clone_run_with_rPr(context_run)
    if not markup:
        return r
    rPr = r.find("w:rPr", namespaces=NS)
    if rPr is None:
        rPr = etree.Element(w_tag("rPr"))
        r.insert(0, rPr)
    _set_bold_and_highlight_yellow(rPr)
    return r


def _paragraph_text(p: etree._Element) -> str:
    # Include tabs and line breaks so paragraph structure isn't silently changed.
    parts: List[str] = []
    for r in p.xpath("./w:r", namespaces=NS):
        for child in r:
            if child.tag == w_tag("t"):
                parts.append(child.text or "")
            elif child.tag == w_tag("tab"):
                parts.append("\t")
            elif child.tag == w_tag("br"):
                parts.append("\n")
    return "".join(parts)


def _paragraph_text_all_runs(p: etree._Element) -> str:
    # Include text from nested runs (e.g., inside hyperlinks) in document order.
    parts: List[str] = []
    for r in p.xpath(".//w:r", namespaces=NS):
        for child in r:
            if child.tag == w_tag("t"):
                parts.append(child.text or "")
            elif child.tag == w_tag("tab"):
                parts.append("\t")
            elif child.tag == w_tag("br"):
                parts.append("\n")
    return "".join(parts)


def _paragraph_is_simple(p: etree._Element) -> bool:
    # We only rewrite paragraphs that are "run-only" with pPr first.
    children = list(p)
    pPr = p.find("w:pPr", namespaces=NS)
    if pPr is not None and children and children[0] is not pPr:
        # Avoid reordering odd/rare paragraphs where pPr isn't first.
        return False
    # Skip hyperlinks/fields; rewriting them safely needs more logic.
    if p.xpath(".//w:hyperlink|.//w:fldChar|.//w:instrText", namespaces=NS):
        return False
    return True


def _first_textual_run_in_paragraph(p: etree._Element) -> Optional[etree._Element]:
    runs = p.xpath(".//w:r[w:t]", namespaces=NS)
    if runs:
        return runs[0]
    runs = p.xpath(".//w:r", namespaces=NS)
    return runs[0] if runs else None


def _apply_full_replace_to_paragraph(p: etree._Element, new_text: str, *, markup: bool) -> bool:
    old_text = _paragraph_text_all_runs(p)
    if old_text == new_text:
        return False
    context_run = _first_textual_run_in_paragraph(p)
    new_children = _emit_changed_text(new_text, context_run, markup=markup)
    _rewrite_paragraph_in_place(p, new_children)
    return True


@dataclass(frozen=True)
class Atom:
    kind: str  # text|tab|br|fn|endnote|run_special|p_special
    start: int
    end: int
    run: Optional[etree._Element]
    elem: etree._Element
    text: str = ""


def _paragraph_atoms(p: etree._Element) -> Tuple[List[Atom], str]:
    atoms: List[Atom] = []
    pos = 0
    for child in p:
        if child.tag == w_tag("pPr"):
            continue
        if child.tag != w_tag("r"):
            atoms.append(Atom("p_special", pos, pos, None, child))
            continue

        run = child
        for rc in run:
            if rc.tag == w_tag("rPr"):
                continue
            if rc.tag == w_tag("t"):
                txt = rc.text or ""
                if txt:
                    atoms.append(Atom("text", pos, pos + len(txt), run, rc, txt))
                    pos += len(txt)
                else:
                    atoms.append(Atom("text", pos, pos, run, rc, ""))
            elif rc.tag == w_tag("tab"):
                atoms.append(Atom("tab", pos, pos + 1, run, rc, "\t"))
                pos += 1
            elif rc.tag == w_tag("br"):
                atoms.append(Atom("br", pos, pos + 1, run, rc, "\n"))
                pos += 1
            elif rc.tag == w_tag("footnoteReference"):
                atoms.append(Atom("fn", pos, pos, run, rc))
            elif rc.tag == w_tag("endnoteReference"):
                atoms.append(Atom("endnote", pos, pos, run, rc))
            else:
                atoms.append(Atom("run_special", pos, pos, run, rc))
    old_text = "".join(a.text for a in atoms if a.kind in ("text", "tab", "br"))
    return atoms, old_text


def _context_run_for_pos(atoms: List[Atom], pos: int) -> Optional[etree._Element]:
    # Prefer body-text runs so inserted text does not inherit footnote/endnote
    # superscript styling when edits occur adjacent to references.
    for a in atoms:
        if a.run is not None and a.kind in ("text", "tab", "br") and a.start <= pos < a.end:
            return a.run
        if a.run is not None and a.kind == "text" and a.start == pos and a.end == pos:
            return a.run

    # If no direct hit, prefer the nearest textual run behind the insertion
    # point, then the nearest textual run ahead.
    for a in reversed(atoms):
        if a.run is not None and a.kind in ("text", "tab", "br") and a.end <= pos:
            return a.run
    for a in atoms:
        if a.run is not None and a.kind in ("text", "tab", "br") and a.start >= pos:
            return a.run

    # Last resort for non-text-only paragraphs.
    for a in reversed(atoms):
        if a.run is not None:
            return a.run
    return None


def _emit_atom(atom: Atom, *, text_override: Optional[str] = None) -> etree._Element:
    if atom.kind == "p_special":
        return deepcopy(atom.elem)

    run = _clone_run_with_rPr(atom.run)
    if atom.kind == "text":
        run.append(_t(text_override if text_override is not None else atom.text))
    elif atom.kind == "tab":
        run.append(deepcopy(atom.elem))
    elif atom.kind == "br":
        run.append(deepcopy(atom.elem))
    else:
        run.append(deepcopy(atom.elem))
    return run


def _emit_changed_text(text: str, context_run: Optional[etree._Element], *, markup: bool) -> List[etree._Element]:
    out: List[etree._Element] = []
    if not text:
        return out

    # Preserve tabs/line breaks as Word elements, not literal characters.
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "\t":
            r = _clone_run_for_changed_text(context_run, markup=markup)
            r.append(etree.Element(w_tag("tab")))
            out.append(r)
            i += 1
            continue
        if ch == "\n":
            r = _clone_run_for_changed_text(context_run, markup=markup)
            r.append(etree.Element(w_tag("br")))
            out.append(r)
            i += 1
            continue

        j = i
        while j < len(text) and text[j] not in ("\t", "\n"):
            j += 1
        chunk = text[i:j]
        r = _clone_run_for_changed_text(context_run, markup=markup)
        r.append(_t(chunk))
        out.append(r)
        i = j
    return out


def _rewrite_paragraph_in_place(p: etree._Element, new_children: List[etree._Element]) -> None:
    pPr = p.find("w:pPr", namespaces=NS)
    for child in list(p):
        if child is pPr:
            continue
        p.remove(child)
    if pPr is None:
        for c in new_children:
            p.append(c)
        return
    insert_at = list(p).index(pPr) + 1
    for c in new_children:
        p.insert(insert_at, c)
        insert_at += 1


def _apply_diff_to_paragraph(p: etree._Element, new_text: str, *, markup: bool) -> bool:
    atoms, old_text = _paragraph_atoms(p)
    if old_text == new_text:
        return False

    old_tokens = _tokenize(old_text)
    new_tokens = _tokenize(new_text)

    sm = difflib.SequenceMatcher(a=old_tokens, b=new_tokens, autojunk=False)
    opcodes = sm.get_opcodes()

    # Precompute token -> char offsets.
    old_tok_starts: List[int] = []
    cur = 0
    for tok in old_tokens:
        old_tok_starts.append(cur)
        cur += len(tok)
    old_total = cur

    new_tok_starts: List[int] = []
    cur = 0
    for tok in new_tokens:
        new_tok_starts.append(cur)
        cur += len(tok)

    def old_char_range(i1: int, i2: int) -> Tuple[int, int]:
        if i1 == i2:
            return (old_tok_starts[i1] if i1 < len(old_tok_starts) else old_total, old_tok_starts[i1] if i1 < len(old_tok_starts) else old_total)
        start = old_tok_starts[i1]
        end = old_tok_starts[i2 - 1] + len(old_tokens[i2 - 1])
        return start, end

    def new_char_range(j1: int, j2: int) -> Tuple[int, int]:
        if j1 == j2:
            at = new_tok_starts[j1] if j1 < len(new_tok_starts) else len(new_text)
            return at, at
        start = new_tok_starts[j1]
        end = new_tok_starts[j2 - 1] + len(new_tokens[j2 - 1])
        return start, end

    emitted_specials: set[int] = set()

    def emit_old_segment(start: int, end: int, *, include_text: bool) -> List[etree._Element]:
        out: List[etree._Element] = []
        for a in atoms:
            if a.kind == "text":
                if not include_text:
                    continue
                if a.end <= start or a.start >= end:
                    continue
                s = max(start, a.start)
                e = min(end, a.end)
                if s >= e:
                    continue
                slice_text = a.text[s - a.start : e - a.start]
                out.append(_emit_atom(a, text_override=slice_text))
            elif a.kind in ("tab", "br"):
                if not include_text:
                    continue
                if a.start >= start and a.end <= end:
                    out.append(_emit_atom(a))
            else:
                # Zero-length or anchored elements: preserve them even when text is replaced/deleted.
                if start <= a.start <= end and id(a.elem) not in emitted_specials:
                    emitted_specials.add(id(a.elem))
                    out.append(_emit_atom(a))
        return out

    new_children: List[etree._Element] = []
    for tag, i1, i2, j1, j2 in opcodes:
        o_start, o_end = old_char_range(i1, i2)
        n_start, n_end = new_char_range(j1, j2)

        if tag == "equal":
            new_children.extend(emit_old_segment(o_start, o_end, include_text=True))
            continue

        context_run = _context_run_for_pos(atoms, o_start)
        inserted = new_text[n_start:n_end]

        if tag in ("replace", "insert"):
            new_children.extend(_emit_changed_text(inserted, context_run, markup=markup))

        # Preserve anchored elements that were in the replaced/deleted old range (not its old text).
        if tag in ("replace", "delete"):
            new_children.extend(emit_old_segment(o_start, o_end, include_text=False))

    _rewrite_paragraph_in_place(p, new_children)
    return True


def _iter_body_paragraphs(doc_root: etree._Element) -> List[etree._Element]:
    return doc_root.xpath("/w:document/w:body//w:p", namespaces=NS)


def _load_docx_xml(path: Path, part: str) -> etree._Element:
    with zipfile.ZipFile(path, "r") as zf:
        data = zf.read(part)
    return etree.fromstring(data)


def _write_docx_with_replaced_part(original_path: Path, out_path: Path, part: str, xml_root: etree._Element) -> None:
    xml_bytes = etree.tostring(
        xml_root,
        encoding="UTF-8",
        xml_declaration=True,
        standalone=False,
    )
    with zipfile.ZipFile(original_path, "r") as zin:
        with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == part:
                    data = xml_bytes
                zout.writestr(info, data)


def refine_from_amended(original_docx: Path, amended_docx: Path, out_docx: Path, *, markup: bool = True) -> Tuple[int, int]:
    part = "word/document.xml"
    orig_root = _load_docx_xml(original_docx, part)
    amend_root = _load_docx_xml(amended_docx, part)

    orig_paras = _iter_body_paragraphs(orig_root)
    amend_paras = _iter_body_paragraphs(amend_root)

    if len(orig_paras) != len(amend_paras):
        raise ValueError(
            f"Paragraph count mismatch: original has {len(orig_paras)}, amended has {len(amend_paras)}. "
            "Export the amended DOCX so it preserves paragraph breaks, or extend the script to allow structural edits."
        )

    changed = 0
    skipped = 0
    for p_orig, p_amend in zip(orig_paras, amend_paras):
        new_text = _paragraph_text_all_runs(p_amend)
        if _paragraph_is_simple(p_orig):
            if _apply_diff_to_paragraph(p_orig, new_text, markup=markup):
                changed += 1
        else:
            # Complex paragraphs (e.g., with hyperlinks/fields): use full-text
            # replacement with local-style inheritance and additive markup.
            if _apply_full_replace_to_paragraph(p_orig, new_text, markup=markup):
                changed += 1
            else:
                skipped += 1

    _write_docx_with_replaced_part(original_docx, out_docx, part, orig_root)
    return changed, skipped


def _split_amended_text_to_paragraphs(text: str) -> List[str]:
    # Split on blank lines (>=2 newlines). Keep internal single newlines as-is.
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    if not cleaned:
        return []
    return re.split(r"\n{2,}", cleaned)


def refine_from_amended_text(
    original_docx: Path, amended_text_path: Path, out_docx: Path, *, markup: bool = True
) -> Tuple[int, int]:
    part = "word/document.xml"
    orig_root = _load_docx_xml(original_docx, part)
    orig_paras = _iter_body_paragraphs(orig_root)

    amended_text = amended_text_path.read_text(encoding="utf-8")
    amended_paras = _split_amended_text_to_paragraphs(amended_text)

    if len(orig_paras) != len(amended_paras):
        raise ValueError(
            f"Paragraph count mismatch: original DOCX has {len(orig_paras)}, amended text has {len(amended_paras)}. "
            "Ensure the amended text preserves paragraph breaks using blank lines between paragraphs."
        )

    changed = 0
    skipped = 0
    for p_orig, new_text in zip(orig_paras, amended_paras):
        if _paragraph_is_simple(p_orig):
            if _apply_diff_to_paragraph(p_orig, new_text, markup=markup):
                changed += 1
        else:
            if _apply_full_replace_to_paragraph(p_orig, new_text, markup=markup):
                changed += 1
            else:
                skipped += 1

    _write_docx_with_replaced_part(original_docx, out_docx, part, orig_root)
    return changed, skipped


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Copy textual amendments from an amended DOCX into an original DOCX while preserving the original styling. Changes are bold + yellow highlighted."
    )
    ap.add_argument("--original", required=True, type=Path, help="Path to the user's original DOCX (style source).")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--amended", type=Path, help="Path to the amended DOCX (text source).")
    src.add_argument(
        "--amended-txt",
        type=Path,
        help="Path to a UTF-8 text file containing amended content, split into paragraphs by blank lines.",
    )
    ap.add_argument(
        "--out",
        type=Path,
        help="Output DOCX path (default: <original>_refined.docx).",
    )
    ap.add_argument(
        "--no-markup",
        action="store_true",
        help="Do not apply bold/yellow highlighting to changed text (keeps inserted text in the surrounding style).",
    )
    args = ap.parse_args(argv)

    if args.out is None:
        args.out = args.original.with_name(args.original.stem + "_refined.docx")

    markup = not args.no_markup
    if args.amended is not None:
        changed, skipped = refine_from_amended(args.original, args.amended, args.out, markup=markup)
    else:
        changed, skipped = refine_from_amended_text(args.original, args.amended_txt, args.out, markup=markup)
    print(f"Wrote: {args.out}")
    print(f"Paragraphs updated: {changed}")
    if skipped:
        print(f"Paragraphs skipped (complex structures): {skipped}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
