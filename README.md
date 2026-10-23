# Policy Form Compliance Audit

You are a compliance auditor for admitted US insurance products. Your job is to prove that a
system-generated policy document says **exactly** what was filed with and approved by the state
Department of Insurance — nothing missing, nothing added, and the same formatting where
formatting carries legal meaning. If a claim goes to litigation, the issued policy is what a
court reads; any deviation from the filed form can be construed against the carrier
(contra proferentem) and result in monetary loss or a regulatory finding for use of an
unfiled form.

**The filed form is always the source of truth.** Never "correct" the filed form toward the output.

## The three critical checks

| Check | Why it matters | What counts as a defect |
|---|---|---|
| **1. Missing words** | A dropped "not", "any", "or", "except" can reverse coverage. | Any filed word absent from the output. Extra words and changed words (including capitalization changes) are reported too — they are equally non-compliant. |
| **2. Bolding** | Bold words are **defined terms**; courts apply the policy definition only where the term is used as defined. An unbolded **Insured** may be read in its ordinary meaning. | Bold in filed but not in output (BOLD_MISSING); bold in output but not filed (BOLD_ADDED — creates an "undefined defined term"); bolding split within a word (BOLD_PARTIAL). |
| **3. Indentation** | Indentation shows which provisions belong to which lead-in (e.g., whether exclusion "c." is subordinate to "We will not pay for…"), and which exceptions modify which exclusions. A shifted level changes scope. | Different left indent, first-line/hanging indent, list level, or presence of list numbering versus the filed form. |

All three are **CRITICAL**. Never downgrade them to warnings. The only non-critical category is
VARIABLE_FIELD (INFO): a placeholder in the specimen (`XXXXXX`, `____`, `[ ]`, `{ }`, `###`) filled
with policy-specific data — list these so a reviewer can verify the data, but don't fail the form on them.

## Workflow

### 1. Gather and identify the inputs
- **Filed form(s)**: the approved specimen (from SERFF filing, the forms library, or the user). Record the
  form number and edition date (e.g., `CY 00 01 10 26`).
- **Generated document**: the output from the system under test.
- Accept `.docx` and `.pdf`. Convert `.doc`/`.rtf` with `soffice --headless --convert-to docx FILE`.
- If a PDF has no extractable text (scanned), OCR it first (pdf skill) and say so in the report —
  OCR'd text can't prove bolding, so note that bold checks on it are unreliable.
- If the user hasn't said which file is filed and which is generated, ask. Getting this backwards
  inverts every finding.

### 2. Map forms to pages
A generated policy usually bundles the dec page, the coverage form, and several endorsements. Audit
**each form against its own filed specimen**. Find each form's page range in the output by its form
number/edition date in the footer, then pass `--generated-pages` (and `--filed-pages` if the specimen
file holds several forms). A form found in the output with no filed specimen is itself a finding
(possible unfiled form) — report it. A filed form listed on the forms schedule but absent from the
output is also a finding.

### 3. Run the comparison script
Use `scripts/audit_forms.py` from this skill's folder if it exists; otherwise write the copy embedded
at the bottom of this file to `audit_forms.py` and run that. Needs `python-docx`, `pdfplumber`,
`openpyxl` (`pip install --break-system-packages python-docx pdfplumber openpyxl` if missing).

```bash
python3 audit_forms.py --filed FILED.pdf --generated OUTPUT.pdf \
  --generated-pages 7-12 --label "CY 00 01 10 26 Cyber Liability Coverage Form" \
  --out-dir audit_CY0001
```
Options: `--filed-pages`, `--generated-pages` (PDF only, e.g. `1-4,6`), `--indent-tolerance` in points
(default 2.0; keep it tight — do not raise it to make findings go away).

Outputs in `--out-dir`: `findings.json`, `findings.xlsx` (Summary + Findings sheets with a
Disposition dropdown for reviewers), `summary.md`. Exit code 1 means critical findings exist.

How it works, so you can explain it: both documents are tokenized into words with a per-word bold
flag and a per-paragraph (docx) or per-printed-line (pdf) indent; the word streams are aligned with
a diff; unmatched words become wording findings; matched words are compared for bolding; each filed
paragraph/line is mapped to the output paragraph/line its words landed in and the indents compared.
DOCX indent is resolved through direct formatting, list numbering definitions and the style chain,
and leading tab characters are counted (generators often fake indentation with tabs).

### 4. Verify every critical finding before reporting it
The script is a detector, not the verdict. For each CRITICAL finding:
- Open the exact spot in both documents (render the PDF page to an image and look, or read the docx
  paragraph) and confirm it is real.
- **Known false-positive sources** to check for:
  - PDF fonts whose names don't say "Bold" but render bold (or synthetic bold) — look at the page image.
  - Hyphenated words broken across lines in a PDF (`cover-` / `age`) appear as changed words.
  - Headers, footers, page numbers, and form-number footers repeating at page breaks can show up as
    extra/missing words in the middle of text; confirm before reporting, but a *wrong* form number or
    edition date in a footer is a real, critical finding.
  - Docx auto-numbering labels (a., b., 1.) are not extracted as text; list level and numbering presence
    are compared instead. If the labels themselves may differ, check them visually.
  - Cross-format comparisons (docx vs pdf) compare indentation by relative level, not points. Say so.
- If a finding is a false positive, mark it "False positive" with the reason; don't silently drop it.
- Never declare a form compliant because the similarity score is high. One missing word is a failure.

### 5. Report
Deliver:
1. **The audit log workbook** (`findings.xlsx`), one per form or combined, with every finding,
   locations in both documents, the text with the issue marked `[[ ]]`, and the Disposition column.
2. **A short audit summary** in the reply: verdict per form (PASS / FAIL), counts by the three
   categories, and the most consequential findings first — rank by legal impact: negation or
   coverage-grant words missing ("not", "no", "except", "only", "any") > defined-term bolding on
   insuring agreements and exclusions > indentation that changes which lead-in or exception a
   clause belongs to > everything else.
3. For each critical finding, the remediation: what the output must say/look like, quoted exactly
   from the filed form.

Wording for verdicts: use "Not consistent with filed form" / "Consistent with filed form". Don't
offer legal conclusions about coverage outcomes; state the deviation and that it must be corrected
or filed. If the user asks whether a deviation is acceptable, the answer is that only the filed
language may be used unless a revised form is filed and approved.

### Batch / regression use
For a full forms library, run the script per form pair and roll up the Summary rows into one
workbook (form, edition, verdict, counts). For system releases, re-run the same pairs and compare
counts to the prior run so new defects stand out.

## Embedded script: audit_forms.py
Write this verbatim to `audit_forms.py` if `scripts/audit_forms.py` is not available.

```python
#!/usr/bin/env python3
"""
audit_forms.py - Compare a system-generated policy document against the FILED
(approved) form and report three classes of defects:

  1. WORDING  - missing words (critical), plus extra / changed words
  2. BOLDING  - defined terms bold in the filed form but not in the output (and vice versa)
  3. INDENT   - paragraph / line indentation or list level that differs from the filed form

The FILED form is always the source of truth.

Supported inputs: .docx and .pdf (text-based PDFs; scanned PDFs need OCR first).

Usage:
  python3 audit_forms.py --filed FILED.docx|pdf --generated OUTPUT.docx|pdf \
      [--filed-pages 1-4] [--generated-pages 7-10] \
      [--indent-tolerance 2.0] [--out-dir audit_out] [--label "CY 00 01 10 26"]

Outputs (in --out-dir):
  findings.json   - every finding with locations and context
  findings.xlsx   - audit log workbook (Summary + Findings sheets)
  summary.md      - short human-readable summary
Exit code: 0 = no critical findings, 1 = critical findings present, 2 = input error.
"""
import argparse
import difflib
import json
import os
import re
import sys
from collections import Counter, defaultdict

# ----------------------------------------------------------------------------
# Tokenizing
# ----------------------------------------------------------------------------
TOKEN_RE = re.compile(r"[A-Za-z0-9]+(?:['’\-][A-Za-z0-9]+)*|[^\sA-Za-z0-9]")
QUOTE_MAP = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"',
                           "–": "-", "—": "-", " ": " "})
PLACEHOLDER_RE = re.compile(r"^(x{2,}|_+|\[|\]|\{|\}|<|>|#+|\*{2,})$", re.I)


def norm(tok):
    return tok.translate(QUOTE_MAP)


class Token:
    __slots__ = ("text", "bold", "block", "page")

    def __init__(self, text, bold, block, page=None):
        self.text = text
        self.bold = bold          # True / False / "mixed"
        self.block = block        # index into Doc.blocks
        self.page = page


class Block:
    """A paragraph (docx) or a printed line (pdf)."""
    __slots__ = ("idx", "text", "left", "first", "ilvl", "num_id", "page", "where", "level")

    def __init__(self, idx, text, left, first, ilvl=None, num_id=None, page=None, where="body"):
        self.idx, self.text, self.left, self.first = idx, text, left, first
        self.ilvl, self.num_id, self.page, self.where = ilvl, num_id, page, where
        self.level = None


class Doc:
    def __init__(self, path, kind):
        self.path, self.kind = path, kind
        self.blocks, self.tokens = [], []

    def add_block(self, text, char_bold, **kw):
        b = Block(len(self.blocks), text, **kw)
        self.blocks.append(b)
        for m in TOKEN_RE.finditer(text):
            flags = char_bold[m.start():m.end()]
            if all(flags):
                bold = True
            elif any(flags):
                bold = "mixed"
            else:
                bold = False
            self.tokens.append(Token(norm(m.group()), bold, b.idx, kw.get("page")))
        return b


# ----------------------------------------------------------------------------
# DOCX extraction
# ----------------------------------------------------------------------------
def load_docx(path):
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(path)
    out = Doc(path, "docx")
    EMU_PER_PT = 12700

    def twips_attr(el, *names):
        for n in names:
            v = el.get(qn("w:" + n))
            if v is not None:
                try:
                    return int(v) / 20.0
                except ValueError:
                    return None
        return None

    def ind_from_ppr(ppr):
        """Return (left_pt, first_pt) from a w:pPr element, None where unset."""
        if ppr is None:
            return None, None
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            return None, None
        left = twips_attr(ind, "left", "start")
        first = None
        hanging = twips_attr(ind, "hanging")
        firstline = twips_attr(ind, "firstLine")
        if hanging is not None:
            first = -hanging
        elif firstline is not None:
            first = firstline
        return left, first

    # numbering definitions: (numId, ilvl) -> (left, first)
    num_ind = {}
    try:
        numbering = d.part.numbering_part.element
        abstract = {}
        for an in numbering.findall(qn("w:abstractNum")):
            aid = an.get(qn("w:abstractNumId"))
            lv = {}
            for lvl in an.findall(qn("w:lvl")):
                lv[lvl.get(qn("w:ilvl"))] = ind_from_ppr(lvl.find(qn("w:pPr")))
            abstract[aid] = lv
        for num in numbering.findall(qn("w:num")):
            nid = num.get(qn("w:numId"))
            an = num.find(qn("w:abstractNumId"))
            if an is not None:
                for ilvl, ind in abstract.get(an.get(qn("w:val")), {}).items():
                    num_ind[(nid, ilvl)] = ind
    except Exception:
        pass

    def style_chain(style):
        while style is not None:
            yield style
            style = style.base_style

    def para_numbering(p):
        ppr = p._p.pPr
        for el in [ppr] + [s.element.pPr for s in style_chain(p.style) if s.element is not None]:
            if el is None:
                continue
            numpr = el.find(qn("w:numPr"))
            if numpr is not None:
                nid = numpr.find(qn("w:numId"))
                ilvl = numpr.find(qn("w:ilvl"))
                nid = nid.get(qn("w:val")) if nid is not None else None
                ilvl = ilvl.get(qn("w:val")) if ilvl is not None else "0"
                if nid and nid != "0":
                    return nid, ilvl
                return None, None
        return None, None

    def para_indent(p):
        left = first = None
        # style chain (lowest priority, nearest style wins)
        for s in reversed(list(style_chain(p.style))):
            l, f = ind_from_ppr(s.element.pPr if s.element is not None else None)
            left = l if l is not None else left
            first = f if f is not None else first
        nid, ilvl = para_numbering(p)
        if nid:
            l, f = num_ind.get((nid, ilvl), (None, None))
            left = l if l is not None else left
            first = f if f is not None else first
        l, f = ind_from_ppr(p._p.pPr)
        left = l if l is not None else left
        first = f if f is not None else first
        return (left or 0.0), (first or 0.0), nid, ilvl

    def style_bold(p):
        for s in style_chain(p.style):
            if s.font is not None and s.font.bold is not None:
                return bool(s.font.bold)
        return False

    def run_bold(run, pbold):
        if run.bold is not None:
            return bool(run.bold)
        try:
            for s in style_chain(run.style):
                if s.font is not None and s.font.bold is not None:
                    return bool(s.font.bold)
        except Exception:
            pass
        return pbold

    def emit(p, where):
        pbold = style_bold(p)
        text, mask = "", []
        lead_tabs = 0
        # iter runs including those inside hyperlinks / smart tags
        for r_el in p._p.iter(qn("w:r")):
            from docx.text.run import Run
            run = Run(r_el, p)
            b = run_bold(run, pbold)
            for child in r_el:
                tag = child.tag
                if tag == qn("w:t"):
                    t = child.text or ""
                elif tag == qn("w:tab"):
                    t = "\t"
                elif tag in (qn("w:br"), qn("w:cr")):
                    t = " "
                else:
                    continue
                text += t
                mask += [b] * len(t)
        stripped = text.lstrip(" \t")
        lead_tabs = text[: len(text) - len(stripped)].count("\t")
        if not text.strip():
            return
        left, first, nid, ilvl = para_indent(p)
        # leading tabs are a common way generators fake indentation; record them
        # (default tab stop 0.5in = 36pt) so they count toward the visible indent
        first_eff = first + lead_tabs * 36.0
        out.add_block(text, mask, left=left, first=first_eff, ilvl=ilvl, num_id=nid, where=where)

    def walk(parent_el, parent_obj, where):
        for child in parent_el.iterchildren():
            if child.tag == qn("w:p"):
                emit(Paragraph(child, parent_obj), where)
            elif child.tag == qn("w:tbl"):
                t = Table(child, parent_obj)
                for row in t.rows:
                    seen = set()
                    for cell in row.cells:
                        if id(cell._tc) in seen:
                            continue
                        seen.add(id(cell._tc))
                        walk(cell._tc, cell, where + ":table")
            elif child.tag == qn("w:sdt"):
                content = child.find(qn("w:sdtContent"))
                if content is not None:
                    walk(content, parent_obj, where)

    walk(d.element.body, d, "body")
    seen_hf = set()
    for i, sec in enumerate(d.sections):
        for kind, hf in (("header", sec.header), ("footer", sec.footer)):
            try:
                if hf.is_linked_to_previous and i > 0:
                    continue
                key = (kind, hf._element.xml)
                if key in seen_hf:
                    continue
                seen_hf.add(key)
                walk(hf._element, hf, kind)
            except Exception:
                pass
    return out


# ----------------------------------------------------------------------------
# PDF extraction
# ----------------------------------------------------------------------------
BOLD_FONT_RE = re.compile(r"bold|black|heavy|semibold|demi|,b$|-b$|bd", re.I)


def parse_pages(spec, n):
    if not spec:
        return list(range(n))
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            pages.update(range(int(a) - 1, int(b)))
        else:
            pages.add(int(part) - 1)
    return sorted(p for p in pages if 0 <= p < n)


def load_pdf(path, pages_spec=None):
    import pdfplumber
    out = Doc(path, "pdf")
    with pdfplumber.open(path) as pdf:
        page_ids = parse_pages(pages_spec, len(pdf.pages))
        empty = 0
        for pi in page_ids:
            page = pdf.pages[pi]
            chars = [c for c in page.chars if c.get("text", "").strip() or c.get("text") == " "]
            if not [c for c in chars if c["text"].strip()]:
                empty += 1
                continue
            # group chars into lines by baseline
            chars.sort(key=lambda c: (round(c["top"]), c["x0"]))
            lines = []
            for c in chars:
                if lines and abs(c["top"] - lines[-1]["top"]) <= 2.0:
                    lines[-1]["chars"].append(c)
                else:
                    lines.append({"top": c["top"], "chars": [c]})
            for ln in lines:
                cs = sorted(ln["chars"], key=lambda c: c["x0"])
                text, mask, prev = "", [], None
                for c in cs:
                    if prev is not None:
                        gap = c["x0"] - prev["x1"]
                        if gap > max(1.0, c.get("size", 10) * 0.18) and not text.endswith(" ") and c["text"] != " ":
                            text += " "
                            mask.append(False)
                    b = bool(BOLD_FONT_RE.search(c.get("fontname", "") or ""))
                    # synthetic bold (fill+stroke) shows as stroking_color on some generators
                    text += c["text"]
                    mask += [b] * len(c["text"])
                    prev = c
                if not text.strip():
                    continue
                first_visible = next(c for c in cs if c["text"].strip())
                out.add_block(text, mask, left=round(first_visible["x0"], 1), first=0.0,
                              page=pi + 1, where="body")
        if empty and empty == len(page_ids):
            raise SystemExit(f"ERROR: {path} has no extractable text (scanned?). OCR it first.")
    return out


def load(path, pages=None):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        if pages:
            print("note: --pages ignored for .docx", file=s