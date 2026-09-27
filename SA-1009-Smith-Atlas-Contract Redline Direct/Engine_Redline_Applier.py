#!/usr/bin/env python3
# ============================================================================
# Copyright 2026 Smith-Atlas LLC (SA-1009-V1).
# Licensed under the Apache License, Version 2.0 (the "License"); you may not use these files 
# except in compliance with the License. You may obtain a copy of the License at http://
# www.apache.org/licenses/LICENSE-2.0.
#
# Deterministic tracked-changes applier. Takes a .docx and a revisions JSON
# (the schema in the assistant's instructions, Section 5) and returns the SAME
# document with native Word tracked changes and margin comments spliced in.
#
# AI ASSISTANT: run this file VERBATIM. Do not modify, refactor, or reimplement.
#
# Usage:
#   python Engine_Redline_Applier.py <input.docx> <revisions.json> <output.docx> --playbook "Rules_X"
#
# Prints a SLIM JSON summary to stdout (counts, validation, and detail for any
# revision that is not APPLIED) so the assistant can run its verify / repair /
# do-over loop without re-reading successful items. The complete per-revision
# report is written to report.json beside the output file.
# Pure standard library. No third-party dependencies.
#
# CHANGELOG (2026-07-11 — pre-release revision, product version unchanged):
#   * stdout now prints a slim summary; full per-revision report -> report.json.
#   * Validation gate now also enforces the edit-in-place invariant: the
#     reject-all-changes view of the output must reproduce the input document's
#     text paragraph-for-paragraph (nothing outside tracked changes may differ).
#   * Disclaimer comment now reads "qualified attorney or contracting
#     professional," matching the system instructions.
#   * Standalone-clause promotion (SAFETY NET): when an insertion declares no
#     resolvable host anchor, a body carrying its own title (any casing, incl.
#     hyphenated ALL-CAPS) or two or more complete sentences is inserted as its
#     own tracked PARAGRAPH instead of being spliced mid-paragraph into an
#     unrelated section. Untitled promotions get an ACTION NEEDED comment; the
#     engine never invents a heading.
#   * Declared host is honored: an insertion whose context_anchor resolves is
#     spliced into that section regardless of length, so a "provided that..."
#     proviso lands INSIDE the existing liability section (human-attorney
#     convention: amend the existing section, don't spawn a duplicate).
#   * Font inheritance for new sections: heading and body runs now carry the
#     document's DOMINANT body font family and size (voted across body runs),
#     with emphasis toggles stripped and bold re-applied only where the detected
#     heading style calls for it. Previously new sections were emitted with an
#     empty rPr and rendered in Word's default face (Calibri) instead of the
#     contract's. Continuations were already correct (they copy the rPr of the
#     text they modify).
#   * Topic-aware placement: a new section is inserted after the existing
#     section governing the same subject (liability, confidentiality, payment,
#     IP, termination...), falling back to the Entire Agreement / Miscellaneous
#     / signature-block waterfall only when no topical relative exists.
# ============================================================================
import sys, os, re, json, zipfile, shutil, datetime

AUTHOR = "Contract Assistant"
PRODUCT = "Smith-Atlas Contract Redline Direct (SA-1009-V1)"
NOTICE = ("\u00A9 2026 Smith-Atlas LLC (SA-1009-V1). All rights reserved. Use of this tool is "
          "subject to the Terms of Service at https://www.smith-atlas.com/policies/terms-of-service.")
DATE = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

_rev_id = [1000]
def next_id():
    _rev_id[0] += 1
    return _rev_id[0]

# ---------------------------------------------------------------- text utils
def xml_escape(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def xml_unescape(t):
    t = re.sub(r"&#(\d+);", lambda m: chr(int(m.group(1))), t)
    t = re.sub(r"&#x([0-9a-fA-F]+);", lambda m: chr(int(m.group(1), 16)), t)
    return (t.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))

def fold(t):
    """Fold smart quotes/dashes/nbsp for tolerant comparison."""
    t = re.sub(r"[\u2018\u2019\u201B\u0060\u00B4]", "'", t)
    t = re.sub(r"[\u201C\u201D\u201E]", '"', t)
    t = re.sub(r"[\u2013\u2014]", "-", t)
    t = t.replace("\u00A0", " ")
    return t

def norm(t):
    return re.sub(r"\s+", " ", fold(t or "")).strip()

def tolerant_regex(target):
    """Regex matching `target` tolerating quote/dash variants and whitespace runs."""
    out = []
    for ch in fold(target):
        if ch.isspace():
            if out and out[-1] == r"\s+":
                continue
            out.append(r"\s+")
        elif ch == "'":
            out.append("[\u2018\u2019\u201B'`\u00B4]")
        elif ch == '"':
            out.append('[\u201C\u201D\u201E"]')
        elif ch == "-":
            out.append("[-\u2013\u2014]")
        else:
            out.append(re.escape(ch))
    return "".join(out)

def jaccard(a, b):
    A = set(w for w in re.sub(r"[^\w ]", "", norm(a).lower()).split() if len(w) > 2)
    B = set(w for w in re.sub(r"[^\w ]", "", norm(b).lower()).split() if len(w) > 2)
    if not A or not B:
        return 0.0
    return len(A & B) / len(A | B)

def reject_view_texts(xml_doc):
    """Per-paragraph text of the document as it would read if every tracked
    change were REJECTED: <w:ins> content excluded, <w:delText> restored.
    Used by the validation gate to prove non-revised text is untouched."""
    out = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml_doc, re.S):
        p = re.sub(r"<w:ins\b[^>]*>.*?</w:ins>", "", m.group(0), flags=re.S)
        parts = [xml_unescape(t.group(2))
                 for t in re.finditer(r"<w:(t|delText)\b[^>]*>(.*?)</w:\1>", p, re.S)]
        out.append("".join(parts))
    return out

# ------------------------------------------------- ported drafting-side repairs
def repair_trailing_punct(target, redline):
    t, r = target.strip(), redline.strip()
    for mark in (".", ";", ":"):
        if t.endswith(mark) and not r.endswith(mark):
            return redline.rstrip() + mark
    return redline

def normalize_insert_text(t):
    if not t:
        return t
    return re.sub(r"([A-Za-z])([\u2018\u2019'])\s+([sStT])\b", r"\1\2\3", t)

def sentenceize(fragment):
    """Convert a salvaged fragment into a clean standalone sentence."""
    b = re.sub(r"^\s*[;.,]\s*", "", (fragment or "").strip())
    b = re.sub(r"^(and|but|or)\s+", "", b, flags=re.I)
    b = re.sub(r"^([a-z])", lambda m: m.group(1).upper(), b)
    if b and not re.search(r"[.!?]$", b):
        b += "."
    return b

def trim_to_minimal_diff(target, redline):
    """Port of taskpane trimToMinimalDiff: shrink to the words that actually differ.
    Returns dict {target, redline} or {pure_insert, insert_text, side} or None."""
    if target is None or redline is None:
        return {"target": target, "redline": redline}
    if len(redline) > len(target):
        if redline.startswith(target) and redline[len(target):].strip():
            return {"pure_insert": True, "insert_text": redline[len(target):], "side": "after", "target": target}
        if redline.endswith(target) and redline[: len(redline) - len(target)].strip():
            return {"pure_insert": True, "insert_text": redline[: len(redline) - len(target)], "side": "before", "target": target}
    tw = re.split(r"(\s+)", target)
    rw = re.split(r"(\s+)", redline)
    s = 0
    while s < len(tw) and s < len(rw) and tw[s] == rw[s]:
        s += 1
    te, ree = len(tw), len(rw)
    while te > s and ree > s and tw[te - 1] == rw[ree - 1]:
        te -= 1
        ree -= 1
    nt, nr = "".join(tw[s:te]), "".join(rw[s:ree])
    if not nt.strip() and not nr.strip():
        return None
    if not nt.strip():
        side = "before" if s == 0 else "after"
        return {"pure_insert": True, "insert_text": nr, "side": side, "target": target}
    return {"target": nt, "redline": nr}

def looks_like_heading(txt):
    t = (txt or "").strip()
    if not t or len(t.split()) > 9 or re.search(r"[.;:]$", t):
        return False
    return bool(re.match(r"^\s*((\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+)?[A-Z][A-Za-z &/'\-]{2,}$", t))

def strip_leading_section_heading(target):
    """Remove a jammed 'N. Title' fused to the body ('5. Liquidated DamagesIn the event...')."""
    t = str(target or "")
    m = re.match(r"^\s*((?:\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+)([A-Z][A-Za-z]*(?:\s+(?:&|and|of|the|for|to|[A-Z][A-Za-z]*)){0,5})", t)
    if not m:
        return target
    after = t[len(m.group(1)):]
    seam = re.search(r"[a-z][A-Z]", after)
    if not seam or seam.start() > 60:
        return target
    heading, body = after[: seam.start() + 1], after[seam.start() + 1:]
    if not looks_like_heading(heading) or len(body.strip()) < 20:
        return target
    return body.strip()

def largest_single_paragraph_segment(target):
    t = str(target or "")
    if "\n" not in t:
        return t
    segs = [s.strip() for s in re.split(r"\n+", t) if s.strip()]
    if not segs:
        return t
    longest = max(segs, key=len)
    return longest if len(longest) >= 20 else t

def split_embedded_heading(text):
    """'Survival. The provisions...' / 'ANTI-CORRUPTION AND COMPLIANCE. Consultant shall...'
    -> (heading, body). Accepts hyphens, slashes, ampersands, and '.' ':' or '-' delimiters."""
    t = re.sub(r"^\s*[.;,]\s*", "", str(text or "").strip())
    WORD = r"[A-Za-z][A-Za-z]*(?:[-/&][A-Za-z]+)*"          # Anti-Corruption, Indemnification/Defense
    JOIN = r"and|or|of|the|to|&|for"
    m = re.match(r"^(%s(?:\s+(?:%s|%s)){0,6})\s*[.:\u2014\u2013-]\s+(?=[A-Za-z])" % (WORD, WORD, JOIN), t)
    if not m:
        return None
    heading = m.group(1).strip()
    if len(heading) > 60:
        return None
    sig = [w for w in heading.split() if not re.match(r"^(%s)$" % JOIN, w, re.I)]
    if not sig:
        return None
    # Every significant word must be Title-Case or ALL-CAPS (hyphenated parts included).
    def titleish(w):
        parts = re.split(r"[-/&]", w)
        return all(re.match(r"^[A-Z][a-z]*$", p) or re.match(r"^[A-Z]{2,}$", p) for p in parts if p)
    if not all(titleish(w) for w in sig):
        return None
    openers = {"Notwithstanding", "The", "Each", "Customer", "Contractor", "Provider", "Vendor", "Company",
               "Either", "Neither", "This", "All", "Any", "No", "Upon", "During", "Except", "If", "In",
               "Subject", "Unless", "Client", "Consultant", "University", "Party", "Parties"}
    if len(sig) == 1 and sig[0] in openers:
        return None
    body = t[m.end():].strip()
    if len(body) < 20:
        return None
    return (heading, body)

def looks_like_standalone_clause(text):
    """True if an insertion body reads as a self-contained clause that should occupy its own
    paragraph rather than be spliced into an existing one. Structural signals only — word-overlap
    'topic' heuristics were tested and rejected: related legal language frequently shares no
    vocabulary with its host paragraph (a prevailing-party-fees continuation scores the same as
    an unrelated audit-rights clause), so they cannot separate the cases.

    Signal used: multi-sentence bodies. A genuine continuation extends a paragraph with a single
    additional sentence or proviso; a body carrying two or more complete sentences is a clause in
    its own right. Titled bodies are caught earlier by split_embedded_heading()."""
    t = str(text or "").strip()
    sentences = [s for s in re.split(r"(?<=[.!?])\s+(?=[A-Z\"'(])", t) if len(s.strip()) > 25]
    return len(sentences) >= 2

# ---------------------------------------------------------------- document model
RUN_RE = re.compile(r"<w:r\b[^>]*>.*?</w:r>", re.S)
T_RE = re.compile(r"<w:t\b[^>]*>(.*?)</w:t>", re.S)
RPR_RE = re.compile(r"<w:rPr>.*?</w:rPr>", re.S)

class Para:
    """One <w:p> block with a plain-text map over its editable runs."""
    def __init__(self, xml, start, end):
        self.xml, self.start, self.end = xml, start, end
        self.build()

    def build(self):
        # Zones already inside <w:ins>/<w:del> are not editable text.
        forbidden = []
        for m in re.finditer(r"<w:(ins|del)\b[^>]*>.*?</w:\1>", self.xml, re.S):
            forbidden.append((m.start(), m.end()))
        self.runs = []   # (xml_start, xml_end, rpr, text)
        self.text = ""
        for m in RUN_RE.finditer(self.xml):
            if any(a <= m.start() < b for a, b in forbidden):
                continue
            rxml = m.group(0)
            tm = T_RE.search(rxml)
            if not tm:
                continue
            rpr_m = RPR_RE.search(rxml)
            rpr = rpr_m.group(0) if rpr_m else ""
            txt = xml_unescape(tm.group(1))
            self.runs.append({"s": m.start(), "e": m.end(), "rpr": rpr, "text": txt,
                              "off": len(self.text)})
            self.text += txt

    def locate(self, needle):
        """Return (a,b) span in self.text for needle: exact first, then tolerant."""
        if not needle:
            return None
        i = self.text.find(needle)
        if i != -1:
            return (i, i + len(needle))
        m = re.search(tolerant_regex(needle), fold(self.text))
        if m:
            return (m.start(), m.end())
        return None

    def runs_for_span(self, a, b):
        out = []
        for r in self.runs:
            rs, re_ = r["off"], r["off"] + len(r["text"])
            if re_ <= a or rs >= b:
                continue
            out.append((r, max(a, rs) - rs, min(b, re_) - rs))
        return out

def mk_run(rpr, text):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r>'

def mk_del(rpr_text_pairs):
    rid = next_id()
    inner = "".join(f'<w:r>{rpr}<w:delText xml:space="preserve">{xml_escape(t)}</w:delText></w:r>'
                    for rpr, t in rpr_text_pairs if t)
    return f'<w:del w:id="{rid}" w:author="{AUTHOR}" w:date="{DATE}">{inner}</w:del>'

def mk_ins(rpr, text):
    rid = next_id()
    return (f'<w:ins w:id="{rid}" w:author="{AUTHOR}" w:date="{DATE}">'
            f'<w:r>{rpr}<w:t xml:space="preserve">{xml_escape(text)}</w:t></w:r></w:ins>')

def splice(para, a, b, replacement, comment_id=None):
    """Rebuild para.xml turning span (a,b) into tracked del + tracked ins(replacement).
    b==a means pure insertion at a. replacement '' means pure strike."""
    touched = para.runs_for_span(a, b) if b > a else []
    if b > a and not touched:
        return False
    if b > a:
        first_rpr = touched[0][0]["rpr"]
        anchor_run = touched[0][0]
    else:
        host = None
        for r in para.runs:
            if r["off"] <= a <= r["off"] + len(r["text"]):
                host = r
                break
        if host is None:
            return False
        first_rpr = host["rpr"]
        anchor_run = host
        touched = [(host, a - host["off"], a - host["off"])]
    xs = touched[0][0]["s"]
    xe = touched[-1][0]["e"]
    pre = touched[0][0]["text"][: touched[0][1]]
    post = touched[-1][0]["text"][touched[-1][2]:]
    parts = []
    if pre:
        parts.append(mk_run(touched[0][0]["rpr"], pre))
    body = []
    if b > a:
        body.append(mk_del([(r["rpr"], r["text"][s:e]) for r, s, e in touched]))
    if replacement:
        body.append(mk_ins(first_rpr, replacement))
    core = "".join(body)
    if comment_id is not None:
        core = (f'<w:commentRangeStart w:id="{comment_id}"/>' + core +
                f'<w:commentRangeEnd w:id="{comment_id}"/>'
                f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                f'<w:commentReference w:id="{comment_id}"/></w:r>')
    parts.append(core)
    if post:
        parts.append(mk_run(touched[-1][0]["rpr"], post))
    para.xml = para.xml[:xs] + "".join(parts) + para.xml[xe:]
    para.build()
    return True

def para_comment(para, comment_id):
    """Anchor a comment to the whole paragraph (markers as direct children of w:p)."""
    m = re.search(r"(<w:p\b[^>]*>)(\s*(?:<w:pPr>.*?</w:pPr>)?)", para.xml, re.S)
    if not m:
        return False
    ins_at = m.end()
    startmark = f'<w:commentRangeStart w:id="{comment_id}"/>'
    endmark = (f'<w:commentRangeEnd w:id="{comment_id}"/>'
               f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
               f'<w:commentReference w:id="{comment_id}"/></w:r>')
    close = para.xml.rfind("</w:p>")
    para.xml = para.xml[:ins_at] + startmark + para.xml[ins_at:close] + endmark + para.xml[close:]
    para.build()
    return True

# ---------------------------------------------------------------- comments part
class Comments:
    def __init__(self):
        self.items = []  # (id, text)
        self.next = 0

    def add(self, text):
        cid = self.next
        self.next += 1
        self.items.append((cid, text))
        return cid

    def write(self, workdir):
        if not self.items:
            return
        body = "".join(
            f'<w:comment w:id="{cid}" w:author="{AUTHOR}" w:date="{DATE}" w:initials="CA">'
            f'<w:p><w:r><w:t xml:space="preserve">{xml_escape(t)}</w:t></w:r></w:p></w:comment>'
            for cid, t in self.items)
        xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
               + body + "</w:comments>")
        with open(os.path.join(workdir, "word", "comments.xml"), "w", encoding="utf-8") as f:
            f.write(xml)
        ct_path = os.path.join(workdir, "[Content_Types].xml")
        ct = open(ct_path, encoding="utf-8").read()
        if "comments+xml" not in ct:
            ct = ct.replace("</Types>",
                '<Override PartName="/word/comments.xml" ContentType='
                '"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
            open(ct_path, "w", encoding="utf-8").write(ct)
        rel_path = os.path.join(workdir, "word", "_rels", "document.xml.rels")
        rels = open(rel_path, encoding="utf-8").read()
        if "relationships/comments" not in rels:
            rels = rels.replace("</Relationships>",
                '<Relationship Id="rIdComments1009" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
            open(rel_path, "w", encoding="utf-8").write(rels)

def reviewer_comment_text(rev, applied):
    rule = (rev.get("rule_reference") or "Playbook rule").strip()
    issue = (rev.get("issue_summary") or "").strip()
    note = (rev.get("drafting_note") or "").strip()
    if applied:
        return f"[{rule}] {note or issue}"
    lines = [f"ACTION NEEDED \u2014 {rule}"]
    if issue:
        lines.append(f"Issue: {issue}")
    if note:
        lines.append(f"Why: {note}")
    suggested = (rev.get("redlined_text") or rev.get("body_text") or "").strip()
    if suggested:
        lines.append(f'Suggested language: "{suggested}"')
    lines.append("(The exact original text could not be auto-located, so this change was not applied automatically.)")
    return "  ".join(lines)

# ---------------------------------------------------------------- heading style
def detect_body_font(paras):
    """Return an rPr fragment carrying the document's DOMINANT body font family and size, with all
    emphasis toggles (bold/italic/underline/strike/caps) stripped.

    New sections have no existing run to inherit from (unlike continuations, which copy the rPr of
    the text they modify), so without this they are emitted with an empty rPr and Word renders them
    in the document default (usually Calibri) instead of the contract's face. Vote across body runs
    rather than trusting the first one, so a stray heading or signature-block run cannot skew it."""
    font_votes, size_votes = {}, {}
    for p in paras:
        t = (p.text or "").strip()
        if len(t) < 60:               # skip headings, labels, signature lines
            continue
        for m in re.finditer(r"<w:r\b[^>]*>\s*<w:rPr>(.*?)</w:rPr>", p.xml, re.S):
            rpr = m.group(1)
            fm = re.search(r"<w:rFonts\b[^/>]*/?>", rpr)
            if fm:
                font_votes[fm.group(0)] = font_votes.get(fm.group(0), 0) + 1
            sm = re.search(r'<w:sz w:val="(\d+)"\s*/>', rpr)
            if sm:
                size_votes[sm.group(1)] = size_votes.get(sm.group(1), 0) + 1
    parts = []
    if font_votes:
        best = max(font_votes.items(), key=lambda kv: kv[1])[0]
        # Strip any emphasis that rode along inside the rFonts tag (there shouldn't be, but be safe).
        parts.append(best)
    if size_votes:
        sz = max(size_votes.items(), key=lambda kv: kv[1])[0]
        parts.append('<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (sz, sz))
    return "".join(parts)

def detect_heading_style(paras):
    cands, delim_votes, caps, own_line, bold_votes, native = [], {":": 0, ".": 0, "-": 0, "none": 0}, 0, 0, 0, 0
    for p in paras[:60]:
        t = (p.text or "").strip()
        if not t:
            continue
        own = re.match(r"^\s*(\d+|[A-Za-z]|[IVXLCivxlc]+)?\s*[.)]?\s*([A-Z][A-Za-z &/'\-]{1,60})$", t)
        if own and len(t.split()) <= 8 and not re.search(r"[.:;]$", t):
            label, delim = own.group(2).strip(), "none"
            own_line += 1
        else:
            m = re.match(r"^([^.:\-]{2,80}?)\s*([:.\-])", t[:90])
            if not m or len(m.group(1).split()) > 8 or not re.search(r"[A-Za-z]", m.group(1)):
                continue
            if re.match(r"^(name|title|address|e-?mail|date|by)\b", m.group(1), re.I):
                continue
            label, delim = m.group(1).strip(), m.group(2)
        cands.append(p)
        delim_votes[delim] = delim_votes.get(delim, 0) + 1
        if label == label.upper() and re.search(r"[A-Z]", label):
            caps += 1
        if p.runs and ("<w:b/>" in p.runs[0]["rpr"] or "<w:b " in p.runs[0]["rpr"]):
            bold_votes += 1
        if "<w:numPr>" in p.xml.split("</w:pPr>")[0]:
            native += 1
    n = len(cands)
    if n < 2:
        return {"found": False, "allCaps": True, "delimiter": ":", "ownLine": False, "bold": True, "native": False}
    maj = (n + 1) // 2
    best = max(delim_votes, key=lambda k: delim_votes[k])
    return {"found": True, "allCaps": caps >= maj, "delimiter": best,
            "ownLine": own_line >= maj or best == "none", "bold": bold_votes >= maj,
            "native": native >= maj}

# ---------------------------------------------------------------- main engine
def section_topic_host(paras, heading_text, rule_reference, anchor_para_obj=None):
    """Find an EXISTING section whose subject matches a new clause, so the clause lands beside its
    topical relatives instead of being piled in front of the Entire Agreement clause.

    Returns the LAST paragraph of that section's body (insert after it), or None.
    A declared anchor always wins: if the drafter named a host, use it."""
    if anchor_para_obj is not None:
        return anchor_para_obj
    subject = (str(heading_text or "") + " " + str(rule_reference or "")).lower()
    if not subject.strip():
        return None
    # Topic -> regexes matching an existing section heading that covers the same ground.
    TOPICS = [
        (r"indemnif|liabilit|damages|cap\b",      r"liabilit|limitation of liability|indemnif"),
        (r"confidential|non-disclosure|nda",      r"confidential"),
        (r"payment|invoice|late|interest|suspen|accelerat|collection|set-?off",
                                                  r"\bpayment\b|invoic|\bfees\b|compensation"),
        (r"intellectual property|\bip\b|licen|work made for hire|foreground|background",
                                                  r"intellectual property|licens"),
        (r"terminat|wind-?down|transition",       r"terminat"),
        (r"warrant|disclaimer",                   r"warrant|disclaimer|services;\s*disclaimer"),
        (r"insur",                                r"insur"),
        (r"notice",                               r"\bnotices?\b"),
        (r"governing law|venue|jurisdiction|dispute|arbitrat|mediation|jury|limitations period",
                                                  r"governing law|jurisdiction|venue|dispute"),
        (r"assign",                               r"assign"),
        (r"force majeure",                        r"force majeure"),
        (r"independent contractor",               r"independent contractor"),
        (r"scope|change order|sow",               r"\bservices\b|scope"),
        (r"non-?solicit",                         r"non-?solicit|employ"),
    ]
    pat = None
    for subj_re, head_re in TOPICS:
        if re.search(subj_re, subject):
            pat = head_re
            break
    if not pat:
        return None
    # Locate the matching section heading, then return the last paragraph of its body.
    idx = None
    for j, p in enumerate(paras):
        t = p.text.strip()
        if not t:
            continue
        head = t[:90].lower()
        if re.search(pat, head) and (looks_like_heading(t) or re.match(r"^[A-Z][A-Z &;,'\-]{4,}[:.]", t)):
            idx = j
            break
    if idx is None:
        return None
    # Walk forward to the last paragraph before the NEXT section heading.
    last = idx
    for j in range(idx + 1, len(paras)):
        t = paras[j].text.strip()
        if not t:
            continue
        if looks_like_heading(t) or re.match(r"^[A-Z][A-Z &;,'\-]{4,}[:.]", t):
            break
        last = j
    return paras[last]

def apply_revisions(input_docx, revisions, output_docx, playbook):
    workdir = "_engine_work"
    if os.path.isdir(workdir):
        shutil.rmtree(workdir)
    with zipfile.ZipFile(input_docx) as z:
        z.extractall(workdir)
    doc_path = os.path.join(workdir, "word", "document.xml")
    doc = open(doc_path, encoding="utf-8").read()

    # Merge adjacent runs with identical rPr so text is findable.
    def merge(m_p):
        block = m_p.group(0)
        def repl(m2):
            a, b = m2.group(1), m2.group(2)
            ra, rb = RPR_RE.search(a), RPR_RE.search(b)
            if (ra.group(0) if ra else "") != (rb.group(0) if rb else ""):
                return m2.group(0)
            ta, tb = T_RE.search(a), T_RE.search(b)
            if not ta or not tb:
                return m2.group(0)
            merged_t = ta.group(1) + tb.group(1)
            return re.sub(T_RE, lambda _: f'<w:t xml:space="preserve">{merged_t}</w:t>', a, count=1)
        prev = None
        while prev != block:
            prev = block
            block = re.sub(r"(<w:r\b[^>]*>.*?</w:r>)(<w:r\b[^>]*>.*?</w:r>)", repl, block, count=1, flags=re.S)
        return block
    doc = re.sub(r"<w:p\b[^>]*>.*?</w:p>", merge, doc, flags=re.S)

    paras = []
    for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", doc, re.S):
        paras.append(Para(m.group(0), m.start(), m.end()))
    original_texts = [p.text for p in paras]
    original_reject = reject_view_texts(doc)  # snapshot for the invariant gate
    comments = Comments()
    statuses = []

    # Disclaimer comment on the first non-empty paragraph.
    first = next((p for p in paras if p.text.strip()), paras[0] if paras else None)
    if first is not None:
        cid = comments.add(f'DISCLAIMER: The tracked changes and comments below were prepared by the '
                           f'{PRODUCT}, a generative AI tool, using the "{playbook}" playbook. This is not '
                           f'legal advice. All suggestions must be verified by a qualified attorney or '
                           f'contracting professional before use. {NOTICE}')
        para_comment(first, cid)

    def find_para(needle, anchor=None):
        """Locate the paragraph containing needle; disambiguate by anchor."""
        if not needle:
            return None, None
        hits = []
        for p in paras:
            span = p.locate(needle)
            if span:
                hits.append((p, span))
        if not hits:
            return None, None
        if len(hits) == 1 or not anchor:
            return hits[0] if len(hits) == 1 else (None, None) if len(hits) > 1 and not anchor else hits[0]
        best = max(hits, key=lambda h: jaccard(h[0].text, anchor))
        return best

    def anchor_para(anchor):
        if not anchor:
            return None
        frag = norm(anchor)[:120]
        for p in paras:
            if p.locate(frag) or (frag and norm(p.text).find(frag) != -1):
                return p
        words = frag.split()
        if len(words) >= 5:
            head = " ".join(words[:5])
            for p in paras:
                if norm(p.text).find(head) != -1:
                    return p
        return None

    def fallback_comment(rev, host_para):
        host = host_para or first
        cid = comments.add(reviewer_comment_text(rev, applied=False))
        if host is not None:
            para_comment(host, cid)

    def next_body_para(p):
        idx = paras.index(p)
        for q in paras[idx + 1:]:
            if q.text.strip():
                return q
        return p

    heading_style = detect_heading_style(paras)
    body_font = detect_body_font(paras)

    # ------------------------------------------------ per-revision application
    for i, rev in enumerate(revisions):
        et = (rev.get("edit_type") or "").lower().strip()
        if et not in ("surgical", "insertion", "comment"):
            et = "surgical" if (rev.get("target_text") or "").strip() else "insertion"

        # Empty-target surgical salvage -> continuation insertion (behavior 11).
        if et == "surgical" and not (rev.get("target_text") or "").strip():
            salvage = (rev.get("body_text") or rev.get("redlined_text") or "").strip()
            if salvage:
                et = "insertion"
                rev["insertion_style"] = rev.get("insertion_style") or "continuation"
                rev["body_text"] = sentenceize(salvage)
            else:
                fallback_comment(rev, anchor_para(rev.get("context_anchor")))
                statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "COMMENTED",
                                 "detail": "incomplete surgical (no target, nothing to salvage)"})
                continue

        # ---------------- comment-only
        if et == "comment":
            host = anchor_para(rev.get("context_anchor")) or first
            rule = (rev.get("rule_reference") or "Review").strip()
            issue = (rev.get("issue_summary") or rev.get("drafting_note") or "").strip()
            cid = comments.add(f"[{rule}] {issue}")
            para_comment(host, cid)
            statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "APPLIED", "detail": "comment"})
            continue

        # ---------------- insertion
        if et == "insertion":
            style = (rev.get("insertion_style") or "").lower().strip()
            if style not in ("continuation", "new_section"):
                style = "new_section" if (rev.get("heading_text") or "").strip() else "continuation"
            body = normalize_insert_text((rev.get("body_text") or rev.get("redlined_text") or "").strip())
            if not body:
                fallback_comment(rev, anchor_para(rev.get("context_anchor")))
                statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "COMMENTED",
                                 "detail": "blank insertion stub"})
                continue
            # A TITLED body is always its own section — the title is the drafter announcing a
            # standalone clause. A declared anchor then says WHERE that section goes (behavior 17),
            # not whether it is one. This must run regardless of declared_host, or a titled clause
            # aimed at an unrelated paragraph gets buried inside it.
            declared_host = None
            if (rev.get("context_anchor") or "").strip():
                declared_host = anchor_para(rev.get("context_anchor"))
                if declared_host is not None and looks_like_heading(declared_host.text):
                    declared_host = next_body_para(declared_host)

            if style == "continuation" and not (rev.get("heading_text") or "").strip():
                emb = split_embedded_heading(body)
                if emb:
                    rev["heading_text"], body, style = emb[0], emb[1], "new_section"

            # Standalone-clause guard (UNTITLED bodies only, SAFETY NET). An untitled multi-sentence
            # body aimed at a DECLARED host is honored as drafted — that is the human-attorney
            # convention of amending an existing section in place (e.g. a Group E "provided that..."
            # proviso spliced into the existing liability section), and length must not evict it.
            # With no declared host, an untitled multi-sentence clause has nowhere sensible to go
            # inline, so it becomes its own paragraph rather than being spliced arbitrarily.
            promoted_headless = False
            if style == "continuation" and declared_host is None and looks_like_standalone_clause(body):
                style = "new_section"
                promoted_headless = True

            if style == "continuation":
                host = declared_host if declared_host is not None else anchor_para(rev.get("context_anchor"))
                if host is None:
                    fallback_comment(rev, None)
                    statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "COMMENTED",
                                     "detail": "continuation anchor not found"})
                    continue
                if looks_like_heading(host.text):
                    host = next_body_para(host)
                clean = re.sub(r"^\s*[.;,]\s*", "", body.strip())
                if not re.search(r"[.!?]$", clean):
                    clean += "."
                pos = None
                after = (rev.get("after_sentence") or "").strip()
                if after:
                    span = host.locate(after)
                    if span:
                        pos = span[1]
                        seam_prev = host.text[:pos].rstrip()
                        sep = " " if re.search(r"[.!?:;]$", seam_prev) else ". "
                if pos is None:
                    pos = len(host.text)
                    seam_prev = host.text.rstrip()
                    sep = " " if re.search(r"[.!?:;]$", seam_prev) else ". "
                cid = comments.add(reviewer_comment_text(rev, applied=True))
                ok = splice(host, pos, pos, sep + clean, comment_id=cid)
                statuses.append({"i": i, "rule": rev.get("rule_reference"),
                                 "status": "APPLIED" if ok else "FAILED", "detail": "continuation"})
                if not ok:
                    fallback_comment(rev, host)
                continue

            # new_section: anchor waterfall
            label = (rev.get("heading_text") or "").strip()
            label = re.sub(r"[:.\-\s]+$", "", label)
            label = re.sub(r"^\s*(\d+|[A-Za-z])[.)]\s*", "", label)
            headless = promoted_headless or not label
            if heading_style["allCaps"] and not headless:
                label = label.upper()
            target_p = None
            insert_before = False
            # TIER 1: place the new section beside its topical relatives.
            # A declared anchor is trusted ONLY if the drafter genuinely intended a new_section
            # there. When the engine itself split a title out of a continuation body, that anchor
            # pointed at the paragraph the clause was wrongly being spliced into — ignore it and
            # place by topic instead.
            drafter_declared_section = (declared_host if
                                        (rev.get("insertion_style") or "").lower().strip() == "new_section"
                                        else None)
            topic_host = section_topic_host(paras, rev.get("heading_text"),
                                            rev.get("rule_reference"), drafter_declared_section)
            if topic_host is not None:
                target_p = topic_host          # insert AFTER the matching section's body
                insert_before = False
            else:
                # TIER 2: generic anchor waterfall (genuinely miscellaneous clauses).
                for pat in (r"entire agreement|integration|order of precedence",
                            r"^\s*(\d+|[A-Za-z]|[IVXLC]+)?[.)]?\s*(miscellaneous|general)\b",
                            r"signatures? and exhibit", r"in witness whereof"):
                    for p in paras:
                        if re.search(pat, p.text, re.I):
                            target_p = p
                            break
                    if target_p is not None:
                        break
                insert_before = target_p is not None
                if target_p is None:
                    numbered = [p for p in paras if re.match(r"^\s*(\d+|[IVXLC]+)[.)]\s+\S", p.text)]
                    target_p = numbered[-1] if numbered else paras[-1]
            cid = comments.add(reviewer_comment_text(rev, applied=True))
            ppr = ""
            manual_num = True
            if heading_style["native"]:
                npm = re.search(r"<w:numPr>.*?</w:numPr>", target_p.xml.split("</w:pPr>")[0] if "</w:pPr>" in target_p.xml else "", re.S)
                if npm:
                    ppr = f"<w:pPr>{npm.group(0)}<w:rPr><w:ins w:id=\"{next_id()}\" w:author=\"{AUTHOR}\" w:date=\"{DATE}\"/></w:rPr></w:pPr>"
                    manual_num = False
            if not ppr:
                ppr = f"<w:pPr><w:rPr><w:ins w:id=\"{next_id()}\" w:author=\"{AUTHOR}\" w:date=\"{DATE}\"/></w:rPr></w:pPr>"
            hd_rpr = "<w:rPr>%s%s</w:rPr>" % (body_font, "<w:b/>" if heading_style["bold"] else "") \
                     if (body_font or heading_style["bold"]) else ""
            bd_rpr = "<w:rPr>%s</w:rPr>" % body_font if body_font else ""
            marks_start = f'<w:commentRangeStart w:id="{cid}"/>'
            marks_end = (f'<w:commentRangeEnd w:id="{cid}"/>'
                         f'<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                         f'<w:commentReference w:id="{cid}"/></w:r>')
            if headless:
                # Standalone clause with no derivable title: insert as its own tracked paragraph.
                # Never invent a heading — flag it for the reviewer instead.
                new_xml = f"<w:p>{ppr}{marks_start}{mk_ins(bd_rpr, body)}{marks_end}</w:p>"
            elif heading_style["ownLine"]:
                new_xml = (f"<w:p>{ppr}{marks_start}{mk_ins(hd_rpr, label)}{marks_end}</w:p>"
                           f"<w:p>{ppr}{mk_ins(bd_rpr, body)}</w:p>")
            else:
                delim = heading_style["delimiter"] if heading_style["delimiter"] != "none" else ":"
                new_xml = (f"<w:p>{ppr}{marks_start}{mk_ins(hd_rpr, label + delim)}"
                           f"{mk_ins(bd_rpr, ' ' + body)}{marks_end}</w:p>")
            if insert_before:
                target_p.xml = new_xml + target_p.xml
            else:
                target_p.xml = target_p.xml + new_xml
            target_p.build()
            if headless:
                cid3 = comments.add("ACTION NEEDED \u2014 this clause was inserted as its own paragraph because it "
                                    "states a self-contained obligation that did not belong inside the preceding "
                                    "paragraph. Please add a section heading and number matching the contract's "
                                    "scheme, and confirm placement.")
                para_comment(target_p, cid3)
            elif manual_num:
                cid2 = comments.add("ACTION NEEDED \u2014 this contract uses plain (typed) section numbers, so no "
                                    "number was added automatically. Please number this new section per the "
                                    "contract's scheme and update any cross-references.")
                para_comment(target_p, cid2)
            _detail = "new_section (before anchor)" if insert_before else "new_section (appended)"
            if promoted_headless:
                _detail += " [promoted from continuation: unrelated to host paragraph]"
            statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "APPLIED",
                             "detail": _detail})
            continue

        # ---------------- surgical
        target = rev.get("target_text") or ""
        redline = rev.get("redlined_text") or ""
        pure_strike = not redline.strip()
        search = largest_single_paragraph_segment(target)
        search2 = strip_leading_section_heading(search)
        host, span = find_para(search2, rev.get("context_anchor"))
        if host is None and search2 != target:
            host, span = find_para(target, rev.get("context_anchor"))
        if host is None:
            # head-fragment retry
            frag = norm(target)[:80]
            hp = anchor_para(frag)
            if hp is not None:
                sp = hp.locate(search2) or hp.locate(frag)
                if sp:
                    host, span = hp, sp if hp.locate(search2) else None
        if host is None or span is None:
            fallback_comment(rev, anchor_para(rev.get("context_anchor")))
            statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "COMMENTED",
                             "detail": "target not located"})
            continue
        found_text = host.text[span[0]:span[1]]
        if pure_strike:
            cid = comments.add(reviewer_comment_text(rev, applied=True))
            ok = splice(host, span[0], span[1], "", comment_id=cid)
            statuses.append({"i": i, "rule": rev.get("rule_reference"),
                             "status": "APPLIED" if ok else "FAILED", "detail": "pure strike"})
            continue
        redline = repair_trailing_punct(found_text, redline)
        redline = normalize_insert_text(redline)
        if re.match(r"^[A-Z]", found_text.strip()) and re.match(r"^[a-z]", redline.strip()):
            redline = re.sub(r"^(\s*)([a-z])", lambda m: m.group(1) + m.group(2).upper(), redline)
        # bundling guard (behavior 12)
        bundled_extra = None
        tgt = found_text.strip()
        if re.match(r"^[^\s]+@[^\s]+$", tgt) or len(tgt.split()) <= 2:
            enders = len(re.findall(r"[.!?](\s|$)", redline.strip()))
            fsm = re.match(r"^.*?[.!?](\s|$)", redline.strip())
            if enders >= 2 and fsm:
                bundled_extra = redline.strip()[fsm.end():].strip()
                redline = re.sub(r"[.!?]+$", "", fsm.group(0).strip())
        diff = trim_to_minimal_diff(found_text, redline)
        if diff is None:
            statuses.append({"i": i, "rule": rev.get("rule_reference"), "status": "APPLIED",
                             "detail": "no-op (identical)"})
            continue
        cid = comments.add(reviewer_comment_text(rev, applied=True))
        if diff.get("pure_insert"):
            pos = span[0] if diff["side"] == "before" else span[1]
            ok = splice(host, pos, pos, diff["insert_text"], comment_id=cid)
        else:
            off = found_text.find(diff["target"])
            if off == -1:
                off, dt = 0, found_text
            else:
                dt = diff["target"]
            a2, b2 = span[0] + off, span[0] + off + len(dt)
            repl = diff["redline"]
            # trailing-punct double guard (behavior 13b)
            last = repl.rstrip()[-1:] if repl.rstrip() else ""
            if last in ".;:" and not host.text[a2:b2].rstrip().endswith(last) \
               and host.text[b2:b2 + 1] == last:
                repl = re.sub(r"[.;:]+$", "", repl)
            ok = splice(host, a2, b2, repl, comment_id=cid)
        if ok and bundled_extra:
            cid2 = comments.add(f'ACTION NEEDED \u2014 additional requirement (place in the correct clause): "{bundled_extra}"')
            para_comment(host, cid2)
        statuses.append({"i": i, "rule": rev.get("rule_reference"),
                         "status": "APPLIED" if ok else "FAILED", "detail": "surgical"})
        if not ok:
            fallback_comment(rev, host)

    # ------------------------------------------------ [RESERVED] post-pass
    deletions = [r for r in revisions
                 if (r.get("edit_type") or "").lower() == "surgical"
                 and (r.get("target_text") or "").strip()
                 and not (r.get("redlined_text") or "").strip()]
    if deletions:
        # Sections from the ORIGINAL text: inline ("1. TITLE. body...") or own-line headings.
        sections = []
        head_inline = re.compile(r"^\s*(\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+([A-Z][A-Z &/'\-]{2,60}?)[.:]\s+(.*)$", re.S)
        head_own = re.compile(r"^\s*(\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+(.+\S)\s*$")
        cur = None
        for idx, t in enumerate(original_texts):
            t = t or ""
            mi = head_inline.match(t)
            mo = head_own.match(t.strip()) if not mi else None
            if mi:
                if cur:
                    sections.append(cur)
                cur = {"marker": mi.group(1), "title": mi.group(2).strip(), "pidx": idx, "body": mi.group(3)}
            elif mo and len(mo.group(2).split()) <= 8 and not re.search(r"[.;]$", mo.group(2)):
                if cur:
                    sections.append(cur)
                cur = {"marker": mo.group(1), "title": mo.group(2).strip(), "pidx": idx, "body": ""}
            elif cur is not None:
                cur["body"] += (" " if cur["body"] else "") + t.strip()
        if cur:
            sections.append(cur)
        for sec in sections:
            nbody = norm(sec["body"]).lower()
            if not nbody:
                continue
            struck, matched = 0, False
            for d in deletions:
                raw = re.sub(r"^\s*(\d+|[A-Za-z]|[IVXLCivxlc]+)[.)]\s+[^\n]{0,60}\n", "", d.get("target_text") or "")
                raw = strip_leading_section_heading(raw)
                nt = norm(raw).lower()
                if len(nt) >= 15 and nt in nbody:
                    struck = max(struck, len(nt))
                    matched = True
            if not matched:
                continue
            ratio = struck / max(1, len(nbody))
            host = paras[sec["pidx"]]
            if ratio >= 0.85:
                cid = comments.add("ACTION NEEDED \u2014 the entire body of this section has been struck. To finish: "
                                   "replace the section title with \"[RESERVED]\" and delete any remaining struck "
                                   f"body text, while KEEPING the section number (e.g. \"{sec['marker']}. [RESERVED]\") "
                                   "so cross-references stay valid.")
                para_comment(host, cid)
            elif ratio >= 0.40:
                cid = comments.add("ACTION NEEDED \u2014 a substantial part of this section was struck, which may leave "
                                   "incomplete or grammatically broken language. Consider whether this section should "
                                   "be fully reserved (\"[RESERVED]\") or rewritten. Please review.")
                para_comment(host, cid)

    # ------------------------------------------------ reassemble, write, validate
    out_doc = []
    last = 0
    for p in paras:
        out_doc.append(doc[last:p.start])
        out_doc.append(p.xml)
        last = p.end
    out_doc.append(doc[last:])
    new_doc = "".join(out_doc)
    open(doc_path, "w", encoding="utf-8").write(new_doc)
    comments.write(workdir)

    if os.path.exists(output_docx):
        os.remove(output_docx)
    with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(workdir):
            for fn in files:
                fp = os.path.join(root, fn)
                z.write(fp, os.path.relpath(fp, workdir))

    # Validation gate.
    import xml.etree.ElementTree as ET
    problems = []
    with zipfile.ZipFile(output_docx) as z:
        if z.testzip() is not None:
            problems.append("zip integrity failure")
        for name in z.namelist():
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(z.read(name))
                except Exception as e:
                    problems.append(f"{name}: {e}")
    final = open(doc_path, encoding="utf-8").read()

    # Invariant gate: rejecting every tracked change must reproduce the input.
    orig_r = [t for t in original_reject if t.strip()]
    fin_r = [t for t in reject_view_texts(final) if t.strip()]
    invariants = {"original_paragraphs": len(orig_r), "final_paragraphs": len(fin_r),
                  "text_preserved": orig_r == fin_r}
    if not invariants["text_preserved"]:
        k = next((j for j, (a, b) in enumerate(zip(orig_r, fin_r)) if a != b),
                 min(len(orig_r), len(fin_r)))
        a = orig_r[k][:80] if k < len(orig_r) else "<missing>"
        b = fin_r[k][:80] if k < len(fin_r) else "<missing>"
        problems.append(f"invariant: non-revised text changed at paragraph {k}: {a!r} != {b!r}")

    report = {
        "statuses": statuses,
        "counts": {
            "revisions": len(revisions),
            "applied": sum(1 for s in statuses if s["status"] == "APPLIED"),
            "commented": sum(1 for s in statuses if s["status"] == "COMMENTED"),
            "failed": sum(1 for s in statuses if s["status"] == "FAILED"),
            "w_ins": len(re.findall(r"<w:ins ", final)),
            "w_del": len(re.findall(r"<w:del ", final)),
            "comments": len(comments.items),
        },
        "validation": {"ok": not problems, "problems": problems, "invariants": invariants},
        "output": output_docx,
    }
    return report

def main():
    args = sys.argv[1:]
    playbook = "Unknown Playbook"
    if "--playbook" in args:
        k = args.index("--playbook")
        playbook = args[k + 1]
        args = args[:k] + args[k + 2:]
    if len(args) != 3:
        print("Usage: python Engine_Redline_Applier.py <input.docx> <revisions.json> <output.docx> --playbook \"Rules_X\"")
        sys.exit(2)
    input_docx, rev_json, output_docx = args
    with open(rev_json, encoding="utf-8") as f:
        data = json.load(f)
    revisions = data.get("revisions", data) if isinstance(data, dict) else data
    report = apply_revisions(input_docx, revisions, output_docx, playbook)
    report_path = os.path.join(os.path.dirname(os.path.abspath(output_docx)), "report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2)
    slim = {
        "counts": report["counts"],
        "validation": report["validation"],
        "needs_attention": [s for s in report["statuses"] if s["status"] != "APPLIED"],
        "full_report": report_path,
        "output": report["output"],
    }
    print(json.dumps(slim, indent=2))
    sys.exit(0 if report["validation"]["ok"] and report["counts"]["failed"] == 0 else 1)

if __name__ == "__main__":
    main()
