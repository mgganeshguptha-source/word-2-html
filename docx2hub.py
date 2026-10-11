#!/usr/bin/env python3
"""
docx2hub.py — convert the BCBSM authoring Word document into the Knowledge Hub HTML page.

Usage:
    python docx2hub.py input.docx output.html
    python docx2hub.py input.docx output.html --check     (validate only, write nothing)

Contract (decided by STRUCTURE, never by guessing):
    Heading 1                      -> new section + sidebar nav entry
    Heading 2 "Group: <name>"      -> nav group for the sections that follow
    First normal paragraph(s)      -> section lead text
    Table  Value | Label           -> stat boxes
    Table  Title | Description     -> card grid
    Table  Title | Description | Accent          -> cards with accent border
    Table  Title | Description | URL | Type      -> document link cards
    Table  Term  | Definition      -> glossary
    Table  Narration               -> spoken script for the Listen button
    Table  Field | Value           -> page header (hero) config
                                      Title, Tagline, Eyebrow, Brand, Brand subtitle,
                                      Pill 1..n, Stat 1..4 ("20 | guided topics"), Footer, Logo
    Numbered / List Paragraph      -> auto-numbered step cards
    Style 'Intense Quote'          -> orange warning callout
    Style 'Subtle Emphasis'        -> small italic footnote
"""

import sys
import re
import io
import json
import html
import base64
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
from email.utils import parsedate_to_datetime
from pathlib import Path

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("python-docx is required:  pip install python-docx")

try:
    from PIL import Image
except ImportError:
    Image = None   # downscaling disabled; images still embed at original size


# ─────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────

CONTENT_MARKER = "PAGE CONTENT STARTS BELOW THIS LINE"

# Author-facing instruction text that must never reach the page.
INSTRUCTION_PATTERNS = [
    r"^orange boxes like this one",
    r"^the header row is what tells",
    r"^accent values allowed",
    r"^this special section builds",
    r"^the numbered list below becomes",
    r"^this is the table from your question",
    r"^the link in the paragraph above",
    r"^a second title \| description table",
    r"^the paragraph below uses",
    r"^to add a link, add a row",
    r"^this whole section is an example",
    r"^you can also just paste a bare url",
    r"^every section can end with",
    r"^there are two ways to add a link",
    r"^\d+\.\s+(inside a sentence|as a link card)",
]
INSTRUCTION_RE = re.compile("|".join(INSTRUCTION_PATTERNS), re.I)

ACCENTS = {"primary": "var(--primary)", "teal": "var(--teal)",
           "orange": "var(--orange)", "deep": "var(--deep)", "mint": "var(--mint)"}

KINDS = {"sp": ("SP", "SharePoint"), "ext": ("WEB", "External"), "apx": ("APX", "Apexon")}

# ---- image handling ----
MAX_IMG_WIDTH = 1600      # downscale anything wider, to control page weight
JPEG_QUALITY = 82         # used when re-encoding a downscaled photo
PAGE_WEIGHT_WARN = 3_000_000   # warn if the finished HTML exceeds ~3 MB

# ---- news feeds ----
FEED_TIMEOUT = 12          # seconds per feed
FEED_MAX_ITEMS = 5         # default headlines per feed
FEED_CACHE = "news-cache.json"
USER_AGENT = "Mozilla/5.0 (compatible; KnowledgeHubBuilder/1.0)"


def parse_feed_date(raw):
    """RSS (RFC 822) and Atom (ISO 8601) date strings -> datetime, or None."""
    raw = (raw or "").strip()
    if not raw:
        return None
    try:
        dt = parsedate_to_datetime(raw)
        if dt:
            return dt.replace(tzinfo=None) if dt.tzinfo else dt
    except (TypeError, ValueError):
        pass
    for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%SZ",
                "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(raw[:len(raw)], fmt)
            return dt.replace(tzinfo=None) if dt.tzinfo else dt
        except ValueError:
            continue
    return None


def fetch_feed(url, source_name, max_items, warnings):
    """Fetch one RSS or Atom feed. Returns a list of headline dicts.

    Only the headline, date and link are taken — never article body text.
    """
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=FEED_TIMEOUT) as resp:
            raw = resp.read()
    except Exception as exc:                                  # noqa: BLE001
        warnings.append(f"[News] could not fetch '{source_name}' ({url[:60]}): "
                        f"{type(exc).__name__} — using cached items if available")
        return None

    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        warnings.append(f"[News] '{source_name}' did not return valid XML "
                        f"({exc}) — is that URL really an RSS/Atom feed?")
        return None

    ns = {"atom": "http://www.w3.org/2005/Atom"}
    items = []

    # RSS 2.0
    for node in root.findall(".//item"):
        title = (node.findtext("title") or "").strip()
        link = (node.findtext("link") or "").strip()
        date = parse_feed_date(node.findtext("pubDate"))
        if title and link:
            items.append({"title": title, "url": link,
                          "date": date.isoformat() if date else "",
                          "source": source_name})

    # Atom
    if not items:
        for node in root.findall(".//atom:entry", ns):
            title = (node.findtext("atom:title", default="", namespaces=ns) or "").strip()
            link = ""
            for lnk in node.findall("atom:link", ns):
                if lnk.get("rel") in (None, "alternate"):
                    link = lnk.get("href") or ""
                    break
            date = parse_feed_date(
                node.findtext("atom:updated", default="", namespaces=ns)
                or node.findtext("atom:published", default="", namespaces=ns))
            if title and link:
                items.append({"title": title, "url": link,
                              "date": date.isoformat() if date else "",
                              "source": source_name})

    if not items:
        warnings.append(f"[News] '{source_name}' returned no items — feed may have moved")
        return None

    return items[:max_items]


def gather_news(feeds, cache_path, warnings, offline=False):
    """Fetch every configured feed, merge, sort newest first, and cache.

    Falls back to the cache for any feed that fails, so a network blip or a
    proxy block never blanks the news section.
    """
    cache = {}
    if cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            cache = {}

    merged, refreshed = [], 0
    for feed in feeds:
        name, url, limit = feed["name"], feed["url"], feed["max"]
        items = None
        if not offline:
            items = fetch_feed(url, name, limit, warnings)
        if items:
            cache[url] = items
            refreshed += 1
        else:
            items = cache.get(url, [])
            if items:
                warnings.append(f"[News] showing cached headlines for '{name}'")
        merged.extend(items)

    try:
        cache_path.write_text(json.dumps(cache, indent=1), encoding="utf-8")
    except OSError:
        pass

    merged.sort(key=lambda i: i.get("date") or "", reverse=True)

    # De-duplicate by URL, preserving order
    seen, out = set(), []
    for item in merged:
        if item["url"] in seen:
            continue
        seen.add(item["url"])
        out.append(item)

    return out, refreshed

# Word autocorrect artefacts that break URLs / narration
NORMALISE = {"\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"', "\u00a0": " "}


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def norm(text):
    """Normalise Word autocorrect artefacts. Does NOT strip — callers decide,
    because stripping per-run destroys the spaces between runs."""
    for bad, good in NORMALISE.items():
        text = text.replace(bad, good)
    return text


def norms(text):
    """norm() plus strip — for whole-paragraph and table-cell values."""
    return norm(text).strip()


def slug(text):
    s = re.sub(r"[^a-z0-9]+", "-", norms(text).lower()).strip("-")
    return s or "section"


def esc(text):
    """Escape inline run text — preserves surrounding spaces."""
    return html.escape(norm(text), quote=False)


def escs(text):
    """Escape a whole value (heading, cell) — safe to strip."""
    return html.escape(norms(text), quote=False)


def extract_image(blip, part, warnings, section_title):
    """Turn a w:blip reference into an inline base64 <img>. Returns (html, bytes)."""
    rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
    if not rid:
        return "", 0
    try:
        image_part = part.related_parts[rid]
        data = image_part.blob
    except (KeyError, AttributeError):
        warnings.append(f"[{section_title}] an image could not be read — skipped")
        return "", 0

    ctype = getattr(image_part, "content_type", "") or "image/png"

    # EMF/WMF are Windows metafiles that browsers cannot display
    if ctype in ("image/x-emf", "image/x-wmf") or ctype.endswith(("emf", "wmf")):
        warnings.append(
            f"[{section_title}] image is EMF/WMF (pasted as a Windows metafile) "
            f"and cannot render in a browser — re-paste it as a picture (Paste Special > PNG)")
        return "", 0

    # Downscale oversized images so the page stays a reasonable weight
    if Image is not None:
        try:
            with Image.open(io.BytesIO(data)) as im:
                if im.width > MAX_IMG_WIDTH:
                    ratio = MAX_IMG_WIDTH / im.width
                    im = im.resize((MAX_IMG_WIDTH, int(im.height * ratio)),
                                   Image.LANCZOS)
                    buf = io.BytesIO()
                    if im.mode in ("RGBA", "P", "LA"):
                        im.convert("RGBA").save(buf, format="PNG", optimize=True)
                        ctype = "image/png"
                    else:
                        im.convert("RGB").save(buf, format="JPEG",
                                               quality=JPEG_QUALITY, optimize=True)
                        ctype = "image/jpeg"
                    data = buf.getvalue()
        except Exception:
            pass  # unreadable by Pillow — embed the original bytes as-is

    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{ctype};base64,{b64}", len(data)


def image_html(drawing, part, warnings, section_title):
    """Build <figure><img>...</figure> for a w:drawing, using Word alt text."""
    blip = drawing.find(".//" + qn("a:blip"))
    if blip is None:
        return "", 0

    # Word stores alt text on docPr: @descr (description) or @name
    alt, caption = "", ""
    docpr = drawing.find(".//" + qn("wp:docPr"))
    if docpr is not None:
        alt = (docpr.get("descr") or "").strip()
        title_attr = (docpr.get("title") or "").strip()
        caption = title_attr
        if not alt and title_attr:
            alt = title_attr

    src, nbytes = extract_image(blip, part, warnings, section_title)
    if not src:
        return "", 0

    if not alt:
        warnings.append(
            f"[{section_title}] image has no alt text — add one in Word "
            f"(right-click the image > Edit Alt Text) for accessibility")
        alt = "Image"

    cap = f'<figcaption>{escs(caption)}</figcaption>' if caption else ""
    return (f'<figure class="fig">'
            f'<img src="{src}" alt="{html.escape(alt, quote=True)}" loading="lazy">'
            f'{cap}</figure>'), nbytes


def pick_columns(n, max_cols):
    """Choose a column count that avoids a lonely item on the last row.

    Prefers an exact fit; otherwise the layout with the fullest last row.
    """
    if n <= 1:
        return 1
    best, best_penalty = min(n, max_cols), None
    for c in range(min(n, max_cols), 1, -1):
        rem = n % c
        penalty = 0 if rem == 0 else (c - rem)
        if best_penalty is None or penalty < best_penalty:
            best, best_penalty = c, penalty
    return best


COL_OVERRIDE_RE = re.compile(r"^(.*?)\s*\((\d+)\)\s*$")


def split_col_override(header_cell):
    """'Value (4)' -> ('Value', 4). Returns (text, None) when no override."""
    m = COL_OVERRIDE_RE.match(header_cell.strip())
    if m:
        count = int(m.group(2))
        if 1 <= count <= 6:
            return m.group(1).strip(), count
    return header_cell.strip(), None


# Placeholders authors use to mean "nobody" in the Reports To column
NO_MANAGER = {"", "-", "--", "\u2013", "\u2014", "n/a", "na", "none", "nil",
              ".", "x", "top", "root", "self"}


def build_org_tree(rows, section_title, warnings):
    """Turn Name | Title | Reports To rows into a forest of nodes.

    A person whose manager is blank, or names someone not listed in the table,
    becomes a root — so 'CEO' can be referenced without being a row.
    """
    people, order = {}, []
    for r in rows:
        if not r or not r[0].strip():
            continue
        name = r[0].strip()
        title = r[1].strip() if len(r) > 1 else ""
        boss = r[2].strip() if len(r) > 2 else ""
        if boss.lower() in NO_MANAGER:
            boss = ""
        if name in people:
            warnings.append(
                f"[{section_title}] '{name}' appears more than once in the org "
                f"table — only the first entry is used")
            continue
        people[name] = {"name": name, "title": title, "boss": boss, "kids": []}
        order.append(name)

    if not people:
        return []

    roots, implied = [], {}
    for name in order:
        node = people[name]
        boss = node["boss"]
        if not boss:
            roots.append(node)
        elif boss in people:
            people[boss]["kids"].append(node)
        else:
            # Manager named but not listed — create a placeholder root for them
            if boss not in implied:
                implied[boss] = {"name": boss, "title": "", "boss": "",
                                 "kids": [], "implied": True}
                roots.append(implied[boss])
            implied[boss]["kids"].append(node)

    # Cycle guard: anything unreachable from a root is in a loop
    seen = set()

    def walk(node, path):
        if node["name"] in path:
            warnings.append(
                f"[{section_title}] reporting loop involving '{node['name']}' "
                f"— check the Reports To column")
            node["kids"] = []
            return
        seen.add(node["name"])
        for kid in list(node["kids"]):
            walk(kid, path | {node["name"]})

    for root in list(roots):
        walk(root, set())

    stranded = [n for n in order if n not in seen]
    if stranded:
        warnings.append(
            f"[{section_title}] not shown because of a reporting loop: "
            + ", ".join(stranded[:5]))

    if len(roots) > 1:
        real = [r for r in roots if not r.get("implied")]
        if len(real) > 1:
            warnings.append(
                f"[{section_title}] org table has {len(real)} people with no "
                f"manager — they render as separate top-level boxes")

    return roots


def render_org_nodes(nodes, depth=0):
    """Nested <ul> tree. Pure CSS connectors — no scripts, no images."""
    if not nodes:
        return ""
    out = ['<ul class="org-level">']
    for n in nodes:
        cls = "org-box" + (" org-implied" if n.get("implied") else "")
        title = f'<span class="org-title">{escs(n["title"])}</span>' if n["title"] else ""
        out.append(
            f'<li><div class="{cls}" data-depth="{min(depth, 4)}">'
            f'<span class="org-name">{escs(n["name"])}</span>{title}</div>'
            + render_org_nodes(n["kids"], depth + 1) + "</li>")
    out.append("</ul>")
    return "".join(out)


def para_html(paragraph, warnings=None, section_title="", img_stats=None):
    """Inline HTML for a paragraph, preserving Ctrl+K hyperlinks, bold runs,
    soft line breaks (Shift+Enter) and inline images."""
    rels = paragraph.part.rels
    warnings = warnings if warnings is not None else []
    out = []
    figures = []

    def run_html(node):
        """Text of a w:r, with <br> for soft breaks, wrapped in <strong> if bold."""
        pieces = []
        for n in node.iter():
            if n.tag == qn("w:t"):
                pieces.append(esc(n.text or ""))
            elif n.tag == qn("w:br"):
                pieces.append("<br>")
            elif n.tag == qn("w:tab"):
                pieces.append(" ")
            elif n.tag == qn("w:drawing"):
                fig, nb = image_html(n, paragraph.part, warnings, section_title)
                if fig:
                    figures.append(fig)
                    if img_stats is not None:
                        img_stats["count"] += 1
                        img_stats["bytes"] += nb
        text = "".join(pieces)
        if not text.strip() and "<br>" not in text:
            return text
        rpr = node.find(qn("w:rPr"))
        if rpr is not None and rpr.find(qn("w:b")) is not None:
            # Keep spaces outside the tag so bold boundaries don't eat them
            lead = " " if text[:1] == " " else ""
            trail = " " if text[-1:] == " " else ""
            return f"{lead}<strong>{text.strip()}</strong>{trail}"
        return text

    def link_html(node):
        rid = node.get(qn("r:id"))
        url = rels[rid].target_ref if rid and rid in rels else ""
        inner = "".join(run_html(r) for r in node.findall(qn("w:r")))
        if not inner.strip():
            return ""
        if url:
            return (f'<a href="{html.escape(url, quote=True)}" '
                    f'target="_blank" rel="noopener">{inner.strip()}</a>')
        return inner

    # Split field hyperlinks: fldChar(begin) -> instrText HYPERLINK "url"
    # -> fldChar(separate) -> display runs -> fldChar(end).
    field_url = None      # url once we're inside the display part of a field
    field_instr = []      # accumulating instrText while between begin and separate
    in_field = False

    for child in paragraph._p.iterchildren():
        if child.tag == qn("w:hyperlink"):
            out.append(link_html(child))
            continue

        if child.tag == qn("w:fldSimple"):
            instr = child.get(qn("w:instr")) or ""
            m = re.search(r'HYPERLINK\s+"([^"]+)"', instr)
            inner = "".join(run_html(r) for r in child.findall(qn("w:r")))
            if m and inner.strip():
                out.append(f'<a href="{html.escape(m.group(1), quote=True)}" '
                           f'target="_blank" rel="noopener">{inner.strip()}</a>')
            else:
                out.append(inner)
            continue

        if child.tag != qn("w:r"):
            continue

        fld = child.find(qn("w:fldChar"))
        if fld is not None:
            state = fld.get(qn("w:fldCharType"))
            if state == "begin":
                in_field, field_instr, field_url = True, [], None
            elif state == "separate":
                instr = "".join(field_instr)
                m = re.search(r'HYPERLINK\s+"([^"]+)"', instr)
                field_url = m.group(1) if m else None
                in_field = False
            elif state == "end":
                field_url, in_field = None, False
            continue

        instr_node = child.find(qn("w:instrText"))
        if instr_node is not None:
            if in_field:
                field_instr.append(instr_node.text or "")
            continue

        piece = run_html(child)
        if field_url and piece.strip():
            lead = " " if piece[:1] == " " else ""
            out.append(f'{lead}<a href="{html.escape(field_url, quote=True)}" '
                       f'target="_blank" rel="noopener">{piece.strip()}</a>')
        else:
            out.append(piece)

    result = "".join(out)

    # Safety net: if markup-walking produced nothing but the paragraph has text,
    # fall back to plain text rather than silently dropping content.
    if not result.strip() and paragraph.text.strip():
        result = esc(paragraph.text)

    # Auto-link bare URLs Word did not linkify
    result = re.sub(
        r'(?<!["\'>])(https?://[^\s<>"\']+)',
        r'<a href="\1" target="_blank" rel="noopener">\1</a>',
        result,
    )
    # Collapse runs of whitespace, tidy space before punctuation
    result = re.sub(r"[ \t]{2,}", " ", result)
    result = re.sub(r"\s+([.,;:!?])", r"\1", result)
    result = result.strip()
    if figures:
        result = (result + "".join(figures)) if result else "".join(figures)
    return result


def is_instruction(text):
    return bool(INSTRUCTION_RE.match(norms(text)))


def table_rows(table, warnings=None, section_title="", img_stats=None):
    """Header list plus data rows, de-duplicating Word's merged-cell repeats.
    Images inside a cell are appended to that cell's text as <figure> markup."""
    def cells(row):
        seen, out = set(), []
        for c in row.cells:
            if c._tc in seen:
                continue
            seen.add(c._tc)
            value = norms(c.text)
            if warnings is not None:
                figs = []
                for d in c._tc.iter(qn("w:drawing")):
                    fig, nb = image_html(d, table.part, warnings, section_title)
                    if fig:
                        figs.append(fig)
                        if img_stats is not None:
                            img_stats["count"] += 1
                            img_stats["bytes"] += nb
                if figs:
                    value = value + "".join(figs)
            out.append(value)
        return out

    rows = [cells(r) for r in table.rows]
    return rows[0], rows[1:]


# ─────────────────────────────────────────────────────────────
# Parser
# ─────────────────────────────────────────────────────────────

class Parser:
    def __init__(self, path):
        self.doc = Document(path)
        self.hero = {}
        self.sections = []
        self.warnings = []
        self.footer = ""
        self.img_stats = {"count": 0, "bytes": 0}
        self.feeds = []
        self.news = []
        self.news_refreshed = 0

    def blocks(self):
        """Yield ('p', Paragraph) / ('t', Table) in true document order."""
        body = self.doc.element.body
        for child in body.iterchildren():
            if child.tag == qn("w:p"):
                yield "p", Paragraph(child, self.doc)
            elif child.tag == qn("w:tbl"):
                yield "t", Table(child, self.doc)

    def _ph(self, obj, section):
        return para_html(obj, self.warnings,
                         section["title"] if section else "before first section",
                         self.img_stats)

    def parse(self):
        started = False
        group = "More"
        section = None

        for kind, obj in self.blocks():

            # ---- paragraphs ----
            if kind == "p":
                text = norms(obj.text)
                has_drawing = obj._p.find(".//" + qn("w:drawing")) is not None
                if not text and not has_drawing:
                    continue

                if not started:
                    if CONTENT_MARKER.lower() in text.lower():
                        started = True
                    continue

                style = (obj.style.name or "").strip()
                # Word appends digits to CUSTOMISED styles ('Subtle Emphasis1').
                # Built-in headings legitimately end in a digit, so exclude them.
                if re.match(r"^Heading \d+$", style):
                    base = style
                else:
                    base = re.sub(r"\d+$", "", style).strip()

                if base == "Heading 1":
                    section = {"title": text, "id": slug(text), "group": group,
                               "lead": [], "blocks": [], "audio": None,
                               "is_hero": text.lower() == "page header"}
                    self.sections.append(section)
                    continue

                if base == "Heading 2":
                    m = re.match(r"group\s*:\s*(.+)", text, re.I)
                    if m:
                        group = m.group(1).strip()
                    else:
                        self.warnings.append(
                            f"Heading 2 '{text}' ignored — expected the form 'Group: <name>'")
                    continue

                if text and is_instruction(text) and not has_drawing:
                    continue
                if section is None:
                    continue

                if base in ("Heading 3", "Heading 4"):
                    section["blocks"].append(
                        {"type": "subhead", "html": para_html(obj, self.warnings,
                                                              section["title"])})
                elif base == "Intense Quote":
                    section["blocks"].append({"type": "warn", "html": self._ph(obj, section)})
                elif base == "Subtle Emphasis":
                    section["blocks"].append({"type": "note", "html": self._ph(obj, section)})
                elif base == "List Paragraph" or obj._p.find(qn("w:numPr")) is not None:
                    body_html = self._ph(obj, section)
                    m = re.match(r"^(?:<strong>)?(.{2,80}?)(?:</strong>)?\s*:\s*(.+)$",
                                 body_html, re.S)
                    if m:
                        step = {"h": m.group(1).strip(), "p": m.group(2).strip()}
                    else:
                        step = {"h": "", "p": body_html}
                    if section["blocks"] and section["blocks"][-1]["type"] == "steps":
                        section["blocks"][-1]["items"].append(step)
                    else:
                        section["blocks"].append({"type": "steps", "items": [step]})
                else:
                    # Lead text before any table; body paragraph after one
                    if not any(b["type"] != "note" for b in section["blocks"]):
                        section["lead"].append(self._ph(obj, section))
                    else:
                        section["blocks"].append({"type": "para", "html": self._ph(obj, section)})
                continue

            # ---- tables ----
            if section is None or not started:
                continue

            header, rows = table_rows(obj, self.warnings, section["title"], self.img_stats)
            if header:
                header = list(header)
                header[0], col_override = split_col_override(header[0])
            else:
                col_override = None
            sig = [h.lower() for h in header]

            # Recovery: authors sometimes paste the Narration rows onto the BOTTOM
            # of the preceding table instead of making a separate table. Detect the
            # 'Narration' marker row and split it off rather than failing silently.
            split_at = None
            for idx, r in enumerate(rows):
                if r and r[0].strip().lower() == "narration":
                    split_at = idx
                    break
            if split_at is not None and sig != ["narration"]:
                tail = rows[split_at + 1:]
                rows = rows[:split_at]
                narration = " ".join(norms(r[0]) for r in tail if r and r[0])
                if narration:
                    section["audio"] = narration
                self.warnings.append(
                    f"[{section['title']}] Narration was merged into the "
                    f"{'|'.join(header)} table — recovered, but it belongs in its own "
                    f"one-column table headed 'Narration'")

            if sig == ["narration"]:
                section["audio"] = " ".join(norms(r[0]) for r in rows if r and r[0])

            elif sig == ["field", "value"]:
                for r in rows:
                    if len(r) < 2:
                        continue
                    key = r[0].strip().lower()
                    val = r[1]
                    m = re.search(r'<img src="([^"]+)"[^>]*alt="([^"]*)"', val)
                    if m:
                        # An image was pasted into this cell
                        self.hero[key + "__img"] = (m.group(1), m.group(2))
                        val = re.sub(r"<figure.*?</figure>", "", val, flags=re.S).strip()
                    if val or key + "__img" not in self.hero:
                        self.hero[key] = val

            elif sig == ["value", "label"]:
                stat_items = [{"n": r[0], "l": r[1]} for r in rows if len(r) >= 2 and r[0]]
                section["blocks"].append({
                    "type": "stats",
                    "items": stat_items,
                    "cols": col_override or pick_columns(len(stat_items), 4),
                })

            elif sig[:2] == ["title", "description"] and "url" in sig:
                iu, it = sig.index("url"), (sig.index("type") if "type" in sig else -1)
                items = []
                for r in rows:
                    if len(r) <= iu or not r[0]:
                        continue
                    url = r[iu].strip()
                    kind_ = (r[it].strip().lower() if it > -1 and len(r) > it else "ext") or "ext"
                    if kind_ not in KINDS:
                        self.warnings.append(
                            f"[{section['title']}] unknown Type '{kind_}' for '{r[0]}' — using 'ext'")
                        kind_ = "ext"
                    if not url.lower().startswith("http"):
                        self.warnings.append(
                            f"[{section['title']}] '{r[0]}' has no valid URL "
                            f"({url[:40] or 'empty'}) — link card skipped")
                        continue
                    items.append({"title": r[0], "desc": r[1], "url": url, "kind": kind_})
                if items:
                    section["blocks"].append({
                        "type": "docs",
                        "items": items,
                        "cols": col_override or pick_columns(len(items), 2),
                    })

            elif sig[:2] == ["title", "description"]:
                ia = sig.index("accent") if "accent" in sig else -1
                items = []
                for r in rows:
                    if not r or not r[0]:
                        continue
                    accent = r[ia].strip().lower() if ia > -1 and len(r) > ia else ""
                    if accent and accent not in ACCENTS:
                        self.warnings.append(
                            f"[{section['title']}] unknown Accent '{accent}' on '{r[0]}' — ignored")
                        accent = ""
                    items.append({"h": r[0],
                                  "p": r[1] if len(r) > 1 else "",
                                  "accent": accent})
                if items:
                    section["blocks"].append({
                        "type": "cards",
                        "items": items,
                        "cols": col_override or pick_columns(len(items), 3),
                    })

            elif sig[:3] == ["name", "title", "reports to"]:
                roots = build_org_tree(rows, section["title"], self.warnings)
                if roots:
                    section["blocks"].append({"type": "org", "roots": roots})
                else:
                    self.warnings.append(
                        f"[{section['title']}] org table has no usable rows — skipped")

            elif sig[:2] == ["feed", "url"]:
                imax = sig.index("max") if "max" in sig else -1
                for r in rows:
                    if len(r) < 2 or not r[0] or not r[1]:
                        continue
                    url = r[1].strip()
                    if not url.lower().startswith("http"):
                        self.warnings.append(
                            f"[{section['title']}] feed '{r[0]}' has no valid URL — skipped")
                        continue
                    try:
                        limit = int(r[imax]) if imax > -1 and len(r) > imax and r[imax] else FEED_MAX_ITEMS
                    except ValueError:
                        limit = FEED_MAX_ITEMS
                    self.feeds.append({"name": r[0], "url": url, "max": limit})
                section["blocks"].append({"type": "news"})

            elif sig == ["term", "definition"]:
                section["blocks"].append({
                    "type": "glossary",
                    "items": [{"t": r[0], "d": r[1]} for r in rows if len(r) >= 2 and r[0]]
                })

            else:
                # Not a recognised block type — render it as a plain table rather
                # than dropping the content. Covers matrices, SLAs, reference grids.
                data_rows = [r for r in rows if any(c.strip() for c in r)]
                if header and data_rows:
                    section["blocks"].append({
                        "type": "table",
                        "header": header,
                        "rows": data_rows,
                    })
                    self.warnings.append(
                        f"[{section['title']}] table {header} rendered as a plain table "
                        f"— if you meant cards or stats, check the header row spelling")
                else:
                    self.warnings.append(
                        f"[{section['title']}] table with headers {header} is empty — skipped")

        # Hero section is config, not a rendered section
        self.sections = [s for s in self.sections if not s["is_hero"]]
        self.footer = self.hero.get("footer", "")

        # Validation
        for s in self.sections:
            if not s["blocks"] and not s["lead"]:
                self.warnings.append(f"[{s['title']}] section is empty")
            if not s["audio"]:
                self.warnings.append(f"[{s['title']}] no Narration table — no Listen button")

        ids = [s["id"] for s in self.sections]
        for i in set(ids):
            if ids.count(i) > 1:
                self.warnings.append(f"duplicate section id '{i}' — nav links will collide")

        return self


# ─────────────────────────────────────────────────────────────
# Renderer
# ─────────────────────────────────────────────────────────────

def wrap_lead(fragment):
    """A <figure> is not valid inside <p>. Split text and figures into siblings."""
    if "<figure" not in fragment:
        return f'<p class="lead">{fragment}</p>' if fragment.strip() else ""
    parts = re.split(r"(<figure.*?</figure>)", fragment, flags=re.S)
    out = []
    for part in parts:
        if not part.strip():
            continue
        if part.startswith("<figure"):
            out.append(part)
        else:
            out.append(f'<p class="lead">{part.strip()}</p>')
    return "".join(out)


def render_block(b):
    t = b["type"]

    if t == "stats":
        return f'<div class="stats" style="--cols:{b.get("cols", 4)}">' + "".join(
            f'<div class="stat"><div class="n">{escs(i["n"])}</div>'
            f'<div class="l">{escs(i["l"])}</div></div>' for i in b["items"]
        ) + "</div>"

    if t == "cards":
        out = []
        for c in b["items"]:
            style = f' style="border-top:3px solid {ACCENTS[c["accent"]]}"' if c["accent"] else ""
            body = c["p"]
            if "<figure" in body:
                txt, _, figs = body.partition("<figure")
                body = f'<p>{escs(txt)}</p>' + ("<figure" + figs if figs else "")
            else:
                body = f'<p>{escs(body)}</p>'
            out.append(f'<div class="card"{style}><h3>{escs(c["h"])}</h3>{body}</div>')
        return (f'<div class="grid" style="--cols:{b.get("cols", 3)}">'
                + "".join(out) + "</div>")

    if t == "steps":
        items = b["items"]
        cards = '<div class="steps">' + "".join(
            f'<div class="step">'
            + (f'<h4>{i["h"]}</h4>' if i["h"] else "")
            + f'<p>{i["p"]}</p></div>' for i in items
        ) + "</div>"

        # Journey strip: a scannable overview above the detailed cards.
        # Only worth drawing when every step has a short title and the count
        # fits across a row — otherwise it is noise.
        titles = [re.sub(r"<[^>]+>", "", i["h"]).strip() for i in items]
        if (3 <= len(items) <= 7
                and all(titles)
                and max(len(t_) for t_ in titles) <= 44):
            strip = ('<div class="journey-wrap"><div class="journey">' + "".join(
                f'<div class="jstep"><span class="jnum">{n}</span>'
                f'<span class="jtitle">{t_}</span></div>'
                for n, t_ in enumerate(titles, 1)) + "</div></div>")
            return strip + cards
        return cards

    if t == "docs":
        out = []
        for d in b["items"]:
            ico, label = KINDS[d["kind"]]
            out.append(
                f'<a class="doclink" href="{html.escape(d["url"], quote=True)}" '
                f'target="_blank" rel="noopener">'
                f'<span class="ico {d["kind"]}">{ico}</span><span>'
                f'<span class="t">{escs(d["title"])}</span>'
                + (f'<span class="d">{escs(d["desc"])}</span>' if d["desc"] else "")
                + f'<br><span class="badge {d["kind"]}">{label}</span></span></a>')
        return (f'<div class="docs" style="--cols:{b.get("cols", 2)}">'
                + "".join(out) + "</div>")

    if t == "table":
        ncol = len(b["header"])
        head = "".join(f"<th>{escs(c)}</th>" for c in b["header"])
        body = []
        for r in b["rows"]:
            cells = []
            for i in range(ncol):
                v = r[i] if i < len(r) else ""
                # keep inline links/figures that came from the cell
                cells.append(f"<td>{v if ('<a ' in v or '<figure' in v) else escs(v)}</td>")
            body.append("<tr>" + "".join(cells) + "</tr>")
        return ('<div class="tbl-wrap"><table class="tbl">'
                f"<thead><tr>{head}</tr></thead>"
                f"<tbody>{''.join(body)}</tbody></table></div>")

    if t == "glossary":
        return '<dl class="gloss">' + "".join(
            f'<dt>{escs(i["t"])}</dt><dd>{escs(i["d"])}</dd>' for i in b["items"]
        ) + "</dl>"

    if t == "news":
        items = b.get("items", [])
        if not items:
            return ('<p class="note">No headlines were available when this page '
                    'was generated.</p>')
        rows = []
        for i in items:
            when = ""
            if i.get("date"):
                try:
                    when = datetime.fromisoformat(i["date"]).strftime("%d %b %Y")
                except ValueError:
                    when = ""
            rows.append(
                f'<a class="news-item" href="{html.escape(i["url"], quote=True)}" '
                f'target="_blank" rel="noopener">'
                f'<span class="news-meta"><span class="news-src">{escs(i["source"])}</span>'
                + (f'<span class="news-date">{when}</span>' if when else "")
                + f'</span><span class="news-title">{escs(i["title"])}</span></a>')
        stamp = datetime.now().strftime("%d %b %Y")
        return ('<div class="news">' + "".join(rows) + "</div>"
                f'<p class="note">Headlines retrieved {stamp}. '
                f'Re-run the converter to refresh.</p>')

    if t == "org":
        return ('<div class="org-wrap"><div class="org">'
                + render_org_nodes(b["roots"]) + "</div></div>")

    if t == "subhead":
        return f'<h3 class="subhead">{b["html"]}</h3>'

    if t == "warn":
        return f'<div class="warn">{b["html"]}</div>'
    if t == "note":
        return f'<p class="note">{b["html"]}</p>'
    if t == "para":
        return f'<p class="body">{b["html"]}</p>'
    return ""


def render(parser, template_css):
    for s in parser.sections:
        for b in s["blocks"]:
            if b["type"] == "news":
                b["items"] = parser.news
    hero = parser.hero
    pills = [v for k, v in hero.items() if k.startswith("pill") and v and not k.endswith("__img")]

    # ---- Groups, in document order ----
    groups = []                      # [{"name","slug","sections":[...]}]
    for s in parser.sections:
        if not groups or groups[-1]["name"] != s["group"]:
            groups.append({"name": s["group"], "slug": "grp-" + slug(s["group"]), "sections": []})
        groups[-1]["sections"].append(s)
    for g in groups:
        # Tabs jump to the group band when there is one, else straight to the section
        g["target"] = g["slug"] if len(g["sections"]) > 1 else g["sections"][0]["id"]

    # ---- Sections, with a band introducing each multi-topic group ----
    sec_html = []
    for gi, g in enumerate(groups, 1):
        if len(g["sections"]) > 1:
            chips = "".join(f'<a href="#{s["id"]}">{escs(s["title"])}</a>' for s in g["sections"])
            sec_html.append(
                f'<div class="group-band" id="{g["slug"]}" data-grp="{g["slug"]}">'
                f'<span class="eyebrow">Area {gi} of {len(groups)} &middot; '
                f'{len(g["sections"])} topics</span>'
                f'<h2 class="group-title">{escs(g["name"])}</h2>'
                f'<nav class="topic-chips" aria-label="{escs(g["name"])} topics">{chips}</nav></div>')
        for s in g["sections"]:
            audio_attr = (' data-audio="' + html.escape(s["audio"], quote=True) + '"') if s["audio"] else ""
            lead = "".join(wrap_lead(l) for l in s["lead"])
            body = "".join(render_block(b) for b in s["blocks"])
            sec_html.append(
                f'<section id="{s["id"]}" data-grp="{g["slug"]}"{audio_attr}>'
                f'<span class="eyebrow">{escs(g["name"])}</span>'
                f'<h2>{escs(s["title"])}</h2>{lead}<div class="stack">{body}</div></section>')

    # ---- Header tabs ----
    tabs = ['<a href="#overview" data-grp="home">Home</a>'] + [
        f'<a href="#{g["target"]}" data-grp="{g["slug"]}">{escs(g["name"])}</a>' for g in groups]

    # ---- Hero stats: "Stat 1" .. "Stat 4" fields ("20 | guided topics"), else computed ----
    first_steps = 0
    for s in parser.sections:
        for b in s["blocks"]:
            if b["type"] == "steps":
                first_steps = len(b["items"]); break
        if first_steps: break
    defaults = [
        (str(len(parser.sections)), "topics"),
        (str(len(groups)), "content areas"),
        (str(first_steps), "onboarding steps") if first_steps else
        (str(sum(1 for s in parser.sections if s["audio"])), "narrated topics"),
        (str(sum(1 for s in parser.sections if s["audio"])), "narrated topics") if first_steps else
        (str(sum(len(s["blocks"]) for s in parser.sections)), "content blocks"),
    ]
    stats = []
    for i in range(4):
        raw = hero.get(f"stat {i+1}", "").strip()
        if "|" in raw:
            v, l = [x.strip() for x in raw.split("|", 1)]
            stats.append((v, l))
        else:
            stats.append(defaults[i])
    stats_html = "".join(
        f'<div class="hero-stat"><strong>{escs(v)}</strong><span>{escs(l)}</span></div>'
        for v, l in stats)

    # ---- Browse-by-area cards ----
    area_cards = []
    wide_first = len(groups) % 2 == 1
    for i, g in enumerate(groups):
        titles = [s["title"] for s in g["sections"]]
        desc = " &middot; ".join(escs(t) for t in titles[:3])
        if len(titles) > 3:
            desc += f" &middot; and {len(titles) - 3} more"
        n = len(titles)
        cls = "area-card wide" if (wide_first and i == 0) else "area-card"
        area_cards.append(
            f'<a class="{cls}" href="#{g["target"]}">'
            f'<span class="area-num">{i+1:02d}</span>'
            f'<span class="area-body"><span class="area-count">{n} topic{"s" if n != 1 else ""}</span>'
            f'<span class="area-name">{escs(g["name"])}</span>'
            f'<span class="area-desc">{desc}</span></span>'
            f'<span class="area-arrow" aria-hidden="true">&rarr;</span></a>')

    # ---- Brand: logo image (left) and brand mark (right) ----
    logo_src, logo_alt = hero.get("logo__img", (None, ""))
    if logo_src:
        brand_icon = (f'<img class="brand-logo" src="{logo_src}" '
                      f'alt="{html.escape(logo_alt or "Logo", quote=True)}">')
    else:
        brand_icon = ('<span class="brand-icon" aria-hidden="true"><svg viewBox="0 0 24 24">'
                      '<path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v15H6.5A2.5 2.5 0 0 0 4 20.5z"/>'
                      '<path d="M4 20.5A2.5 2.5 0 0 1 6.5 18H20v3H6.5A2.5 2.5 0 0 1 4 20.5z"/>'
                      '<path d="M9 7.5h7M9 11h5"/></svg></span>')

    brand_src, brand_alt = hero.get("brand__img", (None, ""))
    brand_text = hero.get("brand", "").strip()
    if brand_src:
        brand_html = (f'<span class="brand-chip img"><img src="{brand_src}" '
                      f'alt="{html.escape(brand_alt or brand_text or "Logo", quote=True)}"></span>')
    elif brand_text:
        brand_html = f'<span class="brand-chip">{escs(brand_text)}</span>'
    else:
        brand_html = ""

    first = parser.sections[0] if parser.sections else None
    cta = (f'<a class="btn btn-primary" href="#{first["id"]}">Start with '
           f'{escs(first["title"])} <span aria-hidden="true">&rarr;</span></a>') if first else ""

    title = escs(hero.get("title", "Knowledge Hub"))
    return TEMPLATE.format(
        css=template_css,
        title=title,
        brand=title,
        brandsub=escs(hero.get("brand subtitle", "Apexon Engagement")),
        brand_icon=brand_icon,
        brandmark=brand_html,
        tabs="".join(tabs),
        eyebrow=escs(hero.get("eyebrow", hero.get("brand subtitle", "Apexon Engagement"))),
        tagline=escs(hero.get("tagline", "")),
        pills="".join(f'<span class="pill">{escs(p)}</span>' for p in pills),
        cta=cta,
        stats=stats_html,
        areas="".join(area_cards),
        sections="".join(sec_html),
        footer=escs(parser.footer),
    )


# ─────────────────────────────────────────────────────────────
# HTML shell (CSS + progressive-enhancement JS carried verbatim)
# ─────────────────────────────────────────────────────────────

TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<style>{css}</style>
</head>
<body id="top">
<noscript>
  <div class="noscript">
    <strong>Scripts are blocked.</strong> All content is readable below &mdash; only audio narration and
    navigation highlighting are unavailable. To enable them, click <em>Allow blocked content</em>.
  </div>
</noscript>
<a class="skip-link" href="#main">Skip to content</a>

<header class="site-header">
  <div class="header-main">
    <a class="site-brand" href="#overview">{brand_icon}
      <span class="brand-text"><strong>{brand}</strong><small>{brandsub}</small></span>
    </a>
    {brandmark}
  </div>
  <nav class="tabs" aria-label="Areas">{tabs}</nav>
</header>

<div class="page">
  <section class="hero" id="overview" data-grp="home" aria-label="Overview">
    <div class="hero-copy">
      <span class="eyebrow light">{eyebrow}</span>
      <h1>{title}</h1>
      <p class="tag">{tagline}</p>
      <div class="pills">{pills}</div>
      <div class="hero-actions">{cta}<a class="btn btn-ghost" href="#areas">Browse by area</a></div>
      <div class="audio-unsupported" id="ttsWarn">Audio narration unavailable in this browser. Requires Edge or Chrome with scripts allowed.</div>
    </div>
    <div class="hero-stats">{stats}</div>
  </section>

  <div class="areas" id="areas" data-grp="home">
    <span class="eyebrow">Explore</span>
    <h2 class="areas-title">Browse by area</h2>
    <div class="area-grid">{areas}</div>
  </div>

  <main id="main">
    {sections}
  </main>

  <footer>{footer}</footer>
</div>

<div class="audiobar" id="audiobar">
  <span class="now">Narrating: <b id="nowTitle">&mdash;</b></span>
  <select id="voiceSel" title="Voice"></select>
  <select id="rateSel" title="Speed">
    <option value="0.85">0.85&times;</option>
    <option value="1" selected>1&times;</option>
    <option value="1.15">1.15&times;</option>
    <option value="1.3">1.3&times;</option>
    <option value="1.5">1.5&times;</option>
  </select>
  <button id="btnPause">Pause</button>
  <button id="btnStop">Stop</button>
  <span class="voice-hint" id="voiceHint">Open in Microsoft Edge for natural-sounding voices</span>
</div>
<a class="to-top" href="#top" aria-label="Back to top">&uarr;</a>

<script>
(function(){{
  "use strict";
  var tabs = Array.prototype.slice.call(document.querySelectorAll('.tabs a'));
  function setActive(g){{
    tabs.forEach(function(t){{
      var on = t.getAttribute('data-grp') === g;
      t.classList.toggle('active', on);
      if(on && t.scrollIntoView && t.parentNode.scrollWidth > t.parentNode.clientWidth){{
        t.parentNode.scrollLeft = t.offsetLeft - 24;
      }}
    }});
  }}
  setActive('home');
  if('IntersectionObserver' in window){{
    var obs = new IntersectionObserver(function(es){{
      es.forEach(function(e){{ if(e.isIntersecting) setActive(e.target.getAttribute('data-grp')); }});
    }},{{rootMargin:'-35% 0px -60% 0px'}});
    document.querySelectorAll('[data-grp]').forEach(function(el){{
      if(el.tagName !== 'A') obs.observe(el);
    }});
  }}

  var synth = window.speechSynthesis;
  if(!synth || typeof window.SpeechSynthesisUtterance !== 'function'){{
    var w = document.getElementById('ttsWarn'); if(w) w.style.display='block';
    return;
  }}
  var bar=document.getElementById('audiobar'), nowTitle=document.getElementById('nowTitle'),
      voiceSel=document.getElementById('voiceSel'), rateSel=document.getElementById('rateSel'),
      btnPause=document.getElementById('btnPause'), btnStop=document.getElementById('btnStop'),
      hint=document.getElementById('voiceHint');
  var currentBtn=null, queue=[], chunkIdx=0, stopping=false, voices=[];
  var SPEAKER='<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3a4.5 4.5 0 0 0-2.5-4.03v8.05A4.47 4.47 0 0 0 16.5 12zM14 3.23v2.06a7 7 0 0 1 0 13.42v2.06a9 9 0 0 0 0-17.54z"/></svg>';
  var EQ='<span class="eq"><i></i><i></i><i></i></span>';

  /* Rank voices so the most human-sounding one is chosen and listed first.
     Neural / "Natural" / online voices score far above legacy SAPI voices. */
  var NATURAL_RE = /natural|neural|online|multilingual/i;
  var MS_NATURAL_NAMES = /aria|jenny|guy|davis|amber|ana|andrew|ashley|brandon|brian|christopher|cora|elizabeth|emma|eric|jacob|jane|jason|michelle|monica|nancy|roger|ryan|sara|steffan|tony|libby|maisie|noah|oliver|olivia|thomas/i;
  var LEGACY_RE = /david|zira|mark|hazel|susan|sam\\b|desktop|espeak|festival|pico/i;

  function voiceScore(v){{
    var s = 0, n = v.name || '';
    if(NATURAL_RE.test(n)) s += 60;
    if(v.localService === false) s += 25;      // cloud voices are usually neural
    if(/google/i.test(n)) s += 22;
    if(MS_NATURAL_NAMES.test(n)) s += 14;
    if(LEGACY_RE.test(n) && !NATURAL_RE.test(n)) s -= 45;
    if(/en[-_]US/i.test(v.lang)) s += 10;
    else if(/en[-_]GB/i.test(v.lang)) s += 6;
    else if(/^en/i.test(v.lang)) s += 3;
    return s;
  }}

  function prettyVoice(v){{
    var n = (v.name||'').replace(/^Microsoft\\s+|^Google\\s+/i,'')
                        .replace(/\\s*-\\s*English.*$/i,'')
                        .replace(/Online \\(Natural\\)/i,'')
                        .replace(/\\(Natural\\)/i,'').trim();
    return n + (NATURAL_RE.test(v.name||'') ? ' \u2605' : '') + ' \u00b7 ' + v.lang;
  }}

  function loadVoices(){{
    var all = synth.getVoices();
    var en = all.filter(function(v){{ return /^en/i.test(v.lang); }});
    voices = (en.length ? en : all).slice().sort(function(a,b){{
      return voiceScore(b) - voiceScore(a);
    }});
    voiceSel.innerHTML='';
    voices.forEach(function(v,i){{
      var o=document.createElement('option'); o.value=i;
      o.textContent=prettyVoice(v);
      voiceSel.appendChild(o);
    }});
    voiceSel.selectedIndex = 0;   // best-ranked voice
    if(voices.length && voiceScore(voices[0]) < 20 && hint){{
      hint.style.display='block';
    }} else if(hint){{
      hint.style.display='none';
    }}
  }}
  loadVoices();
  if(typeof synth.onvoiceschanged!=='undefined') synth.onvoiceschanged=loadVoices;

  function chunkText(text){{
    var parts=text.match(/[^.!?]+[.!?]+|\\S[^.!?]*$/g)||[text];
    var out=[],buf='';
    parts.forEach(function(p){{
      p=p.trim();
      if((buf+' '+p).length>200){{ if(buf) out.push(buf); buf=p; }}
      else buf = buf?buf+' '+p:p;
    }});
    if(buf) out.push(buf);
    return out;
  }}
  function speakNext(){{
    if(stopping||chunkIdx>=queue.length){{ finish(); return; }}
    var u=new SpeechSynthesisUtterance(queue[chunkIdx]);
    if(voices[voiceSel.value]) u.voice=voices[voiceSel.value];
    u.rate=parseFloat(rateSel.value)||1;
    u.onend=function(){{ chunkIdx++; speakNext(); }};
    u.onerror=function(){{ chunkIdx++; speakNext(); }};
    synth.speak(u);
  }}
  function finish(){{
    if(currentBtn){{
      currentBtn.classList.remove('playing');
      currentBtn.querySelector('.lbl').textContent='Listen';
      currentBtn.setAttribute('aria-pressed','false');
    }}
    currentBtn=null; queue=[]; chunkIdx=0; stopping=false;
    bar.classList.remove('on'); btnPause.textContent='Pause';
  }}
  function stopAll(){{ stopping=true; synth.cancel(); finish(); }}
  function play(btn,title,text){{
    if(currentBtn===btn){{ stopAll(); return; }}
    stopAll(); stopping=false; currentBtn=btn;
    queue=chunkText(text); chunkIdx=0;
    btn.classList.add('playing');
    btn.querySelector('.lbl').textContent='Stop';
    btn.setAttribute('aria-pressed','true');
    nowTitle.textContent=title; bar.classList.add('on');
    btnPause.textContent='Pause'; speakNext();
  }}
  document.querySelectorAll('[data-audio]').forEach(function(sec){{
    var h2=sec.querySelector('h2'); if(!h2) return;
    var title=h2.textContent.trim();
    var btn=document.createElement('button');
    btn.className='audio-btn'; btn.type='button';
    btn.setAttribute('aria-pressed','false');
    btn.setAttribute('aria-label','Listen to the '+title+' section');
    btn.innerHTML=SPEAKER+EQ+'<span class="lbl">Listen</span>';
    btn.addEventListener('click',function(e){{ e.preventDefault(); play(btn,title,sec.getAttribute('data-audio')); }});
    h2.appendChild(btn);
  }});
  btnPause.addEventListener('click',function(){{
    if(synth.paused){{ synth.resume(); btnPause.textContent='Pause'; }}
    else {{ synth.pause(); btnPause.textContent='Resume'; }}
  }});
  btnStop.addEventListener('click',stopAll);
  function restartCurrent(){{
    if(!currentBtn) return;
    var btn=currentBtn,title=nowTitle.textContent,sec=btn.closest('[data-audio]');
    stopAll(); setTimeout(function(){{ play(btn,title,sec.getAttribute('data-audio')); }},60);
  }}
  voiceSel.addEventListener('change',restartCurrent);
  rateSel.addEventListener('change',restartCurrent);
  document.addEventListener('visibilitychange',function(){{ if(document.hidden) stopAll(); }});
  window.addEventListener('beforeunload',function(){{ synth.cancel(); }});
}})();
</script>
</body>
</html>
"""


# ─────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    check_only = "--check" in sys.argv
    offline = "--offline" in sys.argv

    if not args:
        sys.exit(__doc__)

    src = Path(args[0])
    dst = Path(args[1]) if len(args) > 1 else src.with_suffix(".html")

    if not src.exists():
        sys.exit(f"Input not found: {src}")

    css_path = Path(__file__).with_name("hub.css")
    if not css_path.exists():
        sys.exit("hub.css must sit next to docx2hub.py")
    css = css_path.read_text(encoding="utf-8")

    p = Parser(src).parse()

    if p.feeds:
        cache_path = dst.parent / FEED_CACHE
        p.news, p.news_refreshed = gather_news(
            p.feeds, cache_path, p.warnings, offline=offline)

    print(f"Parsed {len(p.sections)} sections from {src.name}")
    for s in p.sections:
        kinds = ", ".join(sorted({b["type"] for b in s["blocks"]})) or "text only"
        audio = "audio" if s["audio"] else "NO AUDIO"
        print(f"   • {s['title']:<42} [{s['group']}]  {kinds}  ({audio})")

    if p.feeds:
        print(f"\nNews: {len(p.feeds)} feed(s) configured, "
              f"{p.news_refreshed} refreshed, {len(p.news)} headline(s) shown"
              + (" (offline mode — cache only)" if offline else ""))

    if p.warnings:
        print(f"\n{len(p.warnings)} warning(s):")
        for w in p.warnings:
            print(f"   ! {w}")
    else:
        print("\nNo warnings.")

    if check_only:
        print("\n--check: nothing written.")
        return

    dst.write_text(render(p, css), encoding="utf-8")
    size = dst.stat().st_size
    if p.img_stats["count"]:
        print(f"\nEmbedded {p.img_stats['count']} image(s), "
              f"{p.img_stats['bytes']:,} bytes before base64 encoding")
    print(f"Wrote {dst}  ({size:,} bytes)")
    if size > PAGE_WEIGHT_WARN:
        print(f"   ! Page is over {PAGE_WEIGHT_WARN // 1_000_000} MB — large for an "
              f"email attachment. Consider fewer or smaller images.")


if __name__ == "__main__":
    main()
