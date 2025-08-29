import re
from bs4 import BeautifulSoup, FeatureNotFound
from bs4.element import Tag, NavigableString, Comment, Doctype, ProcessingInstruction

from .fetch import fetch_html
from .settings import ALWAYS_STRIP, INLINE_TAGS
from .utils import (
    normalise_keep_newlines,
    is_noise,
    uk_today_str,
    clean_slug_to_name,
    fallback_page_name_from_url,
)
from .types import ExtractOptions

def annotate_anchor_text(a: Tag, annotate_links: bool) -> str:
    text = a.get_text(" ", strip=True)
    href = a.get("href", "")
    return f"{text} (→ {href})" if (annotate_links and href) else text

def extract_text_preserve_breaks(node: Tag | NavigableString, annotate_links: bool) -> str:
    """Extract visible text; convert <br> to \\n; handle anchors as one unit."""
    if isinstance(node, NavigableString):
        return str(node)
    parts = []
    for child in node.children:
        if isinstance(child, NavigableString):
            parts.append(str(child))
        elif isinstance(child, Tag):
            if child.name == "br":
                parts.append("\n")
            elif child.name == "a":
                parts.append(annotate_anchor_text(child, annotate_links))
            else:
                parts.append(extract_text_preserve_breaks(child, annotate_links))
    return "".join(parts)

def extract_signposted_lines_from_body(body: Tag, annotate_links: bool, include_img_src: bool = False) -> list[str]:
    """
    Emit ONLY:
      - <h1> … <h6> lines
      - <p> lines
      - <img alt="…"> (or <img alt="…" src="…"> when enabled) for every <img> encountered

    Lists are flattened to <p>. Critically, <p> is split on <br> and blank lines preserved
    (blank <p> emitted as '<p>' with no text).

    Additionally, capture stray text nodes (bare text in containers) as <p>, but skip
    comments/doctype/processing instructions and obvious UI/analytics noise.
    """
    lines: list[str] = []

    def emit_lines(tag_name: str, text: str):
        text = normalise_keep_newlines(text)
        segments = text.split("\n")
        for seg in segments:
            seg_stripped = seg.strip()
            if seg_stripped:
                if tag_name == "p" and is_noise(seg_stripped):
                    continue
                lines.append(f"<{tag_name}> {seg_stripped}")
            else:
                if tag_name == "p":
                    lines.append("<p>")

    def emit_img(img_tag: Tag):
        if not isinstance(img_tag, Tag) or img_tag.name != "img":
            return
        alt = (img_tag.get("alt") or "").strip().replace('"', '\\"')
        if include_img_src:
            src = (img_tag.get("src") or "").strip().replace('"', '\\"')
            if src:
                lines.append(f'<img alt="{alt}" src="{src}">')
                return
        lines.append(f'<img alt="{alt}">')

    def handle(tag: Tag):
        name = tag.name
        if name in ALWAYS_STRIP:
            return

        # Headings
        if name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            txt = extract_text_preserve_breaks(tag, annotate_links)
            if txt.strip():
                emit_lines(name, txt)
            return

        # Paragraphs
        if name == "p":
            txt = tag.get_text(" ", strip=True)
            if txt.strip():
                emit_lines("p", txt)
            for img in tag.find_all("img"):
                emit_img(img)
            return

        # Lists
        if name in {"ul", "ol"}:
            for li in tag.find_all("li", recursive=False):
                txt = extract_text_preserve_breaks(li, annotate_links)
                if txt.strip():
                    emit_lines("p", txt)
                for img in li.find_all("img"):
                    emit_img(img)
                for sub in li.find_all(["ul", "ol"], recursive=False):
                    for sub_li in sub.find_all("li", recursive=False):
                        sub_txt = extract_text_preserve_breaks(sub_li, annotate_links)
                        if sub_txt.strip():
                            emit_lines("p", sub_txt)
                        for img in sub_li.find_all("img"):
                            emit_img(img)
            return

        # Generic containers
        buf = []
        def flush_buf():
            if not buf:
                return
            joined = normalise_keep_newlines("".join(buf))
            if joined.strip() and not is_noise(joined):
                emit_lines("p", joined)
            buf.clear()

        for child in tag.children:
            if isinstance(child, (Comment, Doctype, ProcessingInstruction)):
                continue
            if isinstance(child, NavigableString):
                buf.append(str(child))
            elif isinstance(child, Tag):
                if child.name == "br":
                    buf.append("\n")
                elif child.name == "img":
                    flush_buf()
                    emit_img(child)
                elif child.name in INLINE_TAGS:
                    buf.append(extract_text_preserve_breaks(child, annotate_links))
                else:
                    flush_buf()
                    handle(child)
        flush_buf()

    for child in body.children:
        if isinstance(child, (Comment, Doctype, ProcessingInstruction)):
            continue
        if isinstance(child, NavigableString):
            raw = normalise_keep_newlines(str(child))
            if raw.strip() and not is_noise(raw):
                emit_lines("p", raw)
        elif isinstance(child, Tag):
            if child.name == "img":
                emit_img(child)
            else:
                handle(child)

    # Deduplicate trivial adjacent repeats
    deduped, prev = [], None
    for ln in lines:
        if ln != prev:
            deduped.append(ln)
        prev = ln
    return deduped

def first_h1_text(soup: BeautifulSoup) -> str | None:
    if not soup.body:
        return None
    h1 = soup.body.find("h1")
    if not h1:
        return None
    txt = extract_text_preserve_breaks(h1, annotate_links=False)
    txt = normalise_keep_newlines(txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip() or None

def process_url(url: str, opts: ExtractOptions):
    final_url, html_bytes = fetch_html(url)
    # Prefer lxml but gracefully fall back if unavailable
    try:
        soup = BeautifulSoup(html_bytes, "lxml")
    except FeatureNotFound:
        soup = BeautifulSoup(html_bytes, "html.parser")

    # global strip (script/style/noscript/template)
    for el in soup.find_all(list(ALWAYS_STRIP)):
        el.decompose()

    body = soup.body or soup

    # exclude universal blocks
    for sel in opts.exclude_selectors:
        try:
            for el in body.select(sel):
                el.decompose()
        except Exception:
            pass

    # hard-kill: ensure any element with all three classes is removed even if selector order changes
    try:
        for el in body.find_all(lambda t: isinstance(t, Tag) and t.has_attr('class') and {'sr-main','js-searchpage-content','visible'}.issubset(set(t.get('class', [])))):
            el.decompose()
    except Exception:
        pass

    # Also explicitly remove via robust CSS selectors (belt-and-braces)
    for sel in [
        '.sr-main.js-searchpage-content.visible',
        "[class~='sr-main'][class~='js-searchpage-content'][class~='visible']",
        "[class*='js-searchpage-content']",
        "[class*='searchpage-content']",
        ".lmd-map-modal-create.js-lmd-map-modal-map",
    ]:
        try:
            for el in body.select(sel):
                el.decompose()
        except Exception:
            pass

    # If requested, remove everything before the first <h1> but keep the rest of body intact
    if opts.remove_before_h1 and body.name == "body":
        first_h1 = body.find("h1")
        if first_h1 is not None:
            top = first_h1
            while top is not None and top.parent is not None and top.parent != body:
                top = top.parent
            if top is not None and top in body.contents:
                for el in list(body.contents):
                    if el == top:
                        break
                    try:
                        el.decompose()
                    except Exception:
                        continue

    # extract signposted lines
    lines = extract_signposted_lines_from_body(
        body,
        annotate_links=opts.annotate_links,
        include_img_src=opts.include_img_src,
    )

    # meta
    head = soup.head or soup
    title = head.title.string.strip() if (head and head.title and head.title.string) else "N/A"
    meta_el = head.find("meta", attrs={"name": "description"}) if head else None
    description = meta_el.get("content").strip() if (meta_el and meta_el.get("content")) else "N/A"

    # page name: prefer H1
    page_name = first_h1_text(soup) or fallback_page_name_from_url(final_url)

    meta = {
        "page": page_name,
        "date": uk_today_str(),
        "url": final_url,
        "title": title,
        "title_len": len(title) if title != "N/A" else 0,
        "description": description,
        "description_len": len(description) if description != "N/A" else 0,
    }
    return meta, lines
