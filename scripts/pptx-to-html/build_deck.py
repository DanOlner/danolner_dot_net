#!/usr/bin/env python3
"""
Turn a PowerPoint deck into a self-contained HTML slide viewer for the site.

Each slide is exported by Microsoft PowerPoint itself (so fonts and layout
match the original exactly), rasterised to WebP, and wrapped in a viewer
that can be driven by keyboard, mouse, wheel or touch. On top of the images:

  - hyperlinks and plain-text URLs on the slides become clickable areas;
  - embedded videos play in place;
  - each slide's text goes into the page (hidden visually) for screen
    readers and search engines.

Output goes to presentations/<slug>/ in the repo. Quarto copies everything
under presentations/*/ into docs/ on each render (see `resources` in
_quarto.yml), so a full render never wipes it.

Usage (from the repo root):

  python3 scripts/pptx-to-html/build_deck.py \\
      scripts/pptx-to-html/decks/models-in-the-wild-2013.json  path/to/deck.pptx

  # skip PowerPoint and use a PDF you already exported from it:
  python3 scripts/pptx-to-html/build_deck.py CONFIG.json deck.pptx --pdf deck.pdf

Needs: macOS with Microsoft PowerPoint (unless --pdf), and poppler
(pdftoppm, pdftotext, pdfinfo) plus cwebp, e.g. `brew install poppler webp`.

The config JSON holds the deck's metadata (title, slug, description, ...)
and optional per-slide alt text for slides that are only pictures.
"""

import argparse
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent.parent
TEMPLATE = HERE / "viewer_template.html"

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
R_ID = "{%s}id" % NS["r"]
R_LINK = "{%s}link" % NS["r"]
R_EMBED = "{%s}embed" % NS["r"]
SKIP_PLACEHOLDERS = {"dt", "ftr", "sldNum", "hdr"}

SLIDE_WIDTH_PX = 2048
THUMB_WIDTH_PX = 480
WEBP_QUALITY = "82"
THUMB_QUALITY = "70"


def run(cmd, **kw):
    return subprocess.run(cmd, check=True, **kw)


# --------------------------------------------------------------------------
# 1. PowerPoint -> PDF
# --------------------------------------------------------------------------

def export_pdf_with_powerpoint(pptx: Path, out_pdf: Path):
    """Ask PowerPoint to save the deck as PDF.

    PowerPoint is sandboxed, so the deck is copied into its own container
    first; that avoids file-access prompts.
    """
    box = Path.home() / "Library/Containers/com.microsoft.Powerpoint/Data/tmp"
    work = box / f"pptx2html-{os.getpid()}"
    work.mkdir(parents=True, exist_ok=True)
    try:
        src = work / "deck.pptx"
        pdf = work / "deck.pdf"
        shutil.copyfile(pptx, src)
        script = f'''
        tell application "Microsoft PowerPoint"
            open POSIX file "{src}"
            delay 3
            set p to active presentation
            save p in POSIX file "{pdf}" as save as PDF
            close p saving no
        end tell
        '''
        run(["osascript", "-e", script])
        shutil.move(str(pdf), out_pdf)
    finally:
        shutil.rmtree(work, ignore_errors=True)


# --------------------------------------------------------------------------
# 2. PDF -> WebP slides and thumbnails
# --------------------------------------------------------------------------

def pdf_page_size(pdf: Path):
    out = run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
    m = re.search(r"Page size:\s*([\d.]+)\s*x\s*([\d.]+)", out)
    pages = int(re.search(r"Pages:\s*(\d+)", out).group(1))
    return float(m.group(1)), float(m.group(2)), pages


def rasterise(pdf: Path, out_dir: Path, n_pages: int, tmp: Path):
    png_dir = tmp / "png"
    png_dir.mkdir()
    run(["pdftoppm", "-png", "-scale-to-x", str(SLIDE_WIDTH_PX), "-scale-to-y", "-1",
         str(pdf), str(png_dir / "p")])
    pngs = sorted(png_dir.glob("p-*.png"), key=lambda p: int(p.stem.split("-")[1]))
    assert len(pngs) == n_pages, (len(pngs), n_pages)

    (out_dir / "slides").mkdir(parents=True, exist_ok=True)
    (out_dir / "thumbs").mkdir(parents=True, exist_ok=True)

    def convert(i_png):
        i, png = i_png
        name = f"{i:03d}.webp"
        run(["cwebp", "-quiet", "-q", WEBP_QUALITY, "-m", "6", str(png),
             "-o", str(out_dir / "slides" / name)])
        run(["cwebp", "-quiet", "-q", THUMB_QUALITY, "-resize", str(THUMB_WIDTH_PX), "0",
             str(png), "-o", str(out_dir / "thumbs" / name)])

    with ThreadPoolExecutor(max_workers=os.cpu_count() or 4) as pool:
        list(pool.map(convert, enumerate(pngs, 1)))

    # Social-media preview image from the title slide.
    run(["pdftoppm", "-jpeg", "-jpegopt", "quality=85", "-f", "1", "-l", "1",
         "-scale-to-x", "1200", "-scale-to-y", "-1", "-singlefile",
         str(pdf), str(out_dir / "og")])


# --------------------------------------------------------------------------
# 3. Links: real hyperlinks (PDF annotations) + plain-text URLs (pdftotext)
# --------------------------------------------------------------------------

def _pdf_string(raw: bytes) -> str:
    """Decode a PDF literal string body (without the outer parentheses)."""
    out, i = bytearray(), 0
    esc = {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"\b", b"f": b"\f"}
    while i < len(raw):
        c = raw[i:i + 1]
        if c == b"\\" and i + 1 < len(raw):
            n = raw[i + 1:i + 2]
            if n in esc:
                out += esc[n]; i += 2
            elif n.isdigit():
                m = re.match(rb"[0-7]{1,3}", raw[i + 1:])
                out.append(int(m.group(0), 8) & 0xFF); i += 1 + len(m.group(0))
            else:
                out += n; i += 2
        else:
            out += c; i += 1
    return out.decode("latin-1")


def pdf_link_annotations(pdf: Path):
    """Return {page_index: [(x0, y0, x1, y1, uri), ...]} in PDF points,
    origin bottom-left. Minimal parser: enough for PowerPoint/Quartz PDFs,
    which store pages and annotations as plain objects."""
    b = pdf.read_bytes()
    objs = {int(m.group(1)): m.group(2) for m in
            re.finditer(rb"(?<![0-9])(\d+)\s+0\s+obj\b(.*?)\bendobj", b, re.S)}

    def ref(d, key):
        m = re.search(rb"/" + key + rb"\s+(\d+)\s+0\s+R", d)
        return int(m.group(1)) if m else None

    def ref_list(d, key):
        m = re.search(rb"/" + key + rb"\s*\[(.*?)\]", d, re.S)
        if not m:
            r = ref(d, key)  # indirect array
            if r is None or r not in objs:
                return []
            body = objs[r]
        else:
            body = m.group(1)
        return [int(x) for x in re.findall(rb"(\d+)\s+0\s+R", body)]

    catalog = next((n for n, d in objs.items() if re.search(rb"/Type\s*/Catalog\b", d)), None)
    if catalog is None:
        print("  ! no PDF catalog found; skipping hyperlink areas", file=sys.stderr)
        return {}
    pages = []

    def walk(n, seen=set()):
        if n in seen or n not in objs:
            return
        seen.add(n)
        d = objs[n]
        if re.search(rb"/Type\s*/Pages\b", d):
            for k in ref_list(d, b"Kids"):
                walk(k)
        elif re.search(rb"/Type\s*/Page\b", d):
            pages.append(n)

    walk(ref(objs[catalog], b"Pages"))

    links = {}
    for pi, pn in enumerate(pages):
        for an in ref_list(objs[pn], b"Annots"):
            a = objs.get(an, b"")
            if not re.search(rb"/Subtype\s*/Link\b", a):
                continue
            rect = re.search(rb"/Rect\s*\[\s*([-\d.\s]+)\]", a)
            if not rect:
                continue
            action = a
            r = ref(a, b"A")
            if r is not None:
                action = objs.get(r, b"")
            m = re.search(rb"/URI\s*\((.*?)(?<!\\)\)", action, re.S)
            if m:
                uri = _pdf_string(m.group(1))
            else:
                r = ref(action, b"URI")
                m = re.search(rb"^\s*\((.*)\)\s*$", objs.get(r, b""), re.S) if r else None
                if not m:
                    continue
                uri = _pdf_string(m.group(1))
            if not re.match(r"(https?://|mailto:)", uri, re.I):
                continue
            x0, y0, x1, y1 = [float(v) for v in rect.group(1).split()[:4]]
            links.setdefault(pi, []).append((min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1), uri))
    return links


def pdf_text_urls(pdf: Path, ignore):
    """Plain-text URLs on each page: {page_index: [(x0, top, x1, bottom, url)]}
    in PDF points, origin top-left."""
    out = run(["pdftotext", "-bbox-layout", str(pdf), "-"],
              capture_output=True, text=True).stdout
    result = {}
    for pi, page in enumerate(re.split(r"<page\b", out)[1:]):
        for m in re.finditer(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" '
                             r'yMax="([\d.]+)">(.*?)</word>', page):
            word = html.unescape(m.group(5))
            if not re.match(r"(https?://|www\.)\S{4,}", word, re.I):
                continue
            url = word.rstrip(".,;:)]}\"'’”")
            if any(re.search(p, url, re.I) for p in ignore):
                continue
            if not url.lower().startswith("http"):
                url = "http://" + url
            x0, y0, x1, y1 = (float(m.group(k)) for k in range(1, 5))
            result.setdefault(pi, []).append((x0, y0, x1, y1, url))
    return result


def merge_rects(rects):
    """PowerPoint writes one link rectangle per text run, so a single URL can
    arrive in many pieces. Union same-URL pieces that share a line."""
    merged = []
    for r in sorted(rects, key=lambda r: (r[4], (r[1] + r[3]) / 2, r[0])):
        mid = (r[1] + r[3]) / 2
        for k, m in enumerate(merged):
            if m[4] == r[4] and m[1] - 2 <= mid <= m[3] + 2:
                merged[k] = (min(m[0], r[0]), min(m[1], r[1]), max(m[2], r[2]), max(m[3], r[3]), m[4])
                break
        else:
            merged.append(r)
    return merged


def overlaps(a, b):
    return not (a[2] <= b[0] or b[2] <= a[0] or a[3] <= b[1] or b[3] <= a[1])


def collect_links(pdf: Path, page_w, page_h, n_pages, ignore):
    annots = pdf_link_annotations(pdf)
    words = pdf_text_urls(pdf, ignore)
    per_page = []
    for pi in range(n_pages):
        # Annotation rects -> top-left origin to match pdftotext.
        rects = [(x0, page_h - y1, x1, page_h - y0, u) for x0, y0, x1, y1, u in annots.get(pi, [])
                 if not any(re.search(p, u, re.I) for p in ignore)]
        rects = merge_rects(rects)
        for w in words.get(pi, []):
            if not any(overlaps(w, r) for r in rects):
                rects.append(w)
        pad = 2.0
        per_page.append([{
            "url": u,
            "left": max(0, (x0 - pad) / page_w * 100),
            "top": max(0, (y0 - pad) / page_h * 100),
            "width": min(100, (x1 - x0 + 2 * pad) / page_w * 100),
            "height": min(100, (y1 - y0 + 2 * pad) / page_h * 100),
        } for x0, y0, x1, y1, u in rects])
    return per_page


# --------------------------------------------------------------------------
# 4. Slide text and embedded videos, straight from the pptx XML
# --------------------------------------------------------------------------

def _rels(z, part):
    d, f = os.path.split(part)
    path = f"{d}/_rels/{f}.rels"
    if path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(path))
    out = {}
    for r in root:
        tgt = r.get("Target")
        if r.get("TargetMode") != "External":
            tgt = os.path.normpath(os.path.join(d, tgt)).replace("\\", "/")
        out[r.get("Id")] = tgt
    return out


def _para_text(p):
    bits = []
    for el in p.iter():
        tag = el.tag.split("}")[1]
        if tag == "t" and el.text:
            bits.append(el.text)
        elif tag == "br":
            bits.append(" ")
    return re.sub(r"\s+", " ", "".join(bits)).strip()


def read_pptx(pptx: Path, footer_patterns):
    footer = [re.compile(p, re.I) for p in footer_patterns]
    z = zipfile.ZipFile(pptx)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find("p:sldSz", NS)
    sw, sh = int(sz.get("cx")), int(sz.get("cy"))
    prels = _rels(z, "ppt/presentation.xml")

    slides = []
    for sid in pres.find("p:sldIdLst", NS):
        part = prels[sid.get(R_ID)]
        root = ET.fromstring(z.read(part))
        if root.get("show") == "0":
            continue  # hidden slides are left out of PowerPoint's PDF too
        rels = _rels(z, part)
        title, body, videos = [], [], []

        def walk(tree, xf=lambda x, y: (x, y), sc=(1.0, 1.0)):
            for el in tree:
                tag = el.tag.split("}")[1]
                if tag == "grpSp":
                    g = el.find("p:grpSpPr/a:xfrm", NS)
                    if g is not None and g.find("a:chExt", NS) is not None:
                        off, ext = g.find("a:off", NS), g.find("a:ext", NS)
                        coff, cext = g.find("a:chOff", NS), g.find("a:chExt", NS)
                        kx = int(ext.get("cx")) / max(1, int(cext.get("cx")))
                        ky = int(ext.get("cy")) / max(1, int(cext.get("cy")))
                        ox, oy = int(off.get("x")), int(off.get("y"))
                        cx, cy = int(coff.get("x")), int(coff.get("y"))
                        inner = lambda x, y, kx=kx, ky=ky, ox=ox, oy=oy, cx=cx, cy=cy: \
                            xf(ox + (x - cx) * kx, oy + (y - cy) * ky)
                        walk(el, inner, (sc[0] * kx, sc[1] * ky))
                    else:
                        walk(el, xf, sc)
                elif tag == "sp":
                    ph = el.find("p:nvSpPr/p:nvPr/p:ph", NS)
                    ptype = ph.get("type", "body") if ph is not None else None
                    if ptype in SKIP_PLACEHOLDERS:
                        continue
                    txt = [t for t in (_para_text(p) for p in el.iterfind(".//a:p", NS)) if t]
                    if not txt or all(any(f.search(t) for f in footer) for t in txt):
                        continue
                    (title if ptype in ("title", "ctrTitle") else body).extend(txt)
                elif tag == "pic":
                    vf = el.find("p:nvPicPr/p:nvPr/a:videoFile", NS)
                    if vf is None:
                        continue
                    target = rels.get(vf.get(R_LINK) or vf.get(R_EMBED) or "")
                    # Office 2010+ keeps the embedded copy in an extension.
                    for ext_el in el.iterfind("p:nvPicPr/p:nvPr/p:extLst/p:ext/*", NS):
                        if ext_el.tag.endswith("}media") and ext_el.get(R_EMBED) in rels:
                            target = rels[ext_el.get(R_EMBED)]
                    if not target or target not in z.namelist():
                        print(f"  ! video on {part} is linked, not embedded; skipped",
                              file=sys.stderr)
                        continue
                    x = el.find("p:spPr/a:xfrm", NS)
                    off, ext = x.find("a:off", NS), x.find("a:ext", NS)
                    x0, y0 = xf(int(off.get("x")), int(off.get("y")))
                    w, h = int(ext.get("cx")) * sc[0], int(ext.get("cy")) * sc[1]
                    name = el.find("p:nvPicPr/p:cNvPr", NS).get("name", "")
                    videos.append({
                        "member": target, "label": name,
                        "left": x0 / sw * 100, "top": y0 / sh * 100,
                        "width": w / sw * 100, "height": h / sh * 100,
                    })

        walk(root.find("p:cSld/p:spTree", NS))
        slides.append({"title": " ".join(title), "body": body, "videos": videos})
    return z, (sw, sh), slides


# --------------------------------------------------------------------------
# 5. HTML
# --------------------------------------------------------------------------

def pct(v):
    return f"{v:.3f}".rstrip("0").rstrip(".") + "%"


def box_style(d):
    return (f"left:{pct(d['left'])};top:{pct(d['top'])};"
            f"width:{pct(d['width'])};height:{pct(d['height'])}")


def slide_html(i, n, s, links, img_w, img_h, alt_override):
    e = html.escape
    num = f"{i:03d}"
    label = s["short"] or ""
    alt = alt_override or (f"Slide {i}: {label}" if label else f"Slide {i}")
    lazy = "" if i == 1 else 'loading="lazy" '
    parts = [
        f'<section class="slide" id="slide-{i}" data-title="{e(label)}" '
        f'aria-label="Slide {i} of {n}{": " + e(label) if label else ""}">',
        '<div class="frame">',
        f'<img src="slides/{num}.webp" width="{img_w}" height="{img_h}" alt="{e(alt)}" '
        f'{lazy}decoding="async">',
    ]
    for ln in links:
        parts.append(f'<a class="hot" href="{e(ln["url"])}" target="_blank" rel="noopener" '
                     f'title="{e(ln["url"])}" style="{box_style(ln)}">'
                     f'<span class="sr">{e(ln["url"])}</span></a>')
    for v in s["videos"]:
        parts.append(f'<button class="play" type="button" data-src="{e(v["src"])}" '
                     f'style="{box_style(v)}" aria-label="Play video{": " + e(v["label"]) if v["label"] else ""}">'
                     '<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="32" cy="32" r="30"/>'
                     '<path d="M26 20v24l19-12z"/></svg></button>')
    parts.append("</div>")
    text = []
    if s["title"]:
        text.append(f"<h2>{e(s['title'])}</h2>")
    text += [f"<p>{e(t)}</p>" for t in s["body"]]
    if alt_override:
        text.append(f"<p>{e(alt_override)}</p>")
    if text:
        parts.append(f'<div class="slide-text">{"".join(text)}</div>')
    parts.append("</section>")
    return "\n".join(parts)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("config", type=Path)
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--pdf", type=Path, help="use this PDF (exported from PowerPoint) instead of exporting")
    ap.add_argument("--html-only", action="store_true",
                    help="keep the existing slide images; only rewrite index.html and media "
                         "(quick, for template or config changes)")
    args = ap.parse_args()

    cfg = json.loads(args.config.read_text())
    if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", cfg.get("slug", "")):
        sys.exit("config 'slug' must be lower-case letters, digits and hyphens")
    out_dir = REPO / "presentations" / cfg["slug"]
    if args.html_only:
        if not (out_dir / "slides").is_dir():
            sys.exit(f"--html-only needs an existing build in {out_dir}")
    else:
        if out_dir.exists():
            shutil.rmtree(out_dir)
        out_dir.mkdir(parents=True)

    with tempfile.TemporaryDirectory() as t:
        tmp = Path(t)
        pdf = tmp / "deck.pdf"
        if args.pdf:
            shutil.copyfile(args.pdf, pdf)
        else:
            print("Exporting PDF with PowerPoint ...")
            export_pdf_with_powerpoint(args.pptx, pdf)

        page_w, page_h, n_pages = pdf_page_size(pdf)
        z, (sw, sh), slides = read_pptx(args.pptx, cfg.get("footer_patterns", []))
        if len(slides) != n_pages:
            sys.exit(f"PDF has {n_pages} pages but the deck has {len(slides)} visible slides")

        if not args.html_only:
            print(f"Rasterising {n_pages} slides ...")
            rasterise(pdf, out_dir, n_pages, tmp)
        links = collect_links(pdf, page_w, page_h, n_pages, cfg.get("ignore_url_patterns", []))

    # Embedded videos.
    for i, s in enumerate(slides, 1):
        for k, v in enumerate(s["videos"], 1):
            ext = Path(v["member"]).suffix.lower()
            ext = ".mp4" if ext in (".m4v", ".mp4") else ext
            name = f"slide{i:02d}-video{k}{ext}"
            (out_dir / "media").mkdir(exist_ok=True)
            (out_dir / "media" / name).write_bytes(z.read(v["member"]))
            v["src"] = f"media/{name}"
            v["label"] = re.sub(r"\.\w+$", "", v["label"]).replace("_", " ").strip()

    titles = {int(k): v for k, v in cfg.get("titles", {}).items()}
    alts = {int(k): v for k, v in cfg.get("alt", {}).items()}
    for i, s in enumerate(slides, 1):
        first = s["body"][0] if s["body"] else ""
        if len(first) > 60:
            first = first[:60].rsplit(" ", 1)[0] + "…"
        s["short"] = titles.get(i) or s["title"] or first

    img_w = SLIDE_WIDTH_PX
    img_h = round(SLIDE_WIDTH_PX * page_h / page_w)
    n = len(slides)
    body = "\n".join(slide_html(i, n, s, links[i - 1], img_w, img_h, alts.get(i))
                     for i, s in enumerate(slides, 1))

    url = cfg["site_url"].rstrip("/") + f"/presentations/{cfg['slug']}/"
    e = html.escape
    page = TEMPLATE.read_text()
    for key, val in {
        "__TITLE__": e(cfg["title"]),
        "__SUBTITLE__": e(cfg.get("subtitle", "")),
        "__AUTHOR__": e(cfg["author"]),
        "__EVENT__": e(cfg.get("event", "")),
        "__DATE__": e(cfg.get("date", "")),
        "__DESCRIPTION__": e(cfg.get("description", "")),
        "__URL__": e(url),
        "__AR_W__": str(round(page_w)),
        "__AR_H__": str(round(page_h)),
        "__COUNT__": str(n),
        "__SLIDES__": body,
    }.items():
        page = page.replace(key, val)
    (out_dir / "index.html").write_text(page)

    total = sum(f.stat().st_size for f in out_dir.rglob("*") if f.is_file())
    nlinks = sum(len(l) for l in links)
    nvid = sum(len(s["videos"]) for s in slides)
    print(f"Wrote {out_dir.relative_to(REPO)}/ : {n} slides, {nlinks} link areas, "
          f"{nvid} video(s), {total / 1e6:.1f} MB")


if __name__ == "__main__":
    main()
