"""FastAPI server: serve the frontend and generate .pptx decks (Groq + images)."""

from __future__ import annotations

import hashlib
import io
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup, NavigableString
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from PIL import Image
from pypdf import PdfReader
from starlette.concurrency import run_in_threadpool

from dotenv import load_dotenv

# Load GROQ_API_KEY (and optional GROQ_MODEL) from a local .env file.
# Must run BEFORE importing presentation, which reads GROQ_MODEL at import time.
# On hosts like Render, real environment variables are used and .env is simply absent.
load_dotenv()

from presentation import DEFAULT_MODEL, Deck, build_pptx, generate_deck_groq  # noqa: E402

if not os.environ.get("GROQ_API_KEY"):
    print("WARNING: GROQ_API_KEY is not set. Add it to your .env file or host environment variables.")

app = FastAPI(title="Doc2PPT")

ROOT = Path(__file__).parent
OUTPUTS = ROOT / "outputs"
OUTPUTS.mkdir(exist_ok=True)

NAME_RE = re.compile(r"^[a-zA-Z0-9_\-\.]+$")

# Keep the prompt under Groq free-tier token limits. Lower this if you see rate-limit errors.
MAX_SOURCE_CHARS = int(os.environ.get("MAX_SOURCE_CHARS", "18000"))
MAX_IMAGES = 40  # max candidate images collected per source

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# ============================================================================
# Data types
# ============================================================================

@dataclass
class ImageRef:
    desc: str
    data: bytes | None = None        # already-downloaded image (PDF)
    url: str | None = None           # download lazily, only if the model picks it
    fallback_url: str | None = None
    context: str = ""                # text around the image, used for automatic matching


@dataclass
class Source:
    text: str
    images: dict[str, ImageRef] = field(default_factory=dict)
    label: str = ""


# ============================================================================
# Image helpers
# ============================================================================

def prepare_image(raw: bytes, min_w: int = 200, min_h: int = 150) -> bytes | None:
    """Validate, filter out icons/banners, downscale, and convert to PNG/JPEG for PPTX."""
    try:
        im = Image.open(io.BytesIO(raw))
        im.load()
    except Exception:
        return None
    w, h = im.size
    if w < min_w or h < min_h:
        return None
    if w / h > 6 or h / w > 6:  # thin banners / divider lines
        return None
    im.thumbnail((1600, 1600))
    buf = io.BytesIO()
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im.convert("RGBA").save(buf, "PNG")
    else:
        im.convert("RGB").save(buf, "JPEG", quality=85)
    return buf.getvalue()


def fetch_image(ref: ImageRef) -> bytes | None:
    if ref.data:
        return ref.data
    for u in (ref.url, ref.fallback_url):
        if not u:
            continue
        try:
            r = requests.get(u, headers=HEADERS, timeout=15)
            r.raise_for_status()
            ctype = r.headers.get("content-type", "")
            if not ctype.startswith("image/"):
                print(f"DEBUG image skipped (content-type {ctype!r}): {u}")
                continue
            if len(r.content) > 10_000_000:
                print(f"DEBUG image skipped (too large): {u}")
                continue
            prepared = prepare_image(r.content)
            if prepared:
                return prepared
            print(f"DEBUG image skipped (too small / unreadable): {u}")
        except Exception as exc:
            print(f"DEBUG image download failed: {u} -> {exc}")
            continue
    return None


def _tokens(s: str) -> set[str]:
    return {w for w in re.findall(r"[a-z0-9]{4,}", s.lower()) if w not in _STOPWORDS}


def auto_assign_images(deck: Deck, images: dict[str, ImageRef]) -> None:
    """Give slides without a model-chosen image the best-matching unused image.

    Matching is by word overlap between the slide's text and the text around
    each image (the slide text on SlideShare, the page text in a PDF, the
    caption/paragraph on a web page). This does not depend on the model.
    """
    used = {s.image_id for s in deck.slides if s.image_id}
    for s in deck.slides:
        if s.image_id:
            continue
        st = _tokens(s.title + " " + " ".join(s.bullets))
        best_id, best_score = "", 0.0
        for iid, ref in images.items():
            if iid in used:
                continue
            ct = _tokens(ref.context)
            if len(ct) < 3:
                continue
            overlap = len(st & ct)
            score = overlap / len(ct)
            if overlap >= 3 and score > best_score and score >= 0.25:
                best_id, best_score = iid, score
        if best_id:
            s.image_id = best_id
            used.add(best_id)


def attach_images(deck: Deck, images: dict[str, ImageRef]) -> None:
    """Pick images for slides, download them, and attach the bytes."""
    model_picked = sum(1 for s in deck.slides if s.image_id)
    auto_assign_images(deck, images)
    attached = 0
    for s in deck.slides:
        if s.image_id and s.image_id in images:
            s.image = fetch_image(images[s.image_id])
            if s.image:
                attached += 1
    print(
        f"DEBUG images: {len(images)} candidates, model picked {model_picked}, "
        f"{sum(1 for s in deck.slides if s.image_id)} matched in total, {attached} attached"
    )


# ============================================================================
# Source Loading (PDF + URL)
# ============================================================================

_STOPWORDS = {
    "the", "a", "an", "and", "or", "of", "to", "in", "on", "for", "with",
    "is", "are", "was", "were", "be", "by", "as", "at", "it", "this", "that",
}


def _relevance_trim(text: str, topic: str, budget: int = MAX_SOURCE_CHARS) -> str:
    """If text exceeds budget, keep the paragraphs most relevant to the topic."""
    if len(text) <= budget:
        return text

    keywords = {
        w for w in (topic.lower().split()) if len(w) > 2 and w not in _STOPWORDS
    }

    # Build paragraphs; glue [IMAGE ...] marker lines onto the paragraph before them
    paragraphs: list[str] = []
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("[IMAGE") and paragraphs:
            paragraphs[-1] += " " + line
        elif len(line) > 40:
            paragraphs.append(line)
    if not paragraphs:
        return text[:budget]

    def score(p: str) -> int:
        pl = p.lower()
        return sum(pl.count(k) for k in keywords)

    order = {id(p): i for i, p in enumerate(paragraphs)}
    ranked = sorted(paragraphs, key=score, reverse=True)
    kept, total = [], 0
    for p in ranked:
        if total + len(p) > budget:
            continue  # skip too-long ones but keep trying shorter ones
        kept.append(p)
        total += len(p)
    if not kept:
        return text[:budget]
    return "\n".join(sorted(kept, key=lambda p: order[id(p)]))


def extract_pdf(data: bytes) -> Source:
    reader = PdfReader(io.BytesIO(data))
    parts: list[str] = []
    images: dict[str, ImageRef] = {}
    seen: set[str] = set()
    has_text = False

    for pno, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").strip()
        if text:
            has_text = True

        tags: list[str] = []
        if len(images) < MAX_IMAGES:
            try:
                for img in page.images:
                    raw = img.data
                    digest = hashlib.md5(raw).hexdigest()
                    if digest in seen:  # logos / repeated images
                        continue
                    prepared = prepare_image(raw)
                    if prepared is None:
                        continue
                    seen.add(digest)
                    iid = f"p{pno}_{len(tags) + 1}"
                    images[iid] = ImageRef(
                        desc=f"figure on page {pno}", data=prepared, context=text[:1500]
                    )
                    tags.append(f"[IMAGE {iid}: figure on page {pno}]")
                    if len(images) >= MAX_IMAGES:
                        break
            except Exception:
                pass  # some PDFs have image formats pypdf can't decode

        block = "\n".join([text] + tags).strip()
        if block:
            parts.append(block)

    if not has_text:
        raise ValueError("No extractable text found in this PDF (it may be a scanned/image PDF).")
    return Source(text="\n\n".join(parts), images=images)


def _extract_slideshare(soup: BeautifulSoup, page_url: str) -> Source | None:
    """SlideShare pages hold each slide as an <img> whose alt text is the slide's text."""
    slides: list[tuple[str, str]] = []
    seen: set[str] = set()
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src") or ""
        if "image.slidesharecdn.com" not in src or src in seen:
            continue
        alt = " ".join((img.get("alt") or "").split())
        if alt:
            seen.add(src)
            slides.append((src, alt))
    if len(slides) < 3:
        return None

    title = soup.title.get_text(strip=True).split("|")[0].strip() if soup.title else page_url
    lines = [f"Source title: {title}", f"URL: {page_url}", ""]
    images: dict[str, ImageRef] = {}
    n = len(slides)
    for i, (src, alt) in enumerate(slides, start=1):
        line = f"Slide {i}: {alt}"
        # Slides with very little text are usually a figure/table/diagram saved as a picture
        if 1 < i < n and len(alt) < 150 and len(images) < MAX_IMAGES:
            iid = f"s{i}"
            images[iid] = ImageRef(
                desc=f"slide {i} visual",
                url=re.sub(r"-320\.(jpg|png|webp)", r"-1024.\1", src),
                fallback_url=src,
                context=alt,
            )
            line += f"\n[IMAGE {iid}: slide {i} is mostly a figure/table/diagram]"
        lines.append(line)
    return Source(text="\n".join(lines), images=images)


def extract_url(url: str) -> Source:
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise ValueError("Please provide a full http(s) URL.")

    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    final_url = resp.url
    soup = BeautifulSoup(resp.text, "html.parser")

    if "slideshare.net" in urlparse(final_url).netloc:
        ss = _extract_slideshare(soup, final_url)
        if ss:
            ss.label = final_url
            return ss

    for tag in soup(["script", "style", "noscript", "iframe", "svg", "header", "footer", "nav"]):
        tag.decompose()

    title = soup.title.get_text(strip=True) if soup.title else final_url
    main = soup.find("article") or soup.find("main") or soup.body or soup

    # Replace <img> tags with text markers so the model sees where each image sits
    images: dict[str, ImageRef] = {}
    for img in list(main.find_all("img")):
        src = img.get("data-src") or img.get("src") or ""
        if not src or src.startswith("data:") or src.lower().split("?")[0].endswith(".svg"):
            img.decompose()
            continue
        try:
            w = int(str(img.get("width") or "0").replace("px", "") or 0)
            h = int(str(img.get("height") or "0").replace("px", "") or 0)
        except ValueError:
            w = h = 0
        if (w and w < 150) or (h and h < 100) or len(images) >= MAX_IMAGES:
            img.decompose()
            continue
        alt = " ".join((img.get("alt") or "").split())
        fig = img.find_parent("figure")
        cap_tag = fig.find("figcaption") if fig else None
        caption = " ".join(cap_tag.get_text(" ").split()) if cap_tag else ""
        desc = (caption or alt or "image")[:120]
        iid = f"w{len(images) + 1}"
        near = img.find_parent(["p", "figure", "li"])
        near_text = " ".join(near.get_text(" ").split())[:400] if near else ""
        images[iid] = ImageRef(
            desc=desc,
            url=urljoin(final_url, src),
            context=f"{caption} {alt} {near_text}",
        )
        img.replace_with(NavigableString(f"\n[IMAGE {iid}: {desc}]\n"))

    text = "\n".join(
        line.strip() for line in main.get_text("\n").splitlines() if line.strip()
    )
    if len(text) < 100:
        raise ValueError("Could not extract readable content from this URL.")
    return Source(
        text=f"Source title: {title}\nURL: {final_url}\n\n{text}",
        images=images,
        label=final_url,
    )


def load_source(
    pdf_bytes: bytes | None, url: str | None, topic: str = "", pdf_name: str = ""
) -> Source:
    """Return extracted text (topic-relevance trimmed) plus candidate images."""
    if pdf_bytes:
        src = extract_pdf(pdf_bytes)
        src.label = pdf_name or "uploaded PDF"
    elif url and url.strip():
        src = extract_url(url.strip())
    else:
        raise ValueError("Provide a PDF file or a URL.")
    src.text = _relevance_trim(src.text, topic)
    return src


# ============================================================================
# API Routes
# ============================================================================

@app.post("/api/generate")
async def generate(
    topic: str = Form(...),
    num_slides: int = Form(6),
    model: str = Form(""),  # accepted for frontend compatibility, ignored
    url: str = Form(""),
    pdf: UploadFile | None = File(None),
):
    topic = topic.strip()
    if not topic:
        raise HTTPException(400, "Topic is required.")

    pdf_bytes = await pdf.read() if pdf is not None else None
    pdf_name = (pdf.filename or "") if pdf is not None else ""
    if pdf_bytes is not None and not pdf_bytes.strip():
        pdf_bytes = None

    if pdf_bytes is None and not url.strip():
        raise HTTPException(400, "Upload a PDF or provide a URL.")

    try:
        source = await run_in_threadpool(load_source, pdf_bytes, url, topic, pdf_name)
    except Exception as exc:
        raise HTTPException(400, f"Could not read source: {exc}") from exc

    print(
        f"DEBUG source: {len(source.text)} chars, {len(source.images)} images, "
        f"head={source.text[:300]!r}"
    )
    if len(source.text) < 300:
        raise HTTPException(
            400,
            "Could not extract enough text from this source. "
            "Try the direct page URL instead of a share link, or upload the PDF.",
        )

    # Only images whose marker survived trimming are visible to the model
    visible_ids = set(re.findall(r"\[IMAGE (\w+):", source.text))

    try:
        deck = await run_in_threadpool(
            generate_deck_groq,
            source.text,
            topic,
            max(3, min(num_slides, 30)),
            DEFAULT_MODEL,
            visible_ids,
        )
    except Exception as exc:
        raise HTTPException(502, str(exc)) from exc

    deck.source = source.label
    await run_in_threadpool(attach_images, deck, source.images)

    pptx_bytes = build_pptx(deck)
    slug = re.sub(r"[^a-z0-9]+", "_", deck.deck_title.lower()).strip("_")[:60] or "deck"
    filename = f"{slug}_{datetime.now():%Y%m%d_%H%M%S}.pptx"
    (OUTPUTS / filename).write_bytes(pptx_bytes)

    return {
        "deck_title": deck.deck_title,
        "subtitle": deck.subtitle,
        "filename": filename,
        "image_count": sum(1 for s in deck.slides if s.image),
        "slides": [
            {
                "title": s.title,
                "bullets": s.bullets,
                "notes": s.notes,
                "has_image": bool(s.image),
            }
            for s in deck.slides
        ],
    }


@app.get("/api/download/{filename}")
def download(filename: str):
    if not NAME_RE.match(filename) or not filename.endswith(".pptx"):
        raise HTTPException(400, "Bad filename.")
    path = OUTPUTS / filename
    if not path.is_file():
        raise HTTPException(404, "File not found.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=filename,
    )


@app.get("/")
def index():
    return FileResponse(ROOT / "static" / "index.html")


app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))