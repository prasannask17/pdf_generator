"""FastAPI server: serve the frontend and generate .pptx decks (Groq + images)."""

from __future__ import annotations

import hashlib
import io
import json
import logging
import os
import re
import threading
import uuid
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

# ============================================================================
# Logging — writes to both console (Render logs) and a persistent file
# ============================================================================

LOG_FILE = ROOT / "search_logs.jsonl"

logger = logging.getLogger("doc2ppt")
logger.setLevel(logging.INFO)
logger.handlers.clear()  # avoid duplicate handlers on reload

# Console handler — shows in Render logs and terminal
_console = logging.StreamHandler()
_console.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"))
logger.addHandler(_console)

# File handler — appends every search as a JSON line for persistence
try:
    _file = logging.FileHandler(str(LOG_FILE))
    _file.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(_file)
except Exception:
    pass  # read-only filesystem on some hosts


def log_search(**fields) -> None:
    """Log a structured search record to console and file."""
    fields["timestamp"] = datetime.now().isoformat(timespec="seconds")
    logger.info(json.dumps(fields, ensure_ascii=False))


NAME_RE = re.compile(r"^[a-zA-Z0-9_\-\.]+$")

# Keep the prompt under Groq free-tier token limits. Lower this if you see rate-limit errors.
MAX_SOURCE_CHARS = int(os.environ.get("MAX_SOURCE_CHARS", "30000"))
MAX_IMAGES = 40  # max candidate images collected per source

# ============================================================================
# Background job state (avoids Render's 60s request timeout)
# ============================================================================

JOB_STATE: dict[str, dict] = {}
JOB_LOCK = threading.Lock()
JOB_ID_RE = re.compile(r"^[a-f0-9]{16}$")


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _update_job(job_id: str, **changes) -> None:
    with JOB_LOCK:
        job = JOB_STATE.get(job_id)
        if job is None:
            return
        job.update(changes)
        job["updated_at"] = _now()

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
    has_text = False

    # NOTE: Image extraction is disabled by default. pypdf's page.images
    # decompresses all embedded images into memory, which causes OOM and
    # hangs on Render's 512MB free tier — especially for slide-deck PDFs
    # where every page has a full-page background image. Most PDFs that
    # users upload are text-based and don't need image extraction.
    # To re-enable, set the env var EXTRACT_PDF_IMAGES=1.

    extract_images = os.environ.get("EXTRACT_PDF_IMAGES", "0") == "1"

    for pno, page in enumerate(reader.pages, start=1):
        text = (page.extract_text() or "").strip()
        if text:
            has_text = True

        tags: list[str] = []
        if extract_images and len(images) < MAX_IMAGES:
            try:
                # Get page dimensions for background-image detection
                mb = page.mediabox
                page_w, page_h = float(mb.width), float(mb.height)

                for img in page.images:
                    raw = img.data
                    digest = hashlib.md5(raw).hexdigest()
                    if digest in seen:  # logos / repeated images
                        continue
                    if _is_background_image(img, page_w, page_h):
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


def _is_background_image(img, page_w: float, page_h: float) -> bool:
    """Detect if an embedded image is a full-page background (not a figure).

    Compares the image's aspect ratio to the page's aspect ratio. If they
    match closely and the image is large, it's likely a slide/page background.
    """
    if not page_w or not page_h:
        return False
    try:
        from PIL import Image as PILImage
        im = PILImage.open(io.BytesIO(img.data))
        iw, ih = im.size
    except Exception:
        return False
    # Skip small images (icons, logos) — those aren't backgrounds
    if iw < 400 or ih < 300:
        return False
    page_ar = page_w / page_h
    img_ar = iw / ih
    # If the image aspect ratio matches the page within 5%, it's a background
    return abs(img_ar - page_ar) / page_ar < 0.05


# ---------------------------------------------------------------------------
# Unit filtering — detect "unit N" in the topic and keep only matching pages
# ---------------------------------------------------------------------------

_UNIT_RE = re.compile(r"\bunit\s*([0-9]+)\b", re.IGNORECASE)


def filter_by_unit(source: Source, topic: str) -> Source:
    """If the topic mentions 'unit N', keep only pages whose header matches.

    Works for PDFs where each page starts with a line like 'UNIT 1 · TITLE'.
    If no unit is mentioned, or no unit headers are found in the text, returns
    the source unchanged.
    """
    match = _UNIT_RE.search(topic)
    if not match:
        return source

    unit_num = match.group(1)
    unit_pattern = re.compile(rf"^UNIT\s*{unit_num}\b", re.IGNORECASE)

    # Split source text into page blocks (they were joined with \n\n)
    blocks = source.text.split("\n\n")
    kept_blocks: list[str] = []
    kept_page_nums: set[int] = set()

    for block in blocks:
        first_line = block.strip().split("\n")[0].strip()
        if unit_pattern.match(first_line):
            kept_blocks.append(block)
            # Track page number from image markers like [IMAGE p5_1: ...]
            for m in re.finditer(r"\[IMAGE p(\d+)_", block):
                kept_page_nums.add(int(m.group(1)))

    if not kept_blocks:
        return source  # no unit headers found — don't filter

    # Filter images to only those from kept pages
    filtered_images: dict[str, ImageRef] = {}
    for iid, ref in source.images.items():
        m = re.match(r"p(\d+)_", iid)
        if m and int(m.group(1)) in kept_page_nums:
            filtered_images[iid] = ref

    new_source = Source(
        text="\n\n".join(kept_blocks),
        images=filtered_images,
        label=source.label,
    )
    print(
        f"DEBUG unit filter: topic mentions unit {unit_num}, "
        f"kept {len(kept_blocks)} page blocks, {len(filtered_images)} images"
    )
    return new_source


def filter_by_page_range(source: Source, page_start: int, page_end: int) -> Source:
    """Keep only pages within [page_start, page_end] (1-indexed).

    Page numbers are extracted from image markers like [IMAGE p5_1: ...].
    Text blocks without image markers are kept (we can't determine their page
    number, so we err on the side of inclusion — _relevance_trim handles
    trimming later).
    """
    if not page_start and not page_end:
        return source

    start = page_start if page_start > 0 else 1
    end = page_end if page_end > 0 else 999999

    blocks = source.text.split("\n\n")
    kept_blocks: list[str] = []
    kept_page_nums: set[int] = set()

    for block in blocks:
        # Find page numbers from image markers in this block
        block_pages = set(int(m.group(1)) for m in re.finditer(r"\[IMAGE p(\d+)_", block))
        if block_pages:
            # Keep block if any of its images are in range
            in_range = any(start <= p <= end for p in block_pages)
            if in_range:
                kept_blocks.append(block)
                kept_page_nums.update(p for p in block_pages if start <= p <= end)
        else:
            # No image markers — can't determine page number, keep it
            kept_blocks.append(block)

    if not kept_blocks:
        return source

    # Filter images to only those in range
    filtered_images: dict[str, ImageRef] = {}
    for iid, ref in source.images.items():
        m = re.match(r"p(\d+)_", iid)
        if m and start <= int(m.group(1)) <= end:
            filtered_images[iid] = ref

    new_source = Source(
        text="\n\n".join(kept_blocks),
        images=filtered_images,
        label=source.label,
    )
    print(
        f"DEBUG page range filter: pages {start}-{end}, "
        f"kept {len(kept_blocks)} blocks, {len(filtered_images)} images"
    )
    return new_source


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

    # For SlideShare, use a scraping-friendly headers set
    is_slideshare = "slideshare.net" in parsed.netloc

    fetch_headers = HEADERS.copy()
    if is_slideshare:
        fetch_headers.update({
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Upgrade-Insecure-Requests": "1",
        })

    resp = requests.get(url, headers=fetch_headers, timeout=30)
    resp.raise_for_status()
    final_url = resp.url
    soup = BeautifulSoup(resp.text, "html.parser")

    if "slideshare.net" in urlparse(final_url).netloc:
        ss = _extract_slideshare(soup, final_url)
        if ss:
            ss.label = final_url
            return ss
        # SlideShare blocked normal scraping — try extracting text directly from page
        raw_text = soup.get_text(" ", strip=True)
        # Look for slide content patterns
        slide_texts = re.findall(r'\d+\s*/\s*\d+\s*(.*?)(?=\d+\s*/\s*\d+|$)', raw_text, re.DOTALL)
        if slide_texts and len(" ".join(slide_texts)) > 500:
            combined = "\n".join(t.strip() for t in slide_texts if t.strip())
            return Source(text=combined, images={}, label=final_url)
        # Last resort: use all meaningful text from the page
        lines = [l.strip() for l in raw_text.split("\n") if len(l.strip()) > 40]
        combined = "\n".join(lines)
        if len(combined) > 600:
            return Source(text=combined[:MAX_SOURCE_CHARS], images={}, label=final_url)
        raise ValueError(
            "SlideShare blocked this request. Please download the PDF from SlideShare and upload it, "
            "or use the 'Paste Text' tab to paste the slide content."
        )

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
    pdf_bytes: bytes | None,
    url: str | None,
    topic: str = "",
    pdf_name: str = "",
    page_start: int = 0,
    page_end: int = 0,
) -> Source:
    """Return extracted text (topic-relevance trimmed) plus candidate images."""
    if pdf_bytes:
        src = extract_pdf(pdf_bytes)
        src.label = pdf_name or "uploaded PDF"
    elif url and url.strip():
        src = extract_url(url.strip())
    else:
        raise ValueError("Provide a PDF file or a URL.")
    # Apply page range first, then unit filter (both only for PDFs)
    if pdf_bytes:
        src = filter_by_page_range(src, page_start, page_end)
        src = filter_by_unit(src, topic)
    src.text = _relevance_trim(src.text, topic)
    return src


# ============================================================================
# API Routes — Background job pattern (avoids Render's 60s request timeout)
# ============================================================================

def _run_job(
    job_id: str,
    pdf_bytes: bytes | None,
    url: str,
    text: str,
    topic: str,
    num_slides: int,
    pdf_name: str,
    page_start: int,
    page_end: int,
) -> None:
    """Background worker: extract source, generate deck, build PPTX, update job state."""
    _update_job(job_id, status="running", stage="extracting", message="Reading PDF and extracting text…")
    try:
        # Step 1: Load source (PDF extraction, URL scraping, or text)
        if text and pdf_bytes is None and not url.strip():
            source = Source(
                text=_relevance_trim(text, topic),
                images={},
                label="pasted text",
            )
        else:
            try:
                source = load_source(pdf_bytes, url or None, topic, pdf_name, page_start, page_end)
            except Exception as exc:
                _update_job(job_id, status="error", stage="error", message="Failed", error=f"Could not read source: {exc}")
                return

        # Free the raw PDF bytes — no longer needed after extraction
        del pdf_bytes

        print(
            f"DEBUG source: {len(source.text)} chars, {len(source.images)} images, "
            f"head={source.text[:300]!r}"
        )
        if len(source.text) < 600 or "couldn't load" in source.text.lower() or "client challenge" in source.text.lower():
            _update_job(job_id, status="error", stage="error", message="Failed",
                        error="Could not extract enough content from this URL. SlideShare and similar sites block automated access — please download the PDF and upload it instead.")
            return

        # Step 2: Generate deck with Groq
        _update_job(job_id, stage="generating", message="Generating slides with AI…")
        visible_ids = set(re.findall(r"\[IMAGE (\w+):", source.text))
        deck = generate_deck_groq(
            source.text, topic, max(3, min(num_slides, 20)),
            DEFAULT_MODEL, visible_ids,
        )
        deck.source = source.label

        # Step 3: Attach images
        _update_job(job_id, stage="images", message="Attaching images…")
        attach_images(deck, source.images)

        # Step 4: Build PPTX
        _update_job(job_id, stage="building", message="Building PPTX file…")
        pptx_bytes = build_pptx(deck)
        slug = re.sub(r"[^a-z0-9]+", "_", deck.deck_title.lower()).strip("_")[:60] or "deck"
        filename = f"{slug}_{datetime.now():%Y%m%d_%H%M%S}.pptx"
        (OUTPUTS / filename).write_bytes(pptx_bytes)

        result = {
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
        _update_job(job_id, status="done", stage="done", message="Finished", result=result)
        log_search(
            event="job_done",
            job_id=job_id,
            topic=topic,
            deck_title=deck.deck_title,
            slides=len(deck.slides),
            images=result["image_count"],
            source_chars=len(source.text),
            filename=filename,
        )
    except Exception as exc:
        _update_job(job_id, status="error", stage="error", message="Failed", error=str(exc))
        log_search(
            event="job_error",
            job_id=job_id,
            topic=topic,
            error=str(exc),
        )


@app.post("/api/generate")
async def generate(
    topic: str = Form(...),
    num_slides: int = Form(6),
    model: str = Form(""),  # accepted for frontend compatibility, ignored
    url: str = Form(""),
    text: str = Form(""),
    page_start: int = Form(0),
    page_end: int = Form(0),
    pdf: UploadFile | None = File(None),
):
    topic = topic.strip()
    if not topic:
        raise HTTPException(400, "Topic is required.")

    pdf_bytes = await pdf.read() if pdf is not None else None
    pdf_name = (pdf.filename or "") if pdf is not None else ""
    if pdf_bytes is not None and not pdf_bytes.strip():
        pdf_bytes = None

    text = text.strip()

    if pdf_bytes is None and not url.strip() and not text:
        raise HTTPException(400, "Upload a PDF, provide a URL, or paste some text.")

    if text and pdf_bytes is None and not url.strip() and len(text) < 100:
        raise HTTPException(400, "Pasted text is too short. Please provide more content.")

    # Determine source type for logging
    source_type = "pdf" if pdf_bytes else ("url" if url.strip() else "text")
    pdf_size_mb = round(len(pdf_bytes) / (1024 * 1024), 2) if pdf_bytes else 0

    # Create a background job immediately — all heavy work (PDF extraction,
    # Groq API call, PPTX building) runs in the thread, not in this request.
    # This avoids Render's 60s request timeout.
    job_id = uuid.uuid4().hex[:16]

    log_search(
        event="job_submitted",
        job_id=job_id,
        topic=topic,
        source_type=source_type,
        pdf_name=pdf_name or None,
        pdf_size_mb=pdf_size_mb or None,
        url=url.strip() or None,
        text_length=len(text) if text else 0,
        num_slides=num_slides,
        page_start=page_start or None,
        page_end=page_end or None,
    )

    with JOB_LOCK:
        JOB_STATE[job_id] = {
            "id": job_id,
            "status": "queued",
            "stage": "queued",
            "message": "Queued",
            "error": None,
            "result": None,
            "created_at": _now(),
            "updated_at": _now(),
        }

    thread = threading.Thread(
        target=_run_job,
        args=(job_id, pdf_bytes, url, text, topic, num_slides, pdf_name, page_start, page_end),
        daemon=True,
    )
    thread.start()

    return {"job_id": job_id, "status_url": f"/api/jobs/{job_id}"}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    if not JOB_ID_RE.match(job_id):
        raise HTTPException(400, "Bad job id.")
    with JOB_LOCK:
        job = JOB_STATE.get(job_id)
    if job is None:
        raise HTTPException(404, "Job not found.")
    return job


@app.get("/api/logs")
def get_logs(limit: int = 50):
    """Return recent search logs as JSON."""
    if not LOG_FILE.is_file():
        return {"logs": []}
    try:
        lines = LOG_FILE.read_text(encoding="utf-8").strip().splitlines()
    except Exception:
        return {"logs": []}
    # Parse each line as JSON, return most recent first
    logs = []
    for line in reversed(lines):
        line = line.strip()
        if not line:
            continue
        try:
            logs.append(json.loads(line))
        except Exception:
            continue
        if len(logs) >= limit:
            break
    return {"logs": logs, "total": len(lines)}


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


# ============================================================================
# ============================================================================
# PDF export — pure Python conversion via python-pptx + reportlab
# No system dependencies required; works on Render free tier.
# ============================================================================

def _pptx_to_pdf(pptx_path: Path, pdf_path: Path) -> None:
    """Convert a PPTX to a styled PDF matching the app theme.

    Shape naming convention used by build_pptx():
      Title slide : TextBox 2 = deck title, TextBox 3 = subtitle, TextBox 4 = source
      Content slide: TextBox 1 = slide title, TextBox 3 = bullets (multi-para),
                     TextBox 4 = page number, Rectangle 2 = accent bar (no text)
    """
    from pptx import Presentation as PptxPresentation
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader
    import textwrap

    prs = PptxPresentation(str(pptx_path))
    PW, PH = landscape(A4)   # 841 × 595 pt
    M = 0.5 * inch

    # ── Theme palette (matches presentation.py) ──────────────────────────────
    C_DARK   = colors.HexColor("#1F2937")   # header / title-slide bg
    C_ACCENT = colors.HexColor("#2563EB")   # blue bar
    C_WHITE  = colors.HexColor("#FFFFFF")
    C_BODY   = colors.HexColor("#1F2937")   # body text on light bg
    C_MUTED  = colors.HexColor("#6B7280")
    C_LIGHT  = colors.HexColor("#D1D5DB")   # subtitle on dark bg
    C_BG     = colors.HexColor("#F9FAFB")   # content slide bg
    C_DOT    = colors.HexColor("#2563EB")   # bullet dot

    # ── Text sanitizer ────────────────────────────────────────────────────────
    # Helvetica uses WinAnsiEncoding; characters outside it (e.g. the
    # non-breaking hyphen U+2011 produced by the model) render as black boxes.
    _UNICODE_MAP = {
        "\u2010": "-", "\u2011": "-", "\u2012": "-", "\u2013": "-",
        "\u2014": "-", "\u2015": "-", "\u2212": "-", "\u00ad": "",
        "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
        "\u2022": "-", "\u25cf": "-", "\u00b7": "-", "\u2026": "...",
        "\u00a0": " ", "\u200b": "", "\u2192": "->", "\u2190": "<-",
        "\u2265": ">=", "\u2264": "<=", "\u00d7": "x",
    }

    def _clean(text) -> str:
        s = "".join(_UNICODE_MAP.get(ch, ch) for ch in str(text))
        return s.encode("cp1252", errors="replace").decode("cp1252")

    # ── Helpers ───────────────────────────────────────────────────────────────
    def _wrap(text: str, x, y, w, c_obj, font, size, clr, lh, max_l=99):
        c_obj.setFont(font, size)
        c_obj.setFillColor(clr)
        cpl = max(1, int(w / (size * 0.50)))
        lines, drawn = [], 0
        for para in str(text).split("\n"):
            lines += textwrap.wrap(para.strip(), cpl) or [""]
        for line in lines:
            if drawn >= max_l:
                break
            c_obj.drawString(x, y, line)
            y -= lh
            drawn += 1
        return y

    def _bullets(items, x, y, w, c_obj, size, lh, bottom):
        cpl = max(1, int((w - 20) / (size * 0.50)))
        placed = 0
        for i, item in enumerate(items):
            wrapped = textwrap.wrap(item.strip(), cpl) or [""]
            needed = len(wrapped) * lh + 3
            if y - needed < bottom and placed > 0:
                return items[i:]
            # dot
            c_obj.setFillColor(C_DOT)
            c_obj.circle(x + 6, y + size * 0.28, 2.8, fill=1, stroke=0)
            # text
            c_obj.setFillColor(C_BODY)
            c_obj.setFont("Helvetica", size)
            for line in wrapped:
                c_obj.drawString(x + 18, y, line)
                y -= lh
            y -= 3
            placed += 1
        return []

    # ── Extract slide data ────────────────────────────────────────────────────
    slides_data: list[dict] = []
    for slide in prs.slides:
        shapes = {s.name: s for s in slide.shapes}
        notes_text = ""
        if slide.has_notes_slide:
            try:
                notes_text = _clean((slide.notes_slide.notes_text_frame.text or "").strip())
            except Exception:
                pass

        def texts_of(name: str) -> list[str]:
            s = shapes.get(name)
            if s and s.has_text_frame:
                return [_clean(p.text.strip()) for p in s.text_frame.paragraphs if p.text.strip()]
            return []

        # Detect title slide: has TextBox 2 (deck title), no TextBox 1
        if "TextBox 2" in shapes and "TextBox 1" not in shapes:
            slides_data.append({
                "kind": "title",
                "title": (texts_of("TextBox 2") or [""])[0],
                "subtitle": (texts_of("TextBox 3") or [""])[0],
                "source": (texts_of("TextBox 4") or [""])[0],
                "notes": notes_text,
                "img": None,
            })
        else:
            # Content slide
            title_texts = texts_of("TextBox 1")
            bullet_texts = texts_of("TextBox 3")
            # Collect embedded pictures
            img_bytes = None
            for s in slide.shapes:
                if s.shape_type == 13 and img_bytes is None:
                    try:
                        img_bytes = s.image.blob
                    except Exception:
                        pass
            slides_data.append({
                "kind": "content",
                "title": (title_texts or [""])[0],
                "bullets": bullet_texts,
                "notes": notes_text,
                "img": img_bytes,
            })

    # ── Render ────────────────────────────────────────────────────────────────
    c = rl_canvas.Canvas(str(pdf_path), pagesize=landscape(A4))
    total = len(slides_data)

    for idx, sd in enumerate(slides_data):

        if sd["kind"] == "title":
            # ── TITLE SLIDE ──────────────────────────────────────────────────
            # Full dark background
            c.setFillColor(C_DARK)
            c.rect(0, 0, PW, PH, fill=1, stroke=0)

            # Left colour bar (vertical accent)
            c.setFillColor(C_ACCENT)
            c.rect(M, PH * 0.25, 0.06 * inch, PH * 0.5, fill=1, stroke=0)

            # Deck title — large, white, left of centre
            tx = M + 0.22 * inch
            ty = PH * 0.65
            _wrap(sd["title"] or "Presentation", tx, ty, PW * 0.8,
                  c, "Helvetica-Bold", 38, C_WHITE, 46, max_l=3)

            # Subtitle
            if sd["subtitle"]:
                _wrap(sd["subtitle"], tx, PH * 0.37, PW * 0.8,
                      c, "Helvetica", 18, C_LIGHT, 24, max_l=2)

            # Source
            if sd["source"]:
                c.setFont("Helvetica-Oblique", 10)
                c.setFillColor(C_MUTED)
                src = sd["source"]
                cpl = int((PW - tx - M) / (10 * 0.50))
                c.drawString(tx, M + 0.1 * inch,
                             src if len(src) <= cpl else src[:cpl - 1] + "…")

            c.showPage()

        else:
            # ── CONTENT SLIDE (may span multiple pages) ──────────────────────
            has_img = sd["img"] is not None
            title = sd["title"] or f"Slide {idx + 1}"
            notes = sd["notes"]
            HH = 0.78 * inch
            notes_h = 0.28 * inch if notes else 0.12 * inch
            ct = PH - HH - 0.05 * inch - 0.28 * inch   # content top
            cb = notes_h + 0.04 * inch                   # content bottom
            ch = ct - cb                                  # full content height

            if has_img:
                text_w_first = PW * 0.52 - M
                img_x = PW * 0.54
                img_w = PW - img_x - M * 0.5
            else:
                text_w_first = PW - 2 * M

            # Pick a bullet font size that fits the first-page column
            bullets = list(sd["bullets"])
            fs, lh = 10, 10 * 1.55
            for cand in (18, 16, 14, 13, 12, 11, 10):
                lh_c = cand * 1.55
                cpl = max(1, int((text_w_first - 20) / (cand * 0.50)))
                est = sum(len(textwrap.wrap(b.strip(), cpl)) or 1 for b in bullets) * lh_c
                if est <= ch * 0.92:
                    fs, lh = cand, lh_c
                    break

            page_items = bullets
            first = True
            while True:
                # Light background
                c.setFillColor(C_BG)
                c.rect(0, 0, PW, PH, fill=1, stroke=0)

                # Dark header bar
                c.setFillColor(C_DARK)
                c.rect(0, PH - HH, PW, HH, fill=1, stroke=0)

                # Blue accent strip under header
                c.setFillColor(C_ACCENT)
                c.rect(0, PH - HH - 0.05 * inch, PW, 0.05 * inch, fill=1, stroke=0)

                # Title text inside header (white, bold, 24pt)
                hdr = title if first else f"{title} (cont.)"
                cpl_title = int((PW - 2 * M) / (24 * 0.52))
                short = hdr if len(hdr) <= cpl_title else hdr[:cpl_title - 1] + "…"
                c.setFont("Helvetica-Bold", 24)
                c.setFillColor(C_WHITE)
                c.drawString(M, PH - HH + (HH - 24) / 2, short)

                text_w = text_w_first if (first and has_img) else (PW - 2 * M)
                leftover = _bullets(page_items, M, ct, text_w, c, fs, lh, cb)

                # Image — right column, first page only
                if first and has_img:
                    try:
                        buf = io.BytesIO(sd["img"])
                        pil = Image.open(buf)
                        pil.thumbnail((2000, 2000))
                        out = io.BytesIO()
                        fmt = "PNG" if pil.mode in ("RGBA", "LA", "P") else "JPEG"
                        pil.save(out, fmt)
                        out.seek(0)
                        iw, ih = pil.size
                        scale = min(img_w / iw, ch / ih)
                        dw, dh = iw * scale, ih * scale
                        dx = img_x + (img_w - dw) / 2
                        dy = ct - dh
                        c.setFillColor(colors.HexColor("#E2E8F0"))
                        c.roundRect(dx - 2, dy - 2, dw + 4, dh + 4, 4, fill=1, stroke=0)
                        c.drawImage(ImageReader(out), dx, dy, dw, dh,
                                    preserveAspectRatio=True, mask="auto")
                    except Exception as e:
                        print(f"DEBUG pdf img error slide {idx}: {e}")

                # Notes footer — first page only
                if first and notes:
                    c.setFillColor(colors.HexColor("#F3F4F6"))
                    c.rect(0, 0, PW, notes_h + 0.04 * inch, fill=1, stroke=0)
                    c.setStrokeColor(colors.HexColor("#E5E7EB"))
                    c.setLineWidth(0.5)
                    c.line(0, notes_h + 0.04 * inch, PW, notes_h + 0.04 * inch)
                    c.setFont("Helvetica-Oblique", 8.5)
                    c.setFillColor(C_MUTED)
                    note = notes[:180].replace("\n", " ")
                    cpl_n = int((PW - 2 * M) / (8.5 * 0.50))
                    c.drawString(M, 0.09 * inch, note if len(note) <= cpl_n else note[:cpl_n - 1] + "…")

                # Slide number (bottom-right)
                c.setFont("Helvetica", 8)
                c.setFillColor(C_MUTED)
                label = f"{idx + 1}" if first else f"{idx + 1} (cont.)"
                c.drawRightString(PW - M * 0.5, 0.09 * inch, f"{label} / {total}")

                c.showPage()
                if not leftover:
                    break
                page_items = leftover
                first = False

    c.save()


@app.get("/api/download-pdf/{filename}")
def download_pdf(filename: str):
    """Convert an already-generated .pptx to PDF using pure Python (no LibreOffice)."""
    if not NAME_RE.match(filename) or not filename.endswith(".pptx"):
        raise HTTPException(400, "Bad filename.")
    pptx_path = OUTPUTS / filename
    if not pptx_path.is_file():
        raise HTTPException(404, "PPTX file not found.")

    pdf_filename = filename.replace(".pptx", ".pdf")
    pdf_path = OUTPUTS / pdf_filename

    # Always regenerate so fixes to the converter apply to existing decks
    try:
        _pptx_to_pdf(pptx_path, pdf_path)
    except Exception as exc:
        # Clean up partial file
        pdf_path.unlink(missing_ok=True)
        raise HTTPException(500, f"PDF conversion failed: {exc}")

    return FileResponse(
        pdf_path,
        media_type="application/pdf",
        filename=pdf_filename,
    )


@app.get("/")
def index():
    return FileResponse(ROOT / "static" / "index.html")


app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))