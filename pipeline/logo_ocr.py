"""Logo OCR — extract text from cited mark logos so it can feed back into
the Word-axis matcher.

BR-IMG-001 (Phase 1 of the image-similarity workstream, 10 Jun 2026).

Many cited marks in the trademark scrape are tagged Type='Figurative' but
actually contain readable wordmark text rendered as graphics. Without OCR,
those marks get dropped by the Word-axis eligibility gate (the spreadsheet
metadata has no mark text to match against the operator's word_searches),
even though they're a genuine Word threat. Adding OCR recovers this class
of mark.

Strategy
--------
1. **Lazy-load** EasyOCR English reader once per process. Constructing the
   reader downloads ~70 MB of model weights and pulls in torch — we do
   NOT do this at import time so audits that skip OCR pay nothing.
2. **Multi-pass preprocessing** — each cited image is fed through four
   pipelines and OCR is run on each:
       - original colour
       - greyscale + Otsu threshold (Manchester recipe)
       - inverted greyscale + Otsu threshold (white-on-black wordmarks)
       - adaptive threshold (gradient backgrounds, uneven lighting)
3. **Union recognised tokens** — dedupe across passes; return the joined
   string for the caller to pass to mark_matches_any().

Failure-tolerant by design
--------------------------
Any error inside this module returns '' (an empty OCR result). The caller
treats that as "no OCR match" and the Word-axis gate falls back to its
mark-text-only check. OCR is a recall enhancement, never a correctness
requirement, so a missing easyocr install or a corrupt image MUST NOT
break the audit.
"""
from __future__ import annotations

import threading


# ---------------------------------------------------------------------------
# Lazy-loaded singleton EasyOCR reader.
# ---------------------------------------------------------------------------

_reader_lock = threading.Lock()
_reader = None
_reader_init_failed = False  # cache failures so we don't retry every call


def _get_reader():
    """Return a singleton EasyOCR reader, constructing it on first call.

    Returns None if the import or construction fails — callers must
    handle this case by returning '' (no OCR text).
    """
    global _reader, _reader_init_failed
    if _reader is not None:
        return _reader
    if _reader_init_failed:
        return None
    with _reader_lock:
        if _reader is not None:
            return _reader
        if _reader_init_failed:
            return None
        try:
            import easyocr  # heavy import — torch + model weights download
            # gpu=False forces CPU. Streamlit Cloud has no GPU. verbose=False
            # silences the progress bar so it doesn't spam Streamlit logs.
            _reader = easyocr.Reader(['en'], gpu=False, verbose=False)
        except Exception:
            # easyocr not installed, model download failed, OOM, etc.
            # Cache the failure so we don't retry 300 times in one audit.
            _reader_init_failed = True
            return None
    return _reader


# ---------------------------------------------------------------------------
# Multi-pass image preprocessing.
# ---------------------------------------------------------------------------

def _preprocess_passes(image_bytes: bytes):
    """Yield (label, image_array) pairs for multi-pass OCR.

    Each pass is a different binarisation that EasyOCR will read better
    than the raw image. Failures in any single pass are swallowed and the
    pass is skipped — at least one pass usually succeeds.

    Returns nothing if the image is undecodeable.
    """
    try:
        import cv2
        import numpy as np
    except Exception:
        # opencv not installed — fail silently
        return

    try:
        nparr = np.frombuffer(image_bytes, dtype=np.uint8)
        bgr = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        if bgr is None:
            return
    except Exception:
        return

    # Pass 1: original colour. EasyOCR handles colour reasonably well on
    # clean logos; this is the baseline pass.
    yield 'colour', bgr

    # Pass 2: greyscale + Otsu (Manchester's recipe). Good for dark text on
    # light background — the most common case.
    try:
        grey = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        _, otsu = cv2.threshold(grey, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
        yield 'otsu', otsu
    except Exception:
        pass

    # Pass 3: inverted greyscale + Otsu. Catches white-on-black wordmarks
    # which Pass 2 would binarise the wrong way round.
    try:
        grey = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        _, otsu_inv = cv2.threshold(grey, 0, 255,
                                     cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)
        yield 'otsu_inv', otsu_inv
    except Exception:
        pass

    # Pass 4: adaptive threshold. Catches gradient backgrounds and uneven
    # lighting where a global threshold fails. Block size 31 + C=10 is a
    # generic recipe; tighter values catch fine detail but increase noise.
    try:
        grey = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        adapt = cv2.adaptiveThreshold(grey, 255,
                                       cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                       cv2.THRESH_BINARY, 31, 10)
        yield 'adaptive', adapt
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Public API.
# ---------------------------------------------------------------------------

# Junk tokens that frequently appear in EasyOCR output for stylised logos
# but don't carry word-search signal. Mostly registration symbols and
# stray punctuation that survived the alnum filter.
_JUNK_TOKENS = {
    'TM', 'C', 'R',           # registration / copyright symbols misread
    'CO', 'LTD', 'LIMITED', 'LLP', 'INC', 'PLC', 'GROUP', 'GMBH', 'SA',
    'WWW', 'COM', 'NET', 'ORG', 'UK', 'EU', 'IO',
}


def extract_text_from_logo_bytes(image_bytes: bytes, *,
                                  min_confidence: float = 0.30) -> str:
    """Extract text from a logo image using multi-pass OCR.

    `image_bytes` is the raw JPEG/PNG bytes of the cited mark logo.
    Returns a single space-separated string of unique recognised tokens
    in upper case. Returns '' on any failure (caller treats as 'no OCR
    text').

    `min_confidence` (0.0–1.0) is the EasyOCR per-detection threshold.
    0.30 is permissive — better to over-recover and let the downstream
    word_search matcher filter than to miss a genuine stylised wordmark.

    The output is structured for `mark_matches_any()`: whitespace-
    separated tokens, upper-cased. The matcher then runs the same
    Exact / Starts With / Contains / Similar To logic against this
    OCR'd text as it does against the spreadsheet's mark column.
    """
    if not image_bytes:
        return ''

    reader = _get_reader()
    if reader is None:
        # easyocr unavailable — silently degrade
        return ''

    seen: set[str] = set()
    tokens: list[str] = []

    for _label, arr in _preprocess_passes(image_bytes):
        try:
            results = reader.readtext(arr, detail=1)
        except Exception:
            # OCR failed on this pass (likely a tiny image or weird format).
            # Skip to the next preprocessing pass — multi-pass is the whole
            # point.
            continue

        for entry in results:
            try:
                _bbox, text, conf = entry
            except Exception:
                continue
            if conf is None or conf < min_confidence:
                continue
            text = (text or '').strip()
            if not text:
                continue

            # Tokenise on whitespace + non-alnum so 'FOO-BAR' and 'FOO BAR'
            # both yield ['FOO', 'BAR']. The Word-axis matcher works on
            # whole tokens (Exact / Starts With) and substrings (Contains)
            # so this normalisation matches what the spreadsheet's mark
            # column looks like.
            for tok in text.upper().split():
                tok_clean = ''.join(c for c in tok if c.isalnum() or c == '-')
                tok_clean = tok_clean.strip('-')
                if len(tok_clean) < 2:
                    continue
                if tok_clean in _JUNK_TOKENS:
                    continue
                if tok_clean in seen:
                    continue
                seen.add(tok_clean)
                tokens.append(tok_clean)

    return ' '.join(tokens)


def reader_available() -> bool:
    """True if the EasyOCR reader can be constructed (deps present and
    model weights accessible). Useful for the report builder to know
    whether to surface 'OCR n/a' vs 'OCR yielded no text' messaging."""
    return _get_reader() is not None
