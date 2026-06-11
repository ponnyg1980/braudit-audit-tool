"""Visual similarity between the client logo and each cited mark logo.

BR-IMG-007 (Phase 2-bis, 11 Jun 2026) — ONNX runtime rewrite.

This is the same module API as the original torch + OpenCLIP version
(BR-IMG-002), but the backbone is now ONNX runtime + a quantised CLIP
ViT-B/16 vision tower. The torch path was withdrawn because the
`--extra-index-url` flag broke Streamlit Cloud's pip resolver, taking
down the whole app. ONNX runtime has no such packaging quirk.

Stack
-----
* `onnxruntime` (~80 MB on disk) instead of `torch + torchvision` (~800 MB).
* OpenAI CLIP ViT-B/16 vision tower in INT8-quantised ONNX form
  (~83 MB cached weights, downloaded once from HuggingFace Hub).
* CPU inference at ~30 ms per image (vs ~180 ms for torch).
* Peak RAM ~400 MB instead of ~1.3 GB — fits comfortably on Streamlit
  Cloud's 1 GB ceiling.

Thresholds were re-calibrated against the new CLIP variant because OpenAI
CLIP's embedding space has a narrower cosine distribution than the
OpenCLIP LAION-2B model the torch version used — random unrelated pairs
score in the 0.55-0.65 band rather than 0.40-0.50.

The public API (`score_visual_similarity_batch`, `model_available`,
`unload_clip`) is unchanged from the torch version, so process_trademarks
needs no edits.

Failure-tolerant by design
--------------------------
Any error (model download fails, weights load fails, image decode fails)
returns `decision = 'skipped'` for the affected row. The caller falls
back to the base Logo-axis score and the audit continues without CLIP.
"""
from __future__ import annotations

import gc
import os
import threading
from io import BytesIO
from typing import Optional


# ---------------------------------------------------------------------------
# Thresholds — calibrated for OpenAI CLIP ViT-B/16 INT8-quantised.
# Re-calibrate if the model or weights change. See module docstring for
# context on why these are higher than the torch-era values.
# ---------------------------------------------------------------------------

PHASH_DISTANCE_IDENTICAL = 5      # ≤ this → 'identical' bucket (no CLIP needed)
PHASH_DISTANCE_UNRELATED = 30     # > this → 'unrelated' bucket (no CLIP run)

# OpenAI CLIP ViT-B/16 INT8 calibration (11 Jun 2026):
#   Friars cited marks (presumed unrelated) had cosine mean 0.55, max 0.69.
#   Self-similarity was 1.000.
#   Identical (visually) is expected to be ≥ 0.92; mild similarity 0.85-0.91;
#   moderate 0.78-0.84; below 0.78 = unrelated.
CLIP_COSINE_IDENTICAL = 0.92      # ≥ this → 'identical' (strong visual match)
CLIP_COSINE_SIMILAR   = 0.85      # 0.85-0.91 → 'similar' (mild)
CLIP_COSINE_WEAK      = 0.78      # 0.78-0.84 → 'weak' (background score level)
# below CLIP_COSINE_WEAK → 'unrelated'


# ---------------------------------------------------------------------------
# Model source — Xenova's pre-converted OpenAI CLIP ViT-B/16 on HuggingFace.
# We use the INT8-quantised vision tower: 83 MB on disk, ~30 ms inference,
# accuracy within a hair of full precision for our threshold-based use.
# ---------------------------------------------------------------------------

CLIP_HF_REPO = 'Xenova/clip-vit-base-patch16'
CLIP_HF_FILE = 'onnx/vision_model_quantized.onnx'

# OpenAI CLIP image preprocessing constants. Hard-coded so we don't pull in
# the transformers library just for an image processor.
CLIP_IMG_SIZE = 224
_CLIP_MEAN = (0.48145466, 0.4578275, 0.40821073)
_CLIP_STD  = (0.26862954, 0.26130258, 0.27577711)


# ---------------------------------------------------------------------------
# Lazy-loaded singleton ONNX inference session.
# ---------------------------------------------------------------------------

_clip_lock = threading.Lock()
_clip_session = None
_clip_input_name = None
_clip_mean_arr = None
_clip_std_arr = None
_clip_init_failed = False


def _get_clip():
    """Return the ONNX InferenceSession, loading on first call.

    Returns None if loading fails (network, missing deps, corrupt file).
    Callers must handle this case by returning 'skipped' decisions.
    """
    global _clip_session, _clip_input_name, _clip_mean_arr, _clip_std_arr, _clip_init_failed
    if _clip_session is not None:
        return _clip_session
    if _clip_init_failed:
        return None
    with _clip_lock:
        if _clip_session is not None:
            return _clip_session
        if _clip_init_failed:
            return None
        try:
            import onnxruntime as ort
            import numpy as np
            from huggingface_hub import hf_hub_download
            # Cached after first call (lives in ~/.cache/huggingface/hub/).
            path = hf_hub_download(repo_id=CLIP_HF_REPO, filename=CLIP_HF_FILE)
            # Single-threaded inference — Streamlit Cloud's CPU container
            # benefits from less context-switching when running 100+
            # sequential embeds.
            so = ort.SessionOptions()
            so.intra_op_num_threads = 1
            so.inter_op_num_threads = 1
            session = ort.InferenceSession(path,
                                            providers=['CPUExecutionProvider'],
                                            sess_options=so)
            _clip_session = session
            _clip_input_name = session.get_inputs()[0].name  # 'pixel_values'
            # Reshape mean/std to (1, 3, 1, 1) so we can broadcast against
            # the (1, 3, H, W) input tensor.
            _clip_mean_arr = np.array(_CLIP_MEAN, dtype=np.float32).reshape(1, 3, 1, 1)
            _clip_std_arr  = np.array(_CLIP_STD,  dtype=np.float32).reshape(1, 3, 1, 1)
        except Exception:
            _clip_init_failed = True
            return None
    return _clip_session


def unload_clip() -> None:
    """Free the ONNX session memory. Safe to call multiple times.

    Useful for an explicit RAM cleanup after a batch — the session itself
    is only ~85 MB resident but Streamlit Cloud's 1 GB ceiling adds up
    quickly when we're also holding ~200 cited-mark images and a
    docx-in-progress in the same process.
    """
    global _clip_session, _clip_input_name, _clip_mean_arr, _clip_std_arr
    with _clip_lock:
        _clip_session = None
        _clip_input_name = None
        _clip_mean_arr = None
        _clip_std_arr = None
    gc.collect()


# ---------------------------------------------------------------------------
# pHash — perceptual hash, sub-millisecond per image.
# ---------------------------------------------------------------------------

def compute_phash(image_bytes: bytes):
    """Compute the perceptual hash of an image. Returns an
    `imagehash.ImageHash` instance or None on failure."""
    if not image_bytes:
        return None
    try:
        import imagehash
        from PIL import Image
        img = Image.open(BytesIO(image_bytes)).convert('RGB')
        return imagehash.phash(img)
    except Exception:
        return None


def phash_distance(h1, h2) -> int:
    """Hamming distance between two perceptual hashes (0-64). Returns -1
    if either hash is None (couldn't be computed)."""
    if h1 is None or h2 is None:
        return -1
    try:
        return int(h1 - h2)
    except Exception:
        return -1


# ---------------------------------------------------------------------------
# Image preprocessing — pure numpy + Pillow, no torch / transformers.
# Implements the CLIP recipe: resize to 224x224 bicubic, scale to [0,1],
# normalise with CLIP_MEAN / CLIP_STD, transpose to (1, 3, H, W).
# ---------------------------------------------------------------------------

def _preprocess_for_clip(image_bytes: bytes):
    """Return a (1, 3, 224, 224) float32 numpy array ready for the ONNX
    session. None on failure."""
    import numpy as np
    from PIL import Image
    try:
        img = Image.open(BytesIO(image_bytes)).convert('RGB')
        img = img.resize((CLIP_IMG_SIZE, CLIP_IMG_SIZE), Image.BICUBIC)
        arr = np.asarray(img, dtype=np.float32) / 255.0    # (H, W, 3)
        arr = arr.transpose(2, 0, 1)[None, :, :, :]        # (1, 3, H, W)
        arr = (arr - _clip_mean_arr) / _clip_std_arr
        return arr.astype(np.float32)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# CLIP — visual embedding + cosine similarity.
# ---------------------------------------------------------------------------

def clip_embed(image_bytes: bytes):
    """Encode an image into a 512-dim CLIP visual embedding.

    Returns a unit-norm numpy array, or None on failure. Cosine
    similarity between two embeddings is just `np.dot(v1, v2)`.
    """
    if not image_bytes:
        return None
    session = _get_clip()
    if session is None:
        return None
    try:
        import numpy as np
        arr = _preprocess_for_clip(image_bytes)
        if arr is None:
            return None
        out = session.run(None, {_clip_input_name: arr})
        feat = out[0][0]                          # (512,)
        norm = float(np.linalg.norm(feat))
        if norm <= 0:
            return None
        return feat / norm                         # unit-norm
    except Exception:
        return None


def cosine_similarity(v1, v2) -> float:
    """Cosine similarity between two unit-norm vectors. Returns -1.0 if
    either is None."""
    if v1 is None or v2 is None:
        return -1.0
    try:
        import numpy as np
        return float(np.dot(v1, v2))
    except Exception:
        return -1.0


# ---------------------------------------------------------------------------
# Batch entry point — used by process_trademarks.
# Public API matches the torch version, so no caller changes needed.
# ---------------------------------------------------------------------------

def score_visual_similarity_batch(
    client_image_bytes: bytes,
    cited_images: dict,
    *,
    auto_unload: bool = True,
) -> dict:
    """Compute visual-similarity decisions for a batch of cited mark logos
    against a single client logo.

    See BR-IMG-002 module docstring (in git history) for the full API
    rationale. Returns dict mapping cited-mark identifier to:
        {
            'phash_distance': int,    # 0-64 or -1 on failure
            'clip_cosine':    float,  # -1.0 if CLIP wasn't run or failed
            'decision':       str,    # 'identical' | 'similar' | 'weak' |
                                      # 'unrelated' | 'skipped'
        }
    """
    out: dict = {}
    if not client_image_bytes or not cited_images:
        return out

    # Stage 1 — pHash both sides
    client_phash = compute_phash(client_image_bytes)
    if client_phash is None:
        return out
    phash_by_id = {cid: compute_phash(blob) for cid, blob in cited_images.items()}

    # Stage 2 — bucket + identify which need CLIP
    needs_clip: list = []
    for cid, ph in phash_by_id.items():
        if ph is None:
            out[cid] = {'phash_distance': -1, 'clip_cosine': -1.0,
                         'decision': 'skipped'}
            continue
        d = phash_distance(client_phash, ph)
        if d <= PHASH_DISTANCE_IDENTICAL:
            out[cid] = {'phash_distance': d, 'clip_cosine': -1.0,
                         'decision': 'identical'}
        elif d > PHASH_DISTANCE_UNRELATED:
            out[cid] = {'phash_distance': d, 'clip_cosine': -1.0,
                         'decision': 'unrelated'}
        else:
            needs_clip.append(cid)
            out[cid] = {'phash_distance': d, 'clip_cosine': -1.0,
                         'decision': 'skipped'}  # placeholder

    # Stage 3 — CLIP for the ambiguous bucket
    if needs_clip:
        client_vec = clip_embed(client_image_bytes)
        if client_vec is None:
            # CLIP unavailable — leave 'skipped' for everything ambiguous.
            # Caller falls back to base score for these rows.
            pass
        else:
            for cid in needs_clip:
                blob = cited_images.get(cid)
                cited_vec = clip_embed(blob)
                cos = cosine_similarity(client_vec, cited_vec)
                out[cid]['clip_cosine'] = round(cos, 4)
                if cos >= CLIP_COSINE_IDENTICAL:
                    out[cid]['decision'] = 'identical'
                elif cos >= CLIP_COSINE_SIMILAR:
                    out[cid]['decision'] = 'similar'
                elif cos >= CLIP_COSINE_WEAK:
                    out[cid]['decision'] = 'weak'
                else:
                    out[cid]['decision'] = 'unrelated'

    # Stage 4 — release model weights
    if auto_unload:
        unload_clip()

    return out


# ---------------------------------------------------------------------------
# Convenience reporter for the report builder.
# ---------------------------------------------------------------------------

def model_available() -> bool:
    """True if onnxruntime + huggingface_hub + imagehash + PIL are all
    importable AND the ONNX model can be downloaded + loaded. Useful for
    the report builder to know whether to surface 'visual sim n/a' vs a
    real cosine score."""
    return _get_clip() is not None
