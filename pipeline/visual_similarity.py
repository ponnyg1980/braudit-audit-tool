"""Visual similarity between the client logo and each cited mark logo.

BR-IMG-002 (Phase 2 of the image-similarity workstream, 10 Jun 2026).

Closes the loop on the Logo-axis scoring rubric: without this module the
initial audit had no visual signal at all — every Registered + Figurative
+ class-overlap mark scored High Risk on the Logo axis just on those
three metadata facts, even when the actual logo looked nothing like the
client's. We cap at Medium for those records elsewhere; this module
selectively *un-caps* (or further *demotes*) based on real visual
similarity to the client logo.

Two-stage pipeline
------------------
1. **pHash pre-filter** — perceptual hash on the client logo and each
   cited logo. pHash is sub-millisecond per image, no model needed.
   Buckets by Hamming distance (out of 64 bits):
       distance ≤ 5  → essentially identical pixels   → 'identical'
       distance > 30 → visually unrelated             → 'unrelated' (skip CLIP)
       distance 5-30 → ambiguous, CLIP earns the run  → 'ambiguous'

2. **CLIP visual embedding** — for the ambiguous bucket only, encode each
   image with OpenCLIP ViT-B/16 (LAION-2B weights) and take cosine
   similarity. Thresholds:
       cosine ≥ 0.85 → strong visual match → 'identical'
       0.65 ≤ cosine < 0.85 → mild similarity → 'similar'
       0.50 ≤ cosine < 0.65 → weak similarity → 'weak'
       cosine < 0.50 → clearly different → 'unrelated'

The decision labels feed into process_trademarks where they drive the
risk-grade adjustments per Jonathan's rule (10 Jun 2026):
    'identical' + Live + class overlap → High Risk
    'identical' + Live + no class overlap → Medium Risk
    'identical' + Ended (Dead) → Medium Risk (upgrades from Negligible)
    'similar' → keep base (Medium-capped) score
    'weak'    → demote to Low Risk
    'unrelated' → keep base, no adjustment

Lazy load / explicit unload
---------------------------
The OpenCLIP model + weights weighs ~340 MB resident. We load it on the
first `clip_embed` call and provide an `unload_clip()` for callers to
explicitly free memory after a batch completes. On Streamlit Cloud's
~1 GB container this keeps idle RAM low and only peaks during the visual-
similarity phase of an audit.

Failure-tolerant
----------------
Any error in the model load, image decode, or embedding returns
`{'decision': 'skipped', ...}` for that pair so the caller falls back
to the base Logo-axis score and the audit continues.
"""
from __future__ import annotations

import gc
import threading
from io import BytesIO
from typing import Optional


# ---------------------------------------------------------------------------
# Thresholds — Jonathan's calibration knobs. Surface them as module
# constants so the operator can tune without code-spelunking.
# ---------------------------------------------------------------------------

PHASH_DISTANCE_IDENTICAL = 5      # ≤ this → 'identical' bucket (no CLIP needed)
PHASH_DISTANCE_UNRELATED = 30     # > this → 'unrelated' bucket (no CLIP run)

CLIP_COSINE_IDENTICAL = 0.85      # ≥ this → 'identical'
CLIP_COSINE_SIMILAR   = 0.65      # 0.65–0.85 → 'similar' (keep base score)
CLIP_COSINE_WEAK      = 0.50      # 0.50–0.65 → 'weak' (demote to Low)
# below CLIP_COSINE_WEAK → 'unrelated'

# OpenCLIP model + LAION-2B weights. Best size/quality tradeoff for CPU
# inference. ViT-B/16 with LAION-2B weights beats OpenAI's CLIP ViT-B/32
# on retrieval benchmarks while running at similar latency.
CLIP_MODEL_NAME = 'ViT-B-16'
CLIP_PRETRAINED = 'laion2b_s34b_b88k'


# ---------------------------------------------------------------------------
# Lazy-loaded CLIP model + preprocess transform.
# ---------------------------------------------------------------------------

_clip_lock = threading.Lock()
_clip_model = None
_clip_preprocess = None
_clip_init_failed = False


def _get_clip():
    """Return (model, preprocess) tuple, loading on first call.

    Returns (None, None) if loading fails (open_clip not installed,
    weights download failed, OOM). Callers must handle this case.
    """
    global _clip_model, _clip_preprocess, _clip_init_failed
    if _clip_model is not None:
        return _clip_model, _clip_preprocess
    if _clip_init_failed:
        return None, None
    with _clip_lock:
        if _clip_model is not None:
            return _clip_model, _clip_preprocess
        if _clip_init_failed:
            return None, None
        try:
            import torch
            import open_clip
            model, _, preprocess = open_clip.create_model_and_transforms(
                CLIP_MODEL_NAME, pretrained=CLIP_PRETRAINED,
            )
            model.eval()
            # No GPU on Streamlit Cloud; pin to CPU explicitly so we
            # don't even try to move tensors to a CUDA device.
            _ = torch  # quiet linter
            _clip_model = model
            _clip_preprocess = preprocess
        except Exception:
            _clip_init_failed = True
            return None, None
    return _clip_model, _clip_preprocess


def unload_clip() -> None:
    """Explicitly free the OpenCLIP model weights. Call this after the
    visual-similarity phase of an audit completes so the ~340 MB of
    resident memory is released back to the container.

    Safe to call multiple times or before any clip_embed call.
    """
    global _clip_model, _clip_preprocess
    with _clip_lock:
        _clip_model = None
        _clip_preprocess = None
    gc.collect()


# ---------------------------------------------------------------------------
# pHash — perceptual hash, sub-millisecond per image.
# ---------------------------------------------------------------------------

def compute_phash(image_bytes: bytes):
    """Compute the perceptual hash of an image.

    Returns an `imagehash.ImageHash` instance, or None on failure.
    Use `phash_distance` to compare two hashes; equivalent to Hamming
    distance on the underlying 64-bit hash.
    """
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
    """Hamming distance between two perceptual hashes (0-64).

    Returns -1 if either hash is None (couldn't be computed).
    """
    if h1 is None or h2 is None:
        return -1
    try:
        return int(h1 - h2)
    except Exception:
        return -1


# ---------------------------------------------------------------------------
# CLIP — visual embedding + cosine similarity.
# ---------------------------------------------------------------------------

def clip_embed(image_bytes: bytes):
    """Encode an image into a 512-dim CLIP visual embedding.

    Returns a normalised numpy array, or None on failure. The embedding
    is L2-normalised so cosine similarity is just a dot product.
    """
    if not image_bytes:
        return None
    model, preprocess = _get_clip()
    if model is None:
        return None
    try:
        import torch
        from PIL import Image
        img = Image.open(BytesIO(image_bytes)).convert('RGB')
        tensor = preprocess(img).unsqueeze(0)
        with torch.no_grad():
            feat = model.encode_image(tensor)
            # L2-normalise so cosine sim = dot product
            feat = feat / feat.norm(dim=-1, keepdim=True)
        return feat.squeeze().cpu().numpy()
    except Exception:
        return None


def cosine_similarity(v1, v2) -> float:
    """Cosine similarity between two normalised vectors (range -1..1).

    Returns -1.0 if either vector is None.
    """
    if v1 is None or v2 is None:
        return -1.0
    try:
        import numpy as np
        return float(np.dot(v1, v2))
    except Exception:
        return -1.0


# ---------------------------------------------------------------------------
# Batch entry point — used by process_trademarks.
# ---------------------------------------------------------------------------

def score_visual_similarity_batch(
    client_image_bytes: bytes,
    cited_images: dict,
    *,
    auto_unload: bool = True,
) -> dict:
    """Compute visual-similarity decisions for a batch of cited mark logos
    against a single client logo.

    Parameters
    ----------
    client_image_bytes : raw JPEG/PNG bytes of the client's logo (from
        Order Form B30). If empty/None the function returns an empty dict
        and the caller treats it as 'no visual signal available'.
    cited_images : dict mapping cited-mark identifier (typically the
        spreadsheet row number) to raw image bytes.
    auto_unload : if True (default), call `unload_clip()` after the
        batch completes so the resident memory peak is transient. Set
        False if the caller intends to reuse the loaded model immediately.

    Returns
    -------
    dict mapping the same cited-mark identifier to:
        {
            'phash_distance': int,    # 0-64 or -1 on failure
            'clip_cosine':    float,  # -1.0 if CLIP wasn't run or failed
            'decision': str,          # 'identical' | 'similar' | 'weak' |
                                      # 'unrelated' | 'skipped'
        }

    The 'decision' label is the input to process_trademarks' Logo-axis
    risk adjustments. 'skipped' covers any per-image failure (decode,
    embed, etc.) and tells the caller to fall back to the base score.
    """
    out: dict = {}
    if not client_image_bytes or not cited_images:
        return out

    # ---- Stage 1: pHash both sides --------------------------------------
    client_phash = compute_phash(client_image_bytes)
    if client_phash is None:
        # Can't even pHash the client logo — bail. Caller falls back to
        # base score for everything.
        return out

    phash_by_id = {}
    for cid, blob in cited_images.items():
        phash_by_id[cid] = compute_phash(blob)

    # ---- Stage 2: bucket + identify which need CLIP ---------------------
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
            # Ambiguous — run CLIP
            needs_clip.append(cid)
            out[cid] = {'phash_distance': d, 'clip_cosine': -1.0,
                         'decision': 'skipped'}  # placeholder

    # ---- Stage 3: CLIP for the ambiguous bucket -------------------------
    if needs_clip:
        client_vec = clip_embed(client_image_bytes)
        if client_vec is None:
            # CLIP unavailable — leave 'skipped' for everything in the
            # ambiguous bucket. Caller falls back to base score for these.
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

    # ---- Stage 4: release model weights ---------------------------------
    if auto_unload:
        unload_clip()

    return out


# ---------------------------------------------------------------------------
# Convenience reporters
# ---------------------------------------------------------------------------

def model_available() -> bool:
    """True if both pHash (imagehash + PIL) and CLIP (open_clip + torch)
    are importable and the model can be constructed. Useful for the
    report builder to know whether to surface 'visual sim n/a' vs an
    actual cosine score."""
    model, _ = _get_clip()
    return model is not None
