"""Step 6 — Forensic verification layer.

Two distinct deliverables share this module:

  * Pre-Application forensic report  — likelihood of objection / similarity & risk
  * Post-Registration forensic report — potential infringements / reasons to act

Both consume the same verified-trademark dataset. The DIFFERENCE between the two
report types is in the narrative templates (handled in a separate module), not in
the data layer. This file is just the data layer: it fetches verbatim trademark
records from Signa.so (the unified trademark API) and normalises them into a
consistent shape the renderer can consume.

The Signa API is documented at https://docs.signa.so/ and is reached at
https://api.signa.so/v1/ . Bearer-token auth. JSON responses. Search by query
term + office filter, then narrow to a single application number.
"""
from __future__ import annotations
import time
from dataclasses import dataclass, asdict
from typing import Iterable
import requests


SIGNA_BASE_URL = 'https://api.signa.so/v1'
# Be polite — Signa probably allows a few requests per second, but a small
# inter-request delay protects us from accidental hammering on a large audit.
DEFAULT_RATE_LIMIT_SEC = 0.25
# Signa request timeout. Most lookups take under a second; 15s is generous.
DEFAULT_TIMEOUT_SEC = 15
# Maximum retries on transient failures (network blips, 5xx, rate limits).
DEFAULT_MAX_RETRIES = 3


@dataclass
class VerifiedRecord:
    """Normalised trademark record returned from Signa, in a shape the report
    builder can consume without caring what office it came from.

    Field mapping confirmed against Signa's live response shape (see
    documentation at https://docs.signa.so/). Some fields (source_url) are
    derived rather than returned directly.
    """
    office: str                 # Uppercased \u2014 'USPTO', 'UKIPO', 'EUIPO', 'WIPO'
    jurisdiction: str           # e.g. 'US', 'GB', 'EU'
    app_number: str             # Application number (string, may contain leading zeros)
    registration_number: str    # Often distinct, may be empty for pending apps
    mark_text: str              # Verbatim mark text
    mark_type: str              # 'word' / 'figurative' / 'combined' / etc.
    mark_legal_category: str    # 'standard' / 'stylised' / etc.
    status: str                 # 'active' / 'pending' / 'dead' / ...
    status_stage: str           # 'registered' / 'examination' / 'opposition' / ...
    filing_date: str            # ISO 'YYYY-MM-DD'
    registration_date: str      # ISO, may be ''
    expiry_date: str            # ISO, may be ''
    renewal_due_date: str       # ISO, may be ''
    owner_name: str
    nice_classes: list[int]     # All Nice classes the mark covers
    goods_services: list[str]   # 'Class N \u2014 verbatim recital' per class
    primary_image_url: str      # Signa-hosted image URL (figurative marks)
    has_media: bool             # True if Signa has an image for this record
    source_url: str             # Direct link to the source register record (derived)
    signa_id: str               # Signa's internal id, useful for follow-up calls
    verified: bool              # True if Signa returned a record; False if not found
    verification_note: str      # Free-text note explaining any verification gap

    def to_dict(self) -> dict:
        return asdict(self)


class SignaError(RuntimeError):
    """Raised when the Signa API returns an error we can't recover from."""


class SignaClient:
    """Thin REST wrapper over the Signa trademark API.

    Built against the public Signa documentation at https://docs.signa.so/ .
    Uses `requests` rather than the official SDK to keep the deployment surface
    small and to give us full control over retry / rate-limit behaviour.
    """

    def __init__(self, api_key: str, *,
                 base_url: str = SIGNA_BASE_URL,
                 rate_limit_sec: float = DEFAULT_RATE_LIMIT_SEC,
                 timeout_sec: int = DEFAULT_TIMEOUT_SEC,
                 max_retries: int = DEFAULT_MAX_RETRIES):
        if not api_key:
            raise ValueError('SignaClient requires a non-empty api_key')
        self._api_key = api_key
        self._base_url = base_url.rstrip('/')
        self._rate_limit = rate_limit_sec
        self._timeout = timeout_sec
        self._max_retries = max_retries
        self._last_request_at = 0.0

    # -- internal helpers --

    def _headers(self) -> dict:
        return {
            'Authorization': f'Bearer {self._api_key}',
            'Accept': 'application/json',
            'User-Agent': 'BrauditAuditTool/1.0 (+forensic-layer)',
        }

    def _sleep_for_rate_limit(self) -> None:
        elapsed = time.time() - self._last_request_at
        if elapsed < self._rate_limit:
            time.sleep(self._rate_limit - elapsed)

    def _get(self, path: str, params: dict | None = None) -> dict:
        url = f'{self._base_url}{path}'
        last_err: Exception | None = None
        for attempt in range(1, self._max_retries + 1):
            self._sleep_for_rate_limit()
            try:
                resp = requests.get(url, headers=self._headers(),
                                    params=params or {}, timeout=self._timeout)
            except requests.RequestException as exc:
                last_err = exc
                time.sleep(min(2 ** attempt, 8))
                continue
            finally:
                self._last_request_at = time.time()

            if resp.status_code == 200:
                try:
                    return resp.json()
                except ValueError as exc:
                    raise SignaError(f'Signa returned non-JSON 200 response: {exc}') from exc
            if resp.status_code == 429:
                # Rate-limited. Honour Retry-After if present, else exponential backoff.
                retry_after = float(resp.headers.get('Retry-After', '') or 2 ** attempt)
                time.sleep(min(retry_after, 10))
                last_err = SignaError(f'Rate limited (429); attempt {attempt}')
                continue
            if 500 <= resp.status_code < 600:
                last_err = SignaError(f'Signa server error {resp.status_code}')
                time.sleep(min(2 ** attempt, 8))
                continue
            # 4xx other than 429 — bubble up with the body for diagnostics
            body = (resp.text or '')[:400]
            raise SignaError(f'Signa {resp.status_code}: {body}')
        # Out of retries
        raise SignaError(f'Signa request failed after {self._max_retries} attempts: {last_err}')

    # -- public methods --

    def search(self, *, q: str, offices: Iterable[str] | str | None = None,
               limit: int = 10) -> list[dict]:
        """Free-text + office-filtered search. Returns the raw 'data' array."""
        params = {'q': q, 'limit': limit}
        if offices:
            if isinstance(offices, str):
                params['offices'] = offices
            else:
                params['offices'] = ','.join(offices)
        body = self._get('/trademarks', params)
        # Signa standard envelope: { "data": [...], "object": "list", ... }
        return body.get('data', []) if isinstance(body, dict) else []

    def lookup_by_app_number(self, *, app_number: str, office: str) -> dict | None:
        """Find a single trademark by application number scoped to one office.

        Returns the raw Signa record dict, or None if no match.

        Signa returns `office_code` lowercase (e.g. 'uspto') so we normalise on
        both ends \u2014 send the office in whatever case the caller used, but compare
        case-insensitively when filtering responses.
        """
        candidates = self.search(q=app_number, offices=office, limit=5)
        target = str(app_number).strip()
        target_office = (office or '').strip().lower()
        # Pass 1: exact match on application_number AND office_code
        for rec in candidates:
            rec_office = str(rec.get('office_code') or '').lower()
            rec_app = str(rec.get('application_number') or '').strip()
            if rec_app == target and (not target_office or rec_office == target_office):
                return rec
        # Pass 2: relax the office requirement (some scraped data may have
        # office mismatches; still return if the app number is an exact hit)
        for rec in candidates:
            if str(rec.get('application_number') or '').strip() == target:
                return rec
        # Pass 3: try registration_number (some scraped sheets store reg
        # numbers in the app number column for registered marks)
        for rec in candidates:
            if str(rec.get('registration_number') or '').strip() == target:
                return rec
        # Otherwise no confident match
        return None


# ---------- response normalisation ----------

def _safe_str(v) -> str:
    if v is None:
        return ''
    return str(v).strip()


def _coerce_nice_classes(rec: dict) -> list[int]:
    out: list[int] = []
    # Signa documentation indicates a 'classifications' array with nice_class
    classifications = rec.get('classifications') or []
    if isinstance(classifications, list):
        for c in classifications:
            if not isinstance(c, dict):
                continue
            n = c.get('nice_class')
            if n is None:
                continue
            try:
                out.append(int(n))
            except (TypeError, ValueError):
                continue
    # Fallback: some offices return a flat 'nice_classes' or 'classes' list
    if not out:
        flat = rec.get('nice_classes') or rec.get('classes') or []
        for n in flat:
            try:
                out.append(int(n))
            except (TypeError, ValueError):
                continue
    return sorted(set(out))


def _coerce_goods_services(rec: dict) -> list[str]:
    out: list[str] = []
    classifications = rec.get('classifications') or []
    if isinstance(classifications, list):
        for c in classifications:
            if not isinstance(c, dict):
                continue
            t = _safe_str(c.get('goods_services_text') or c.get('goods_services'))
            if t:
                # Prefix with the class number so the report can show 'Class 11 — ...'
                cls = c.get('nice_class')
                if cls:
                    out.append(f'Class {cls} — {t}')
                else:
                    out.append(t)
    return out


def _source_url_for(office: str, app_number: str) -> str:
    """Construct a direct link to the source register record.

    Signa doesn't return this so we build it ourselves using each office's
    known URL pattern. Matches the office-aware tm_url() in report_builder.py.
    """
    if not app_number:
        return ''
    o = (office or '').upper().strip()
    n = str(app_number).strip()
    if o in ('US', 'USPTO'):
        return f'https://tsdr.uspto.gov/statusview/sn{n}'
    if o in ('UK', 'UKIPO', 'GB'):
        return f'https://trademarks.ipo.gov.uk/ipo-tmcase/page/Results/1/{n}'
    if o in ('EU', 'EUIPO', 'EM'):
        return f'https://euipo.europa.eu/eSearch/#details/trademarks/{n}'
    if o in ('WIPO', 'MADRID', 'IR', 'WO'):
        return f'https://www3.wipo.int/branddb/en/showData.jsp?ID={n}'
    return ''


def normalise_signa_record(rec: dict, *, fallback_office: str = '',
                           fallback_app_number: str = '') -> VerifiedRecord:
    """Map a raw Signa response into our internal VerifiedRecord shape.

    Field mappings confirmed against the live Signa response shape:
      office_code, jurisdiction_code, application_number, registration_number,
      mark_text, mark_feature_type, mark_legal_category, status.primary,
      status.stage, filing_date, registration_date, expiry_date,
      renewal_due_date, owner_name, classifications[], primary_image_url,
      has_media, id.
    """
    status = rec.get('status') or {}
    if not isinstance(status, dict):
        status = {'primary': str(status)}

    office_raw = _safe_str(rec.get('office_code') or fallback_office)
    # Signa returns office_code lowercased ('uspto'). Uppercase for UI/code paths.
    office = office_raw.upper()
    app_number = _safe_str(rec.get('application_number') or fallback_app_number)

    return VerifiedRecord(
        office=office,
        jurisdiction=_safe_str(rec.get('jurisdiction_code')),
        app_number=app_number,
        registration_number=_safe_str(rec.get('registration_number')),
        mark_text=_safe_str(rec.get('mark_text')),
        mark_type=_safe_str(rec.get('mark_feature_type')),
        mark_legal_category=_safe_str(rec.get('mark_legal_category')),
        status=_safe_str(status.get('primary')),
        status_stage=_safe_str(status.get('stage')),
        filing_date=_safe_str(rec.get('filing_date')),
        registration_date=_safe_str(rec.get('registration_date')),
        expiry_date=_safe_str(rec.get('expiry_date')),
        renewal_due_date=_safe_str(rec.get('renewal_due_date')),
        owner_name=_safe_str(rec.get('owner_name')),
        nice_classes=_coerce_nice_classes(rec),
        goods_services=_coerce_goods_services(rec),
        primary_image_url=_safe_str(rec.get('primary_image_url')),
        has_media=bool(rec.get('has_media')),
        source_url=_source_url_for(office, app_number),
        signa_id=_safe_str(rec.get('id')),
        verified=True,
        verification_note='Verified via Signa API',
    )


def unverified_record(office: str, app_number: str, *, reason: str = '') -> VerifiedRecord:
    """Placeholder VerifiedRecord for trademarks we couldn't verify (no match,
    or transient API error). The renderer flags these explicitly so the
    operator can fact-check manually."""
    return VerifiedRecord(
        office=(office or '').upper(),
        jurisdiction='',
        app_number=app_number,
        registration_number='',
        mark_text='', mark_type='', mark_legal_category='',
        status='', status_stage='',
        filing_date='', registration_date='',
        expiry_date='', renewal_due_date='',
        owner_name='',
        nice_classes=[], goods_services=[],
        primary_image_url='', has_media=False,
        source_url=_source_url_for(office, app_number),
        signa_id='',
        verified=False,
        verification_note=reason or 'No matching record returned by Signa.',
    )


def verify_records(client: SignaClient, records: list[dict],
                   progress_callback=None) -> list[VerifiedRecord]:
    """Verify a list of (office, app_number) pairs against Signa.

    `records` is a list of dicts that must each contain at least
    'office' and 'app' (matches the existing scored_trademarks shape from
    pipeline/filters.py).

    Returns a list of VerifiedRecord — one per input record, in the same order.
    Failed lookups are returned as unverified placeholders rather than dropped,
    so the report renderer can flag them explicitly.
    """
    out: list[VerifiedRecord] = []
    total = len(records)
    for idx, t in enumerate(records, start=1):
        office = (t.get('office') or '').strip()
        app_number = (t.get('app') or t.get('app_number') or '').strip()
        if progress_callback:
            progress_callback(idx, total, office, app_number)
        if not office or not app_number:
            out.append(unverified_record(office, app_number,
                                         reason='Missing office or app number'))
            continue
        try:
            raw = client.lookup_by_app_number(app_number=app_number, office=office)
        except SignaError as exc:
            out.append(unverified_record(office, app_number,
                                         reason=f'Signa API error: {exc}'))
            continue
        if raw is None:
            out.append(unverified_record(office, app_number,
                                         reason='No matching record returned by Signa.'))
            continue
        out.append(normalise_signa_record(raw, fallback_office=office,
                                          fallback_app_number=app_number))
    return out
