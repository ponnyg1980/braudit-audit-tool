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

# --- BR-001 fix (28 May 2026) ---
# Signa's /v1/trademarks endpoint exposes two parallel filter parameters:
#   offices=         — accepts Signa's internal office codes (uspto, euipo,
#                      cipo, wipo, inpi-fr, ipau, ipi, ipos, nipo, prv).
#                      UKIPO not yet in production (planned code: ukipo).
#   jurisdictions=   — accepts ISO 3166-1 alpha-2 codes plus EU and WO,
#                      e.g. jurisdictions=US, jurisdictions=EU.
#
# Braudit's scraped spreadsheets carry ISO-style codes (US, GB, EU and the
# legacy OHIM-era EM for EUIPO), not Signa office codes. The previous code
# was sending these as offices=US / offices=EM / offices=GB and getting an
# HTTP 400 ("Unknown office code") every time. The fix is to (a) normalise
# legacy aliases to the canonical jurisdiction code used elsewhere in the
# tool (see pipeline/jurisdictions.py) and (b) send under jurisdictions=,
# not offices=.
#
# Side effect worth noting: jurisdictions=US returns USPTO direct nationals
# AND Madrid designations into the US via WIPO. For a clearance audit this
# is the correct behaviour — Madrid-routed rights are enforceable in the
# territory and matter for opposition risk.
#
# Note: until Signa ships the planned `ukipo` connector, jurisdictions=GB
# will only return Madrid designations into the UK (via WIPO) plus EUIPO
# pre-Brexit rights that designated GB. Native UK filings (UK00003xxxxxxx)
# remain unverifiable through Signa. See BR-002 in the build backlog.
_NORMALISE_TO_JURISDICTION = {
    # US
    'US': 'US', 'USA': 'US', 'USPTO': 'US',
    # EUIPO (EM is the legacy OHIM-era designator, 1996–2016)
    'EU': 'EU', 'EM': 'EU', 'EUTM': 'EU', 'EUIPO': 'EU',
    # UK
    'GB': 'GB', 'UK': 'GB', 'UKIPO': 'GB',
    # WIPO Madrid
    'WO': 'WO', 'WIPO': 'WO', 'IR': 'WO', 'MADRID': 'WO',
    # Benelux
    'BX': 'BX', 'BOIP': 'BX',
}


def _to_jurisdiction(code: str) -> str:
    """Normalise an office/jurisdiction string to the canonical code accepted
    by Signa's `jurisdictions=` filter. Returns the input uppercased if no
    mapping is found (lets unknown ISO codes pass through unchanged).
    """
    if not code:
        return ''
    key = str(code).strip().upper()
    return _NORMALISE_TO_JURISDICTION.get(key, key)


def _brexit_clone_to_eutm(uk_number: str) -> str | None:
    """Map a UK Brexit-clone application number back to its parent EUTM number.

    At midnight on 31 December 2020, EUIPO registrations covering the UK were
    cloned into the UK register under the `UK009` prefix. The clone number is
    formed by `UK009` + the parent EUTM's digits with the leading zero of
    the 9-digit EUTM stripped. Examples:
        EUTM 015873649  ->  UK00915873649
        EUTM 018452071  ->  UK00918452071
    Strip the `UK009` prefix and re-pad to 9 digits to recover the parent.

    Returns None for native UK numbers (UK00003xxxxxxxx and friends) and any
    input that does not match the Brexit-clone shape.
    """
    n = str(uk_number or '').strip().upper()
    if not n.startswith('UK009'):
        return None
    suffix = n[5:]
    if not suffix or not suffix.isdigit():
        return None
    # The EUTM register uses 9-digit numbers with a leading zero; the clone
    # encodes 8 digits. Anything other than 8 digits is not a recognisable
    # clone shape.
    if len(suffix) != 8:
        return None
    return suffix.zfill(9)


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
    verified: bool              # True if a registry returned a record; False if not
    verification_note: str      # Free-text note explaining any verification gap
    # Where the data actually came from. Added in the BR-002 Temmy patch.
    # One of: 'signa' (default, direct hit) | 'signa_brexit_clone_proxy'
    # (parent EUTM standing in for a UK clone) | 'temmy' (UKIPO via TemmyDB)
    # | 'unverified' (no source returned a record). Default 'signa' is kept
    # so any existing caller code that constructs VerifiedRecord without
    # this field keeps working.
    verification_source: str = 'signa'

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
        """Free-text + jurisdiction-filtered search. Returns the raw 'data' array.

        The `offices` kwarg name is kept for backwards-compatibility with
        existing callers, but the values are normalised through
        `_to_jurisdiction()` and sent under Signa's `jurisdictions=` filter.
        See module docstring (BR-001 fix note) for the rationale.
        """
        params = {'q': q, 'limit': limit}
        if offices:
            if isinstance(offices, str):
                codes = [_to_jurisdiction(offices)]
            else:
                codes = [_to_jurisdiction(o) for o in offices]
            codes = [c for c in codes if c]  # drop empties from bad input
            if codes:
                params['jurisdictions'] = ','.join(codes)
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
        # Pass 4 (Brexit clone fallback): if this is a UK Brexit-clone
        # application number (UK009xxxxxxxx), the clone itself is not in
        # Signa — UKIPO is on Signa's roadmap but not yet in production.
        # The clone shares its recital, owner, classes and status with the
        # parent EUTM (the two are the same underlying right, cloned at
        # midnight on 31 Dec 2020). Fetch the parent EUTM and re-badge it
        # as the UK record so the report renderer can show real data for
        # the UK row instead of an unverified placeholder.
        parent_eutm = _brexit_clone_to_eutm(app_number)
        if parent_eutm and target_office in ('gb', 'uk', 'ukipo', ''):
            eu_candidates = self.search(q=parent_eutm, offices='EU', limit=5)
            for rec in eu_candidates:
                rec_app = str(rec.get('application_number') or '').strip()
                if rec_app == parent_eutm:
                    proxy = dict(rec)
                    # Re-badge the proxy: caller asked about the UK clone,
                    # not the parent EUTM. The marker field lets the
                    # normaliser write an honest verification_note.
                    proxy['office_code'] = 'ukipo'
                    proxy['jurisdiction_code'] = 'GB'
                    proxy['application_number'] = app_number
                    proxy['_brexit_clone_parent_eutm'] = parent_eutm
                    return proxy
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

    # Brexit-clone proxy (see lookup_by_app_number Pass 4): the record carries
    # parent EUTM data but has been re-badged as UKIPO. Write an honest
    # verification note so the report renderer can show the provenance.
    # Also detect Temmy-sourced UK records via the `_temmy_proxy` marker the
    # temmy adapter writes — these are first-class UKIPO verifications.
    clone_parent = _safe_str(rec.get('_brexit_clone_parent_eutm'))
    temmy_proxy = bool(rec.get('_temmy_proxy'))
    if temmy_proxy:
        temmy_id = rec.get('_temmy_id') or ''
        verification_note = (
            f'Verified via TemmyDB (the Trademark Helpline’s internal UKIPO '
            f'mirror; refreshed weekly via live FTP feed from UKIPO). '
            f'Temmy record id {temmy_id}.'
        )
        verification_source = 'temmy'
    elif clone_parent:
        verification_note = (
            f'Verified via parent EUTM {clone_parent} — this UK record is a '
            f'Brexit clone created from the parent EUTM at 23:00 GMT on '
            f'31 December 2020. The record was not found in TemmyDB (UKIPO '
            f'mirror), so the recital, owner, classes and status shown here '
            f'are taken from the parent EUTM that the clone was derived from. '
            f'Confirm direct on IPSUM (https://trademarks.ipo.gov.uk/) before '
            f'relying for filing.'
        )
        verification_source = 'signa_brexit_clone_proxy'
    else:
        verification_note = 'Verified via Signa API'
        verification_source = 'signa'

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
        verification_note=verification_note,
        verification_source=verification_source,
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
        verification_note=reason or 'No matching record returned by any registry.',
        verification_source='unverified',
    )


def _is_uk_office(office: str) -> bool:
    """True if the input office code resolves to the UK jurisdiction."""
    return _to_jurisdiction(office) == 'GB'


def verify_records(client: SignaClient, records: list[dict],
                   progress_callback=None, *,
                   temmy_client=None) -> list[VerifiedRecord]:
    """Verify a list of (office, app_number) pairs.

    Dispatch:
      * UK records → TemmyDB primary (if `temmy_client` is provided). If
        TemmyDB returns no record AND the number is a Brexit-clone shape
        (UK009xxxxxxxx), fall back to the Signa parent-EUTM proxy. If
        both fail, return an unverified placeholder routed to the v3
        Data Quality block.
      * All other jurisdictions → Signa, unchanged.

    `records` is a list of dicts that must each contain at least 'office'
    and 'app' (matches the existing scored_trademarks shape from
    pipeline/filters.py).

    Returns a list of VerifiedRecord — one per input record, in the same
    order. Failed lookups are returned as unverified placeholders rather
    than dropped, so the report renderer can flag them explicitly.

    `temmy_client` is the TemmyClient instance (see temmy.py). If omitted
    or None, UK records skip the Temmy primary path and go straight to the
    Signa Brexit-clone fallback — useful for environments that don't have
    the Temmy credential configured.
    """
    # Local import to avoid a hard dependency on temmy.py when callers
    # don't configure a TemmyClient. Tries the in-package relative import
    # first (production layout where forensic.py and temmy.py both live
    # inside pipeline/), falls back to a top-level import (flat workspace
    # layout used for ad-hoc local testing).
    if temmy_client is not None:
        try:
            from .temmy import verify_uk_record_via_temmy, TemmyError
        except ImportError:
            from temmy import verify_uk_record_via_temmy, TemmyError
    else:
        verify_uk_record_via_temmy = None
        TemmyError = None  # type: ignore

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

        # ---- UK records: Temmy primary, Signa Brexit-clone backstop ----
        if _is_uk_office(office):
            temmy_raw = None
            if temmy_client is not None and verify_uk_record_via_temmy is not None:
                try:
                    temmy_raw = verify_uk_record_via_temmy(temmy_client, app_number)
                except Exception as exc:  # TemmyError or transient network
                    # Don't blow up the whole batch on a Temmy outage — fall
                    # through to the Signa Brexit-clone fallback. Capture the
                    # reason in case the fallback also fails.
                    temmy_raw = None
                    _last_temmy_error: str | None = f'TemmyDB lookup failed: {exc}'
                else:
                    _last_temmy_error = None
            else:
                _last_temmy_error = None

            if temmy_raw is not None:
                out.append(normalise_signa_record(
                    temmy_raw, fallback_office='UKIPO',
                    fallback_app_number=app_number))
                continue

            # Try the Signa Brexit-clone fallback for UK009xxxxxxxx numbers.
            # lookup_by_app_number internally runs Pass 4 (Brexit clone) when
            # the office is UK/GB/UKIPO and the number matches the clone shape.
            try:
                raw = client.lookup_by_app_number(
                    app_number=app_number, office=office)
            except SignaError as exc:
                reason = (
                    f'TemmyDB returned no record; Signa Brexit-clone fallback '
                    f'failed: {exc}.'
                    if _last_temmy_error is None
                    else f'{_last_temmy_error}; Signa Brexit-clone fallback '
                         f'also failed: {exc}.'
                )
                out.append(unverified_record(office, app_number,
                                             reason=reason))
                continue
            if raw is None:
                reason = (
                    'No matching record returned by TemmyDB or Signa '
                    '(native UK records pre-2018 may not be in TemmyDB; '
                    'Brexit-clone fallback only applies to UK009xxxxxxxx '
                    'numbers).'
                    if _last_temmy_error is None
                    else f'{_last_temmy_error}; Signa Brexit-clone fallback '
                         f'returned no record either.'
                )
                out.append(unverified_record(office, app_number,
                                             reason=reason))
                continue
            out.append(normalise_signa_record(raw, fallback_office=office,
                                              fallback_app_number=app_number))
            continue

        # ---- All other jurisdictions: existing Signa path ----
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
