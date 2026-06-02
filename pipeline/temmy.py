"""TemmyDB adapter — UK trademark verification via the Trademark Helpline's
own UKIPO mirror.

TemmyDB is Braudit's internal source of truth for UK trademark data. It
holds the full UKIPO trademark register from 2018 onwards and is refreshed
weekly via a live FTP feed. For UK records, TemmyDB takes priority over
Signa (which does not yet have UKIPO in production — UKIPO is on Signa's
"Planned" list with code `ukipo`).

This module exposes two HTTP clients:

  * `TemmyClient` — standard X-API-Key surface. Used for record verification
    (`get_trademark`), prefix search (`search_trademarks`), applicant lookup
    (`search_applicants`), and the weekly update feed (`updated_since` /
    `count_updated_since`). This is what the forensic appendix uses to fill
    in UK rows in the scoring table.

  * `TemmyQueryRunsClient` — privileged X-Query-Runs-Key surface for
    bounded read-only SQL. Useful for ad-hoc analytics on the UK register
    (e.g. "all live registrations naming applicant X across class Y").
    NOT used by the forensic appendix.

The adapter function `verify_uk_record_via_temmy()` wraps the trademark
detail call and returns a raw dict shaped like the `VerifiedRecord`
fields so the orchestrator in `forensic.py` can normalise it through
the same path it uses for Signa records.

API reference: see `temmy-api-standard-key.md` and
`temmy-query-runs-privileged.md` in the External Developer Pack.
"""
from __future__ import annotations
import time
from typing import Iterable, Any
import requests


TEMMY_BASE_URL = 'https://temmy-api-prod-zfxujusd3q-nw.a.run.app'
# Defaults aligned with SignaClient — polite to upstream by default.
DEFAULT_RATE_LIMIT_SEC = 0.25
DEFAULT_TIMEOUT_SEC = 15
DEFAULT_MAX_RETRIES = 3


class TemmyError(RuntimeError):
    """Raised when the Temmy API returns an error we can't recover from."""


# ============================================================================
# Standard API key client — read-only trademark / applicant / update endpoints
# ============================================================================

class TemmyClient:
    """Thin REST wrapper over the Temmy HTTP API v1.

    Authenticates via `X-API-Key` (the standard external key). Use this
    for every UK trademark verification path. For bounded SQL analytics
    use `TemmyQueryRunsClient` separately — the two credentials are
    distinct on Temmy's side.
    """

    def __init__(self, api_key: str, *,
                 base_url: str = TEMMY_BASE_URL,
                 rate_limit_sec: float = DEFAULT_RATE_LIMIT_SEC,
                 timeout_sec: int = DEFAULT_TIMEOUT_SEC,
                 max_retries: int = DEFAULT_MAX_RETRIES):
        if not api_key:
            raise ValueError('TemmyClient requires a non-empty api_key')
        self._api_key = api_key
        self._base_url = base_url.rstrip('/')
        self._rate_limit = rate_limit_sec
        self._timeout = timeout_sec
        self._max_retries = max_retries
        self._last_request_at = 0.0

    # -- internal helpers --

    def _headers(self) -> dict:
        return {
            'X-API-Key': self._api_key,
            'Accept': 'application/json',
            'User-Agent': 'BrauditAuditTool/2.0 (+forensic-layer; temmy-adapter)',
        }

    def _sleep_for_rate_limit(self) -> None:
        elapsed = time.time() - self._last_request_at
        if elapsed < self._rate_limit:
            time.sleep(self._rate_limit - elapsed)

    def _request(self, method: str, path: str, *,
                 params: dict | None = None,
                 json: Any = None,
                 accept_404_as_none: bool = False) -> dict | None:
        """One HTTP call with rate-limit + retry behaviour matching SignaClient.

        Returns parsed JSON on 200. Returns None on 404 if
        `accept_404_as_none` is set (used for "not found" semantics on
        single-record lookups).
        """
        url = f'{self._base_url}{path}'
        last_err: Exception | None = None
        for attempt in range(1, self._max_retries + 1):
            self._sleep_for_rate_limit()
            try:
                resp = requests.request(method, url, headers=self._headers(),
                                        params=params or None, json=json,
                                        timeout=self._timeout)
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
                    raise TemmyError(f'Temmy returned non-JSON 200 response: {exc}') from exc
            if resp.status_code == 404 and accept_404_as_none:
                return None
            if resp.status_code == 429:
                retry_after = float(resp.headers.get('Retry-After', '') or 2 ** attempt)
                time.sleep(min(retry_after, 10))
                last_err = TemmyError(f'Rate limited (429); attempt {attempt}')
                continue
            if 500 <= resp.status_code < 600:
                last_err = TemmyError(f'Temmy server error {resp.status_code}')
                time.sleep(min(2 ** attempt, 8))
                continue
            # 4xx other than 429/404 — surface the body for diagnostics
            body = (resp.text or '')[:400]
            raise TemmyError(f'Temmy {resp.status_code}: {body}')
        raise TemmyError(f'Temmy request failed after {self._max_retries} attempts: {last_err}')

    # -- public methods (standard X-API-Key surface) --

    def health(self) -> dict:
        """Unauthenticated health check. Returns `{"status": "ok"}` when up."""
        # /health is public and accepts no headers — but our wrapper passes
        # the API key header regardless; the server tolerates it.
        return self._request('GET', '/health') or {}

    def get_trademark(self, trademark_id: str | int) -> dict | None:
        """Full trademark detail. Accepts either Temmy internal id (e.g. 248616)
        or a UK application number (e.g. 'UK00003553322'). Returns None on 404.
        """
        return self._request('GET', f'/api/v1/trademarks/{trademark_id}',
                             accept_404_as_none=True)

    def get_trademark_status(self, trademark_id: str | int) -> dict | None:
        """Compact status/date fields. Returns None on 404."""
        return self._request('GET', f'/api/v1/trademarks/{trademark_id}/status',
                             accept_404_as_none=True)

    def search_trademarks(self, *, text: str, limit: int = 25,
                          page: int = 1) -> dict:
        """Prefix search by verbal element text. Returns a paginated envelope:

            { "query": {...}, "pagination": {...}, "items": [...] }
        """
        params = {'text': text, 'limit': limit, 'page': page}
        body = self._request('GET', '/api/v1/trademarks/search', params=params)
        return body or {'items': [], 'pagination': {}}

    def search_applicants(self, *, name: str | None = None,
                          ipo_identifier: int | None = None,
                          exact_match: bool = False,
                          limit: int = 25, page: int = 1) -> dict:
        """Applicant lookup by name OR IPO identifier (exactly one). Returns
        a paginated envelope with applicants and their active trademarks.
        """
        if (name is None) == (ipo_identifier is None):
            raise ValueError('search_applicants requires exactly one of name or ipo_identifier')
        params: dict = {'limit': limit, 'page': page}
        if name is not None:
            params['name'] = name
            params['exact_match'] = 'true' if exact_match else 'false'
        else:
            params['id'] = ipo_identifier
        body = self._request('GET', '/api/v1/applicants/search', params=params)
        return body or {'items': [], 'pagination': {}}

    def updated_since(self, *, since: str, limit: int = 1000,
                      page: int = 1) -> dict:
        """Paged feed of trademarks where `last_updated_on >= since`.
        `since` is an ISO date or datetime. Use this for the weekly
        trademark-monitoring delta against watched brands.
        """
        params = {'since': since, 'limit': limit, 'page': page}
        body = self._request('GET', '/api/v1/trademarks/updated-since',
                             params=params)
        return body or {'items': [], 'pagination': {}}

    def count_updated_since(self, *, since: str) -> int:
        """Just the total count of trademarks updated since the given date."""
        body = self._request('GET', '/api/v1/trademarks/updated-since/count',
                             params={'since': since})
        return int((body or {}).get('total', 0))


# ============================================================================
# Privileged Query Runs client — bounded read-only SQL over TemmyDB
# ============================================================================

class TemmyQueryRunsClient:
    """Privileged wrapper over Temmy's /api/v2/query-runs endpoints.

    Uses `X-Query-Runs-Key`, NOT the standard `X-API-Key`. Used for
    bounded read-only SELECT analytics over the UK register. Not used
    by the forensic appendix; exposed for ad-hoc analytical workflows
    and any future Streamlit / monitoring tooling that needs SQL.
    """

    def __init__(self, query_runs_key: str, *,
                 base_url: str = TEMMY_BASE_URL,
                 timeout_sec: int = DEFAULT_TIMEOUT_SEC):
        if not query_runs_key:
            raise ValueError('TemmyQueryRunsClient requires a non-empty query_runs_key')
        self._key = query_runs_key
        self._base_url = base_url.rstrip('/')
        self._timeout = timeout_sec

    def _headers(self) -> dict:
        return {
            'X-Query-Runs-Key': self._key,
            'Accept': 'application/json',
            'User-Agent': 'BrauditAuditTool/2.0 (+forensic-layer; temmy-queryruns)',
        }

    def submit(self, *, sql: str, page_size: int = 1000,
               ttl_seconds: int = 600, preview_limit: int = 25) -> dict:
        """Submit one bounded SELECT/WITH statement. Returns a manifest with
        `query_id`, `columns`, `pagination`, `preview` and `expires_at`.

        See `temmy-query-runs-privileged.md` for full schema. Statements
        must start with SELECT or WITH; write/control keywords are rejected
        server-side.
        """
        payload = {
            'sql': sql,
            'page_size': page_size,
            'ttl_seconds': ttl_seconds,
            'preview_limit': preview_limit,
        }
        resp = requests.post(f'{self._base_url}/api/v2/query-runs',
                             headers=self._headers(), json=payload,
                             timeout=self._timeout)
        if resp.status_code != 200:
            raise TemmyError(f'Query Runs submit failed: {resp.status_code} {resp.text[:400]}')
        return resp.json()

    def get_manifest(self, query_id: str) -> dict:
        """Read the manifest for a previously-submitted query run."""
        resp = requests.get(f'{self._base_url}/api/v2/query-runs/{query_id}',
                            headers=self._headers(), timeout=self._timeout)
        if resp.status_code != 200:
            raise TemmyError(f'Query Runs manifest failed: {resp.status_code} {resp.text[:400]}')
        return resp.json()

    def get_page(self, query_id: str, page: int) -> dict:
        """Fetch a materialised result page (1-indexed)."""
        resp = requests.get(
            f'{self._base_url}/api/v2/query-runs/{query_id}/pages/{page}',
            headers=self._headers(), timeout=self._timeout)
        if resp.status_code != 200:
            raise TemmyError(f'Query Runs page fetch failed: {resp.status_code} {resp.text[:400]}')
        return resp.json()

    def delete(self, query_id: str) -> dict:
        """Delete temporary result objects for a query run."""
        resp = requests.delete(
            f'{self._base_url}/api/v2/query-runs/{query_id}',
            headers=self._headers(), timeout=self._timeout)
        if resp.status_code != 200:
            raise TemmyError(f'Query Runs delete failed: {resp.status_code} {resp.text[:400]}')
        return resp.json()


# ============================================================================
# Adapter — Temmy detail response → fields the orchestrator can normalise
# ============================================================================

def _safe_str(v) -> str:
    if v is None:
        return ''
    return str(v).strip()


def _coerce_uk_classes(rec: dict) -> list[int]:
    """UK trademarks expose `classes: [9, 25]` and may also expose a richer
    `nice_class_trademarks` array. Prefer the flat `classes` list when
    present and fall back to nice_class_trademarks otherwise.
    """
    out: list[int] = []
    flat = rec.get('classes') or []
    if isinstance(flat, list):
        for n in flat:
            try:
                out.append(int(n))
            except (TypeError, ValueError):
                continue
    if not out:
        nct = rec.get('nice_class_trademarks') or []
        if isinstance(nct, list):
            for c in nct:
                if not isinstance(c, dict):
                    continue
                n = c.get('nice_class') or c.get('class')
                if n is None:
                    continue
                try:
                    out.append(int(n))
                except (TypeError, ValueError):
                    continue
    return sorted(set(out))


def _coerce_uk_goods_services(rec: dict) -> list[str]:
    """Pull the verbatim recital per Nice class from `nice_class_trademarks`,
    defending against the several plausible shapes the nested object can take.
    Returns 'Class N — recital text' strings ready for the report renderer.
    """
    out: list[str] = []
    nct = rec.get('nice_class_trademarks') or []
    if not isinstance(nct, list):
        return out
    for c in nct:
        if not isinstance(c, dict):
            continue
        # Try a few plausible field names — Temmy doc states "additional
        # database fields may appear"; treat missing fields as empty.
        text = _safe_str(
            c.get('goods_services_text')
            or c.get('goods_services')
            or c.get('description')
            or c.get('text')
        )
        if not text:
            continue
        cls = c.get('nice_class') or c.get('class')
        if cls:
            out.append(f'Class {cls} — {text}')
        else:
            out.append(text)
    return out


def _coerce_uk_owner(rec: dict) -> str:
    """First applicant on the record. UK records can have multiple
    applicants; the report renderer expects a single owner_name string
    so we join multiples with semicolons.
    """
    applicants = rec.get('applicants') or []
    if not isinstance(applicants, list):
        return ''
    names = []
    for a in applicants:
        if not isinstance(a, dict):
            continue
        n = _safe_str(a.get('name'))
        if n:
            names.append(n)
    return '; '.join(names)


def _ukipo_source_url(app_number: str) -> str:
    n = _safe_str(app_number)
    if not n:
        return ''
    return f'https://trademarks.ipo.gov.uk/ipo-tmcase/page/Results/1/{n}'


def _filing_date_from(rec: dict) -> str:
    """Temmy returns `application_date_time` as an ISO datetime. The report
    expects an ISO date. Split off the date portion defensively.
    """
    raw = _safe_str(rec.get('application_date_time')
                    or rec.get('application_date'))
    if not raw:
        return ''
    # Common shapes: '2020-01-15T10:00:00', '2020-01-15T10:00:00+00:00',
    # '2020-01-15 10:00:00', '2020-01-15'.
    return raw.split('T', 1)[0].split(' ', 1)[0]


def verify_uk_record_via_temmy(client: TemmyClient, app_number: str
                               ) -> dict | None:
    """Look a UK trademark up in TemmyDB and return a dict shaped like the
    fields `normalise_signa_record()` expects, so the orchestrator in
    forensic.py can route Temmy results through the same normaliser.

    Returns None if the record is not in TemmyDB (404 from the API), or
    if the input is empty / non-UK shaped. Raises TemmyError on transient
    or authentication failures so the orchestrator can flag the record
    rather than silently treating it as "not found".

    The returned dict is decorated with `_temmy_proxy=True` so the
    normaliser can write a Temmy-specific verification_note rather than
    the Signa default.
    """
    if not app_number:
        return None

    raw = client.get_trademark(app_number)
    if raw is None:
        return None

    mark = raw.get('mark') or {}
    mark_text = (_safe_str(raw.get('verbal_element_text'))
                 or _safe_str(mark.get('verbal_element_text'))
                 or _safe_str(mark.get('description')))

    # Build a dict in the shape normalise_signa_record() reads.
    proxy = {
        'id': f'temmy_{raw.get("temmy_id") or ""}',
        'office_code': 'ukipo',
        'jurisdiction_code': 'GB',
        'application_number': _safe_str(raw.get('application_number')) or _safe_str(app_number),
        'registration_number': _safe_str(raw.get('application_number')),  # UKIPO uses one number throughout
        'mark_text': mark_text,
        'mark_feature_type': _safe_str(mark.get('feature')),
        'mark_legal_category': _safe_str(mark.get('kind_mark')),
        'status': {
            'primary': _safe_str(raw.get('status')),
            'stage': _safe_str(raw.get('status')),  # Temmy collapses these
        },
        'filing_date': _filing_date_from(raw),
        'registration_date': _safe_str(raw.get('registration_date')),
        'expiry_date': _safe_str(raw.get('expiry_date')),
        'renewal_due_date': _safe_str(raw.get('expiry_date')),  # UKIPO: renewal = expiry
        'owner_name': _coerce_uk_owner(raw),
        # These two keys are read by Signa's _coerce_* helpers in forensic.py:
        'classifications': [
            {'nice_class': n, 'goods_services_text': _extract_recital_for_class(raw, n)}
            for n in _coerce_uk_classes(raw)
        ],
        'primary_image_url': _safe_str(raw.get('image_url')),
        'has_media': bool(raw.get('image_url')),
        # Marker so the normaliser knows this came from Temmy not Signa:
        '_temmy_proxy': True,
        '_temmy_id': raw.get('temmy_id'),
    }
    return proxy


def _extract_recital_for_class(rec: dict, nice_class: int) -> str:
    """Pull the recital text for a single Nice class out of the
    nice_class_trademarks array, defending against the several plausible
    shapes the nested object can take. Returns empty string if no recital
    is recorded against that class.
    """
    nct = rec.get('nice_class_trademarks') or []
    if not isinstance(nct, list):
        return ''
    for c in nct:
        if not isinstance(c, dict):
            continue
        cls = c.get('nice_class') or c.get('class')
        try:
            if cls is None or int(cls) != int(nice_class):
                continue
        except (TypeError, ValueError):
            continue
        return _safe_str(
            c.get('goods_services_text')
            or c.get('goods_services')
            or c.get('description')
            or c.get('text')
        )
    return ''
