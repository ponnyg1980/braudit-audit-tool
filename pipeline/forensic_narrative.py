"""Day 2 — Forensic narrative generation via Anthropic Sonnet 4.6.

Two-pass design, two report types:

  Pass 1  Deterministic forensic scoring per VerifiedRecord (0–10 across four
          criteria). Pure Python, no LLM. Reproducible and auditable.

  Pass 2  LLM-generated per-record commentary in one batched Sonnet call.
          Hybrid mode: verbatim TSDR/Signa data is preserved, only the
          forensic commentary paragraph is generated.

  Pass 3  LLM-generated executive summary + final recommendation in one Sonnet
          call. Different system prompt depending on report type:

            PRE_APPLICATION   — "likelihood of objection / similarity & risk"
            POST_REGISTRATION — "potential infringements / reasons to act"

The actual docx rendering lives in pipeline/forensic_report.py (Day 3).
"""
from __future__ import annotations
import json
from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Iterable

# Anthropic SDK is imported lazily inside NarrativeClient so that the
# deterministic scoring functions in this module can be used in environments
# (tests, local dev) where the SDK isn't installed.
from .forensic import VerifiedRecord


# Model choice: Sonnet 4.6 is the sweet-spot for narrative quality + cost.
# Opus 4.6 produces marginally better prose at ~5x the price; we route to
# Opus only when an audit is explicitly escalated (Phase B Run Card flag).
DEFAULT_MODEL = 'claude-sonnet-4-6'
DEFAULT_MAX_TOKENS_PER_RECORD = 260
# Summary call must accommodate executive_summary (300-500w) +
# final_recommendation (250-400w) + recommended_specification or
# recommended_actions (variable). 4000 tokens covers the worst case
# comfortably and leaves room for JSON overhead and newlines.
DEFAULT_MAX_TOKENS_SUMMARY = 4000


def _robust_json_loads(raw: str) -> dict:
    """Parse JSON returned by the LLM, tolerating common quirks.

    Handles:
      * ```json / ``` code fences around the JSON
      * Leading or trailing prose around the JSON object
      * Truncated/malformed JSON \u2014 by extracting the longest
        prefix that parses cleanly via JSONDecoder.raw_decode().
    """
    cleaned = (raw or '').strip()
    if cleaned.startswith('```'):
        nl = cleaned.find('\n')
        if nl >= 0:
            cleaned = cleaned[nl + 1:]
        if cleaned.endswith('```'):
            cleaned = cleaned[:-3]
        cleaned = cleaned.strip()
    # Pass 1: parse as-is
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass
    # Pass 2: locate the first '{' and try raw_decode from there
    start = cleaned.find('{')
    if start < 0:
        raise json.JSONDecodeError('No JSON object found in LLM response', cleaned, 0)
    try:
        obj, _ = json.JSONDecoder().raw_decode(cleaned[start:])
        if isinstance(obj, dict):
            return obj
    except json.JSONDecodeError:
        pass
    raise json.JSONDecodeError(
        'Failed to parse LLM response as JSON (possibly truncated)',
        cleaned, 0,
    )


class ReportType(str, Enum):
    PRE_APPLICATION = 'pre_application'
    POST_REGISTRATION = 'post_registration'


# ---------- deterministic forensic scoring (0–10) ----------

@dataclass
class ForensicScore:
    """Reproducible 0–10 score with a breakdown by criterion, so the report
    can show the rubric alongside the headline number."""
    total: int                # 0–10
    risk_band: str            # 'Very High' / 'High' / 'Medium' / 'Low' / 'Negligible'
    classes_terms: int        # 0–3
    mark_type: int            # 0–2
    age_proof_of_use: int     # 0–3
    region: int               # 0–2
    explanation: list[str]    # Per-criterion human-readable lines

    def to_dict(self) -> dict:
        return asdict(self)


def _risk_band_for(score: int) -> str:
    if score >= 9: return 'Very High'
    if score >= 7: return 'High'
    if score >= 5: return 'Medium'
    if score >= 3: return 'Low'
    return 'Negligible'


def score_record(record: VerifiedRecord, *,
                 client_classes: list[int],
                 client_mark: str,
                 client_jurisdiction: str | list[str] = 'US') -> ForensicScore:
    """Apply the four-criteria forensic rubric (same one in the Run Card).

    Pure deterministic Python — no LLM, no randomness, no API calls. The
    LLM commentary is generated separately and references these scores.
    """
    notes: list[str] = []

    # 1. Classes and terms overlap (0–3)
    shared_classes = set(record.nice_classes) & set(client_classes or [])
    if len(shared_classes) >= 2:
        classes_terms = 3
        notes.append(f'High class overlap: shares {len(shared_classes)} classes with client filing ({", ".join(str(c) for c in sorted(shared_classes))})')
    elif len(shared_classes) == 1:
        cls = next(iter(shared_classes))
        # Bump higher if the senior recital text contains key client terms
        client_keywords = [w.upper() for w in (client_mark or '').split() if len(w) > 2]
        recital_blob = ' '.join(record.goods_services).upper()
        keyword_hits = sum(1 for k in client_keywords if k in recital_blob)
        if keyword_hits >= 1:
            classes_terms = 3
            notes.append(f'Same class (Class {cls}) AND recital references client keywords ({keyword_hits} match{"es" if keyword_hits>1 else ""})')
        else:
            classes_terms = 2
            notes.append(f'Same class (Class {cls}) but recital differs from client goods')
    else:
        classes_terms = 1 if record.nice_classes else 0
        notes.append('No shared classes with client filing' if record.nice_classes else 'No class data available')

    # 2. Mark type (0–2). Word marks have broadest scope.
    mt = (record.mark_type or '').lower()
    if 'word' in mt or 'standard' in mt:
        mark_type = 2
        notes.append('Word/standard-character mark — broadest scope')
    elif 'combined' in mt:
        mark_type = 1
        notes.append('Combined mark — moderate scope')
    elif 'figurative' in mt or 'stylised' in mt or 'stylized' in mt:
        mark_type = 1
        notes.append('Figurative / stylised — narrower scope')
    else:
        mark_type = 0
        notes.append(f'Mark type unknown or unusual ({record.mark_type or "blank"})')

    # 3. Age and proof of use (0–3)
    status_l = (record.status or '').lower() + ' ' + (record.status_stage or '').lower()
    if 'dead' in status_l or 'ended' in status_l or 'abandoned' in status_l or 'cancelled' in status_l:
        age_proof_of_use = 0
        notes.append('Record is dead / ended / cancelled — unenforceable rights')
    elif record.registration_date:
        # Live + registered + has a registration date → enforceable rights
        age_proof_of_use = 3
        notes.append('Live registered right with confirmed registration date')
    elif 'pending' in status_l or 'examination' in status_l or 'opposition' in status_l:
        age_proof_of_use = 2
        notes.append('Pending application — rights not yet vested but progressing')
    elif 'active' in status_l:
        age_proof_of_use = 2
        notes.append('Active record — verification of registration date recommended')
    else:
        age_proof_of_use = 1
        notes.append('Status uncertain — manual verification recommended')

    # 4. Region / jurisdiction (0–2). Now supports a LIST of client
    # jurisdictions: highest score if record matches ANY client jurisdiction.
    rec_juris = (record.jurisdiction or '').upper()
    if isinstance(client_jurisdiction, str):
        cli_juris_list = [client_jurisdiction.upper()] if client_jurisdiction else []
    else:
        cli_juris_list = [str(j).upper() for j in (client_jurisdiction or []) if j]
    cli_juris_display = ', '.join(cli_juris_list) if cli_juris_list else '(none)'

    if rec_juris and rec_juris in cli_juris_list:
        region = 2
        notes.append(f'Same jurisdiction as client filing ({rec_juris} \u2208 {{{cli_juris_display}}})')
    elif rec_juris and cli_juris_list and rec_juris in ('EU', 'WO', 'IR'):
        # Regional / international rights overlap most national filings
        region = 1
        notes.append(f'International/regional right ({rec_juris}) may extend to client jurisdictions ({cli_juris_display})')
    elif rec_juris and cli_juris_list and ('EU' in cli_juris_list and rec_juris in ('AT','BE','BG','HR','CY','CZ','DK','EE','FI','FR','DE','GR','HU','IE','IT','LV','LT','LU','MT','NL','PL','PT','RO','SK','SI','ES','SE')):
        # Client has EUIPO filing, record is from an EU member state
        region = 2
        notes.append(f'Client has EU coverage and record is from EU member state ({rec_juris})')
    elif not rec_juris:
        region = 1
        notes.append('Jurisdiction unknown')
    else:
        region = 0
        notes.append(f'Different jurisdiction ({rec_juris} vs client {cli_juris_display})')

    total = classes_terms + mark_type + age_proof_of_use + region
    return ForensicScore(
        total=total, risk_band=_risk_band_for(total),
        classes_terms=classes_terms, mark_type=mark_type,
        age_proof_of_use=age_proof_of_use, region=region,
        explanation=notes,
    )


# ---------- Anthropic client wrapper ----------

class NarrativeClient:
    """Thin wrapper over the Anthropic SDK, scoped to the narrative use-case."""

    def __init__(self, api_key: str, *, model: str = DEFAULT_MODEL):
        if not api_key:
            raise ValueError('NarrativeClient requires a non-empty api_key')
        # Lazy import so the rest of this module can be used without the SDK
        from anthropic import Anthropic
        self._client = Anthropic(api_key=api_key)
        self._model = model

    def _call(self, *, system: str, user: str, max_tokens: int) -> str:
        """Single Sonnet round-trip. Returns the text content of the response."""
        msg = self._client.messages.create(
            model=self._model,
            max_tokens=max_tokens,
            system=system,
            messages=[{'role': 'user', 'content': user}],
        )
        # SDK returns content as a list of blocks; we use 'text' blocks only
        text_chunks = []
        for block in msg.content:
            if getattr(block, 'type', None) == 'text':
                text_chunks.append(block.text)
        return ''.join(text_chunks).strip()


# ---------- Per-record commentary ----------

def _pre_application_system_prompt() -> str:
    return (
        "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist at "
        "Braudit. You are writing the per-record forensic commentary for a "
        "PRE-APPLICATION trademark clearance audit. The client is considering filing "
        "a new trademark and wants to know which existing senior rights might cause "
        "the USPTO (or other examining authority) to refuse the application on grounds "
        "of likelihood of confusion under section 2(d), or which owners might file an "
        "opposition during the publication period.\n\n"
        "For each cited senior record, your job is to write ONE paragraph (80–120 words) "
        "explaining:\n"
        "  • Whether the Examiner would likely cite this record in a 2(d) refusal (high / "
        "moderate / low likelihood), with a brief reason grounded in classes, recital "
        "overlap, and mark identity.\n"
        "  • If applicable, recommended response strategy (narrow recital, distinguishing "
        "matter, coexistence claim).\n\n"
        "Tone: professional, second-person voice as if writing to the client. Personable "
        "and direct rather than dry forensic register. Do not hedge unnecessarily, but be "
        "explicit when the data is incomplete.\n\n"
        "Output: strict JSON. An object mapping each record_id to its commentary paragraph "
        "as a single string. No extra prose outside the JSON."
    )


def _post_registration_system_prompt() -> str:
    return (
        "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist at "
        "Braudit. You are writing the per-record forensic commentary for a "
        "POST-REGISTRATION trademark audit. The client already has a registered "
        "trademark and wants to know which third parties might be operating in their "
        "space — potential infringement risks, reasons to take action.\n\n"
        "For each cited record, write ONE paragraph (80–120 words) explaining:\n"
        "  • Whether this third party is plausibly infringing the client's live rights "
        "(critical / high / medium / low priority).\n"
        "  • The reason — classes, recital overlap, mark proximity, channel of trade.\n"
        "  • Recommended action: cease-and-desist letter, opposition watch, monitor only, "
        "or no action warranted.\n\n"
        "Tone: professional, second-person, direct. Frame enforcement decisions in terms "
        "of commercial impact on the client, not just legal similarity. Avoid hedging.\n\n"
        "Output: strict JSON. An object mapping each record_id to its commentary "
        "paragraph as a single string. No extra prose outside the JSON."
    )


def _build_record_payload(record: VerifiedRecord, score: ForensicScore,
                          record_id: str) -> dict:
    """Compact dict per record for the LLM, including only the fields the model
    needs to write commentary. Keeps prompts under token budgets."""
    return {
        'record_id': record_id,
        'office': record.office,
        'app_number': record.app_number,
        'registration_number': record.registration_number,
        'mark_text': record.mark_text,
        'mark_type': record.mark_type,
        'status': record.status,
        'status_stage': record.status_stage,
        'filing_date': record.filing_date,
        'registration_date': record.registration_date,
        'owner_name': record.owner_name,
        'nice_classes': record.nice_classes,
        'goods_services': record.goods_services,
        'forensic_score': score.total,
        'risk_band': score.risk_band,
        'score_explanation': score.explanation,
        'verified': record.verified,
    }


def generate_per_record_commentary(client: NarrativeClient, *,
                                   records: list[VerifiedRecord],
                                   scores: list[ForensicScore],
                                   report_type: ReportType,
                                   client_brand: dict,
                                   is_monitoring: bool = False) -> dict[str, str]:
    """One Sonnet call generates commentary for all records.

    Returns {record_id: commentary_paragraph}. Records keyed by their position
    in the input list (record_0 ... record_N), which the caller maps back.

    `is_monitoring` (BR-011, 11 Jun 2026) swaps in the monitoring system
    prompt — observational, non-alarming tone for scheduled monitoring
    updates. When False (default) the existing PRE/POST application
    branches apply.
    """
    if len(records) != len(scores):
        raise ValueError('records and scores must be the same length')
    if not records:
        return {}

    payload = [
        _build_record_payload(r, s, f'record_{i}')
        for i, (r, s) in enumerate(zip(records, scores))
    ]

    if is_monitoring:
        system = _MONITORING_PER_RECORD_PROMPT
    else:
        system = (_pre_application_system_prompt()
                  if report_type == ReportType.PRE_APPLICATION
                  else _post_registration_system_prompt())

    user = (
        'CLIENT BRAND CONTEXT (the mark being protected):\n'
        f'{json.dumps(client_brand, indent=2)}\n\n'
        'CITED SENIOR RECORDS TO COMMENT ON:\n'
        f'{json.dumps(payload, indent=2)}\n\n'
        'Return STRICT JSON only. The shape is: {"record_0": "paragraph...", '
        '"record_1": "paragraph...", ...} — one key per record_id above.'
    )

    raw = client._call(
        system=system, user=user,
        max_tokens=len(records) * DEFAULT_MAX_TOKENS_PER_RECORD + 800,
    )
    return _robust_json_loads(raw)


# ---------- Monitoring report prompts (BR-011, 11 Jun 2026) ----------------

_MONITORING_PER_RECORD_PROMPT = (
    "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist "
    "at Braudit. You are writing per-record commentary for a SCHEDULED "
    "MONITORING forensic appendix.\n\n"
    "Audience: an existing client receiving their regular monitoring update. "
    "Tone: observational, second-person, brief, non-alarming. The client has "
    "already filed their mark — your job is to surface what we noticed during "
    "this period, NOT to recommend filings or enforcement actions.\n\n"
    "For each cited senior record below, write a single paragraph (80-150 "
    "words) noting: what the record is, whether it appeared in earlier "
    "monitoring periods or is new, and what (if anything) the client should "
    "look at. Do NOT include: cease-and-desist language, priority ranking, "
    "filing recommendations, or alarming framing. If a record looks routine "
    "say so plainly.\n\n"
    'Return STRICT JSON only. Shape: {"record_0": "paragraph...", '
    '"record_1": "paragraph...", ...} — one key per record_id.'
)

_MONITORING_SUMMARY_PROMPT = (
    "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist "
    "at Braudit. You are writing the executive summary and final invitation "
    "for a SCHEDULED MONITORING forensic appendix.\n\n"
    "Audience: an existing client. Tone: observational, second-person, brief, "
    "non-alarming. This is a regular monitoring update, not an audit or "
    "enforcement decision.\n\n"
    "Open by stating that this is the scheduled monitoring period summary. "
    "Briefly note the volume of new records and the risk distribution. "
    "Identify the records that most warrant the client's attention — but do "
    "NOT make filing, enforcement, or priority-ranked recommendations. The "
    "client decides whether to instruct further; you are surfacing what we "
    "saw.\n\n"
    "Then write a Final Recommendation block phrased as an invitation: "
    'review the flagged items, and if they would like to investigate any of '
    'them further, please get in touch and we can advise on the appropriate '
    'next step.\n\n'
    "Output strict JSON with these keys:\n"
    '  "executive_summary"   : (string, 200-350 words, observational)\n'
    '  "final_recommendation": (string, 150-250 words, invitation to engage)\n'
    'No viability score, no recommended specification, no recommended actions.'
)


# ---------- Web-section commentary (BR-010, 11 Jun 2026) ------------------

_EXTRAS_SECTION_BRIEFS = {
    'google': (
        'Google web hit. Comment on how the cited use of the mark relates '
        'to the client\'s brand: is the cited site a competitor, a news '
        'mention, a parking page, a defensive registration? Note the risk '
        'grade returned by the scrape and whether it looks justified.'
    ),
    'companies': (
        'UK Companies House registered company. Comment on whether the '
        'cited company is an active commercial threat (filed accounts, '
        'trading, SIC overlap with client) or a dormant / different-trade '
        'entity. The data shown is verbatim from Companies House.'
    ),
    'domains': (
        'Domain name registration. Comment on whether the cited domain '
        'represents an active site, a defensive registration by an '
        'unrelated party, or a speculative reservation. Note whether the '
        'TLD pattern (.com / .co.uk / .net etc) suggests targeting.'
    ),
    'social': (
        'Social media handle / page. Comment on whether the cited handle '
        'is an active brand presence requiring action, or a dormant '
        'reservation. Mention which platforms are populated.'
    ),
}


def _build_extras_payload(section: str, row: dict, idx: int) -> dict:
    """Pack one selected 3a–3d row into a small JSON-friendly dict that
    the LLM can reason about. Drop bulky fields (image_bytes, raw HTML)
    but keep risk/score/identifiers."""
    pruned = {k: v for k, v in row.items()
              if k not in ('image_bytes', 'raw_html', 'keyword_image_bytes')
              and not isinstance(v, (bytes, bytearray))}
    return {'extra_id': f'{section}_{idx}', **pruned}


def generate_extras_commentary(client: NarrativeClient, *,
                                extras_rows: dict,
                                report_type: ReportType,
                                client_brand: dict) -> dict:
    """Single batched Sonnet call generates commentary for the operator-
    selected rows across sections 3a (Google), 3b (Companies),
    3c (Domains), 3d (Social).

    `extras_rows` is the dict produced by the app's custom-selection UI:
        {'google': [row, row, ...], 'companies': [...], 'domains': [...], 'social': [...]}

    Returns a parallel dict where each row is augmented with a
    'commentary' field:
        {'google': [{'row': row, 'commentary': 'paragraph...'}, ...], ...}

    Empty input sections produce empty output sections. Total cost
    scales with the number of selected rows, not the number of original
    results — operators control spend by selecting carefully.
    """
    # Build a flat payload across sections so we can do it in one call.
    flat_payload: list[dict] = []
    for section in ('google', 'companies', 'domains', 'social'):
        rows = extras_rows.get(section) or []
        for idx, row in enumerate(rows):
            flat_payload.append({
                'section': section,
                'briefing': _EXTRAS_SECTION_BRIEFS[section],
                **_build_extras_payload(section, row, idx),
            })

    out = {'google': [], 'companies': [], 'domains': [], 'social': []}
    if not flat_payload:
        return out

    system = (
        "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist "
        "at Braudit. You are writing forensic commentary for a "
        f"{'PRE-APPLICATION' if report_type == ReportType.PRE_APPLICATION else 'POST-REGISTRATION'} "
        "trademark audit appendix.\n\n"
        "The operator has hand-picked individual hits from the Google / "
        "Companies House / Domains / Social Media scrapes for forensic "
        "treatment. For each one, write a single short paragraph "
        "(80-150 words) that explains what the hit means for the client.\n\n"
        "Tone: professional, second-person, concrete. Cite the data from "
        "the row directly rather than speculating. If the data is "
        "ambiguous, say so honestly — don't invent commercial detail.\n\n"
        "Output STRICT JSON. The shape is exactly:\n"
        '{\n'
        '  "google":    {"google_0": "paragraph", "google_1": "paragraph", ...},\n'
        '  "companies": {"companies_0": "paragraph", ...},\n'
        '  "domains":   {"domains_0": "paragraph", ...},\n'
        '  "social":    {"social_0": "paragraph", ...}\n'
        '}\n'
        'Include keys only for sections that have entries. The keys must '
        'match the extra_id values from the input.'
    )

    user = (
        'CLIENT BRAND CONTEXT (the mark being protected):\n'
        f'{json.dumps(client_brand, indent=2)}\n\n'
        'SELECTED ROWS TO COMMENT ON:\n'
        f'{json.dumps(flat_payload, indent=2)}\n\n'
        'Return STRICT JSON only, per the schema above.'
    )

    raw = client._call(
        system=system, user=user,
        max_tokens=len(flat_payload) * DEFAULT_MAX_TOKENS_PER_RECORD + 600,
    )
    parsed = _robust_json_loads(raw)

    # Stitch commentaries back onto rows in the original order.
    for section in ('google', 'companies', 'domains', 'social'):
        rows = extras_rows.get(section) or []
        commentaries = (parsed.get(section) if isinstance(parsed, dict) else None) or {}
        for idx, row in enumerate(rows):
            extra_id = f'{section}_{idx}'
            commentary = commentaries.get(extra_id, '')
            out[section].append({'row': row, 'commentary': commentary})
    return out


# ---------- Executive summary + final recommendation ----------

def _summary_system_prompt(report_type: ReportType, is_monitoring: bool = False) -> str:
    if is_monitoring:
        # BR-011 — scheduled monitoring updates use a softer prompt that
        # doesn't pre-judge filing or enforcement action.
        return _MONITORING_SUMMARY_PROMPT
    if report_type == ReportType.PRE_APPLICATION:
        return (
            "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist at "
            "Braudit. You are writing the executive summary and final recommendation for "
            "a PRE-APPLICATION forensic trademark audit.\n\n"
            "Audience: the client and their account manager. Tone: professional, "
            "personable, second-person.\n\n"
            "Open with the 'common word' framing if the client's mark is a common English "
            "word (multiple senior rights are normal, not automatically blocking). State "
            "what the examining authority actually weighs (similarity of goods/services, "
            "commercial impression, channels of trade). Identify the strongest senior "
            "records by name/number/recital. Give the headline filing-viability score.\n\n"
            "Then write a Final Recommendation block in second-person: file as-is / file "
            "with narrowed recital / file with distinguishing matter / don't file. "
            "Recommend specialist in-jurisdiction counsel for prosecution where helpful "
            "(NOT Madrid Protocol only).\n\n"
            "Output strict JSON with these keys:\n"
            '  "executive_summary"   : (string, 300–500 words)\n'
            '  "viability_score"     : (integer 1–10)\n'
            '  "viability_label"     : (string, one of: Strongly Viable / Viable / '
            'Conditionally Viable / At Risk / Not Recommended)\n'
            '  "final_recommendation": (string, 250–400 words, personable second-person)\n'
            '  "recommended_specification": (object mapping class number to suggested '
            'recital text with negative limitations walking past the top-cited records)'
        )
    return (
        "You are Alex Pugh, Trademark Intelligence & Brand Protection Specialist at "
        "Braudit. You are writing the executive summary and final recommendation for a "
        "POST-REGISTRATION forensic trademark audit.\n\n"
        "Audience: the client and their account manager. Tone: professional, personable, "
        "second-person.\n\n"
        "Frame the audit in terms of enforcement strength and commercial harm. Identify "
        "the highest-priority third-party records for action. Distinguish between "
        "'infringement' (legal conflict with the client's registered rights) and "
        "'co-existence' (similar marks in different channels). Give the headline "
        "enforcement-priority distribution.\n\n"
        "Then write a Final Recommendation block in second-person: which records to "
        "enforce against in priority order, which to put on watch, which to take no "
        "action on, and the rationale for each. Recommend specialist in-jurisdiction "
        "counsel for enforcement steps where helpful.\n\n"
        "Output strict JSON with these keys:\n"
        '  "executive_summary"   : (string, 300–500 words)\n'
        '  "enforcement_priority": (string, one of: Critical / High / Medium / Low / '
        'Monitor Only)\n'
        '  "final_recommendation": (string, 250–400 words, personable second-person)\n'
        '  "recommended_actions" : (array of objects, each with keys '
        '"record_id", "action", "rationale" — action one of: '
        '"Cease-and-desist", "Opposition watch", "Monitor", "No action")'
    )


def generate_summary(client: NarrativeClient, *,
                     records: list[VerifiedRecord],
                     scores: list[ForensicScore],
                     report_type: ReportType,
                     client_brand: dict,
                     is_monitoring: bool = False) -> dict:
    """One Sonnet call for executive summary + final recommendation +
    type-specific block (recommended specification or recommended actions).

    `is_monitoring` (BR-011, 11 Jun 2026) routes to the monitoring summary
    prompt which omits viability score / recommended spec / recommended
    actions and frames the close as an invitation to engage further.

    Returns the JSON object from the model.
    """
    payload = [
        _build_record_payload(r, s, f'record_{i}')
        for i, (r, s) in enumerate(zip(records, scores))
    ]

    user = (
        'CLIENT BRAND CONTEXT:\n'
        f'{json.dumps(client_brand, indent=2)}\n\n'
        'ALL CITED RECORDS WITH FORENSIC SCORES:\n'
        f'{json.dumps(payload, indent=2)}\n\n'
        'Return STRICT JSON only in the exact shape specified in the system prompt.'
    )
    raw = client._call(
        system=_summary_system_prompt(report_type, is_monitoring=is_monitoring),
        user=user,
        max_tokens=DEFAULT_MAX_TOKENS_SUMMARY,
    )
    return _robust_json_loads(raw)


# ---------- Top-level orchestration ----------

@dataclass
class ForensicReport:
    """Everything the docx renderer needs to produce the Step 6 appendix."""
    report_type: ReportType
    records: list[VerifiedRecord]
    scores: list[ForensicScore]
    commentaries: dict[str, str]    # record_id -> paragraph
    summary: dict                   # exec_summary, final_recommendation, etc.
    client_brand: dict
    generation_meta: dict           # model used, token spend, timestamps
    # BR-010 (11 Jun 2026) — operator-selected rows from sections 3a–3d
    # get LLM commentary so they appear alongside the trademark cards in
    # the appendix. Each top-level key is a section ('google', 'companies',
    # 'domains', 'social'); each value is a list of {'row': dict,
    # 'commentary': str} pairs. Empty by default — sections with no
    # selection render nothing.
    extras: dict = field(default_factory=lambda: {
        'google': [], 'companies': [], 'domains': [], 'social': [],
    })
    # BR-011 (11 Jun 2026) — scheduled-monitoring deliverable flag.
    # When True the forensic_report renderer skips the "Recommended
    # Specification" / "Recommended Actions" sections (which presume a
    # filing or enforcement decision) and the title prefix changes to
    # "Monitoring". Defaults False so legacy callers behave identically.
    is_monitoring: bool = False

    def to_dict(self) -> dict:
        return {
            'report_type': self.report_type.value,
            'records': [r.to_dict() for r in self.records],
            'scores': [s.to_dict() for s in self.scores],
            'commentaries': self.commentaries,
            'summary': self.summary,
            'client_brand': self.client_brand,
            'generation_meta': self.generation_meta,
            'extras': self.extras,
        }


def _fallback_summary(report_type: ReportType, error: str) -> dict:
    """Minimal summary dict so the appendix renders even when the
    LLM summary call fails (truncated JSON, API timeout, etc.)."""
    base = {
        'executive_summary': (
            '[Executive summary generation failed.] '
            'The deterministic forensic scoring (see Section 4) is still '
            f'valid and the per-record cards below have been preserved. '
            f'Cause: {error}.'
        ),
        'final_recommendation': (
            '[Final recommendation generation failed.] '
            'Specialist counsel review of the scoring table and per-record '
            'cards is recommended.'
        ),
    }
    if report_type == ReportType.PRE_APPLICATION:
        base.update({
            'viability_score': 0,
            'viability_label': 'Assessment unavailable',
            'recommended_specification': {},
        })
    else:
        base.update({
            'enforcement_priority': 'Unknown',
            'recommended_actions': [],
        })
    return base


def run_forensic_layer(*, signa_records: list[VerifiedRecord],
                       client_brand: dict,
                       report_type: ReportType,
                       narrative_client: NarrativeClient,
                       client_classes: list[int],
                       client_mark: str,
                       client_jurisdiction: str | list[str] = 'US',
                       extras_rows: dict | None = None,
                       is_monitoring: bool = False) -> ForensicReport:
    """Top-level orchestrator. Scores all records, generates per-record
    commentary in one batched call, generates executive summary in one call,
    returns a ForensicReport ready for docx rendering.

    `extras_rows` (BR-010, 11 Jun 2026) carries operator-selected rows from
    sections 3a–3d (Google / Companies / Domains / Social). When supplied,
    a separate batched LLM call generates a short commentary for each
    selected row and the appendix grows new sections 6a–6d. Pass None or
    an all-empty dict for the trademark-only default behaviour.

    If the executive summary fails (e.g. JSON parse error from a truncated
    LLM response), the audit still completes with a fallback summary so the
    user gets a usable appendix instead of a hard crash.
    """
    from datetime import datetime
    started = datetime.utcnow().isoformat() + 'Z'

    scores = [
        score_record(r, client_classes=client_classes, client_mark=client_mark,
                     client_jurisdiction=client_jurisdiction)
        for r in signa_records
    ]

    # Per-record commentary failure also falls back gracefully.
    try:
        commentaries = generate_per_record_commentary(
            narrative_client,
            records=signa_records, scores=scores,
            report_type=report_type, client_brand=client_brand,
            is_monitoring=is_monitoring,
        )
    except Exception as exc:
        commentaries = {
            f'record_{i}': f'[Commentary generation failed: {exc}]'
            for i in range(len(signa_records))
        }

    # BR-010: web-section commentary for operator-selected 3a-3d rows.
    # Empty selection -> empty extras dict, zero LLM cost, no new sections.
    extras = {'google': [], 'companies': [], 'domains': [], 'social': []}
    if extras_rows and any(extras_rows.get(s) for s in ('google', 'companies', 'domains', 'social')):
        try:
            extras = generate_extras_commentary(
                narrative_client,
                extras_rows=extras_rows,
                report_type=report_type,
                client_brand=client_brand,
            )
        except Exception as exc:
            # Failure-tolerant: render the rows verbatim with a placeholder.
            extras = {
                section: [
                    {'row': row, 'commentary': f'[Commentary generation failed: {exc}]'}
                    for row in (extras_rows.get(section) or [])
                ]
                for section in ('google', 'companies', 'domains', 'social')
            }

    try:
        summary = generate_summary(
            narrative_client,
            records=signa_records, scores=scores,
            report_type=report_type, client_brand=client_brand,
            is_monitoring=is_monitoring,
        )
    except Exception as exc:
        summary = _fallback_summary(report_type, str(exc))

    return ForensicReport(
        report_type=report_type,
        records=signa_records,
        scores=scores,
        commentaries=commentaries,
        summary=summary,
        client_brand=client_brand,
        generation_meta={
            'started_at': started,
            'completed_at': datetime.utcnow().isoformat() + 'Z',
            'model': narrative_client._model,
            'record_count': len(signa_records),
            'extras_count': sum(len(extras.get(s, [])) for s in
                                ('google', 'companies', 'domains', 'social')),
            'is_monitoring': is_monitoring,
        },
        extras=extras,
        is_monitoring=is_monitoring,
    )
