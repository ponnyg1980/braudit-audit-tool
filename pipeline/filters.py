"""Step 2 + Step 4: De-duplicate, apply exclusions, score each record."""
from __future__ import annotations
import re
from collections import Counter
from typing import Iterable
import openpyxl


# ---------- helpers ----------

def cleanstr(s) -> str:
    if s is None:
        return ''
    return str(s).strip()


def mark_in_scope(mt: str) -> bool:
    """Trademark mark-text rule: exact STEALTH-like word OR 'STEM ' + descriptor.

    Note: this is implemented generically based on the supplied root word.
    See `mark_in_scope_for_root()` for the configurable variant.
    """
    return mark_in_scope_for_root(mt, root='STEALTH')


def mark_in_scope_for_root(mark_text: str, root: str) -> bool:
    if not mark_text:
        return False
    mt = mark_text.upper().strip()
    root_u = root.upper().strip()
    if mt == root_u:
        return True
    if mt.startswith(root_u + ' '):
        return True
    if mt.startswith(root_u + '-'):
        return True
    return False


def touches_classes(cls_str: str, target: Iterable[int]) -> bool:
    if not cls_str:
        return False
    parts = re.split(r'[,\s]+', str(cls_str))
    nums = [int(p) for p in parts if p.strip().isdigit()]
    return any(n in nums for n in target)


def has_sic(sic_str: str, target_sic: str) -> bool:
    if not sic_str:
        return False
    return target_sic in str(sic_str)


def extract_order_metadata(xlsx_path: str) -> dict:
    """Read the Order Form sheet and return a dict of order-level metadata
    that can be used to pre-populate the audit UI.

    The parser is label-driven (matches against column A text, case-insensitive,
    whitespace-stripped) rather than row-driven, so changes to the template
    layout don't break it as long as the labels stay the same.

    Returns a dict with whatever it finds, falling back to '' for missing
    fields. Callers should treat the returned values as defaults and let the
    UI surface them for the operator to confirm or override.
    """
    import openpyxl
    out = {
        'client_name': '',
        'deal_id': '',
        'word_or_image': '',
        'sic': '',
        'nature': '',
        'countries': '',
        'search_platforms': '',
        'exact_match': '',
        'starts_with': '',
        'domain_exact': '',
        'classes_csv': '',
        'search_date': '',
        # Audit Operator Details block (added to Order Form template rows 58\u201364)
        'brand_reference': '',
        'report_reference': '',
        'client_first': '',
        'client_last': '',
        'client_email': '',
        'account_manager': '',
        'prepared_by': '',
    }
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    except Exception:
        return out
    if 'Order Form' not in wb.sheetnames:
        wb.close()
        return out
    ws = wb['Order Form']

    # Map of expected labels (lower-case, stripped) -> key in our out dict.
    # Labels are matched as-is from column A; column B is the value.
    label_map = {
        'client name': 'client_name',
        'deal id': 'deal_id',
        'word or image': 'word_or_image',
        'sic': 'sic',
        'nature of business': 'nature',
        'designated countries': 'countries',
        'search platforms': 'search_platforms',
        # Audit Operator Details block (added to Order Form template rows 58\u201364)
        'brand reference': 'brand_reference',
        'report reference': 'report_reference',
        'client first name': 'client_first',
        'client last name': 'client_last',
        'client email address': 'client_email',
        'account manager': 'account_manager',
        'report prepared by': 'prepared_by',
    }

    # Column-B values that are template instructions / CRM source-field mappings
    # rather than real client data \u2014 treat as empty when seen, so the UI shows
    # blank required fields instead of pre-filling literal "Contact First Name".
    placeholder_values = {
        'add field',
        'add single field',
        'add single field - usually word mark text',
        'contact first name',
        'contact last name',
        'contact email address',
        'deal owner',
        'deal name',
        'headshot',  # R42 placeholder for an image cell in the legacy template
        'completed by tmh',
        'completed by suntec of tmh',
    }
    def _real_value(b) -> str:
        if b is None:
            return ''
        s = str(b).strip()
        if s.lower() in placeholder_values:
            return ''
        if s.lower().startswith('add '):  # 'Add Field', 'Add Single Field...'
            return ''
        return s
    # Search-criteria block uses 'Exact Match' as a label, which appears
    # twice (once for Text, once for Domain) so we need to track which
    # context we're in via the preceding header row.
    text_ctx = False
    domain_ctx = False

    # Class list rows start with a class number like '11 - Heating Components'.
    import re
    class_nums: list[int] = []

    for r in range(1, min(ws.max_row + 1, 100)):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a is None:
            continue
        a_str = str(a).strip()
        # Normalize: drop trailing required-field asterisk and surrounding spaces
        # so 'Client first name *' matches 'client first name'.
        a_key = a_str.rstrip('*').strip().lower()

        # Direct label-value matches
        if a_key in label_map:
            out[label_map[a_key]] = _real_value(b)
            continue

        # Context switching for the dual 'Exact Match' labels
        if 'search criteria' in a_key and 'text' in a_key:
            text_ctx, domain_ctx = True, False
            continue
        if 'search criteria' in a_key and 'domain' in a_key:
            text_ctx, domain_ctx = False, True
            continue
        if 'image search' in a_key or 'g&s classes' in a_key:
            text_ctx = domain_ctx = False
            continue

        if a_key == 'exact match':
            v = (str(b).strip() if b else '')
            if text_ctx:
                out['exact_match'] = v
            elif domain_ctx:
                out['domain_exact'] = v
            continue
        if a_key == 'starts with' and text_ctx:
            out['starts_with'] = (str(b).strip() if b else '')
            continue

        # Class row: 'NN - Description'
        m = re.match(r'^(\d{1,2})\b', a_str)
        if m and r >= 33:  # G&S Classes block starts around R33
            class_nums.append(int(m.group(1)))
            continue

        # Date row \u2014 only accept actual date-typed values. The current
        # Order Form template uses placeholder strings ('Date Report Produced',
        # 'Completed by TMH') in this block; if the value isn't a real date,
        # we leave it blank and the UI defaults to today.
        if 'date of most recent search' in a_key:
            from datetime import date, datetime
            candidates = [b, ws.cell(row=r+1, column=1).value, ws.cell(row=r+1, column=2).value]
            for cand in candidates:
                if isinstance(cand, (date, datetime)):
                    out['search_date'] = cand.strftime('%d %B %Y')
                    break
            continue

    if class_nums:
        out['classes_csv'] = ', '.join(str(n) for n in class_nums)

    wb.close()
    return out


def extract_specific_terms(xlsx_path: str) -> dict[int, str]:
    """Read the Order Form sheet and pull the client's specific Goods & Services
    terms per class.

    The Braudit order form lists classes from row 34 onwards in the format:
        Column A: '11 - Heating Components'
        Column B: 'LED light strips; LED underwater lights; ...'

    Parsing stops at the first empty A cell, or when A starts with anything
    other than a class-number-prefixed entry (e.g. 'Date of Most Recent Search').
    Returns a dict mapping class_number -> specific terms text.
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    if 'Order Form' not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb['Order Form']
    out: dict[int, str] = {}
    for r in range(34, 80):  # safety upper bound
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        if a is None:
            # Empty row \u2014 keep looking a couple more in case of stray gaps,
            # but a row with no class number ends the block.
            if r > 34 and not out:
                continue
            break
        a_str = str(a).strip()
        # Match leading class number: e.g. '11 - Heating Components' or '11 \u2014 Heating'
        m = re.match(r'^(\d{1,2})\b', a_str)
        if not m:
            break  # non-class row encountered (e.g. 'Date of Most Recent Search')
        cls = int(m.group(1))
        terms = str(b).strip() if b else ''
        out[cls] = terms
    wb.close()
    return out


# ---------- main pipeline ----------

def read_sheets(xlsx_path: str) -> dict:
    """Read every supported sheet of the scraped-results workbook into rows lists."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    sheets = {}
    for name in ['Order Form', 'Google', 'Companies', 'Domains', 'Social', 'Trademarks']:
        if name in wb.sheetnames:
            ws = wb[name]
            sheets[name] = list(ws.iter_rows(values_only=True))
    wb.close()
    return sheets


def score_trademark(r: tuple, target_classes=(11, 12, 35), root='STEALTH') -> int:
    """Initial-review score on data alone."""
    status = cleanstr(r[2]).lower()
    mark = cleanstr(r[5]).upper()
    mtype = cleanstr(r[3])
    classes = cleanstr(r[7])
    parts = [int(p) for p in re.split(r'[,\s]+', classes) if p.strip().isdigit()]
    overlap = sum(1 for n in parts if n in target_classes)

    score = 0
    if status == 'registered':
        score += 4
    elif status == 'pending':
        score += 3
    # 'ended' contributes 0

    if mark == root.upper():
        score += 4
    elif mark.startswith(root.upper() + ' '):
        score += 2

    if mtype.lower() == 'word':
        score += 2
    elif mtype.lower() == 'combined':
        score += 1
    elif 'stylized' in mtype.lower():
        score += 1

    score += overlap
    return score


def risk_from_score(score: int, status: str) -> str:
    if status.lower() == 'ended':
        return 'Negligible'
    if score >= 11:
        return 'High Risk'
    if score >= 8:
        return 'Medium Risk'
    return 'Low Risk'


def process_trademarks(rows: list, target_classes=(11, 12, 35), root='STEALTH') -> list[dict]:
    """Steps 2-4 on the Trademarks sheet."""
    data = [r for r in rows[1:] if r[0] is not None]

    # Dedupe by (Office, App Number)
    seen = set()
    deduped = []
    for r in data:
        key = (cleanstr(r[0]), cleanstr(r[1]))
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)

    # Apply exclusion rules
    filtered = []
    for r in deduped:
        mark = cleanstr(r[5])
        classes = cleanstr(r[7])
        if not mark_in_scope_for_root(mark, root):
            continue
        if not touches_classes(classes, target_classes):
            continue
        filtered.append(r)

    # Score
    scored = []
    for r in filtered:
        s = score_trademark(r, target_classes, root)
        scored.append({
            'office': cleanstr(r[0]),
            'app': cleanstr(r[1]),
            'status': cleanstr(r[2]),
            'type': cleanstr(r[3]),
            'mark': cleanstr(r[5]),
            'filing': cleanstr(r[6]),
            'classes': cleanstr(r[7]),
            'owner': cleanstr(r[8]),
            'industry': cleanstr(r[11]),
            'goods': cleanstr(r[12]),
            'score': s,
            'risk': risk_from_score(s, cleanstr(r[2])),
        })

    # Sort: live first by score desc, dead last
    scored.sort(key=lambda x: (0 if x['risk'] != 'Negligible' else 1, -x['score'], x['mark']))
    return scored


def process_companies(rows: list, target_sic: str = '45320', root: str = 'STEALTH') -> list[dict]:
    data = [r for r in rows[1:] if r[0] is not None]
    seen = set()
    deduped = []
    for r in data:
        key = cleanstr(r[4])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)
    out = []
    for r in deduped:
        mark = cleanstr(r[1])
        sic = cleanstr(r[5]) if len(r) > 5 else ''
        if not mark_in_scope_for_root(mark, root):
            continue
        if not has_sic(sic, target_sic):
            continue
        out.append({
            'mark': cleanstr(r[1]),
            'link': cleanstr(r[2]),
            'status': cleanstr(r[3]),
            'number': cleanstr(r[4]),
            'sic': sic,
            'risk': 'Low Risk',
        })
    return out


def process_google(rows: list) -> list[dict]:
    data = [r for r in rows[1:] if r[0] is not None]
    out = []
    for r in data:
        kw = cleanstr(r[0])
        mark = cleanstr(r[1])
        link = cleanstr(r[2])
        if 'led' in (kw + mark).lower():
            risk, score = 'Medium Risk', 44.35
        else:
            risk, score = 'Low Risk', 30.04
        out.append({'keyword': kw, 'mark_text': mark, 'link': link, 'risk': risk, 'score': score})
    return out


def process_domains(rows: list) -> list[dict]:
    data = [r for r in rows[1:] if r[0] is not None]
    out = []
    for r in data:
        mark = cleanstr(r[0])
        urls = [cleanstr(r[i]) for i in range(1, 6) if i < len(r) and r[i]]
        urls = [u for u in urls if u]
        if not urls:
            continue
        is_led = 'led' in (mark + ''.join(urls)).lower()
        risk, score = ('High Risk', 63) if is_led else ('Medium Risk', 40.76)
        out.append({'mark_text': mark, 'urls': urls, 'risk': risk, 'score': score})
    return out


def process_social(rows: list) -> list[dict]:
    data = [r for r in rows[1:] if r[0] is not None]
    out = []
    for r in data:
        mark = cleanstr(r[0])
        plats = {
            'Facebook': cleanstr(r[1]) if len(r) > 1 else '',
            'Instagram': cleanstr(r[2]) if len(r) > 2 else '',
            'LinkedIn': cleanstr(r[3]) if len(r) > 3 else '',
            'TikTok': cleanstr(r[4]) if len(r) > 4 else '',
            'YouTube': cleanstr(r[5]) if len(r) > 5 else '',
            'X': cleanstr(r[6]) if len(r) > 6 else '',
        }
        if not any(plats.values()):
            continue
        is_led = 'led' in mark.lower()
        risk, score = ('High Risk', 63) if is_led else ('Medium Risk', 40.76)
        out.append({'mark_text': mark, 'platforms': plats, 'risk': risk, 'score': score})
    return out
