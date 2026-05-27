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
