"""Trademark jurisdictions / offices the audit tool supports.

Display labels are formatted "Country Name (CODE)"; the parenthetical CODE
is the lookup key the rest of the pipeline uses when comparing client
jurisdiction(s) vs record jurisdiction.

Codes match Signa's `jurisdiction_code` / `office_code` field uppercased,
i.e. ISO 3166-1 alpha-2 for most countries, plus 'EU' (EUIPO), 'WO' (WIPO
Madrid), and the three regional systems we keep separate ('ARIPO', 'OAPI',
'BX' for Benelux/BOIP).
"""
from __future__ import annotations
import re


# (display_label, lookup_code) tuples, grouped by region for human review.
# Order matters for the multi-select dropdown: most common first.
JURISDICTIONS: list[tuple[str, str]] = [
    # --- Most common ---
    ('United Kingdom (GB)', 'GB'),
    ('EUIPO \u2014 European Union (EU)', 'EU'),
    ('United States (US)', 'US'),
    ('China (CN)', 'CN'),
    ('Japan (JP)', 'JP'),
    ('South Korea (KR)', 'KR'),
    ('Canada (CA)', 'CA'),
    ('Australia (AU)', 'AU'),
    ('India (IN)', 'IN'),
    ('WIPO / Madrid (WO)', 'WO'),

    # --- Tier 2 ---
    ('UAE (AE)', 'AE'),
    ('Saudi Arabia (SA)', 'SA'),
    ('Singapore (SG)', 'SG'),
    ('Hong Kong (HK)', 'HK'),
    ('Brazil (BR)', 'BR'),
    ('Mexico (MX)', 'MX'),
    ('South Africa (ZA)', 'ZA'),

    # --- Europe ---
    ('Albania (AL)', 'AL'),
    ('Andorra (AD)', 'AD'),
    ('Armenia (AM)', 'AM'),
    ('Austria (AT)', 'AT'),
    ('Azerbaijan (AZ)', 'AZ'),
    ('Belarus (BY)', 'BY'),
    ('Belgium (BE)', 'BE'),
    ('Benelux \u2014 BOIP (BX)', 'BX'),
    ('Bosnia and Herzegovina (BA)', 'BA'),
    ('Bulgaria (BG)', 'BG'),
    ('Croatia (HR)', 'HR'),
    ('Cyprus (CY)', 'CY'),
    ('Czech Republic (CZ)', 'CZ'),
    ('Denmark (DK)', 'DK'),
    ('Estonia (EE)', 'EE'),
    ('Finland (FI)', 'FI'),
    ('France (FR)', 'FR'),
    ('Georgia (GE)', 'GE'),
    ('Germany (DE)', 'DE'),
    ('Greece (GR)', 'GR'),
    ('Hungary (HU)', 'HU'),
    ('Iceland (IS)', 'IS'),
    ('Ireland (IE)', 'IE'),
    ('Italy (IT)', 'IT'),
    ('Kosovo (XK)', 'XK'),
    ('Latvia (LV)', 'LV'),
    ('Liechtenstein (LI)', 'LI'),
    ('Lithuania (LT)', 'LT'),
    ('Luxembourg (LU)', 'LU'),
    ('Malta (MT)', 'MT'),
    ('Moldova (MD)', 'MD'),
    ('Monaco (MC)', 'MC'),
    ('Montenegro (ME)', 'ME'),
    ('Netherlands (NL)', 'NL'),
    ('North Macedonia (MK)', 'MK'),
    ('Norway (NO)', 'NO'),
    ('Poland (PL)', 'PL'),
    ('Portugal (PT)', 'PT'),
    ('Romania (RO)', 'RO'),
    ('San Marino (SM)', 'SM'),
    ('Serbia (RS)', 'RS'),
    ('Slovakia (SK)', 'SK'),
    ('Slovenia (SI)', 'SI'),
    ('Spain (ES)', 'ES'),
    ('Sweden (SE)', 'SE'),
    ('Switzerland (CH)', 'CH'),
    ('Turkey (TR)', 'TR'),
    ('Ukraine (UA)', 'UA'),

    # --- Middle East ---
    ('Bahrain (BH)', 'BH'),
    ('Iraq (IQ)', 'IQ'),
    ('Jordan (JO)', 'JO'),
    ('Kuwait (KW)', 'KW'),
    ('Lebanon (LB)', 'LB'),
    ('Oman (OM)', 'OM'),
    ('Qatar (QA)', 'QA'),
    ('Syria (SY)', 'SY'),
    ('Yemen (YE)', 'YE'),

    # --- Africa ---
    ('ARIPO (regional)', 'ARIPO'),
    ('OAPI (regional)', 'OAPI'),
    ('Algeria (DZ)', 'DZ'),
    ('Angola (AO)', 'AO'),
    ('Botswana (BW)', 'BW'),
    ('Cabo Verde (CV)', 'CV'),
    ('Egypt (EG)', 'EG'),
    ('Eswatini (SZ)', 'SZ'),
    ('Ethiopia (ET)', 'ET'),
    ('Gambia (GM)', 'GM'),
    ('Ghana (GH)', 'GH'),
    ('Kenya (KE)', 'KE'),
    ('Lesotho (LS)', 'LS'),
    ('Liberia (LR)', 'LR'),
    ('Madagascar (MG)', 'MG'),
    ('Malawi (MW)', 'MW'),
    ('Mauritius (MU)', 'MU'),
    ('Morocco (MA)', 'MA'),
    ('Mozambique (MZ)', 'MZ'),
    ('Namibia (NA)', 'NA'),
    ('Nigeria (NG)', 'NG'),
    ('Rwanda (RW)', 'RW'),
    ('Sierra Leone (SL)', 'SL'),
    ('Tanzania (TZ)', 'TZ'),
    ('Uganda (UG)', 'UG'),
    ('Zambia (ZM)', 'ZM'),
    ('Zimbabwe (ZW)', 'ZW'),

    # --- Asia (excluding Tier 1/2) ---
    ('Afghanistan (AF)', 'AF'),
    ('Bangladesh (BD)', 'BD'),
    ('Bhutan (BT)', 'BT'),
    ('Brunei (BN)', 'BN'),
    ('Cambodia (KH)', 'KH'),
    ('Indonesia (ID)', 'ID'),
    ('Kazakhstan (KZ)', 'KZ'),
    ('Kyrgyzstan (KG)', 'KG'),
    ('Lao PDR (LA)', 'LA'),
    ('Malaysia (MY)', 'MY'),
    ('Maldives (MV)', 'MV'),
    ('Mongolia (MN)', 'MN'),
    ('Myanmar (MM)', 'MM'),
    ('Nepal (NP)', 'NP'),
    ('Pakistan (PK)', 'PK'),
    ('Philippines (PH)', 'PH'),
    ('Sri Lanka (LK)', 'LK'),
    ('Tajikistan (TJ)', 'TJ'),
    ('Thailand (TH)', 'TH'),
    ('Turkmenistan (TM)', 'TM'),
    ('Uzbekistan (UZ)', 'UZ'),
    ('Vietnam (VN)', 'VN'),

    # --- Americas (LATAM & Caribbean) ---
    ('Antigua and Barbuda (AG)', 'AG'),
    ('Argentina (AR)', 'AR'),
    ('Bahamas (BS)', 'BS'),
    ('Barbados (BB)', 'BB'),
    ('Belize (BZ)', 'BZ'),
    ('Bolivia (BO)', 'BO'),
    ('Chile (CL)', 'CL'),
    ('Colombia (CO)', 'CO'),
    ('Costa Rica (CR)', 'CR'),
    ('Cuba (CU)', 'CU'),
    ('Cura\u00e7ao (CW)', 'CW'),
    ('Dominica (DM)', 'DM'),
    ('Dominican Republic (DO)', 'DO'),
    ('Ecuador (EC)', 'EC'),
    ('El Salvador (SV)', 'SV'),
    ('Grenada (GD)', 'GD'),
    ('Guatemala (GT)', 'GT'),
    ('Honduras (HN)', 'HN'),
    ('Jamaica (JM)', 'JM'),
    ('Nicaragua (NI)', 'NI'),
    ('Panama (PA)', 'PA'),
    ('Paraguay (PY)', 'PY'),
    ('Peru (PE)', 'PE'),
    ('Saint Kitts and Nevis (KN)', 'KN'),
    ('Saint Lucia (LC)', 'LC'),
    ('Saint Vincent and the Grenadines (VC)', 'VC'),
    ('Trinidad and Tobago (TT)', 'TT'),
    ('Uruguay (UY)', 'UY'),
    ('Venezuela (VE)', 'VE'),

    # --- Oceania ---
    ('Fiji (FJ)', 'FJ'),
    ('Papua New Guinea (PG)', 'PG'),
    ('Samoa (WS)', 'WS'),
    ('Tonga (TO)', 'TO'),
    ('Vanuatu (VU)', 'VU'),
]

# Lookup tables built once at import time
LABEL_TO_CODE: dict[str, str] = {label: code for label, code in JURISDICTIONS}
CODE_TO_LABEL: dict[str, str] = {code: label for label, code in JURISDICTIONS}

# Common informal aliases the operator might have typed into Order Form R10
# (Designated countries). Maps an uppercased word to a JURISDICTION code.
ALIASES: dict[str, str] = {
    'UK': 'GB', 'GREAT BRITAIN': 'GB', 'BRITAIN': 'GB',
    'UNITED KINGDOM': 'GB', 'ENGLAND': 'GB', 'SCOTLAND': 'GB', 'WALES': 'GB',
    'EUIPO': 'EU', 'EU': 'EU', 'EUROPEAN UNION': 'EU', 'EUROPE': 'EU',
    'COMMUNITY': 'EU', 'CTM': 'EU',
    'USA': 'US', 'US': 'US', 'UNITED STATES': 'US', 'AMERICA': 'US', 'USPTO': 'US',
    'WIPO': 'WO', 'MADRID': 'WO', 'INTERNATIONAL': 'WO',
    'PRC': 'CN', 'CHINA': 'CN', 'HONG KONG': 'HK', 'HK': 'HK',
    'JAPAN': 'JP', 'KOREA': 'KR', 'SOUTH KOREA': 'KR',
    'CANADA': 'CA', 'AUSTRALIA': 'AU', 'INDIA': 'IN',
    'UAE': 'AE', 'EMIRATES': 'AE',
    'SOUTH AFRICA': 'ZA', 'SINGAPORE': 'SG',
    'BRAZIL': 'BR', 'MEXICO': 'MX',
    'GERMANY': 'DE', 'DEUTSCHLAND': 'DE', 'FRANCE': 'FR',
    'SPAIN': 'ES', 'ITALY': 'IT', 'PORTUGAL': 'PT',
    'NETHERLANDS': 'NL', 'HOLLAND': 'NL',
    'BENELUX': 'BX', 'BOIP': 'BX',
    'IRELAND': 'IE', 'BELGIUM': 'BE', 'AUSTRIA': 'AT',
    'SWITZERLAND': 'CH', 'NORWAY': 'NO', 'SWEDEN': 'SE', 'DENMARK': 'DK',
    'FINLAND': 'FI', 'POLAND': 'PL', 'TURKEY': 'TR',
}


def all_labels() -> list[str]:
    """All display labels in display order (most common first)."""
    return [label for label, _ in JURISDICTIONS]


def labels_for_codes(codes: list[str]) -> list[str]:
    """Convert codes back to display labels for the multiselect default."""
    return [CODE_TO_LABEL[c] for c in codes if c in CODE_TO_LABEL]


def codes_for_labels(labels: list[str]) -> list[str]:
    """Convert selected display labels into the underlying codes."""
    return [LABEL_TO_CODE[l] for l in labels if l in LABEL_TO_CODE]


def parse_country_string(s: str) -> list[str]:
    """Best-effort parse of a free-text 'Designated countries' string.

    Accepts comma/semicolon/slash/'and'/'&' separated lists of either
    country names ('United Kingdom', 'Germany'), informal aliases ('UK',
    'USA', 'EUIPO'), or raw ISO codes ('GB', 'DE', 'US'). Returns a
    de-duplicated list of canonical codes.

    Unknown tokens are silently dropped \u2014 the operator can still
    multi-select missing jurisdictions manually in the UI.
    """
    if not s:
        return []
    parts = re.split(r'[,;/&\n]|\band\b', s, flags=re.IGNORECASE)
    out: list[str] = []
    for part in parts:
        p = part.strip().upper()
        if not p:
            continue
        # Direct ISO code match (2 letters)
        if p in CODE_TO_LABEL:
            out.append(p)
            continue
        # Alias match
        if p in ALIASES:
            out.append(ALIASES[p])
            continue
        # Country-name match against the JURISDICTIONS table
        # (strip the trailing parenthetical from each label)
        matched = False
        for label, code in JURISDICTIONS:
            country_name = label.split('(')[0].strip().upper()
            # Handle "Brunei" matching "Brunei (BN)" but also partial-word
            # matches like "BENELUX" → "Benelux \u2014 BOIP (BX)"
            if country_name == p or country_name.startswith(p) and len(p) >= 4:
                out.append(code)
                matched = True
                break
        if matched:
            continue
    # Dedupe preserving order
    seen: set[str] = set()
    deduped: list[str] = []
    for code in out:
        if code not in seen:
            seen.add(code)
            deduped.append(code)
    return deduped
