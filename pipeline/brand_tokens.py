"""TMH brand tokens — single source of truth for colours, fonts, logo paths
and tagline used across all generated reports.

If the brand ever changes, edit the values here and both the initial
monitoring report (`report_builder.py`) and the forensic appendix
(`forensic_report.py`) pick it up on the next render.

Colours are stored as 6-char hex strings (no leading "#") so they can be
passed directly to python-docx's `RGBColor.from_string` and to our
`set_cell_bg` / `add_hyperlink` helpers.

Values verified against the live TMH logo SVG at
https://www.thetrademarkhelpline.com/wp-content/uploads/2024/06/logo.svg
and the canonical pink supplied by Jonathan (08 Jun 2026).
"""
from __future__ import annotations
import os


# ---------- Colours -----------------------------------------------------------

# Primary — TMH pink. Operator-authoritative value (08 Jun 2026). The logo
# SVG uses a near-cousin (#E0004D); we prefer Jonathan's stated value.
BRAND_PINK = 'E51652'

# Secondary navy — used for headings, table-header fills, key/value labels.
# Sampled directly from the wordmark "The Trademark Helpline" in the logo.
BRAND_NAVY = '2D455A'

# Slate — secondary heading / muted body / borders that need to read.
BRAND_SLATE = '617383'

# Light slate — italic disclaimers, "no results" lines, footnote captions.
BRAND_LIGHT_SLATE = '96A2AC'

# Near-black — body text. Slightly off pure black for a softer feel.
BRAND_BODY = '1D1D1B'

# White — for text on dark fills.
BRAND_WHITE = 'FFFFFF'

# Neutral cell border — kept as the generic grey so risk-coloured cells
# stay legible against it.
BRAND_BORDER = 'BFBFBF'

# Risk-band colours — semantic, NOT brand. Kept stable across rebrands so
# the colour-coded scoring table remains readable.
RISK_VERY_HIGH = 'C00000'   # dark red
RISK_HIGH = 'F4B084'        # orange
RISK_MEDIUM = 'FFE699'      # amber
RISK_LOW = 'C5E0B4'         # green
RISK_NEGLIGIBLE = 'D9D9D9'  # grey


# ---------- Typography -------------------------------------------------------

# Font face used for runs. Arial is universal in Word and renders the same
# whether the user opens the docx on macOS, Windows or in Google Docs.
# (TMH's website uses a Google Font, but Word can't embed Google Fonts
# reliably — Arial is the safe cross-platform stand-in.)
BRAND_FONT = 'Arial'


# ---------- Brand assets -----------------------------------------------------

# Tagline that lives under the logo on the cover page and in the footer.
BRAND_TAGLINE = 'Search. Register. Protect.'

# Logo image paths, resolved relative to the directory this module lives in.
# The PNG file is committed alongside this module under pipeline/assets/.
_HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH_FULL = os.path.join(_HERE, 'assets', 'tmh_logo.png')
LOGO_PATH_SMALL = os.path.join(_HERE, 'assets', 'tmh_logo_small.png')


def logo_exists(path: str = LOGO_PATH_FULL) -> bool:
    """True if the logo file is on disk. Lets callers gracefully skip the
    image insert if the asset hasn't been deployed yet."""
    return os.path.isfile(path) and os.path.getsize(path) > 0


# ---------- Layout constants -------------------------------------------------

# US Letter is 8.5"; with 0.75" margins each side we have 7.0" usable.
# Every table in every report MUST fit within this width or it will overflow
# the right margin (the bug we shipped a fix for on 10 Jun 2026).
USABLE_PAGE_WIDTH_IN = 7.0


# ---------- Risk-band thresholds for the trademark scoring rubric -----------

# These are the boundaries used by `risk_from_score()` in filters.py.
# Mirrored here so the legend rendered in the report stays in sync with
# the actual scoring logic — change one place, change both.
RISK_THRESHOLDS = [
    ('High Risk',   '≥ 11'),   # >= 11
    ('Medium Risk', '8 – 10'),
    ('Low Risk',    '≤ 7'),    # <= 7
    ('Negligible',  'status = Ended'),
]
