"""Step 5: build the Braudit-style Word document from the filtered/scored data."""
from __future__ import annotations
from io import BytesIO
from docx import Document
from docx.shared import Pt, Inches, RGBColor, Cm
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsmap
from docx.oxml import OxmlElement
from datetime import date
from io import BytesIO
from .nice_classes import NICE_HEADINGS, parse_classes


# ---------- styling helpers ----------

def set_cell_bg(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    shd.set(qn('w:val'), 'clear')
    tc_pr.append(shd)


def set_cell_border(cell, color='BFBFBF', size=4):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), str(size))
        b.set(qn('w:color'), color)
        borders.append(b)
    tc_pr.append(borders)


def risk_fill(risk):
    r = (risk or '').lower()
    if 'very high' in r: return 'C00000'
    if 'high' in r: return 'F4B084'
    if 'medium' in r: return 'FFE699'
    if 'low' in r: return 'C5E0B4'
    if 'negligible' in r: return 'D9D9D9'
    return 'FFFFFF'


def tm_url(office: str, app_num: str) -> str:
    """Return the official-register URL for a trademark, based on the office code.

    Supports US (USPTO TSDR), UK (UKIPO), EU (EUIPO eSearch) and WIPO/Madrid
    (WIPO Brand Database). Returns '' for unknown offices so callers can fall
    back to plain text instead of a broken link.
    """
    if not app_num:
        return ''
    o = (office or '').upper().strip()
    n = str(app_num).strip()
    if o in ('US', 'USPTO'):
        return f'https://tsdr.uspto.gov/statusview/sn{n}'
    if o in ('UK', 'UKIPO', 'GB'):
        return f'https://trademarks.ipo.gov.uk/ipo-tmcase/page/Results/1/{n}'
    if o in ('EU', 'EUIPO', 'EM'):
        return f'https://euipo.europa.eu/eSearch/#details/trademarks/{n}'
    if o in ('WIPO', 'MADRID', 'IR', 'WO'):
        return f'https://www3.wipo.int/branddb/en/showData.jsp?ID={n}'
    return ''


def add_hyperlink(paragraph, url, text, color='0563C1', underline=True, font_size=9, bold=False):
    """Inserts an external hyperlink into a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    r_pr = OxmlElement('w:rPr')

    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial'); rFonts.set(qn('w:hAnsi'), 'Arial')
    r_pr.append(rFonts)

    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(font_size * 2)); r_pr.append(sz)
    color_el = OxmlElement('w:color'); color_el.set(qn('w:val'), color); r_pr.append(color_el)
    if underline:
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); r_pr.append(u)
    if bold:
        b = OxmlElement('w:b'); r_pr.append(b)
    new_run.append(r_pr)
    t = OxmlElement('w:t'); t.text = text; t.set(qn('xml:space'), 'preserve')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def run(p, text, *, bold=False, italic=False, color='000000', size=11, font='Arial'):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def add_para(doc, text='', *, bold=False, italic=False, color='000000', size=11, align=None, space_after=4):
    p = doc.add_paragraph()
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        run(p, text, bold=bold, italic=italic, color=color, size=size)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    sizes = {1: 16, 2: 13, 3: 11}
    colors = {1: '1F3864', 2: '1F3864', 3: '2F5496'}
    run(p, text, bold=True, color=colors.get(level, '000000'), size=sizes.get(level, 12))
    if level == 1:
        # Bottom border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '8'); bottom.set(qn('w:space'), '4'); bottom.set(qn('w:color'), '1F3864')
        pBdr.append(bottom); pPr.append(pBdr)
    return p


def add_table(doc, col_widths_in, header_row, body_rows, *, header_fill='1F3864', header_color='FFFFFF',
              risk_col_index=None, hyperlink_col_indexes=None, font_size=9):
    """Generic styled table.
    body_rows: list[list[str]] or list[list[(text, opts_dict)]]
    hyperlink_col_indexes: dict {col_index: lambda row -> url} to render that cell as a link
    risk_col_index: int — colour that cell by risk text
    """
    table = doc.add_table(rows=1 + len(body_rows), cols=len(header_row))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    # set column widths
    for col_idx, w_in in enumerate(col_widths_in):
        for cell in table.columns[col_idx].cells:
            cell.width = Inches(w_in)

    # header
    hdr = table.rows[0]
    for i, txt in enumerate(header_row):
        c = hdr.cells[i]
        c.text = ''
        set_cell_bg(c, header_fill)
        set_cell_border(c)
        c.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run(p, txt, bold=True, color=header_color, size=font_size)

    # body
    hyperlink_col_indexes = hyperlink_col_indexes or {}
    for r_idx, row in enumerate(body_rows):
        tr = table.rows[1 + r_idx]
        for c_idx, val in enumerate(row):
            cell = tr.cells[c_idx]
            cell.text = ''
            set_cell_border(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP

            # risk cell fill
            if risk_col_index is not None and c_idx == risk_col_index:
                set_cell_bg(cell, risk_fill(str(val)))

            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)

            # NEW: image cell support \u2014 if val is a dict with 'image_bytes',
            # render the JPEG/PNG inline at the requested width. Falls back to
            # empty if there are no bytes.
            if isinstance(val, dict) and 'image_bytes' in val:
                img_bytes = val.get('image_bytes')
                img_w = float(val.get('width_in', 0.55))
                if img_bytes:
                    try:
                        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        cell.paragraphs[0].add_run().add_picture(BytesIO(img_bytes), width=Inches(img_w))
                    except Exception as exc:
                        run(p, f'[img err: {exc.__class__.__name__}]', size=font_size, color='C00000')
                # If no image bytes, leave the cell visually empty
                continue

            # NEW: multi-link cell support \u2014 if val is a list of (text, url)
            # tuples, render each as its own hyperlink on its own paragraph.
            if isinstance(val, list) and val and isinstance(val[0], tuple):
                first_text, first_url = val[0]
                if first_url:
                    add_hyperlink(p, first_url, str(first_text), font_size=font_size)
                else:
                    run(p, str(first_text), size=font_size)
                for txt, url in val[1:]:
                    p_next = cell.add_paragraph()
                    p_next.paragraph_format.space_after = Pt(0)
                    if url:
                        add_hyperlink(p_next, url, str(txt), font_size=font_size)
                    else:
                        run(p_next, str(txt), size=font_size)
                continue

            if c_idx in hyperlink_col_indexes:
                url = hyperlink_col_indexes[c_idx](row)
                if url:
                    add_hyperlink(p, url, str(val), font_size=font_size)
                else:
                    run(p, str(val), size=font_size)
            else:
                txt = str(val) if val is not None else ''
                # Split on \n into multiple paragraphs
                lines = txt.split('\n')
                run(p, lines[0], size=font_size,
                    bold=(c_idx == risk_col_index))
                for line in lines[1:]:
                    p2 = cell.add_paragraph()
                    p2.paragraph_format.space_after = Pt(0)
                    run(p2, line, size=font_size)

    return table


# ---------- main builder ----------

def build_step5_report(*, order_meta: dict,
                       trademarks_live: list[dict],
                       trademarks_dead: list[dict],
                       companies: list[dict],
                       google: list[dict],
                       domains: list[dict],
                       social: list[dict],
                       raw_counts: dict) -> bytes:
    """Generate the Step 5 Braudit-style monitoring report as docx bytes."""
    doc = Document()

    # Set default font + margins
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # ---- Header / Title ----
    add_para(doc, order_meta.get('client_name', ''), bold=True, color='1F3864', size=16, space_after=2)
    add_para(doc, 'Monitoring or Representation Report', bold=True, size=18, space_after=10)

    # ---- Cover / Order Detail Table ----
    # Build the "Trademark Classes" value with UKIPO standardised definitions
    class_nums = parse_classes(order_meta.get('classes', ''))
    if class_nums:
        class_lines = []
        for n in class_nums:
            heading = NICE_HEADINGS.get(n, '')
            if heading:
                class_lines.append(f"Class {n} \u2014 {heading}")
            else:
                class_lines.append(f"Class {n}")
        classes_display = "\n".join(class_lines)
    else:
        classes_display = order_meta.get('classes', '')

    # Build the "Specific Terms" value from the client's actual G&S per class
    specific_terms = order_meta.get('specific_terms') or {}
    if specific_terms:
        st_lines = []
        for n in class_nums or sorted(specific_terms.keys()):
            terms = specific_terms.get(n) or specific_terms.get(str(n)) or ''
            if terms:
                st_lines.append(f"Class {n}: {terms}")
        specific_terms_display = "\n".join(st_lines) if st_lines else '\u2014'
    else:
        specific_terms_display = '\u2014'

    cover_rows = [
        ['Prepared for', order_meta.get('client_name', '')],
        ['Brand Reference', order_meta.get('brand_reference', '') or '\u2014'],
        ['Report Reference', order_meta.get('report_reference', '') or '\u2014'],
        ['Client Contact', f"{order_meta.get('client_first','')} {order_meta.get('client_last','')}".strip()],
        ['Client Email', order_meta.get('client_email', '')],
        ['Account Manager', order_meta.get('account_manager', '')],
        ['Report Prepared By', order_meta.get('prepared_by', '')],
        ['Date of Search', order_meta.get('search_date', '')],
        ['Type of Search', order_meta.get('search_type', 'Word')],
        ['Mark', order_meta.get('mark_label', '')],
        ['Trademark Classes', classes_display],
        ['Specific Terms', specific_terms_display],
        ['SIC Code', order_meta.get('sic', '')],
        ['Nature of Business', order_meta.get('nature', '')],
        ['Designated Countries', order_meta.get('countries', '')],
        ['Filtering Rules', order_meta.get('filtering_rules', '')],
    ]
    t = doc.add_table(rows=len(cover_rows), cols=2)
    t.autofit = False
    for i, (k, v) in enumerate(cover_rows):
        for ci, txt in enumerate([k, v]):
            cell = t.rows[i].cells[ci]
            cell.text = ''
            set_cell_border(cell)
            cell.width = Inches(2.2 if ci == 0 else 5.3)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if ci == 0:
                set_cell_bg(cell, '1F3864')
                run(p, txt, bold=True, color='FFFFFF', size=10)
            else:
                # Multi-line cell support \u2014 each \n becomes a paragraph
                lines = str(txt).split('\n')
                run(p, lines[0], size=10)
                for line in lines[1:]:
                    p2 = cell.add_paragraph()
                    p2.paragraph_format.space_after = Pt(0)
                    run(p2, line, size=10)

    # ---- Section 1 ----
    add_heading(doc, '1. Report Overview', level=1)
    add_para(doc, 'This report provides a monitoring update based on searches carried out in the last 30 days. It compiles findings from multiple sources, including search engines, company registries, domain name databases, social media platforms, and trademark registries. The purpose is to identify any existing use of the name or logo, potential infringements, or similar trademarks that could conflict with the client\u2019s branding.', size=11)
    add_para(doc, 'This report is for guidance only and does not replace expert trademark advice or professional legal advice.', size=11)

    add_heading(doc, 'Results Overview', level=2)
    overview_rows = [
        ['Search Engine (Google)', str(raw_counts.get('google_raw', 0)), str(len(google))],
        ['Companies House (UK)', str(raw_counts.get('companies_raw', 0)), str(len(companies))],
        ['Domains', str(raw_counts.get('domains_raw', 0)), str(len(domains))],
        ['Social Media', str(raw_counts.get('social_raw', 0)), str(len(social))],
        ['Trademark Registers \u2014 Live', str(raw_counts.get('tm_live_raw', 0)), str(len(trademarks_live))],
        ['Trademark Registers \u2014 Dead (Negligible)', str(raw_counts.get('tm_dead_raw', 0)), str(len(trademarks_dead))],
    ]
    add_table(doc, [3.5, 1.8, 2.2],
              ['Platform', 'Total Results', 'Flagged On This Report'],
              overview_rows, font_size=10)

    # ---- Section 2 ----
    add_heading(doc, '2. Search Criteria', level=1)
    add_para(doc, 'Word Search', bold=True)
    add_para(doc, f"Exact Match: {order_meta.get('exact','')}", space_after=2)
    add_para(doc, f"Similar Match: {order_meta.get('similar','')}", space_after=2)
    add_para(doc, 'Starts With: \u2014', space_after=8)

    add_para(doc, 'Exclusion Rules Applied', bold=True)
    add_para(doc, 'Trademarks: marks where the search root is not the leading word were dropped (e.g. BAY STEALTH, SUPER STEALTH, SPOTLIGHT STEALTH). Records that did not touch any of the client\u2019s classes were dropped. Dead/ended records were retained but tagged Negligible Risk.', size=10)
    add_para(doc, 'Companies House: records were retained only where the registered SIC included the client\u2019s SIC.', size=10)
    add_para(doc, 'Domains / Google / Social: the raw scrape contained one aggregated row per mark variant, so no further filtering was required.', size=10)

    # ---- Section 3 ----
    doc.add_page_break()
    add_heading(doc, '3. Search Results, Risk Grading and Analysis', level=1)

    # 3a Google
    add_heading(doc, '3a. Google Search Results', level=2)
    add_para(doc, 'Word Searches use Google\u2019s matching criteria. Image Searches use Google Lens to identify similar images online.', size=10)
    if google:
        g_body = [[g['keyword'], f"{g['score']}", g['risk'], g['link']] for g in google]
        add_table(doc, [2.4, 1.3, 1.7, 2.4],
                  ['Keyword', 'Score', 'Risk', 'Link'],
                  g_body, risk_col_index=2,
                  hyperlink_col_indexes={3: lambda row: row[3]},
                  font_size=10)
    else:
        add_para(doc, 'No Google results flagged.', italic=True, color='595959')

    # 3b Companies
    add_heading(doc, '3b. Companies House UK Search Results', level=2)
    add_para(doc, 'Company name search using UK Companies House. Records retained only where the SIC included client SIC. Company name and company number link directly to the Companies House register.', size=10)
    if companies:
        # Build a {number: link} map so both Name and Co. No. cells can use it
        co_link_by_number = {c['number']: c.get('link') or f"https://find-and-update.company-information.service.gov.uk/company/{c['number']}" for c in companies}
        c_body = [[c['mark'], c['status'], c['number'], c['sic'], c['risk']] for c in companies]
        add_table(doc, [2.4, 1.0, 1.4, 1.7, 1.3],
                  ['Registered Company', 'Status', 'Co. No.', 'SIC', 'Risk'],
                  c_body, risk_col_index=4,
                  hyperlink_col_indexes={
                      0: lambda row: co_link_by_number.get(row[2], ''),  # mark name -> CH page
                      2: lambda row: co_link_by_number.get(row[2], ''),  # co. no -> CH page
                  },
                  font_size=10)
    else:
        add_para(doc, 'No matching companies flagged.', italic=True, color='595959')

    # 3c Domains
    add_heading(doc, '3c. Domain Name Search Results', level=2)
    add_para(doc, 'Domain searches using variations of the client\u2019s mark across .com, .net, .co.uk, .co and .uk.', size=10)
    if domains:
        # Build domain cells as list[(text, url)] so each URL is a clickable link
        d_body = [
            [d['mark_text'],
             [(u, u) for u in d['urls']],
             f"{d['risk']} ({d['score']})"]
            for d in domains
        ]
        add_table(doc, [2.0, 4.4, 1.4],
                  ['Mark Variant', 'Domains', 'Risk (Score)'], d_body,
                  risk_col_index=2, font_size=10)
    else:
        add_para(doc, 'No domain results flagged.', italic=True, color='595959')

    # 3d Social
    add_heading(doc, '3d. Social Media Search Results', level=2)
    add_para(doc, 'Social media searches across Facebook, Instagram, LinkedIn, TikTok, YouTube and X. Each platform name is a clickable link to that account.', size=10)
    if social:
        s_body = []
        for s in social:
            # Build platform cells as list[(label, url)] so each platform name
            # is a clickable hyperlink straight to that profile/page.
            plats = [(k, v) for k, v in s['platforms'].items() if v]
            s_body.append([s['mark_text'], plats, f"{s['risk']} ({s['score']})"])
        add_table(doc, [2.0, 4.4, 1.4],
                  ['Mark Variant', 'Platforms', 'Risk (Score)'], s_body,
                  risk_col_index=2, font_size=9)
    else:
        add_para(doc, 'No social results flagged.', italic=True, color='595959')

    # Helper: build the per-row Class cell value as "11 \u2014 heading\n12 \u2014 heading"
    def _class_cell(cls_str):
        nums = parse_classes(cls_str)
        if not nums:
            return str(cls_str or '')
        lines = []
        for n in nums:
            h = NICE_HEADINGS.get(n, '')
            if h:
                lines.append(f"{n} \u2014 {h}")
            else:
                lines.append(str(n))
        return "\n".join(lines)

    # 3e Trademarks Live
    doc.add_page_break()
    add_heading(doc, '3e. Trademark Search Results \u2014 Live', level=1)
    add_para(doc, 'Live trademark records (Registered and Pending) where the mark is in scope and the recital touches one or more of the client\u2019s classes. Class definitions are the UKIPO standardised Nice Classification headings. Sorted by initial-review score descending.', size=10)
    # Shorten the "Stylized characters" Mark Type label so it fits in the Type
    # column on one line. Other values ("Word", "Combined") already fit.
    def _short_type(t):
        s = str(t or '').strip()
        if 'stylized' in s.lower():
            return 'Stylized'
        return s

    # Column widths re-balanced again with Image column inserted after Mark Text.
    # Total: 0.4 + 0.65 + 0.95 + 0.75 + 1.25 + 0.55 + 1.05 + 0.7 + 0.7 = 7.0 inches
    tm_widths  = [0.4,    0.65,    0.95,        0.75,    1.25,                       0.55,   1.05,    0.7,       0.7]
    tm_headers = ['Office','App #', 'Mark Text', 'Image', 'Class & UKIPO Definition', 'Type', 'Owner', 'Status', 'Risk']

    # Office-aware hyperlink for the App # column (index 1 now)
    # We pull the office from column 0 of the row so each link goes to the
    # correct register (USPTO TSDR / UKIPO / EUIPO / WIPO).
    def _app_link(row):
        return tm_url(row[0], row[1])

    def _img_cell(t):
        # Render the mark image inline at ~0.6" wide; empty cell if no image.
        return {'image_bytes': t.get('image_bytes'), 'width_in': 0.6}

    if trademarks_live:
        tl_body = [[t['office'], t['app'], t['mark'], _img_cell(t), _class_cell(t['classes']), _short_type(t['type']), t['owner'], t['status'], t['risk']] for t in trademarks_live]
        add_table(doc, tm_widths, tm_headers,
                  tl_body, risk_col_index=8,
                  hyperlink_col_indexes={1: _app_link},
                  font_size=7)
    else:
        add_para(doc, 'No live trademarks flagged.', italic=True, color='595959')

    # 3f Trademarks Dead
    doc.add_page_break()
    add_heading(doc, '3f. Trademark Search Results \u2014 Dead (Negligible Risk)', level=1)
    add_para(doc, 'Trademark records with status \u201cEnded\u201d. These have no enforceable rights and would not, on their own, support an opposition or refusal. Retained for completeness and for audit of the search sweep.', size=10)
    if trademarks_dead:
        td_body = [[t['office'], t['app'], t['mark'], _img_cell(t), _class_cell(t['classes']), _short_type(t['type']), t['owner'], t['status'], t['risk']] for t in trademarks_dead]
        add_table(doc, tm_widths, tm_headers,
                  td_body, risk_col_index=8,
                  hyperlink_col_indexes={1: _app_link},
                  font_size=7)

    # ---- Footer message ----
    doc.add_page_break()
    add_heading(doc, 'A message from our founder', level=1)
    add_para(doc, 'Thank you for taking the time to review this report.')
    add_para(doc, 'We have been helping businesses protect their brands since 2008, and if anything within this document needs clarification or gives you cause for concern, please do not hesitate to contact us.')
    add_para(doc, 'If you require full representation or only assistance with a particular stage, our team is ready to support you.')
    add_para(doc, 'We continually look to improve our service, so any feedback you can provide is always welcome.')
    add_para(doc, ' ')
    add_para(doc, 'Jonathan Paton', bold=True)
    add_para(doc, 'Founder and Director')
    add_para(doc, 'The Trademark Helpline')

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
