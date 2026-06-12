"""Braudit Audit Tool — Streamlit MVP for Steps 2–5.

Internal tool for The Trademark Helpline / Braudit. Run with:
    streamlit run app.py
"""
from __future__ import annotations
import os
import tempfile
import hashlib
import streamlit as st
from datetime import date

from pipeline.filters import (
    read_sheets, process_trademarks, process_companies,
    process_google, process_domains, process_social,
    extract_specific_terms, extract_order_metadata,
    extract_trademark_images, extract_google_image_cells,
    report_type_from_meta,
)
from pipeline.report_builder import build_step5_report
from pipeline.forensic import SignaClient, verify_records
from pipeline.temmy import TemmyClient
from pipeline.forensic_narrative import (
    NarrativeClient, ReportType, run_forensic_layer,
)
from pipeline.forensic_report import build_forensic_appendix
from pipeline.jurisdictions import (
    all_labels as all_jurisdiction_labels,
    labels_for_codes as juris_labels_for_codes,
    codes_for_labels as juris_codes_for_labels,
    parse_country_string,
)


# ---------- page setup ----------

st.set_page_config(
    page_title='Braudit Audit Tool',
    page_icon='🔍',
    layout='centered',
)


# ---------- password gate ----------

def check_password():
    """Simple password gate using Streamlit secrets."""
    if 'auth_ok' in st.session_state and st.session_state.auth_ok:
        return True

    st.title('🔍 Braudit Audit Tool')
    st.caption('Internal tool — staff access only')

    pwd = st.text_input('Password', type='password', help='Ask your account manager for the shared password.')
    if st.button('Sign in'):
        expected = st.secrets.get('app_password', os.environ.get('APP_PASSWORD', 'braudit-dev'))
        if pwd == expected:
            st.session_state.auth_ok = True
            st.rerun()
        else:
            st.error('Incorrect password.')
    return False


if not check_password():
    st.stop()


# ---------- main UI ----------

st.title('🔍 Braudit Audit Tool')
st.caption('Steps 2–5 of the audit pipeline · de-duplicate, exclude, score, generate report.')
st.divider()

# Step 1: upload + auto-parse the Order Form so the form fields can pre-fill.
# The uploader sits OUTSIDE the form so Streamlit reruns on file change.
st.subheader('1. Upload scraped-results spreadsheet')
uploaded = st.file_uploader(
    'Upload the scraped-results .xlsx file *',
    type=['xlsx'],
    help='The workbook produced by the Braudit scrape job. The app reads the Order Form sheet and pre-fills the search criteria below.',
)

# Persist the uploaded file path + parsed metadata across reruns so the form
# fields stay populated as the operator types into the manual fields.
def _ingest_upload(uploaded_file):
    raw = uploaded_file.getvalue()
    digest = hashlib.sha256(raw).hexdigest()
    if st.session_state.get('uploaded_digest') == digest:
        return  # already ingested this file
    # Save the bytes to a temp file we can re-read on each rerun
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    tmp.write(raw); tmp.close()
    st.session_state['uploaded_digest'] = digest
    st.session_state['uploaded_path'] = tmp.name
    st.session_state['uploaded_name'] = uploaded_file.name
    meta = extract_order_metadata(tmp.name)
    st.session_state['order_meta'] = meta
    st.session_state['specific_terms'] = extract_specific_terms(tmp.name)

    # ----- Push captured search criteria into form-widget session_state -----
    # The form widgets in _search_block use both `value=` AND `key=`. Streamlit's
    # documented behaviour is: on subsequent renders the session_state value
    # (under the widget key) wins over `value=`. So if the form is rendered ONCE
    # with empty defaults (e.g. before a file is uploaded), session_state[key]
    # is locked to '' and a later `value=` from a freshly-parsed upload is
    # silently ignored. The fix is to explicitly overwrite those keys here
    # the moment we know what the spreadsheet contains.
    for i in range(5):
        ws = (meta.get('word_searches') or [])[i] if i < len(meta.get('word_searches') or []) else None
        st.session_state[f'wordsearch_type_{i}'] = (ws or {}).get('type', 'Exact Match')
        st.session_state[f'wordsearch_phrase_{i}'] = (ws or {}).get('phrase', '')
        st.session_state[f'wordsearch_remarks_{i}'] = (ws or {}).get('remarks', '')
        ds = (meta.get('domain_searches') or [])[i] if i < len(meta.get('domain_searches') or []) else None
        st.session_state[f'domsearch_type_{i}'] = (ds or {}).get('type', 'Exact Match')
        st.session_state[f'domsearch_phrase_{i}'] = (ds or {}).get('phrase', '')
        st.session_state[f'domsearch_remarks_{i}'] = (ds or {}).get('remarks', '')

if uploaded is not None:
    _ingest_upload(uploaded)
    meta = st.session_state.get('order_meta') or {}
    if meta.get('client_name'):
        # Build a concise summary line covering everything the parser now captures.
        word_n = len(meta.get('word_searches') or [])
        domain_n = len(meta.get('domain_searches') or [])
        # 'NO SEARCH' is a legitimate sentinel meaning the slot is intentionally
        # unused. Count only filenames as actual image searches.
        def _is_real_image(v: str) -> bool:
            v = (v or '').strip()
            return bool(v) and v.upper() != 'NO SEARCH'
        image_n = sum(1 for k in ('image_1', 'image_2') if _is_real_image(meta.get(k, '')))
        vienna = meta.get('vienna_classes', '')
        st.success(
            f"📋 Order Form parsed: **{meta['client_name']}** · classes {meta.get('classes_csv','—')} "
            f"· {word_n} word search{'es' if word_n != 1 else ''}, "
            f"{domain_n} domain search{'es' if domain_n != 1 else ''}, "
            f"{image_n} image search{'es' if image_n != 1 else ''}"
            + (f' · Vienna: {vienna}' if vienna else '')
        )

        # Expandable preview of every captured search criterion. Read-only —
        # it shows what landed from the spreadsheet so the operator can verify
        # before running the audit. Edits still happen in the form below.
        with st.expander('📑 Show all captured search criteria from the Order Form'):
            st.markdown('**Word searches (Order Form rows 14–18)**')
            ws_list = meta.get('word_searches') or []
            if ws_list:
                st.dataframe(
                    [{'Type': w['type'], 'Phrase': w['phrase'], 'Remarks': w['remarks']}
                     for w in ws_list],
                    use_container_width=True, hide_index=True,
                )
            else:
                st.caption('_No word searches captured._')

            st.markdown('**Domain searches (Order Form rows 22–26)**')
            ds_list = meta.get('domain_searches') or []
            if ds_list:
                st.dataframe(
                    [{'Type': d['type'], 'Phrase': d['phrase'], 'Remarks': d['remarks']}
                     for d in ds_list],
                    use_container_width=True, hide_index=True,
                )
            else:
                st.caption('_No domain searches captured._')

            st.markdown('**Image searches (Order Form rows 30–31)**')
            img_rows = [
                {
                    'Slot': 'Image Mark 1',
                    'Value': meta.get('image_1', '') or '—',
                    'Embedded': (
                        f'✓ {meta.get("image_1_format","").upper()} '
                        f'({len(meta.get("image_1_bytes") or b"")} bytes)'
                        if meta.get('image_1_bytes') else '—'
                    ),
                },
                {
                    'Slot': 'Image Mark 2',
                    'Value': meta.get('image_2', '') or '—',
                    'Embedded': (
                        f'✓ {meta.get("image_2_format","").upper()} '
                        f'({len(meta.get("image_2_bytes") or b"")} bytes)'
                        if meta.get('image_2_bytes') else '—'
                    ),
                },
            ]
            st.dataframe(img_rows, use_container_width=True, hide_index=True)
            # Visual preview of embedded image marks so the operator can
            # confirm the right logo was extracted from the spreadsheet.
            img1_bytes = meta.get('image_1_bytes')
            img2_bytes = meta.get('image_2_bytes')
            if img1_bytes or img2_bytes:
                preview_cols = st.columns(2)
                if img1_bytes:
                    with preview_cols[0]:
                        st.image(img1_bytes, caption='Image Mark 1 (B30)', width=240)
                if img2_bytes:
                    with preview_cols[1]:
                        st.image(img2_bytes, caption='Image Mark 2 (B31)', width=240)
            if vienna:
                st.markdown(f'**Vienna classes (shared):** `{vienna}`')
            else:
                st.caption('_No Vienna codes captured yet (column C of the image rows). Operator can add them on the Order Form._')
    else:
        st.warning('Uploaded — but the Order Form sheet was not found or could not be read. Fill the fields below manually.')
else:
    st.info('Upload a scraped-results spreadsheet to pre-fill the search criteria below.')

st.divider()

# All the parsed defaults flow into the form via session_state.
meta = st.session_state.get('order_meta') or {}

with st.form('audit_form', clear_on_submit=False):

    st.subheader('2. Audit operator details')
    st.caption('Pre-filled from the spreadsheet (Order Form rows 58–64). Edit any field to override for this audit only. The two reference fields are optional — they don’t affect the audit but flow through to the report header.')
    # Optional reference fields (rows 58/59) shown first
    c0a, c0b = st.columns(2)
    with c0a:
        brand_reference = st.text_input(
            'Brand reference',
            value=meta.get('brand_reference', ''),
            help='Optional. From Order Form R58. Usually the word mark text.',
        )
    with c0b:
        report_reference = st.text_input(
            'Report reference',
            value=meta.get('report_reference', ''),
            help='Optional. From Order Form R59. Usually the Deal Name.',
        )
    # Required operator fields (rows 60–64)
    c1, c2 = st.columns(2)
    with c1:
        client_first = st.text_input(
            'Client first name *',
            value=meta.get('client_first', ''),
            help='From Order Form R60.',
        )
        client_email = st.text_input(
            'Client email address *',
            value=meta.get('client_email', ''),
            help='From Order Form R62.',
        )
        account_manager = st.text_input(
            'Account manager *',
            value=meta.get('account_manager', ''),
            help='From Order Form R63.',
        )
    with c2:
        client_last = st.text_input(
            'Client last name *',
            value=meta.get('client_last', ''),
            help='From Order Form R61.',
        )
        prepared_by = st.text_input(
            'Report prepared by *',
            value=meta.get('prepared_by', ''),
            help='From Order Form R64.',
        )

    st.subheader('3. Search criteria from your Order Form')
    st.caption('Pre-filled from the spreadsheet. Edit any field below to override the value for this audit only.')
    c3, c4 = st.columns(2)
    with c3:
        client_name = st.text_input('Client / company name *', value=meta.get('client_name', ''), help='From Order Form R5.')
        classes_text = st.text_input('Trademark classes (comma-separated) *', value=meta.get('classes_csv', ''), help='From Order Form G&S Classes rows.')
        sic_code = st.text_input('Client SIC code *', value=meta.get('sic', ''), help='From Order Form R8.')
    with c4:
        deal_id = st.text_input('Deal ID', value=meta.get('deal_id', ''), help='From Order Form R6.')
        countries = st.text_input('Designated countries', value=meta.get('countries', ''), help='From Order Form R10.')
        nature = st.text_input('Nature of business', value=meta.get('nature', ''), help='From Order Form R9.')

    # ----- Word & Domain search criteria (multi-row) -----
    # Up to 5 rows per block, each with a type (Exact Match / Starts With /
    # Contains / Similar To) + phrase + optional remarks. Pre-populated from
    # the Order Form parser (extract_order_metadata captures word_searches /
    # domain_searches). The operator can edit any row in-place; empty phrase
    # rows are dropped on submit so blank slots don't pollute the report.
    SEARCH_TYPES = ['Exact Match', 'Starts With', 'Contains', 'Similar To']

    def _search_block(title: str, caption: str, key_prefix: str,
                       meta_key: str, n_rows: int = 5):
        """Render a 5-row block of (type, phrase, remarks). Returns a list of
        dicts in the same shape the parser produces, only including rows
        whose phrase is non-empty."""
        st.markdown(f'**{title}**')
        st.caption(caption)
        # Tight header row
        h0, h1, h2 = st.columns([1.4, 3.0, 2.4])
        with h0: st.markdown('Type')
        with h1: st.markdown('Phrase')
        with h2: st.markdown('Remarks (optional)')

        captured = meta.get(meta_key) or []
        rows: list[dict] = []
        for i in range(n_rows):
            default = captured[i] if i < len(captured) else {'type': 'Exact Match', 'phrase': '', 'remarks': ''}
            try:
                type_idx = SEARCH_TYPES.index(default.get('type', 'Exact Match'))
            except ValueError:
                type_idx = 0
            r0, r1, r2 = st.columns([1.4, 3.0, 2.4])
            with r0:
                t = st.selectbox(
                    f'{key_prefix}_type_{i}', SEARCH_TYPES, index=type_idx,
                    key=f'{key_prefix}_type_{i}', label_visibility='collapsed',
                )
            with r1:
                p = st.text_input(
                    f'{key_prefix}_phrase_{i}', value=default.get('phrase', ''),
                    key=f'{key_prefix}_phrase_{i}', label_visibility='collapsed',
                    placeholder=f'phrase {i+1}',
                )
            with r2:
                rmk = st.text_input(
                    f'{key_prefix}_remarks_{i}', value=default.get('remarks', ''),
                    key=f'{key_prefix}_remarks_{i}', label_visibility='collapsed',
                    placeholder='optional remark',
                )
            if (p or '').strip():
                rows.append({'type': t, 'phrase': p.strip(), 'remarks': (rmk or '').strip()})
        return rows

    st.subheader('3a. Word Search Criteria *')
    word_searches_input = _search_block(
        title='Up to 5 word search phrases (at least 1 required)',
        caption='Pre-filled from Order Form rows 14–18. Edit any row to override; leave a row blank to drop it.',
        key_prefix='wordsearch', meta_key='word_searches',
    )

    st.subheader('3b. Domain Search Criteria')
    domain_searches_input = _search_block(
        title='Up to 5 domain search phrases',
        caption='Pre-filled from Order Form rows 22–26. Edit any row to override; leave a row blank to drop it.',
        key_prefix='domsearch', meta_key='domain_searches',
    )

    # Back-compat single-value fields for the few places downstream that
    # still expect them — the first Exact Match (or first phrase if no Exact)
    # for "exact", and the first Starts With for "similar".
    exact = ''
    similar = ''
    for w in word_searches_input:
        if not exact and w['type'].lower() == 'exact match':
            exact = w['phrase']
        if not similar and w['type'].lower() == 'starts with':
            similar = w['phrase']
    if not exact and word_searches_input:
        # Fall back to the first phrase (whatever its type) so downstream
        # report headers don't print blank.
        exact = word_searches_input[0]['phrase']

    search_platforms = st.text_area(
        'Search platforms',
        value=meta.get('search_platforms', ''),
        help='From Order Form R11 — comma-separated list of platforms the scrape covered.',
        height=68,
    )

    submitted = st.form_submit_button('▶ Run Audit', type='primary', use_container_width=True)


# ---------- processing ----------

if submitted:
    missing = []
    for label, val in [
        ('Client first name', client_first), ('Client last name', client_last),
        ('Client email', client_email), ('Account manager', account_manager),
        ('Report prepared by', prepared_by), ('Client / company name', client_name),
        ('Trademark classes', classes_text),
        ('Client SIC code', sic_code),
    ]:
        if not val.strip():
            missing.append(label)
    if uploaded is None or 'uploaded_path' not in st.session_state:
        missing.append('Scraped data .xlsx file')
    # At least one word search phrase is required. This replaces the old
    # "Exact match search term *" required field, since the form now accepts
    # any of the four search types (Exact / Starts With / Contains / Similar To).
    if not word_searches_input:
        missing.append('At least one Word Search phrase (3a)')

    if missing:
        st.error('Please complete: ' + ', '.join(missing))
        st.stop()

    tmp_path = st.session_state['uploaded_path']
    specific_terms = st.session_state.get('specific_terms') or {}

    try:
        with st.spinner('Reading scraped data...'):
            sheets = read_sheets(tmp_path)
            tm_images = extract_trademark_images(tmp_path)

        # BR-011 (11 Jun 2026) — Monitoring Report support. When the Order
        # Form A2 cell says "Monitoring or Representation Report", filter
        # every data sheet to rows whose Search Date equals the latest in
        # that sheet (i.e. drop earlier monitoring periods). The latest
        # date is also surfaced to order_meta so the report cover can show
        # it in the Section 1 monitoring intro.
        is_monitoring = (meta.get('document_type') or 'audit').lower() == 'monitoring'
        latest_search_date = ''
        if is_monitoring:
            from pipeline.filters import filter_to_latest_search_date
            for sheet_name in ('Google', 'Companies', 'Domains', 'Social', 'Trademarks'):
                raw = sheets.get(sheet_name, [[]])
                filtered, latest = filter_to_latest_search_date(raw)
                sheets[sheet_name] = filtered
                if latest and latest > latest_search_date:
                    latest_search_date = latest
            st.info(
                f'📅 **Monitoring report mode** — data filtered to latest scrape '
                f'period only{": " + latest_search_date if latest_search_date else ""}. '
                f'Earlier monitoring periods are excluded.'
            )

        target_classes = tuple(
            int(x.strip()) for x in classes_text.split(',') if x.strip().isdigit()
        ) or (11, 12, 35)

        root_word = exact.strip().split()[0].upper() if exact.strip() else 'STEALTH'

        with st.spinner('Step 2: de-duplicating and applying exclusions...'):
            # The operator's edited word/domain search rows now drive the
            # filter+scoring decisions. Trademarks and companies are
            # filtered against word_searches_input with type-aware matching
            # (Exact Match / Starts With / Contains / Similar To). Domains
            # filter against domain_searches_input the same way. The
            # legacy single-root behaviour still applies when no rows are
            # supplied (back-compat for any caller that hasn't moved over).
            # Tri-state report type drives dual scoring + duplicate-row
            # behaviour in process_trademarks. Combined reports get every
            # record scored on BOTH word and image axes; duplicate rows
            # emerge when both axes are non-Negligible.
            report_type = report_type_from_meta(meta)
            tm_all = process_trademarks(
                sheets.get('Trademarks', [[]]),
                target_classes=target_classes,
                root=root_word,
                images=tm_images,
                word_searches=word_searches_input,
                report_type=report_type,
                client_vienna=meta.get('vienna_classes', ''),
                # BR-IMG-002 (Phase 2): client logo for visual similarity.
                # Triggers CLIP visual comparison against every cited mark
                # logo on Combined / Image reports. On first-ever audit on
                # a fresh Streamlit container the OpenCLIP model downloads
                # (~340 MB) which adds ~90 sec; subsequent audits reuse
                # the cached weights and add ~10–15 sec per audit.
                client_image_bytes=meta.get('image_1_bytes'),
                # BR-IMG-003 (Phase 2.1): client name used to detect when
                # the cited mark's owner is the client themselves — the
                # row is tagged 'Client Likely' instead of being scored
                # as a third-party threat.
                client_name=client_name,
            )
            companies = process_companies(
                sheets.get('Companies', [[]]),
                target_sic=sic_code.strip(),
                root=root_word,
                word_searches=word_searches_input,
            )
            # Picture-in-Cell images on the Google sheet (column A) — modern
            # Excel rich-values that openpyxl reports as #VALUE!. Extract them
            # so the report can render the actual image instead of an error.
            google_imgs = extract_google_image_cells(tmp_path)
            google = process_google(sheets.get('Google', [[]]), images=google_imgs)
            domains = process_domains(
                sheets.get('Domains', [[]]),
                domain_searches=domain_searches_input,
            )
            social = process_social(sheets.get('Social', [[]]))

        tm_live = [t for t in tm_all if t['risk'] != 'Negligible']
        tm_dead = [t for t in tm_all if t['risk'] == 'Negligible']

        with st.spinner('Step 5: building the report...'):
            search_date_value = (meta.get('search_date') or '').strip() or date.today().strftime('%d %B %Y')
            order_meta = {
                'client_name': client_name,
                'client_first': client_first,
                'client_last': client_last,
                'client_email': client_email,
                'account_manager': account_manager,
                'prepared_by': prepared_by,
                'deal_id': deal_id,
                'brand_reference': brand_reference,
                'report_reference': report_reference,
                'search_date': search_date_value,
                'search_type': 'Word',
                'mark_label': f'Exact: {exact}   ·   Similar: {similar}' if similar else f'Exact: {exact}',
                'exact': exact, 'similar': similar,
                # NEW: structured search criteria — used by report_builder
                # section 2 to render all 5 word + 5 domain rows. Back-compat
                # 'exact' / 'similar' single-value fields are still set above
                # for any caller that hasn't moved over yet.
                'word_searches': word_searches_input,
                'domain_searches': domain_searches_input,
                # Pass through the Word Mark / Image Mark / Vienna fields the
                # parser captured, so the cover panel can show them.
                'word_or_image': meta.get('word_or_image', ''),
                'image_1': meta.get('image_1', ''),
                'image_1_bytes': meta.get('image_1_bytes'),
                'image_1_format': meta.get('image_1_format', ''),
                'image_2': meta.get('image_2', ''),
                'image_2_bytes': meta.get('image_2_bytes'),
                'image_2_format': meta.get('image_2_format', ''),
                'vienna_classes': meta.get('vienna_classes', ''),
                # NEW (10 Jun 2026): structured Vienna classifications from
                # D31:D40 + E31:E40 of the Order Form. The report renders
                # these as a small (code, description) table on the cover
                # for Image and Combined reports.
                'vienna_classifications': meta.get('vienna_classifications', []),
                # BR-011 (11 Jun 2026): monitoring vs audit deliverable.
                # Detected via Order Form A2. Drives title prefix + Section
                # 1 narrative + skips Recommended Spec / Actions in
                # forensic appendix. latest_search_date is computed from
                # the data sheets and rendered in the monitoring intro.
                'document_type': meta.get('document_type', 'audit'),
                'latest_search_date': latest_search_date,
                'classes': classes_text,
                'sic': sic_code,
                'nature': nature,
                'countries': countries,
                'search_platforms': search_platforms,
                'filtering_rules': (
                    f'Mark scope: exact {root_word} or "{root_word} " + descriptor; '
                    f'class touch any of {classes_text}; '
                    f'SIC = {sic_code} (companies); '
                    'dead trademarks retained but tagged Negligible.'
                ),
                'specific_terms': specific_terms,
            }
            raw_counts = {
                'google_raw': max(0, len(sheets.get('Google', [[]])) - 1),
                'companies_raw': max(0, len(sheets.get('Companies', [[]])) - 1),
                'domains_raw': max(0, len(sheets.get('Domains', [[]])) - 1),
                'social_raw': max(0, len(sheets.get('Social', [[]])) - 1),
                'tm_live_raw': len(tm_live),
                'tm_dead_raw': len(tm_dead),
            }

            docx_bytes = build_step5_report(
                order_meta=order_meta,
                trademarks_live=tm_live,
                trademarks_dead=tm_dead,
                companies=companies,
                google=google,
                domains=domains,
                social=social,
                raw_counts=raw_counts,
            )

        # Stash the audit result so it survives session reruns triggered by
        # the optional Step 6 forensic-layer button below. Without this,
        # clicking "Run Forensic Audit" would lose tm_live etc.
        st.session_state['step5_result'] = {
            'order_meta': order_meta,
            'tm_live': tm_live,
            'tm_dead': tm_dead,
            'companies_count': len(companies),
            'domains_count': len(domains),
            'google_count': len(google),
            'social_count': len(social),
            'classes_text': classes_text,
            'exact': exact,
            'countries': countries,
            'client_name': client_name,
            'docx_bytes': docx_bytes,
        }
        # Clear any prior forensic output so the page doesn't show a stale
        # forensic appendix from a previous spreadsheet upload.
        st.session_state.pop('forensic_result', None)

        st.success('✅ Audit complete.')

    except Exception as e:
        st.error(f'Audit failed: {e}')
        raise


# ---------- step 5 display (persistent across reruns) ----------

step5 = st.session_state.get('step5_result')
if step5:
    st.subheader('Summary of flagged results')
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric('Live Trademarks', len(step5['tm_live']))
    m2.metric('Dead Trademarks', len(step5['tm_dead']))
    m3.metric('Companies', step5['companies_count'])
    m4.metric('Domains', step5['domains_count'])
    m5.metric('Google + Social', step5['google_count'] + step5['social_count'])

    from collections import Counter
    risk_dist = Counter(t['risk'] for t in step5['tm_live'])
    st.write('**Live trademark risk distribution:** ' + ', '.join(
        f'{k}: {v}' for k, v in risk_dist.most_common()
    ) if risk_dist else '_no live trademarks flagged_')

    st.divider()
    st.subheader('Download the report')
    safe_client = ''.join(c if c.isalnum() else '_' for c in step5['client_name'])[:40]
    # BR-011: distinguish Monitoring vs Audit in the download filename so the
    # operator can tell them apart when filing both for the same client.
    _doc_type = (step5['order_meta'].get('document_type') or 'audit').lower()
    _prefix = 'Braudit Monitoring Report' if _doc_type == 'monitoring' else 'Braudit Audit Report'
    filename = f'{_prefix} – {safe_client} – {date.today().isoformat()}.docx'
    st.download_button(
        label='⬇ Download Word report',
        data=step5['docx_bytes'],
        file_name=filename,
        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        use_container_width=True,
    )

    with st.expander('Preview — Live trademarks'):
        st.dataframe([
            {k: v for k, v in t.items() if k in ('app','mark','classes','type','owner','status','score','risk')}
            for t in step5['tm_live']
        ], use_container_width=True)
    if step5['tm_dead']:
        with st.expander(f'Preview — Dead trademarks ({len(step5["tm_dead"])})'):
            st.dataframe([
                {k: v for k, v in t.items() if k in ('app','mark','classes','type','owner','status','risk')}
                for t in step5['tm_dead']
            ], use_container_width=True)


# ---------- step 6: forensic layer (optional) ----------

if step5:
    st.divider()
    st.subheader('6. Add Forensic Layer (optional)')
    st.caption(
        'Verify each cited trademark against the Signa unified register API and '
        'generate forensic commentary via Anthropic Sonnet 4.6. Produces a separate '
        'Word appendix to accompany the Braudit monitoring report above.'
    )

    has_signa = bool(st.secrets.get('signa_api_key', ''))
    has_anthropic = bool(st.secrets.get('anthropic_api_key', ''))
    if not (has_signa and has_anthropic):
        missing = []
        if not has_signa: missing.append('`signa_api_key`')
        if not has_anthropic: missing.append('`anthropic_api_key`')
        st.warning(
            'Forensic layer requires the following keys in Streamlit Secrets: '
            + ', '.join(missing)
            + '. Ask the admin to add them in **Manage app → Secrets** before running.'
        )

    f_c1, f_c2 = st.columns([1, 1])
    with f_c1:
        report_type_label = st.radio(
            'Report type',
            options=['Pre-Application — likelihood of objection',
                     'Post-Registration — potential infringements'],
            index=0,
            help='Pre-Application: client is filing a new mark and wants to know which senior rights might block. '
                 'Post-Registration: client already has a registered mark and wants to identify potential infringers.',
        )
    with f_c2:
        # BR-010 (11 Jun 2026) — selection mode:
        #   Default: top-N live trademarks by score (existing behaviour)
        #   Custom : operator hand-picks rows from 3a / 3b / 3c / 3d / 3e / 3f
        live_count = len(step5['tm_live'])
        dead_count = len(step5['tm_dead']) if step5.get('tm_dead') else 0
        default_top_n = min(10, live_count) if live_count else 1
        selection_mode = st.radio(
            'Selection mode',
            options=['Top N trademarks by risk (default)',
                     'Custom selection across 3a–3f'],
            index=0,
            help='Default: forensically audits the top-N live trademarks by risk score. '
                 'Custom: opens per-section selectors so you can hand-pick Google hits, '
                 'Companies House records, domains, social handles, and trademarks.',
        )
        is_custom_mode = selection_mode.startswith('Custom')
        if not is_custom_mode:
            top_n = st.number_input(
                'Records to forensically audit',
                min_value=1, max_value=max(1, live_count), value=default_top_n, step=1,
                help='Top N live trademarks by score will be verified via Signa '
                     'and forensically commented. Higher N → higher API cost. '
                     '10 is a sensible default.',
            )
        else:
            top_n = 0  # ignored in custom mode

    # ---------- Custom selection UI (BR-010) ----------
    # Six per-section dataframes with multi-row selection. Defaults are empty;
    # the operator ticks rows to include. Each section's selection is read
    # back via st.session_state under a stable widget key.
    custom_selection: dict = {'tm_3e': [], 'tm_3f': [],
                              'google': [], 'companies': [],
                              'domains': [], 'social': []}
    if is_custom_mode:
        st.markdown('**Pick rows from each section to include in the forensic audit.** '
                    'Leaving a section empty means no entries from that section '
                    'will be forensically analysed. The trademarks you pick from '
                    '3e + 3f are verified against Signa / TemmyDB and get '
                    'full per-record cards. Selections from 3a–3d get LLM-generated '
                    'forensic commentary blocks in new appendix sections (6a–6d).')

        # ----- 3a Google -----
        with st.expander(f'3a Google — {len(step5.get("google") or [])} results',
                         expanded=False):
            google = step5.get('google') or []
            if google:
                df_g = [{'#': i + 1,
                         'Keyword': (r.get('keyword') or r.get('mark_text') or '')[:60],
                         'URL': (r.get('urls') or [r.get('link', '')])[0][:80],
                         'Risk': r.get('risk', ''),
                         'Score': r.get('score', '')}
                        for i, r in enumerate(google)]
                ev = st.dataframe(df_g, use_container_width=True, hide_index=True,
                                  on_select='rerun', selection_mode='multi-row',
                                  key='sel_3a_google')
                custom_selection['google'] = list(getattr(ev, 'selection', {}).get('rows', []) or [])
                st.caption(f'Selected: {len(custom_selection["google"])}')
            else:
                st.caption('_No Google results to choose from._')

        # ----- 3b Companies -----
        with st.expander(f'3b Companies House — {len(step5.get("companies") or [])} results',
                         expanded=False):
            companies = step5.get('companies') or []
            if companies:
                df_c = [{'#': i + 1,
                         'Company': (r.get('company') or r.get('name') or r.get('mark_text') or '')[:60],
                         'Status': r.get('status', ''),
                         'Co. number': r.get('company_number') or r.get('co_no') or '',
                         'SIC': r.get('sic', ''),
                         'Risk': r.get('risk', ''),
                         'Score': r.get('score', '')}
                        for i, r in enumerate(companies)]
                ev = st.dataframe(df_c, use_container_width=True, hide_index=True,
                                  on_select='rerun', selection_mode='multi-row',
                                  key='sel_3b_companies')
                custom_selection['companies'] = list(getattr(ev, 'selection', {}).get('rows', []) or [])
                st.caption(f'Selected: {len(custom_selection["companies"])}')
            else:
                st.caption('_No Companies House records to choose from._')

        # ----- 3c Domains -----
        with st.expander(f'3c Domains — {len(step5.get("domains") or [])} results',
                         expanded=False):
            domains = step5.get('domains') or []
            if domains:
                df_d = [{'#': i + 1,
                         'Mark': (r.get('mark_text') or '')[:40],
                         'Domain': ((r.get('urls') or [r.get('domain', '')])[0] if r.get('urls') or r.get('domain') else '')[:60],
                         'Risk': r.get('risk', ''),
                         'Score': r.get('score', '')}
                        for i, r in enumerate(domains)]
                ev = st.dataframe(df_d, use_container_width=True, hide_index=True,
                                  on_select='rerun', selection_mode='multi-row',
                                  key='sel_3c_domains')
                custom_selection['domains'] = list(getattr(ev, 'selection', {}).get('rows', []) or [])
                st.caption(f'Selected: {len(custom_selection["domains"])}')
            else:
                st.caption('_No domain records to choose from._')

        # ----- 3d Social -----
        with st.expander(f'3d Social media — {len(step5.get("social") or [])} results',
                         expanded=False):
            social = step5.get('social') or []
            if social:
                def _plats_str(r):
                    p = r.get('platforms') or {}
                    if isinstance(p, dict):
                        return ', '.join(k for k, v in p.items() if v)
                    return str(p)
                df_s = [{'#': i + 1,
                         'Mark': (r.get('mark_text') or '')[:40],
                         'Platforms': _plats_str(r)[:60],
                         'Risk': r.get('risk', ''),
                         'Score': r.get('score', '')}
                        for i, r in enumerate(social)]
                ev = st.dataframe(df_s, use_container_width=True, hide_index=True,
                                  on_select='rerun', selection_mode='multi-row',
                                  key='sel_3d_social')
                custom_selection['social'] = list(getattr(ev, 'selection', {}).get('rows', []) or [])
                st.caption(f'Selected: {len(custom_selection["social"])}')
            else:
                st.caption('_No social media records to choose from._')

        # ----- 3e Live trademarks -----
        sorted_live_for_pick = sorted(step5['tm_live'],
                                      key=lambda t: -int(t.get('score') or 0))
        with st.expander(f'3e Live trademarks — {len(sorted_live_for_pick)} records '
                         '(sorted by risk score, highest first)',
                         expanded=True):
            if sorted_live_for_pick:
                df_e = [{'#': i + 1,
                         'Office': r.get('office', ''),
                         'App #': r.get('app', ''),
                         'Mark': (r.get('mark') or '')[:35],
                         'Owner': (r.get('owner') or '')[:35],
                         'Classes': r.get('classes', ''),
                         'Risk': r.get('risk', ''),
                         'Score': r.get('score', '')}
                        for i, r in enumerate(sorted_live_for_pick)]
                ev = st.dataframe(df_e, use_container_width=True, hide_index=True,
                                  on_select='rerun', selection_mode='multi-row',
                                  key='sel_3e_trademarks')
                custom_selection['tm_3e'] = list(getattr(ev, 'selection', {}).get('rows', []) or [])
                st.caption(f'Selected: {len(custom_selection["tm_3e"])} of {len(sorted_live_for_pick)}')
            else:
                st.caption('_No live trademarks to choose from._')

        # ----- 3f Dead trademarks -----
        with st.expander(f'3f Dead trademarks (Negligible Risk) — {dead_count} records',
                         expanded=False):
            tm_dead_list = step5.get('tm_dead') or []
            if tm_dead_list:
                df_f = [{'#': i + 1,
                         'Office': r.get('office', ''),
                         'App #': r.get('app', ''),
                         'Mark': (r.get('mark') or '')[:35],
                         'Owner': (r.get('owner') or '')[:35],
                         'Classes': r.get('classes', '')}
                        for i, r in enumerate(tm_dead_list)]
                ev = st.dataframe(df_f, use_container_width=True, hide_index=True,
                                  on_select='rerun', selection_mode='multi-row',
                                  key='sel_3f_trademarks')
                custom_selection['tm_3f'] = list(getattr(ev, 'selection', {}).get('rows', []) or [])
                st.caption(f'Selected: {len(custom_selection["tm_3f"])} of {len(tm_dead_list)}')
            else:
                st.caption('_No dead trademarks to choose from._')

        # Summary line
        total_selected = sum(len(v) for v in custom_selection.values())
        st.info(
            f'Total selection across all sections: **{total_selected}** records · '
            f'TM 3e {len(custom_selection["tm_3e"])} · '
            f'TM 3f {len(custom_selection["tm_3f"])} · '
            f'Google {len(custom_selection["google"])} · '
            f'Companies {len(custom_selection["companies"])} · '
            f'Domains {len(custom_selection["domains"])} · '
            f'Social {len(custom_selection["social"])}'
        )

    # Multi-select jurisdictions — default to whatever we parsed out
    # of the Order Form’s 'Designated countries' field, falling back
    # to United States if parsing returns nothing.
    parsed_codes = parse_country_string(step5.get('countries', '') or '')
    if not parsed_codes:
        parsed_codes = ['US']
    default_jurisdiction_labels = juris_labels_for_codes(parsed_codes)
    selected_jurisdiction_labels = st.multiselect(
        'Client jurisdictions',
        options=all_jurisdiction_labels(),
        default=default_jurisdiction_labels,
        help='All jurisdictions the client has filed (or plans to file) in. '
             'Pre-populated from the Designated countries field above; add or '
             'remove as needed. Used by the scoring rubric for the region '
             'criterion — records in any of these jurisdictions score highest.',
    )
    client_jurisdictions = juris_codes_for_labels(selected_jurisdiction_labels)

    forensic_clicked = st.button(
        '▶ Run Forensic Audit',
        type='primary',
        use_container_width=True,
        disabled=not (has_signa and has_anthropic),
    )

    if forensic_clicked:
        try:
            # BR-010 (11 Jun 2026) — choose the verification target list:
            #   Default mode: top-N live trademarks by score (existing).
            #   Custom mode : operator-picked rows from 3e + 3f.
            sorted_live = sorted(
                step5['tm_live'],
                key=lambda t: -int(t.get('score') or 0),
            )
            tm_dead_list = step5.get('tm_dead') or []
            extras_rows: dict = {'google': [], 'companies': [],
                                  'domains': [], 'social': []}
            if is_custom_mode:
                # Combine 3e + 3f picks. The dataframe used the
                # sorted_live order, so the indices map directly.
                target_records = [sorted_live[i] for i in custom_selection['tm_3e']
                                  if 0 <= i < len(sorted_live)]
                target_records += [tm_dead_list[i] for i in custom_selection['tm_3f']
                                   if 0 <= i < len(tm_dead_list)]
                # Pull the row payloads for 3a-3d so they reach the
                # narrative layer for forensic commentary.
                google_list = step5.get('google') or []
                companies_list = step5.get('companies') or []
                domains_list = step5.get('domains') or []
                social_list = step5.get('social') or []
                extras_rows['google']    = [google_list[i]    for i in custom_selection['google']    if 0 <= i < len(google_list)]
                extras_rows['companies'] = [companies_list[i] for i in custom_selection['companies'] if 0 <= i < len(companies_list)]
                extras_rows['domains']   = [domains_list[i]   for i in custom_selection['domains']   if 0 <= i < len(domains_list)]
                extras_rows['social']    = [social_list[i]    for i in custom_selection['social']    if 0 <= i < len(social_list)]
                if not target_records and not any(extras_rows.values()):
                    st.error('Custom mode selected but no rows were ticked in any section. Please select at least one row and try again.')
                    raise RuntimeError('Empty custom selection')
            else:
                target_records = sorted_live[:int(top_n)]

            # Build clients
            signa_client = SignaClient(api_key=st.secrets['signa_api_key'])
            narrative_client = NarrativeClient(api_key=st.secrets['anthropic_api_key'])
            # TemmyDB is Braudit's source of truth for UK records (BR-002).
            # Optional: if the temmy_api_key secret is not set, UK records
            # skip the Temmy primary path and fall back to the Signa
            # Brexit-clone proxy for UK009xxxxxxxx numbers only.
            temmy_client = (
                TemmyClient(api_key=st.secrets['temmy_api_key'])
                if st.secrets.get('temmy_api_key') else None
            )

            # Resolve report type
            report_type = (ReportType.PRE_APPLICATION
                           if report_type_label.startswith('Pre-')
                           else ReportType.POST_REGISTRATION)

            # Parse client classes from the comma-separated text
            try:
                client_classes = [
                    int(x.strip()) for x in step5['classes_text'].split(',') if x.strip().isdigit()
                ]
            except Exception:
                client_classes = []

            # --- Phase 1: Signa verification (with progress bar) ---
            prog = st.progress(0.0, text='Verifying records against Signa...')
            def _cb(idx, total, office, app):
                pct = idx / max(1, total)
                prog.progress(pct, text=f'Verifying {idx}/{total} — {office} {app}')
            with st.spinner('Verifying records against Signa unified register API (UK via TemmyDB)...'):
                signa_records = verify_records(
                    signa_client, target_records,
                    progress_callback=_cb,
                    temmy_client=temmy_client,
                )
            prog.empty()

            verified_count = sum(1 for r in signa_records if r.verified)
            # Per-source breakdown for the BR-001/BR-002 fix — lets operators
            # see at a glance which records came from Signa, Temmy or the
            # Brexit-clone proxy.
            by_source: dict = {}
            for r in signa_records:
                by_source[r.verification_source] = by_source.get(r.verification_source, 0) + 1
            src_summary = ', '.join(f'{k}={v}' for k, v in sorted(by_source.items()))
            st.info(
                f'Verification: {verified_count} of {len(signa_records)} records confirmed '
                f'({src_summary}).'
            )

            # --- Phase 2: scoring + narrative (one batched LLM call per pass) ---
            if not client_jurisdictions:
                # Always supply at least US so the scoring rubric has something
                # to compare record jurisdictions against.
                client_jurisdictions = ['US']
            client_brand = {
                'mark': step5['exact'],
                'classes': client_classes,
                'jurisdictions': client_jurisdictions,
                'brand_reference': step5['order_meta'].get('brand_reference', ''),
                'countries': step5['countries'],
            }
            with st.spinner('Generating forensic commentary (Sonnet 4.6)...'):
                forensic_report = run_forensic_layer(
                    signa_records=signa_records,
                    client_brand=client_brand,
                    report_type=report_type,
                    narrative_client=narrative_client,
                    client_classes=client_classes,
                    client_mark=step5['exact'],
                    client_jurisdiction=client_jurisdictions,
                    # BR-010 — pass selected 3a–3d rows; empty dict in
                    # default mode means run_forensic_layer skips the
                    # extras LLM call entirely.
                    extras_rows=extras_rows,
                    # BR-011 — monitoring report mode switches the
                    # narrative tone to observational and skips the
                    # Recommended Specification / Actions sections. Read
                    # from the order_meta we built in step 5 (auto-set
                    # from Order Form A2 = "Monitoring or Representation
                    # Report").
                    is_monitoring=(step5['order_meta'].get('document_type') == 'monitoring'),
                )

            # --- Phase 3: render docx ---
            with st.spinner('Rendering forensic appendix...'):
                forensic_bytes = build_forensic_appendix(forensic_report, step5['order_meta'])

            st.session_state['forensic_result'] = {
                'bytes': forensic_bytes,
                'report_type': report_type.value,
                'record_count': len(signa_records),
                'verified_count': verified_count,
            }
            st.success('✅ Forensic layer complete.')
        except Exception as e:
            st.error(f'Forensic audit failed: {e}')
            raise

    # Display forensic download if available
    forensic = st.session_state.get('forensic_result')
    if forensic:
        st.divider()
        st.subheader('Download the forensic appendix')
        rt_label = ('Pre-Application' if forensic['report_type'] == 'pre_application'
                    else 'Post-Registration')
        safe_client = ''.join(c if c.isalnum() else '_' for c in step5['client_name'])[:40]
        f_filename = f'Braudit Forensic Appendix ({rt_label}) – {safe_client} – {date.today().isoformat()}.docx'
        st.download_button(
            label=f'⬇ Download forensic appendix ({rt_label})',
            data=forensic['bytes'],
            file_name=f_filename,
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True,
        )
        st.caption(
            f'{forensic["record_count"]} records audited · '
            f'{forensic["verified_count"]} verified via Signa · '
            f'{rt_label} report type.'
        )


# ---------- footer ----------
st.divider()
st.caption('Braudit Audit Tool · v2.0 · Internal use only · Steps 2–6 (Step 6 forensic layer powered by Signa + Anthropic Sonnet 4.6)')
