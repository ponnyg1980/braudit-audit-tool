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
    extract_trademark_images,
)
from pipeline.report_builder import build_step5_report
from pipeline.forensic import SignaClient, verify_records
from pipeline.temmy import TemmyClient
from pipeline.forensic_narrative import (
    NarrativeClient, ReportType, run_forensic_layer,
)
from pipeline.forensic_report import build_forensic_appendix


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
    st.session_state['order_meta'] = extract_order_metadata(tmp.name)
    st.session_state['specific_terms'] = extract_specific_terms(tmp.name)

if uploaded is not None:
    _ingest_upload(uploaded)
    meta = st.session_state.get('order_meta') or {}
    if meta.get('client_name'):
        st.success(
            f"📋 Order Form parsed: **{meta['client_name']}** · classes {meta.get('classes_csv','—')} "
            f"· exact match \"{meta.get('exact_match','—')}\""
        )
    else:
        st.warning('Uploaded \u2014 but the Order Form sheet was not found or could not be read. Fill the fields below manually.')
else:
    st.info('Upload a scraped-results spreadsheet to pre-fill the search criteria below.')

st.divider()

# All the parsed defaults flow into the form via session_state.
meta = st.session_state.get('order_meta') or {}

with st.form('audit_form', clear_on_submit=False):

    st.subheader('2. Audit operator details')
    st.caption('Pre-filled from the spreadsheet (Order Form rows 58\u201364). Edit any field to override for this audit only. The two reference fields are optional \u2014 they don\u2019t affect the audit but flow through to the report header.')
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
    # Required operator fields (rows 60\u201364)
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
        exact = st.text_input('Exact match search term *', value=meta.get('exact_match', ''), help='From Order Form R14.')
        classes_text = st.text_input('Trademark classes (comma-separated) *', value=meta.get('classes_csv', ''), help='From Order Form G&S Classes rows.')
        sic_code = st.text_input('Client SIC code *', value=meta.get('sic', ''), help='From Order Form R8.')
    with c4:
        deal_id = st.text_input('Deal ID', value=meta.get('deal_id', ''), help='From Order Form R6.')
        similar = st.text_input('Similar match (Starts With) search term', value=meta.get('starts_with', ''), help='From Order Form R15.')
        countries = st.text_input('Designated countries', value=meta.get('countries', ''), help='From Order Form R10.')
        nature = st.text_input('Nature of business', value=meta.get('nature', ''), help='From Order Form R9.')

    search_platforms = st.text_area(
        'Search platforms',
        value=meta.get('search_platforms', ''),
        help='From Order Form R11 \u2014 comma-separated list of platforms the scrape covered.',
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
        ('Exact match search term', exact), ('Trademark classes', classes_text),
        ('Client SIC code', sic_code),
    ]:
        if not val.strip():
            missing.append(label)
    if uploaded is None or 'uploaded_path' not in st.session_state:
        missing.append('Scraped data .xlsx file')

    if missing:
        st.error('Please complete: ' + ', '.join(missing))
        st.stop()

    tmp_path = st.session_state['uploaded_path']
    specific_terms = st.session_state.get('specific_terms') or {}

    try:
        with st.spinner('Reading scraped data...'):
            sheets = read_sheets(tmp_path)
            tm_images = extract_trademark_images(tmp_path)

        target_classes = tuple(
            int(x.strip()) for x in classes_text.split(',') if x.strip().isdigit()
        ) or (11, 12, 35)

        root_word = exact.strip().split()[0].upper() if exact.strip() else 'STEALTH'

        with st.spinner('Step 2: de-duplicating and applying exclusions...'):
            tm_all = process_trademarks(
                sheets.get('Trademarks', [[]]),
                target_classes=target_classes,
                root=root_word,
                images=tm_images,
            )
            companies = process_companies(
                sheets.get('Companies', [[]]),
                target_sic=sic_code.strip(),
                root=root_word,
            )
            google = process_google(sheets.get('Google', [[]]))
            domains = process_domains(sheets.get('Domains', [[]]))
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
    filename = f'Braudit Report – {safe_client} – {date.today().isoformat()}.docx'
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

    f_c1, f_c2, f_c3 = st.columns([2, 2, 1.4])
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
        # Default to top 10 by score to keep the cost predictable. Operators
        # can dial up to all live records if they want full coverage.
        live_count = len(step5['tm_live'])
        default_top_n = min(10, live_count) if live_count else 1
        top_n = st.number_input(
            'Records to forensically audit',
            min_value=1, max_value=max(1, live_count), value=default_top_n, step=1,
            help='Top N live trademarks by score will be verified via Signa and forensically commented. '
                 'Higher N → higher API cost. 10 is a sensible default.',
        )
    with f_c3:
        client_jurisdiction = st.selectbox(
            'Client jurisdiction',
            options=['US', 'GB', 'EU', 'WO'],
            index=0,
            help='Client\u2019s primary filing jurisdiction. Used by the scoring rubric for the region criterion.',
        )

    forensic_clicked = st.button(
        '▶ Run Forensic Audit',
        type='primary',
        use_container_width=True,
        disabled=not (has_signa and has_anthropic),
    )

    if forensic_clicked:
        try:
            # Sort live records by score descending, take top N
            sorted_live = sorted(
                step5['tm_live'],
                key=lambda t: -int(t.get('score') or 0),
            )
            target_records = sorted_live[:int(top_n)]

            # Build clients
            signa_client = SignaClient(api_key=st.secrets['signa_api_key'])
            narrative_client = NarrativeClient(api_key=st.secrets['anthropic_api_key'])
            # TemmyDB is Braudit's source of truth for UK records (BR-002).
            # Optional: if the temmy_api_key secret is not set, UK records
            # skip the Temmy primary path and fall back to the Signa
            # Brexit-clone proxy for UK009xxx numbers only.
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
            # Per-source breakdown — useful for diagnosing UK coverage at a glance.
            by_source: dict = {}
            for r in signa_records:
                by_source[r.verification_source] = by_source.get(r.verification_source, 0) + 1
            src_summary = ', '.join(f'{k}={v}' for k, v in sorted(by_source.items()))
            st.info(
                f'Verification: {verified_count} of {len(signa_records)} records confirmed '
                f'({src_summary}).'
            )

            # --- Phase 2: scoring + narrative (one batched LLM call per pass) ---
            client_brand = {
                'mark': step5['exact'],
                'classes': client_classes,
                'jurisdiction': client_jurisdiction,
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
                    client_jurisdiction=client_jurisdiction,
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
