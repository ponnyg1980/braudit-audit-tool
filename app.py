"""Braudit Audit Tool — Streamlit MVP for Steps 2–5.

Internal tool for The Trademark Helpline / Braudit. Run with:
    streamlit run app.py
"""
from __future__ import annotations
import os
import tempfile
import streamlit as st
from datetime import date

from pipeline.filters import (
    read_sheets, process_trademarks, process_companies,
    process_google, process_domains, process_social,
)
from pipeline.report_builder import build_step5_report


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

with st.form('audit_form', clear_on_submit=False):
    st.subheader('Client & order details')
    c1, c2 = st.columns(2)
    with c1:
        client_first = st.text_input('Client first name *')
        client_email = st.text_input('Client email address *')
        account_manager = st.text_input('Account manager *')
    with c2:
        client_last = st.text_input('Client last name *')
        client_name = st.text_input('Client / company name *', help="e.g. UK Performance Parts Ltd")
        prepared_by = st.text_input('Report prepared by *')

    st.subheader('Search criteria')
    c3, c4 = st.columns(2)
    with c3:
        exact = st.text_input('Exact match search term *', value='Stealth LED')
        classes_text = st.text_input('Trademark classes (comma-separated)', value='11, 12, 35')
        sic_code = st.text_input('Client SIC code', value='45320')
    with c4:
        similar = st.text_input('Similar match search term', value='Stealth')
        countries = st.text_input('Designated countries', value='USA Office')
        nature = st.text_input('Nature of business', value='Retail trade of motor vehicle parts and accessories')

    st.subheader('Scraped data')
    uploaded = st.file_uploader(
        'Upload the scraped-results .xlsx file *',
        type=['xlsx'],
        help='The workbook produced by the Braudit scrape job, with the standard sheet names (Trademarks, Companies, Google, Domains, Social).'
    )

    submitted = st.form_submit_button('▶ Run Audit', type='primary', use_container_width=True)


# ---------- processing ----------

if submitted:
    missing = []
    for label, val in [
        ('Client first name', client_first), ('Client last name', client_last),
        ('Client email', client_email), ('Account manager', account_manager),
        ('Report prepared by', prepared_by), ('Client / company name', client_name),
        ('Exact match search term', exact),
    ]:
        if not val.strip():
            missing.append(label)
    if uploaded is None:
        missing.append('Scraped data .xlsx file')

    if missing:
        st.error('Please complete: ' + ', '.join(missing))
        st.stop()

    # Save upload to a tempfile so openpyxl can read it from disk
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
        with st.spinner('Reading scraped data...'):
            sheets = read_sheets(tmp_path)

        # Parse classes
        target_classes = tuple(
            int(x.strip()) for x in classes_text.split(',') if x.strip().isdigit()
        ) or (11, 12, 35)

        # Use the exact-match search term as the mark root (uppercase)
        # Strip any descriptive suffix to get to the root word for filtering
        root_word = exact.strip().split()[0].upper() if exact.strip() else 'STEALTH'

        with st.spinner('Step 2: de-duplicating and applying exclusions...'):
            tm_all = process_trademarks(
                sheets.get('Trademarks', [[]]),
                target_classes=target_classes,
                root=root_word,
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
            order_meta = {
                'client_name': client_name,
                'client_first': client_first,
                'client_last': client_last,
                'client_email': client_email,
                'account_manager': account_manager,
                'prepared_by': prepared_by,
                'search_date': date.today().strftime('%d %B %Y'),
                'search_type': 'Word',
                'mark_label': f'Exact: {exact}   ·   Similar: {similar}' if similar else f'Exact: {exact}',
                'exact': exact, 'similar': similar,
                'classes': classes_text,
                'sic': sic_code,
                'nature': nature,
                'countries': countries,
                'filtering_rules': (
                    f'Mark scope: exact {root_word} or "{root_word} " + descriptor; '
                    f'class touch any of {classes_text}; '
                    f'SIC = {sic_code} (companies); '
                    'dead trademarks retained but tagged Negligible.'
                ),
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

        st.success('✅ Audit complete.')

        # --- Summary ---
        st.subheader('Summary of flagged results')
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric('Live Trademarks', len(tm_live))
        m2.metric('Dead Trademarks', len(tm_dead))
        m3.metric('Companies', len(companies))
        m4.metric('Domains', len(domains))
        m5.metric('Google + Social', len(google) + len(social))

        # Risk band breakdown for live trademarks
        from collections import Counter
        risk_dist = Counter(t['risk'] for t in tm_live)
        st.write('**Live trademark risk distribution:** ' + ', '.join(
            f'{k}: {v}' for k, v in risk_dist.most_common()
        ) if risk_dist else '_no live trademarks flagged_')

        # --- Download ---
        st.divider()
        st.subheader('Download the report')
        safe_client = ''.join(c if c.isalnum() else '_' for c in client_name)[:40]
        filename = f'Braudit Report – {safe_client} – {date.today().isoformat()}.docx'
        st.download_button(
            label='⬇ Download Word report',
            data=docx_bytes,
            file_name=filename,
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True,
        )

        # Show preview tables
        with st.expander('Preview — Live trademarks'):
            st.dataframe([
                {k: v for k, v in t.items() if k in ('app','mark','classes','type','owner','status','score','risk')}
                for t in tm_live
            ], use_container_width=True)
        if tm_dead:
            with st.expander(f'Preview — Dead trademarks ({len(tm_dead)})'):
                st.dataframe([
                    {k: v for k, v in t.items() if k in ('app','mark','classes','type','owner','status','risk')}
                    for t in tm_dead
                ], use_container_width=True)

    finally:
        os.unlink(tmp_path)

# ---------- footer ----------
st.divider()
st.caption('Braudit Audit Tool · v1.0 · Internal use only · Steps 2–5 only (Step 6 forensic audit coming in v2)')
