"""
FYW Pending Order Report — Streamlit App
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime
from report_engine import (
    load_fyw_csv,
    load_shopee_files,
    build_report_rows,
    generate_excel,
    merge_manual_orders,
)
from tc_reconciliation import (
    run_full_reconciliation,
    build_manual_entry_template,
    apply_manual_overrides,
    manual_confirmed_to_df,
)

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="FYW Pending Order Report",
    page_icon="📦",
    layout="wide",
)

st.title("📦 FYW Pending Order Report")
st.markdown("Upload your daily FYW dashboard export and marketplace files to generate reports.")

tab_report, tab_recon = st.tabs(["📊 Pending Order Report", "🔄 TC Reconciliation Check"])

# Persist manually-confirmed "pushed to TC" orders across reruns/reuploads.
# Structure: { marketplace_label: [ {eOrder Number, Invoice Number, ...}, ... ] }
if 'manual_confirmed' not in st.session_state:
    st.session_state.manual_confirmed = {}

# ════════════════════════════════════════════════════════════════════════════
# SIDEBAR — Shared file uploads
# ════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📁 Upload Files")

    fyw_csv = st.file_uploader(
        "FYW Dashboard CSV (ALL-DD-Mon-YYYY.csv) *",
        type=["csv"],
        help="Required — the main FYW order export (this is 'TC')"
    )

    st.markdown("---")
    st.markdown("**Shopee Seller Centre Files** _(per brand)_")
    melissa_file = st.file_uploader("Melissa Shopee Order (.xlsx)", type=["xlsx"], key="melissa_shopee")
    ipanema_file = st.file_uploader("Ipanema Shopee Order (.xlsx)", type=["xlsx"], key="ipanema_shopee")
    cspace_file  = st.file_uploader("CSpace Shopee Order (.xlsx)",  type=["xlsx"], key="cspace_shopee")

    st.markdown("---")
    st.markdown("**Other Marketplace Files** _(for TC reconciliation)_")
    tiktok_file = st.file_uploader("TikTok Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="tiktok_file")

    st.markdown("**Lazada** _(per brand)_")
    lazada_melissa_file = st.file_uploader("Lazada - Melissa Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="lazada_melissa")
    lazada_ipanema_file = st.file_uploader("Lazada - Ipanema Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="lazada_ipanema")
    lazada_cspace_file  = st.file_uploader("Lazada - CSpace Order Export (.xlsx/.csv)",  type=["xlsx", "csv"], key="lazada_cspace")

    zalora_file = st.file_uploader("Zalora Order Export (.xlsx/.csv)", type=["xlsx", "csv"], key="zalora_file")

    st.markdown("---")
    report_date = st.date_input("Report Date", value=datetime.today())

if not fyw_csv:
    st.info("👈 Upload the FYW CSV file to get started.")
    st.stop()

# ════════════════════════════════════════════════════════════════════════════
# TC RECONCILIATION — computed here (code order) so results are ready for the
# Pending Order Report tab's NOT PUSHED sheet + preview below. This still
# renders into the "🔄 TC Reconciliation Check" tab visually (tab order is
# controlled by st.tabs(), not by code order).
# ════════════════════════════════════════════════════════════════════════════
results = None  # populated below if marketplace files are uploaded

with tab_recon:
    st.markdown(
        "Checks whether orders from each marketplace export have successfully been "
        "**pushed to TC** (i.e. appear in the FYW dashboard CSV). Confirm any orders "
        "that were pushed but missed by the auto-match here — updates flow straight "
        "into the **Pending Order Report** tab's FILTERED DATA, PIVOT, and NOT PUSHED sheet."
    )

    mp_files = {
        'Shopee - Melissa': melissa_file,
        'Shopee - Ipanema': ipanema_file,
        'Shopee - CSpace':  cspace_file,
        'TikTok':           tiktok_file,
        'Lazada - Melissa': lazada_melissa_file,
        'Lazada - Ipanema': lazada_ipanema_file,
        'Lazada - CSpace':  lazada_cspace_file,
        'Zalora':           zalora_file,
    }
    uploaded_mp_files = {k: v for k, v in mp_files.items() if v is not None}

    if not uploaded_mp_files:
        st.info("👈 Upload at least one marketplace file (Shopee, TikTok, Lazada, or Zalora) in the sidebar to run the reconciliation check.")
    else:
        # ── Reupload TC file to refresh the check ────────────────────────────
        with st.expander("🔄 Reupload TC order file (refresh check with updated TC data)"):
            st.caption(
                "If you've since re-exported the FYW dashboard CSV (e.g. after pushing more "
                "orders), upload it here to re-run the match. Orders you've already confirmed "
                "below stay confirmed either way."
            )
            tc_reupload = st.file_uploader(
                "Updated FYW Dashboard CSV", type=["csv"], key="tc_reupload_csv"
            )

        recon_error = None
        with st.spinner("Loading FYW CSV for reconciliation..."):
            active_tc_file = tc_reupload if tc_reupload is not None else fyw_csv
            active_tc_file.seek(0)
            fyw_raw_df = pd.read_csv(active_tc_file)

        with st.spinner("Cross-checking marketplace orders against TC..."):
            try:
                base_results = run_full_reconciliation(fyw_raw_df, uploaded_mp_files)
            except Exception as e:
                recon_error = str(e)
                base_results = None

        if recon_error:
            st.error(f"❌ Reconciliation failed: {recon_error}")
        else:
            # Fold in previously-confirmed manual entries so they stay applied
            # across reruns/reuploads.
            results = apply_manual_overrides(base_results, st.session_state.manual_confirmed)

            if tc_reupload is not None:
                st.success("✅ Reconciliation refreshed using the reuploaded TC file.")

            st.subheader("📊 Reconciliation Summary")

            total_missing   = sum(r.get('missing_count', 0) for r in results.values())
            total_mp_orders = sum(r.get('total_mp_orders', 0) for r in results.values())
            total_manual    = sum(r.get('manual_confirmed_count', 0) for r in results.values())

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Marketplaces Checked", len(results))
            col2.metric("Total MP Orders",      total_mp_orders)
            col3.metric("Manually Confirmed",   total_manual)
            col4.metric(
                "Total Missing from TC",
                total_missing,
                delta=f"{total_missing} not pushed" if total_missing > 0 else "All synced ✅",
                delta_color="inverse" if total_missing > 0 else "normal",
            )

            summary_rows = []
            for label, res in results.items():
                if 'error' in res:
                    summary_rows.append({
                        'Marketplace': label,
                        'Status': f"⚠️ {res['error']}",
                        'Total MP Orders': '-',
                        'Matched in TC': '-',
                        'Manually Confirmed': '-',
                        'Missing from TC': '-',
                        'Match Rate': '-',
                    })
                else:
                    status = "✅ All synced" if res['missing_count'] == 0 else f"🔴 {res['missing_count']} missing"
                    summary_rows.append({
                        'Marketplace': label,
                        'Status': status,
                        'Total MP Orders': res['total_mp_orders'],
                        'Matched in TC': res['matched_count'],
                        'Manually Confirmed': res.get('manual_confirmed_count', 0),
                        'Missing from TC': res['missing_count'],
                        'Match Rate': f"{res['match_rate']}%" if res['match_rate'] is not None else '-',
                    })

            summary_df = pd.DataFrame(summary_rows)

            def highlight_status(row):
                if '🔴' in str(row['Status']):
                    return ['background-color: #FFCCCC'] * len(row)
                elif '✅' in str(row['Status']):
                    return ['background-color: #C6EFCE'] * len(row)
                elif '⚠️' in str(row['Status']):
                    return ['background-color: #FFE699'] * len(row)
                return [''] * len(row)

            st.dataframe(
                summary_df.style.apply(highlight_status, axis=1),
                use_container_width=True,
                hide_index=True,
            )

            # ── Not Pushed to TC ──────────────────────────────────────────────
            st.subheader("🔴 Not Pushed to TC")

            has_missing = any(r.get('missing_count', 0) > 0 for r in results.values())
            if not has_missing and total_manual == 0:
                st.success("🎉 All marketplace orders have been successfully pushed to TC!")
            else:
                for label, res in results.items():
                    if res.get('missing_count', 0) == 0:
                        continue
                    with st.expander(f"🔴 {label} — {res['missing_count']} orders not pushed to TC"):
                        missing_df = pd.DataFrame({
                            '#': range(1, len(res['missing_ids']) + 1),
                            'Order ID / Order Number': res['missing_ids'],
                        })
                        st.dataframe(missing_df, use_container_width=True, hide_index=True)

                        st.markdown(
                            "**Already pushed to TC but not auto-matched?** Tick **Confirmed** and "
                            "fill in what you know for each order below, then update the report."
                        )

                        editor_key = f"manual_editor_{label}"
                        template_df = build_manual_entry_template(res['missing_ids'])
                        edited_df = st.data_editor(
                            template_df,
                            key=editor_key,
                            hide_index=True,
                            use_container_width=True,
                            num_rows="fixed",
                            column_config={
                                'Confirmed': st.column_config.CheckboxColumn('Confirmed', default=False),
                                'eOrder Number': st.column_config.TextColumn('eOrder Number', disabled=True),
                                'Invoice Number': st.column_config.TextColumn('Invoice Number'),
                                'Payment Status': st.column_config.TextColumn('Payment Status'),
                                'Order Item Status': st.column_config.TextColumn('Order Item Status'),
                                'Ordered Date': st.column_config.TextColumn('Ordered Date'),
                                'Nickname': st.column_config.TextColumn('Nickname'),
                                'MP SLA': st.column_config.TextColumn('MP SLA'),
                            },
                        )

                        if st.button(f"🔄 Update Report — {label}", key=f"apply_{label}"):
                            confirmed_rows = edited_df[edited_df['Confirmed'] == True]
                            if confirmed_rows.empty:
                                st.warning("No rows ticked as Confirmed — nothing to apply.")
                            else:
                                new_entries = confirmed_rows.drop(columns=['Confirmed']).to_dict('records')
                                existing = st.session_state.manual_confirmed.setdefault(label, [])
                                existing_ids = {r['eOrder Number'] for r in existing}
                                added = 0
                                for row in new_entries:
                                    if row['eOrder Number'] not in existing_ids:
                                        existing.append(row)
                                        added += 1
                                st.success(
                                    f"Updated report — {added} order(s) moved out of Not Pushed and into "
                                    "FILTERED DATA/PIVOT. See the Pending Order Report tab."
                                )
                                st.rerun()

                if st.session_state.manual_confirmed and any(st.session_state.manual_confirmed.values()):
                    with st.expander("✅ Manually Confirmed Orders (applied)", expanded=False):
                        st.dataframe(
                            manual_confirmed_to_df(st.session_state.manual_confirmed),
                            use_container_width=True,
                            hide_index=True,
                        )
                        if st.button("🗑️ Clear all manual confirmations"):
                            st.session_state.manual_confirmed = {}
                            st.rerun()

# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — PENDING ORDER REPORT
# ════════════════════════════════════════════════════════════════════════════
with tab_report:
    with st.spinner("Loading files..."):
        try:
            pending = load_fyw_csv(fyw_csv)
        except Exception as e:
            st.error(f"❌ Failed to load FYW CSV: {e}")
            st.stop()

        shopee_map = {'MELISSA': melissa_file, 'IPANEMA': ipanema_file, 'CSPACE': cspace_file}
        mp_sla_map, shopee_tracking = load_shopee_files(shopee_map)
        base_df_report = build_report_rows(pending, mp_sla_map, shopee_tracking)

    # Fold in any orders manually confirmed as pushed (from the TC Reconciliation
    # Check tab) so FILTERED DATA / PIVOT / the download all reflect them.
    df_report = merge_manual_orders(base_df_report, st.session_state.manual_confirmed)
    manual_row_count = len(df_report) - len(base_df_report)
    total_not_pushed = sum(r.get('missing_count', 0) for r in results.values()) if results else 0

    st.subheader("📊 Summary")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Pending Items",  len(df_report))
    col2.metric("Unique Orders",        df_report['Order ID'].nunique())
    col3.metric("Total Value (MYR)",    f"{df_report['Sold Price (MYR)'].sum():,.2f}")
    col4.metric("Active Channels",      df_report['Channel'].nunique())

    if results is not None:
        n1, n2, n3 = st.columns(3)
        n1.metric("Manually Confirmed (added)", manual_row_count)
        n2.metric(
            "Not Pushed to TC",
            total_not_pushed,
            delta="✅ All synced" if total_not_pushed == 0 else f"{total_not_pushed} pending",
            delta_color="normal" if total_not_pushed == 0 else "inverse",
        )
        n3.metric("Marketplaces Checked", len(results))
        st.caption(
            "Not Pushed to TC comes from the 🔄 TC Reconciliation Check tab. Confirm orders there "
            "to move them into Filtered Data / Pivot below, then re-download."
        )
    else:
        st.caption(
            "👉 Upload marketplace files (Shopee/TikTok/Lazada/Zalora) in the sidebar to also check "
            "which orders haven't been pushed to TC yet — see the 🔄 TC Reconciliation Check tab."
        )

    st.subheader("🗂️ Pivot: Orders by Nickname × FYW SLA")
    pivot_raw = df_report.pivot_table(
        index='Nickname', columns='FYW SLA', values='Order ID',
        aggfunc='count', fill_value=0
    )
    pivot_raw['Grand Total'] = pivot_raw.sum(axis=1)
    total_row = pivot_raw.sum(axis=0)
    total_row.name = 'Grand Total'
    pivot_display = pd.concat([pivot_raw, total_row.to_frame().T]).astype(int)
    pivot_display = pivot_display.replace(0, '-')

    def style_pivot(df):
        styles = pd.DataFrame('', index=df.index, columns=df.columns)
        today  = datetime.today().date()
        yest   = (datetime.today() - pd.Timedelta(days=1)).date()
        for col in df.columns:
            if col == 'Grand Total':
                styles[col] = 'background-color: #DDEEFF; font-weight: bold'
            else:
                try:
                    col_date = datetime.strptime(col, '%d/%m/%Y').date()
                    if col_date < yest:
                        styles[col] = 'background-color: #FFCCCC; color: #9C0006; font-weight: bold'
                    elif col_date == yest:
                        styles[col] = 'background-color: #FCE4D6; color: #833C00; font-weight: bold'
                    elif col_date == today:
                        styles[col] = 'background-color: #C6EFCE; color: #006100; font-weight: bold'
                except Exception:
                    pass
        return styles

    st.dataframe(pivot_display.style.apply(style_pivot, axis=None), use_container_width=True)

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.markdown("🔴 **Past due**")
    col_b.markdown("🟠 **Yesterday**")
    col_c.markdown("🟢 **Today**")
    col_d.markdown("🔵 **Grand Total**")

    st.subheader("🏷️ Brand Summary")
    brand_summary = df_report.groupby('Brand').agg(
        Orders=('Order ID', 'nunique'),
        Items=('Qty', 'sum'),
        Value=('Sold Price (MYR)', 'sum'),
        Channels=('Channel', lambda x: ', '.join(sorted(set(x))))
    ).reset_index()
    brand_summary['Value'] = brand_summary['Value'].apply(lambda x: f"MYR {x:,.2f}")
    st.dataframe(brand_summary, use_container_width=True, hide_index=True)

    st.subheader("📋 All Pending Orders")
    with st.expander("🔍 Filter options"):
        c1, c2, c3 = st.columns(3)
        brand_filter   = c1.multiselect("Brand",   options=sorted(df_report['Brand'].unique()),   default=list(df_report['Brand'].unique()))
        channel_filter = c2.multiselect("Channel", options=sorted(df_report['Channel'].unique()), default=list(df_report['Channel'].unique()))
        status_filter  = c3.multiselect("Status",  options=sorted(df_report['FYW Status'].unique()), default=list(df_report['FYW Status'].unique()))

    filtered = df_report[
        df_report['Brand'].isin(brand_filter) &
        df_report['Channel'].isin(channel_filter) &
        df_report['FYW Status'].isin(status_filter)
    ]
    st.dataframe(filtered.drop(columns=['Nickname']), use_container_width=True, hide_index=True)
    st.caption(f"Showing {len(filtered)} of {len(df_report)} orders")

    if results is not None and total_not_pushed > 0:
        st.subheader("🔴 Not Pushed to TC")
        not_pushed_rows = []
        for label, res in results.items():
            for oid in res.get('missing_ids', []):
                not_pushed_rows.append({'Marketplace': label, 'Order ID / Order Number': oid})
        st.dataframe(pd.DataFrame(not_pushed_rows), use_container_width=True, hide_index=True)
        st.caption(
            "Go to the 🔄 TC Reconciliation Check tab to confirm any of these that were actually pushed."
        )

    st.subheader("⬇️ Download Excel Report")
    with st.spinner("Generating Excel..."):
        excel_buffer = generate_excel(
            df_report,
            report_date=datetime.combine(report_date, datetime.min.time()),
            not_pushed_results=results,
        )

    filename = f"FYW_Pending_Orders_{report_date.strftime('%d%b%Y')}.xlsx"
    st.download_button(
        label="📥 Download Excel Report",
        data=excel_buffer,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="download_pending_report",
    )
    st.success(f"✅ **{filename}** — {len(df_report)} items | MYR {df_report['Sold Price (MYR)'].sum():,.2f}")
