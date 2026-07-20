"""
PMV Toolkit Delivery Tracker — Streamlit App (Karad Division v1)
==================================================================

Workflow:
  1. Upload Office Master File, Old PMV Report, New PMV Report (3 tabs).
  2. Click "Process Reports" -> pipeline runs using the CURRENT persistent
     RTS exclusion list (fetched from GitHub).
  3. Review the on-screen summary dashboard.
  4. In the RTS List Manager, review Pending items and confirm any that are
     actually Returned to Sender -> saves to GitHub -> re-processes so the
     final download reflects the update.
  5. Download the finished Excel workbook.
"""

from __future__ import annotations

import io
from datetime import date

import pandas as pd
import streamlit as st

from pmv_core import process, COL_APP_NO, COL_ARTISAN_NAME, COL_PIN, COL_STATUS
from pmv_excel_writer import write_workbook
from rts_github_store import load_rts_list, save_rts_list

st.set_page_config(page_title="PMV Toolkit Delivery Tracker — Karad", layout="wide")

# --------------------------------------------------------------------------
# Session state
# --------------------------------------------------------------------------
if "result" not in st.session_state:
    st.session_state.result = None
if "rts_df" not in st.session_state or "rts_sha" not in st.session_state:
    st.session_state.rts_df, st.session_state.rts_sha = load_rts_list()

st.title("PMV Toolkit Delivery Tracker — Karad Division")
st.caption(
    "Tracks toolkit delivery status across successive PMV report exports, "
    "restricted to Karad Division jurisdiction, with a persistent "
    "Returned-to-Sender exclusion list."
)

# --------------------------------------------------------------------------
# Step 1 — Uploads
# --------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["1. Office Master Data", "2. Old PMV Report", "3. New PMV Report"])

with tab1:
    st.write("Karad Division Office Master File (.xlsx) — Division / Sub Division / Taluka / PIN structure.")
    master_file = st.file_uploader("Upload Office Master File", type=["xlsx"], key="master_upload")

with tab2:
    st.write("The OLDER of the two PMV toolkit delivery report exports (.csv).")
    old_file = st.file_uploader("Upload Old PMV Report", type=["csv"], key="old_upload")
    old_date_override = st.text_input(
        "Old report date (dd.mm.yyyy) — leave blank to auto-detect from filename",
        key="old_date_override",
    )

with tab3:
    st.write("The NEWER of the two PMV toolkit delivery report exports (.csv).")
    new_file = st.file_uploader("Upload New PMV Report", type=["csv"], key="new_upload")
    new_date_override = st.text_input(
        "New report date (dd.mm.yyyy) — leave blank to auto-detect from filename",
        key="new_date_override",
    )

st.divider()

ready = master_file is not None and old_file is not None and new_file is not None
process_clicked = st.button("Process Reports", type="primary", disabled=not ready)
if not ready:
    st.info("Upload all three files above to enable processing.")


def _run_pipeline():
    rts_set = set(st.session_state.rts_df[COL_APP_NO].astype(str)) if not st.session_state.rts_df.empty else set()
    result = process(
        old_csv_path=old_file,
        new_csv_path=new_file,
        master_xlsx_path=master_file,
        rts_app_numbers=rts_set,
        old_date_label=old_date_override or None,
        new_date_label=new_date_override or None,
    )
    st.session_state.result = result


if process_clicked:
    # File uploader objects need to be re-read from the start on each pass.
    master_file.seek(0)
    old_file.seek(0)
    new_file.seek(0)
    try:
        _run_pipeline()
        st.success("Processed successfully.")
    except Exception as e:
        st.error(f"Processing failed: {e}")

# --------------------------------------------------------------------------
# Step 2 — Summary dashboard
# --------------------------------------------------------------------------
result = st.session_state.result
if result is not None:
    st.header("Summary")
    s = result.summary

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Delivered", s["delivered_count"])
    c2.metric("Pending (RTS-excluded)", s["pending_count"])
    c3.metric("Returned to Sender", s["returned_to_sender_count"])
    c4.metric(f"Newly Delivered", s["newly_delivered_count"],
              help=f"Pending in {result.old_date_label}, Delivered in {result.new_date_label}")
    c5.metric("Other Divisions (excluded)", s["new_report_other_division_rows"])

    with st.expander("Full run details"):
        st.json(s)
        if "warning_ambiguous_taluka_pins" in s:
            st.warning(
                f"These PINs map to more than one Taluka in the master file — "
                f"worth reviewing: {s['warning_ambiguous_taluka_pins']}"
            )

    # ----------------------------------------------------------------
    # Step 3 — RTS List Manager
    # ----------------------------------------------------------------
    st.header("Returned-to-Sender List Manager")
    st.write(
        "This list is the permanent record of Application Numbers confirmed as "
        "Returned to Sender. It is stored in GitHub and applied automatically "
        "on every future run. Mark any genuinely RTS items below, then save."
    )

    colA, colB = st.columns(2)

    with colA:
        st.subheader("Currently Pending — mark any that are actually RTS")
        if result.pending.empty:
            st.write("No pending items in this run.")
            newly_marked = []
        else:
            pending_display = result.pending[[COL_APP_NO, COL_ARTISAN_NAME, COL_PIN, COL_STATUS]].reset_index(drop=True)
            pending_display.insert(0, "Mark as RTS", False)
            edited = st.data_editor(
                pending_display,
                hide_index=True,
                use_container_width=True,
                disabled=[COL_APP_NO, COL_ARTISAN_NAME, COL_PIN, COL_STATUS],
                key="pending_editor",
            )
            newly_marked = edited.loc[edited["Mark as RTS"], COL_APP_NO].tolist()

        if st.button("Save marked items to RTS list", disabled=(len(newly_marked) == 0)):
            current = set(st.session_state.rts_df[COL_APP_NO].astype(str)) if not st.session_state.rts_df.empty else set()
            updated = current | set(newly_marked)
            new_sha = save_rts_list(
                updated,
                st.session_state.rts_sha,
                commit_message=f"Add {len(newly_marked)} RTS entr{'y' if len(newly_marked)==1 else 'ies'} — {date.today().isoformat()}",
            )
            st.session_state.rts_df = pd.DataFrame({COL_APP_NO: sorted(updated)})
            st.session_state.rts_sha = new_sha
            st.success(f"Saved. RTS list now has {len(updated)} entries. Re-processing...")
            _run_pipeline()
            st.rerun()

    with colB:
        st.subheader("Current RTS exclusion list")
        st.write(f"{len(st.session_state.rts_df)} entries stored.")
        if not st.session_state.rts_df.empty:
            rts_display = st.session_state.rts_df.reset_index(drop=True).copy()
            rts_display.insert(0, "Remove", False)
            edited_rts = st.data_editor(
                rts_display, hide_index=True, use_container_width=True,
                disabled=[COL_APP_NO], key="rts_editor",
            )
            to_remove = edited_rts.loc[edited_rts["Remove"], COL_APP_NO].tolist()
            if st.button("Remove selected from RTS list", disabled=(len(to_remove) == 0)):
                current = set(st.session_state.rts_df[COL_APP_NO].astype(str))
                updated = current - set(to_remove)
                new_sha = save_rts_list(
                    updated,
                    st.session_state.rts_sha,
                    commit_message=f"Remove {len(to_remove)} RTS entr{'y' if len(to_remove)==1 else 'ies'} — {date.today().isoformat()}",
                )
                st.session_state.rts_df = pd.DataFrame({COL_APP_NO: sorted(updated)})
                st.session_state.rts_sha = new_sha
                st.success("Removed. Re-processing...")
                _run_pipeline()
                st.rerun()

    # ----------------------------------------------------------------
    # Step 4 — Download
    # ----------------------------------------------------------------
    st.header("Download")
    buf = io.BytesIO()
    write_workbook(result, buf)
    buf.seek(0)
    st.download_button(
        "Download Excel Report",
        data=buf,
        file_name=f"PMV_Karad_Report_{result.old_date_label}_to_{result.new_date_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
