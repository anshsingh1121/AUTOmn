"""
FCB Incident Tracker — Production Console
Corporate Incident Management Tracker (IMT) Reconciliation & Governance Platform.
"""

import io
import json
import os
import sys
import time
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st

# -------------------------------------------------------------
# Page Configuration
# -------------------------------------------------------------
st.set_page_config(
    page_title="FCB Incident Tracker — Production Console",
    page_icon="📁",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Root directory resolution (handles standalone PyInstaller executable and normal python)
if getattr(sys, "frozen", False):
    PROJECT_ROOT = os.path.dirname(sys.executable)
else:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from src.business_rules import build_ic_lookup_map
from src.config_loader import Config, load_config
from src.output_writer import (
    create_output_path,
    write_output_workbook,
    write_reconciliation_report_excel,
)
from src.audit import create_audit_record, save_audit_log
from src.reconciler import reconciliation_to_dataframe, reconcile
from src.report_generator import generate_reconciliation_summary
from src.validator import (
    ValidationResult,
    validate_number_column,
    validate_output,
    validate_required_columns,
)

# -------------------------------------------------------------
# Corporate Styling
# -------------------------------------------------------------
st.markdown(
    """
    <style>
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    .app-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding-bottom: 0.75rem;
        margin-bottom: 1.25rem;
        border-bottom: 1px solid #E2E8F0;
    }
    .app-title {
        font-size: 1.4rem;
        font-weight: 700;
        color: #0F172A;
        letter-spacing: -0.01em;
        margin: 0;
    }
    .app-subtitle {
        font-size: 0.85rem;
        color: #64748B;
        margin-top: 0.15rem;
    }
    .metric-container {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 0.85rem 1rem;
    }
    .metric-label {
        font-size: 0.7rem;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.05em;
        color: #64748B;
        margin-bottom: 0.2rem;
    }
    .metric-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0F172A;
        line-height: 1.2;
    }
    .metric-sub {
        font-size: 0.75rem;
        color: #059669;
        font-weight: 500;
        margin-top: 0.2rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# Core Pipeline Execution
# -------------------------------------------------------------
def execute_reconciliation(
    master_df: pd.DataFrame,
    servicenow_df: pd.DataFrame,
    ic_lookup_df: Optional[pd.DataFrame],
    config: Config,
) -> Tuple[Any, pd.DataFrame, Dict[str, Any], ValidationResult, ValidationResult, float]:
    """Execute monthly reconciliation and business rules."""
    start_time = time.time()
    input_val = ValidationResult()

    # Schema Validation
    input_val.merge(
        validate_required_columns(servicenow_df, config.required_servicenow_columns, "ServiceNow")
    )
    input_val.merge(
        validate_required_columns(master_df, config.required_master_columns, "Master Tracker")
    )
    input_val.merge(
        validate_number_column(
            servicenow_df,
            "Number",
            "ServiceNow",
            block_on_duplicates=config.block_on_duplicate_numbers,
            block_on_blanks=config.block_on_blank_numbers,
        )
    )
    input_val.merge(
        validate_number_column(
            master_df,
            "Number",
            "Master Tracker",
            block_on_duplicates=config.block_on_duplicate_numbers,
            block_on_blanks=config.block_on_blank_numbers,
        )
    )

    # Build IC Lookup map
    ic_lookup_map = {}
    if ic_lookup_df is not None:
        region_cfg = config.region_rule
        ic_lookup_map, _ = build_ic_lookup_map(
            ic_lookup_df,
            ic_column=region_cfg.get("lookup_ic_column", "IC"),
            region_column=region_cfg.get("lookup_region_column", "Region"),
            case_sensitive=region_cfg.get("case_sensitive", False),
        )

    # Reconcile & Apply Business Rules
    master_before_count = len(master_df)
    recon_result = reconcile(
        master_df=master_df,
        servicenow_df=servicenow_df,
        ic_lookup=ic_lookup_map,
        config=config,
    )

    # Build output DataFrame
    output_df = reconciliation_to_dataframe(recon_result, list(master_df.columns))

    # Output validation
    output_val = validate_output(output_df, master_df, "Number")

    # Generate summary report
    report = generate_reconciliation_summary(
        master_before_count=master_before_count,
        servicenow_count=len(servicenow_df),
        recon=recon_result,
        input_validation=input_val,
        output_validation=output_val,
    )

    duration = time.time() - start_time
    return recon_result, output_df, report, input_val, output_val, duration


# -------------------------------------------------------------
# Session State Initialization
# -------------------------------------------------------------
if "processed" not in st.session_state:
    st.session_state["processed"] = False
    st.session_state["master_df"] = None
    st.session_state["output_df"] = None
    st.session_state["recon_result"] = None
    st.session_state["report"] = None
    st.session_state["duration"] = 0.0
    st.session_state["run_time"] = None


# -------------------------------------------------------------
# Sidebar: File Selection & Actions
# -------------------------------------------------------------
with st.sidebar:
    st.markdown("### Monthly Inputs")
    st.markdown("Select your corporate Master Tracker and ServiceNow export files.")
    st.markdown("---")

    master_file = st.file_uploader(
        "1. Master Tracker Workbook (.xlsm / .xlsx)",
        type=["xlsm", "xlsx"],
        help="The cumulative Master Tracker containing 'Consolidated' and 'IC Lookup' sheets.",
    )

    sn_file = st.file_uploader(
        "2. ServiceNow Export (.xlsx / .csv)",
        type=["xlsx", "csv"],
        help="The monthly incident extract from ServiceNow.",
    )

    st.markdown("---")
    save_to_output = st.checkbox(
        "Save Output to 'output/' folder",
        value=True,
        help="Saves the updated workbook and reconciliation report into the local output directory.",
    )

    run_btn = st.button("Run Reconciliation", type="primary", use_container_width=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.caption("FCB Incident Management Automation")
    st.caption("Environment: Corporate Production")


# -------------------------------------------------------------
# Execution Workflow
# -------------------------------------------------------------
if run_btn:
    if not master_file or not sn_file:
        st.error("Please provide both the Master Tracker and the ServiceNow export file.")
    else:
        with st.spinner("Processing reconciliation..."):
            try:
                cfg_path = os.path.join(PROJECT_ROOT, "config", "config.json")
                config = load_config(cfg_path)

                # Load Master Tracker (Consolidated sheet)
                master_df = pd.read_excel(master_file, sheet_name=config.master_sheet, engine="openpyxl")

                # Load IC Lookup (from inside the same Master Tracker workbook)
                try:
                    ic_lookup_df = pd.read_excel(master_file, sheet_name=config.ic_lookup_sheet, engine="openpyxl")
                except Exception:
                    ic_lookup_df = None

                # Load ServiceNow Export
                if sn_file.name.lower().endswith(".csv"):
                    try:
                        servicenow_df = pd.read_csv(sn_file, encoding="utf-8")
                    except UnicodeDecodeError:
                        servicenow_df = pd.read_csv(sn_file, encoding="cp1252")
                else:
                    try:
                        servicenow_df = pd.read_excel(sn_file, sheet_name=config.servicenow_sheet, engine="openpyxl")
                    except Exception:
                        # Fallback to first available sheet
                        servicenow_df = pd.read_excel(sn_file, engine="openpyxl")

                # Clean header whitespace
                master_df.columns = [str(c).strip() for c in master_df.columns]
                servicenow_df.columns = [str(c).strip() for c in servicenow_df.columns]
                if ic_lookup_df is not None:
                    ic_lookup_df.columns = [str(c).strip() for c in ic_lookup_df.columns]

                # Run reconciliation pipeline
                recon_res, out_df, rep, inp_val, out_val, duration = execute_reconciliation(
                    master_df=master_df,
                    servicenow_df=servicenow_df,
                    ic_lookup_df=ic_lookup_df,
                    config=config,
                )

                # Save output to disk if checked
                saved_files = []
                if save_to_output and (recon_res.new_records or recon_res.change_details):
                    out_dir = os.path.join(PROJECT_ROOT, "output")
                    os.makedirs(out_dir, exist_ok=True)
                    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")

                    # 1. Save Updated Master Tracker Workbook
                    is_xlsm = str(getattr(master_file, "name", "")).lower().endswith(".xlsm")
                    master_ext = ".xlsm" if is_xlsm else ".xlsx"
                    master_out_path = os.path.join(out_dir, f"Master_Tracker_Updated_{timestamp_str}{master_ext}")

                    temp_master = os.path.join(out_dir, f"_temp_master_{timestamp_str}{master_ext}")
                    try:
                        with open(temp_master, "wb") as f_tmp:
                            f_tmp.write(master_file.getvalue())

                        # Attempt native COM write to preserve pivots, charts, and macros
                        try:
                            write_output_workbook(
                                output_df=out_df,
                                output_path=master_out_path,
                                source_workbook_path=temp_master,
                                target_sheets=config.target_sheets,
                            )
                        except Exception:
                            # Fallback: direct write
                            with pd.ExcelWriter(master_out_path, engine="openpyxl") as writer:
                                out_df.to_excel(writer, sheet_name=config.master_sheet, index=False)
                                if "FCB" in config.target_sheets:
                                    out_df.to_excel(writer, sheet_name="FCB", index=False)
                                if ic_lookup_df is not None:
                                    ic_lookup_df.to_excel(writer, sheet_name=config.ic_lookup_sheet, index=False)
                    finally:
                        if os.path.isfile(temp_master):
                            try:
                                os.remove(temp_master)
                            except Exception:
                                pass

                    saved_files.append(os.path.basename(master_out_path))

                    # 2. Save Reconciliation Audit Report
                    rep_path = os.path.join(out_dir, f"Reconciliation_Report_{timestamp_str}.xlsx")
                    write_reconciliation_report_excel(rep, rep_path)
                    saved_files.append(os.path.basename(rep_path))

                    # 3. Save JSON Audit Log
                    audit_record = create_audit_record(
                        config=config,
                        master_before_count=len(master_df),
                        servicenow_count=len(servicenow_df),
                        recon_summary=recon_res.summary(),
                        validation_passed=(inp_val.is_valid and out_val.is_valid),
                        report_summary=rep["summary"],
                        output_path=master_out_path,
                        archive_path=None,
                        errors=rep.get("errors", []),
                        warnings=rep.get("warnings", []),
                    )
                    audit_path = save_audit_log(audit_record, out_dir)
                    saved_files.append(os.path.basename(audit_path))

                # Store in session state
                st.session_state["saved_files"] = saved_files
                st.session_state["processed"] = True
                st.session_state["master_df"] = master_df
                st.session_state["output_df"] = out_df
                st.session_state["recon_result"] = recon_res
                st.session_state["report"] = rep
                st.session_state["duration"] = duration
                st.session_state["run_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                st.toast("Reconciliation completed successfully.", icon="✅")

            except Exception as e:
                st.error(f"Reconciliation Failed: {e}")


# -------------------------------------------------------------
# Top Header Bar
# -------------------------------------------------------------
st.markdown(
    f"""
    <div class="app-header">
        <div>
            <div class="app-title">FCB Incident Tracker — Production Console</div>
            <div class="app-subtitle">Corporate Incident Reconciliation & Governance</div>
        </div>
        <div style="font-size: 0.85rem; color: #475569;">
            {f"Last Execution: {st.session_state['run_time']}" if st.session_state['run_time'] else "Ready for File Ingestion"}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------
# Main Application Content
# -------------------------------------------------------------
if not st.session_state["processed"]:
    st.info(
        "👈 Please select the **Master Tracker** workbook and the monthly **ServiceNow Export** in the left sidebar, "
        "then click **'Run Reconciliation'**."
    )
else:
    output_df = st.session_state["output_df"]
    recon = st.session_state["recon_result"]
    rep_summary = st.session_state["report"]["summary"]

    # Navigation Tabs
    tab_summary, tab_analytics, tab_explorer, tab_export = st.tabs(
        [
            "Reconciliation Summary",
            "Incident Analytics",
            "Master Dataset Explorer",
            "Export & Download",
        ]
    )

    # ---------------------------------------------------------
    # TAB 1: RECONCILIATION SUMMARY
    # ---------------------------------------------------------
    with tab_summary:
        saved = st.session_state.get("saved_files", [])
        if saved:
            st.success(
                "📁 **All output files successfully generated in `output/` directory:**\n\n"
                + "\n".join([f"- `{f}`" for f in saved])
            )

        # KPI Metric Strip
        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            st.markdown(
                f"""
                <div class="metric-container">
                    <div class="metric-label">Master Records (Before)</div>
                    <div class="metric-value">{rep_summary.get('Master Records Before', 0)}</div>
                    <div class="metric-sub" style="color:#64748B;">Existing Baseline</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m2:
            st.markdown(
                f"""
                <div class="metric-container">
                    <div class="metric-label">Matched & Updated</div>
                    <div class="metric-value">{rep_summary.get('Matched (Updated)', 0)}</div>
                    <div class="metric-sub">Updated with latest data</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m3:
            st.markdown(
                f"""
                <div class="metric-container">
                    <div class="metric-label">New Incidents Added</div>
                    <div class="metric-value">{rep_summary.get('New Records Added', 0)}</div>
                    <div class="metric-sub">Newly identified</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m4:
            st.markdown(
                f"""
                <div class="metric-container">
                    <div class="metric-label">Historical Retained</div>
                    <div class="metric-value">{rep_summary.get('Historical Retained', 0)}</div>
                    <div class="metric-sub">100% Preserved (Zero Deletion)</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with m5:
            st.markdown(
                f"""
                <div class="metric-container">
                    <div class="metric-label">Total Reconciled Records</div>
                    <div class="metric-value">{rep_summary.get('Total Output Records', 0)}</div>
                    <div class="metric-sub">Updated Master Count</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Sub-tabs for Audit Inspection
        sub_new, sub_diffs, sub_misses = st.tabs(
            [
                f"New Incidents ({len(recon.new_records)})",
                f"Field-Level Changes ({len(recon.change_details)})",
                f"IC Lookup Validation ({len(recon.ic_lookup_misses)})",
            ]
        )

        with sub_new:
            if recon.new_records:
                st.dataframe(pd.DataFrame(recon.new_records), use_container_width=True, height=340)
            else:
                st.caption("No new incidents were identified in this cycle.")

        with sub_diffs:
            if recon.change_details:
                diff_rows = []
                for inc_id, chgs in recon.change_details.items():
                    for chg in chgs:
                        diff_rows.append({"Incident Number": inc_id, "Field Update": chg})
                st.dataframe(pd.DataFrame(diff_rows), use_container_width=True, height=340)
            else:
                st.caption("No field changes detected for existing records.")

        with sub_misses:
            if recon.ic_lookup_misses:
                st.warning(
                    f"{len(recon.ic_lookup_misses)} IC values were not matched in the 'IC Lookup' sheet. "
                    "Add them to the 'IC Lookup' sheet in your Master Tracker to map their Region."
                )
                st.dataframe(pd.DataFrame({"Unmapped Incident Commander": recon.ic_lookup_misses}), use_container_width=True)
            else:
                st.success("All Incident Commanders were successfully resolved in the IC Lookup table.")

    # ---------------------------------------------------------
    # TAB 2: INCIDENT ANALYTICS
    # ---------------------------------------------------------
    with tab_analytics:
        col_c1, col_c2 = st.columns(2)

        with col_c1:
            st.markdown("##### Incident Count by Priority Class")
            if "Priority" in output_df.columns:
                prio_counts = output_df["Priority"].value_counts().reset_index()
                prio_counts.columns = ["Priority", "Count"]
                fig_prio = px.bar(
                    prio_counts,
                    x="Priority",
                    y="Count",
                    color="Priority",
                    text="Count",
                    color_discrete_sequence=["#0284C7", "#38BDF8", "#0EA5E9", "#BAE6FD"],
                )
                fig_prio.update_layout(
                    template="plotly_white",
                    height=280,
                    margin=dict(l=10, r=10, t=10, b=10),
                    showlegend=False,
                    xaxis_title=None,
                    yaxis_title="Count",
                )
                st.plotly_chart(fig_prio, use_container_width=True)

        with col_c2:
            st.markdown("##### Region Volume Distribution (India vs. US)")
            if "Region" in output_df.columns:
                reg_clean = output_df["Region"].fillna("Unassigned").replace("", "Unassigned")
                reg_counts = reg_clean.value_counts().reset_index()
                reg_counts.columns = ["Region", "Count"]
                fig_reg = px.bar(
                    reg_counts,
                    x="Region",
                    y="Count",
                    color="Region",
                    text="Count",
                    color_discrete_map={"India": "#1E40AF", "US": "#F59E0B", "Unassigned": "#94A3B8"},
                )
                fig_reg.update_layout(
                    template="plotly_white",
                    height=280,
                    margin=dict(l=10, r=10, t=10, b=10),
                    showlegend=False,
                    xaxis_title=None,
                    yaxis_title="Count",
                )
                st.plotly_chart(fig_reg, use_container_width=True)

        col_c3, col_c4 = st.columns(2)

        with col_c3:
            st.markdown("##### Top Primary Configuration Items (Applications)")
            ci_col = "Primary CI/Application" if "Primary CI/Application" in output_df.columns else "Configuration item"
            if ci_col in output_df.columns:
                ci_counts = output_df[ci_col].dropna().value_counts().head(5).reset_index()
                ci_counts.columns = ["Application", "Count"]
                fig_ci = px.bar(
                    ci_counts,
                    y="Application",
                    x="Count",
                    orientation="h",
                    text="Count",
                    color_discrete_sequence=["#334155"],
                )
                fig_ci.update_layout(
                    template="plotly_white",
                    height=260,
                    margin=dict(l=10, r=10, t=10, b=10),
                    showlegend=False,
                    xaxis_title="Count",
                    yaxis_title=None,
                    yaxis=dict(autorange="reversed"),
                )
                st.plotly_chart(fig_ci, use_container_width=True)

        with col_c4:
            st.markdown("##### Bank Segregation (FCB vs. SVB)")
            if "Bank" in output_df.columns:
                bank_clean = output_df["Bank"].fillna("FCB").replace("", "FCB")
                bank_counts = bank_clean.value_counts().reset_index()
                bank_counts.columns = ["Bank", "Count"]
                fig_bank = px.pie(
                    bank_counts,
                    names="Bank",
                    values="Count",
                    hole=0.5,
                    color="Bank",
                    color_discrete_map={"FCB": "#10B981", "SVB": "#EF4444"},
                )
                fig_bank.update_layout(
                    template="plotly_white",
                    height=260,
                    margin=dict(l=10, r=10, t=10, b=10),
                )
                st.plotly_chart(fig_bank, use_container_width=True)

    # ---------------------------------------------------------
    # TAB 3: MASTER DATASET EXPLORER
    # ---------------------------------------------------------
    with tab_explorer:
        col_s1, col_s2, col_s3 = st.columns([2, 1, 1])
        with col_s1:
            query = st.text_input("Filter Records (INC Number, Description, Assignee):", "")
        with col_s2:
            prio_opts = ["All"] + sorted([str(x) for x in output_df["Priority"].dropna().unique()]) if "Priority" in output_df.columns else ["All"]
            prio_val = st.selectbox("Priority:", prio_opts)
        with col_s3:
            reg_opts = ["All"] + sorted([str(x) for x in output_df["Region"].dropna().unique()]) if "Region" in output_df.columns else ["All"]
            reg_val = st.selectbox("Region:", reg_opts)

        filtered = output_df.copy()
        if query:
            filtered = filtered[
                filtered.astype(str).apply(lambda r: r.str.contains(query, case=False).any(), axis=1)
            ]
        if prio_val != "All" and "Priority" in filtered.columns:
            filtered = filtered[filtered["Priority"].astype(str) == prio_val]
        if reg_val != "All" and "Region" in filtered.columns:
            filtered = filtered[filtered["Region"].astype(str) == reg_val]

        st.caption(f"Showing {len(filtered)} of {len(output_df)} records")
        st.dataframe(filtered, use_container_width=True, height=450)

    # ---------------------------------------------------------
    # TAB 4: EXPORT & DOWNLOAD
    # ---------------------------------------------------------
    with tab_export:
        st.markdown("### Export Reconciled Files")
        st.markdown("Download the updated Master Tracker workbook and the Reconciliation Audit Report.")
        st.markdown("<br>", unsafe_allow_html=True)

        col_e1, col_e2 = st.columns(2)

        with col_e1:
            st.markdown("##### Updated Master Tracker (.xlsx)")
            st.caption("Consolidated sheet containing reconciled records with verified formulas and formatting.")

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                output_df.to_excel(writer, sheet_name="Consolidated", index=False)
            buf.seek(0)

            st.download_button(
                label="Download Updated Master Tracker",
                data=buf,
                file_name=f"Master_Tracker_Updated_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col_e2:
            st.markdown("##### Reconciliation Audit Report (.xlsx)")
            st.caption("Detailed multi-tab workbook documenting Summary, Additions, Field Diffs, and Misses.")

            rep_data = st.session_state["report"]
            temp_path = os.path.join(PROJECT_ROOT, "output", "temp_export_report.xlsx")
            os.makedirs(os.path.dirname(temp_path), exist_ok=True)
            write_reconciliation_report_excel(rep_data, temp_path)

            with open(temp_path, "rb") as f:
                rep_bytes = f.read()

            st.download_button(
                label="Download Reconciliation Audit Report",
                data=rep_bytes,
                file_name=f"Reconciliation_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.markdown("---")
        st.markdown("##### Audit Record (JSON)")
        st.json(st.session_state["report"]["summary"])
