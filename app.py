from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from core.spss_io import read_sav, write_sav
from core.excel_io import read_excel, write_excel
from core.metadata import build_metadata
from core.auto_mapper import generate_auto_mapping, auto_plan_to_apply_mapping
from core.transform import apply_mapping
from core.qc import run_qc

st.set_page_config(page_title="DP Automation v1.2", page_icon="⚙️", layout="wide")

st.title("⚙️ DP Automation v1.2 — SPSS + Excel Zero-Mapping")
st.caption("Upload raw Sawtooth SPSS or Excel → auto-map → review exceptions → rename/relabel/recode → QC → final SAV/Excel/CSV")

with st.expander("How this version works", expanded=False):
    st.markdown(
        """
**No external mapping file is required.** The app generates its own transformation plan.

**Supported inputs**
- SPSS `.sav`: reads variable names, variable labels and value labels directly.
- Excel `.xlsx`: uses respondent data from `Raw_Data` / `Data` / the first sheet. If a `SPSS_Metadata_Reference` or `Metadata` sheet is present, the app also reads variable labels and value labels from it.

**Workflow**
1. Upload a `.sav` or `.xlsx` file.
2. The app detects likely question structures, proposed DP variable names, cleaned labels, DK/Refused labels and question types.
3. Review only rows flagged **Needs Review**. Edit names, labels or `new_value` inside the browser when required.
4. Generate final SPSS, Excel and CSV plus the auto-mapping plan, metadata, audit log and QC report.

Codes are preserved by default. Semantic recodes are never guessed silently.
        """
    )

uploaded = st.file_uploader("1. Upload raw survey data (.sav or .xlsx)", type=["sav", "xlsx"])

if not uploaded:
    st.info("Upload a `.sav` or `.xlsx` file to begin.")
    st.stop()

with tempfile.TemporaryDirectory() as tmpdir:
    suffix = Path(uploaded.name).suffix.lower()
    raw_path = Path(tmpdir) / f"raw_input{suffix}"
    raw_path.write_bytes(uploaded.getvalue())

    try:
        if suffix == ".sav":
            df, meta = read_sav(raw_path)
            source_detail = "SPSS metadata"
        else:
            df, meta, data_sheet = read_excel(raw_path)
            source_detail = f"Excel sheet: {data_sheet}"
    except Exception as e:
        st.error(f"Could not read the input file: {e}")
        st.stop()

    col_labels = getattr(meta, "column_names_to_labels", {}) or {}
    val_labels = getattr(meta, "variable_value_labels", {}) or {}

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Respondents", f"{len(df):,}")
    c2.metric("Variables", f"{len(df.columns):,}")
    c3.metric("Input type", "SPSS" if suffix == ".sav" else "Excel")
    c4.metric("Metadata source", source_detail)

    if suffix == ".xlsx" and not val_labels:
        st.warning("This Excel workbook does not contain a recognized metadata sheet with value labels. Auto-mapping will use variable names and observed data patterns; label-related confidence may be lower.")

    metadata_df = build_metadata(df, meta)
    plan_df = generate_auto_mapping(df, col_labels, val_labels)

    st.subheader("2. Automatic DP mapping")
    summary_vars = plan_df[["original_variable", "new_variable", "question_type", "confidence", "needs_review"]].drop_duplicates()
    renamed_n = int((summary_vars["original_variable"] != summary_vars["new_variable"]).sum())
    review_n = int((summary_vars["needs_review"] == "YES").sum())
    dk_n = int((plan_df["dk_refused_flag"] == "YES").sum())

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Proposed renames", renamed_n)
    s2.metric("Needs review", review_n)
    s3.metric("DK/Refused labels", dk_n)
    s4.metric("Mapping rows", len(plan_df))

    review_only = st.toggle("Show only rows needing review", value=review_n > 0)
    view_df = plan_df[plan_df["needs_review"] == "YES"].copy() if review_only else plan_df.copy()

    st.caption("Edit proposed names/labels or `new_value` directly here. No external mapping file is needed.")
    edited_view = st.data_editor(
        view_df,
        use_container_width=True,
        height=430,
        num_rows="fixed",
        disabled=[
            "original_variable", "original_variable_label", "question_type", "old_value",
            "old_value_label", "dk_refused_flag", "confidence", "needs_review", "reason", "action"
        ],
        key="auto_mapping_editor",
    )

    edited_plan = plan_df.copy()
    if review_only:
        edited_plan.loc[edited_view.index, edited_view.columns] = edited_view
    else:
        edited_plan = edited_view.copy()

    st.download_button(
        "⬇️ Download auto-generated mapping/audit plan (optional)",
        edited_plan.to_csv(index=False).encode("utf-8-sig"),
        file_name="Project_Auto_Mapping_Plan.csv",
        mime="text/csv",
    )

    apply_df = auto_plan_to_apply_mapping(edited_plan)
    processed_df, final_col_labels, final_val_labels, audit_df, warnings = apply_mapping(
        df, col_labels, val_labels, apply_df
    )

    st.subheader("3. Transformation + QC preview")
    if warnings:
        for w in warnings:
            st.warning(w)
    else:
        st.success("Automatic transformation plan passed structural validation.")

    q1, q2, q3 = st.columns(3)
    q1.metric("Final variables", len(processed_df.columns))
    q2.metric("Recorded changes", len(audit_df))
    q3.metric("Review exceptions", int((edited_plan["needs_review"] == "YES").sum()))

    left, right = st.columns(2)
    with left:
        st.markdown("**Audit preview**")
        st.dataframe(audit_df.head(400), use_container_width=True, height=300)
    with right:
        st.markdown("**Processed data preview**")
        st.dataframe(processed_df.head(100), use_container_width=True, height=300)

    qc_df = run_qc(processed_df, final_col_labels, final_val_labels)
    st.markdown("**Automated QC**")
    st.dataframe(qc_df, use_container_width=True, height=260)

    blocking = [w for w in warnings if "duplicate" in w.lower() or "missing required" in w.lower()]
    confirm = st.checkbox("I reviewed the exceptions/warnings and approve generation of the final files.")

    if st.button("4. Generate Final Files", type="primary", disabled=(not confirm or bool(blocking))):
        out_sav = Path(tmpdir) / "Project_Final.sav"
        out_xlsx = Path(tmpdir) / "Project_Final.xlsx"

        try:
            write_sav(processed_df, out_sav, final_col_labels, final_val_labels)
        except Exception as e:
            st.error(f"SAV export failed: {e}")
            st.stop()

        try:
            write_excel(processed_df, out_xlsx, final_col_labels, final_val_labels)
        except Exception as e:
            st.error(f"Excel export failed: {e}")
            st.stop()

        final_meta_rows = []
        for var in processed_df.columns:
            vls = final_val_labels.get(var, {}) or {}
            if vls:
                for code, lab in vls.items():
                    final_meta_rows.append({
                        "variable": var,
                        "variable_label": final_col_labels.get(var, ""),
                        "value": code,
                        "value_label": lab,
                        "dtype": str(processed_df[var].dtype),
                    })
            else:
                final_meta_rows.append({
                    "variable": var,
                    "variable_label": final_col_labels.get(var, ""),
                    "value": "",
                    "value_label": "",
                    "dtype": str(processed_df[var].dtype),
                })
        final_metadata_df = pd.DataFrame(final_meta_rows)

        st.success("Final DP files generated successfully.")
        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button("⬇️ Final SPSS", out_sav.read_bytes(), "Project_Final.sav", "application/octet-stream")
            st.download_button("⬇️ Final Excel", out_xlsx.read_bytes(), "Project_Final.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with d2:
            st.download_button("⬇️ Final CSV", processed_df.to_csv(index=False).encode("utf-8-sig"), "Project_Final.csv", "text/csv")
            st.download_button("⬇️ Final Metadata", final_metadata_df.to_csv(index=False).encode("utf-8-sig"), "Project_Metadata.csv", "text/csv")
        with d3:
            st.download_button("⬇️ Audit Log", audit_df.to_csv(index=False).encode("utf-8-sig"), "Project_Audit_Log.csv", "text/csv")
            st.download_button("⬇️ QC Report", qc_df.to_csv(index=False).encode("utf-8-sig"), "Project_QC_Report.csv", "text/csv")
            st.download_button("⬇️ Auto Mapping Plan", edited_plan.to_csv(index=False).encode("utf-8-sig"), "Project_Auto_Mapping_Plan.csv", "text/csv")
