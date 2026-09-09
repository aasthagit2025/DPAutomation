from __future__ import annotations

# DP Automation v1.3 - single-file deployment build
# All processing logic is contained in this file; no core/ folder is required.


# ===== spss_io.py =====

from pathlib import Path
from typing import Any, Dict, Tuple

import pandas as pd
import pyreadstat


def read_sav(path: str | Path) -> Tuple[pd.DataFrame, Any]:
    """Read SPSS .sav preserving user-defined missing values where possible."""
    df, meta = pyreadstat.read_sav(str(path), user_missing=True, apply_value_formats=False)
    return df, meta


def write_sav(
    df: pd.DataFrame,
    path: str | Path,
    column_labels: Dict[str, str],
    variable_value_labels: Dict[str, Dict[Any, str]],
    file_label: str = "DP Automation v1 output",
) -> None:
    """Write an SPSS .sav with variable labels and value labels."""
    pyreadstat.write_sav(
        df,
        str(path),
        file_label=file_label,
        column_labels=column_labels,
        variable_value_labels=variable_value_labels,
    )

# ===== excel_io.py =====

from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, Tuple
import re

import pandas as pd


def _coerce_code(text: str) -> Any:
    s = str(text).strip()
    if s == "":
        return s
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except Exception:
        return s


def _parse_value_labels(text: Any) -> Dict[Any, str]:
    """Parse '1=Male; 2=Female' style metadata cells."""
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return {}
    s = str(text).strip()
    if not s:
        return {}
    out: Dict[Any, str] = {}
    for piece in re.split(r"\s*;\s*", s):
        if not piece or "=" not in piece:
            continue
        code, label = piece.split("=", 1)
        out[_coerce_code(code)] = label.strip()
    return out


def read_excel(path: str | Path) -> Tuple[pd.DataFrame, Any, str]:
    """Read an Excel survey file and synthesize SPSS-like metadata.

    Preferred workbook layout:
      - Raw_Data (or first sheet): respondent-level data, headers = variable names
      - SPSS_Metadata_Reference (optional): columns Variable, Variable Label, Value Labels

    If metadata is absent, labels default to variable names and value labels remain blank.
    """
    xls = pd.ExcelFile(path)
    preferred_data = next((s for s in xls.sheet_names if s.lower() in {"raw_data", "data", "raw data"}), xls.sheet_names[0])
    df = pd.read_excel(path, sheet_name=preferred_data)
    df.columns = [str(c).strip() for c in df.columns]

    column_labels: Dict[str, str] = {c: c for c in df.columns}
    value_labels: Dict[str, Dict[Any, str]] = {}

    meta_sheet = next(
        (s for s in xls.sheet_names if s.lower() in {
            "spss_metadata_reference", "metadata", "variable_metadata", "variable dictionary", "dictionary"
        }),
        None,
    )

    if meta_sheet:
        mdf = pd.read_excel(path, sheet_name=meta_sheet)
        normalized = {str(c).strip().lower(): c for c in mdf.columns}
        var_col = normalized.get("variable") or normalized.get("original_variable") or normalized.get("var")
        label_col = normalized.get("variable label") or normalized.get("variable_label") or normalized.get("label")
        values_col = normalized.get("value labels") or normalized.get("value_labels") or normalized.get("values")

        if var_col:
            for _, row in mdf.iterrows():
                var = row.get(var_col)
                if pd.isna(var):
                    continue
                var = str(var).strip()
                if var not in df.columns:
                    continue
                if label_col and not pd.isna(row.get(label_col)):
                    column_labels[var] = str(row.get(label_col)).strip()
                if values_col and not pd.isna(row.get(values_col)):
                    parsed = _parse_value_labels(row.get(values_col))
                    if parsed:
                        value_labels[var] = parsed

    # Create a minimal metadata object compatible with the existing code.
    meta = SimpleNamespace(
        column_names_to_labels=column_labels,
        variable_value_labels=value_labels,
        readstat_variable_types={c: str(df[c].dtype) for c in df.columns},
        original_variable_types={c: str(df[c].dtype) for c in df.columns},
        missing_ranges={},
    )
    return df, meta, preferred_data


def write_excel(
    df: pd.DataFrame,
    path: str | Path,
    column_labels: Dict[str, str],
    variable_value_labels: Dict[str, Dict[Any, str]],
) -> None:
    """Write processed data plus metadata to an Excel workbook."""
    meta_rows = []
    for var in df.columns:
        labels = variable_value_labels.get(var, {}) or {}
        value_text = "; ".join(f"{k}={v}" for k, v in labels.items())
        meta_rows.append({
            "Variable": var,
            "Variable Label": column_labels.get(var, ""),
            "Value Labels": value_text,
            "Data Type": str(df[var].dtype),
        })
    meta_df = pd.DataFrame(meta_rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Final_Data", index=False)
        meta_df.to_excel(writer, sheet_name="Metadata", index=False)

# ===== metadata.py =====

from typing import Any, Dict, List

import pandas as pd


def _safe_attr(meta: Any, name: str, default):
    value = getattr(meta, name, default)
    return default if value is None else value


def build_metadata(df: pd.DataFrame, meta: Any) -> pd.DataFrame:
    """Flatten SPSS metadata into a review-friendly table."""
    labels: Dict[str, str] = _safe_attr(meta, "column_names_to_labels", {})
    value_labels: Dict[str, Dict[Any, str]] = _safe_attr(meta, "variable_value_labels", {})
    readstat_types: Dict[str, str] = _safe_attr(meta, "readstat_variable_types", {})
    original_types: Dict[str, str] = _safe_attr(meta, "original_variable_types", {})
    missing_ranges: Dict[str, Any] = _safe_attr(meta, "missing_ranges", {})

    rows: List[dict] = []
    for var in df.columns:
        vlabels = value_labels.get(var, {}) or {}
        base = {
            "original_variable": var,
            "variable_label": labels.get(var, "") or "",
            "storage_type": readstat_types.get(var, str(df[var].dtype)),
            "original_format": original_types.get(var, ""),
            "missing_definition": str(missing_ranges.get(var, "")) if var in missing_ranges else "",
            "non_null_n": int(df[var].notna().sum()),
            "unique_n": int(df[var].nunique(dropna=True)),
        }
        if vlabels:
            for code, lab in vlabels.items():
                rows.append({**base, "value": code, "value_label": lab})
        else:
            rows.append({**base, "value": "", "value_label": ""})
    return pd.DataFrame(rows)


def mapping_template_from_metadata(metadata_df: pd.DataFrame) -> pd.DataFrame:
    """Create a mapping template users can edit and re-upload."""
    cols = [
        "original_variable",
        "new_variable",
        "new_variable_label",
        "old_value",
        "new_value",
        "new_value_label",
        "action",
        "notes",
    ]
    rows = []
    for _, r in metadata_df.iterrows():
        rows.append({
            "original_variable": r["original_variable"],
            "new_variable": r["original_variable"],
            "new_variable_label": r["variable_label"],
            "old_value": r["value"],
            "new_value": r["value"],
            "new_value_label": r["value_label"],
            "action": "KEEP",
            "notes": "",
        })
    return pd.DataFrame(rows, columns=cols)

# ===== auto_mapper.py =====

import html
import re
from collections import defaultdict
from typing import Any, Dict, List, Tuple

import pandas as pd


DK_PATTERNS = [
    r"\bdon'?t\s*know\b", r"\bdo not know\b", r"\bdk\b",
    r"\brefus(?:ed|al)?\b", r"prefer not to (?:say|answer)",
    r"\bnot applicable\b", r"\bn/?a\b"
]


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    s = html.unescape(str(value))
    s = re.sub(r"<[^>]+>", " ", s)
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def standardize_value_label(label: Any) -> str:
    s = clean_text(label)
    low = s.lower().replace("’", "'")
    if re.fullmatch(r"dk|don't know|do not know", low):
        return "Don't know"
    if re.fullmatch(r"refused|refusal", low):
        return "Refused"
    if low in {"prefer not to say", "prefer not to answer"}:
        return "Prefer not to answer"
    if low in {"n/a", "na", "not applicable"}:
        return "Not applicable"
    return s


def _question_prefix(name: str) -> Tuple[str, List[int]]:
    """Parse common Sawtooth-style names while preserving unknown names safely."""
    n = str(name).strip()
    # Q7_1_2 / Q7r1c2 / Q7.1.2 style-ish structures
    m = re.match(r"^([A-Za-z]+\d+[A-Za-z]?)(?:[_\.](\d+))(?:[_\.](\d+))$", n)
    if m:
        return m.group(1), [int(m.group(2)), int(m.group(3))]
    m = re.match(r"^([A-Za-z]+\d+[A-Za-z]?)[_\.](\d+)$", n)
    if m:
        return m.group(1), [int(m.group(2))]
    m = re.match(r"^([A-Za-z]+\d+[A-Za-z]?)$", n)
    if m:
        return m.group(1), []
    return n, []


def propose_variable_name(name: str) -> Tuple[str, float, str]:
    """Return standardized DP name, confidence, and rationale.

    Conservative by design: recognizable question suffixes are standardized;
    unknown/system fields are retained.
    """
    raw = str(name).strip()
    prefix, idx = _question_prefix(raw)
    if len(idx) == 2:
        return f"{prefix}r{idx[0]}c{idx[1]}", 0.96, "Two-dimensional question suffix detected"
    if len(idx) == 1:
        return f"{prefix}r{idx[0]}", 0.90, "Single question-item suffix detected"
    if prefix != raw or re.match(r"^[A-Za-z]+\d+[A-Za-z]?$", raw):
        return raw, 0.99, "Already question-like"
    # Preserve system / unfamiliar variables rather than inventing semantics.
    safe = re.sub(r"[^A-Za-z0-9_@#$]", "_", raw)
    safe = re.sub(r"_+", "_", safe).strip("_") or "VAR"
    if not re.match(r"^[A-Za-z@#$]", safe):
        safe = "V_" + safe
    return safe[:64], 0.65, "Unrecognized structure; conservative cleanup only"


def classify_variable(var: str, series: pd.Series, value_labels: Dict[Any, str], sibling_count: int) -> Tuple[str, float]:
    labels = [clean_text(x).lower() for x in (value_labels or {}).values()]
    vals = set(series.dropna().unique().tolist())

    if pd.api.types.is_string_dtype(series) or series.dtype == object:
        if not value_labels:
            return "OPEN_END", 0.95
    if sibling_count > 1 and vals and vals.issubset({0, 1, 0.0, 1.0}):
        return "MULTI_RESPONSE", 0.97
    if sibling_count > 1:
        return "GRID_ITEM", 0.82
    if value_labels:
        n = len(value_labels)
        if n in {4, 5, 7, 10, 11}:
            return "RATING_OR_SINGLE", 0.72
        return "SINGLE_CODE", 0.82
    if pd.api.types.is_numeric_dtype(series):
        return "NUMERIC", 0.70
    return "OTHER", 0.55


def is_dk_refused(label: str) -> bool:
    low = clean_text(label).lower().replace("’", "'")
    return any(re.search(p, low, flags=re.I) for p in DK_PATTERNS)


def generate_auto_mapping(
    df: pd.DataFrame,
    column_labels: Dict[str, str],
    value_labels: Dict[str, Dict[Any, str]],
) -> pd.DataFrame:
    """Generate the complete internal transformation plan from SPSS metadata.

    Codes are preserved unless a deterministic future rule explicitly requests a
    recode. This avoids semantic recodes based on guesses while eliminating the
    external mapping-file workflow.
    """
    groups: Dict[str, List[str]] = defaultdict(list)
    for var in df.columns:
        prefix, _ = _question_prefix(var)
        groups[prefix].append(var)

    rows: List[dict] = []
    for var in df.columns:
        prefix, _ = _question_prefix(var)
        new_var, name_conf, name_reason = propose_variable_name(var)
        raw_label = column_labels.get(var, "") or ""
        new_var_label = clean_text(raw_label)
        vlabels = value_labels.get(var, {}) or {}
        qtype, type_conf = classify_variable(var, df[var], vlabels, len(groups[prefix]))
        confidence = min(name_conf, type_conf if qtype != "OTHER" else name_conf)

        action_bits = []
        if new_var != var:
            action_bits.append("RENAME")
        if new_var_label != raw_label:
            action_bits.append("RELABEL")

        if vlabels:
            for old_value, old_label in vlabels.items():
                new_value_label = standardize_value_label(old_label)
                row_action = list(action_bits)
                if new_value_label != str(old_label):
                    row_action.append("VALUE_RELABEL")
                rows.append({
                    "original_variable": var,
                    "new_variable": new_var,
                    "original_variable_label": raw_label,
                    "new_variable_label": new_var_label,
                    "question_type": qtype,
                    "old_value": old_value,
                    "new_value": old_value,
                    "old_value_label": old_label,
                    "new_value_label": new_value_label,
                    "dk_refused_flag": "YES" if is_dk_refused(str(old_label)) else "",
                    "action": "+".join(row_action) if row_action else "KEEP",
                    "confidence": round(confidence, 2),
                    "needs_review": "YES" if confidence < 0.80 else "",
                    "reason": name_reason,
                    "notes": "",
                })
        else:
            rows.append({
                "original_variable": var,
                "new_variable": new_var,
                "original_variable_label": raw_label,
                "new_variable_label": new_var_label,
                "question_type": qtype,
                "old_value": "",
                "new_value": "",
                "old_value_label": "",
                "new_value_label": "",
                "dk_refused_flag": "",
                "action": "+".join(action_bits) if action_bits else "KEEP",
                "confidence": round(confidence, 2),
                "needs_review": "YES" if confidence < 0.80 else "",
                "reason": name_reason,
                "notes": "",
            })

    plan = pd.DataFrame(rows)

    # Flag rename collisions for review instead of applying silently.
    var_plan = plan[["original_variable", "new_variable"]].drop_duplicates()
    collisions = var_plan.groupby("new_variable")["original_variable"].nunique()
    bad_names = set(collisions[collisions > 1].index)
    if bad_names:
        mask = plan["new_variable"].isin(bad_names)
        plan.loc[mask, "needs_review"] = "YES"
        plan.loc[mask, "confidence"] = plan.loc[mask, "confidence"].clip(upper=0.40)
        plan.loc[mask, "reason"] = plan.loc[mask, "reason"] + "; rename collision"

    return plan


def auto_plan_to_apply_mapping(plan: pd.DataFrame) -> pd.DataFrame:
    """Convert the rich auto-plan into the existing transformation-engine schema."""
    out = pd.DataFrame({
        "original_variable": plan["original_variable"],
        "new_variable": plan["new_variable"],
        "new_variable_label": plan["new_variable_label"],
        "old_value": plan["old_value"],
        "new_value": plan["new_value"],
        "new_value_label": plan["new_value_label"],
        "action": "KEEP",
        "notes": plan.get("notes", ""),
    })

    # The existing engine combines variable-level rename/relabel with value-label updates.
    # Explicit code changes are treated as recodes.
    for i, r in plan.iterrows():
        rename = str(r["new_variable"]) != str(r["original_variable"])
        relabel = clean_text(r["new_variable_label"]) != clean_text(r["original_variable_label"])
        old_v, new_v = r["old_value"], r["new_value"]
        recode = str(old_v) != str(new_v) and str(old_v).strip() != "" and str(new_v).strip() != ""
        if recode and rename:
            action = "RENAME_RECODE"
        elif recode:
            action = "RECODE"
        elif rename:
            action = "RENAME"
        elif relabel:
            action = "RELABEL"
        else:
            action = "KEEP"
        out.at[i, "action"] = action
    return out

# ===== transform.py =====

import math
import re
from collections import defaultdict
from typing import Any, Dict, Iterable, List, Tuple

import pandas as pd

VALID_ACTIONS = {"KEEP", "RENAME", "RELABEL", "RECODE", "RENAME_RECODE", "DROP"}


def clean_variable_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r"[^A-Za-z0-9_@#$]", "_", name)
    name = re.sub(r"_+", "_", name)
    if not name:
        name = "VAR"
    if not re.match(r"^[A-Za-z@#$]", name):
        name = "V_" + name
    return name[:64]


def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    return str(v).strip() == ""


def _coerce_for_series(v: Any, series: pd.Series) -> Any:
    if _is_blank(v):
        return v
    if pd.api.types.is_numeric_dtype(series):
        try:
            f = float(v)
            return int(f) if f.is_integer() else f
        except Exception:
            return v
    return str(v)


def _coerce_like_value(v: Any, sample: Any) -> Any:
    if _is_blank(v):
        return v
    if isinstance(sample, (int, float)) and not isinstance(sample, bool):
        try:
            f = float(v)
            return int(f) if f.is_integer() else f
        except Exception:
            return v
    return str(v)


def validate_mapping(mapping: pd.DataFrame, existing_columns: Iterable[str]) -> List[str]:
    warnings: List[str] = []
    required = {
        "original_variable", "new_variable", "new_variable_label", "old_value",
        "new_value", "new_value_label", "action", "notes"
    }
    missing = required - set(mapping.columns)
    if missing:
        return [f"Mapping file is missing required column(s): {', '.join(sorted(missing))}"]

    existing = set(existing_columns)
    bad_vars = sorted(set(mapping["original_variable"].dropna().astype(str)) - existing)
    if bad_vars:
        warnings.append(f"Mapping contains variables not found in the SAV: {', '.join(bad_vars[:10])}")

    actions = mapping["action"].fillna("KEEP").astype(str).str.upper().str.strip()
    invalid = sorted(set(actions) - VALID_ACTIONS)
    if invalid:
        warnings.append(f"Invalid action(s): {', '.join(invalid)}")

    # Proposed final names by variable.
    proposals = {}
    for var, g in mapping.groupby("original_variable", dropna=False):
        if _is_blank(var):
            continue
        vals = [clean_variable_name(x) for x in g["new_variable"].dropna().astype(str) if str(x).strip()]
        proposals[str(var)] = vals[0] if vals else str(var)
        if len(set(vals)) > 1:
            warnings.append(f"{var}: multiple different new_variable values were supplied.")

    seen = defaultdict(list)
    for old, new in proposals.items():
        seen[new].append(old)
    for new, olds in seen.items():
        if len(olds) > 1:
            warnings.append(f"Duplicate final variable name '{new}' proposed for: {', '.join(olds)}")

    return warnings


def apply_mapping(
    df: pd.DataFrame,
    original_column_labels: Dict[str, str],
    original_value_labels: Dict[str, Dict[Any, str]],
    mapping: pd.DataFrame,
) -> Tuple[pd.DataFrame, Dict[str, str], Dict[str, Dict[Any, str]], pd.DataFrame, List[str]]:
    """Apply controlled rename/relabel/recode/drop transformations."""
    out = df.copy()
    column_labels = dict(original_column_labels or {})
    value_labels = {k: dict(v or {}) for k, v in (original_value_labels or {}).items()}
    audit: List[dict] = []
    warnings = validate_mapping(mapping, df.columns)
    if any(w.startswith("Mapping file is missing") for w in warnings):
        return out, column_labels, value_labels, pd.DataFrame(), warnings

    mapping = mapping.copy()
    mapping["action"] = mapping["action"].fillna("KEEP").astype(str).str.upper().str.strip()

    # Process each original variable as an atomic unit.
    rename_map: Dict[str, str] = {}
    drop_vars: List[str] = []

    for var, g in mapping.groupby("original_variable", sort=False):
        var = str(var)
        if var not in out.columns:
            continue
        actions = set(g["action"])
        if "DROP" in actions:
            drop_vars.append(var)
            audit.append({"variable": var, "action": "DROP", "detail": "Variable dropped"})
            continue

        proposed_names = [clean_variable_name(x) for x in g["new_variable"].dropna().astype(str) if x.strip()]
        new_var = proposed_names[0] if proposed_names else var

        if new_var != var or actions & {"RENAME", "RENAME_RECODE"}:
            rename_map[var] = new_var
            audit.append({"variable": var, "action": "RENAME", "detail": f"{var} -> {new_var}"})

        # Label: take first non-blank provided label.
        provided_labels = [str(x) for x in g["new_variable_label"].dropna() if str(x).strip()]
        if provided_labels:
            new_lab = provided_labels[0].strip()
            if new_lab != (column_labels.get(var, "") or ""):
                column_labels[var] = new_lab
                audit.append({"variable": var, "action": "RELABEL", "detail": new_lab})

        # Recode rows with explicit old/new values and RECODE-type action.
        recode_rows = g[g["action"].isin(["RECODE", "RENAME_RECODE"])]
        if not recode_rows.empty:
            recode_dict = {}
            new_label_dict = {}
            old_vlabels = value_labels.get(var, {}) or {}
            for _, r in recode_rows.iterrows():
                if _is_blank(r["old_value"]) or _is_blank(r["new_value"]):
                    continue
                old = _coerce_for_series(r["old_value"], out[var])
                new = _coerce_for_series(r["new_value"], out[var])
                recode_dict[old] = new
                lab = "" if _is_blank(r["new_value_label"]) else str(r["new_value_label"]).strip()
                if lab:
                    new_label_dict[new] = lab
                elif old in old_vlabels:
                    new_label_dict[new] = old_vlabels[old]

            if recode_dict:
                before_nonmissing = int(out[var].notna().sum())
                out[var] = out[var].replace(recode_dict)
                after_nonmissing = int(out[var].notna().sum())
                value_labels[var] = {**{k: v for k, v in old_vlabels.items() if k not in recode_dict}, **new_label_dict}
                audit.append({
                    "variable": var,
                    "action": "RECODE",
                    "detail": f"{recode_dict} | nonmissing {before_nonmissing}->{after_nonmissing}"
                })
        else:
            # Value-label-only updates without code changes.
            current = dict(value_labels.get(var, {}) or {})
            changed = False
            for _, r in g.iterrows():
                if _is_blank(r["old_value"]) or _is_blank(r["new_value_label"]):
                    continue
                old = _coerce_for_series(r["old_value"], out[var])
                new_lab = str(r["new_value_label"]).strip()
                if current.get(old) != new_lab:
                    current[old] = new_lab
                    changed = True
            if changed:
                value_labels[var] = current
                audit.append({"variable": var, "action": "VALUE_RELABEL", "detail": "Value labels updated"})

    if drop_vars:
        out = out.drop(columns=[v for v in drop_vars if v in out.columns])
        for v in drop_vars:
            column_labels.pop(v, None)
            value_labels.pop(v, None)

    # Validate rename collisions before applying.
    final_cols = [rename_map.get(c, c) for c in out.columns]
    if len(final_cols) != len(set(final_cols)):
        warnings.append("Rename not applied because it would create duplicate variable names.")
        rename_map = {}

    if rename_map:
        out = out.rename(columns=rename_map)
        column_labels = {rename_map.get(k, k): v for k, v in column_labels.items() if k not in drop_vars}
        value_labels = {rename_map.get(k, k): v for k, v in value_labels.items() if k not in drop_vars}

    # Ensure every output column has a label key.
    column_labels = {c: column_labels.get(c, "") or "" for c in out.columns}
    value_labels = {c: value_labels.get(c, {}) or {} for c in out.columns if value_labels.get(c)}

    audit_df = pd.DataFrame(audit, columns=["variable", "action", "detail"])
    return out, column_labels, value_labels, audit_df, warnings

# ===== qc.py =====

from typing import Any, Dict, List

import pandas as pd


def run_qc(df: pd.DataFrame, column_labels: Dict[str, str], value_labels: Dict[str, Dict[Any, str]]) -> pd.DataFrame:
    issues: List[dict] = []

    if df.columns.duplicated().any():
        for c in df.columns[df.columns.duplicated()].tolist():
            issues.append({"severity": "ERROR", "variable": c, "issue": "Duplicate variable name"})

    for var in df.columns:
        if not (column_labels.get(var, "") or "").strip():
            issues.append({"severity": "WARNING", "variable": var, "issue": "Blank variable label"})

        vlabels = value_labels.get(var, {}) or {}
        if vlabels and pd.api.types.is_numeric_dtype(df[var]):
            observed = set(df[var].dropna().unique().tolist())
            labelled = set(vlabels.keys())
            unlabelled = observed - labelled
            if unlabelled:
                preview = ", ".join(map(str, list(unlabelled)[:10]))
                issues.append({
                    "severity": "WARNING", "variable": var,
                    "issue": f"Observed value(s) without value labels: {preview}"
                })

    if not issues:
        issues.append({"severity": "OK", "variable": "", "issue": "No structural QC issues detected"})
    return pd.DataFrame(issues)

# ===== STREAMLIT UI =====

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(page_title="DP Automation v1.3", page_icon="⚙️", layout="wide")

st.title("⚙️ DP Automation v1.3 — SPSS + Excel Zero-Mapping")
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
