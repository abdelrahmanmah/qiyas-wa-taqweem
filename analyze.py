"""
analyze.py — Redaa2 survey analysis CLI (standalone Python tool)

Reads a YAML schema + Excel file and outputs the same result JSON that the
React frontend produces, making it usable for batch processing or scripting.

Usage:
  python analyze.py <schema.yaml> <survey.xlsx> [--out result.json]
  python analyze.py <schemas/> <survey.xlsx>     # auto-detect schema
"""
import sys
import re
import json
import unicodedata
import argparse
from pathlib import Path
import pandas as pd
import yaml


# ── Schema compilation (mirrors src/schemas/compileSchema.js) ─────────────────

SCALE5_VALUES = [
    {"code": "1", "label": "لا أوافق بشدة", "score": 1},
    {"code": "2", "label": "لا أوافق",      "score": 2},
    {"code": "3", "label": "محايد",          "score": 3},
    {"code": "4", "label": "أوافق",          "score": 4},
    {"code": "5", "label": "أوافق بشدة",     "score": 5},
]

SCALE3_VALUES = [
    {"code": "agree",    "label": "أوافق",    "score": 3},
    {"code": "neutral",  "label": "محايد",    "score": 2},
    {"code": "disagree", "label": "لا أوافق", "score": 1},
]

INTERP5 = [
    {"min": 4.5, "label": "أوافق بشدة",  "tier": "excellent", "color": "0d6e3a"},
    {"min": 3.5, "label": "أوافق",       "tier": "good",      "color": "1a5276"},
    {"min": 2.5, "label": "محايد",       "tier": "neutral",   "color": "784212"},
    {"min": 1.0, "label": "لا أوافق",   "tier": "low",       "color": "922b21"},
]

INTERP3 = [
    {"min": 85, "label": "موافقة قوية",  "tier": "excellent", "color": "0d6e3a"},
    {"min": 70, "label": "موافقة",       "tier": "good",      "color": "1a5276"},
    {"min": 50, "label": "محايد",        "tier": "neutral",   "color": "784212"},
    {"min":  0, "label": "عدم الموافقة", "tier": "low",       "color": "922b21"},
]


def _as_list(v):
    if v is None:
        return []
    return v if isinstance(v, list) else [v]


def compile_schema(raw: dict) -> dict:
    is5 = raw["scale"] == 5
    scale = {
        "type": "likert-5",
        "values": SCALE5_VALUES,
        "agreementCodes": ["4", "5"],
        "tokenPattern": r"\((\d)\)",
    } if is5 else {
        "type": "likert-3",
        "values": SCALE3_VALUES,
        "agreementCodes": ["agree"],
        "detectFn": "arabic3point",
    }

    m = raw.get("meta") or {}
    metadata = {
        "timestampCol":  _as_list(m.get("timestamp")) or ["Timestamp"],
        "freeTextCols":  _as_list(m.get("freetext")),
    }
    if m.get("level"):   metadata["levelCol"]      = _as_list(m["level"])
    if m.get("email"):   metadata["emailCol"]      = _as_list(m["email"])
    if m.get("degree"):  metadata["degreeCol"]     = _as_list(m["degree"])
    if m.get("dept"):    metadata["departmentCol"] = _as_list(m["dept"])
    if m.get("program"): metadata["programCol"]    = _as_list(m["program"])

    # Build axes
    axes_src = raw.get("axes") or [{"name": raw["name"], "questions": raw["questions"],
                                     "recommendation": raw.get("recommendation")}]
    global_seq = 1
    auto_col = 0
    axes_out = []
    for ai, ax in enumerate(axes_src):
        ax_id = f"ax{ai+1:02d}"
        qs_out = []
        for q in ax["questions"]:
            text = q if isinstance(q, str) else q["text"]
            q_id = f"q{global_seq:02d}"
            entry = {"id": q_id, "seq": global_seq, "text": text}
            if is5:
                if isinstance(q, dict) and q.get("col") is not None:
                    entry["colIndex"] = q["col"]
                    auto_col = q["col"] + 1
                else:
                    entry["colIndex"] = auto_col
                    auto_col += 1
            qs_out.append(entry)
            global_seq += 1
        ax_entry = {"id": ax_id, "name": ax["name"], "questions": qs_out}
        if ax.get("recommendation"):
            ax_entry["recommendation"] = ax["recommendation"]
        axes_out.append(ax_entry)

    # questionStartIndex for likert-3
    q_start = raw.get("questionStartIndex")
    if not is5 and q_start is None:
        q_start = 1 + sum(bool(m.get(k)) for k in ("email", "degree", "dept", "program"))

    schema = {
        "id": raw["id"],
        "label": raw["name"],
        "fileHints": _as_list(raw.get("hints")),
        "scale": scale,
        "metadata": metadata,
        "interpretation": raw.get("interpretation") or (INTERP5 if is5 else INTERP3),
        "axes": axes_out,
    }
    if not is5:
        schema["questionStartIndex"] = q_start
    return schema


def load_schema(path: Path) -> dict:
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    return compile_schema(raw)


# ── Arabic normalization (mirrors analyze.js normalize / similarity) ───────────

def normalize(v: str) -> str:
    if not v:
        return ""
    v = str(v).strip()
    v = re.sub(r"[ً-ٰٟ]", "", v)       # strip tashkeel
    v = re.sub(r"[أإآ]", "ا", v)
    v = re.sub(r"ى", "ي", v)
    v = re.sub(r"ة", "ه", v)
    v = re.sub(r"\s+", " ", v)
    return v.lower()


def similarity(a: str, b: str) -> float:
    na, nb = normalize(a), normalize(b)
    if not na or not nb:
        return 0.0
    sa, sb = set(na), set(nb)
    overlap = len(sa & sb) / max(len(sa | sb), 1)
    bonus = 0.2 if (na in nb or nb in na) else 0.0
    return min(1.0, overlap + bonus)


# ── Schema auto-detection ─────────────────────────────────────────────────────

def detect_schema(schemas_dir: Path, filename: str) -> dict:
    schemas = [load_schema(p) for p in sorted(schemas_dir.glob("*.yaml"))]
    fname = normalize(filename)

    # Pass 1: filename hints
    for s in schemas:
        if any(normalize(h) in fname for h in s["fileHints"]):
            return s

    # Pass 2: header similarity (load headers from Excel)
    raise ValueError(
        f"Could not auto-detect schema for '{filename}'. "
        "Pass an explicit schema path instead."
    )


# ── Column mapping ────────────────────────────────────────────────────────────

def _find_col(headers: list, keywords: list) -> int | None:
    for kw in keywords:
        nkw = normalize(kw)
        for i, h in enumerate(headers):
            if nkw in normalize(str(h)):
                return i
    return None


def map_columns(df: pd.DataFrame, schema: dict) -> dict:
    """Returns {question_id: column_name} mapping."""
    headers = list(df.columns)
    meta = schema["metadata"]
    scale_type = schema["scale"]["type"]
    mapping = {}

    if scale_type == "likert-5":
        # Exclude metadata columns to build the filtered index list
        exclude_keywords = (
            meta.get("timestampCol", []) +
            meta.get("levelCol", []) +
            meta.get("degreeCol", []) +
            meta.get("departmentCol", []) +
            meta.get("emailCol", []) +
            meta.get("freeTextCols", [])
        )
        def is_meta(h):
            nh = normalize(str(h))
            return any(normalize(kw) in nh for kw in exclude_keywords)

        filtered = [h for h in headers if not is_meta(h)]
        for ax in schema["axes"]:
            for q in ax["questions"]:
                idx = q["colIndex"]
                if idx < len(filtered):
                    mapping[q["id"]] = filtered[idx]

    else:  # likert-3 — fuzzy header matching
        q_start = schema.get("questionStartIndex", 0)
        candidate_headers = headers[q_start:]
        all_questions = [q for ax in schema["axes"] for q in ax["questions"]]
        used = set()
        for q in all_questions:
            best_col, best_score = None, 0.0
            for h in candidate_headers:
                if h in used:
                    continue
                sc = similarity(q["text"], str(h))
                if sc > best_score:
                    best_score, best_col = sc, h
            if best_col and best_score >= 0.35:
                mapping[q["id"]] = best_col
                used.add(best_col)

    return mapping


# ── Response parsing ──────────────────────────────────────────────────────────

_TOKEN5 = re.compile(r"\((\d)\)")

def parse_response5(v) -> str | None:
    s = str(v).strip()
    m = _TOKEN5.search(s)
    if m:
        return m.group(1)
    if s in ("1", "2", "3", "4", "5"):
        return s
    return None


def parse_response3(v) -> str | None:
    nv = normalize(str(v))
    if "لا اوافق" in nv or "لا أوافق" in nv or nv == "disagree":
        return "disagree"
    if "محايد" in nv or nv == "neutral":
        return "neutral"
    if "اوافق" in nv or "أوافق" in nv or nv == "agree":
        return "agree"
    return None


def parse_series(series: pd.Series, scale_type: str) -> pd.Series:
    fn = parse_response5 if scale_type == "likert-5" else parse_response3
    return series.map(fn)


# ── Interpretation ────────────────────────────────────────────────────────────

def interpret(value: float, schema: dict) -> dict:
    for tier in schema["interpretation"]:
        if value >= tier["min"]:
            return {"direction": tier["label"], "tier": tier["tier"], "color": tier["color"]}
    last = schema["interpretation"][-1]
    return {"direction": last["label"], "tier": last["tier"], "color": last["color"]}


# ── Core analysis ─────────────────────────────────────────────────────────────

def analyze_question(series: pd.Series, schema: dict) -> dict:
    scale = schema["scale"]
    codes = [v["code"] for v in scale["values"]]
    agree_codes = set(scale["agreementCodes"])
    parsed = parse_series(series.dropna(), scale["type"])
    valid = parsed.dropna()
    total = len(valid)

    counts = {c: int((valid == c).sum()) for c in codes}
    pcts = {c: round(counts[c] / total * 100, 1) if total else 0.0 for c in codes}

    mean = None
    if scale["type"] == "likert-5":
        score_map = {v["code"]: v["score"] for v in scale["values"]}
        mean = round(sum(counts[c] * score_map[c] for c in codes) / total, 2) if total else 0.0

    agree_pct = round(sum(counts[c] for c in agree_codes) / total * 100, 1) if total else 0.0
    return {"counts": counts, "pcts": pcts, "mean": mean, "agreePct": agree_pct, "total": total}


def bucket_key(v: str) -> str:
    return normalize(str(v))


def count_by_col(df: pd.DataFrame, col: str) -> dict:
    result = {}
    for raw_val in df[col].dropna():
        key = bucket_key(raw_val)
        label = str(raw_val).strip()
        if key not in result:
            result[key] = {"label": label, "count": 0}
        result[key]["count"] += 1
    return {v["label"]: v["count"] for v in result.values()}


def cross_tab(df: pd.DataFrame, row_col: str, col_col: str) -> dict:
    rows_labels = sorted(df[row_col].dropna().unique(), key=lambda x: normalize(str(x)))
    cols_labels = sorted(df[col_col].dropna().unique(), key=lambda x: normalize(str(x)))
    matrix = {r: {c: 0 for c in cols_labels} for r in rows_labels}
    row_totals = {r: 0 for r in rows_labels}
    col_totals = {c: 0 for c in cols_labels}
    grand = 0
    for _, row in df.iterrows():
        r, c = row.get(row_col), row.get(col_col)
        if pd.isna(r) or pd.isna(c):
            continue
        r, c = str(r).strip(), str(c).strip()
        if r in matrix and c in matrix[r]:
            matrix[r][c] += 1
            row_totals[r] += 1
            col_totals[c] += 1
            grand += 1
    return {
        "rows": [str(r) for r in rows_labels],
        "cols": [str(c) for c in cols_labels],
        "matrix": {str(r): {str(c): matrix[r][c] for c in cols_labels} for r in rows_labels},
        "rowTotals": {str(r): row_totals[r] for r in rows_labels},
        "colTotals": {str(c): col_totals[c] for c in cols_labels},
        "grandTotal": grand,
    }


def analyze(excel_path: Path, schema: dict) -> dict:
    df = pd.read_excel(excel_path, header=0)
    meta = schema["metadata"]

    # Deduplication by email
    email_col_idx = None
    for kw in meta.get("emailCol", []):
        for col in df.columns:
            if normalize(kw) in normalize(str(col)):
                email_col_idx = col
                break
        if email_col_idx:
            break
    if email_col_idx:
        df = df.drop_duplicates(subset=[email_col_idx], keep="last")

    col_map = map_columns(df, schema)
    scale_type = schema["scale"]["type"]
    n = len(df)

    global_seq = 1
    axes_out = []
    total_q_counts = []
    total_agree_sum = 0
    total_mean_sum = 0
    total_q = 0

    for ax in schema["axes"]:
        qs_out = []
        ax_agree_sum = 0
        ax_mean_sum = 0
        ax_n = 0
        ax_count = 0

        for q in ax["questions"]:
            col = col_map.get(q["id"])
            series = df[col] if col else pd.Series(dtype=object)
            stats = analyze_question(series, schema)
            interp_val = stats["mean"] if scale_type == "likert-5" else stats["agreePct"]
            interp = interpret(interp_val or 0, schema)
            qs_out.append({
                "seq": q["seq"], "id": q["id"], "text": q["text"],
                "counts": stats["counts"], "pcts": stats["pcts"],
                "mean": stats["mean"], "agreePct": stats["agreePct"],
                "total": stats["total"],
                **interp,
            })
            ax_agree_sum += stats["agreePct"]
            if stats["mean"] is not None:
                ax_mean_sum += stats["mean"]
            ax_n += stats["total"]
            ax_count += 1

        axis_agree_pct = round(ax_agree_sum / ax_count, 1) if ax_count else 0
        axis_mean = round(ax_mean_sum / ax_count, 2) if (ax_count and scale_type == "likert-5") else None
        ax_interp_val = axis_mean if scale_type == "likert-5" else axis_agree_pct
        ax_interp = interpret(ax_interp_val or 0, schema)

        axes_out.append({
            "id": ax["id"], "name": ax["name"],
            **({} if "recommendation" not in ax else {"recommendation": ax["recommendation"]}),
            "axisMean": axis_mean,
            "axisAgreePct": axis_agree_pct,
            **ax_interp,
            "questions": qs_out,
        })
        total_agree_sum += ax_agree_sum
        if axis_mean is not None:
            total_mean_sum += axis_mean
        total_q += ax_count

    overall_agree_pct = round(total_agree_sum / total_q, 1) if total_q else 0
    overall_mean = round(total_mean_sum / len(axes_out), 2) if (axes_out and scale_type == "likert-5") else None
    overall_val = overall_mean if scale_type == "likert-5" else overall_agree_pct
    overall_interp = interpret(overall_val or 0, schema)

    # Demographic groupings
    by_degree, by_department, cross = {}, {}, None
    degree_col = _find_col_name(df, meta.get("degreeCol", []))
    dept_col   = _find_col_name(df, meta.get("departmentCol", []))
    if degree_col:
        by_degree = count_by_col(df, degree_col)
    if dept_col:
        by_department = count_by_col(df, dept_col)
    if degree_col and dept_col:
        cross = cross_tab(df, degree_col, dept_col)

    return {
        "schemaId":         schema["id"],
        "schemaLabel":      schema["label"],
        "scaleType":        scale_type,
        "n":                n,
        "axes":             axes_out,
        "overallMean":      overall_mean,
        "overallAgreePct":  overall_agree_pct,
        "overallDirection": overall_interp["direction"],
        "overallTier":      overall_interp["tier"],
        "byDegree":         by_degree,
        "byDepartment":     by_department,
        "crossDegreeByDept": cross,
        "totalQuestions":   total_q,
    }


def _find_col_name(df: pd.DataFrame, keywords: list) -> str | None:
    for kw in keywords:
        nkw = normalize(kw)
        for col in df.columns:
            if nkw in normalize(str(col)):
                return col
    return None


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Redaa2 survey analyzer")
    parser.add_argument("schema",  help="Path to .yaml schema file or schemas/ directory")
    parser.add_argument("excel",   help="Path to .xlsx survey file")
    parser.add_argument("--out",   help="Output JSON path (default: stdout)")
    args = parser.parse_args()

    schema_path = Path(args.schema)
    excel_path  = Path(args.excel)

    if schema_path.is_dir():
        schema = detect_schema(schema_path, excel_path.name)
    else:
        schema = load_schema(schema_path)

    result = analyze(excel_path, schema)
    output = json.dumps(result, ensure_ascii=False, indent=2)

    if args.out:
        Path(args.out).write_text(output, encoding="utf-8")
        print(f"Saved → {args.out}")
    else:
        print(output)


if __name__ == "__main__":
    main()
