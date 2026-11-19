"""
score_loader.py
讀取分數記錄表（scores.xlsx）：支援 MC 選項自動評分、二選一大題、扣分欄
load_data_from_bytes() 回傳格式與舊版相同：
    (score_data, max_scores, absent_set, paper_map, class_info)
load_data_with_extras() 額外回傳 extras 字典（供 exam_item_analysis 使用）。
"""
import io, re
from collections import defaultdict
import numpy as np
import pandas as pd

PAPER_LABELS = {"P1", "P2", "P3", "P4"}
INFO_COLS    = ["班別", "班號", "英文姓名", "中文姓名"]
ABSENT_COLS  = ["缺席", "Absent", "absent", "缺考"]
DEDUCT_NAMES = ("扣分",)
MC_LETTERS   = set("ABCD")


def _num(v):
    try:
        f = float(v)
        return f if f == f else np.nan
    except (TypeError, ValueError):
        return np.nan


def _mc_marks_from_notes(df_raw, default=2.0):
    text = " ".join(str(x) for x in df_raw.iloc[:2].values.ravel() if pd.notna(x))
    m = re.search(r"多項選擇題[^，。,；;]*?每題\s*(\d+(?:\.\d+)?)\s*分", text)
    return float(m.group(1)) if m else default


def _parent(q):
    m = re.match(r"(Q?\d+)", q.split("_", 1)[-1])
    return m.group(1) if m else q


def load_data_with_extras(file_bytes: bytes, mc_marks: float = None):
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, sheet_name=0)
    col_names = df_raw.iloc[2].tolist()
    row3 = df_raw.iloc[3].tolist()
    row3_str = [str(v).strip() for v in row3 if pd.notna(v) and str(v).strip()]
    has_paper_row = any(v in PAPER_LABELS for v in row3_str)
    if has_paper_row:
        paper_row, max_row, student_raw = row3, df_raw.iloc[4].tolist(), df_raw.iloc[5:].copy()
    else:
        paper_row, max_row, student_raw = None, row3, df_raw.iloc[4:].copy()
    student_raw.columns = col_names
    student_raw = student_raw[
        student_raw["中文姓名"].notna()
        & ~student_raw["中文姓名"].astype(str).str.contains("說明|輸入|請在此行", na=False)
    ]
    names = student_raw["中文姓名"].astype(str).str.strip().values

    absent_set = set()
    absent_col_idx = next((i for i, c in enumerate(col_names) if str(c) in ABSENT_COLS), None)
    if absent_col_idx is not None:
        for nm, v in zip(names, student_raw.iloc[:, absent_col_idx]):
            if pd.notna(v) and str(v).strip() not in ("", "0", "nan"):
                absent_set.add(nm)

    if mc_marks is None:
        mc_marks = _mc_marks_from_notes(df_raw)

    raw_q = []
    deduct_idx = {}
    for i, c in enumerate(col_names):
        if not (isinstance(c, str) and c.strip()):
            continue
        c = c.strip()
        if c in INFO_COLS or c in ABSENT_COLS:
            continue
        pv = str(paper_row[i]).strip() if (paper_row is not None and pd.notna(paper_row[i])) else "P1"
        paper = pv if pv in PAPER_LABELS else "P1"
        if c in DEDUCT_NAMES:
            deduct_idx.setdefault(paper, []).append(i)
            continue
        mv = max_row[i] if i < len(max_row) else np.nan
        if isinstance(mv, str) and mv.strip().upper() in MC_LETTERS:
            raw_q.append((c, i, paper, "mc", mv.strip().upper()))
        else:
            f = _num(mv)
            if f == f and f > 0:
                raw_q.append((c, i, paper, "num", f))

    multi_paper = has_paper_row and len({p for _, _, p, _, _ in raw_q}) > 1
    seen, final = {}, []
    for orig, i, paper, kind, val in raw_q:
        base = f"{paper}_{orig}" if multi_paper else orig
        if base in seen:
            seen[base] += 1
            base = f"{base}_{seen[base]}"
        else:
            seen[base] = 1
        final.append((base, i, paper, kind, val))

    cols, max_scores, paper_map, mc_info = {}, {}, {}, {}
    for name, i, paper, kind, val in final:
        s = student_raw.iloc[:, i]
        if kind == "mc":
            ans = s.astype(str).str.strip().str.upper()
            cols[name] = pd.Series(np.where(ans == val, mc_marks, 0.0), index=s.index)
            max_scores[name] = mc_marks
            mc_info[name] = {"key": val, "marks": mc_marks,
                             "answers": ans.where(ans.isin(MC_LETTERS), "-").values}
        else:
            cols[name] = pd.to_numeric(s, errors="coerce")
            max_scores[name] = val
        paper_map[name] = paper

    raw_scores = pd.DataFrame(cols)
    raw_scores.index = names
    raw_scores.index.name = "姓名"
    max_scores = pd.Series(max_scores)
    present = ~raw_scores.index.isin(absent_set)

    elective, attempted = {}, pd.DataFrame(True, index=raw_scores.index, columns=raw_scores.columns)
    for paper in sorted(set(paper_map.values())):
        groups = defaultdict(list)
        for q in raw_scores.columns:
            if paper_map[q] == paper and q not in mc_info:
                groups[_parent(q)].append(q)
        cand = {}
        for g, qs in groups.items():
            blank_all = raw_scores.loc[present, qs].isna().all(axis=1)
            if 0 < blank_all.mean() < 1:
                cand[g] = (qs, float(max_scores[qs].sum()))
        if len(cand) < 2:
            continue
        mx_set = {round(v[1], 6) for v in cand.values()}
        if len(mx_set) != 1:
            continue
        gl = list(cand)
        nonblank = pd.DataFrame({g: ~raw_scores.loc[present, cand[g][0]].isna().all(axis=1) for g in gl})
        if (nonblank.sum(axis=1) == 1).mean() < 0.8:
            continue
        tot = pd.DataFrame({g: raw_scores[cand[g][0]].fillna(0).sum(axis=1) for g in gl})
        has = pd.DataFrame({g: ~raw_scores[cand[g][0]].isna().all(axis=1) for g in gl})
        tot = tot.where(has, -1)
        chosen = tot.idxmax(axis=1).where(has.any(axis=1))
        for g in gl:
            for q in cand[g][0]:
                attempted[q] = (chosen == g).values
        elective[paper] = {"groups": {g: cand[g][0] for g in gl}, "chosen": chosen,
                           "n_selected": chosen[present].value_counts().to_dict(),
                           "group_max": mx_set.pop()}

    score_data = raw_scores.fillna(0.0)
    for q in score_data.columns:
        score_data[q] = score_data[q].where(attempted[q], 0.0)

    deductions = {}
    for paper, idxs in deduct_idx.items():
        d = sum(pd.to_numeric(student_raw.iloc[:, i], errors="coerce").fillna(0.0).values for i in idxs)
        deductions[paper] = pd.Series(d, index=names)

    paper_max_eff = {}
    for p in sorted(set(paper_map.values())):
        tot = float(max_scores[[q for q in max_scores.index if paper_map[q] == p]].sum())
        if p in elective:
            tot -= elective[p]["group_max"] * (len(elective[p]["groups"]) - 1)
        paper_max_eff[p] = tot

    ci = pd.DataFrame({"班別": student_raw["班別"].values, "班號": student_raw["班號"].values,
                       "中文姓名": names})
    extras = {"attempted": attempted, "elective": elective, "deductions": deductions,
              "mc_info": mc_info, "paper_max_eff": paper_max_eff, "mc_marks": mc_marks}
    return score_data, max_scores, absent_set, paper_map, ci, extras


def load_data_from_bytes(file_bytes: bytes):
    """與舊版相容的介面。"""
    sd, ms, ab, pm, ci, _ = load_data_with_extras(file_bytes)
    return sd, ms, ab, pm, ci


def total_scores(score_data, extras=None):
    """各生總分（含扣分）。"""
    tot = score_data.sum(axis=1)
    for s in (extras or {}).get("deductions", {}).values():
        tot = tot + s.reindex(tot.index).fillna(0.0)
    return tot
