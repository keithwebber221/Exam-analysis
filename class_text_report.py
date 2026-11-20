"""
class_text_report.py
全班試卷分析文字報告模組
輸入：df, max_scores, paper_map, class_info, exam_info, absent_set,
       item_df, group_df, student_df, stats_df
輸出：Word (.docx) BytesIO
整合：由 app.py 呼叫 generate_text_report()
"""

import io, re
import numpy as np
import pandas as pd
from collections import defaultdict

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ══════════════════════════════════════════════════════════════
# 樣式工具
# ══════════════════════════════════════════════════════════════

LATIN_FONT     = "Times New Roman"   # 英文、數字
EAST_ASIA_FONT = "新細明體"           # 中文
EXCELLENT_PCT  = 80                  # 優秀：得分率 80% 或以上


def _set_font(run, size, bold=False, color=None):
    run.font.name = LATIN_FONT
    rPr = run._element.get_or_add_rPr()
    rf  = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for _a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(_a), LATIN_FONT)
    rf.set(qn("w:eastAsia"), EAST_ASIA_FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def _no_page_break(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ["w:pageBreakBefore", "w:keepNext", "w:keepLines"]:
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
        el = OxmlElement(tag)
        el.set(qn("w:val"), "0")
        pPr.append(el)


def _add_heading(doc, text, level=1, color=(0, 0, 0)):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7 if level == 1 else 4)
    p.paragraph_format.space_after  = Pt(2)
    _no_page_break(p)
    r = p.add_run(text)
    _set_font(r, 13 if level == 1 else 11, bold=True, color=color)
    return p


def _add_body(doc, text, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.line_spacing = Pt(16)
    _no_page_break(p)
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    r = p.add_run(text)
    _set_font(r, 10.5)
    return p


def _add_bullet(doc, text, color=None):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Cm(0.6)
    _no_page_break(p)
    r = p.add_run(text)
    _set_font(r, 10.5, color=color)
    return p


def _add_divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(1)
    _no_page_break(p)
    r = p.add_run("─" * 60)
    _set_font(r, 8, color=(190, 190, 190))


def _add_spacer(doc, after_pt=2):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after_pt)
    _no_page_break(p)


def _set_cell_bg(cell, hex_color):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _add_note_box(doc, label, body_lines,
                  bg_hex="EEF4FF", label_rgb=(31, 71, 136), body_rgb=(40, 40, 80)):
    tbl  = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.rows[0].cells[0]
    _set_cell_bg(cell, bg_hex)

    trPr      = tbl.rows[0]._tr.get_or_add_trPr()
    cantSplit = OxmlElement("w:cantSplit")
    cantSplit.set(qn("w:val"), "0")
    trPr.append(cantSplit)

    cell.paragraphs[0].clear()
    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(3)
    p0.paragraph_format.space_after  = Pt(1)
    _no_page_break(p0)
    rl = p0.add_run(f"{label}　")
    _set_font(rl, 10, bold=True, color=label_rgb)
    rb = p0.add_run(body_lines[0])
    _set_font(rb, 10, color=body_rgb)

    for line in body_lines[1:]:
        pe = cell.add_paragraph()
        pe.paragraph_format.space_before = Pt(0)
        pe.paragraph_format.space_after  = Pt(1)
        pe.paragraph_format.left_indent  = Cm(0.2)
        _no_page_break(pe)
        re_ = pe.add_run(line)
        _set_font(re_, 10, color=body_rgb)
    return tbl


# ══════════════════════════════════════════════════════════════
# 數據計算輔助
# ══════════════════════════════════════════════════════════════

def _pl(q):
    """顯示用題號：P2_Q8a → 卷二 Q8a"""
    s = str(q)
    for p, l in [("P1_", "卷一 "), ("P2_", "卷二 "), ("P3_", "卷三 "), ("P4_", "卷四 ")]:
        s = s.replace(p, l)
    return s


def _short(q):
    q = str(q)
    return q.split("_", 1)[1] if "_" in q else q


def _group_of(q, mc_q):
    q = str(q)
    if q in mc_q:
        return (q.split("_", 1)[0] + "_" if "_" in q else "") + "MC選擇題"
    return _get_parent_q(q)[1]


def _genuine_poor(item_df):
    """真正需要檢視的低鑑別度題目：鑑別度 < 0.2，且得分率介乎 20%–85%。
    得分率過高（題目太易）或過低（題目太難）時，鑑別度偏低屬自然結果。"""
    if "鑑別度" not in item_df.columns:
        return item_df.iloc[0:0]
    rate = item_df["得分率%"]
    return item_df[(item_df["鑑別度"] < 0.2) & (rate >= 20) & (rate <= 85)]


def _extreme_low_disc(item_df):
    if "鑑別度" not in item_df.columns:
        return item_df.iloc[0:0]
    rate = item_df["得分率%"]
    return item_df[(item_df["鑑別度"] < 0.2) & ((rate > 85) | (rate < 20))]



def _get_parent_q(q_name):
    parts = q_name.split("_", 1)
    if len(parts) < 2:
        # No paper prefix — single-paper mode
        item = q_name
        paper = "P1"
        if "Part" in item or "part" in item:
            return paper, "PartA（選擇題）"
        # Match "Q1a" → "Q1",  "1a" → "1",  "Q10b" → "Q10",  "10b" → "10"
        m = re.match(r"(Q?[\d/]+)", item)
        if m:
            return paper, m.group(1)
        return paper, item
    paper, item = parts
    if "Part" in item or "part" in item:
        return paper, f"{paper}_PartA（選擇題）"
    m = re.match(r"(Q?[\d/]+)", item)
    if m:
        return paper, f"{paper}_{m.group(1)}"
    return paper, f"{paper}_{item}"


def _build_group_stats(df, max_scores, paper_map, absent_set, total_scores,
                       item_df=None, overall_rate=None, mc_q=None, pass_line=40.0):
    """Build parent-question group statistics.
    大題平均 = 各小題平均分之和（小題平均已排除缺席學生，二選一題只計選答學生），
    故與試題分析一致；MC 選擇題自成一個大題，避免與乙部 Q1a、Q1b… 混在一起。"""
    present_mask   = ~df.index.isin(absent_set)
    present_scores = df[present_mask]
    if overall_rate is None:
        overall_rate = total_scores.mean() / max_scores.sum() * 100
    mc_q = set(mc_q or [])

    item_mean = {}
    if item_df is not None:
        _qc = "題目" if "題目" in item_df.columns else item_df.columns[0]
        item_mean = dict(zip(item_df[_qc].astype(str), item_df["平均分"]))

    parent_groups = defaultdict(list)
    for q in df.columns:
        if q in mc_q:
            _pre = q.split("_", 1)[0] + "_" if "_" in q else ""
            parent_groups[_pre + "MC選擇題"].append(q)
            continue
        _, parent = _get_parent_q(q)
        parent_groups[parent].append(q)

    rows = []
    for parent, qs in parent_groups.items():
        grp_max    = sum(max_scores[q] for q in qs)
        if item_mean and all(str(q) in item_mean for q in qs):
            grp_avg = float(sum(item_mean[str(q)] for q in qs))
        else:
            grp_avg = present_scores[qs].sum(axis=1).mean()
        grp_rate   = grp_avg / grp_max * 100 if grp_max else 0

        diff = ("容易" if grp_rate >= EXCELLENT_PCT else "困難" if grp_rate < pass_line else "適中")
        # Clean up display name
        display = parent
        for prefix, label in [("P1_", ""), ("P2_", "卷二 "), ("P3_", "卷三 "), ("P4_", "卷四 ")]:
            display = display.replace(prefix, label)
        display = display.strip()
        rows.append({
            "parent": parent, "display": display,
            "qs": qs, "max": grp_max, "avg": round(grp_avg, 2),
            "rate": round(grp_rate, 1), "diff": diff,
            "paper": qs[0].split("_")[0] if "_" in qs[0] else "P1",
        })

    group_df2 = pd.DataFrame(rows).sort_values("rate", ascending=False)
    return group_df2, overall_rate


# ══════════════════════════════════════════════════════════════
# 各部分建構函式
# ══════════════════════════════════════════════════════════════

def _build_cover(doc, exam_info):
    subject_label   = exam_info.get("subject_label", "")
    form_label      = exam_info.get("form_label", "")
    year_label      = exam_info.get("year_label", "")
    exam_type_label = exam_info.get("exam_type_label", "")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(2)
    _no_page_break(p)
    r = p.add_run("全班試卷分析文字報告")
    _set_font(r, 20, bold=True, color=(0, 0, 0))

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    p2.paragraph_format.space_after  = Pt(2)
    _no_page_break(p2)
    r2 = p2.add_run(f"{year_label}　{exam_type_label}　{form_label}　{subject_label}")
    _set_font(r2, 11, color=(0, 0, 0))

    _add_spacer(doc, 2)
    _add_divider(doc)

    toc_p = doc.add_paragraph()
    toc_p.paragraph_format.space_before = Pt(2)
    toc_p.paragraph_format.space_after  = Pt(2)
    _no_page_break(toc_p)
    r_toc = toc_p.add_run(
        "本報告共分五部分：一、試卷總覽　二、各題表現分析　三、大題分析　四、學生分層分析　五、後續跟進建議"
    )
    _set_font(r_toc, 9.5, color=(0, 0, 0))

    _add_spacer(doc, 2)
    _add_note_box(
        doc, label="📌 名詞說明",
        bg_hex="EEF4FF", label_rgb=(31, 71, 136), body_rgb=(40, 40, 80),
        body_lines=[
            "閱讀本報告前，請留意以下兩個常用指標的含義：",
            "",
            "【pp（百分點 / percentage points）】兩個百分比之間的直接差距。"
            "例如某大題得分率 34%、全卷整體得分率 61%，相差 27pp。▲ 高於全卷整體得分率；▼ 低於全卷整體得分率。",
            "",
            "【評級標準】優秀：得分率 80% 或以上；合格：達合格線；不合格：低於合格線；其餘為中等。題目難度同樣以此劃分（80% 或以上為容易、低於合格線為困難）。",
            "",
            "【鑑別度（Discrimination Index）】衡量題目能否有效區分高、低能力學生，數值介乎 −1 至 1，數值愈高愈理想。",
            "    • 0.40 或以上 → 優良　• 0.30–0.39 → 良好　"
            "• 0.20–0.29 → 尚可（可考慮修訂）　• 0.20 以下 → 偏低",
            "    若題目得分率極高（高於 85%）或極低（低於 20%），鑑別度偏低屬自然結果，不視為題目設計問題。",
        ]
    )
    _add_spacer(doc, 2)


def _build_overview(doc, exam_info, n_total, n_absent, n_present,
                    n_pass, pass_pct, mean_score, mean_pct,
                    max_score_val, min_score_val, std_score, median_score,
                    total_max, pass_threshold, paper_max_str):
    subject_label   = exam_info.get("subject_label", "")
    form_label      = exam_info.get("form_label", "")
    year_label      = exam_info.get("year_label", "")
    exam_type_label = exam_info.get("exam_type_label", "")
    pass_rate_pct   = int(exam_info.get("pass_rate", 0.4) * 100)

    _add_heading(doc, "一、試卷總覽")

    tbl = doc.add_table(rows=3, cols=4)
    tbl.style = "Table Grid"
    rows_data = [
        ["考試名稱", f"{year_label} {exam_type_label}", "科目", subject_label],
        ["年級",     form_label, "總滿分", f"{total_max} 分{paper_max_str}"],
        ["合格線",   f"{int(pass_threshold)} 分（{pass_rate_pct}%）",
         "出席人數", f"{n_present} 人（全班 {n_total} 人，缺席 {n_absent} 人）"],
    ]
    for ri, row_data in enumerate(rows_data):
        for ci, val in enumerate(row_data):
            cell = tbl.rows[ri].cells[ci]
            cell.text = val
            run  = cell.paragraphs[0].runs[0]
            cell.paragraphs[0].paragraph_format.space_before = Pt(1)
            cell.paragraphs[0].paragraph_format.space_after  = Pt(1)
            _no_page_break(cell.paragraphs[0])
            _set_font(run, 9.5, bold=(ci % 2 == 0),
                      color=(0, 0, 0))

    _add_spacer(doc, 1)

    overall_comment = (
        f"整體表現理想，合格率達 {pass_pct:.1f}%，反映大部分同學已掌握本科核心概念。"
        if pass_pct >= 80 else
        f"整體表現尚可，合格率為 {pass_pct:.1f}%，仍有部分同學需要針對性跟進。"
        if pass_pct >= 60 else
        f"整體表現欠佳，合格率僅 {pass_pct:.1f}%，建議全面檢視教學重點。"
    )
    _add_body(doc,
        f"全班共 {n_present} 名同學出席考試，合格人數為 {n_pass} 人，"
        f"合格率為 {pass_pct:.1f}%。{overall_comment}"
    )
    _add_body(doc,
        f"全班平均分為 {mean_score:.1f} 分（得分率 {mean_pct:.1f}%），"
        f"最高分 {round(max_score_val, 1):g} 分，最低分 {round(min_score_val, 1):g} 分，"
        f"標準差 {std_score:.1f}，中位數 {median_score:.1f} 分。"
    )
    _add_divider(doc)


def _build_mc_section(doc, df, extras, item_df, absent_set, q_col, pass_line):
    mc = (extras or {}).get("mc_info") or {}
    if not mc:
        return
    present = ~df.index.isin(absent_set)
    n = int(present.sum())
    rows = []
    for q, d in mc.items():
        ans_all = np.asarray(d["answers"])
        if len(ans_all) != len(df):
            continue
        ans = ans_all[present]
        cnt = {o: int((ans == o).sum()) for o in "ABCD"}
        blank = n - sum(cnt.values())
        rows.append({"q": q, "key": d["key"], "cnt": cnt, "blank": blank,
                     "rate": cnt.get(d["key"], 0) / n * 100 if n else 0.0})
    if not rows:
        return
    _add_heading(doc, "（丁）MC 選擇題選項分析", level=2, color=(0, 0, 0))
    avg_rate = float(np.mean([r["rate"] for r in rows]))
    _add_body(doc, f"本次共 {len(rows)} 題 MC，平均答對率 {avg_rate:.1f}%。下表列出各選項的選答百分比（粗體為正確答案）。")

    hdr = ["題號", "答案", "A", "B", "C", "D", "未答／無效", "答對率"]
    tbl = doc.add_table(rows=1, cols=len(hdr))
    tbl.style = "Table Grid"
    for c, h in zip(tbl.rows[0].cells, hdr):
        c.text = h
        _set_cell_bg(c, "D9D9D9")
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_font(c.paragraphs[0].runs[0], 9, bold=True, color=(0, 0, 0))
    for r in rows:
        cells = tbl.add_row().cells
        vals = [_short(r["q"]), r["key"]] + [f"{r['cnt'][o] / n * 100:.0f}%" for o in "ABCD"] + \
               [f"{r['blank'] / n * 100:.0f}%", f"{r['rate']:.0f}%"]
        for ci, (c, v) in enumerate(zip(cells, vals)):
            c.text = v
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            c.paragraphs[0].paragraph_format.space_before = Pt(0)
            c.paragraphs[0].paragraph_format.space_after  = Pt(0)
            is_key = ci >= 2 and ci <= 5 and "ABCD"[ci - 2] == r["key"]
            _set_font(c.paragraphs[0].runs[0], 9, bold=is_key or ci in (0, 7), color=(0, 0, 0))
    _add_spacer(doc, 1)

    low = sorted([r for r in rows if r["rate"] < EXCELLENT_PCT], key=lambda r: r["rate"])[:3]
    for r in low:
        wrong = {o: c for o, c in r["cnt"].items() if o != r["key"]}
        top_o = max(wrong, key=wrong.get)
        share = wrong[top_o] / n * 100 if n else 0
        extra = "，比正確答案更多人選，屬強力誘答項" if wrong[top_o] > r["cnt"].get(r["key"], 0) else ""
        _add_bullet(doc,
            f"{_pl(r['q'])}：答對率 {r['rate']:.0f}%（答案 {r['key']}），"
            f"最多人誤選 {top_o}（{share:.0f}%）{extra}。建議檢視選項 {top_o} 對應的概念誤解。")
    if not low:
        _add_body(doc, "各 MC 題答對率均達優秀水平，未見明顯的誘答項問題。", indent=True)


def _build_elective_section(doc, df, extras, item_df, absent_set, q_col):
    el  = (extras or {}).get("elective") or {}
    att = (extras or {}).get("attempted")
    if not el or att is None:
        return
    present = ~df.index.isin(absent_set)
    n_pres  = int(present.sum())
    if n_pres == 0:
        return
    item_mean = dict(zip(item_df[q_col].astype(str), item_df["平均分"]))
    labels = {"P1": "卷一", "P2": "卷二", "P3": "卷三", "P4": "卷四"}
    _add_heading(doc, "（戊）二選一題選答情況", level=2, color=(0, 0, 0))
    for paper, info in el.items():
        parts = []
        for g, qs in info["groups"].items():
            n_g = int(att.loc[present, qs[0]].sum())
            mx  = float(info["group_max"])
            avg = float(sum(item_mean.get(str(q), 0.0) for q in qs))
            parts.append({"g": g, "n": n_g, "avg": avg, "mx": mx,
                          "rate": avg / mx * 100 if mx else 0.0})
        label = labels.get(paper, paper)
        txt = "；".join(f"{p['g']} 選答 {p['n']} 人（{p['n'] / n_pres * 100:.0f}%），"
                       f"平均 {p['avg']:.1f}/{p['mx']:g} 分（得分率 {p['rate']:.0f}%）" for p in parts)
        _add_bullet(doc, f"{label}二選一題：{txt}。")
        rates = [p["rate"] for p in parts]
        if len(parts) > 1 and max(rates) - min(rates) >= 10:
            lo = min(parts, key=lambda p: p["rate"])
            _add_body(doc, f"兩題得分率相差 {max(rates) - min(rates):.0f}pp，{lo['g']} 較低，"
                           f"宜留意兩題難度是否平衡。", indent=True)
        few = [p for p in parts if p["n"] / n_pres < 0.15]
        for p in few:
            _add_body(doc, f"{p['g']} 選答比例偏低（{p['n'] / n_pres * 100:.0f}%），"
                           f"同學傾向避開該題，相關課題的教學覆蓋可再檢視。", indent=True)


def _build_item_analysis(doc, item_df, overall_rate, pass_line=40.0, mc_q=None,
                         extras=None, df=None, absent_set=None):
    mc_q = set(mc_q or [])
    _add_heading(doc, "二、各題表現分析")
    _add_body(doc, f"以下按各題得分率及鑑別度分類呈現（優秀：{EXCELLENT_PCT}% 或以上；合格線：{pass_line:g}%）：")

    item_df = item_df.copy()
    if "得分率%" not in item_df.columns:
        item_df["得分率%"] = (item_df["平均分"] / item_df["滿分"] * 100).round(1)
    rate_col = "得分率%"
    q_col = "題目" if "題目" in item_df.columns else item_df.columns[0]

    weak_qs   = item_df[item_df[rate_col] < pass_line].sort_values(rate_col)
    strong_qs = item_df[item_df[rate_col] >= EXCELLENT_PCT].sort_values(rate_col, ascending=False)

    def _disc_label(d):
        if d >= 0.4: return "優良"
        if d >= 0.3: return "良好"
        if d >= 0.2: return "尚可"
        return "偏低"

    # 甲：弱項（按大題合併）
    _add_heading(doc, f"（甲）表現較弱的題目（得分率低於合格線 {pass_line:g}%）", level=2, color=(180, 40, 40))
    if len(weak_qs) > 0:
        buckets = defaultdict(list)
        for _, row in weak_qs.iterrows():
            buckets[_group_of(row[q_col], mc_q)].append(row)
        total_in = defaultdict(int)
        for q in item_df[q_col]:
            total_in[_group_of(q, mc_q)] += 1
        order = sorted(buckets.items(), key=lambda kv: float(np.mean([r[rate_col] for r in kv[1]])))
        for parent, rows in order:
            rows = sorted(rows, key=lambda r: r[rate_col])
            shown = "、".join(f"{_short(r[q_col])} {r[rate_col]:.0f}%" for r in rows[:10])
            if len(rows) > 10:
                shown += f" 等共 {len(rows)} 題"
            w = rows[0]
            d = w.get("鑑別度", np.nan)
            d_txt = f"，鑑別度 {d:.2f}（{_disc_label(d)}）" if pd.notna(d) else ""
            n_note = ""
            if "作答人數" in item_df.columns and w["作答人數"] < item_df["作答人數"].max():
                n_note = f"（選答 {int(w['作答人數'])} 人）"
            _add_bullet(doc,
                f"{_pl(parent)}：{len(rows)}/{total_in[parent]} 條小題低於合格線（{shown}）；"
                f"最弱為 {_short(w[q_col])}{n_note}，平均 {w['平均分']:.1f}/{int(w['滿分'])} 分{d_txt}。"
                f"建議針對相關概念安排補充練習。")
    else:
        _add_body(doc, f"本次考試所有題目得分率均達合格線（{pass_line:g}%）以上。", indent=True)

    # 乙：強項
    _add_heading(doc, f"（乙）表現優秀的題目（得分率達 {EXCELLENT_PCT}% 或以上）", level=2, color=(0, 120, 60))
    if len(strong_qs) > 0:
        for _, row in strong_qs.head(5).iterrows():
            _add_bullet(doc,
                f"{_pl(row[q_col])}（滿分 {int(row['滿分'])} 分）：平均 {row['平均分']:.1f} 分，"
                f"得分率 {row[rate_col]:.1f}%，同學整體掌握良好。")
        if len(strong_qs) > 5:
            _add_body(doc, f"（另有 {len(strong_qs) - 5} 題得分率達 {EXCELLENT_PCT}% 或以上，從略。）", indent=True)
    else:
        _add_body(doc, f"本次考試沒有題目得分率達 {EXCELLENT_PCT}% 或以上。", indent=True)

    # 丙：鑑別度
    _add_heading(doc, "（丙）鑑別度偏低的題目（鑑別度低於 0.2）", level=2, color=(160, 90, 0))
    genuine = _genuine_poor(item_df)
    extreme = _extreme_low_disc(item_df)
    if len(genuine) > 0:
        for _, row in genuine.iterrows():
            n_note = ""
            if "作答人數" in item_df.columns and row["作答人數"] < 10:
                n_note = f"（作答人數僅 {int(row['作答人數'])} 人，數值僅供參考）"
            _add_bullet(doc,
                f"{_pl(row[q_col])}：鑑別度 {row['鑑別度']:.2f}，得分率 {row[rate_col]:.0f}%{n_note}，"
                f"未能有效區分不同能力的學生，建議檢視題目設計。")
    if len(extreme) > 0:
        names = "、".join(_pl(q) for q in extreme[q_col].tolist()[:12])
        if len(extreme) > 12:
            names += f" 等共 {len(extreme)} 題"
        _add_body(doc,
            f"另有 {len(extreme)} 題鑑別度偏低，但得分率極高（高於 85%）或極低（低於 20%）：{names}。"
            f"這是題目過易或過難的自然結果，不代表題目設計有問題，毋須修訂。", indent=True)
    if len(genuine) == 0 and len(extreme) == 0:
        _add_body(doc, "本次考試各題鑑別度整體理想。", indent=True)

    if df is not None:
        _build_mc_section(doc, df, extras, item_df, absent_set or set(), q_col, pass_line)
        _build_elective_section(doc, df, extras, item_df, absent_set or set(), q_col)

    _add_divider(doc)



def _natural_key(s):
    return [int(x) if x.isdigit() else x for x in re.split(r"(\d+)", str(s))]


def _build_group_analysis(doc, df, max_scores, paper_map, absent_set,
                           total_scores, item_df, overall_rate=None, mc_q=None, pass_line=40.0):
    group_df2, overall_rate = _build_group_stats(
        df, max_scores, paper_map, absent_set, total_scores,
        item_df=item_df, overall_rate=overall_rate, mc_q=mc_q, pass_line=pass_line)

    _add_heading(doc, "三、大題分析")
    _add_body(doc, f"以下按大題匯總各小題表現，並與全卷整體得分率（{overall_rate:.1f}%）作比較：")

    q_col     = "題目" if "題目" in item_df.columns else item_df.columns[0]
    rate_col  = "得分率%" if "得分率%" in item_df.columns else None
    papers_in = sorted(group_df2["paper"].unique())

    multi_paper_mode = len(papers_in) > 1

    for paper_key in papers_in:
        label_map = {"P1": "卷一（Paper 1）", "P2": "卷二（Paper 2）",
                     "P3": "卷三（Paper 3）", "P4": "卷四（Paper 4）"}
        if multi_paper_mode:
            paper_label = label_map.get(paper_key, paper_key)
            _add_heading(doc, f"【{paper_label}】", level=2, color=(0, 0, 0))

        paper_rows = sorted(group_df2[group_df2["paper"] == paper_key].to_dict("records"), key=lambda r: _natural_key(r["parent"]))

        for row in paper_rows:
            rate    = row["rate"];  delta   = rate - overall_rate
            mx      = row["max"];   avg     = row["avg"]
            diff    = row["diff"];  qs      = row["qs"];  display = row["display"]

            rate_color   = (0, 120, 60) if rate >= EXCELLENT_PCT else (160, 100, 0) if rate >= pass_line else (180, 40, 40)
            delta_symbol = "▲" if delta >= 0 else "▼"
            delta_color  = (0, 120, 60) if delta >= 0 else (180, 40, 40)

            # sub-question breakdown
            sub_parts = []
            for q in qs:
                q_short  = q.split("_", 1)[1] if "_" in q else q
                q_row    = item_df[item_df[q_col] == q]
                if len(q_row):
                    if rate_col:
                        q_rate = q_row[rate_col].values[0]
                    else:
                        q_rate = q_row["平均分"].values[0] / q_row["滿分"].values[0] * 100
                    q_delta = q_rate - overall_rate
                    sym = "▲" if q_delta >= 0 else "▼"
                    sub_parts.append(f"{q_short} {q_rate:.0f}%（{sym}{abs(q_delta):.0f}pp）")

            # comment
            if   delta >= 15:  comment = f"表現顯著高於全卷平均（+{delta:.1f}pp），屬本次考試強項大題。"
            elif delta >= 5:   comment = f"表現略高於全卷平均（+{delta:.1f}pp），同學整體掌握尚可。"
            elif delta >= -5:  comment = f"表現與全卷平均相近（{delta:+.1f}pp），屬正常水平。"
            elif delta >= -15: comment = f"表現低於全卷平均（{delta:.1f}pp），建議加強相關概念的練習。"
            else:              comment = f"表現顯著低於全卷平均（{delta:.1f}pp），屬本次最弱大題，建議優先補底。"

            # worst sub-question
            sub_rates_list = []
            for q in qs:
                q_row = item_df[item_df[q_col] == q]
                if len(q_row):
                    r_val = q_row[rate_col].values[0] if rate_col else q_row["平均分"].values[0] / q_row["滿分"].values[0] * 100
                    sub_rates_list.append((q, r_val))
            weakness_str = ""
            if len(sub_rates_list) > 1:
                worst_q, worst_r = min(sub_rates_list, key=lambda x: x[1])
                wq_short = worst_q.split("_", 1)[1] if "_" in worst_q else worst_q
                weakness_str = f" 其中 {wq_short} 得分率 {worst_r:.0f}%，為本大題最弱小題。"

            # Line 1: title + stats + delta
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after  = Pt(0)
            p.paragraph_format.left_indent  = Cm(0.3)
            _no_page_break(p)
            r1 = p.add_run(f"{display}　");                _set_font(r1, 10.5, bold=True, color=rate_color)
            r2 = p.add_run(f"滿分 {int(mx)} 分　平均 {avg:.1f} 分　得分率 {rate:.1f}%（{diff}）　"); _set_font(r2, 10.5, color=(0, 0, 0))
            r3 = p.add_run(f"{delta_symbol} {abs(delta):.1f}pp（{"高於" if delta >= 0 else "低於"}全卷整體得分率 {overall_rate:.1f}%）"); _set_font(r3, 10.5, bold=True, color=delta_color)

            # Line 2: sub-question
            if sub_parts:
                p_sub = doc.add_paragraph()
                p_sub.paragraph_format.left_indent  = Cm(1.0)
                p_sub.paragraph_format.space_before = Pt(0)
                p_sub.paragraph_format.space_after  = Pt(0)
                _no_page_break(p_sub)
                r_sub = p_sub.add_run(f"各小題得分率：{"　".join(sub_parts)}")
                _set_font(r_sub, 9.5, color=(0, 0, 0))

            # Line 3: comment
            p_comm = doc.add_paragraph()
            p_comm.paragraph_format.left_indent  = Cm(1.0)
            p_comm.paragraph_format.space_before = Pt(0)
            p_comm.paragraph_format.space_after  = Pt(5)
            _no_page_break(p_comm)
            r_comm = p_comm.add_run(f"▸ {comment}{weakness_str}")
            _set_font(r_comm, 9.5, color=(0, 0, 0))

    best_grp  = group_df2.iloc[0]
    worst_grp = group_df2.iloc[-1]
    _add_body(doc,
        f"綜合而言，以「{best_grp['display']}」表現最佳"
        f"（{best_grp['rate']:.1f}%，▲{best_grp['rate'] - overall_rate:.1f}pp），"
        f"以「{worst_grp['display']}」表現最遜"
        f"（{worst_grp['rate']:.1f}%，▼{overall_rate - worst_grp['rate']:.1f}pp），"
        f"建議列為優先補底重點。"
    )
    _add_divider(doc)
    return group_df2, overall_rate, best_grp, worst_grp


def _build_tier_analysis(doc, total_scores, total_max, pass_threshold, n_present):
    _add_heading(doc, "四、學生分層分析")

    tier_top = total_scores[total_scores >= total_max * 0.8]
    tier_mid = total_scores[(total_scores >= pass_threshold) & (total_scores < total_max * 0.8)]
    tier_low = total_scores[total_scores < pass_threshold]

    tbl = doc.add_table(rows=4, cols=4)
    tbl.style = "Table Grid"
    cm_list = [(31, 71, 136), (0, 120, 60), (160, 100, 0), (180, 40, 40)]
    for ri, row_data in enumerate([
        ["分層", "標準", "人數", "佔出席比例"],
        ["優秀",   f"≥ {int(total_max * 0.8)} 分（80%+）",                  str(len(tier_top)), f"{len(tier_top) / n_present * 100:.1f}%"],
        ["合格",   f"{int(pass_threshold)}–{int(total_max * 0.8) - 1} 分",  str(len(tier_mid)), f"{len(tier_mid) / n_present * 100:.1f}%"],
        ["待改善", f"< {int(pass_threshold)} 分",                            str(len(tier_low)), f"{len(tier_low) / n_present * 100:.1f}%"],
    ]):
        for ci, val in enumerate(row_data):
            cell = tbl.rows[ri].cells[ci]
            cell.text = val
            run  = cell.paragraphs[0].runs[0]
            cell.paragraphs[0].paragraph_format.space_before = Pt(1)
            cell.paragraphs[0].paragraph_format.space_after  = Pt(1)
            _no_page_break(cell.paragraphs[0])
            _set_font(run, 9.5, bold=(ri == 0 or ci == 0),
                      color=cm_list[ri] if (ci == 0 or ri == 0) else (30, 30, 30))

    _add_spacer(doc, 1)

    if len(tier_low) > 0:
        _add_body(doc,
            f"共 {len(tier_low)} 名同學未達合格線，建議優先安排個別跟進（名單見附錄）。"
        )
    if len(tier_top) > 0:
        _add_body(doc,
            f"共 {len(tier_top)} 名同學表現優秀，可提供延伸練習（名單見附錄）。"
        )
    _add_divider(doc)
    return tier_top, tier_mid, tier_low


def _build_appendix(doc, tier_top, tier_low):
    """附錄：學生姓名名單（獨立一頁，對外傳閱前可整頁刪除）。"""
    doc.add_page_break()
    _add_heading(doc, "附錄：學生名單（內部參考，對外傳閱前可刪除本頁）")
    def _fmt(s):
        return "、".join(f"{n}（{v:g}）" for n, v in s.sort_values().round(1).items())
    if len(tier_low) > 0:
        _add_body(doc, f"未達合格線（{len(tier_low)} 人，括號內為總分）：{_fmt(tier_low)}。")
    else:
        _add_body(doc, "沒有同學未達合格線。")
    if len(tier_top) > 0:
        _add_body(doc, f"表現優秀（{len(tier_top)} 人，括號內為總分）：{_fmt(tier_top.sort_values(ascending=False))}。")


def _build_suggestions(doc, weak_qs, poor_disc, worst_grp, overall_rate,
                        tier_low, std_score, pass_pct, item_df, pass_line=40.0, std_pct=None):
    _add_heading(doc, "五、後續跟進建議")

    q_col    = "題目" if "題目" in item_df.columns else item_df.columns[0]
    rate_col = "得分率%" if "得分率%" in item_df.columns else None

    if rate_col is None:
        weak_qs = item_df.copy()
        weak_qs["得分率%"] = (weak_qs["平均分"] / weak_qs["滿分"] * 100).round(1)
        weak_qs = weak_qs[weak_qs["得分率%"] < 40].sort_values("得分率%")
        rate_col = "得分率%"

    suggestions = []

    if len(weak_qs) > 0:
        _wl = weak_qs[q_col].tolist()
        wq_names = "、".join(_pl(r) for r in _wl[:8])
        if len(_wl) > 8:
            wq_names += f" 等共 {len(_wl)} 題"
        suggestions.append(
            f"【重點補底】{wq_names}，得分率低於合格線（{pass_line:g}%），"
            f"建議下次課堂重點複習，並提供針對性練習。"
        )

    suggestions.append(
        f"【大題跟進】「{worst_grp['display']}」為得分率最低大題"
        f"（{worst_grp['rate']:.1f}%，▼{overall_rate - worst_grp['rate']:.1f}pp），"
        f"建議下一教學單元前安排專題複習。"
    )

    if len(tier_low) > 0:
        suggestions.append(
            f"【個別跟進】共 {len(tier_low)} 名同學未達合格線，"
            f"建議安排個別面談及補底支援。"
        )

    if poor_disc is not None and len(poor_disc) > 0:
        pd_names = "、".join([
            _pl(r) for r in poor_disc[q_col].tolist()
        ])
        suggestions.append(
            f"【題目檢視】{pd_names} 鑑別度低於 0.2 且難度適中，"
            f"建議下次命題時重新設計以提升評估效度。"
        )

    _sp = std_pct if std_pct is not None else std_score
    if _sp > 20:
        suggestions.append(
            f"【差異照顧】全班標準差 {std_score:.1f} 分（佔總滿分 {_sp:.1f}%），能力差距較大，"
            f"建議課堂加入分層練習，照顧不同程度學生。"
        )

    suggestions.append(
        f"【持續監察】建議下次考試後比較大題得分率趨勢，"
        f"追蹤各弱項大題的改善情況。"
    )

    for s in suggestions:
        _add_bullet(doc, s)

    # 教師備註欄
    _add_spacer(doc, 2)
    _add_heading(doc, "教師備註", level=2)
    _nt = doc.add_table(rows=1, cols=1)
    _nt.style = "Table Grid"
    _nt.rows[0].height = Cm(3.0)
    _nt.rows[0].cells[0].paragraphs[0].add_run(" ")

    _add_spacer(doc, 2)
    _add_divider(doc)

    p_footer = doc.add_paragraph()
    p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_footer.paragraph_format.space_before = Pt(1)
    _no_page_break(p_footer)
    r_footer = p_footer.add_run("本報告由試卷分析系統自動生成　如有疑問請核對原始數據")
    _set_font(r_footer, 8.5, color=(160, 160, 160))


# ══════════════════════════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════════════════════════

def generate_text_report(
    df, max_scores, paper_map, class_info, exam_info,
    absent_set, item_df, group_df, student_df, stats_df, extras=None
) -> bytes:
    """
    生成全班試卷分析文字報告（Word .docx）。

    Parameters
    ----------
    df          : pd.DataFrame  學生得分矩陣（index=姓名，columns=題目）
    max_scores  : pd.Series     各題滿分
    paper_map   : dict          題目 -> 試卷編號
    class_info  : pd.DataFrame  班別名冊（含 班別/班號/中文姓名）
    exam_info   : dict          考試資訊（subject_label, form_label, year_label, ...）
    absent_set  : set           缺席學生姓名集合
    item_df     : pd.DataFrame  試題分析表（含 題目/滿分/平均分/鑑別度 等欄位）
    group_df    : pd.DataFrame  大題分析表（來自 ea.question_group_analysis）
    student_df  : pd.DataFrame  學生成績表
    stats_df    : pd.DataFrame  全班統計表

    Returns
    -------
    bytes  Word docx 內容
    """
    # ── 欄位名稱統一（ea.item_analysis 的欄名為「題號」「鑑別度 D」「難度指數 P」）──
    item_df = item_df.copy()
    _ren = {}
    if "鑑別度" not in item_df.columns and "鑑別度 D" in item_df.columns:
        _ren["鑑別度 D"] = "鑑別度"
    if "題目" not in item_df.columns and "題號" in item_df.columns:
        _ren["題號"] = "題目"
    item_df = item_df.rename(columns=_ren)

    # ── 基本統計：總分一律取自學生成績表（多卷＝加權分；已計入扣分、二選一大題）──
    present_mask   = ~df.index.isin(absent_set)
    score_col      = "總分(加權)" if "總分(加權)" in student_df.columns else "總分"
    _sc            = pd.to_numeric(student_df[score_col], errors="coerce")
    total_scores   = _sc[~_sc.index.isin(absent_set)].dropna()
    total_max      = None
    if stats_df is not None and "統計項目" in stats_df.columns:
        _r = stats_df[(stats_df["區塊"] == "整體") & (stats_df["統計項目"] == "滿分")]
        if len(_r):
            total_max = float(_r["數值"].iloc[0])
    if not total_max:
        total_max = 100.0 if score_col == "總分(加權)" else float(max_scores.sum())
    total_max      = int(total_max) if float(total_max).is_integer() else float(total_max)
    pass_rate      = float(exam_info.get("pass_rate", 0.4))
    pass_threshold = total_max * pass_rate

    n_total   = len(df)
    n_absent  = len(absent_set)
    n_present = n_total - n_absent
    n_pass    = int((total_scores >= pass_threshold).sum())
    pass_pct  = n_pass / n_present * 100 if n_present else 0

    mean_score    = float(total_scores.mean())
    mean_pct      = mean_score / total_max * 100
    max_score_val = float(total_scores.max())
    min_score_val = float(total_scores.min())
    std_score     = float(total_scores.std())
    median_score  = float(total_scores.median())
    mc_q          = list((extras or {}).get("mc_info", {}) or {})

    # ── 試卷最大分拆解字串 ──
    papers = sorted(set(paper_map.values()))
    if len(papers) > 1:
        _w   = exam_info.get("paper_weights", {}) or {}
        _eff = (extras or {}).get("paper_max_eff", {}) or {}
        parts = []
        for p_key in papers:
            qs_in_p = [q for q, pv in paper_map.items() if pv == p_key]
            p_max   = float(_eff.get(p_key, sum(max_scores[q] for q in qs_in_p)))
            label   = p_key.replace("P1", "卷一").replace("P2", "卷二").replace("P3", "卷三").replace("P4", "卷四")
            wtxt    = f"，佔 {float(_w[p_key]) * 100:g}%" if p_key in _w else ""
            parts.append(f"{label} {p_max:g} 分{wtxt}")
        paper_max_str = "（加權計算；" + "　".join(parts) + "）" if score_col == "總分(加權)" else "（" + "　".join(parts) + "）"
    else:
        paper_max_str = ""

    # ── 弱項 / 強項（item_df 增補得分率欄）──
    _item = item_df.copy()
    q_col = "題目" if "題目" in _item.columns else _item.columns[0]
    if "得分率%" not in _item.columns:
        _item["得分率%"] = (_item["平均分"] / _item["滿分"] * 100).round(1)
    pass_line = pass_rate * 100
    weak_qs   = _item[_item["得分率%"] < pass_line].sort_values("得分率%")
    poor_disc = _genuine_poor(_item)

    # ── 建立 Word 文件 ──
    doc = Document()
    _normal = doc.styles["Normal"]
    _normal.font.name = LATIN_FONT
    _normal.element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), EAST_ASIA_FONT)
    for section in doc.sections:
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin   = Cm(1.8)
        section.right_margin  = Cm(1.8)

    _build_cover(doc, exam_info)
    _build_overview(
        doc, exam_info,
        n_total, n_absent, n_present, n_pass, pass_pct,
        mean_score, mean_pct, max_score_val, min_score_val,
        std_score, median_score, total_max, pass_threshold, paper_max_str
    )
    _build_item_analysis(doc, _item, mean_pct, pass_line=pass_line, mc_q=mc_q,
                         extras=extras, df=df, absent_set=absent_set)
    _, overall_rate, best_grp, worst_grp = _build_group_analysis(
        doc, df, max_scores, paper_map, absent_set, total_scores, _item,
        overall_rate=mean_pct, mc_q=mc_q, pass_line=pass_line)
    tier_top, tier_mid, tier_low = _build_tier_analysis(
        doc, total_scores, total_max, pass_threshold, n_present)
    _build_suggestions(
        doc, weak_qs, poor_disc, worst_grp, overall_rate,
        tier_low, std_score, pass_pct, _item,
        pass_line=pass_line, std_pct=std_score / total_max * 100 if total_max else None)
    _build_appendix(doc, tier_top, tier_low)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
