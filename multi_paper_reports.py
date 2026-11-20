"""
multi_paper_reports.py
多份試卷：每卷個人報告 ＋ 總成績報告；並提供「黑白列印版」（bw=True）。
所有個人報告均不顯示排名。
只依賴 individual_report.py 內既有的版面工具（不需要修改 individual_report.py）。
"""
import io, math
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib
matplotlib.use("Agg")
from matplotlib import pyplot as plt
from matplotlib.patches import Patch
import individual_report as ir

FONT = ir.DEFAULT_FONT
MAX_PER_ROW = 18
BW_HEADER = "D9D9D9"


def _label(q):
    return q.split("_", 1)[1] if "_" in q else q


def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=100, bbox_inches="tight", facecolor="white")
    buf.seek(0); plt.close(fig)
    return buf


def create_question_order_chart(df_student, max_scores, pass_rate=0.4,
                                max_per_row=MAX_PER_ROW, bw=False):
    """各題得分率（依題目順序；題多時分多個圖表）。回傳 (buf, strong, good, weak)。"""
    try:
        items = []
        for q in df_student.index:
            mx = float(max_scores[q])
            items.append((q, float(df_student[q]) / mx * 100 if mx > 0 else 0.0))
        if not items:
            return None, [], [], []
        pr = pass_rate * 100
        strong = [(q, r) for q, r in items if r >= 80]
        good   = [(q, r) for q, r in items if pr <= r < 80]
        weak   = [(q, r) for q, r in items if r < pr]
        if bw:
            style = lambda v: (("#4D4D4D", "") if v >= 80 else ("#BFBFBF", "") if v >= pr else ("white", "xxx"))
        else:
            style = lambda v: (("#2ecc71", "") if v >= 80 else ("#f39c12", "") if v >= pr else ("#e74c3c", ""))
        n = len(items)
        n_rows = math.ceil(n / max_per_row)
        per = math.ceil(n / n_rows)
        fig, axes = plt.subplots(n_rows, 1, figsize=(11, 2.9 * n_rows + 0.4), dpi=100, squeeze=False)
        for r, ax in enumerate(axes[:, 0]):
            lo = r * per
            chunk = items[lo:lo + per]
            labels = [c[0] for c in chunk]
            rates = [c[1] for c in chunk]
            for i, v in enumerate(rates):
                col, hatch = style(v)
                ax.bar(i, v, color=col, hatch=hatch, alpha=1 if bw else 0.85, width=0.62,
                       edgecolor="black", linewidth=0.8)
                ax.text(i, v + 2, f"{v:.0f}%", ha="center", va="bottom", fontsize=9, fontweight="bold")
            ax.set_xticks(range(len(chunk)))
            fs = 8.5 if max(len(s) for s in labels) > 6 else 10
            ax.set_xticklabels(labels, fontsize=fs, fontweight="bold")
            ax.set_xlim(-0.7, per - 0.3)
            ax.set_ylim(0, 118)
            ax.set_ylabel("得分率 %", fontsize=10, fontweight="bold")
            lc = "black" if bw else None
            ax.axhline(80, color=lc or "green", linestyle="--", alpha=0.5 if bw else 0.4, linewidth=1.0)
            ax.axhline(pr, color=lc or "orange", linestyle=":", alpha=0.7 if bw else 0.5, linewidth=1.2)
            ax.grid(axis="y", alpha=0.25, linestyle="--", linewidth=0.6)
            ax.set_axisbelow(True)
            part = f"（第 {lo + 1}–{lo + len(chunk)} 題）" if n_rows > 1 else ""
            ax.set_title("各題得分率（依題目順序）" + part, fontsize=12, fontweight="bold", loc="left")
        handles = [Patch(facecolor=style(90)[0], hatch=style(90)[1], edgecolor="black", label="優秀 ≥80%"),
                   Patch(facecolor=style(pr)[0], hatch=style(pr)[1], edgecolor="black", label=f"合格 ≥{pr:g}%"),
                   Patch(facecolor=style(0)[0], hatch=style(0)[1], edgecolor="black", label=f"不合格 <{pr:g}%")]
        fig.legend(handles=handles, loc="upper right", ncol=3, fontsize=9, frameon=False)
        fig.patch.set_facecolor("white")
        plt.tight_layout(rect=(0, 0, 1, 0.97))
        return _png(fig), strong, good, weak
    except Exception as e:
        print(f"   ⚠️  各題得分率圖表生成失敗：{e}")
        plt.close("all")
        return None, [], [], []


ir.create_strong_weak_analysis_chart_v2 = create_question_order_chart


def _progress_chart(student, mx, avg, bw=False):
    if not bw:
        return ir.create_single_bar_progress(student, mx, avg, mx)
    sp = student / mx * 100 if mx > 0 else 0
    ap = avg / mx * 100 if mx > 0 else 0
    fig, ax = plt.subplots(figsize=(11, 2), dpi=100)
    ax.barh(1, 100, height=0.3, color="white", edgecolor="black", linewidth=1)
    ax.barh(1, ap, height=0.25, color="#DDDDDD", edgecolor="black", hatch="///", linewidth=1.2,
            label=f"全班平均 ({ap:.1f}%)")
    ax.barh(1, sp, height=0.15, color="#222222", edgecolor="black", linewidth=1.2,
            label=f"個人成績 ({sp:.1f}%)")
    ax.text(sp + 2, 1.08, f"{ir.fmt_score(student)}/{ir.fmt_score(mx)}", fontsize=11, fontweight="bold", va="center")
    ax.text(ap + 2, 0.92, f"{ir.fmt_score(avg)}/{ir.fmt_score(mx)}", fontsize=10, va="center")
    ax.axvline(80, color="black", linestyle="--", alpha=0.5, linewidth=1.2)
    ax.text(80, 1.35, "80% 優秀線", fontsize=10, ha="center", fontweight="bold")
    ax.set_xlim(0, 110); ax.set_ylim(0.5, 1.5)
    ax.set_xlabel("得分率 %", fontsize=12, fontweight="bold")
    ax.set_title("個人成績與全班平均比較", fontsize=14, fontweight="bold", pad=15)
    ax.set_yticks([]); ax.legend(fontsize=11, loc="lower right")
    ax.grid(axis="x", alpha=0.3, linestyle="--", linewidth=0.8); ax.set_axisbelow(True)
    fig.patch.set_facecolor("white"); plt.tight_layout()
    return _png(fig)


def _paper_bar_chart(papers, stu_pct, avg_pct, pass_rate, bw=False):
    fig, ax = plt.subplots(figsize=(10, 1.4 + 1.1 * len(papers)), dpi=100)
    y = np.arange(len(papers))[::-1]
    if bw:
        ax.barh(y + 0.18, avg_pct, height=0.32, color="#DDDDDD", edgecolor="black", hatch="///", label="全班平均")
        ax.barh(y - 0.18, stu_pct, height=0.32, color="#222222", edgecolor="black", label="個人")
        c_avg = c_stu = c_line = "black"
    else:
        ax.barh(y + 0.18, avg_pct, height=0.32, color="#f39c12", alpha=0.85, label="全班平均")
        ax.barh(y - 0.18, stu_pct, height=0.32, color="#3498db", alpha=0.95, label="個人")
        c_avg, c_stu, c_line = "#b9770e", "#1f618d", "#e74c3c"
    for yy, a_, s_ in zip(y, avg_pct, stu_pct):
        ax.text(a_ + 1, yy + 0.18, f"{a_:.1f}%", va="center", fontsize=10, color=c_avg)
        ax.text(s_ + 1, yy - 0.18, f"{s_:.1f}%", va="center", fontsize=10, color=c_stu, fontweight="bold")
    ax.axvline(pass_rate * 100, color=c_line, linestyle="--", alpha=0.6)
    ax.text(pass_rate * 100, len(papers) - 0.35, f"合格線 {pass_rate * 100:g}%", color=c_line,
            fontsize=9, ha="center")
    ax.set_yticks(y); ax.set_yticklabels(papers, fontsize=12, fontweight="bold")
    ax.set_xlim(0, 112); ax.set_xlabel("得分率 %", fontsize=11)
    ax.set_title("各卷得分率：個人與全班平均", fontsize=13, fontweight="bold")
    ax.grid(axis="x", alpha=0.3, linestyle="--"); ax.set_axisbelow(True)
    ax.legend(loc="lower right", fontsize=10)
    fig.patch.set_facecolor("white"); plt.tight_layout()
    return _png(fig)


def compute_tables(df, max_scores, paper_map, paper_weights, extras, absent_set):
    """各卷得分（含扣分）、有效滿分、得分率、加權總分、全班平均。"""
    extras = extras or {}
    papers = sorted(set(paper_map.values()))
    present = ~df.index.isin(absent_set or set())
    raw = pd.DataFrame(index=df.index)
    mx = {}
    for p in papers:
        cols = [q for q in df.columns if paper_map[q] == p]
        s = df[cols].sum(axis=1)
        d = extras.get("deductions", {}).get(p)
        if d is not None:
            s = s + d.reindex(df.index).fillna(0.0)
        raw[p] = s
        mx[p] = float(extras.get("paper_max_eff", {}).get(p, max_scores[cols].sum()))
    mx = pd.Series(mx)
    pct = raw.div(mx) * 100
    w = pd.Series({p: float(paper_weights.get(p, 0)) for p in papers})
    weighted = (pct * w).sum(axis=1)
    return {
        "papers": papers, "raw": raw, "max": mx, "pct": pct, "weights": w,
        "weighted": weighted, "present": present, "n_present": int(present.sum()),
        "avg_raw": raw[present].mean(), "avg_pct": pct[present].mean(),
        "avg_weighted": float(weighted[present].mean()),
    }


def absent_doc(label="", bw=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(1.0)
    sec.left_margin = sec.right_margin = Inches(1.2)
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("本次缺席")
    r.font.size = Pt(36); r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0) if bw else RGBColor(0xC0, 0x39, 0x2B)
    if label:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run(label)
        r2.font.size = Pt(14)
    return doc


def _cell_fmt(cell, size=10, bold=False, color=None, bg=None, align="c"):
    if bg:
        ir.set_cell_background(cell, bg)
    for para in cell.paragraphs:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "c" else WD_ALIGN_PARAGRAPH.LEFT
        for run in para.runs:
            run.font.size = Pt(size); run.font.bold = bold; run.font.name = FONT
            if color:
                run.font.color.rgb = RGBColor(*color)


def _header_cells(tbl, texts, bw, size=10):
    for c, t in zip(tbl.rows[0].cells, texts):
        c.text = t
        if bw:
            _cell_fmt(c, size, True, (0, 0, 0), BW_HEADER)
        else:
            _cell_fmt(c, size, True, (255, 255, 255), "1f4e78")


def _paper_doc(name, title, score, paper_max, stu, mx_sub, skip, notes, ded,
               avg_score, class_info, pass_rate, bw=False):
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Inches(11.69); sec.page_width = Inches(8.27)
    sec.top_margin = sec.bottom_margin = Inches(0.3)
    sec.left_margin = sec.right_margin = Inches(0.3)

    def line(text, size, bold=False, color=None):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT
        c = (0, 0, 0) if bw else color
        if c:
            r.font.color.rgb = RGBColor(*c)
        return p

    line(title, 16, True, (25, 25, 112))
    line("學生個人成績報告", 12, True)
    cn = cnum = "N/A"
    if class_info is not None:
        m = class_info[class_info["中文姓名"] == name]
        if not m.empty:
            cn, cnum = str(m.iloc[0]["班別"]), str(m.iloc[0]["班號"])
    line(f"班別：{cn}  班號：{cnum}  姓名：{name}  "
         f"總分：{ir.fmt_score(score)}/{ir.fmt_score(paper_max)}  "
         f"({score / paper_max * 100:.1f}%)", 11, True)
    doc.add_paragraph()

    tbl = doc.add_table(rows=1, cols=4)
    tbl.style = "Table Grid" if bw else "Light Grid Accent 1"
    tbl.autofit = False
    _header_cells(tbl, ["題號", "滿分", "得分", "得分率"], bw)
    for c, w in zip(tbl.rows[0].cells, [1.4, 0.8, 0.8, 1.0]):
        c.width = Inches(w)

    shown = []
    for q in stu.index:
        if q in skip:
            continue
        sc, mx = float(stu[q]), float(mx_sub[q])
        shown.append(q)
        cells = tbl.add_row().cells
        cells[0].text = str(q) + (f"  {notes[q]}" if q in notes else "")
        cells[1].text = ir.fmt_score(mx)
        cells[2].text = ir.fmt_score(sc)
        cells[3].text = f"{sc / mx * 100:.1f}%" if mx > 0 else "0.0%"
        for c in cells:
            _cell_fmt(c, 9, color=(0, 0, 0) if bw else None)
    if ded:
        cells = tbl.add_row().cells
        for c, t in zip(cells, ["扣分", "-", ir.fmt_score(ded), "-"]):
            c.text = t; _cell_fmt(c, 9, color=(0, 0, 0) if bw else None)
    doc.add_paragraph()

    try:
        buf = _progress_chart(score, paper_max, avg_score, bw)
        if buf:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_picture(buf, width=Inches(7.5)); buf.close()
    except Exception:
        pass
    doc.add_paragraph()
    out = create_question_order_chart(stu[shown], mx_sub[shown], pass_rate, bw=bw)
    if out[0]:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_picture(out[0], width=Inches(7.5)); out[0].close()
    return doc


def build_paper_report(name, p, T, df, max_scores, paper_map, extras, exam_info,
                       class_info, pass_rate, item_df=None, bw=False, single=False):
    """每卷個人報告（不顯示排名）。single=True：單份試卷（標題不加比重）。"""
    extras = extras or {}
    cols = [q for q in df.columns if paper_map[q] == p]
    att = extras.get("attempted")
    skip = set()
    if att is not None:
        for q in cols:
            if not att[q].all() and not bool(att.loc[name, q]):
                skip.add(q)
    labels = {q: (q if single else _label(q)) for q in cols}
    if len(set(labels.values())) < len(cols):
        labels = {q: q for q in cols}

    stu = df.loc[name, cols].copy()
    stu.index = [labels[q] for q in cols]
    mx_sub = pd.Series({labels[q]: float(max_scores[q]) for q in cols})

    notes = {}   # 個人報告不顯示 MC 的選項／答案

    ded = 0.0
    if p in extras.get("deductions", {}):
        ded = float(extras["deductions"][p].get(name, 0.0))
    score = float(T["raw"].loc[name, p])
    paper_max = float(T["max"][p])
    if single:
        title = exam_info["exam_title"]
    else:
        title = f"{exam_info['exam_title']}　{p}（佔總成績 {float(T['weights'][p]) * 100:g}%）"
    return _paper_doc(name, title, score, paper_max, stu, mx_sub,
                      {labels[q] for q in skip}, notes, ded,
                      float(T["avg_raw"][p]), class_info, pass_rate, bw=bw)


def build_total_report(name, T, df, extras, exam_info, class_info, pass_rate, bw=False):
    papers = T["papers"]
    extras = extras or {}
    doc = Document()
    sec = doc.sections[0]
    sec.page_height = Inches(11.69); sec.page_width = Inches(8.27)
    sec.top_margin = sec.bottom_margin = Inches(0.3)
    sec.left_margin = sec.right_margin = Inches(0.3)

    def para(text, size=11, bold=False, color=None, align=WD_ALIGN_PARAGRAPH.CENTER, space=2):
        p = doc.add_paragraph(); p.alignment = align
        p.paragraph_format.space_after = Pt(space)
        r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT
        c = (0, 0, 0) if bw else color
        if c:
            r.font.color.rgb = RGBColor(*c)
        return p

    para(exam_info["exam_title"], 16, True, (25, 25, 112))
    para("學生總成績報告", 12, True)
    cn = cnum = "N/A"
    if class_info is not None:
        m = class_info[class_info["中文姓名"] == name]
        if not m.empty:
            cn, cnum = str(m.iloc[0]["班別"]), str(m.iloc[0]["班號"])
    para(f"班別：{cn}  班號：{cnum}  姓名：{name}", 11, True)

    wscore = float(T["weighted"][name]); avg_w = T["avg_weighted"]
    passed = wscore >= pass_rate * 100
    para(f"加權總分：{wscore:.1f} / 100　　{'合格' if passed else '不合格'}", 15, True,
         (0, 110, 50) if passed else (192, 57, 43), space=6)

    hdr = ["試卷", "比重", "得分 / 滿分", "得分率", "全班平均", "與平均差距"]
    tbl = doc.add_table(rows=1, cols=len(hdr))
    tbl.style = "Table Grid" if bw else "Light Grid Accent 1"
    _header_cells(tbl, hdr, bw)
    for p in papers:
        s = float(T["raw"].loc[name, p]); m = float(T["max"][p])
        pc = float(T["pct"].loc[name, p]); ap = float(T["avg_pct"][p])
        d = pc - ap
        cells = tbl.add_row().cells
        vals = [p, f"{T['weights'][p] * 100:g}%", f"{ir.fmt_score(s)} / {ir.fmt_score(m)}",
                f"{pc:.1f}%", f"{ap:.1f}%", f"{d:+.1f}pp"]
        for c, t in zip(cells, vals):
            c.text = t
        for i, c in enumerate(cells):
            col = (0, 0, 0)
            if i == 5 and not bw:
                col = (0, 110, 50) if d >= 0 else (192, 57, 43)
            _cell_fmt(c, 10, i in (0, 3), col)
    doc.add_paragraph()

    buf = _progress_chart(wscore, 100.0, avg_w, bw)
    if buf:
        doc.add_picture(buf, width=Inches(7.5)); buf.close()
    cb = _paper_bar_chart(papers, [float(T["pct"].loc[name, p]) for p in papers],
                          [float(T["avg_pct"][p]) for p in papers], pass_rate, bw=bw)
    doc.add_picture(cb, width=Inches(7.0)); cb.close()

    formula = " ＋ ".join(f"{p}得分率 × {T['weights'][p] * 100:g}%" for p in papers)
    para(f"計算方式：加權總分 ＝ {formula}", 9.5, False, (90, 90, 90), WD_ALIGN_PARAGRAPH.LEFT)
    ded_txt = []
    for p in papers:
        d = extras.get("deductions", {}).get(p)
        if d is not None and float(d.get(name, 0)) != 0:
            ded_txt.append(f"{p} 扣 {abs(float(d[name])):g} 分")
    if ded_txt:
        para("已計入扣分：" + "、".join(ded_txt), 9.5, False, (90, 90, 90), WD_ALIGN_PARAGRAPH.LEFT)
    diffs = {p: float(T["pct"].loc[name, p] - T["avg_pct"][p]) for p in papers}
    best = max(diffs, key=diffs.get); worst = min(diffs, key=diffs.get)
    if len(papers) > 1:
        if min(diffs.values()) >= 0:
            msg = f"各卷均不低於全班平均，其中 {best} 最突出（{diffs[best]:+.1f}pp）。"
        elif max(diffs.values()) < 0:
            msg = f"各卷均低於全班平均，差距最大的是 {worst}（{diffs[worst]:+.1f}pp），宜優先跟進。"
        else:
            msg = (f"相對全班平均，{best} 表現較佳（{diffs[best]:+.1f}pp），"
                   f"{worst} 較需加強（{diffs[worst]:+.1f}pp）。")
        para("重點：" + msg, 10.5, True, (25, 25, 112), WD_ALIGN_PARAGRAPH.LEFT)
    return doc
