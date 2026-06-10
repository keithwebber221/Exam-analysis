"""
class_item_analysis.py
全級班際項目分析模組（第二階段）
依賴：pandas, numpy, openpyxl
整合方式：由 app.py 呼叫 generate_class_analysis_excel() / get_class_summary_df()
"""

import io
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ══════════════════════════════════════════════════════════════
# 配色常數（最多支援 8 班）
# ══════════════════════════════════════════════════════════════
PALETTE = [
    ("1F9E9E", "E0F4F4", "1A6B6B"),   # 5A 深藍綠
    ("6C5FD1", "ECEAF8", "4A3F8C"),   # 5B 深紫
    ("E05A4E", "FBECEB", "C0392B"),   # 5C 深紅
    ("D4A153", "FDF3E3", "B87333"),   # 5D 棕橙
    ("E67E22", "FEF0E6", "A04000"),
    ("27AE60", "E9F7EF", "1A6B3A"),
    ("2980B9", "EAF4FB", "1A5276"),
    ("8E44AD", "F5EEF8", "6C3483"),
]

# 各班熱力圖標題色（固定，不隨palette循環）
HEATMAP_HEADER_FILLS = [
    ("1F3864", "FFFFFF"),  # 5A 深海軍藍
    ("2E75B6", "FFFFFF"),  # 5B 中藍
    ("843C0C", "FFFFFF"),  # 5C 深棕紅
    ("375623", "FFFFFF"),  # 5D 深森綠
    ("7F6000", "FFFFFF"),
    ("1A5276", "FFFFFF"),
    ("4A235A", "FFFFFF"),
    ("1A3A6B", "FFFFFF"),
]

def _build_palette(classes):
    return {cls: PALETTE[i % len(PALETTE)] for i, cls in enumerate(classes)}

# ══════════════════════════════════════════════════════════════
# 樣式工具
# ══════════════════════════════════════════════════════════════
def _cfill(hex6):
    return PatternFill("solid", fgColor=hex6)

NO_FILL = PatternFill(fill_type=None)

def _fnt(bold=False, size=12, color="000000", name="Times New Roman"):
    return Font(bold=bold, size=size, color=color, name=name)

def _hdr_fnt(bold=True, size=10, color="000000"):
    """標題用字體（新細明體）"""
    return Font(bold=bold, size=size, color=color, name="新細明體")

_CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LFT = Alignment(horizontal="left",   vertical="center", indent=1)

def _thin_border(color="BBCAD6"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _med_border(color="7F9AB5"):
    s = Side(style="medium", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _sc(cell, val=None, fill=None, font=None, aln=None, bdr=None, fmt=None):
    if val  is not None: cell.value        = val
    if fill is not None: cell.fill         = fill
    if font is not None: cell.font         = font
    if aln  is not None: cell.alignment    = aln
    if bdr  is not None: cell.border       = bdr
    if fmt  is not None: cell.number_format= fmt

def _set_merged(ws, r1, c1, r2, c2, fill_hex, font_obj, val, fmt="@", bdr=None):
    if bdr is None: bdr = _thin_border()
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    none_s = Side(style=None)
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            cell = ws.cell(r, c)
            cell.fill      = _cfill(fill_hex)
            cell.alignment = _CTR
            cell.border = Border(
                left   = bdr.left   if c == c1 else none_s,
                right  = bdr.right  if c == c2 else none_s,
                top    = bdr.top    if r == r1 else none_s,
                bottom = bdr.bottom if r == r2 else none_s,
            )
    main = ws.cell(r1, c1)
    main.value = val; main.font = font_obj
    main.number_format = fmt; main.alignment = _CTR

def _make_bar(pct, width=20):
    if pct is None or (isinstance(pct, float) and (pct != pct)):  # NaN check
        return "░" * width
    pct = float(pct)
    filled = max(0, min(width, round(pct * width)))
    return "█" * filled + "░" * (width - filled)

def _diff_fill(d):
    """差距欄填色（與範本v12完全一致）"""
    if d is None or (isinstance(d, float) and d != d): d = 0.0
    if d >=  0.08: return _cfill("C6EFCE")   # 強正：淡綠
    if d >=  0.0:  return _cfill("FFEB9C")   # 弱正：淡黃
    if d >= -0.05: return _cfill("FFF0F0")   # 弱負：極淡紅
    if d >= -0.10: return _cfill("FFE0E0")   # 中負：淡紅
    return _cfill("FFD0D0")                   # 強負：較深紅

def _diff_color(d):
    """差距條字色"""
    if d is None or (isinstance(d, float) and d != d): d = 0.0
    if d >= 0: return "27AE60"   # 正差距：綠
    return "E74C3C"              # 負差距：紅

def _diff_pct_color(d):
    """差距%數值字色"""
    if d is None or (isinstance(d, float) and d != d): d = 0.0
    if d >= 0: return "375623"
    return "9C0006"

def _make_center_bar(d, scale=0.25, half=10):
    """中線式差距條（範本v12風格）：░░░░│████░░░  總長=half*2+1"""
    if d is None or (isinstance(d, float) and d != d): d = 0.0
    filled = max(0, min(half, round(abs(d) / scale * half)))
    if d >= 0:
        # 正差距：│右邊填
        return "░" * half + "│" + "█" * filled + "░" * (half - filled)
    else:
        # 負差距：│左邊填（靠近│）
        return "░" * (half - filled) + "█" * filled + "│" + "░" * half

# ══════════════════════════════════════════════════════════════
# 統計計算
# ══════════════════════════════════════════════════════════════
def _compute_stats(df, q_cols, max_scores, paper_map, class_col):
    classes = sorted(class_col.unique().tolist())
    df2 = df[q_cols].copy()
    df2["_cls"] = class_col.values

    def _safe_mean(s):
        v = s.dropna().mean()
        return 0.0 if (v != v) else float(v)
    def _safe_std(s):
        v = s.dropna().std(ddof=1) if len(s.dropna()) > 1 else 0.0
        return 0.0 if (v != v) else float(v)

    class_stats = {}
    for cls in classes:
        sub = df2[df2["_cls"] == cls][q_cols]
        class_stats[cls] = {}
        for q in q_cols:
            mx = float(max_scores[q])
            m  = _safe_mean(sub[q])
            class_stats[cls][q] = {
                "mean":     round(m, 2),
                "mean_pct": m / mx if mx > 0 else 0.0,
                "std":      round(_safe_std(sub[q]), 2),
            }

    grade_stats = {}
    for q in q_cols:
        mx = float(max_scores[q])
        m  = _safe_mean(df2[q])
        grade_stats[q] = {
            "mean":     round(m, 2),
            "mean_pct": m / mx if mx > 0 else 0.0,
            "std":      round(_safe_std(df2[q]), 2),
        }

    papers = sorted(set(paper_map.values()))

    def _pt_stats(sub_df, p_cols):
        raw = sub_df[p_cols].sum(axis=1)
        mx  = float(max_scores[p_cols].sum())
        return {"mean": round(raw.mean(),2), "std": round(raw.std(ddof=1),2),
                "max": int(raw.max()), "min": int(raw.min()),
                "max_score": mx, "mean_pct": raw.mean()/mx if mx>0 else 0}

    paper_totals = {}
    for p in papers:
        p_cols = [q for q in q_cols if paper_map[q] == p]
        if not p_cols: continue
        paper_totals[p] = {cls: _pt_stats(df2[df2["_cls"]==cls], p_cols) for cls in classes}
        paper_totals[p]["grade"] = _pt_stats(df2, p_cols)

    combined_totals = {cls: _pt_stats(df2[df2["_cls"]==cls], q_cols) for cls in classes}
    combined_totals["grade"] = _pt_stats(df2, q_cols)

    return class_stats, grade_stats, classes, paper_totals, combined_totals

# ══════════════════════════════════════════════════════════════
# 各班分析工作表
# ══════════════════════════════════════════════════════════════
def _make_class_sheet(wb, cls, q_cols, max_scores, paper_map,
                      class_stats, grade_stats, df, class_col, palette):
    bar_hex, bg_hex, dark_hex = palette[cls]
    other_classes = [c for c in sorted(class_col.unique()) if c != cls]
    df_other = df[class_col != cls]
    other_stats = {}
    for q in q_cols:
        mx   = float(max_scores[q])
        omean = df_other[q].dropna().mean() if len(df_other) > 0 else 0.0
        omean = 0.0 if (omean != omean) else omean   # NaN → 0
        ostd  = df_other[q].dropna().std(ddof=1) if len(df_other) > 1 else 0.0
        ostd  = 0.0 if (ostd != ostd) else ostd
        other_stats[q] = {
            "mean":     round(omean, 2),
            "mean_pct": omean / mx if mx > 0 else 0.0,
            "std":      round(ostd, 2),
        }

    papers = sorted(set(paper_map.values()))
    scale_map = {}
    for p in papers:
        p_qs = [q for q in q_cols if paper_map[q] == p]
        max_d = max((abs(class_stats[cls][q]["mean_pct"] - grade_stats[q]["mean_pct"]) for q in p_qs), default=0.05)
        scale_map[p] = max(0.05, round(max_d + 0.02, 2))

    ws = wb.create_sheet(f"{cls} 分析")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    n_cls   = int((class_col == cls).sum())
    n_other = len(df_other)
    n_all   = len(df)
    other_label = "+".join(other_classes)
    NCOLS = 18
    THIN  = _thin_border()

    ws.row_dimensions[2].height = 34
    ws.merge_cells(f"B2:{get_column_letter(1+NCOLS)}2")
    _sc(ws["B2"], val=f"{cls}班　項目分析　·　本班 vs 全級　／　本班 vs {other_label}合拼",
        fill=_cfill(dark_hex), font=_hdr_fnt(bold=True, size=14, color="FFFFFF"), aln=_CTR)
    ws.row_dimensions[3].height = 15
    ws.merge_cells(f"B3:{get_column_letter(1+NCOLS)}3")
    _sc(ws["B3"], val=f"本班人數：{n_cls} 人　　{other_label} 人數：{n_other} 人",
        font=Font(size=10, color="595959", name="新細明體"), aln=_CTR, fill=NO_FILL)

    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 18
    col = 2
    for lbl in ["題號", "滿分"]:
        ws.merge_cells(start_row=4, start_column=col, end_row=5, end_column=col)
        _sc(ws.cell(4, col), val=lbl, fill=_cfill("D9D9D9"),
            font=_hdr_fnt(bold=True, size=12, color="1F3864"), aln=_CTR)
        col += 1

    headers = [
        (f"本班（{cls}）", dark_hex, bg_hex, dark_hex, ["人數","平均分","平均%","平均%圖表","S.D."], 5),
        (f"全　級（{cls}+{other_label}）", "5D6A77", "EAECEE","3D4E5C",["平均分","平均%","平均%圖表","S.D."],       4),
        ("本班 vs 全級",   "2C3E50", "D5F5E3","1E8449",["差距%","差距圖表"],                        2),
        (f"{other_label} 合拼","6C3483","F5EEF8","6C3483",["平均分","平均%","平均%圖表","S.D."],    4),
        (f"本班 vs {other_label}","6C3483","E8DAEF","6C3483",["差距%","差距圖表"],                  2),
    ]
    for title, title_bg, sub_bg, sub_fg, sub_labels, span in headers:
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col+span-1)
        _sc(ws.cell(4, col), val=title, fill=_cfill(title_bg),
            font=_hdr_fnt(bold=True, size=12, color="FFFFFF"), aln=_CTR)
        for o, lbl in enumerate(sub_labels):
            _sc(ws.cell(5, col+o), val=lbl,
                fill=_cfill(sub_bg), font=_hdr_fnt(bold=True, size=10, color=sub_fg), aln=_CTR)
        col += span

    prev_paper = None
    data_row   = 6
    BARS = 10  # legacy，差距條已改用 _make_center_bar

    for ri, q in enumerate(q_cols):
        paper = paper_map[q]
        scale = scale_map[paper]
        mx    = float(max_scores[q])
        cs    = class_stats[cls][q]
        gs    = grade_stats[q]
        os_   = other_stats[q]
        bg    = _cfill("F7F9FB") if ri % 2 == 1 else NO_FILL

        if paper != prev_paper:
            ws.row_dimensions[data_row].height = 15
            p_qs  = [q2 for q2 in q_cols if paper_map[q2] == paper]
            p_max = float(max_scores[p_qs].sum())
            ws.merge_cells(start_row=data_row, start_column=2, end_row=data_row, end_column=1+NCOLS)
            _sc(ws.cell(data_row, 2),
                val=f"　{paper}　·　滿分 {int(p_max)}　（本班 vs 全級差距圖表刻度：±{int(round(scale*100))}%）",
                fill=_cfill(dark_hex), font=_hdr_fnt(bold=True, size=11, color="FFFFFF"), aln=_LFT)
            data_row += 1
            prev_paper = paper

        ws.row_dimensions[data_row].height = 18
        dc = 2
        _sc(ws.cell(data_row, dc), val=q,       fill=bg, font=_fnt(bold=True,size=11), aln=_CTR, bdr=THIN); dc+=1
        _sc(ws.cell(data_row, dc), val=int(mx), fill=bg, font=_fnt(size=12),            aln=_CTR, bdr=THIN, fmt="0"); dc+=1
        _sc(ws.cell(data_row, dc), val=n_cls,          fill=bg, font=_fnt(size=12), aln=_CTR, bdr=THIN, fmt="0"); dc+=1
        _sc(ws.cell(data_row, dc), val=cs["mean"],     fill=bg, font=_fnt(size=12), aln=_CTR, bdr=THIN, fmt="0.00"); dc+=1
        _sc(ws.cell(data_row, dc), val=cs["mean_pct"], fill=bg, font=_fnt(size=12), aln=_CTR, bdr=THIN, fmt="0.0%"); dc+=1
        c = ws.cell(data_row, dc); c.value=_make_bar(cs["mean_pct"]); c.font=Font(name="Courier New",size=10,color=bar_hex); c.alignment=_LFT; c.border=THIN; c.fill=bg; dc+=1
        _sc(ws.cell(data_row, dc), val=cs["std"],      fill=bg, font=_fnt(size=12), aln=_CTR, bdr=THIN, fmt="0.00"); dc+=1

        gf = _cfill("F2F3F4")
        _sc(ws.cell(data_row, dc), val=gs["mean"],     fill=gf, font=_fnt(size=12,color="3D4E5C"), aln=_CTR, bdr=THIN, fmt="0.00"); dc+=1
        _sc(ws.cell(data_row, dc), val=gs["mean_pct"], fill=gf, font=_fnt(size=12,color="3D4E5C"), aln=_CTR, bdr=THIN, fmt="0.0%"); dc+=1
        c = ws.cell(data_row, dc); c.value=_make_bar(gs["mean_pct"]); c.font=Font(name="Courier New",size=10,color="7D8B99"); c.alignment=_LFT; c.border=THIN; c.fill=gf; dc+=1
        _sc(ws.cell(data_row, dc), val=gs["std"],      fill=gf, font=_fnt(size=12,color="3D4E5C"), aln=_CTR, bdr=THIN, fmt="0.00"); dc+=1

        d1 = cs["mean_pct"] - gs["mean_pct"]
        d1 = 0.0 if (d1 != d1) else d1
        _sc(ws.cell(data_row, dc), val=d1, fill=_diff_fill(d1), font=_fnt(bold=True,size=12,color=_diff_pct_color(d1)), aln=_CTR, bdr=THIN, fmt="+0.0%;-0.0%;0.0%"); dc+=1
        bar1 = _make_center_bar(d1)
        c1 = ws.cell(data_row, dc); c1.value=bar1; c1.font=Font(name="Courier New",size=11,color=_diff_color(d1)); c1.alignment=_CTR; c1.border=THIN; c1.fill=_diff_fill(d1); dc+=1

        of = _cfill("F5EEF8")
        _sc(ws.cell(data_row, dc), val=os_["mean"],     fill=of, font=_fnt(size=12,color="6C3483"), aln=_CTR, bdr=THIN, fmt="0.00"); dc+=1
        _sc(ws.cell(data_row, dc), val=os_["mean_pct"], fill=of, font=_fnt(size=12,color="6C3483"), aln=_CTR, bdr=THIN, fmt="0.0%"); dc+=1
        c = ws.cell(data_row, dc); c.value=_make_bar(os_["mean_pct"]); c.font=Font(name="Courier New",size=10,color="8E6BB0"); c.alignment=_LFT; c.border=THIN; c.fill=of; dc+=1
        _sc(ws.cell(data_row, dc), val=os_["std"],      fill=of, font=_fnt(size=12,color="6C3483"), aln=_CTR, bdr=THIN, fmt="0.00"); dc+=1

        d2 = cs["mean_pct"] - os_["mean_pct"]
        d2 = 0.0 if (d2 != d2) else d2
        _sc(ws.cell(data_row, dc), val=d2, fill=_diff_fill(d2), font=_fnt(bold=True,size=12,color=_diff_pct_color(d2)), aln=_CTR, bdr=THIN, fmt="+0.0%;-0.0%;0.0%"); dc+=1
        bar2 = _make_center_bar(d2)
        c2 = ws.cell(data_row, dc); c2.value=bar2; c2.font=Font(name="Courier New",size=11,color=_diff_color(d2)); c2.alignment=_CTR; c2.border=THIN; c2.fill=_diff_fill(d2)

        data_row += 1

    widths = [8,6, 7,9,9,22,8, 9,9,22,8, 10,36, 9,9,22,8, 10,36]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(2+i)].width = w
    ws.freeze_panes = "B6"

# ══════════════════════════════════════════════════════════════
# 分數分佈工作表
# ══════════════════════════════════════════════════════════════
def _make_distribution_sheet(wb, df, max_scores, paper_map, class_col, classes, palette):
    ws = wb.create_sheet("分數分佈")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    q_cols   = list(max_scores.index)
    n_cls_cols = len(classes)
    total_cols = 2 + n_cls_cols * 2 + 2
    end_c = get_column_letter(1 + total_cols)

    ws.row_dimensions[2].height = 32
    ws.merge_cells(f"B2:{end_c}2")
    _sc(ws["B2"], val="分數分佈　·　各卷及合併總分",
        fill=_cfill("1F3864"), font=_hdr_fnt(bold=True, size=14, color="FFFFFF"), aln=_CTR)

    papers = sorted(set(paper_map.values()))
    sections = []
    for p in papers:
        p_qs  = [q for q in q_cols if paper_map[q]==p]
        sections.append((p, df[p_qs].sum(axis=1), float(max_scores[p_qs].sum())))
    if len(papers) > 1:
        sections.append(("合併", df[q_cols].sum(axis=1), float(max_scores[q_cols].sum())))

    THICK_B = _med_border("7F9AB5")
    STAT_B  = _thin_border("5D8AA8")
    cur_row = 4

    for sec_label, raw_series, s_max in sections:
        pass_th = s_max * 0.4
        ws.row_dimensions[cur_row].height = 22
        ws.merge_cells(start_row=cur_row, start_column=2, end_row=cur_row, end_column=1+total_cols)
        _sc(ws.cell(cur_row,2),
            val=f"◆  {sec_label}　　滿分：{int(s_max)}　　及格線：{round(pass_th)} 分（40%）",
            fill=_cfill("2C3E50"), font=_hdr_fnt(bold=True, size=12, color="FFFFFF"),
            aln=Alignment(horizontal="left",vertical="center",indent=1))
        cur_row += 1

        ws.row_dimensions[cur_row].height = 20
        ws.row_dimensions[cur_row+1].height = 18
        col = 2
        for lbl in ["分佈（%）","分數範圍"]:
            ws.merge_cells(start_row=cur_row, start_column=col, end_row=cur_row+1, end_column=col)
            _sc(ws.cell(cur_row,col), val=lbl, fill=_cfill("D9D9D9"),
                font=_hdr_fnt(bold=True, size=11, color="1F3864"), aln=_CTR)
            col += 1
        for cls in classes:
            _, bg_hex, dark_hex = palette[cls]
            ws.merge_cells(start_row=cur_row, start_column=col, end_row=cur_row, end_column=col+1)
            _sc(ws.cell(cur_row,col), val=cls, fill=_cfill(dark_hex),
                font=_hdr_fnt(bold=True, size=11, color="FFFFFF"), aln=_CTR)
            for o, lbl in enumerate(["人數 (%)","高於本段人數 (%)"]):
                _sc(ws.cell(cur_row+1,col+o), val=lbl, fill=_cfill(bg_hex),
                    font=_hdr_fnt(bold=True, size=9, color=dark_hex), aln=_CTR)
            col += 2
        ws.merge_cells(start_row=cur_row, start_column=col, end_row=cur_row, end_column=col+1)
        _sc(ws.cell(cur_row,col), val="全　級", fill=_cfill("5D6A77"),
            font=_hdr_fnt(bold=True, size=11, color="FFFFFF"), aln=_CTR)
        for o, lbl in enumerate(["人數 (%)","高於本段人數 (%)"]):
            _sc(ws.cell(cur_row+1,col+o), val=lbl, fill=_cfill("EAECEE"),
                font=_hdr_fnt(bold=True, size=9, color="3D4E5C"), aln=_CTR)
        cur_row += 2

        bins   = [i/10 for i in range(0,11)]
        bin_lo = [round(b*s_max) for b in bins[:-1]]
        bin_hi = [round((b+0.1)*s_max)-1 for b in bins[:-1]]
        bin_hi[-1] = int(s_max)
        bin_labels = [f"{int(b*100)}%–{int((b+0.1)*100)-1}%" if b<0.9 else "90%–100%" for b in bins[:-1]]

        cls_raw = {}
        for cls in classes:
            mask = class_col == cls
            if sec_label == "合併":
                cls_raw[cls] = df[mask][q_cols].sum(axis=1)
            else:
                p_qs = [q for q in q_cols if paper_map[q]==sec_label]
                cls_raw[cls] = df[mask][p_qs].sum(axis=1) if p_qs else pd.Series(dtype=float)

        n_all = len(raw_series)
        for ri, bi in enumerate(range(len(bin_labels)-1,-1,-1)):
            lo, hi = bin_lo[bi], bin_hi[bi]
            row = cur_row + ri
            ws.row_dimensions[row].height = 17
            bg = _cfill("F7F9FB") if ri%2==1 else NO_FILL
            col = 2
            _sc(ws.cell(row,col), val=bin_labels[bi], fill=_cfill("EBF5FB"),
                font=_fnt(bold=True,size=11,color="1A5276"), aln=_CTR); col+=1
            _sc(ws.cell(row,col), val=f"{lo}–{hi}", fill=_cfill("EBF5FB"),
                font=_fnt(size=11,color="1A5276"), aln=_CTR); col+=1
            for cls in classes:
                s   = cls_raw[cls]
                n_c = len(s)
                cnt   = int(((s>=lo-0.001)&(s<=hi+0.001)).sum()) if n_c else 0
                above = int((s>hi+0.001).sum()) if n_c else 0
                _, bg_hex, _ = palette[cls]
                _sc(ws.cell(row,col), val=f"{cnt} ({cnt/n_c:.0%})" if n_c else "-",
                    fill=bg, font=_fnt(size=12), aln=_CTR); col+=1
                _sc(ws.cell(row,col), val=f"{above} ({above/n_c:.0%})" if n_c else "-",
                    fill=bg, font=_fnt(size=12), aln=_CTR); col+=1
            cnt_a   = int(((raw_series>=lo-0.001)&(raw_series<=hi+0.001)).sum())
            above_a = int((raw_series>hi+0.001).sum())
            _sc(ws.cell(row,col), val=f"{cnt_a} ({cnt_a/n_all:.0%})",
                fill=_cfill("F2F3F4"), font=_fnt(bold=True,size=11,color="3D4E5C"), aln=_CTR); col+=1
            _sc(ws.cell(row,col), val=f"{above_a} ({above_a/n_all:.0%})",
                fill=_cfill("F2F3F4"), font=_fnt(bold=True,size=11,color="3D4E5C"), aln=_CTR)
        cur_row += len(bin_labels)

        ws.row_dimensions[cur_row].height = 20
        _set_merged(ws,cur_row,2,cur_row,3,"FEF9E7",
                    _fnt(bold=True,size=11,color="7D6608"),
                    f"合格人數 / 合格率（≥{round(pass_th)}分）", bdr=THICK_B)
        col = 4
        for cls in classes:
            s=cls_raw[cls]; n_c=len(s)
            pss=int((s>=pass_th).sum()) if n_c else 0
            _set_merged(ws,cur_row,col,cur_row,col+1,"FEF9E7",
                        _fnt(bold=True,size=11,color="7D6608"),
                        f"{pss} ({pss/n_c:.1%})" if n_c else "-", bdr=THICK_B)
            col+=2
        pss_all=int((raw_series>=pass_th).sum())
        _set_merged(ws,cur_row,col,cur_row,col+1,"FEF9E7",
                    _fnt(bold=True,size=11,color="7D6608"),
                    f"{pss_all} ({pss_all/n_all:.1%})", bdr=THICK_B)
        cur_row+=1

        for stat_lbl, stat_fn, num_fmt in [
            ("平均分", lambda s: round(s.mean(),1), "0.0"),
            ("S.D.",   lambda s: round(s.std(ddof=1),1), "0.0"),
            ("最高分", lambda s: int(s.max()), "0"),
            ("最低分", lambda s: int(s.min()), "0"),
        ]:
            ws.row_dimensions[cur_row].height = 17
            _set_merged(ws,cur_row,2,cur_row,3,"EBF5FB",
                        _fnt(bold=True,size=11,color="1A5276"),stat_lbl,bdr=STAT_B)
            col=4
            for cls in classes:
                s=cls_raw[cls]; _,bg_hex,dark_hex=palette[cls]
                _set_merged(ws,cur_row,col,cur_row,col+1,bg_hex,
                            _fnt(bold=True,size=11,color=dark_hex),
                            stat_fn(s) if len(s)>0 else "-",fmt=num_fmt,bdr=STAT_B)
                col+=2
            _set_merged(ws,cur_row,col,cur_row,col+1,"EAECEE",
                        _fnt(bold=True,size=11,color="3D4E5C"),
                        stat_fn(raw_series),fmt=num_fmt,bdr=STAT_B)
            cur_row+=1
        cur_row+=2

    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    for ci in range(len(classes)):
        ws.column_dimensions[get_column_letter(4+ci*2)].width   = 14
        ws.column_dimensions[get_column_letter(5+ci*2)].width   = 16
    ws.column_dimensions[get_column_letter(4+len(classes)*2)].width = 14
    ws.column_dimensions[get_column_letter(5+len(classes)*2)].width = 16
    ws.freeze_panes = "B4"

# ══════════════════════════════════════════════════════════════
# 班際熱力圖工作表
# ══════════════════════════════════════════════════════════════
def _make_heatmap_sheet(wb, class_stats, grade_stats, classes, q_cols, paper_map, palette):
    """班際熱力圖 - 對齊範本v12：B=試卷, C=題號, D..=各班, 最後欄=全級"""
    ws = wb.create_sheet("班際熱力圖")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    n_cls  = len(classes)
    # 欄結構：A(margin) B(試卷) C(題號) D..D+n-1(各班) D+n(全級)
    END_COL = 3 + n_cls + 1  # = 4+n_cls
    end_c   = get_column_letter(END_COL)

    # ── 行2：主標題 ──
    ws.row_dimensions[2].height = 30
    ws.merge_cells(f"B2:{end_c}2")
    _sc(ws["B2"], val="班際比較熱力圖　·　各題平均得分率",
        fill=_cfill("1F3864"), font=_hdr_fnt(bold=True, size=14, color="FFFFFF"), aln=_CTR)

    # ── 行3：副標題 ──
    ws.row_dimensions[3].height = 15
    ws.merge_cells(f"B3:{end_c}3")
    _sc(ws["B3"], val="顏色由深紅（低）→ 黃（中）→ 深綠（高）。全級欄顯示各題整體水平。",
        font=Font(size=9, color="595959", name="新細明體"), aln=_LFT)

    # ── 行4：標題列 ──
    ws.row_dimensions[4].height = 20
    _sc(ws.cell(4,2), val="試卷", fill=_cfill("D6E4F7"),
        font=_hdr_fnt(bold=True, size=11, color="1F3864"), aln=_CTR)
    _sc(ws.cell(4,3), val="題號", fill=_cfill("D6E4F7"),
        font=_hdr_fnt(bold=True, size=11, color="1F3864"), aln=_CTR)

    # 各班標題色（固定配色，對齊範本）
    for i, cls in enumerate(classes):
        hf, _ = HEATMAP_HEADER_FILLS[i % len(HEATMAP_HEADER_FILLS)]
        _sc(ws.cell(4, 4+i), val=cls, fill=_cfill(hf),
            font=_hdr_fnt(bold=True, size=11, color="FFFFFF"), aln=_CTR)
    # 全級標題
    _sc(ws.cell(4, 4+n_cls), val="全級",
        fill=_cfill("FFF2CC"), font=_hdr_fnt(bold=True, size=11, color="7F4F00"), aln=_CTR)

    # ── 數據行 ──
    prev_paper = None
    row = 5
    for ri, q in enumerate(q_cols):
        paper = paper_map.get(q, "P1")

        # 試卷分組行（首次出現合拼試卷格）
        if paper != prev_paper:
            ws.row_dimensions[row].height = 14
            _sc(ws.cell(row, 2), val=paper, fill=_cfill("2E75B6"),
                font=_hdr_fnt(bold=True, size=11, color="FFFFFF"), aln=_CTR)
            prev_paper = paper

        ws.row_dimensions[row].height = 17
        row_fill = _cfill("F2F7FC") if ri % 2 == 0 else PatternFill(fill_type=None)

        # 試卷欄（已在上面寫，數據行空著即可）
        ws.cell(row, 2).fill = row_fill if row_fill.fill_type else PatternFill(fill_type=None)

        # 題號欄
        _sc(ws.cell(row, 3), val=q, fill=row_fill,
            font=_fnt(bold=True, size=11, color="000000"), aln=_CTR)

        # 各班熱力色
        pcts = [class_stats[cls][q]["mean_pct"] for cls in classes]
        grade_pct = grade_stats[q]["mean_pct"]
        all_pcts = pcts + [grade_pct]
        mn2, mx2 = min(all_pcts), max(all_pcts)

        for i, cls in enumerate(classes):
            pct = pcts[i]
            t   = (pct - mn2)/(mx2 - mn2) if mx2 > mn2 else 0.5
            # 深紅(0) → 黃(0.5) → 深綠(1)
            if t < 0.5:
                t2 = t * 2
                r2 = int(192 + (1-t2)*63); g2 = int(t2*192); b2 = int(t2*50)
            else:
                t2 = (t-0.5)*2
                r2 = int((1-t2)*192); g2 = int(192 - t2*30); b2 = int(50 + t2*20)
            r2 = max(0,min(255,r2)); g2=max(0,min(255,g2)); b2=max(0,min(255,b2))
            hex_bg   = f"{r2:02X}{g2:02X}{b2:02X}"
            txt_col  = "FFFFFF" if t < 0.3 or t > 0.75 else "1A1A1A"
            _sc(ws.cell(row, 4+i), val=pct, fill=_cfill(hex_bg),
                font=_fnt(size=11, color=txt_col), aln=_CTR, fmt="0.0%")

        # 全級欄
        _sc(ws.cell(row, 4+n_cls), val=grade_pct,
            fill=_cfill("FFF2CC"), font=_fnt(bold=True, size=11, color="7F4F00"),
            aln=_CTR, fmt="0.0%")
        row += 1

    # ── 欄寬（對齊範本v12）──
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 8
    for i in range(n_cls + 1):
        ws.column_dimensions[get_column_letter(4+i)].width = 10
    ws.freeze_panes = "D5"

# ══════════════════════════════════════════════════════════════
# 主要公開函數
# ══════════════════════════════════════════════════════════════
def generate_class_analysis_excel(
    df: pd.DataFrame,
    max_scores: pd.Series,
    paper_map: dict,
    class_info: pd.DataFrame,
    exam_info: dict,
) -> bytes:
    # 用位置對應（避免重複姓名導致班別誤判）
    class_col = pd.Series(
        class_info["班別"].astype(str).values,
        index=df.index
    )

    q_cols  = list(max_scores.index)
    classes = sorted(class_col.unique().tolist())
    palette = _build_palette(classes)

    class_stats, grade_stats, classes, paper_totals, combined_totals = _compute_stats(
        df, q_cols, max_scores, paper_map, class_col
    )

    wb = Workbook()
    wb.remove(wb.active)

    for cls in classes:
        _make_class_sheet(wb, cls, q_cols, max_scores, paper_map,
                          class_stats, grade_stats, df, class_col, palette)
    _make_distribution_sheet(wb, df, max_scores, paper_map, class_col, classes, palette)
    _make_heatmap_sheet(wb, class_stats, grade_stats, classes, q_cols, paper_map, palette)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def get_class_summary_df(
    df: pd.DataFrame,
    max_scores: pd.Series,
    paper_map: dict,
    class_info: pd.DataFrame,
) -> pd.DataFrame:
    # 用位置對應（避免重複姓名導致班別誤判）
    class_col = pd.Series(
        class_info["班別"].astype(str).values,
        index=df.index
    )

    q_cols   = list(max_scores.index)
    total_mx = float(max_scores.sum())
    pass_th  = total_mx * 0.4
    rows = []
    for cls in sorted(class_col.unique()):
        sub = df[class_col==cls][q_cols]
        raw = sub.sum(axis=1)
        rows.append({
            "班別":      cls,
            "人數":      len(sub),
            "合併平均分": round(raw.mean(),1),
            "合併平均%":  f"{raw.mean()/total_mx*100:.1f}%",
            "合格率":     f"{(raw>=pass_th).mean()*100:.1f}%",
        })
    all_raw = df[q_cols].sum(axis=1)
    rows.append({
        "班別":      "全　級",
        "人數":      len(df),
        "合併平均分": round(all_raw.mean(),1),
        "合併平均%":  f"{all_raw.mean()/total_mx*100:.1f}%",
        "合格率":     f"{(all_raw>=pass_th).mean()*100:.1f}%",
    })
    return pd.DataFrame(rows)
