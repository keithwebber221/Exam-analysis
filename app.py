"""
DSE 試卷分析系統 - 網頁版
Streamlit app (app.py)
"""
import streamlit as st
import pandas as pd
import numpy as np
import io, os, sys, zipfile, tempfile, re

# ── 載入分析模組（與 app.py 同資料夾）──
sys.path.insert(0, os.path.dirname(__file__))
import exam_item_analysis as ea
import individual_report  as ir

# ══════════════════════════════════════════════════════════════
# 頁面設定
# ══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="DSE 試卷分析系統",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 全域 CSS ──
st.markdown("""
<style>
    .main-header{font-size:1.8rem;font-weight:700;color:#1F4788;margin-bottom:0.2rem;}
    .sub-header{font-size:1rem;color:#555;margin-bottom:1.5rem;}
    .step-badge{background:#1F4788;color:white;border-radius:50%;
                width:28px;height:28px;display:inline-flex;
                align-items:center;justify-content:center;
                font-weight:bold;margin-right:8px;}
    .success-box{background:#d4edda;border-left:4px solid #28a745;
                 padding:0.75rem 1rem;border-radius:4px;margin:0.5rem 0;}
    .warn-box{background:#fff3cd;border-left:4px solid #ffc107;
              padding:0.75rem 1rem;border-radius:4px;margin:0.5rem 0;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# 工具函數：load_data 的 BytesIO 版本
# ══════════════════════════════════════════════════════════════
def load_data_from_bytes(file_bytes: bytes):
    """從 bytes 讀取 scores.xlsx（適用 Streamlit UploadedFile）"""
    buf = io.BytesIO(file_bytes)
    df_raw    = pd.read_excel(buf, header=None)
    col_names = df_raw.iloc[2].tolist()

    row3_vals = df_raw.iloc[3].tolist()
    paper_labels = {"P1","P2","P3","P4"}
    row3_str  = [str(v).strip() for v in row3_vals if pd.notna(v) and str(v).strip()!=""]
    has_paper_row = any(v in paper_labels for v in row3_str)

    if has_paper_row:
        paper_row   = row3_vals
        max_row     = df_raw.iloc[4].tolist()
        student_raw = df_raw.iloc[5:].copy()
    else:
        paper_row   = None
        max_row     = row3_vals
        student_raw = df_raw.iloc[4:].copy()

    student_raw.columns = col_names
    student_raw = student_raw[
        student_raw["中文姓名"].notna() &
        ~student_raw["中文姓名"].astype(str).str.contains("說明|輸入", na=False)
    ]

    absent_set = set()
    absent_col_candidates = ["缺席","Absent","absent","缺考"]
    absent_col = next((c for c in col_names if str(c) in absent_col_candidates), None)
    if absent_col:
        for _, row in student_raw.iterrows():
            val = row.get(absent_col,"")
            if pd.notna(val) and str(val).strip() not in ["","0","nan"]:
                absent_set.add(str(row["中文姓名"]).strip())

    info_cols     = ["班別","班號","英文姓名","中文姓名"] + absent_col_candidates
    question_cols = [c for c in col_names
                     if isinstance(c,str) and c.strip()!="" and c not in info_cols]
    max_scores = pd.Series(max_row, index=col_names)[question_cols].astype(float)
    score_data = student_raw[question_cols].astype(float)
    score_data.index = student_raw["中文姓名"].values
    score_data.index.name = "姓名"
    score_data.fillna(0, inplace=True)

    # class_info
    ci = pd.DataFrame()
    ci["班別"]    = student_raw.iloc[:,0].values
    ci["班號"]    = student_raw.iloc[:,1].values
    ci["中文姓名"] = student_raw["中文姓名"].values
    ci = ci[ci["中文姓名"].notna()].reset_index(drop=True)

    paper_map = {}
    if has_paper_row:
        paper_series = pd.Series(paper_row, index=col_names)
        for q in question_cols:
            val = str(paper_series.get(q,"P1")).strip()
            paper_map[q] = val if val in paper_labels else "P1"
    else:
        paper_map = {q:"P1" for q in question_cols}

    return score_data, max_scores, absent_set, paper_map, ci


def export_excel_bytes(item_df, group_df, student_df, stats_df, exam_title):
    """匯出 Excel 到 BytesIO"""
    buf = io.BytesIO()
    from openpyxl.styles import (PatternFill, Font, Border, Side,
                                  Alignment, GradientFill)
    from openpyxl.utils import get_column_letter

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"))
    header_fill = PatternFill("solid", fgColor="1F4788")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    color_fills = {
        "🟢 容易": PatternFill("solid", fgColor="C6EFCE"),
        "🟡 適中": PatternFill("solid", fgColor="FFEB9C"),
        "🔴 困難": PatternFill("solid", fgColor="FFC7CE"),
        "⭐ 優良": PatternFill("solid", fgColor="C6EFCE"),
        "✅ 良好": PatternFill("solid", fgColor="DDEBF7"),
        "⚠️ 尚可": PatternFill("solid", fgColor="FFEB9C"),
        "❌ 不佳": PatternFill("solid", fgColor="FFC7CE"),
    }

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        item_df.to_excel(writer, sheet_name="1_試題分析", index=False)
        group_df.to_excel(writer, sheet_name="2_大題分析", index=False)

        student_df_sorted = student_df.copy()
        rank_col = pd.to_numeric(student_df_sorted["排名"], errors="coerce")
        student_df_sorted = student_df_sorted.iloc[rank_col.argsort(kind="stable")]
        absent_mask = student_df_sorted["出席狀態"] == "缺席"
        student_df_sorted = pd.concat([student_df_sorted[~absent_mask],
                                       student_df_sorted[absent_mask]])
        student_df_sorted.to_excel(writer, sheet_name="3_學生成績", index=True)
        stats_df.to_excel(writer, sheet_name="4_全班統計", index=False)

        wb = writer.book
        # 格式化試題分析
        ws = wb["1_試題分析"]
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center")
                for keyword, fill in color_fills.items():
                    if str(cell.value) == keyword:
                        cell.fill = fill
        for col in ws.columns:
            ws.column_dimensions[get_column_letter(col[0].column)].width = 14

    buf.seek(0)
    return buf.read()


def generate_reports_zip(df, max_scores, item_df, exam_info, class_info,
                          pass_rate, absent_set):
    """生成個人報告 ZIP（BytesIO）"""
    buf = io.BytesIO()
    total_scores = df.sum(axis=1)
    total_max    = int(max_scores.sum())
    class_avg    = item_df["平均分"].sum()

    class_info_dict = {}
    if class_info is not None and len(class_info):
        class_info_dict = dict(zip(
            class_info["中文姓名"],
            zip(class_info["班別"].astype(str), class_info["班號"].astype(str))
        ))

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for student_name in df.index:
            student_score = total_scores[student_name]
            class_code, class_num = class_info_dict.get(student_name, ("", "00"))
            try:
                class_num_int = int(float(class_num))
            except:
                class_num_int = 0

            if student_name in absent_set:
                from docx import Document
                from docx.shared import Pt, Inches
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from docx.oxml.ns import qn
                from docx.oxml import OxmlElement
                from docx.shared import RGBColor
                doc = Document()
                sec = doc.sections[0]
                sec.top_margin = sec.bottom_margin = Inches(1.0)
                sec.left_margin = sec.right_margin = Inches(1.2)
                for _ in range(3): doc.add_paragraph()
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run("本次缺席")
                r.font.size = Pt(36); r.font.bold = True
                r.font.color.rgb = RGBColor(0xC0,0x39,0x2B)
            else:
                doc = ir.create_personal_report_v2_4(
                    student_name, student_score, total_max,
                    df.loc[student_name], max_scores, item_df,
                    exam_info, class_avg, total_max,
                    class_info, pass_rate
                )

            fname = f"{class_code}{class_num_int:02d}{student_name}_個人報告.docx"
            doc_buf = io.BytesIO()
            doc.save(doc_buf)
            doc_buf.seek(0)
            zf.writestr(fname, doc_buf.read())

    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
# 追蹤分析工具函數
# ══════════════════════════════════════════════════════════════
def export_tracking_excel_bytes(pct_matrix, rank_matrix, student_info,
                                 class_stats, exam_labels, pass_rate, subject):
    """追蹤報告 Excel → BytesIO"""
    buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = os.path.join(tmpdir, "tracking.xlsx")
        import performance_tracker as pt
        pt.export_tracking_excel(
            pct_matrix, rank_matrix, student_info, class_stats,
            exam_labels, tmp_path, pass_rate, subject
        )
        with open(tmp_path, "rb") as f:
            buf.write(f.read())
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
# 側邊欄導航
# ══════════════════════════════════════════════════════════════
st.sidebar.markdown("## 📊 DSE 試卷分析系統")
st.sidebar.markdown("---")
page = st.sidebar.radio(
    "功能選擇",
    ["📋 試卷分析", "📈 成績追蹤"],
    label_visibility="collapsed"
)
st.sidebar.markdown("---")



# ══════════════════════════════════════════════════════════════
# 頁面一：試卷分析
# ══════════════════════════════════════════════════════════════
if page == "📋 試卷分析":
    st.markdown('''<div class="main-header">📋 試卷分析</div>
<div class="sub-header">上載成績表，即時生成試題分析、學生報告及圖表</div>''',
                unsafe_allow_html=True)

    # ── Step 1: 考試資訊 ──
    st.markdown("### ① 考試資訊")
    col1, col2, col3 = st.columns(3)
    with col1:
        year_input = st.text_input("年度（如 2526）", value="2526", max_chars=4)
        exam_type  = st.selectbox("考試類別", ["上學期測驗","上學期考試","下學期測驗","下學期考試"])
    with col2:
        form_label  = st.selectbox("年級", ["F1","F2","F3","F4","F5","F6"])
        subject_label = st.text_input("科目（如 BAFS）", value="BAFS")
    with col3:
        pass_rate_pct = st.selectbox("及格線", ["40%（高中）","50%（初中）"])
        pass_rate = 0.4 if "40" in pass_rate_pct else 0.5

    # 試卷設定
    st.markdown("### ② 試卷結構")
    num_papers = st.selectbox("試卷數目", [1,2,3,4],
                               format_func=lambda x: f"{x} 份試卷")
    paper_weights = {"P1": 1.0}
    if num_papers > 1:
        st.info("各試卷比例合計必須為 100%")
        cols = st.columns(num_papers)
        tmp_weights = {}
        for i, col in enumerate(cols, 1):
            with col:
                w = col.number_input(f"卷{i}（P{i}）比例 %",
                                      min_value=1, max_value=99,
                                      value=100//num_papers, step=1,
                                      key=f"pw_{i}")
                tmp_weights[f"P{i}"] = w
        total_w = sum(tmp_weights.values())
        if total_w != 100:
            st.warning(f"⚠️ 目前合計：{total_w}%，必須等於 100%")
        else:
            paper_weights = {k: v/100 for k, v in tmp_weights.items()}
            st.success(f"✅ 各卷比例設定正確")

    # 組合 exam_info
    EXAM_CODES = {"上學期測驗":"T1T","上學期考試":"T1E","下學期測驗":"T2T","下學期考試":"T2E"}
    year_label  = f"20{year_input[:2]}-20{year_input[2:]}" if len(year_input)==4 else year_input
    exam_title  = f"{year_label} {exam_type}｜{form_label} {subject_label}"
    file_prefix = f"{year_input}_{EXAM_CODES.get(exam_type,'EXAM')}_{form_label}_{subject_label}"
    exam_info = {
        "year_label": year_label, "exam_type_label": exam_type,
        "exam_type_code": EXAM_CODES.get(exam_type,"EXAM"),
        "subject_label": subject_label, "form_label": form_label,
        "exam_title": exam_title, "file_prefix": file_prefix,
        "pass_rate": pass_rate, "paper_weights": paper_weights,
        "num_papers": num_papers,
    }

    # ── Step 3: 上載檔案 ──
    st.markdown("### ③ 上載成績表")
    uploaded = st.file_uploader("選擇 scores.xlsx", type=["xlsx"],
                                  help="第3行為題號，第4行為【試卷】（多試卷時），之後為【滿分】，再下為學生成績")

    if uploaded:
        raw_bytes = uploaded.read()
        try:
            df, max_scores, absent_set, paper_map, class_info = load_data_from_bytes(raw_bytes)
            papers_in_excel = sorted(set(paper_map.values()))
            if len(papers_in_excel) > 1 and num_papers == 1:
                st.warning(f"⚠️ 偵測到試卷行 {papers_in_excel}，但設定為單試卷，請調整上方試卷數目")

            st.markdown(f'''<div class="success-box">
            ✅ 成功載入：<b>{len(df)} 名學生</b>，<b>{len(df.columns)} 道題目</b>
            {"｜缺席：" + "、".join(sorted(absent_set)) if absent_set else ""}
            </div>''', unsafe_allow_html=True)

            # 預覽
            with st.expander("📋 預覽學生名單", expanded=False):
                st.dataframe(class_info, use_container_width=True)

            # ── 分析按鈕 ──
            st.markdown("### ④ 開始分析")
            if st.button("🚀 開始分析", type="primary", use_container_width=True):
                with st.spinner("分析中，請稍候..."):

                    # 加權計算
                    weighted_scores, paper_pct, _ = ea.calc_weighted_scores(
                        df, max_scores, paper_weights, paper_map)

                    item_df    = ea.item_analysis(df.copy(), max_scores, absent_set)
                    student_df, stats_df = ea.student_summary(
                        df.copy(), max_scores, pass_rate, absent_set,
                        paper_weights=paper_weights, paper_pct=paper_pct,
                        weighted_scores=weighted_scores, num_papers=num_papers)
                    group_df   = ea.question_group_analysis(df.copy(), max_scores, item_df)

                st.success("✅ 分析完成！")

                # ── 結果顯示 ──
                tab1, tab2, tab3, tab4 = st.tabs(
                    ["📊 試題分析","📋 大題分析","👨‍🎓 學生成績","📈 全班統計"])

                with tab1:
                    st.dataframe(item_df, use_container_width=True, height=400)
                with tab2:
                    st.dataframe(group_df, use_container_width=True)
                with tab3:
                    st.dataframe(student_df, use_container_width=True, height=400)
                with tab4:
                    st.dataframe(stats_df, use_container_width=True)

                # ── 互動圖表 ──
                st.markdown("### 📊 互動圖表")
                import plotly.express as px
                import plotly.graph_objects as go

                c1, c2 = st.columns(2)
                with c1:
                    item_plot = item_df.copy()
                    item_plot["得分率 %"] = (item_plot["平均分"] / item_plot["滿分"] * 100).round(1)
                    color_map = {"🟢 容易":"#2ecc71","🟡 適中":"#f39c12","🔴 困難":"#e74c3c"}
                    fig1 = px.scatter(item_plot, x="難度指數 P", y="鑑別度 D",
                                      text="題號", color="難度評級",
                                      color_discrete_map=color_map,
                                      title="試題難度-鑑別度分佈")
                    fig1.add_hline(y=0.3, line_dash="dash", line_color="gray")
                    fig1.add_vline(x=0.25, line_dash="dash", line_color="gray")
                    fig1.add_vline(x=0.75, line_dash="dash", line_color="gray")
                    st.plotly_chart(fig1, use_container_width=True)

                with c2:
                    item_sorted = item_plot.sort_values("得分率 %")
                    fig2 = px.bar(item_sorted, x="得分率 %", y="題號",
                                  orientation="h", color="難度評級",
                                  color_discrete_map=color_map,
                                  title="各題得分率排行", text="得分率 %")
                    fig2.update_traces(texttemplate="%{text}%", textposition="outside")
                    fig2.add_vline(x=50, line_dash="dash", line_color="gray")
                    st.plotly_chart(fig2, use_container_width=True)

                # 總分分佈
                scores_num = pd.to_numeric(student_df.get(
                    "總分(加權)" if "總分(加權)" in student_df.columns else "總分",
                    pd.Series()), errors="coerce").dropna()
                if len(scores_num) >= 2:
                    fig3 = px.histogram(scores_num, nbins=10, title="全班總分分佈",
                                        labels={"value":"總分","count":"人數"})
                    fig3.add_vline(x=scores_num.mean(), line_dash="dash",
                                   annotation_text=f"平均 {scores_num.mean():.1f}")
                    st.plotly_chart(fig3, use_container_width=True)

                # ── 下載區 ──
                st.markdown("### ⬇️ 下載報告")
                dl1, dl2, dl3 = st.columns(3)

                with dl1:
                    excel_bytes = export_excel_bytes(
                        item_df, group_df, student_df, stats_df, exam_title)
                    st.download_button(
                        "📥 下載 Excel 分析報告",
                        data=excel_bytes,
                        file_name=f"{file_prefix}_analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)

                with dl2:
                    with st.spinner("生成個人報告中..."):
                        zip_bytes = generate_reports_zip(
                            df, max_scores, item_df, exam_info,
                            class_info, pass_rate, absent_set)
                    st.download_button(
                        "📥 下載個人報告 ZIP",
                        data=zip_bytes,
                        file_name=f"{file_prefix}_個人報告.zip",
                        mime="application/zip",
                        use_container_width=True)

                with dl3:
                    st.download_button(
                        "📥 下載原始成績表",
                        data=raw_bytes,
                        file_name=f"{file_prefix}_scores.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)


        except Exception as e:
            st.error(f"❌ 載入失敗：{e}")
            import traceback
            st.code(traceback.format_exc())


# ══════════════════════════════════════════════════════════════
# 頁面二：成績追蹤
# ══════════════════════════════════════════════════════════════
elif page == "📈 成績追蹤":
    st.markdown('''<div class="main-header">📈 成績追蹤</div>
<div class="sub-header">上載多次考試的分析報告，生成跨試成績追蹤</div>''',
                unsafe_allow_html=True)

    import performance_tracker as pt

    # ── Step 1: 設定 ──
    st.markdown("### ① 追蹤設定")
    col1, col2 = st.columns(2)
    with col1:
        track_subject  = st.text_input("科目名稱", value="BAFS")
        track_pass_pct = st.selectbox("及格線", ["40%（高中）","50%（初中）"], key="tp")
        track_pass_rate = 0.4 if "40" in track_pass_pct else 0.5
    with col2:
        st.markdown("**分析檔案命名規則**")
        st.code("2526_T1E_F5_BAFS_analysis.xlsx\n2526_T2E_F5_BAFS_analysis.xlsx", language=None)
        st.caption("格式：年度_類別_年級_科目_analysis.xlsx")

    # ── Step 2: 上載 ──
    st.markdown("### ② 上載分析報告（可多選）")
    track_files = st.file_uploader(
        "選擇 *_analysis.xlsx 檔案（可多選）",
        type=["xlsx"], accept_multiple_files=True,
        help="即之前試卷分析下載的 Excel 報告")

    if track_files:
        st.markdown(f"已上載 **{len(track_files)}** 個檔案：")
        exam_files_data = []
        for f in track_files:
            year, etype, form, subject = pt.parse_filename(f.name)
            tag = " ".join(filter(None,[form,subject]))
            st.caption(f"· {f.name}  [{tag or '舊格式'}]  {year} {etype}")
            if year and etype:
                exam_files_data.append({
                    "file_obj": f, "file": f.name,
                    "year": year, "type": etype,
                    "form": form, "subject": subject,
                })
            else:
                st.warning(f"⚠️ {f.name} 無法解析檔名，請確認命名格式")

        # 過濾選項
        forms_found    = sorted(set(ef["form"]    for ef in exam_files_data if ef["form"]))
        subjects_found = sorted(set(ef["subject"] for ef in exam_files_data if ef["subject"]))

        col1, col2 = st.columns(2)
        with col1:
            form_filter = st.selectbox(
                "年級過濾", ["（全部）"] + forms_found) if len(forms_found)>1 else (
                forms_found[0] if forms_found else "（全部）")
        with col2:
            subj_filter = st.selectbox(
                "科目過濾", ["（全部）"] + subjects_found) if len(subjects_found)>1 else (
                subjects_found[0] if subjects_found else "（全部）")

        filtered = [ef for ef in exam_files_data
                    if (form_filter in ["（全部）",""] or ef["form"]==form_filter)
                    and (subj_filter in ["（全部）",""] or ef["subject"]==subj_filter)]
        st.info(f"過濾後：{len(filtered)} 個檔案納入分析")

        # ── Step 3: 班別班號（選填）──
        st.markdown("### ③ 上載最新 scores.xlsx（選填，補充班別班號）")
        ci_file = st.file_uploader("選擇 scores.xlsx（選填）", type=["xlsx"], key="ci_upload")
        class_info_df = None
        if ci_file:
            try:
                raw = pd.read_excel(io.BytesIO(ci_file.read()), header=None)
                ci = pd.DataFrame()
                ci["班別"]    = raw.iloc[4:,0].values
                ci["班號"]    = raw.iloc[4:,1].values
                ci["中文姓名"] = raw.iloc[4:,3].values
                ci = ci[ci["中文姓名"].notna() &
                        ~ci["中文姓名"].astype(str).str.contains("說明|輸入",na=False)]
                class_info_df = ci.reset_index(drop=True)
                st.success(f"✅ 讀取 {len(ci)} 位學生班別資訊")
            except Exception as e:
                st.warning(f"⚠️ 讀取 scores.xlsx 失敗：{e}")

        # ── Step 4: 生成 ──
        if st.button("🚀 生成追蹤報告", type="primary",
                     use_container_width=True, disabled=len(filtered)==0):
            with st.spinner("建立成績矩陣中..."):
                # 將上載檔案寫入暫存目錄
                with tempfile.TemporaryDirectory() as tmpdir:
                    ef_list = []
                    for ef in filtered:
                        tmp_path = os.path.join(tmpdir, ef["file"])
                        ef["file_obj"].seek(0)
                        with open(tmp_path, "wb") as out:
                            out.write(ef["file_obj"].read())
                        ef_list.append({
                            "file": tmp_path, "year": ef["year"],
                            "type": ef["type"], "form": ef["form"],
                            "subject": ef["subject"],
                        })

                    pct_matrix, rank_matrix, exam_labels, student_info = (
                        pt.build_tracking_matrix(ef_list, class_info_df))

                    if pct_matrix is None:
                        st.error("❌ 無法建立成績矩陣，請確認檔案格式")
                    else:
                        class_stats = pt.calc_class_stats(pct_matrix, track_pass_rate)
                        st.success(f"✅ 矩陣：{pct_matrix.shape[0]} 位學生 × {pct_matrix.shape[1]} 次考試")

                        # 顯示結果
                        tab1, tab2 = st.tabs(["📊 全班趨勢","👨‍🎓 學生成績矩陣"])
                        with tab1:
                            import plotly.graph_objects as go
                            fig = go.Figure()
                            for col in class_stats.columns[1:]:
                                vals = pd.to_numeric(class_stats[col], errors="coerce")
                                if not vals.isna().all():
                                    fig.add_trace(go.Scatter(
                                        x=class_stats["考試"],
                                        y=vals, mode="lines+markers",
                                        name=col))
                            fig.update_layout(title="全班成績趨勢",
                                              xaxis_title="考試",
                                              yaxis_title="分數")
                            st.plotly_chart(fig, use_container_width=True)
                        with tab2:
                            st.dataframe(pct_matrix.round(1), use_container_width=True, height=400)

                        # 下載 Excel
                        st.markdown("### ⬇️ 下載追蹤報告")
                        try:
                            excel_bytes = export_tracking_excel_bytes(
                                pct_matrix, rank_matrix, student_info,
                                class_stats, exam_labels,
                                track_pass_rate, track_subject)
                            years = sorted(set(ef["year"] for ef in filtered))
                            out_name = f"{'_'.join(years)}_{track_subject}_成績追蹤.xlsx"
                            st.download_button(
                                "📥 下載成績追蹤 Excel",
                                data=excel_bytes,
                                file_name=out_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True)
                        except Exception as e:
                            st.error(f"❌ 生成 Excel 失敗：{e}")
                            import traceback; st.code(traceback.format_exc())
