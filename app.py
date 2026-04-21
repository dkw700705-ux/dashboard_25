import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np

st.set_page_config(
    page_title="25학번 대학 적응 조사 대시보드",
    password = st.sidebar.text_input("비밀번호", type="password")
if password != "4872":
    st.warning("비밀번호를 입력해주세요.")
    st.stop()
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    [data-testid="stSidebar"] { background-color: #1e2a3a; }
    [data-testid="stSidebar"] * { color: #e0e8f0 !important; }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px; padding: 18px 22px; color: white;
        text-align: center; margin-bottom: 8px;
    }
    .metric-card .val { font-size: 2rem; font-weight: 700; }
    .metric-card .lbl { font-size: 0.85rem; opacity: 0.85; margin-top: 2px; }
    .type-card {
        border-radius: 10px; padding: 14px 18px;
        text-align: center; color: white; margin-bottom: 6px;
    }
    .section-title {
        font-size: 1.15rem; font-weight: 700; color: #2c3e50;
        border-left: 4px solid #667eea; padding-left: 10px;
        margin: 18px 0 10px;
    }
    .stTabs [data-baseweb="tab"] { font-size: 0.9rem; }
</style>
""", unsafe_allow_html=True)


@st.cache_data
def load_data():
    path = "대시보드_구축_데이터.xlsx"
    xl = pd.ExcelFile(path)

    # 시트1: 유형분류
    raw1 = pd.read_excel(xl, sheet_name="유형분류", header=None)
    type_rows = []
    for i in range(2, len(raw1)):
        r = raw1.iloc[i]
        학과 = r[0]
        if pd.isna(학과):
            continue
        type_rows.append({
            "학과": 학과,
            "저적응위기형_빈도": r[1], "저적응위기형_비율": r[2],
            "적응취약형_빈도":  r[3], "적응취약형_비율":  r[4],
            "적응안정형_빈도":  r[5], "적응안정형_비율":  r[6],
            "고적응성취형_빈도": r[7], "고적응성취형_비율": r[8],
            "총계": r[9],
        })
    df_type = pd.DataFrame(type_rows)

    # 시트2: 영역결과
    raw2 = pd.read_excel(xl, sheet_name="영역결과", header=None)
    area_rows = []
    for i in range(2, len(raw2)):
        r = raw2.iloc[i]
        학과 = r[0]
        if pd.isna(학과):
            continue
        area_rows.append({
            "학과": 학과, "빈도": r[1],
            "학업적응_평균": r[2],  "학업적응_SD": r[3],
            "사회적응_평균": r[4],  "사회적응_SD": r[5],
            "대학적응_평균": r[6],  "대학적응_SD": r[7],
            "경제적응_평균": r[8],  "경제적응_SD": r[9],
            "정서적응_평균": r[10], "정서적응_SD": r[11],
            "진로적응_평균": r[12], "진로적응_SD": r[13],
            "중도탈락의도_평균": r[14], "중도탈락의도_SD": r[15],
        })
    df_area = pd.DataFrame(area_rows)

    # 시트3: 문항결과
    raw3 = pd.read_excel(xl, sheet_name="문항결과", header=None)

    section_cols = {
        "학업적응": 26, "사회적응": 58, "대학적응": 78,
        "경제적응": 90, "정서적응": 114, "진로적응": 140, "중도탈락의도": 152
    }

    area_order = ["학업적응", "사회적응", "대학적응", "경제적응", "정서적응", "진로적응", "중도탈락의도"]
    item_map = {}
    for idx, area in enumerate(area_order):
        start = 2
        if idx > 0:
            prev_area = area_order[idx - 1]
            start = section_cols[prev_area] + 2
        end = section_cols[area]
        items = []
        for c in range(start, end, 2):
            item_name = raw3.iloc[0, c]
            if pd.notna(item_name):
                items.append((str(item_name), c, c + 1))
        item_map[area] = items

    item_rows = []
    for i in range(2, len(raw3)):
        r = raw3.iloc[i]
        학과 = r[0]
        if pd.isna(학과):
            continue
        row = {"학과": 학과, "빈도": r[1]}
        for area, items in item_map.items():
            for item_name, mc, sc in items:
                row[f"{area}||{item_name}||평균"] = r[mc]
                row[f"{area}||{item_name}||SD"] = r[sc]
        item_rows.append(row)
    df_item = pd.DataFrame(item_rows)

    return df_type, df_area, df_item, item_map


df_type, df_area, df_item, item_map = load_data()

학과_목록 = [d for d in df_type["학과"].tolist() if d != "총계"]
전체_key = "총계"

AREA_LABELS = ["학업적응", "사회적응", "대학적응", "경제적응", "정서적응", "진로적응", "중도탈락의도"]
TYPE_COLORS = {
    "저적응위기형": "#e74c3c",
    "적응취약형":   "#e67e22",
    "적응안정형":   "#3498db",
    "고적응성취형": "#2ecc71",
}

# 사이드바
with st.sidebar:
    st.markdown("## 🎓 대학 적응 대시보드")
    st.markdown("---")
    st.markdown("### 학과 선택")
    selected = st.selectbox(
        "학과를 선택하세요",
        학과_목록,
        index=0,
        label_visibility="collapsed"
    )
    st.markdown("---")
    row_s = df_area[df_area["학과"] == selected]
    if not row_s.empty:
        n = int(row_s["빈도"].values[0])
        st.markdown(f"**응답자 수**: {n}명")
    st.markdown("---")
    st.markdown("<small style='opacity:0.5'>대학 적응 유형 분류 및 영역별 분석 대시보드</small>", unsafe_allow_html=True)

st.markdown(f"# {selected}")
st.markdown("---")


def get_type_row(dept):
    r = df_type[df_type["학과"] == dept]
    return r.iloc[0] if not r.empty else None

def get_area_row(dept):
    r = df_area[df_area["학과"] == dept]
    return r.iloc[0] if not r.empty else None

def get_item_row(dept):
    r = df_item[df_item["학과"] == dept]
    return r.iloc[0] if not r.empty else None


tr_s = get_type_row(selected)
ar_s = get_area_row(selected)
ir_s = get_item_row(selected)
tr_all = get_type_row(전체_key)
ar_all = get_area_row(전체_key)
ir_all = get_item_row(전체_key)

tab1, tab2, tab3 = st.tabs(["📊 유형 분류", "📈 영역별 적응", "📋 문항별 상세"])

# ── TAB 1: 유형 분류 ──────────────────────────────────────
with tab1:
    if tr_s is None:
        st.warning("데이터 없음")
    else:
        types = ["저적응위기형", "적응취약형", "적응안정형", "고적응성취형"]
        빈도s = [tr_s[f"{t}_빈도"] for t in types]
        비율s = [tr_s[f"{t}_비율"] for t in types]
        전체비율s = [tr_all[f"{t}_비율"] for t in types] if tr_all is not None else [0] * 4

        cols = st.columns(5)
        total = int(tr_s["총계"]) if pd.notna(tr_s["총계"]) else 0
        with cols[0]:
            st.markdown(f"""
            <div class='metric-card'>
                <div class='val'>{total}명</div>
                <div class='lbl'>총 응답자</div>
            </div>""", unsafe_allow_html=True)
        card_colors = ["#e74c3c", "#e67e22", "#3498db", "#2ecc71"]
        for i, (t, b, r) in enumerate(zip(types, 빈도s, 비율s)):
            with cols[i + 1]:
                pct = f"{r*100:.1f}%" if pd.notna(r) else "0.0%"
                cnt = int(b) if pd.notna(b) else 0
                st.markdown(f"""
                <div class='type-card' style='background:{card_colors[i]}'>
                    <div style='font-size:0.75rem;opacity:0.9'>{t}</div>
                    <div style='font-size:1.6rem;font-weight:700'>{cnt}명</div>
                    <div style='font-size:1rem'>{pct}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("")
        c1, c2 = st.columns(2)

        with c1:
            st.markdown(f"<div class='section-title'>{selected} 유형 분포</div>", unsafe_allow_html=True)
            valid = [(t, b, r) for t, b, r in zip(types, 빈도s, 비율s) if pd.notna(b) and b > 0]
            if valid:
                fig = go.Figure(go.Pie(
                    labels=[v[0] for v in valid],
                    values=[v[1] for v in valid],
                    hole=0.55,
                    marker_colors=[TYPE_COLORS[v[0]] for v in valid],
                    textinfo="label+percent",
                    textfont_size=13,
                    hovertemplate="%{label}<br>%{value}명 (%{percent})<extra></extra>"
                ))
                fig.update_layout(
                    height=340, margin=dict(t=20, b=20, l=20, r=20),
                    showlegend=False,
                    annotations=[dict(text=f"<b>{total}</b><br>명", x=0.5, y=0.5,
                                      font_size=16, showarrow=False)]
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("해당 학과 데이터 없음")

        with c2:
            st.markdown(f"<div class='section-title'>전체 대학 비교</div>", unsafe_allow_html=True)
            fig2 = go.Figure()
            x_labels = [t.replace("형", "") for t in types]
            dept_pct = [r * 100 if pd.notna(r) else 0 for r in 비율s]
            total_pct = [r * 100 if pd.notna(r) else 0 for r in 전체비율s]

            fig2.add_trace(go.Bar(
                name=selected, x=x_labels, y=dept_pct,
                marker_color=[TYPE_COLORS[t] for t in types],
                text=[f"{v:.1f}%" for v in dept_pct],
                textposition="outside", textfont_size=12
            ))
            fig2.add_trace(go.Scatter(
                name="전체 평균", x=x_labels, y=total_pct,
                mode="lines+markers+text",
                line=dict(color="#2c3e50", width=2, dash="dash"),
                marker=dict(size=8, color="#2c3e50"),
                text=[f"{v:.1f}%" for v in total_pct],
                textposition="top center", textfont_size=11
            ))
            fig2.update_layout(
                height=340, margin=dict(t=30, b=30, l=10, r=10),
                yaxis=dict(title="%", range=[0, max(max(dept_pct), max(total_pct)) * 1.25 + 5]),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                plot_bgcolor="white", paper_bgcolor="white",
                bargap=0.3
            )
            fig2.update_xaxes(showgrid=False)
            fig2.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("<div class='section-title'>전체 학과 유형 분포 비교</div>", unsafe_allow_html=True)
        all_dept_types = df_type[df_type["학과"] != "총계"].copy()
        heat_data = []
        for t in types:
            heat_data.append([
                round(v * 100, 1) if pd.notna(v) else 0
                for v in all_dept_types[f"{t}_비율"].tolist()
            ])
        fig3 = go.Figure(go.Heatmap(
            z=heat_data,
            x=all_dept_types["학과"].tolist(),
            y=[t.replace("형", "") for t in types],
            colorscale="RdYlGn",
            text=[[f"{v:.1f}%" for v in row] for row in heat_data],
            texttemplate="%{text}",
            textfont_size=10,
            hovertemplate="%{x}<br>%{y}: %{z:.1f}%<extra></extra>"
        ))
        if selected in all_dept_types["학과"].tolist():
            xi = all_dept_types["학과"].tolist().index(selected)
            fig3.add_shape(
                type="rect", x0=xi - 0.5, x1=xi + 0.5, y0=-0.5, y1=3.5,
                line=dict(color="#e74c3c", width=3), fillcolor="rgba(0,0,0,0)"
            )
        fig3.update_layout(
            height=220, margin=dict(t=10, b=80, l=80, r=10),
            xaxis=dict(tickangle=-45, tickfont_size=10),
        )
        st.plotly_chart(fig3, use_container_width=True)


# ── TAB 2: 영역별 적응 ────────────────────────────────────
with tab2:
    if ar_s is None:
        st.warning("데이터 없음")
    else:
        AREAS_RADAR = ["학업적응", "사회적응", "대학적응", "경제적응", "정서적응", "진로적응"]
        dept_vals = [ar_s.get(f"{a}_평균", np.nan) for a in AREAS_RADAR]
        all_vals  = [ar_all.get(f"{a}_평균", np.nan) for a in AREAS_RADAR] if ar_all is not None else [None] * 6
        dept_sd   = [ar_s.get(f"{a}_SD", 0) for a in AREAS_RADAR]

        c1, c2 = st.columns([1, 1])

        with c1:
            st.markdown(f"<div class='section-title'>적응 영역 레이더 차트</div>", unsafe_allow_html=True)
            theta = AREAS_RADAR + [AREAS_RADAR[0]]
            dept_r = [v if pd.notna(v) else 0 for v in dept_vals] + [dept_vals[0] if pd.notna(dept_vals[0]) else 0]
            all_r  = [v if pd.notna(v) else 0 for v in all_vals]  + [all_vals[0] if pd.notna(all_vals[0]) else 0]

            fig = go.Figure()
            fig.add_trace(go.Scatterpolar(
                r=all_r, theta=theta, name="전체 평균",
                fill="toself", fillcolor="rgba(44,62,80,0.1)",
                line=dict(color="#2c3e50", dash="dash", width=2),
                marker=dict(size=6)
            ))
            fig.add_trace(go.Scatterpolar(
                r=dept_r, theta=theta, name=selected,
                fill="toself", fillcolor="rgba(102,126,234,0.25)",
                line=dict(color="#667eea", width=2.5),
                marker=dict(size=8)
            ))
            fig.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[1, 5], tickfont_size=10)),
                height=360, margin=dict(t=30, b=30, l=30, r=30),
                legend=dict(orientation="h", yanchor="bottom", y=-0.15),
                showlegend=True
            )
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            st.markdown(f"<div class='section-title'>영역별 점수 비교</div>", unsafe_allow_html=True)
            fig2 = go.Figure()
            colors = ["#e74c3c" if v < a - 0.1 else "#2ecc71" if v > a + 0.1 else "#667eea"
                      for v, a in zip(dept_vals, all_vals)]
            fig2.add_trace(go.Bar(
                name=selected,
                x=AREAS_RADAR,
                y=[v if pd.notna(v) else 0 for v in dept_vals],
                marker_color=colors,
                error_y=dict(type="data", array=[s if pd.notna(s) else 0 for s in dept_sd], visible=True, color="#999"),
                text=[f"{v:.2f}" if pd.notna(v) else "-" for v in dept_vals],
                textposition="outside", textfont_size=12
            ))
            if ar_all is not None:
                fig2.add_trace(go.Scatter(
                    name="전체 평균", x=AREAS_RADAR,
                    y=[v if pd.notna(v) else 0 for v in all_vals],
                    mode="lines+markers",
                    line=dict(color="#2c3e50", dash="dash", width=2),
                    marker=dict(size=8, color="#2c3e50")
                ))
            fig2.update_layout(
                height=360, margin=dict(t=30, b=30, l=10, r=10),
                yaxis=dict(range=[0, 5.5], title="평균 점수"),
                plot_bgcolor="white", paper_bgcolor="white",
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                bargap=0.35
            )
            fig2.update_xaxes(showgrid=False)
            fig2.update_yaxes(showgrid=True, gridcolor="#f0f0f0")
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown("<div class='section-title'>⚠️ 중도탈락의도</div>", unsafe_allow_html=True)
        d_val = ar_s.get("중도탈락의도_평균", None)
        a_val = ar_all.get("중도탈락의도_평균", None) if ar_all is not None else None
        d_sd  = ar_s.get("중도탈락의도_SD", None)
        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            if pd.notna(d_val):
                diff = d_val - a_val if pd.notna(a_val) else 0
                color = "#e74c3c" if diff > 0.1 else "#2ecc71" if diff < -0.1 else "#667eea"
                st.markdown(f"""
                <div class='metric-card' style='background:{color}'>
                    <div class='val'>{d_val:.2f}</div>
                    <div class='lbl'>{selected} 평균</div>
                </div>""", unsafe_allow_html=True)
        with mc2:
            if pd.notna(a_val):
                st.markdown(f"""
                <div class='metric-card' style='background:#7f8c8d'>
                    <div class='val'>{a_val:.2f}</div>
                    <div class='lbl'>전체 평균</div>
                </div>""", unsafe_allow_html=True)
        with mc3:
            if pd.notna(d_val) and pd.notna(a_val):
                diff = d_val - a_val
                sign = "▲" if diff > 0 else "▼"
                color = "#e74c3c" if diff > 0 else "#2ecc71"
                st.markdown(f"""
                <div class='metric-card' style='background:{color}'>
                    <div class='val'>{sign} {abs(diff):.2f}</div>
                    <div class='lbl'>전체 대비 차이</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<div class='section-title'>전체 학과 영역별 분포</div>", unsafe_allow_html=True)
        area_sel = st.selectbox("영역 선택", AREAS_RADAR + ["중도탈락의도"], key="area_scatter")
        col_name = f"{area_sel}_평균"
        df_plot = df_area[df_area["학과"] != "총계"][["학과", col_name]].dropna().copy()
        df_plot = df_plot.sort_values(col_name, ascending=True)
        df_plot["color"] = df_plot["학과"].apply(lambda x: "#e74c3c" if x == selected else "#aab7c4")
        df_plot["size"]  = df_plot["학과"].apply(lambda x: 14 if x == selected else 8)

        fig3 = go.Figure()
        for _, row in df_plot.iterrows():
            fig3.add_trace(go.Scatter(
                x=[row[col_name]], y=[row["학과"]],
                mode="markers+text",
                marker=dict(size=row["size"], color=row["color"]),
                text=[f"  {row[col_name]:.2f}"] if row["학과"] == selected else [""],
                textposition="middle right",
                showlegend=False,
                hovertemplate=f"{row['학과']}: {row[col_name]:.2f}<extra></extra>"
            ))
        if ar_all is not None and pd.notna(ar_all.get(col_name)):
            fig3.add_vline(x=ar_all[col_name], line_dash="dash", line_color="#2c3e50",
                           annotation_text=f"전체 {ar_all[col_name]:.2f}", annotation_position="top right")
        fig3.update_layout(
            height=max(400, len(df_plot) * 18),
            margin=dict(t=10, b=10, l=10, r=60),
            xaxis=dict(title="평균 점수", range=[0, 5.5]),
            plot_bgcolor="white", paper_bgcolor="white"
        )
        fig3.update_yaxes(showgrid=True, gridcolor="#f5f5f5")
        st.plotly_chart(fig3, use_container_width=True)


# ── TAB 3: 문항별 상세 ────────────────────────────────────
with tab3:
    if ir_s is None:
        st.warning("데이터 없음")
    else:
        area_tabs = st.tabs([f"📌 {a}" for a in item_map.keys()])

        for ai, (area, items) in enumerate(item_map.items()):
            with area_tabs[ai]:
                if not items:
                    st.info("문항 없음")
                    continue

                st.caption("🔄 역문항은 점수 변환 처리되어 있으며, 모든 문항에서 점수가 높을수록 긍정적 의미입니다.")

                dept_means = []
                dept_sds   = []
                all_means  = []
                labels     = []

                for item_name, mc, sc in items:
                    key_m = f"{area}||{item_name}||평균"
                    key_s = f"{area}||{item_name}||SD"
                    dm = ir_s.get(key_m, np.nan)
                    ds = ir_s.get(key_s, np.nan)
                    am = ir_all.get(key_m, np.nan) if ir_all is not None else np.nan
                    dept_means.append(dm if pd.notna(dm) else 0)
                    dept_sds.append(ds if pd.notna(ds) else 0)
                    all_means.append(am if pd.notna(am) else 0)
                    is_reverse = item_name.endswith("*")
                    clean = item_name.replace("나는 ", "").replace("나의 ", "")
                    if len(clean) > 50:
                        clean = clean[:50] + "…"
                    label = f"🔄 {clean}" if is_reverse else clean
                    labels.append(label)

                colors = ["#e74c3c" if d < a - 0.15 else "#2ecc71" if d > a + 0.15 else "#667eea"
                          for d, a in zip(dept_means, all_means)]

                fig = go.Figure()
                fig.add_trace(go.Bar(
                    name=selected, y=labels, x=dept_means,
                    orientation="h",
                    marker_color=colors,
                    error_x=dict(type="data", array=dept_sds, visible=True, color="#bbb"),
                    text=[f"{v:.2f}" for v in dept_means],
                    textposition="outside",
                    textfont_size=11
                ))
                if ir_all is not None:
                    fig.add_trace(go.Scatter(
                        name="전체 평균", y=labels, x=all_means,
                        mode="markers",
                        marker=dict(symbol="diamond", size=9, color="#2c3e50"),
                    ))
                fig.update_layout(
                    height=max(400, len(items) * 32),
                    margin=dict(t=10, b=10, l=220, r=80),
                    xaxis=dict(range=[0, 5.8], title="평균 점수"),
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="h", yanchor="bottom", y=1.01),
                    bargap=0.2
                )
                fig.update_xaxes(showgrid=True, gridcolor="#f0f0f0")
                fig.update_yaxes(showgrid=False)
                st.plotly_chart(fig, use_container_width=True)

                st.markdown("""
                <small>
                🟥 전체 평균보다 낮음 &nbsp;&nbsp;
                🟦 전체 평균과 유사 &nbsp;&nbsp;
                🟩 전체 평균보다 높음 &nbsp;&nbsp;
                ◆ 전체 평균
                </small>
                """, unsafe_allow_html=True)

                with st.expander("📋 수치 테이블 보기"):
                    tbl = []
                    for item_name, mc, sc in items:
                        key_m = f"{area}||{item_name}||평균"
                        key_s = f"{area}||{item_name}||SD"
                        dm = ir_s.get(key_m, np.nan)
                        ds = ir_s.get(key_s, np.nan)
                        am = ir_all.get(key_m, np.nan) if ir_all is not None else np.nan
                        diff = dm - am if pd.notna(dm) and pd.notna(am) else np.nan
                        is_reverse = "🔄" if item_name.endswith("*") else ""
                        tbl.append({
                            "역": is_reverse,
                            "문항": item_name.rstrip("*").strip(),
                            f"{selected} 평균": f"{dm:.2f}" if pd.notna(dm) else "-",
                            "SD": f"{ds:.2f}" if pd.notna(ds) else "-",
                            "전체 평균": f"{am:.2f}" if pd.notna(am) else "-",
                            "차이": f"{diff:+.2f}" if pd.notna(diff) else "-",
                        })
                    st.dataframe(pd.DataFrame(tbl), use_container_width=True, hide_index=True)
