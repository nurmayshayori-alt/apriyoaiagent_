import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import os
from io import BytesIO
from datetime import datetime
from groq import Groq
from dotenv import load_dotenv

# =====================================================
# 1. PAGE CONFIG & CLEAN MODERN LIGHT THEME CSS
# =====================================================
st.set_page_config(
    page_title="APRIYO | AI Financial & Accounting Suite",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #0f172a;
    }

    .stApp {
        background-color: #f8fafc;
        color: #0f172a;
    }

    section[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e2e8f0;
    }

    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 18px 22px;
        border-radius: 20px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="metric-container"] label {
        color: #64748b !important;
        font-size: 13.5px !important;
        font-weight: 600;
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #0f172a !important;
        font-size: 24px !important;
        font-weight: 700;
    }

    .hero-banner {
        background: linear-gradient(120deg, #eef2ff 0%, #f0f9ff 60%, #ffffff 100%);
        border: 1px solid #c7d2fe;
        padding: 24px 30px;
        border-radius: 22px;
        margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(79, 70, 229, 0.05);
    }
    .hero-badge {
        display: inline-block;
        background-color: #e0e7ff;
        color: #3730a3;
        padding: 4px 14px;
        border-radius: 999px;
        font-size: 11px;
        font-weight: 700;
        margin-bottom: 8px;
        border: 1px solid #c7d2fe;
    }
    .hero-banner h1 {
        color: #0f172a !important;
        margin: 0;
        font-size: 26px;
        font-weight: 700;
    }
    .hero-banner p {
        color: #475569 !important;
        margin: 5px 0 0 0;
        font-size: 14px;
    }

    .empty-state-box {
        background-color: #ffffff;
        border: 2px dashed #cbd5e1;
        border-radius: 22px;
        padding: 48px 32px;
        text-align: center;
        margin-top: 16px;
    }

    .status-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 14px 16px;
        margin-top: 10px;
        margin-bottom: 12px;
        font-size: 12.5px;
        line-height: 1.7;
        color: #1e293b;
    }

    .audit-alert {
        background-color: #fff1f2;
        border: 1px solid #fecdd3;
        border-left: 5px solid #e11d48;
        border-radius: 14px;
        padding: 13px 20px;
        margin-bottom: 18px;
        font-size: 13.5px;
        color: #881337;
    }
    .audit-ok {
        background-color: #ecfdf5;
        border: 1px solid #a7f3d0;
        border-left: 5px solid #10b981;
        border-radius: 14px;
        padding: 13px 20px;
        margin-bottom: 18px;
        font-size: 13.5px;
        color: #065f46;
    }

    .white-panel {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 20px 24px;
        margin-top: 10px;
        margin-bottom: 12px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.03);
    }
    .highlight-row {
        background-color: #f8fafc;
        border: 1px solid #f1f5f9;
        border-radius: 12px;
        padding: 10px 14px;
        margin-bottom: 8px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 13.5px;
    }

    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #ffffff;
        padding: 6px;
        border-radius: 999px;
        border: 1px solid #e2e8f0;
        width: fit-content;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 999px;
        color: #475569;
        padding: 7px 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #4f46e5, #0284c7) !important;
        color: #ffffff !important;
    }

    .stButton > button, .stDownloadButton > button {
        border-radius: 999px !important;
        padding: 8px 18px !important;
        font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# 2. LOAD GROQ API KEY & HELPERS
# =====================================================
load_dotenv()
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

try:
    if "GROQ_API_KEY" in st.secrets:
        GROQ_API_KEY = st.secrets["GROQ_API_KEY"]
except Exception:
    pass

if not GROQ_API_KEY:
    st.error("🚨 GROQ_API_KEY belum diatur di Streamlit Secrets!")
    st.stop()

client = Groq(api_key=GROQ_API_KEY)

def to_excel_bytes(dataframe, sheet_name="Data"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def generate_executive_report(title, df_table, ai_summary):
    now = datetime.now().strftime("%d %B %Y, %H:%M WIB")
    report = f"""============================================================
APRIYO AI FINANCIAL SUITE - EXECUTIVE REPORT
Modul   : {title}
Tanggal : {now}
============================================================

1. RINGKASAN DATA KEUANGAN
------------------------------------------------------------
{df_table.to_string(index=False)}

2. ANALISIS & REKOMENDASI AI
------------------------------------------------------------
{ai_summary if ai_summary else "Belum ada analisis AI yang di-generate."}

============================================================
Generated by APRIYO AI Accounting & FP&A Suite
"""
    return report.encode("utf-8")

# =====================================================
# 3. SIDEBAR NAVIGATION
# =====================================================
with st.sidebar:
    st.markdown("## ✨ APRIYO AI")
    st.caption("Intelligent Accounting & FP&A Suite")
    st.markdown("<br>", unsafe_allow_html=True)

    app_mode = st.radio(
        "🧭 Pilih Modul Analisis:",
        ["📊 Budget vs. Actuals", "🚀 Scenario Planning"],
        index=0
    )

    st.markdown("---")
    selected_model = st.selectbox(
        "🤖 Model AI (Groq)",
        ["openai/gpt-oss-20b", "openai/gpt-oss-120b", "llama-3.3-70b-versatile"],
        index=0
    )

# =====================================================
# MODUL 1: BUDGET VS. ACTUALS (VARIANCE & AUDIT)
# =====================================================
if app_mode == "📊 Budget vs. Actuals":
    sample_var = pd.DataFrame({
        "Category": ["Sales Revenue", "COGS", "Marketing", "Salaries", "IT & Cloud", "Operations"],
        "Budget": [150000000, 60000000, 15000000, 35000000, 8000000, 5000000],
        "Actual": [145000000, 64000000, 18500000, 35000000, 7200000, 5600000]
    })

    with st.sidebar:
        uploaded_var = st.file_uploader(
            "📂 Upload Data Budget (.xlsx)",
            type=["xlsx"],
            key="upload_var",
            help="Kolom wajib: Category, Budget, Actual"
        )
        st.download_button(
            label="📥 Contoh Excel Variance",
            data=to_excel_bytes(sample_var, "Budget_vs_Actual"),
            file_name="APRIYO_Template_Variance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("---")
        audit_threshold = st.slider(
            "🛡️ Batas Toleransi Audit (%)",
            min_value=5, max_value=30, value=10, step=1
        )

    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • VARIANCE & AUDIT INTELLIGENCE</span>
<h1>Budget vs. Actuals Analytics</h1>
<p>Deteksi anomali anggaran otomatis, skor kesehatan finansial, dan asisten audit berbasis AI.</p>
</div>
""", unsafe_allow_html=True)

    if uploaded_var is not None:
        try:
            df = pd.read_excel(uploaded_var)
            if not all(col in df.columns for col in ["Category", "Budget", "Actual"]):
                st.error("⚠️ File wajib memiliki kolom: 'Category', 'Budget', dan 'Actual'.")
                st.stop()

            df["Budget"] = pd.to_numeric(df["Budget"], errors="coerce")
            df["Actual"] = pd.to_numeric(df["Actual"], errors="coerce")

            with st.sidebar:
                st.markdown("---")
                cats = st.multiselect(
                    "🔍 Filter Kategori Aktif:",
                    options=df["Category"].unique().tolist(),
                    default=df["Category"].unique().tolist()
                )

            df_f = df[df["Category"].isin(cats)].copy()
            df_f["Variance"] = df_f["Actual"] - df_f["Budget"]
            df_f["Variance %"] = (df_f["Variance"] / df_f["Budget"].replace(0, float("nan"))) * 100
            df_f["Audit Flag"] = df_f["Variance %"].apply(
                lambda x: "🚨 Anomali (>Batas)" if abs(x) > audit_threshold else "✅ Normal"
            )

            with st.sidebar:
                st.markdown(f"""
<div class="status-card">
<div style="color: #0f172a; font-weight: 700; margin-bottom: 4px;">🤖 AI Copilot Status</div>
<div style="color: #059669; font-weight: 600;">● Engine: Groq Ultra-Fast</div>
<div style="color: #334155;">● File: {uploaded_var.name}</div>
<div style="color: #334155;">● Data: {len(df_f)} Kategori Valid</div>
</div>
""", unsafe_allow_html=True)
                st.download_button(
                    "📄 Unduh Memo Eksekutif",
                    data=generate_executive_report("Budget vs. Actuals Variance", df_f, st.session_state.get("var_ai", "")),
                    file_name="APRIYO_Memo_Eksekutif.txt",
                    mime="text/plain",
                    use_container_width=True
                )

            tot_b = df_f["Budget"].sum()
            tot_a = df_f["Actual"].sum()
            tot_v = tot_a - tot_b
            v_pct = (tot_v / tot_b * 100) if tot_b != 0 else 0
            realisasi = (tot_a / tot_b * 100) if tot_b != 0 else 0

            avg_abs_dev = df_f["Variance %"].abs().mean() if not df_f.empty else 0
            health_score = max(0, min(100, int(100 - avg_abs_dev)))

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 Total Budget", f"{tot_b:,.0f}", f"{len(df_f)} Pos Anggaran")
            c2.metric("💳 Total Actual", f"{tot_a:,.0f}", f"Rasio: {realisasi:.1f}%")
            c3.metric("📉 Total Variance", f"{tot_v:,.0f}", delta=f"{v_pct:+.1f}%", delta_color="inverse")
            c4.metric("🛡️ Budget Health Score", f"{health_score} / 100", delta="Sehat" if health_score >= 80 else "Perlu Audit")

            anomalies = df_f[df_f["Variance %"].abs() > audit_threshold]
            if not anomalies.empty:
                anom_list = ", ".join([f"{row['Category']} ({row['Variance %']:+.1f}%)" for _, row in anomalies.iterrows()])
                st.markdown(f"""
<div class="audit-alert">
<strong>🚨 Peringatan Audit Otomatis:</strong> Ditemukan <b>{len(anomalies)} kategori</b> dengan penyimpangan melebihi batas toleransi ({audit_threshold}%): <b>{anom_list}</b>.
</div>
""", unsafe_allow_html=True)
            else:
                st.markdown(f"""
<div class="audit-ok">
<strong>✅ Status Audit Aman:</strong> Seluruh kategori berada dalam batas toleransi deviasi ({audit_threshold}%).
</div>
""", unsafe_allow_html=True)

            t1, t2, t3 = st.tabs(["✨ Visualisasi Grafik", "📋 Rincian & Export", "🤖 AI Audit Analyst"])

            with t1:
                col_g1, col_g2, col_g3 = st.columns([1.2, 1.2, 0.9])

                with col_g1:
                    fig_area = go.Figure()
                    fig_area.add_trace(go.Scatter(
                        x=df_f["Category"], y=df_f["Budget"],
                        name="Budget", mode="lines+markers",
                        line=dict(color="#4f46e5", width=3, shape="spline"),
                        fill="tozeroy", fillcolor="rgba(79, 70, 229, 0.10)"
                    ))
                    fig_area.add_trace(go.Scatter(
                        x=df_f["Category"], y=df_f["Actual"],
                        name="Actual", mode="lines+markers",
                        line=dict(color="#0ea5e9", width=3, shape="spline"),
                        fill="tozeroy", fillcolor="rgba(14, 165, 233, 0.10)"
                    ))
                    fig_area.update_layout(
                        title=dict(text="Tren Budget vs. Actual", font=dict(color="#0f172a")),
                        template="plotly_white",
                        paper_bgcolor="#ffffff",
                        plot_bgcolor="#ffffff",
                        font=dict(color="#0f172a"),
                        xaxis=dict(showgrid=False, tickfont=dict(color="#334155")),
                        yaxis=dict(gridcolor="#f1f5f9", tickfont=dict(color="#334155")),
                        legend=dict(orientation="h", y=1.12, x=1, xanchor="right"),
                        margin=dict(t=55, b=30, l=20, r=20),
                        height=350
                    )
                    st.plotly_chart(fig_area, use_container_width=True)

                with col_g2:
                    colors = ["#e11d48" if v > 0 else "#10b981" for v in df_f["Variance"]]
                    fig_var = go.Figure(go.Bar(
                        x=df_f["Variance"],
                        y=df_f["Category"],
                        orientation="h",
                        marker=dict(color=colors, cornerradius=8),
                        text=df_f["Variance"],
                        texttemplate="%{text:,.0f}",
                        textposition="auto"
                    ))
                    fig_var.add_vline(x=0, line_width=1, line_dash="dot", line_color="#94a3b8")
                    fig_var.update_layout(
                        title=dict(text="Deviasi (Variance) per Kategori", font=dict(color="#0f172a")),
                        template="plotly_white",
                        paper_bgcolor="#ffffff",
                        plot_bgcolor="#ffffff",
                        font=dict(color="#0f172a"),
                        xaxis=dict(gridcolor="#f1f5f9", tickfont=dict(color="#334155")),
                        yaxis=dict(autorange="reversed", showgrid=False, tickfont=dict(color="#334155")),
                        margin=dict(t=55, b=30, l=20, r=20),
                        height=350
                    )
                    st.plotly_chart(fig_var, use_container_width=True)

                with col_g3:
                    fig_donut = go.Figure(data=[go.Pie(
                        labels=df_f["Category"],
                        values=df_f["Actual"].abs(),
                        hole=0.62,
                        textinfo="percent",
                        marker=dict(colors=["#4f46e5", "#0ea5e9", "#e11d48", "#10b981", "#f59e0b", "#6366f1"])
                    )])
                    fig_donut.update_layout(
                        title=dict(text="Proporsi Realisasi", font=dict(color="#0f172a")),
                        template="plotly_white",
                        paper_bgcolor="#ffffff",
                        font=dict(color="#0f172a"),
                        showlegend=True,
                        legend=dict(orientation="h", y=-0.15, font=dict(size=10)),
                        margin=dict(t=55, b=40, l=10, r=10),
                        height=350
                    )
                    st.plotly_chart(fig_donut, use_container_width=True)

                col_b1, col_b2 = st.columns(2)

                with col_b1:
                    top_devs = df_f.reindex(df_f["Variance"].abs().sort_values(ascending=False).index).head(3)
                    rows_html = ""
                    for idx, (_, r) in enumerate(top_devs.iterrows(), 1):
                        badge_col = "#e11d48" if r["Variance"] > 0 else "#059669"
                        status_txt = "Over Budget" if r["Variance"] > 0 else "Efisiensi"
                        rows_html += f'<div class="highlight-row"><span style="color: #0f172a; font-weight: 600;">{idx}. {r["Category"]}</span><span style="color: {badge_col}; font-weight: 700;">{status_txt} {r["Variance"]:+,.0f} ({r["Variance %"]:+.1f}%)</span></div>'

                    st.markdown(f"""
<div class="white-panel">
<div style="color: #0f172a; font-size: 16px; font-weight: 700; margin-bottom: 12px;">🏆 Top Deviasi & Sorotan Kategori</div>
{rows_html}
</div>
""", unsafe_allow_html=True)

                with col_b2:
                    st.markdown("""
<div class="white-panel">
<div style="color: #0f172a; font-size: 16px; font-weight: 700; margin-bottom: 4px;">⚡ AI Quick Ask</div>
<p style="color: #64748b; font-size: 13px; margin-bottom: 0;">Tanyakan insight cepat seputar data anggaran ini ke AI.</p>
</div>
""", unsafe_allow_html=True)
                    q_col1, q_col2 = st.columns([3, 1.2])
                    quick_q = q_col1.text_input("Tanya cepat ke AI...", label_visibility="collapsed", placeholder="Misal: Strategi hemat biaya...")
                    if q_col2.button("Kirim", key="btn_quick_var", use_container_width=True) and quick_q:
                        with st.spinner("Menjawab..."):
                            ans_q = client.chat.completions.create(
                                messages=[
                                    {"role": "system", "content": "You are an AI financial analyst. Answer briefly in Indonesian."},
                                    {"role": "user", "content": f"Data:\n{df_f.to_string(index=False)}\n\nPertanyaan: {quick_q}"}
                                ],
                                model=selected_model
                            ).choices[0].message.content
                            st.success(ans_q)

            with t2:
                col_hd, col_dl1, col_dl2 = st.columns([2.5, 1, 1])
                col_hd.subheader("Tabel Audit & Variance")
                col_dl1.download_button(
                    "💾 Export Excel",
                    data=to_excel_bytes(df_f, "Variance"),
                    file_name="APRIYO_Hasil_Variance.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                col_dl2.download_button(
                    "📄 Download Memo (.txt)",
                    data=generate_executive_report("Budget vs. Actuals Variance", df_f, st.session_state.get("var_ai", "")),
                    file_name="APRIYO_Executive_Memo.txt",
                    mime="text/plain",
                    use_container_width=True
                )
                st.dataframe(
                    df_f.style.format({
                        "Budget": "{:,.2f}",
                        "Actual": "{:,.2f}",
                        "Variance": "{:,.2f}",
                        "Variance %": "{:.2f}%"
                    }),
                    use_container_width=True
                )

            with t3:
                c_ai1, c_ai2 = st.columns([1.2, 1])
                with c_ai1:
                    st.subheader("📑 Executive AI Audit Commentary")
                    if st.button("✨ Generate Analisis Lengkap", type="primary"):
                        with st.spinner("AI sedang mengaudit data..."):
                            res = client.chat.completions.create(
                                messages=[
                                    {"role": "system", "content": "You are a senior financial auditor. Provide structured insights, highlight audit anomalies, overspending risks, and recommendations in Indonesian."},
                                    {"role": "user", "content": f"Analisis data Budget vs Actual ini (Batas toleransi audit {audit_threshold}%):\n{df_f.to_string(index=False)}"}
                                ],
                                model=selected_model
                            )
                            st.session_state.var_ai = res.choices[0].message.content
                    if "var_ai" in st.session_state:
                        st.markdown(st.session_state.var_ai)

                with c_ai2:
                    st.subheader("💬 Tanya AI Auditor")
                    q_var = st.text_input("Tanyakan sesuatu tentang data ini:")
                    if st.button("Kirim Pertanyaan") and q_var:
                        with st.spinner("Menjawab..."):
                            chat_res = client.chat.completions.create(
                                messages=[
                                    {"role": "system", "content": "You are an AI financial analyst. Answer concisely in Indonesian."},
                                    {"role": "user", "content": f"Data:\n{df_f.to_string(index=False)}\n\nPertanyaan: {q_var}"}
                                ],
                                model=selected_model
                            )
                            st.info(chat_res.choices[0].message.content)

        except Exception as e:
            st.error(f"Terjadi kesalahan saat membaca file: {e}")

    else:
        st.markdown("""
<div class="empty-state-box">
<h3 style="color: #0f172a; margin-bottom: 8px;">📂 Belum Ada File Data yang Diunggah</h3>
<p style="color: #475569; font-size: 14.5px; max-width: 540px; margin: 0 auto;">
Silakan unggah file Excel <b>Budget vs. Actual (.xlsx)</b> melalui panel di sebelah kiri untuk memulai perhitungan variance, deteksi anomali audit, dan visualisasi grafik.<br><br>
<i>Belum punya file? Klik tombol <b>"📥 Contoh Excel Variance"</b> di sidebar kiri untuk mengunduh template siap pakai.</i>
</p>
</div>
""", unsafe_allow_html=True)

# =====================================================
# MODUL 2: SCENARIO PLANNING & WHAT-IF SIMULATOR
# =====================================================
elif app_mode == "🚀 Scenario Planning":
    sample_fc = pd.DataFrame({
        "Category": ["Product Revenue", "Service Income", "Direct Labor", "Raw Materials", "Marketing Opex"],
        "Base Forecast": [250000000, 90000000, 75000000, 85000000, 30000000]
    })

    with st.sidebar:
        uploaded_fc = st.file_uploader(
            "📂 Upload File Forecast (.xlsx)",
            type=["xlsx"],
            key="upload_fc",
            help="Kolom wajib: Category, Base Forecast"
        )
        st.download_button(
            label="📥 Contoh Excel Forecast",
            data=to_excel_bytes(sample_fc, "Forecast"),
            file_name="APRIYO_Template_Forecast.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("---")
        st.markdown("**🎛️ Interactive What-If Simulator**")
        custom_shock = st.slider(
            "Simulasi Perubahan Custom (%)",
            min_value=-50, max_value=50, value=0, step=5
        )
        scenario_prompt = st.text_area(
            "📝 Asumsi Skenario Strategis:",
            placeholder="Contoh: Pendapatan turun 10% dan biaya bahan baku naik 5%.",
            height=90
        )
        run_sim = st.button("🚀 Jalankan Analisis AI", type="primary", use_container_width=True)

    st.markdown("""
<div class="hero-banner">
<span class="hero-badge">APRIYO • PREDICTIVE FP&A & WHAT-IF SIMULATOR</span>
<h1>Scenario Planning & Strategic Copilot</h1>
<p>Simulasi proyeksi multi-skenario interaktif lengkap dengan What-If Slider dan mitigasi risiko AI.</p>
</div>
""", unsafe_allow_html=True)

    if uploaded_fc is not None:
        df_fc = pd.read_excel(uploaded_fc)
        if not all(col in df_fc.columns for col in ["Category", "Base Forecast"]):
            st.error("⚠️ File wajib memiliki kolom 'Category' dan 'Base Forecast'.")
            st.stop()

        np.random.seed(42)
        df_fc["Optimistic"] = df_fc["Base Forecast"] * np.random.uniform(1.10, 1.25, len(df_fc))
        df_fc["Pessimistic"] = df_fc["Base Forecast"] * np.random.uniform(0.75, 0.90, len(df_fc))
        df_fc["Worst Case"] = df_fc["Base Forecast"] * np.random.uniform(0.55, 0.70, len(df_fc))
        df_fc[f"Custom ({custom_shock:+d}%)"] = df_fc["Base Forecast"] * (1 + custom_shock / 100.0)
        custom_col = f"Custom ({custom_shock:+d}%)"

        with st.sidebar:
            st.markdown(f"""
<div class="status-card">
<div style="color: #0f172a; font-weight: 700; margin-bottom: 4px;">🤖 AI Copilot Status</div>
<div style="color: #059669; font-weight: 600;">● Engine: Groq Ultra-Fast</div>
<div style="color: #334155;">● File: {uploaded_fc.name}</div>
<div style="color: #0284c7; font-weight: 600;">● What-If Shock: {custom_shock:+d}%</div>
</div>
""", unsafe_allow_html=True)
            st.download_button(
                "📄 Unduh Memo Forecast",
                data=generate_executive_report("Scenario Planning Forecast", df_fc, st.session_state.get("fc_ai_summary", "")),
                file_name="APRIYO_Memo_Forecast.txt",
                mime="text/plain",
                use_container_width=True
            )

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("📌 Base Forecast", f"{df_fc['Base Forecast'].sum():,.0f}", f"{len(df_fc)} Kategori")
        k2.metric("🌟 Optimistic", f"{df_fc['Optimistic'].sum():,.0f}", delta="+10% s/d +25%")
        k3.metric("🚨 Worst Case", f"{df_fc['Worst Case'].sum():,.0f}", delta="-30% s/d -45%", delta_color="inverse")
        k4.metric(f"🎛️ {custom_col}", f"{df_fc[custom_col].sum():,.0f}", delta=f"{custom_shock:+d}% Simulasi")

        if run_sim:
            df_prev = df_fc.head(20).to_string(index=False)
            with st.spinner("Menyusun strategi keuangan..."):
                res_fc = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": "You are an AI Financial Copilot specializing in scenario planning, cash flow, and risk strategy. Respond in Indonesian."},
                        {"role": "user", "content": f"Proyeksi:\n{df_prev}\nAsumsi Skenario: {scenario_prompt}\nSimulasi Custom: {custom_shock:+d}%\nBerikan analisis dan rekomendasi jangka pendek & panjang."}
                    ],
                    model=selected_model
                )
                st.session_state.fc_ai_summary = res_fc.choices[0].message.content

        t_fc1, t_fc2, t_fc3 = st.tabs(["✨ Grafik Proyeksi", "📋 Tabel & Export", "🤖 AI Strategic Copilot"])

        with t_fc1:
            fig_fc = go.Figure(data=[
                go.Bar(name="Base Forecast", x=df_fc["Category"], y=df_fc["Base Forecast"], marker=dict(color="#4f46e5", cornerradius=6)),
                go.Bar(name="Optimistic", x=df_fc["Category"], y=df_fc["Optimistic"], marker=dict(color="#10b981", cornerradius=6)),
                go.Bar(name="Pessimistic", x=df_fc["Category"], y=df_fc["Pessimistic"], marker=dict(color="#f59e0b", cornerradius=6)),
                go.Bar(name="Worst Case", x=df_fc["Category"], y=df_fc["Worst Case"], marker=dict(color="#e11d48", cornerradius=6)),
                go.Scatter(name=custom_col, x=df_fc["Category"], y=df_fc[custom_col], mode="lines+markers", line=dict(color="#0f172a", width=3, dash="dash"))
            ])
            fig_fc.update_layout(
                title=dict(text="Simulasi Proyeksi Multi-Skenario & Garis What-If Custom", font=dict(color="#0f172a")),
                barmode="group",
                template="plotly_white",
                paper_bgcolor="#ffffff",
                plot_bgcolor="#ffffff",
                font=dict(color="#0f172a"),
                xaxis=dict(showgrid=False, tickfont=dict(color="#334155")),
                yaxis=dict(gridcolor="#f1f5f9", tickfont=dict(color="#334155")),
                legend=dict(orientation="h", y=1.1, x=1, xanchor="right"),
                height=380
            )
            st.plotly_chart(fig_fc, use_container_width=True)

        with t_fc2:
            col_f1, col_f2, col_f3 = st.columns([2.5, 1, 1])
            col_f1.subheader("Data Proyeksi Skenario")
            col_f2.download_button(
                "💾 Export Excel",
                data=to_excel_bytes(df_fc, "Scenario_Forecast"),
                file_name="APRIYO_Hasil_Forecast.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            col_f3.download_button(
                "📄 Download Memo (.txt)",
                data=generate_executive_report("Scenario Planning Forecast", df_fc, st.session_state.get("fc_ai_summary", "")),
                file_name="APRIYO_Forecast_Memo.txt",
                mime="text/plain",
                use_container_width=True
            )
            st.dataframe(
                df_fc.style.format({
                    "Base Forecast": "{:,.2f}",
                    "Optimistic": "{:,.2f}",
                    "Pessimistic": "{:,.2f}",
                    "Worst Case": "{:,.2f}",
                    custom_col: "{:,.2f}"
                }),
                use_container_width=True
            )

        with t_fc3:
            c_s1, c_s2 = st.columns([1.2, 1])
            df_prev = df_fc.head(20).to_string(index=False)

            with c_s1:
                st.subheader("📑 Rekomendasi Strategis AI")
                if "fc_ai_summary" in st.session_state:
                    st.markdown(st.session_state.fc_ai_summary)
                else:
                    st.info("Klik tombol **🚀 Jalankan Analisis AI** di panel kiri untuk menyusun rekomendasi strategis.")

            with c_s2:
                st.subheader("💬 Diskusi Copilot")
                if "fc_chat" not in st.session_state:
                    st.session_state.fc_chat = []

                q_fc = st.text_input("Tanyakan dampak EBIT, Cash Flow, atau mitigasi:")
                b1, b2 = st.columns([3, 1])
                if b1.button("Kirim Diskusi", type="primary", use_container_width=True) and q_fc:
                    ans = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": "You are an AI Financial Copilot. Explain clearly in Indonesian."},
                            *st.session_state.fc_chat,
                            {"role": "user", "content": f"Data:\n{df_prev}\nAsumsi: {scenario_prompt}\n\nPertanyaan: {q_fc}"}
                        ],
                        model=selected_model
                    ).choices[0].message.content
                    st.session_state.fc_chat.append({"role": "user", "content": q_fc})
                    st.session_state.fc_chat.append({"role": "assistant", "content": ans})

                if b2.button("Reset", use_container_width=True):
                    st.session_state.fc_chat = []
                    st.rerun()

                for m in st.session_state.fc_chat:
                    with st.chat_message(m["role"]):
                        st.write(m["content"])

    else:
        st.markdown("""
<div class="empty-state-box">
<h3 style="color: #0f172a; margin-bottom: 8px;">📂 Belum Ada File Forecast yang Diunggah</h3>
<p style="color: #475569; font-size: 14.5px; max-width: 540px; margin: 0 auto;">
Silakan unggah file Excel <b>Forecast (.xlsx)</b> melalui panel di sebelah kiri untuk memulai simulasi multi-skenario dan analisis strategis AI.<br><br>
<i>Belum punya file? Klik tombol <b>"📥 Contoh Excel Forecast"</b> di sidebar kiri untuk mengunduh template siap pakai.</i>
</p>
</div>
""", unsafe_allow_html=True)
