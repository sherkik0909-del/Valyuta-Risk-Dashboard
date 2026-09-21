import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os


# =========================================================
# 1. SAHIFA SOZLAMALARI
# =========================================================

st.set_page_config(
    page_title="Valyuta Risk Dashboard",
    page_icon="💱",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# 2. PROFESSIONAL CSS DIZAYN
# =========================================================

st.markdown("""
<style>

    /* Umumiy fon */
    .stApp {
        background-color: #f4f6f9;
    }

    /* Asosiy sarlavha */
    .main-title {
        font-size: 34px;
        font-weight: 700;
        color: #172033;
        margin-bottom: 0px;
    }

    .subtitle {
        color: #6b7280;
        font-size: 15px;
        margin-top: 4px;
        margin-bottom: 25px;
    }

    /* KPI karta */
    .kpi-card {
        background: white;
        padding: 18px 20px;
        border-radius: 14px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 3px 12px rgba(0,0,0,0.05);
        min-height: 115px;
    }

    .kpi-title {
        color: #6b7280;
        font-size: 13px;
        font-weight: 500;
        margin-bottom: 8px;
    }

    .kpi-value {
        color: #172033;
        font-size: 25px;
        font-weight: 700;
    }

    .kpi-small {
        color: #6b7280;
        font-size: 12px;
        margin-top: 5px;
    }

    /* Section title */
    .section-title {
        font-size: 21px;
        font-weight: 650;
        color: #172033;
        margin-top: 25px;
        margin-bottom: 10px;
    }

    /* Status */
    .positive {
        color: #059669;
        font-weight: 700;
    }

    .negative {
        color: #dc2626;
        font-weight: 700;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #172033;
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# 3. EXCEL FAYL
# =========================================================

file_path = os.path.join(
    os.path.dirname(__file__),
    "Valyuta.xlsx"
)

try:
    df = pd.read_excel(file_path)

except Exception as e:
    st.error(f"Excel faylni yuklashda xatolik: {e}")
    st.stop()


# =========================================================
# 4. USTUN NOMLARINI TOZALASH
# =========================================================

df.columns = (
    df.columns
    .astype(str)
    .str.replace("\n", " ", regex=False)
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)


# =========================================================
# 5. USTUNLARNI STANDART NOMLARGA O'ZGARTIRISH
# =========================================================

column_map = {
    "Sana": "Date",
    "AQSH dollari (840)": "USD",
    "EVRO (978)": "EUR"
}

df = df.rename(columns=column_map)


required_columns = [
    "Date",
    "USD",
    "EUR"
]

missing = [
    col for col in required_columns
    if col not in df.columns
]

if missing:

    st.error(
        f"Excel faylda quyidagi ustunlar topilmadi: {missing}"
    )

    st.write(
        "Topilgan ustunlar:",
        df.columns.tolist()
    )

    st.stop()


# =========================================================
# 6. DATA TAYYORLASH
# =========================================================

df["Date"] = pd.to_datetime(
    df["Date"],
    errors="coerce"
)

df["USD"] = pd.to_numeric(
    df["USD"],
    errors="coerce"
)

df["EUR"] = pd.to_numeric(
    df["EUR"],
    errors="coerce"
)

df = df.dropna(
    subset=[
        "Date",
        "USD",
        "EUR"
    ]
)

df = df.sort_values(
    "Date"
).reset_index(drop=True)


# =========================================================
# 7. SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        "## 💱 VALYUTA"
    )

    st.markdown(
        "### Risk Dashboard"
    )

    st.markdown("---")

    st.markdown(
        "### 📅 Tahlil davri"
    )

    start_date = st.date_input(
        "Boshlanish sanasi",
        value=df["Date"].min().date()
    )

    end_date = st.date_input(
        "Tugash sanasi",
        value=df["Date"].max().date()
    )

    st.markdown("---")

    st.markdown(
        "### 📊 Ma'lumot"
    )

    st.write(
        f"Qatorlar: **{len(df):,}**"
    )

    st.write(
        f"Boshlanish: **{df['Date'].min().strftime('%d.%m.%Y')}**"
    )

    st.write(
        f"Tugash: **{df['Date'].max().strftime('%d.%m.%Y')}**"
    )


# =========================================================
# 8. SANA TEKSHIRISH
# =========================================================

if start_date > end_date:

    st.error(
        "Boshlanish sanasi tugash sanasidan katta bo'lishi mumkin emas."
    )

    st.stop()


filtered = df[
    (df["Date"].dt.date >= start_date)
    &
    (df["Date"].dt.date <= end_date)
].copy()


if filtered.empty:

    st.warning(
        "Tanlangan davr bo'yicha ma'lumot topilmadi."
    )

    st.stop()


# =========================================================
# 9. HEADER
# =========================================================

st.markdown(
    '<div class="main-title">💱 VALYUTA RISK DASHBOARD</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    f"USD va EUR kurslari dinamikasi • "
    f"{start_date.strftime('%d.%m.%Y')} — "
    f"{end_date.strftime('%d.%m.%Y')}"
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# 10. USD HISOBLASH
# =========================================================

usd_start = filtered["USD"].iloc[0]
usd_end = filtered["USD"].iloc[-1]

usd_change = (
    usd_end / usd_start - 1
) * 100

usd_max = filtered["USD"].max()
usd_min = filtered["USD"].min()

usd_max_date = filtered.loc[
    filtered["USD"].idxmax(),
    "Date"
]

usd_min_date = filtered.loc[
    filtered["USD"].idxmin(),
    "Date"
]


# =========================================================
# 11. EUR HISOBLASH
# =========================================================

eur_start = filtered["EUR"].iloc[0]
eur_end = filtered["EUR"].iloc[-1]

eur_change = (
    eur_end / eur_start - 1
) * 100

eur_max = filtered["EUR"].max()
eur_min = filtered["EUR"].min()

eur_max_date = filtered.loc[
    filtered["EUR"].idxmax(),
    "Date"
]

eur_min_date = filtered.loc[
    filtered["EUR"].idxmin(),
    "Date"
]


# =========================================================
# 12. USD HEADER
# =========================================================

st.markdown(
    '<div class="section-title">🇺🇸 USD / UZS</div>',
    unsafe_allow_html=True
)


c1, c2, c3, c4 = st.columns(4)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Boshlang'ich kurs</div>
            <div class="kpi-value">{usd_start:,.2f}</div>
            <div class="kpi-small">
                {start_date.strftime('%d.%m.%Y')}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Oxirgi kurs</div>
            <div class="kpi-value">{usd_end:,.2f}</div>
            <div class="kpi-small">
                {end_date.strftime('%d.%m.%Y')}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    change_class = (
        "positive"
        if usd_change >= 0
        else "negative"
    )

    arrow = "▲" if usd_change >= 0 else "▼"

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Umumiy o'zgarish</div>
            <div class="kpi-value {change_class}">
                {arrow} {usd_change:.2f}%
            </div>
            <div class="kpi-small">
                Davr bo'yicha
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Maksimum / Minimum</div>
            <div class="kpi-value">
                {usd_max:,.0f}
            </div>
            <div class="kpi-small">
                Min: {usd_min:,.0f}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


st.caption(
    f"📌 Maksimum: {usd_max:,.2f} "
    f"({usd_max_date.strftime('%d.%m.%Y')})  •  "
    f"Minimum: {usd_min:,.2f} "
    f"({usd_min_date.strftime('%d.%m.%Y')})"
)


# =========================================================
# 13. USD GRAFIK
# =========================================================

fig_usd = go.Figure()


fig_usd.add_trace(
    go.Scatter(
        x=filtered["Date"],
        y=filtered["USD"],
        mode="lines+markers",
        name="USD / UZS",

        line=dict(
            color="#2563EB",
            width=3,
            shape="spline",
            smoothing=1.1
        ),

        marker=dict(
            size=5,
            color="#2563EB",
            line=dict(
                width=1,
                color="white"
            )
        ),

        hovertemplate=
        "<b>%{x|%d.%m.%Y}</b><br>"
        "USD: %{y:,.2f}<extra></extra>"
    )
)


fig_usd.update_layout(

    height=470,

    margin=dict(
        l=20,
        r=20,
        t=25,
        b=20
    ),

    plot_bgcolor="white",
    paper_bgcolor="white",

    hovermode="x unified",

    xaxis=dict(
        title="",
        showgrid=False,
        showline=True,
        linecolor="#d1d5db",
        tickfont=dict(
            color="#6b7280"
        )
    ),

    yaxis=dict(
        title="UZS",
        showgrid=True,
        gridcolor="#eef1f5",
        zeroline=False,
        tickfont=dict(
            color="#6b7280"
        )
    ),

    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    )
)


st.plotly_chart(
    fig_usd,
    use_container_width=True
)


# =========================================================
# 14. EUR HEADER
# =========================================================

st.markdown(
    '<div class="section-title">🇪🇺 EUR / UZS</div>',
    unsafe_allow_html=True
)


c1, c2, c3, c4 = st.columns(4)


with c1:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Boshlang'ich kurs</div>
            <div class="kpi-value">{eur_start:,.2f}</div>
            <div class="kpi-small">
                {start_date.strftime('%d.%m.%Y')}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c2:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Oxirgi kurs</div>
            <div class="kpi-value">{eur_end:,.2f}</div>
            <div class="kpi-small">
                {end_date.strftime('%d.%m.%Y')}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c3:

    change_class = (
        "positive"
        if eur_change >= 0
        else "negative"
    )

    arrow = "▲" if eur_change >= 0 else "▼"

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Umumiy o'zgarish</div>
            <div class="kpi-value {change_class}">
                {arrow} {eur_change:.2f}%
            </div>
            <div class="kpi-small">
                Davr bo'yicha
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


with c4:

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">Maksimum / Minimum</div>
            <div class="kpi-value">
                {eur_max:,.0f}
            </div>
            <div class="kpi-small">
                Min: {eur_min:,.0f}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


st.caption(
    f"📌 Maksimum: {eur_max:,.2f} "
    f"({eur_max_date.strftime('%d.%m.%Y')})  •  "
    f"Minimum: {eur_min:,.2f} "
    f"({eur_min_date.strftime('%d.%m.%Y')})"
)


# =========================================================
# 15. EUR GRAFIK
# =========================================================

fig_eur = go.Figure()


fig_eur.add_trace(
    go.Scatter(
        x=filtered["Date"],
        y=filtered["EUR"],
        mode="lines+markers",
        name="EUR / UZS",

        line=dict(
            color="#7C3AED",
            width=3,
            shape="spline",
            smoothing=1.1
        ),

        marker=dict(
            size=5,
            color="#7C3AED",
            line=dict(
                width=1,
                color="white"
            )
        ),

        hovertemplate=
        "<b>%{x|%d.%m.%Y}</b><br>"
        "EUR: %{y:,.2f}<extra></extra>"
    )
)


fig_eur.update_layout(

    height=470,

    margin=dict(
        l=20,
        r=20,
        t=25,
        b=20
    ),

    plot_bgcolor="white",
    paper_bgcolor="white",

    hovermode="x unified",

    xaxis=dict(
        title="",
        showgrid=False,
        showline=True,
        linecolor="#d1d5db",
        tickfont=dict(
            color="#6b7280"
        )
    ),

    yaxis=dict(
        title="UZS",
        showgrid=True,
        gridcolor="#eef1f5",
        zeroline=False,
        tickfont=dict(
            color="#6b7280"
        )
    ),

    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    )
)


st.plotly_chart(
    fig_eur,
    use_container_width=True
)


# =========================================================
# 16. DATA JADVAL
# =========================================================

with st.expander("📊 Batafsil ma'lumotlarni ko'rish"):

    display_df = filtered.copy()

    display_df["Date"] = display_df[
        "Date"
    ].dt.strftime("%d.%m.%Y")

    display_df = display_df.rename(
        columns={
            "Date": "Sana",
            "USD": "USD / UZS",
            "EUR": "EUR / UZS"
        }
    )

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )