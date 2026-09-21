import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os


# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="Valyuta Risk Dashboard",
    page_icon="💱",
    layout="wide"
)


# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>

/* =====================================================
   UMUMIY FON
   ===================================================== */

.stApp {
    background-color: #f4f6f9;
}


/* =====================================================
   CHAP SIDEBAR
   ===================================================== */

section[data-testid="stSidebar"] {
    background-color: #172033 !important;
}

section[data-testid="stSidebar"] * {
    color: white !important;
}

/* Brauzerga sidebar qorong'i ekanini aytamiz (input'lar ham qorong'i bo'ladi) */
[data-testid="stSidebar"] {
    color-scheme: dark;
}


/* =====================================================
   VALYUTA SELECTBOX
   ===================================================== */

/* Yopiq holatdagi katak (barcha qatlamlar). Rang: #2a3752 */
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"],
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] div,
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] input {
    background: #2a3752 !important;
    background-color: #2a3752 !important;
    color: white !important;
    -webkit-text-fill-color: white !important;
}

/* Chegara */
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
    border: 1px solid #6b7280 !important;
    border-radius: 8px !important;
}

/* Pastga ochiladigan strelka */
[data-testid="stSidebar"] [data-testid="stSelectbox"] [data-baseweb="select"] svg {
    fill: white !important;
    color: white !important;
}

/* Tanlov ochilganda chiqadigan ro'yxat (sidebar'dan tashqarida chiqadi) */
div[data-baseweb="popover"] ul[role="listbox"] {
    background-color: #2a3752 !important;
}

div[data-baseweb="popover"] li[role="option"] {
    background-color: #2a3752 !important;
    color: white !important;
}

div[data-baseweb="popover"] li[role="option"]:hover,
div[data-baseweb="popover"] li[aria-selected="true"] {
    background-color: #3b4d72 !important;
}


/* =====================================================
   SANA KIRITILADIGAN KATAK
   Rangni o'zgartirish uchun faqat #2a3752 ni almashtiring
   ===================================================== */

/* 1) Katakning BARCHA qatlamlari (tashqi, ichki va input) */
[data-testid="stSidebar"] [data-testid="stDateInput"] [data-baseweb],
[data-testid="stSidebar"] [data-testid="stDateInput"] [data-baseweb] > div,
[data-testid="stSidebar"] [data-testid="stDateInput"] input,
[data-testid="stSidebar"] [data-testid="stDateInputField"] {
    background: #2a3752 !important;
    background-color: #2a3752 !important;
}

/* 2) Katak chegarasi */
[data-testid="stSidebar"] [data-testid="stDateInput"] [data-baseweb="input"] {
    border: 1px solid #6b7280 !important;
    border-radius: 8px !important;
}

/* 3) Sana yoziladigan joyning matni */
[data-testid="stSidebar"] [data-testid="stDateInput"] input {
    color: white !important;
    -webkit-text-fill-color: white !important;
    caret-color: white !important;
    border: none !important;
    outline: none !important;
    font-weight: 600 !important;
}

/* 4) Placeholder */
[data-testid="stSidebar"] [data-testid="stDateInput"] input::placeholder {
    color: white !important;
    -webkit-text-fill-color: white !important;
    opacity: 1 !important;
}

/* 5) Sana labeli */
[data-testid="stSidebar"] [data-testid="stDateInput"] label {
    color: white !important;
    font-weight: 600 !important;
}


/* =====================================================
   KALENDAR TUGMASI
   ===================================================== */

[data-testid="stSidebar"] [data-testid="stDateInput"] button {
    background-color: #2a3752 !important;
    color: white !important;
    border: none !important;
}

/* Kalendar ikonkasini oq qilish */
[data-testid="stSidebar"] [data-testid="stDateInput"] button svg {
    color: white !important;
    fill: white !important;
}


/* =====================================================
   KPI KARTALAR
   ===================================================== */

div[data-testid="stMetric"] {
    background-color: white;
    padding: 18px;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}


/* =====================================================
   HEADER
   ===================================================== */

.main-title {
    font-size: 32px;
    font-weight: 700;
    color: #172033;
    margin-bottom: 5px;
}

.subtitle {
    font-size: 15px;
    color: #6b7280;
    margin-bottom: 25px;
}


/* =====================================================
   SECTION TITLE
   ===================================================== */

.section-title {
    font-size: 22px;
    font-weight: 700;
    color: #172033;
    margin-top: 25px;
    margin-bottom: 15px;
}


/* =====================================================
   FOOTER
   ===================================================== */

.footer {
    text-align: center;
    color: #6b7280;
    font-size: 13px;
    margin-top: 40px;
    padding: 20px;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# EXCEL FILE
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(BASE_DIR, "Valyuta.xlsx")


# =========================================================
# LOAD DATA
# =========================================================

@st.cache_data
def load_data():

    df = pd.read_excel(FILE_PATH)

    # Ustun nomlarini tozalash
    df.columns = (
        df.columns
        .astype(str)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )

    # Excel ustunlarini standart nomlarga o'tkazish
    rename_map = {}

    for col in df.columns:

        if col == "Sana":
            rename_map[col] = "Date"

        elif "AQSH dollari" in col:
            rename_map[col] = "USD"

        elif "EVRO" in col:
            rename_map[col] = "EUR"

    df = df.rename(columns=rename_map)

    # Sana
    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    # USD
    df["USD"] = pd.to_numeric(
        df["USD"],
        errors="coerce"
    )

    # EUR
    df["EUR"] = pd.to_numeric(
        df["EUR"],
        errors="coerce"
    )

    # Bo'sh sanalarni olib tashlash
    df = df.dropna(
        subset=["Date"]
    )

    # Sana bo'yicha tartiblash
    df = df.sort_values("Date")

    return df


df = load_data()


# =========================================================
# SIDEBAR
# =========================================================

st.sidebar.markdown(
    "## 💱 Valyuta Risk Dashboard"
)

st.sidebar.markdown("---")


# =========================================================
# VALYUTA TANLASH
# =========================================================

currency_option = st.sidebar.selectbox(
    "Valyutani tanlang",
    [
        "USD + EUR",
        "Faqat USD",
        "Faqat EUR"
    ]
)


# =========================================================
# SANA ORALIG'I
# =========================================================

min_date = df["Date"].min().date()
max_date = df["Date"].max().date()


start_date = st.sidebar.date_input(
    "Boshlanish sanasi",
    value=min_date,
    min_value=min_date,
    max_value=max_date
)


end_date = st.sidebar.date_input(
    "Tugash sanasi",
    value=max_date,
    min_value=min_date,
    max_value=max_date
)


# =========================================================
# SANA TEKSHIRISH
# =========================================================

if start_date > end_date:

    st.error(
        "Boshlanish sanasi tugash sanasidan katta bo'lishi mumkin emas."
    )

    st.stop()


# =========================================================
# TANLANGAN DAVR
# =========================================================

selected_df = df[
    (df["Date"].dt.date >= start_date) &
    (df["Date"].dt.date <= end_date)
].copy()


# =========================================================
# HEADER
# =========================================================

st.markdown(
    '<div class="main-title">Valyuta Risk Dashboard</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'USD va EUR kurslarining tanlangan davr bo‘yicha tahlili'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# VALYUTA DASHBOARD FUNKSIYASI
# =========================================================

def show_currency_dashboard(
    data,
    currency,
    currency_name
):

    if data.empty:

        st.warning(
            f"{currency_name} bo‘yicha ma'lumot topilmadi."
        )

        return


    # -----------------------------------------------------
    # BOSHLANG'ICH VA YAKUNIY KURS
    # -----------------------------------------------------

    start_value = data[currency].iloc[0]

    end_value = data[currency].iloc[-1]

    change = (
        (end_value - start_value)
        / start_value
        * 100
    )


    # -----------------------------------------------------
    # MAXIMUM VA MINIMUM
    # -----------------------------------------------------

    max_value = data[currency].max()

    min_value = data[currency].min()

    max_date = data.loc[
        data[currency].idxmax(),
        "Date"
    ]

    min_date = data.loc[
        data[currency].idxmin(),
        "Date"
    ]


    # -----------------------------------------------------
    # NOMI
    # -----------------------------------------------------

    st.markdown(
        f'<div class="section-title">{currency_name}</div>',
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # KPI
    # -----------------------------------------------------

    col1, col2, col3, col4 = st.columns(4)


    with col1:

        st.metric(
            "Boshlang‘ich kurs",
            f"{start_value:,.2f}"
        )


    with col2:

        st.metric(
            "Yakuniy kurs",
            f"{end_value:,.2f}",
            f"{change:+.2f}%"
        )


    with col3:

        st.metric(
            "Maksimum",
            f"{max_value:,.2f}",
            max_date.strftime("%d.%m.%Y")
        )


    with col4:

        st.metric(
            "Minimum",
            f"{min_value:,.2f}",
            min_date.strftime("%d.%m.%Y")
        )


    # -----------------------------------------------------
    # GRAFIK
    # -----------------------------------------------------

    fig = go.Figure()


    fig.add_trace(
        go.Scatter(
            x=data["Date"],
            y=data[currency],
            mode="lines",
            name=currency_name,
            line=dict(
                shape="spline",
                smoothing=1.3,
                width=3
            )
        )
    )


    fig.update_layout(
        height=430,
        template="plotly_white",

        margin=dict(
            l=20,
            r=20,
            t=30,
            b=20
        ),

        xaxis_title="Sana",
        yaxis_title="Kurs",

        hovermode="x unified"
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


    # -----------------------------------------------------
    # JADVAL
    # -----------------------------------------------------

    table_df = data[
        ["Date", currency]
    ].copy()


    table_df["Date"] = table_df[
        "Date"
    ].dt.strftime("%d.%m.%Y")


    table_df.columns = [
        "Sana",
        "Kurs"
    ]


    st.dataframe(
        table_df,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# VALYUTANI KO'RSATISH
# =========================================================

if currency_option == "USD + EUR":

    show_currency_dashboard(
        selected_df,
        "USD",
        "AQSH dollari (USD)"
    )

    st.markdown("---")

    show_currency_dashboard(
        selected_df,
        "EUR",
        "EVRO (EUR)"
    )

elif currency_option == "Faqat USD":

    show_currency_dashboard(
        selected_df,
        "USD",
        "AQSH dollari (USD)"
    )

else:

    show_currency_dashboard(
        selected_df,
        "EUR",
        "EVRO (EUR)"
    )


# =========================================================
# 2025 VS 2026
# =========================================================

st.markdown("---")

st.markdown(
    '<div class="section-title">'
    '2025 vs 2026 taqqoslash'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# OYLIK MA'LUMOT
# =========================================================

comparison_df = df.copy()

comparison_df["Year"] = comparison_df[
    "Date"
].dt.year

comparison_df["Month"] = comparison_df[
    "Date"
].dt.month


monthly = (
    comparison_df
    .groupby(
        ["Year", "Month"]
    )[["USD", "EUR"]]
    .mean()
    .reset_index()
)


# =========================================================
# TAQQOSLASH GRAFIGI
# =========================================================

def comparison_chart(
    currency,
    name
):

    fig = go.Figure()


    # -----------------------------------------------------
    # 2025
    # -----------------------------------------------------

    data_2025 = monthly[
        monthly["Year"] == 2025
    ]


    fig.add_trace(
        go.Scatter(
            x=data_2025["Month"],
            y=data_2025[currency],
            mode="lines+markers",
            name=f"{name} — 2025",

            line=dict(
                shape="spline",
                smoothing=1.3,
                width=3
            )
        )
    )


    # -----------------------------------------------------
    # 2026
    # -----------------------------------------------------

    data_2026 = monthly[
        monthly["Year"] == 2026
    ]


    fig.add_trace(
        go.Scatter(
            x=data_2026["Month"],
            y=data_2026[currency],
            mode="lines+markers",
            name=f"{name} — 2026",

            line=dict(
                shape="spline",
                smoothing=1.3,
                width=3
            )
        )
    )


    # -----------------------------------------------------
    # GRAFIK SOZLAMALARI
    # -----------------------------------------------------

    fig.update_layout(

        height=430,

        template="plotly_white",

        xaxis=dict(

            title="Oy",

            tickmode="array",

            tickvals=list(range(1, 13)),

            ticktext=[
                "Yan",
                "Fev",
                "Mar",
                "Apr",
                "May",
                "Iyun",
                "Iyul",
                "Avg",
                "Sen",
                "Okt",
                "Noy",
                "Dek"
            ]
        ),

        yaxis_title="Kurs",

        hovermode="x unified",

        margin=dict(
            l=20,
            r=20,
            t=30,
            b=20
        )
    )


    st.plotly_chart(
        fig,
        use_container_width=True
    )


# =========================================================
# TAQQOSLASHNI KO'RSATISH
# =========================================================

if currency_option == "USD + EUR":

    comparison_chart(
        "USD",
        "USD"
    )

    comparison_chart(
        "EUR",
        "EUR"
    )

elif currency_option == "Faqat USD":

    comparison_chart(
        "USD",
        "USD"
    )

else:

    comparison_chart(
        "EUR",
        "EUR"
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer">
        Valyuta Risk Dashboard | USD / EUR
    </div>
    """,
    unsafe_allow_html=True
)