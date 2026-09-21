import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Valyuta Risk Dashboard",
    page_icon="💱",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>

    .stApp {
        background-color: #f4f6f9;
    }

    .main-title {
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .sub-title {
        color: #6b7280;
        font-size: 15px;
        margin-bottom: 25px;
    }

    .section-title {
        font-size: 22px;
        font-weight: 700;
        margin-top: 30px;
        margin-bottom: 15px;
    }

    section[data-testid="stSidebar"] {
        background-color: #172033;
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    div[data-testid="stMetric"] {
        background-color: white;
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.06);
        border: 1px solid #e5e7eb;
    }

    div[data-testid="stMetricLabel"] {
        font-size: 14px;
    }

    div[data-testid="stMetricValue"] {
        font-size: 25px;
        font-weight: 700;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# FILE PATH
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(
    BASE_DIR,
    "Valyuta.xlsx"
)


# =========================================================
# LOAD EXCEL
# =========================================================
@st.cache_data
def load_data():

    if not os.path.exists(FILE_PATH):

        st.error(
            f"Excel fayl topilmadi:\n\n{FILE_PATH}"
        )

        st.stop()

    df = pd.read_excel(
        FILE_PATH
    )

    # -----------------------------------------------------
    # Normalize column names
    # -----------------------------------------------------
    df.columns = (
        df.columns
        .astype(str)
        .str.replace(
            "\u00a0",
            " ",
            regex=False
        )
        .str.replace(
            r"\s+",
            " ",
            regex=True
        )
        .str.strip()
    )

    # -----------------------------------------------------
    # Rename columns
    # -----------------------------------------------------
    column_map = {}

    for col in df.columns:

        clean = (
            str(col)
            .replace(
                "\u00a0",
                " "
            )
            .strip()
        )

        if clean == "Sana":

            column_map[col] = "Date"

        elif "AQSH dollari" in clean:

            column_map[col] = "USD"

        elif "EVRO" in clean:

            column_map[col] = "EUR"

    df = df.rename(
        columns=column_map
    )

    # -----------------------------------------------------
    # Check columns
    # -----------------------------------------------------
    required_columns = [
        "Date",
        "USD",
        "EUR"
    ]

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:

        st.error(
            "Excel faylda quyidagi ustunlar topilmadi: "
            + ", ".join(missing)
        )

        st.stop()

    # -----------------------------------------------------
    # Convert types
    # -----------------------------------------------------
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

    # -----------------------------------------------------
    # Remove invalid rows
    # -----------------------------------------------------
    df = df.dropna(
        subset=["Date"]
    )

    df = df.sort_values(
        "Date"
    ).reset_index(
        drop=True
    )

    return df


data = load_data()


# =========================================================
# DATA RANGE
# =========================================================
min_data_date = data["Date"].min().date()
max_data_date = data["Date"].max().date()


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:

    st.markdown(
        "## 💱 Valyuta Risk Dashboard"
    )

    st.markdown("---")

    # =====================================================
    # CURRENCY
    # =====================================================
    currency_choice = st.selectbox(
        "Valyuta",
        [
            "USD + EUR",
            "Faqat USD",
            "Faqat EUR"
        ]
    )

    st.markdown("---")

    # =====================================================
    # DATE RANGE
    # =====================================================
    st.markdown(
        "### 📅 Sana oralig‘i"
    )

    start_date = st.date_input(
        "Boshlanish sanasi",
        value=min_data_date,
        min_value=min_data_date,
        max_value=max_data_date,
        format="DD.MM.YYYY"
    )

    end_date = st.date_input(
        "Tugash sanasi",
        value=max_data_date,
        min_value=min_data_date,
        max_value=max_data_date,
        format="DD.MM.YYYY"
    )

    # =====================================================
    # VALIDATE DATES
    # =====================================================
    if start_date > end_date:

        st.error(
            "Boshlanish sanasi tugash sanasidan "
            "katta bo‘lishi mumkin emas."
        )

        st.stop()

    # =====================================================
    # INFORMATION
    # =====================================================
    st.markdown("---")

    st.markdown(
        "### 📊 Tanlangan davr"
    )

    st.write(
        f"**{start_date.strftime('%d.%m.%Y')}**"
    )

    st.write(
        f"dan"
    )

    st.write(
        f"**{end_date.strftime('%d.%m.%Y')}**"
    )

    # =====================================================
    # AVAILABLE EXCEL RANGE
    # =====================================================
    st.markdown("---")

    st.caption(
        "Excel ma'lumotlari:"
    )

    st.caption(
        f"{min_data_date.strftime('%d.%m.%Y')}"
        f" — "
        f"{max_data_date.strftime('%d.%m.%Y')}"
    )

    st.caption(
        f"Jami: {len(data):,} ta yozuv"
    )


# =========================================================
# SELECTED PERIOD DATA
# =========================================================
selected_data = data[
    (data["Date"].dt.date >= start_date)
    &
    (data["Date"].dt.date <= end_date)
].copy()


# =========================================================
# HEADER
# =========================================================
st.markdown(
    '<div class="main-title">'
    '💱 Valyuta Risk Dashboard'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="sub-title">
    USD va EUR kurslari bo‘yicha tarixiy tahlil,
    tanlangan sana oralig‘i va yillik taqqoslash
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# SELECTED PERIOD INFO
# =========================================================
if selected_data.empty:

    st.warning(
        "Tanlangan sana oralig‘ida ma'lumot mavjud emas."
    )

else:

    st.info(
        f"📅 Tanlangan davr: "
        f"**{start_date.strftime('%d.%m.%Y')}**"
        f" → "
        f"**{end_date.strftime('%d.%m.%Y')}**"
        f" | "
        f"Ma'lumotlar soni: "
        f"**{len(selected_data):,}**"
    )


# =========================================================
# CURRENCY DASHBOARD FUNCTION
# =========================================================
def show_currency_dashboard(
    df,
    currency,
    currency_name
):

    if df.empty:

        st.warning(
            f"{currency_name} bo‘yicha "
            "tanlangan davrda ma'lumot yo‘q."
        )

        return

    # =====================================================
    # CALCULATIONS
    # =====================================================
    start_value = df[currency].iloc[0]

    end_value = df[currency].iloc[-1]

    if start_value != 0:

        change = (
            (end_value - start_value)
            / start_value
            * 100
        )

    else:

        change = 0

    min_value = df[currency].min()

    max_value = df[currency].max()

    min_date = df.loc[
        df[currency].idxmin(),
        "Date"
    ]

    max_date = df.loc[
        df[currency].idxmax(),
        "Date"
    ]

    # =====================================================
    # SECTION
    # =====================================================
    st.markdown(
        f'<div class="section-title">'
        f'{currency_name}'
        f'</div>',
        unsafe_allow_html=True
    )

    # =====================================================
    # KPI
    # =====================================================
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:

        st.metric(
            "Boshlang‘ich kurs",
            f"{start_value:,.2f}"
        )

    with col2:

        st.metric(
            "Oxirgi kurs",
            f"{end_value:,.2f}"
        )

    with col3:

        st.metric(
            "O‘zgarish",
            f"{change:+.2f}%"
        )

    with col4:

        st.metric(
            "Minimum",
            f"{min_value:,.2f}",
            help=(
                "Sana: "
                + min_date.strftime("%d.%m.%Y")
            )
        )

    with col5:

        st.metric(
            "Maksimum",
            f"{max_value:,.2f}",
            help=(
                "Sana: "
                + max_date.strftime("%d.%m.%Y")
            )
        )

    # =====================================================
    # CHART
    # =====================================================
    fig = go.Figure()

    fig.add_trace(
        go.Scatter(

            x=df["Date"],

            y=df[currency],

            mode="lines+markers",

            name=currency_name,

            # Smooth winding line
            line=dict(
                width=4,
                shape="spline",
                smoothing=1.3
            ),

            marker=dict(
                size=6
            ),

            hovertemplate=(
                "%{x|%d.%m.%Y}"
                "<br>"
                "Kurs: %{y:,.2f}"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(

        title=(
            f"{currency_name}"
            f" | "
            f"{start_date.strftime('%d.%m.%Y')}"
            f" — "
            f"{end_date.strftime('%d.%m.%Y')}"
        ),

        height=500,

        template="plotly_white",

        hovermode="x unified",

        margin=dict(
            l=20,
            r=20,
            t=60,
            b=20
        ),

        xaxis=dict(
            title="Sana",
            showgrid=True
        ),

        yaxis=dict(
            title="UZS",
            tickformat=","
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
        fig,
        use_container_width=True
    )

    # =====================================================
    # TABLE
    # =====================================================
    with st.expander(
        f"📋 {currency_name} ma'lumotlari"
    ):

        table = df[
            ["Date", currency]
        ].copy()

        table["Date"] = table[
            "Date"
        ].dt.strftime(
            "%d.%m.%Y"
        )

        table.columns = [
            "Sana",
            "Kurs"
        ]

        table["Kurs"] = table[
            "Kurs"
        ].map(
            lambda x: f"{x:,.2f}"
        )

        st.dataframe(
            table,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# SHOW SELECTED CURRENCY
# =========================================================
if currency_choice == "USD + EUR":

    show_currency_dashboard(
        selected_data,
        "USD",
        "🇺🇸 AQSH dollari (USD)"
    )

    show_currency_dashboard(
        selected_data,
        "EUR",
        "🇪🇺 Yevro (EUR)"
    )

elif currency_choice == "Faqat USD":

    show_currency_dashboard(
        selected_data,
        "USD",
        "🇺🇸 AQSH dollari (USD)"
    )

elif currency_choice == "Faqat EUR":

    show_currency_dashboard(
        selected_data,
        "EUR",
        "🇪🇺 Yevro (EUR)"
    )


# =========================================================
# FULL YEAR DATA
# IMPORTANT:
# This section does NOT depend on selected date range.
# =========================================================
data_2025 = data[
    data["Date"].dt.year == 2025
].copy()

data_2026 = data[
    data["Date"].dt.year == 2026
].copy()


# =========================================================
# MONTHLY COMPARISON DATA
# =========================================================
def prepare_comparison_data(
    df,
    currency
):

    if df.empty:

        return pd.DataFrame(
            columns=[
                "Month",
                "Value"
            ]
        )

    temp = df.copy()

    temp["Month"] = (
        temp["Date"].dt.month
    )

    temp = temp.sort_values(
        "Date"
    )

    # Har oyning oxirgi mavjud kursi
    monthly = (
        temp
        .groupby(
            "Month",
            as_index=False
        )
        .last()
    )

    monthly = monthly[
        [
            "Month",
            currency
        ]
    ]

    monthly = monthly.rename(
        columns={
            currency: "Value"
        }
    )

    return monthly


# =========================================================
# YEAR COMPARISON CHART
# =========================================================
def show_year_comparison_chart(
    data_2025,
    data_2026,
    currency,
    currency_name
):

    monthly_2025 = (
        prepare_comparison_data(
            data_2025,
            currency
        )
    )

    monthly_2026 = (
        prepare_comparison_data(
            data_2026,
            currency
        )
    )

    month_labels = {

        1: "Yan",
        2: "Fev",
        3: "Mar",
        4: "Apr",
        5: "May",
        6: "Iyun",
        7: "Iyul",
        8: "Avg",
        9: "Sen",
        10: "Okt",
        11: "Noy",
        12: "Dek"
    }

    all_month_labels = [
        month_labels[m]
        for m in range(1, 13)
    ]

    fig = go.Figure()

    # =====================================================
    # 2025
    # =====================================================
    if not monthly_2025.empty:

        fig.add_trace(
            go.Scatter(

                x=[
                    month_labels[m]
                    for m in monthly_2025["Month"]
                ],

                y=monthly_2025["Value"],

                mode="lines+markers",

                name="2025",

                line=dict(
                    width=4,
                    shape="spline",
                    smoothing=1.3
                ),

                marker=dict(
                    size=7
                ),

                hovertemplate=(
                    "2025"
                    "<br>"
                    "Oy: %{x}"
                    "<br>"
                    "Kurs: %{y:,.2f}"
                    "<extra></extra>"
                )
            )
        )

    # =====================================================
    # 2026
    # =====================================================
    if not monthly_2026.empty:

        fig.add_trace(
            go.Scatter(

                x=[
                    month_labels[m]
                    for m in monthly_2026["Month"]
                ],

                y=monthly_2026["Value"],

                mode="lines+markers",

                name="2026",

                line=dict(
                    width=4,
                    shape="spline",
                    smoothing=1.3
                ),

                marker=dict(
                    size=7
                ),

                connectgaps=False,

                hovertemplate=(
                    "2026"
                    "<br>"
                    "Oy: %{x}"
                    "<br>"
                    "Kurs: %{y:,.2f}"
                    "<extra></extra>"
                )
            )
        )

    # =====================================================
    # LAYOUT
    # =====================================================
    fig.update_layout(

        title=(
            f"{currency_name}"
            f" — 2025 vs 2026"
        ),

        height=500,

        template="plotly_white",

        hovermode="x unified",

        margin=dict(
            l=20,
            r=20,
            t=60,
            b=20
        ),

        xaxis=dict(

            title="Oy",

            categoryorder="array",

            categoryarray=all_month_labels,

            range=[
                -0.2,
                11.2
            ]
        ),

        yaxis=dict(
            title="UZS",
            tickformat=","
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
        fig,
        use_container_width=True
    )


# =========================================================
# YEAR COMPARISON TABLE
# =========================================================
def show_comparison_table(
    data_2025,
    data_2026,
    currency,
    currency_name
):

    monthly_2025 = (
        prepare_comparison_data(
            data_2025,
            currency
        )
    )

    monthly_2026 = (
        prepare_comparison_data(
            data_2026,
            currency
        )
    )

    month_labels = {

        1: "Yanvar",
        2: "Fevral",
        3: "Mart",
        4: "Aprel",
        5: "May",
        6: "Iyun",
        7: "Iyul",
        8: "Avgust",
        9: "Sentabr",
        10: "Oktabr",
        11: "Noyabr",
        12: "Dekabr"
    }

    comparison = pd.DataFrame({
        "Oy": list(
            month_labels.values()
        )
    })

    # =====================================================
    # 2025
    # =====================================================
    if not monthly_2025.empty:

        temp_2025 = monthly_2025.copy()

        temp_2025["Oy"] = temp_2025[
            "Month"
        ].map(
            month_labels
        )

        temp_2025 = temp_2025[
            ["Oy", "Value"]
        ]

        temp_2025 = temp_2025.rename(
            columns={
                "Value": "2025"
            }
        )

        comparison = comparison.merge(
            temp_2025,
            on="Oy",
            how="left"
        )

    else:

        comparison["2025"] = None

    # =====================================================
    # 2026
    # =====================================================
    if not monthly_2026.empty:

        temp_2026 = monthly_2026.copy()

        temp_2026["Oy"] = temp_2026[
            "Month"
        ].map(
            month_labels
        )

        temp_2026 = temp_2026[
            ["Oy", "Value"]
        ]

        temp_2026 = temp_2026.rename(
            columns={
                "Value": "2026"
            }
        )

        comparison = comparison.merge(
            temp_2026,
            on="Oy",
            how="left"
        )

    else:

        comparison["2026"] = None

    # =====================================================
    # Format values
    # =====================================================
    comparison["2025"] = comparison[
        "2025"
    ].apply(
        lambda x:
        f"{x:,.2f}"
        if pd.notna(x)
        else "—"
    )

    comparison["2026"] = comparison[
        "2026"
    ].apply(
        lambda x:
        f"{x:,.2f}"
        if pd.notna(x)
        else "—"
    )

    # =====================================================
    # SHOW TABLE
    # =====================================================
    with st.expander(
        f"📋 {currency_name}"
        f" — oylik taqqoslash jadvali"
    ):

        st.dataframe(
            comparison,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# YEAR COMPARISON SECTION
# =========================================================
st.markdown("---")

st.markdown(
    '<div class="section-title">'
    '📊 2025 vs 2026 — Yillik taqqoslash'
    '</div>',
    unsafe_allow_html=True
)

st.info(
    "Bu bo‘lim tanlangan sana oralig‘iga "
    "bog‘liq emas. 2025 va 2026 yillarning "
    "mavjud ma'lumotlari to‘liq taqqoslanadi. "
    "Ma'lumot mavjud bo‘lmagan oylar bo‘sh qoladi."
)


# =========================================================
# USD COMPARISON
# =========================================================
if currency_choice in [
    "USD + EUR",
    "Faqat USD"
]:

    show_year_comparison_chart(
        data_2025,
        data_2026,
        "USD",
        "🇺🇸 AQSH dollari (USD)"
    )

    show_comparison_table(
        data_2025,
        data_2026,
        "USD",
        "🇺🇸 AQSH dollari (USD)"
    )


# =========================================================
# EUR COMPARISON
# =========================================================
if currency_choice in [
    "USD + EUR",
    "Faqat EUR"
]:

    show_year_comparison_chart(
        data_2025,
        data_2026,
        "EUR",
        "🇪🇺 Yevro (EUR)"
    )

    show_comparison_table(
        data_2025,
        data_2026,
        "EUR",
        "🇪🇺 Yevro (EUR)"
    )


# =========================================================
# FOOTER
# =========================================================
st.markdown("---")

st.caption(
    "Valyuta Risk Dashboard • "
    "Excel ma'lumotlari asosida"
)
