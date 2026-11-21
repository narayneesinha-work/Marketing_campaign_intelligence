
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import date

# =========================================================
# 1. APP CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Marketing Campaign Intelligence",
    page_icon="📊",
    layout="wide"
)

st.title("Marketing Campaign Intelligence")
st.caption(
    "Campaign performance monitoring, funnel analysis "
    "and evidence-based diagnostics."
)

KEYS = ["Date", "Campaign", "Ad Set", "Ad"]

SHEET_HINTS = {
    "upper": ["meta_upper", "upper funnel", "upper"],
    "lower": ["lower funnel", "lower"],
    "audience": ["placement", "location", "age level", "audience"]
}

# =========================================================
# 2. HELPER FUNCTIONS
# =========================================================

def clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_sheet_role(sheet_name, df):
    name = sheet_name.lower().replace("_", " ")

    for role, hints in SHEET_HINTS.items():
        if any(hint in name for hint in hints):
            return role

    cols = set(df.columns)

    if {"Spend", "Impressions", "Clicks"}.issubset(cols):
        return "upper"

    if {"Purchases", "Revenue"}.intersection(cols):
        return "lower"

    if {"Placement", "Location", "Age"}.intersection(cols):
        return "audience"

    return None


def safe_divide(numerator, denominator, multiplier=1):
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")

    result = numerator / denominator.replace(0, float("nan"))
    return result * multiplier


def make_excel(master, campaign_report, audit):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        master.to_excel(
            writer, sheet_name="Master Data", index=False
        )
        campaign_report.to_excel(
            writer, sheet_name="Campaign Analysis", index=False
        )
        audit.to_excel(
            writer, sheet_name="Data Audit", index=False
        )

    output.seek(0)
    return output


def format_number(value, currency=False):
    if pd.isna(value):
        return "N/A"

    if currency:
        return f"₹{value:,.2f}"

    return f"{value:,.0f}"


# =========================================================
# 3. WORKBOOK PROCESSING
# =========================================================

def process_workbook(uploaded_file):

    excel = pd.ExcelFile(uploaded_file)

    dataframes = {}
    audit_rows = []

    ignored_sheets = []

    for sheet in excel.sheet_names:

        df = clean_columns(
            pd.read_excel(excel, sheet_name=sheet)
        )

        role = find_sheet_role(sheet, df)

        if role is None:
            ignored_sheets.append(sheet)
            continue

        missing_keys = [
            col for col in KEYS if col not in df.columns
        ]

        if missing_keys:
            raise ValueError(
                f"Sheet '{sheet}' is missing key columns: "
                f"{missing_keys}"
            )

        df["Date"] = pd.to_datetime(
            df["Date"], errors="coerce"
        )

        invalid_dates = int(df["Date"].isna().sum())

        df = df.dropna(subset=KEYS).copy()

        duplicate_count = int(
            df.duplicated(subset=KEYS).sum()
        )

        audit_rows.append({
            "Sheet": sheet,
            "Detected type": role,
            "Rows read": len(df),
            "Invalid dates": invalid_dates,
            "Duplicate merge keys": duplicate_count,
            "Status": (
                "Valid" if duplicate_count == 0
                else "Needs review"
            )
        })

        if duplicate_count > 0:
            raise ValueError(
                f"Sheet '{sheet}' contains {duplicate_count} "
                "duplicate Date/Campaign/Ad Set/Ad combinations. "
                "Please resolve these before merging. "
                "A many-to-many merge could multiply spend and results."
            )

        dataframes[role] = (sheet, df)

    if not dataframes:
        raise ValueError(
            "No recognised Meta Ads data sheets were found."
        )

    # Use upper funnel as the preferred base.
    # If it is unavailable, use another available report.

    base_role = (
        "upper" if "upper" in dataframes
        else next(iter(dataframes))
    )

    base_sheet, master = dataframes[base_role]
    master = master.copy()

    merge_audit = []

    for role, (sheet, incoming) in dataframes.items():

        if role == base_role:
            continue

        overlapping = [
            col for col in incoming.columns
            if col not in KEYS and col in master.columns
        ]

        if overlapping:
            raise ValueError(
                f"Overlapping columns found in '{sheet}': "
                f"{overlapping}. Please review the source sheets "
                "before merging to avoid ambiguous metrics."
            )

        before_rows = len(master)

        merged = master.merge(
            incoming,
            on=KEYS,
            how="outer",
            validate="one_to_one",
            indicator=True
        )

        left_only = int(
            (merged["_merge"] == "left_only").sum()
        )
        right_only = int(
            (merged["_merge"] == "right_only").sum()
        )
        matched = int(
            (merged["_merge"] == "both").sum()
        )

        merge_audit.append({
            "Merged sheet": sheet,
            "Matched records": matched,
            "Only in existing data": left_only,
            "Only in incoming data": right_only,
            "Rows before merge": before_rows,
            "Rows after merge": len(merged)
        })

        master = merged.drop(columns=["_merge"])

    # Convert available metrics to numeric.
    key_columns = set(KEYS)

    for col in master.columns:
        if col not in key_columns:
            if col not in [
                "Platform", "Placement", "Location", "Age"
            ]:
                master[col] = pd.to_numeric(
                    master[col], errors="coerce"
                )

    master["Month"] = master["Date"].dt.strftime("%Y-%m")

    audit_df = pd.DataFrame(audit_rows)

    merge_audit_df = pd.DataFrame(merge_audit)

    return (
        master,
        audit_df,
        merge_audit_df,
        ignored_sheets
    )


# =========================================================
# 4. CALCULATE PERFORMANCE METRICS
# =========================================================

def build_campaign_report(df):

    group_columns = ["Campaign"]

    sum_metrics = [
        "Spend",
        "Impressions",
        "Clicks",
        "Reach",
        "Landing Page Views",
        "Add To Cart",
        "Checkout Initiated",
        "Purchases",
        "Revenue",
        "Views",
        "ThruPlay",
        "3 Second Views"
    ]

    available_metrics = [
        col for col in sum_metrics if col in df.columns
    ]

    report = df.groupby(
        group_columns, dropna=False
    )[available_metrics].sum(min_count=1).reset_index()

    if {"Clicks", "Impressions"}.issubset(report.columns):
        report["CTR (%)"] = safe_divide(
            report["Clicks"], report["Impressions"], 100
        )

    if {"Spend", "Clicks"}.issubset(report.columns):
        report["CPC"] = safe_divide(
            report["Spend"], report["Clicks"]
        )

    if {"Spend", "Impressions"}.issubset(report.columns):
        report["CPM"] = safe_divide(
            report["Spend"], report["Impressions"], 1000
        )

    if {"Impressions", "Reach"}.issubset(report.columns):
        report["Frequency"] = safe_divide(
            report["Impressions"], report["Reach"]
        )

    if {"Landing Page Views", "Clicks"}.issubset(report.columns):
        report["LPV Rate (%)"] = safe_divide(
            report["Landing Page Views"],
            report["Clicks"], 100
        )

    if {"Add To Cart", "Landing Page Views"}.issubset(report.columns):
        report["ATC Rate (%)"] = safe_divide(
            report["Add To Cart"],
            report["Landing Page Views"], 100
        )

    if {"Checkout Initiated", "Add To Cart"}.issubset(report.columns):
        report["Checkout Rate (%)"] = safe_divide(
            report["Checkout Initiated"],
            report["Add To Cart"], 100
        )

    if {"Purchases", "Checkout Initiated"}.issubset(report.columns):
        report["Purchase Rate (%)"] = safe_divide(
            report["Purchases"],
            report["Checkout Initiated"], 100
        )

    if {"Purchases", "Spend"}.issubset(report.columns):
        report["CPA"] = safe_divide(
            report["Spend"], report["Purchases"]
        )

    if {"Revenue", "Spend"}.issubset(report.columns):
        report["ROAS"] = safe_divide(
            report["Revenue"], report["Spend"]
        )

    return report


# =========================================================
# 5. SIDEBAR: DATA UPLOAD AND NAVIGATION
# =========================================================

st.sidebar.title("Navigation")

platform = st.sidebar.selectbox(
    "Select platform",
    ["Meta Ads"]
)

uploaded_file = st.sidebar.file_uploader(
    "Upload campaign Excel workbook",
    type=["xlsx"]
)

if uploaded_file is None:
    st.info(
        "Upload your Meta Ads Excel workbook to open "
        "the campaign dashboard."
    )

    st.markdown("""
    ### What this application does

    - Consolidates available campaign reports.
    - Builds a campaign performance dashboard.
    - Analyses funnel conversion and drop-offs.
    - Identifies potential performance issues.
    - Provides downloadable analysis and audit data.
    """)

    st.stop()


# =========================================================
# 6. PROCESS THE UPLOADED WORKBOOK
# =========================================================

try:
    master_df, sheet_audit, merge_audit, ignored_sheets = (
        process_workbook(uploaded_file)
    )

    campaign_report = build_campaign_report(master_df)

except Exception as error:
    st.error(f"Processing error: {error}")
    st.stop()


if master_df.empty:
    st.error("The uploaded workbook contains no usable records.")
    st.stop()


# =========================================================
# 7. GLOBAL FILTERS
# =========================================================

st.sidebar.divider()
st.sidebar.subheader("Analysis filters")

min_date = master_df["Date"].min().date()
max_date = master_df["Date"].max().date()

date_range = st.sidebar.date_input(
    "Reporting date range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

if isinstance(date_range, tuple) and len(date_range) == 2:

    start_date, end_date = date_range

    filtered_df = master_df[
        (master_df["Date"].dt.date >= start_date)
        & (master_df["Date"].dt.date <= end_date)
    ].copy()

else:
    filtered_df = master_df.copy()


campaign_options = sorted(
    filtered_df["Campaign"].dropna().unique().tolist()
)

selected_campaigns = st.sidebar.multiselect(
    "Campaigns",
    campaign_options,
    default=campaign_options
)

filtered_df = filtered_df[
    filtered_df["Campaign"].isin(selected_campaigns)
]


if "Ad Set" in filtered_df.columns:

    adset_options = sorted(
        filtered_df["Ad Set"].dropna().unique().tolist()
    )

    selected_adsets = st.sidebar.multiselect(
        "Ad Sets",
        adset_options,
        default=adset_options
    )

    filtered_df = filtered_df[
        filtered_df["Ad Set"].isin(selected_adsets)
    ]


if "Ad" in filtered_df.columns:

    ad_options = sorted(
        filtered_df["Ad"].dropna().unique().tolist()
    )

    selected_ads = st.sidebar.multiselect(
        "Ads",
        ad_options,
        default=ad_options
    )

    filtered_df = filtered_df[
        filtered_df["Ad"].isin(selected_ads)
    ]


if filtered_df.empty:
    st.warning("No records match the selected filters.")
    st.stop()


analysis_report = build_campaign_report(filtered_df)


# =========================================================
# 8. PAGE NAVIGATION
# =========================================================

page = st.sidebar.radio(
    "Go to page",
    [
        "Performance Dashboard",
        "Funnel & Metric Analysis",
        "Master Data",
        "Data Audit"
    ]
)

st.sidebar.divider()
st.sidebar.caption(
    "Meta Ads analysis module. "
    "Google Ads will be handled separately."
)


# =========================================================
# 9. PERFORMANCE DASHBOARD
# =========================================================

if page == "Performance Dashboard":

    st.header("Performance Dashboard")

    st.caption(
        "Campaign results based on the currently selected filters."
    )

    def total_metric(col):
        if col not in filtered_df.columns:
            return None
        return filtered_df[col].sum(min_count=1)

    spend = total_metric("Spend")
    impressions = total_metric("Impressions")
    clicks = total_metric("Clicks")
    purchases = total_metric("Purchases")
    revenue = total_metric("Revenue")

    kpi_cols = st.columns(5)

    kpis = [
        ("Total Spend", spend, True),
        ("Impressions", impressions, False),
        ("Clicks", clicks, False),
        ("Purchases", purchases, False),
        ("Revenue", revenue, True)
    ]

    for col, (label, value, currency) in zip(kpi_cols, kpis):
        if value is None or pd.isna(value):
            col.metric(label, "N/A")
        else:
            col.metric(
                label,
                format_number(value, currency)
            )

    st.divider()

    # Campaign ranking
    st.subheader("Campaign Performance Ranking")

    ranking_metric = None
    ascending = False

    if "ROAS" in analysis_report.columns and analysis_report["ROAS"].notna().any():
        ranking_metric = "ROAS"
        ascending = False

    elif "CPA" in analysis_report.columns and analysis_report["CPA"].notna().any():
        ranking_metric = "CPA"
        ascending = True

    elif "Purchases" in analysis_report.columns and analysis_report["Purchases"].notna().any():
        ranking_metric = "Purchases"
        ascending = False

    elif "CTR (%)" in analysis_report.columns and analysis_report["CTR (%)"].notna().any():
        ranking_metric = "CTR (%)"
        ascending = False

    if ranking_metric:

        ranked = analysis_report.sort_values(
            ranking_metric,
            ascending=ascending,
            na_position="last"
        ).copy()

        ranked.insert(
            0, "Rank", range(1, len(ranked) + 1)
        )

        st.caption(
            f"Ranked by {ranking_metric}. "
            "This is a metric-based ranking, not a complete "
            "business judgement."
        )

        chart = px.bar(
            ranked,
            x="Campaign",
            y=ranking_metric,
            color=ranking_metric,
            title=f"Campaign ranking by {ranking_metric}",
            text_auto=".2f"
        )

        chart.update_layout(
            xaxis_title="Campaign",
            yaxis_title=ranking_metric,
            xaxis_tickangle=-35
        )

        st.plotly_chart(chart, use_container_width=True)

        st.dataframe(
            ranked,
            use_container_width=True,
            hide_index=True
        )

    else:
        st.info(
            "There is not enough outcome data to rank campaigns "
            "by ROAS, CPA or purchases. Check the available reports."
        )

    st.divider()

    # Spend comparison
    if "Spend" in analysis_report.columns:

        st.subheader("Campaign Spend Distribution")

        spend_chart = px.bar(
            analysis_report.sort_values(
                "Spend", ascending=False
            ),
            x="Campaign",
            y="Spend",
            title="Spend by campaign"
        )

        spend_chart.update_layout(
            xaxis_tickangle=-35,
            yaxis_title="Spend (₹)"
        )

        st.plotly_chart(
            spend_chart,
            use_container_width=True
        )

    # Revenue comparison
    if "Revenue" in analysis_report.columns:

        st.subheader("Campaign Revenue")

        revenue_chart = px.bar(
            analysis_report.sort_values(
                "Revenue", ascending=False
            ),
            x="Campaign",
            y="Revenue",
            title="Revenue by campaign"
        )

        revenue_chart.update_layout(
            xaxis_tickangle=-35,
            yaxis_title="Revenue (₹)"
        )

        st.plotly_chart(
            revenue_chart,
            use_container_width=True
        )

    # Trend
    if "Spend" in filtered_df.columns:

        st.subheader("Performance Trend Over Time")

        trend_metrics = [
            col for col in [
                "Spend", "Clicks", "Purchases", "Revenue"
            ]
            if col in filtered_df.columns
        ]

        if trend_metrics:

            trend = (
                filtered_df
                .groupby("Date")[trend_metrics]
                .sum(min_count=1)
                .reset_index()
                .sort_values("Date")
            )

            selected_trend = st.selectbox(
                "Select trend metric",
                trend_metrics
            )

            trend_chart = px.line(
                trend,
                x="Date",
                y=selected_trend,
                markers=True,
                title=f"{selected_trend} over time"
            )

            st.plotly_chart(
                trend_chart,
                use_container_width=True
            )



    # =========================================================
    # CAMPAIGN PERFORMANCE DIAGNOSIS
    # =========================================================

    st.divider()
    st.header("Campaign Performance Diagnosis")

    st.caption(
        "Evidence-based identification of strong campaigns, "
        "potential underperformance and areas requiring investigation."
    )

    diagnosis = analysis_report.copy()

    if not diagnosis.empty and "Campaign" in diagnosis.columns:

        # Ensure available metrics are numeric
        metric_columns = [
            "Spend", "Revenue", "Purchases", "ROAS",
            "CPA", "CTR", "CPC", "Conversion Rate"
        ]

        for col in metric_columns:
            if col in diagnosis.columns:
                diagnosis[col] = pd.to_numeric(
                    diagnosis[col], errors="coerce"
                )

        # Identify campaigns with usable ROAS
        if "ROAS" in diagnosis.columns and diagnosis["ROAS"].notna().any():

            valid = diagnosis.dropna(subset=["ROAS"]).copy()
            median_roas = valid["ROAS"].median()

            # Best-performing campaigns: above-median ROAS
            best = valid[valid["ROAS"] > median_roas].sort_values(
                "ROAS", ascending=False
            )

            # Potentially weak campaigns: below-median ROAS
            weak = valid[valid["ROAS"] < median_roas].sort_values(
                "ROAS", ascending=True
            )

            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Strong performers")
                st.caption(
                    "Campaigns with ROAS above the selected-data median."
                )

                if not best.empty:
                    display_cols = [
                        c for c in [
                            "Campaign", "Spend", "ROAS",
                            "Purchases", "Revenue"
                        ] if c in best.columns
                    ]
                    st.dataframe(
                        best[display_cols].head(5),
                        width="stretch",
                        hide_index=True
                    )
                else:
                    st.info("No above-median campaigns identified.")

            with col2:
                st.subheader("Potential underperformers")
                st.caption(
                    "Campaigns with ROAS below the selected-data median."
                )

                if not weak.empty:
                    display_cols = [
                        c for c in [
                            "Campaign", "Spend", "ROAS",
                            "Purchases", "Revenue"
                        ] if c in weak.columns
                    ]
                    st.dataframe(
                        weak[display_cols].head(5),
                        width="stretch",
                        hide_index=True
                    )
                else:
                    st.info("No below-median campaigns identified.")

            st.divider()
            st.subheader("High-spend, low-return investigation")

            if "Spend" in valid.columns and valid["Spend"].notna().any():

                median_spend = valid["Spend"].median()

                concern = valid[
                    (valid["Spend"] > median_spend) &
                    (valid["ROAS"] < median_roas)
                ].sort_values("Spend", ascending=False)

                if not concern.empty:
                    st.warning(
                        "These campaigns have above-median spend "
                        "and below-median ROAS. Review them before "
                        "making budget decisions."
                    )

                    display_cols = [
                        c for c in [
                            "Campaign", "Spend", "ROAS",
                            "CPA", "CTR", "Purchases"
                        ] if c in concern.columns
                    ]

                    st.dataframe(
                        concern[display_cols],
                        width="stretch",
                        hide_index=True
                    )
                else:
                    st.success(
                        "No campaigns meet both high-spend "
                        "and below-median ROAS conditions."
                    )

            # Metric-level observations
            st.divider()
            st.subheader("Metric health observations")

            observations = []

            if "CTR" in diagnosis.columns:
                median_ctr = diagnosis["CTR"].median()
                if pd.notna(median_ctr):
                    observations.append(
                        f"Median CTR in the selected data: "
                        f"{median_ctr:.2f}%. Compare individual campaigns "
                        "against this value to investigate engagement."
                    )

            if "CPA" in diagnosis.columns:
                median_cpa = diagnosis["CPA"].median()
                if pd.notna(median_cpa):
                    observations.append(
                        f"Median CPA in the selected data: "
                        f"₹{median_cpa:,.2f}. Higher CPA may indicate "
                        "less efficient acquisition."
                    )

            if "ROAS" in diagnosis.columns:
                observations.append(
                    f"Median ROAS in the selected data: "
                    f"{median_roas:.2f}x. This is a relative benchmark, "
                    "not a business profitability target."
                )

            if observations:
                for observation in observations:
                    st.write("• " + observation)
            else:
                st.info(
                    "No additional metrics are available "
                    "for metric health observations."
                )

            st.info(
                "Important: These are relative comparisons within the "
                "currently filtered data. They do not establish the "
                "business target or prove the cause of poor performance. "
                "Review attribution, margins, campaign objectives and "
                "conversion volume before making decisions."
            )

        else:
            st.warning(
                "ROAS is unavailable. Campaign ranking by return "
                "cannot be performed for this dataset."
            )

    else:
        st.info("Campaign-level analysis is not available.")

# =========================================================
# 10. FUNNEL AND METRIC ANALYSIS
# =========================================================

elif page == "Funnel & Metric Analysis":

    st.header("Funnel & Metric Analysis")

    st.caption(
        "Understand where users are dropping off "
        "between campaign engagement and purchase."
    )

    funnel_stages = [
        "Clicks",
        "Landing Page Views",
        "Add To Cart",
        "Checkout Initiated",
        "Purchases"
    ]

    available_stages = [
        stage for stage in funnel_stages
        if stage in filtered_df.columns
        and filtered_df[stage].notna().any()
    ]

    if len(available_stages) < 2:

        st.warning(
            "At least two funnel stages with available data "
            "are needed for funnel analysis."
        )

    else:

        funnel_values = [
            filtered_df[stage].sum(min_count=1)
            for stage in available_stages
        ]

        funnel_chart = go.Figure(
            go.Funnel(
                y=available_stages,
                x=funnel_values,
                textinfo="value+percent initial"
            )
        )

        funnel_chart.update_layout(
            title="Overall Conversion Funnel"
        )

        st.plotly_chart(
            funnel_chart,
            use_container_width=True
        )

        st.subheader("Stage-to-Stage Conversion")

        conversion_rows = []

        for i in range(len(available_stages) - 1):

            first_stage = available_stages[i]
            next_stage = available_stages[i + 1]

            first_value = filtered_df[first_stage].sum(
                min_count=1
            )

            next_value = filtered_df[next_stage].sum(
                min_count=1
            )

            if (
                pd.isna(first_value)
                or pd.isna(next_value)
                or first_value == 0
            ):
                rate = None
                drop = None
            else:
                rate = next_value / first_value * 100
                drop = first_value - next_value

            conversion_rows.append({
                "From": first_stage,
                "To": next_stage,
                "Starting volume": first_value,
                "Next stage volume": next_value,
                "Conversion rate (%)": rate,
                "Volume difference": drop
            })

        conversion_df = pd.DataFrame(conversion_rows)

        st.dataframe(
            conversion_df,
            use_container_width=True,
            hide_index=True
        )

        st.subheader("Potential Funnel Issues")

        if not conversion_df.empty:

            valid_rates = conversion_df.dropna(
                subset=["Conversion rate (%)"]
            )

            if not valid_rates.empty:

                weakest = valid_rates.loc[
                    valid_rates["Conversion rate (%)"].idxmin()
                ]

                st.warning(
                    f"The lowest observed stage-to-stage conversion "
                    f"is from {weakest['From']} to {weakest['To']} "
                    f"({weakest['Conversion rate (%)']:.2f}%). "
                    "Investigate this stage. The data alone does not "
                    "establish the cause."
                )

        # Campaign-level diagnostic table
        st.divider()
        st.subheader("Campaign Metric Breakdown")

        diagnostic_columns = [
            "Campaign",
            "Spend",
            "CTR (%)",
            "CPC",
            "CPM",
            "LPV Rate (%)",
            "ATC Rate (%)",
            "Checkout Rate (%)",
            "Purchase Rate (%)",
            "CPA",
            "ROAS"
        ]

        available_diagnostic_columns = [
            col for col in diagnostic_columns
            if col in analysis_report.columns
        ]

        st.dataframe(
            analysis_report[available_diagnostic_columns],
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "Diagnostic guidance: a low CTR may warrant a review "
            "of creative, audience or message relevance; a low "
            "landing-page-view rate may warrant checking page load "
            "and tracking; a low add-to-cart rate may warrant "
            "reviewing the offer or product page; a low purchase "
            "rate may warrant checking checkout friction, price, "
            "payment options and tracking. These are hypotheses "
            "to investigate, not proven causes."
        )


# =========================================================
# 11. MASTER DATA
# =========================================================

elif page == "Master Data":

    st.header("Consolidated Master Data")

    st.caption(
        "The merged records used by the dashboard and analysis."
    )

    st.write(
        f"Records: {len(filtered_df):,} | "
        f"Campaigns: {filtered_df['Campaign'].nunique():,}"
    )

    search = st.text_input(
        "Search campaign, ad set or ad"
    )

    display_df = filtered_df.copy()

    if search.strip():

        searchable_columns = [
            col for col in ["Campaign", "Ad Set", "Ad"]
            if col in display_df.columns
        ]

        mask = pd.Series(
            False, index=display_df.index
        )

        for col in searchable_columns:
            mask = mask | display_df[col].astype(str).str.contains(
                search, case=False, na=False
            )

        display_df = display_df[mask]

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "Download Filtered Master Data",
        data=display_df.to_csv(index=False).encode("utf-8"),
        file_name="Meta_Ads_Master_Data.csv",
        mime="text/csv"
    )

    excel_file = make_excel(
        master_df,
        campaign_report,
        sheet_audit
    )

    st.download_button(
        "Download Complete Analysis Excel",
        data=excel_file,
        file_name="Marketing_Campaign_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# 12. DATA AUDIT
# =========================================================

elif page == "Data Audit":

    st.header("Data Processing & Audit")

    st.subheader("Detected and processed worksheets")

    st.dataframe(
        sheet_audit,
        use_container_width=True,
        hide_index=True
    )

    st.subheader("Merge audit")

    if merge_audit.empty:
        st.info(
            "Only one recognised data worksheet was available. "
            "No merge was required."
        )
    else:
        st.dataframe(
            merge_audit,
            use_container_width=True,
            hide_index=True
        )

    st.subheader("Ignored worksheets")

    if ignored_sheets:
        st.write(ignored_sheets)
    else:
        st.write("No unrecognised worksheets.")

    st.subheader("Missing data overview")

    missing_df = pd.DataFrame({
        "Column": master_df.columns,
        "Missing records": master_df.isna().sum().values,
        "Missing percentage": (
            master_df.isna().mean().values * 100
        ).round(2)
    })

    st.dataframe(
        missing_df,
        use_container_width=True,
        hide_index=True
    )

    st.subheader("Data validation summary")

    audit_cols = st.columns(3)

    audit_cols[0].metric(
        "Master records",
        f"{len(master_df):,}"
    )

    audit_cols[1].metric(
        "Campaigns",
        f"{master_df['Campaign'].nunique():,}"
    )

    audit_cols[2].metric(
        "Date range",
        f"{master_df['Date'].min().date()} to "
        f"{master_df['Date'].max().date()}"
    )

    st.caption(
        "Missing values are intentionally retained as unavailable. "
        "They are not automatically converted to zero."
    )


# =========================================================
# 13. FOOTER
# =========================================================

st.divider()

st.caption(
    "Marketing Campaign Intelligence | "
    "Meta Ads | Analysis is based on the uploaded reports "
    "and selected filters."
)
