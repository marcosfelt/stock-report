import os
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches

load_dotenv()
POLYGON_API_KEY = os.getenv("POLYGON_API_KEY")
STOCKS = [
    # "AAPL",
    # "AL",
    # "AX",
    # "DAR",
    # "INMD",
    # "META",
    "MSFT",
    # "NVDA",
    # "PAYC",
    # "SCHW",
    # "SWKS",
    # "TSCO",
    "V",
    # "VRTX",
]


### Functions ###
@st.cache_data
def get_last_close_price(ticker: str):
    """Get last closing price from Polygon.io"""
    # Get last weekday
    today = datetime.now() - timedelta(days=1)
    while today.weekday() > 4:
        today -= timedelta(days=1)
    date_str = today.strftime("%Y-%m-%d")
    res = requests.get(
        f"https://api.polygon.io/v1/open-close/{ticker}/{date_str}",
        params={
            "apiKey": POLYGON_API_KEY,
        },
    )
    if res.status_code != 200:
        st.warning(f"Failed to fetch data from Polygon ({res.status_code}): {res.text}")
        return
    data = res.json()
    return data["close"]


@st.cache_data
def get_financial_reports_polygon(ticker: str, limit: int = 50):
    """Get financial reports from Polygon.io"""
    res = requests.get(
        "https://api.polygon.io/vX/reference/financials",
        params={"apiKey": POLYGON_API_KEY, "ticker": ticker, "limit": limit},
    )
    if res.status_code != 200:
        st.warning(f"Failed to fetch data from Polygon ({res.status_code}): {res.text}")
        return
    data = res.json()
    if "results" not in data:
        return
    return data["results"]


def extract_quarter_financials(quarter: dict):
    q = quarter["fiscal_period"]
    yr = quarter["fiscal_year"]
    eps = quarter["financials"]["income_statement"]["basic_earnings_per_share"]["value"]
    revenue = quarter["financials"]["income_statement"]["revenues"]["value"]
    income = quarter["financials"]["income_statement"][
        "income_loss_from_continuing_operations_before_tax"
    ]["value"]
    return {"quarter": q, "year": yr, "eps": eps, "revenue": revenue, "income": income}


def get_financials_df(ticker: str, limit: int = 50) -> pd.DataFrame:
    reports = get_financial_reports_polygon(ticker, limit)
    if reports is None:
        return
    quarterly_financials = [
        extract_quarter_financials(q)
        for q in reports
        if q["fiscal_period"] in ["Q1", "Q2", "Q3", "Q4"]
    ]
    df = pd.DataFrame(quarterly_financials)
    df = df.set_index(["year", "quarter"])
    yoy_change = (df - df.shift(-4)) / df.shift(-4) * 100
    yoy_change = yoy_change.dropna().rename(columns=lambda x: f"{x}_yoy_change")
    df = pd.concat([df, yoy_change], axis=1)
    df["fiscal_period"] = df.index.get_level_values(
        "quarter"
    ) + df.index.get_level_values("year").astype(str)
    df["pre_tax_profit_margin"] = df["income"] / df["revenue"] * 100
    return df


def make_bar_plot(
    df: pd.DataFrame, col: str, title: str, color="green", target=7.5, n_quarters=4
):
    ax = df.iloc[:n_quarters][::-1].plot.bar(
        x="fiscal_period", y=col, color=color, rot=0
    )
    ax.grid(axis="y", color="gray", linestyle="-", linewidth=0.5, alpha=0.2)
    ax.get_legend().remove()
    # Remove lines
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_xlabel("")
    ax.set_title(title)
    ax.axhline(target, color="red", linewidth=2.0, linestyle="--")
    ax.tick_params(length=0)
    ax.set_ylim()
    ax.set_yticklabels([f"{int(x)}%" for x in ax.get_yticks()])
    for container in ax.containers:
        ax.bar_label(container, fmt="%.1f%%", label_type="edge")
    return ax


def make_ranges_plot(
    current_price: float,
    buy_price: float,
    hold_price: float,
    sell_price: float,
    sell_upper_price: float,
):
    ranges = pd.DataFrame(
        [
            [buy_price, hold_price - buy_price],
            [hold_price, sell_price - hold_price],
            [sell_price, sell_upper_price - sell_price],
        ],
        columns=["Low", "High"],
        index=["Buy", "Hold", "Sell"],
    )
    ax = ranges.plot.bar(stacked=True, color=[(0, 0, 0, 0), "grey"], rot=0)
    ax.get_legend().remove()
    ax.grid(axis="y", color="gray", linestyle="-", linewidth=0.5, alpha=0.2)
    for container in ax.containers:
        ax.bar_label(container, fmt="$%.1f", label_type="edge")
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(length=0)
    ax.axhline(current_price, color="black", linewidth=0.5)
    ax.text(
        2,
        current_price,
        f"Last close: (${current_price:.2f})",
        va="center",
        ha="center",
        backgroundcolor=(1, 1, 1, 0.7),
    )
    return ax


def make_ppt_report(
    ticker: str,
    financial_period: str,
    decision: str,
    comments: str,
    ax_revenue,
    ax_eps,
    ax_ptpm,
    ax_ranges,
):
    prs = Presentation()
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = f"{ticker} {financial_period} Financial Report"
    title.text_frame.paragraphs[0].font.size = Inches(0.3)
    # Left justify title
    title.text_frame.paragraphs[0].alignment = 1
    # Move title up
    title.top = Inches(0.15)
    title.width = Inches(6)
    title.left = Inches(0.1)

    left = Inches(1)
    top = Inches(1.5)
    width = Inches(4)
    height = Inches(3)

    # Add recommendation
    text_box = slide.shapes.add_textbox(
        Inches(0.1), Inches(0.6), width * 2, Inches(0.5)
    ).text_frame
    text_box.text = f"Recommendation: {decision}"
    text_box.paragraphs[0].font.size = Inches(0.15)

    # Add comments
    text_box = slide.shapes.add_textbox(
        Inches(0.1), Inches(1.0), width * 2, Inches(0.5)
    ).text_frame
    text_box.text = comments
    text_box.paragraphs[0].font.size = Inches(0.1)
    text_box.word_wrap = True

    img_stream = BytesIO()
    ax_revenue.figure.savefig(img_stream, format="png")
    slide.shapes.add_picture(img_stream, left, top, width, height)

    img_stream = BytesIO()
    ax_eps.figure.savefig(img_stream, format="png")
    slide.shapes.add_picture(img_stream, left + width, top, width, height)

    img_stream = BytesIO()
    ax_ptpm.figure.savefig(img_stream, format="png")
    slide.shapes.add_picture(img_stream, left, top + height, width, height)

    img_stream = BytesIO()
    ax_ranges.figure.savefig(img_stream, format="png")
    slide.shapes.add_picture(img_stream, left + width, top + height, width, height)
    ppt_buffer = BytesIO()
    prs.save(ppt_buffer)
    return ppt_buffer


### App ###
st.set_page_config(
    page_title="Stock Tracking Report",
    page_icon=":chart_with_upwards_trend:",
)
st.title("Stock Tracking Report")

ticker = st.selectbox("Ticker", STOCKS, index=None)


if not ticker:
    st.stop()
df = get_financials_df(ticker)
current_price = get_last_close_price(ticker)

# Revenue, EPS and PTPM targets
st.write("### Targets")
cols = st.columns(3)
with cols[0]:
    revenue_target = st.number_input("Revenue YoY % Target", 0, value=10)
with cols[1]:
    ptpm_target = st.number_input("PTPM (%) Target", 0, value=10)
with cols[2]:
    eps_target = st.number_input("EPS YoY (%) Target", 0, value=10)

# Buy sell, hold range
st.write("")
st.write("### Buy, Sell, Hold Range")
cols = st.columns(4)
with cols[0]:
    buy = st.number_input("Buy lower price", 0, value=100)
with cols[1]:
    hold = st.number_input("Hold lower price", 0, value=200)
with cols[2]:
    sell_lower = st.number_input("Sell lower price", 0, value=300)
with cols[3]:
    sell_upper = st.number_input("Sell upper price", 0, value=350)

# Comment
st.write("### Recommendation")
decision = st.selectbox("My recommendation", ["Buy", "Sell", "Hold"])

comments = st.text_input(
    "Comments", placeholder=f"This is a {decision.lower()} because..."
)

# Download
download_container = st.container()

st.divider()
# Create report
st.write("_Report preview_")
st.write(f"#### {ticker} {df.iloc[0]['fiscal_period']} Report")
st.write(f"Recommendation: **{decision}**")
st.caption(comments)
if df is not None:
    cols = st.columns(2)
    with cols[0]:
        # Revenue YoY
        ax_revenue = make_bar_plot(
            df,
            "revenue_yoy_change",
            "Revenue YoY Growth Rate (%)",
            color="green",
            target=revenue_target,
        )
        st.pyplot(ax_revenue.figure)

        # EPS YoY
        ax_eps = make_bar_plot(
            df,
            "eps_yoy_change",
            "EPS YoY Growth Rate (%)",
            color="blue",
            target=eps_target,
        )
        st.pyplot(ax_eps.figure)

    with cols[1]:
        # Pre-tax profit margin
        ax_ptpm = make_bar_plot(
            df,
            "pre_tax_profit_margin",
            "Pre-tax Profit Margin (%)",
            color="olive",
            target=ptpm_target,
        )
        st.pyplot(ax_ptpm.figure)

        # Buy, hold sell
        ax_ranges = make_ranges_plot(current_price, buy, hold, sell_lower, sell_upper)
        ax_ranges.set_title("Buy, Hold, Sell Ranges")
        st.pyplot(ax_ranges.figure)
else:
    st.write(f"No data found for {ticker}. ")
st.divider()

with download_container:
    report_buffer = make_ppt_report(
        ticker=ticker,
        decision=decision,
        financial_period=df.iloc[0]["fiscal_period"],
        comments=comments,
        ax_revenue=ax_revenue,
        ax_eps=ax_eps,
        ax_ptpm=ax_ptpm,
        ax_ranges=ax_ranges,
    )

    st.download_button(
        "Download report", report_buffer, "StockReport.pptx", "pptx", type="primary"
    )
