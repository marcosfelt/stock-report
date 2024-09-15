import requests
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt
import os
from dotenv import load_dotenv
import streamlit as st

load_dotenv()
POLYGON_API_KEY= os.getenv("POLYGON_API_KEY")

### Functions ###
@st.cache_data
def get_financial_reports_polygon(ticker: str, limit: int = 50):
    """Get financial reports from Polygon.io"""
    res = requests.get(
        f"https://api.polygon.io/vX/reference/financials",
        params={
            "apiKey": POLYGON_API_KEY, 
            "ticker": ticker, 

            "limit": limit
        }
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
    income = quarter["financials"]["income_statement"]["income_loss_from_continuing_operations_before_tax"]["value"]
    return {
        "quarter": q,
        "year": yr,
        "eps": eps,
        "revenue": revenue,
        "income": income
    }

def get_financials_df(ticker: str, limit: int = 50)-> pd.DataFrame:
    reports = get_financial_reports_polygon(ticker, limit)
    if reports is None:
        return
    quarterly_financials = [
        extract_quarter_financials(q) for q in reports 
        if q["fiscal_period"] in ["Q1", "Q2", "Q3", "Q4"]
    ]
    df = pd.DataFrame(quarterly_financials)
    df = df.set_index(["year", "quarter"])
    yoy_change = (df - df.shift(-4)) / df.shift(-4) * 100
    yoy_change = yoy_change.dropna().rename(columns=lambda x: f"{x}_yoy_change")
    df = pd.concat([df, yoy_change], axis=1)
    df["fiscal_period"] = df.index.get_level_values("quarter") + df.index.get_level_values("year").astype(str)
    df["pre_tax_profit_margin"] = df["income"] / df["revenue"] * 100
    return df

def make_bar_plot(
    df: pd.DataFrame,
    col: str, 
    title: str, 
    color="green", 
    target=7.5,
    n_quarters=4
):
    ax = df.iloc[:n_quarters].plot.bar(x="fiscal_period", y=col, color=color, rot=0)
    ax.grid(axis='y', color='gray', linestyle='-', linewidth=0.5, alpha=0.2)
    ax.get_legend().remove()
    # Remove lines
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.set_xlabel("")
    ax.set_title(title)
    ax.axhline(target, color='red', linewidth=2.0, linestyle='--')
    ax.tick_params(length=0)
    ax.set_ylim()
    ax.set_yticklabels([f"{int(x)}%" for x in ax.get_yticks()])
    for container in ax.containers:
        ax.bar_label(container, fmt='%.1f%%', label_type='edge')
    return ax

### App ###
st.title("Stock Tracking Report")

ticker = st.selectbox("Ticker", ["AAPL", "MSFT", "GOOGL", "AMZN", "TSLA"])


if not ticker:
    st.stop()
df = get_financials_df(ticker)

# Revenue, EPS and PTPM targets
st.write("### Targets")
cols = st.columns(3)
with cols[0]:
    revenue_target = st.number_input("Revenue YoY % Target", 10)
with cols[1]:
    ptpm_target = st.number_input("PTPM (%) Target", 10)
with cols[2]:
    eps_target = st.number_input("EPS YoY (%) Target", 10)

# Buy sell, hold range
st.write("### Buy, Sell, Hold Range")
cols = st.columns(3)
with cols[0]:
    st.number_input("Buy lower price", 100)
with cols[1]:
    st.number_input("Hold lower price", 200)
with cols[2]:
    st.number_input("Sell lower price", 300)

# Comment
st.write("### Recommendation")
decision = st.selectbox("My recommendation", ["Buy", "Sell", "Hold"])

st.text_input("Comments", placeholder="This is a hold because..") 


# Download
download_container = st.container()

st.divider()
# Create report
if df is not None:
    cols = st.columns(2)
    with cols[0]:
        # Revenue YoY
        ax = make_bar_plot(df, "revenue_yoy_change", "Revenue YoY Change (%)", color="green", target=revenue_target)
        st.pyplot(ax.figure)

        # EPS YoY
        ax = make_bar_plot(df, "eps_yoy_change", "EPS YoY Change (%)", color="blue", target=eps_target)
        st.pyplot(ax.figure)

    with cols[1]:
        # Pre-tax profit margin
        ax = make_bar_plot(df, "pre_tax_profit_margin", "Pre-tax Profit Margin (%)", color="olive", target=ptpm_target)
        st.pyplot(ax.figure)

else:
    st.write(f"No data found for {ticker}. ")

with download_container:
    st.download_button("Download report", df.to_csv(), "report.csv", "text/csv", type="primary")