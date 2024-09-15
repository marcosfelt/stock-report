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
def get_financial_reports_polygon( ticker: str, limit: int = 50):
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
    ax.axhline(target, color='red', linewidth=0.5, linestyle='--')
    ax.tick_params(length=0)
    ax.set_ylim()
    ax.set_yticklabels([f"{int(x)}%" for x in ax.get_yticks()])
    for container in ax.containers:
        ax.bar_label(container, fmt='%.1f%%', label_type='edge')
    return ax

### App ###
ticker = st.text_input("Enter a ticker", placeholder="AAPL")

if not ticker:
    st.stop()
df = get_financials_df(ticker)

# Revenue
if df is not None:
    cols = st.columns(2)
    with cols[0]:
        ax = make_bar_plot(df, "revenue_yoy_change", "Revenue YoY Change (%)", color="green", target=7.5)
        st.pyplot(ax.figure)

        # EPS
        ax = make_bar_plot(df, "eps_yoy_change", "EPS YoY Change (%)", color="blue", target=7.5)
        st.pyplot(ax.figure)

    # Pre-tax profit margin
    with cols[1]:
        ax = make_bar_plot(df, "pre_tax_profit_margin", "Pre-tax Profit Margin (%)", color="olive", target=34)
        st.pyplot(ax.figure)


else:
    st.write(f"No data found for {ticker}. ")