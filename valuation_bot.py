import yfinance as yf
import pandas as pd
import math
import os

# --- CONFIGURATION ---
STOCKS = [
    "RELIANCE.NS", "TCS.NS", "HDFCBANK.NS", "ICICIBANK.NS", "INFY.NS",
    "HINDUNILVR.NS", "ITC.NS", "SBIN.NS", "BHARTIARTL.NS", "TATASTEEL.NS",
    "COALINDIA.NS", "MARUTI.NS", "SUNPHARMA.NS", "ONGC.NS", "POWERGRID.NS",
    "NTPC.NS", "M&M.NS", "TITAN.NS", "BAJFINANCE.NS", "WIPRO.NS"
]

print(f"🚀 STARTING EXCEL REPORT GENERATION ({len(STOCKS)} Stocks)...")

results = []

for symbol in STOCKS:
    try:
        stock = yf.Ticker(symbol)
        info = stock.info
        
        price = info.get('currentPrice', 0)
        book_value = info.get('bookValue', 0)
        eps = info.get('trailingEps', 0)
        
        if book_value <= 0 or eps <= 0:
            continue

        # Graham Number Formula: Sqrt(22.5 * EPS * Book Value)
        graham_value = math.sqrt(22.5 * eps * book_value)
        margin_of_safety = ((graham_value - price) / graham_value) * 100
        
        verdict = "HOLD"
        if price < graham_value:
            verdict = "UNDERVALUED"
        else:
            verdict = "OVERVALUED"

        print(f"{symbol} -> Safety: {round(margin_of_safety, 1)}%")

        results.append({
            "Stock": symbol,
            "Price": price,
            "Fair_Value": round(graham_value, 2),
            "Margin_Safety_%": round(margin_of_safety, 2),
            "Verdict": verdict
        })
            
    except Exception as e:
        print(f"Error: {e}")

if results:
    df = pd.DataFrame(results)
    df = df.sort_values(by="Margin_Safety_%", ascending=False)
    filename = "Valuation_Report.csv"
    df.to_csv(filename, index=False)
    print(f"✅ SUCCESS! Report saved as {filename}")