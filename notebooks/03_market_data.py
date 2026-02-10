import yfinance as yf
import pandas as pd
import numpy as np

# Treasury rate
treasury = yf.Ticker("^TNX")
rf = treasury.history(period="5d")['Close'].iloc[-1] / 100

# Market risk premium
mrp = 0.065

# Calculate beta
asml = yf.Ticker("ASML")
spy = yf.Ticker("SPY")

asml_ret = asml.history(period="2y", interval="1mo")['Close'].pct_change().dropna()
spy_ret = spy.history(period="2y", interval="1mo")['Close'].pct_change().dropna()

common = asml_ret.index.intersection(spy_ret.index)
asml_ret = asml_ret.loc[common]
spy_ret = spy_ret.loc[common]

beta = np.cov(asml_ret, spy_ret)[0][1] / np.var(spy_ret)
re = rf + beta * mrp

# Save
market_data = {
    'risk_free_rate': rf,
    'market_risk_premium': mrp,
    'beta': beta,
    'cost_of_equity': re
}

pd.DataFrame([market_data]).to_csv('../data/market_data.csv', index=False)

print(f"""
Market Data:
Risk-Free: {rf*100:.2f}%
Beta: {beta:.3f}
Cost of Equity: {re*100:.2f}%
""")