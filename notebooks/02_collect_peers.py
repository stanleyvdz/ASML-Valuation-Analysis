import yfinance as yf
import pandas as pd

tickers = {
    'ASML': 'ASML',
    'Applied Materials': 'AMAT',
    'Lam Research': 'LRCX',
    'KLA Corp': 'KLAC'
}

data = []
for name, ticker in tickers.items():
    stock = yf.Ticker(ticker)
    info = stock.info
    data.append({
        'Company': name,
        'Market Cap ($B)': round(info.get('marketCap', 0) / 1e9, 2),
        'P/E': round(info.get('trailingPE', 0), 2),
        'EV/EBITDA': round(info.get('enterpriseToEbitda', 0), 2),
        'Beta': round(info.get('beta', 0), 2)
    })

df = pd.DataFrame(data)
df.to_csv('../data/comparables.csv', index=False)
print("Peer data saved!")
print(df)