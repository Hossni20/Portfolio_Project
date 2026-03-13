import yfinance as yf
import pandas as pd
tickers = {
    "Technology": ["MSFT", "NVDA","GOOGL", "META"],
    "E-commerce/Cloud": ["AMZN"],
    "Semiconductors": ["AVGO"],
    "Health": ["LLY", "JNJ"],
    "Energie": ["XOM"],
    "Conso": ["KO"],
    "Crypto": ["BTC-USD", "ETH-USD"]
}
all_tickers = [ticker for sector_tickers in tickers.values() for ticker in sector_tickers]
start = "2022-10-25"
end = "2025-10-25"
data = yf.download(all_tickers, start=start, end=end, group_by='ticker', threads=True)

if data.empty:
    raise RuntimeError("No data downloaded for the requested tickers/date range.")

try:
    data = data["Adj Close"]
except Exception:
    try:
        data = data.xs("Adj Close", axis=1, level=1)
    except Exception:
        pass

data = data.sort_index()
data = data.dropna(how="all", axis=1)
weekly = data.resample("W").mean()
monthly = data.resample("ME").mean()
summary = data.describe().T
corr = data.corr()
sector_map = {ticker: sector for sector, tlist in tickers.items() for ticker in tlist}
metadata = pd.DataFrame.from_dict(sector_map, orient='index', columns=['Sector'])
with pd.ExcelWriter("nasdaq_stock_data.xlsx", engine="openpyxl") as writer:
    data.to_excel(writer, sheet_name="Cleaned")
    weekly.to_excel(writer, sheet_name="Weekly")
    monthly.to_excel(writer, sheet_name="Monthly")
    summary.to_excel(writer, sheet_name="Summary")
    corr.to_excel(writer, sheet_name="Correlation")
    metadata.to_excel(writer, sheet_name="Metadata")
