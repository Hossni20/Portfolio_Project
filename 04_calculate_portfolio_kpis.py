import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime
tickers_dict = {
    "Technology": ["MSFT", "NVDA", "GOOGL", "META"],
    "E-commerce/Cloud": ["AMZN"],
    "Semiconductors": ["AVGO"],
    "Health": ["LLY", "JNJ"],
    "Energie": ["XOM"],
    "Conso": ["KO"],
    "Crypto": ["BTC-USD", "ETH-USD"]
}
my_tickers = [t for sector in tickers_dict.values() for t in sector]
benchmark_ticker = '^GSPC'
custom_weights = {
    'MSFT': 0.100,    
    'NVDA': 0.117,    
    'GOOGL': 0.067,
    'META': 0.083,   
    'AMZN': 0.083,   
    'AVGO': 0.083,  
    'LLY': 0.083,     
    'JNJ': 0.067,     
    'XOM': 0.075,     
    'KO': 0.075,      
    'BTC-USD': 0.100, 
    'ETH-USD': 0.067 
}
start_date = '2022-10-25'
end_date = '2025-10-25'
risk_free_rate = 0.04
ann_factor = 252
output_file = "Portfolio_Performance_Report.xlsx"
data_all = yf.download(my_tickers + [benchmark_ticker], start=start_date, end=end_date, auto_adjust=True, progress=False)

if isinstance(data_all.columns, pd.MultiIndex):
    try:
        prices = data_all['Close']
    except KeyError:
        prices = data_all.xs(data_all.columns.get_level_values(0)[0], axis=1, level=0)
else:
    prices = data_all

prices_stocks = prices[my_tickers]
prices_bench = prices[benchmark_ticker]
aligned_data = prices_stocks.join(prices_bench.to_frame('Benchmark'), how='inner').dropna()
daily_returns = aligned_data.pct_change().dropna()
r_stocks = daily_returns[my_tickers]
r_bench = daily_returns['Benchmark']
weights_list = [custom_weights[ticker] for ticker in r_stocks.columns]
r_portfolio = r_stocks.dot(weights_list)
ann_ret_port = r_portfolio.mean() * ann_factor
ann_ret_bench = r_bench.mean() * ann_factor
ann_vol_port = r_portfolio.std() * np.sqrt(ann_factor)
ann_vol_bench = r_bench.std() * np.sqrt(ann_factor)
sharpe_ratio = (ann_ret_port - risk_free_rate) / ann_vol_port
correlation = r_portfolio.corr(r_bench)
covariance = r_portfolio.cov(r_bench)
variance_bench = r_bench.var()
beta = covariance / variance_bench
alpha = ann_ret_port - (risk_free_rate + beta * (ann_ret_bench - risk_free_rate))
active_return = r_portfolio - r_bench
tracking_error = active_return.std() * np.sqrt(ann_factor)
info_ratio = (active_return.mean() * ann_factor) / tracking_error if tracking_error != 0 else 0
kpi_data = {
    "Metric": [
        "Annual Return (Moyenne Annuelle)", 
        "Volatility (Volatilité)", 
        "Sharpe Ratio", 
        "Beta", 
        "Alpha (Jensen)", 
        "Correlation (vs S&P 500)", 
        "Information Ratio",
        "Tracking Error"
    ],
    "My Portfolio": [
        ann_ret_port, 
        ann_vol_port, 
        sharpe_ratio, 
        beta, 
        alpha, 
        correlation, 
        info_ratio,
        tracking_error
    ],
    "Benchmark (S&P 500)": [
        ann_ret_bench, 
        ann_vol_bench, 
        (ann_ret_bench - risk_free_rate)/ann_vol_bench, 
        1.0, 
        0.0, 
        1.0, 
        0.0, 
        0.0
    ]
}
df_kpi = pd.DataFrame(kpi_data)
df_weights = pd.DataFrame(list(custom_weights.items()), columns=['Ticker', 'Weight'])
df_weights['Sector'] = df_weights['Ticker'].apply(lambda x: next((k for k,v in tickers_dict.items() if x in v), "Unknown"))
df_weights = df_weights.sort_values(by='Weight', ascending=False)
corr_matrix = pd.DataFrame({
    'Portfolio': r_portfolio,
    'Benchmark': r_bench
}).corr()

try:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_kpi.to_excel(writer, sheet_name='Performance_Summary', index=False)
        df_weights.to_excel(writer, sheet_name='Composition', index=False)
        corr_matrix.to_excel(writer, sheet_name='Correlation_Matrix')
        daily_returns_export = pd.DataFrame({'Portfolio': r_portfolio, 'Benchmark': r_bench})
        daily_returns_export.to_excel(writer, sheet_name='Daily_Returns')

    print(f"File created: {output_file}")

except Exception as e:
    print(f"Error: {e}")