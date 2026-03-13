import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = f"Analyse_Sectorielle_FINAL_{timestamp}.xlsx"
tickers_dict = {
    "Technology": ["MSFT", "NVDA","GOOGL", "META"],
    "E-commerce/Cloud": ["AMZN"],
    "Semiconductors": ["AVGO"],
    "Health": ["LLY", "JNJ"],
    "Energie": ["XOM"],
    "Conso": ["KO"],
    "Crypto": ["BTC-USD", "ETH-USD"]
}
all_tickers = [ticker for sector_list in tickers_dict.values() for ticker in sector_list]
BENCHMARK_TICKER = '^GSPC'
start_date = '2022-10-25'
end_date = '2025-10-25'

try:
    raw_stocks = yf.download(all_tickers, start=start_date, end=end_date, auto_adjust=True, progress=False)
    
    if isinstance(raw_stocks.columns, pd.MultiIndex):
        try:
            df_stocks = raw_stocks['Close']
        except KeyError:
            df_stocks = raw_stocks.xs(raw_stocks.columns.get_level_values(0)[0], axis=1, level=0)
    else:
        df_stocks = raw_stocks

    raw_bench = yf.download(BENCHMARK_TICKER, start=start_date, end=end_date, auto_adjust=True, progress=False)
    
    if isinstance(raw_bench.columns, pd.MultiIndex):
        bench_series = raw_bench['Close'] if 'Close' in raw_bench.columns else raw_bench.iloc[:, 0]
    else:
        bench_series = raw_bench['Close'] if 'Close' in raw_bench.columns else raw_bench.iloc[:, 0]
        
    if isinstance(bench_series, pd.DataFrame):
        df_bench = bench_series
        df_bench.columns = ['Benchmark']
    else:
        df_bench = bench_series.to_frame(name='Benchmark')

except Exception as e:
    print(f"ERREUR {e}")
    exit()

data = df_stocks.join(df_bench, how='inner').dropna()
daily_returns = data.pct_change().dropna()
bench_ret = daily_returns['Benchmark']
stock_ret = daily_returns.drop(columns=['Benchmark'])
rf_annual = 0.04
ann_factor = 252
metrics = pd.DataFrame(index=stock_ret.columns)
betas, alphas, correlations, info_ratios, tracking_errors = [], [], [], [], []
var_benchmark = bench_ret.var()
mean_benchmark = bench_ret.mean() * ann_factor

for ticker in stock_ret.columns:
    r_stock = stock_ret[ticker]
    corr = r_stock.corr(bench_ret)
    correlations.append(corr)
    cov = r_stock.cov(bench_ret)
    beta = cov / var_benchmark
    betas.append(beta)
    mean_stock = r_stock.mean() * ann_factor
    alpha = mean_stock - (rf_annual + beta * (mean_benchmark - rf_annual))
    alphas.append(alpha)
    active_ret = r_stock - bench_ret
    te = active_ret.std() * np.sqrt(ann_factor)
    tracking_errors.append(te)
    ir = (active_ret.mean() * ann_factor) / te if te != 0 else 0
    info_ratios.append(ir)

metrics['Secteur'] = metrics.index.map({t: s for s, t_list in tickers_dict.items() for t in t_list})
metrics['Rendement Annuel'] = stock_ret.mean() * ann_factor
metrics['Volatilité'] = stock_ret.std() * np.sqrt(ann_factor)
metrics['Sharpe Ratio'] = (metrics['Rendement Annuel'] - rf_annual) / metrics['Volatilité']
metrics['Bêta'] = betas
metrics['Alpha'] = alphas
metrics['Corrélation'] = correlations
metrics['Info Ratio'] = info_ratios
cumulative_returns = (1 + stock_ret).cumprod() - 1
cumulative_bench = (1 + bench_ret).cumprod() - 1
cumulative_returns['BENCHMARK_SP500'] = cumulative_bench

try:
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        metrics.round(4).to_excel(writer, sheet_name='KPI_Financiers')
        cumulative_returns.round(4).to_excel(writer, sheet_name='Donnees_Cumulees')
        daily_returns.round(4).to_excel(writer, sheet_name='Rendements_Bruts')
        
        # Petit bonus : Une feuille avec la compo par secteur
        pd.DataFrame.from_dict(tickers_dict, orient='index').T.to_excel(writer, sheet_name='Composition_Secteurs')
        
    print(f"Fichier généré : {OUTPUT_FILE}")

except Exception as e:
    print(f"Erreur: {e}")