import yfinance as yf
import pandas as pd
import numpy as np
BENCHMARK_TICKER = '^GSPC'
BENCHMARK_NAME = 'SP500'
start = "2022-10-25"
end = "2025-10-25"
data = yf.download(BENCHMARK_TICKER, start=start, end=end, progress=False)

if data.empty:
    raise RuntimeError(f"Aucune donnée téléchargée pour {BENCHMARK_TICKER} pour la période demandée.")

METRIC_COLUMN = 'Close' 

try:
    data_series = data[(METRIC_COLUMN, BENCHMARK_TICKER)]
except KeyError:
    try:
        data_series = data[METRIC_COLUMN]
    except KeyError as e:
        raise KeyError(f"La colonne '{METRIC_COLUMN}' n'a pas été trouvée, colonnes disponibles: " + str(data.columns.tolist())) from e
    
data_prices = data_series.to_frame(name=BENCHMARK_TICKER)
data_prices = data_prices.sort_index()
data_prices = data_prices.dropna() 
weekly = data_prices.resample("W").mean()
monthly = data_prices.resample("ME").mean()
summary = data_prices.describe()
daily_returns = data_prices.pct_change().dropna()
returns_summary = daily_returns.describe()
returns_summary.columns = ['Stats Rendements Quotidiens']

output_filename = f"{BENCHMARK_NAME}_data.xlsx"
try:
    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        data_prices.to_excel(writer, sheet_name="Prix_Quotidien")
        weekly.to_excel(writer, sheet_name="Moyenne_Hebdomadaire")
        monthly.to_excel(writer, sheet_name="Moyenne_Mensuelle")
        summary.to_excel(writer, sheet_name="Stats_Prix")
        returns_summary.to_excel(writer, sheet_name="Stats_Rendements")

    print(f"Les données du S&P 500 ont été sauvegardées dans: {output_filename}")
    
except Exception as e:
    print(f"\nErreur lors de l'exportation Excel. Erreur: {e}")