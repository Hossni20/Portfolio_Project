import pandas as pd
import numpy as np
import yfinance as yf
import matplotlib.pyplot as plt
import seaborn as sns 
tickers_dict = {
    "Technology": ["MSFT", "NVDA","GOOGL", "META"],
    "E-commerce/Cloud": ["AMZN"],
    "Semiconductors": ["AVGO"],
    "Health": ["LLY", "JNJ"],
    "Energie": ["XOM"],
    "Conso": ["KO"],
    "Crypto": ["BTC-USD", "ETH-USD"]
}
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
my_tickers = [t for sector in tickers_dict.values() for t in sector]
benchmark_ticker = '^GSPC'
start_date = '2022-10-25'
end_date = '2025-10-25' 

try:
    data_all = yf.download(my_tickers + [benchmark_ticker], start=start_date, end=end_date, auto_adjust=True, progress=False)

    if isinstance(data_all.columns, pd.MultiIndex):
        try:
            prices = data_all['Close']
        except KeyError:
            prices = data_all.xs(data_all.columns.get_level_values(0)[0], axis=1, level=0)
    else:
        prices = data_all

    bench_prices = prices[benchmark_ticker]
    stock_prices = prices[my_tickers]
    aligned_data = stock_prices.join(bench_prices.to_frame(name='Benchmark'), how='inner').dropna()

except Exception as e:
    print(f"Erreur: {e}")
    exit()

daily_returns = aligned_data.pct_change().dropna()
r_bench = daily_returns['Benchmark']
r_stocks = daily_returns.drop(columns=['Benchmark'])
weights_list = [custom_weights[ticker] for ticker in r_stocks.columns]
r_portfolio = r_stocks.dot(weights_list)
analysis_df = pd.DataFrame({
    'Mon Portefeuille': r_portfolio,
    'S&P 500': r_bench
})
yearly_periods = [
    ('2022-10-25', '2023-10-25'),
    ('2023-10-25', '2024-10-25'),
    ('2024-10-25', '2025-10-25')
]
results_list = []

print("\nRÉSULTATS ANNUELS :")
print("-" * 80)
print(f"{'Période':<25} | {'Mon Portfolio':<15} | {'S&P 500':<10} | {'Diff (Alpha)'}")
print("-" * 80)

for p_start, p_end in yearly_periods:
    mask = (analysis_df.index >= p_start) & (analysis_df.index <= p_end)
    period_data = analysis_df.loc[mask]
    
    if period_data.empty:
        print(f"{p_start} à {p_end} | Données futures ou indisponibles.")
        continue
        
    total_ret_period = (1 + period_data).prod() - 1
    perf_port = total_ret_period['Mon Portefeuille']
    perf_bench = total_ret_period['S&P 500']
    diff = perf_port - perf_bench
    results_list.append({
        'Année': f"{p_start[:4]}-{p_end[:4]}",
        'Mon Portefeuille': perf_port,
        'S&P 500': perf_bench,
        'Surperformance (Alpha)': diff
    })
    
    print(f"{p_start} -> {p_end} | {perf_port:>14.2%} | {perf_bench:>9.2%} | {diff:>+10.2%}")

print("-" * 80)

if results_list:
    excel_filename = "Resultats_Annuels.xlsx"
    res_df = pd.DataFrame(results_list)
    
    try:
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            res_df.to_excel(writer, sheet_name='Comparaison_Annuelle', index=False)
        print(f"\nFichier Excel généré : {excel_filename}")
    except Exception as e:
        print(f"\nErreur: {e}")

sns.set_theme(style="whitegrid")
fig, axes = plt.subplots(2, 1, figsize=(12, 12))
cumulative_growth = (1 + analysis_df).cumprod() * 100
ax1 = axes[0]
ax1.plot(cumulative_growth.index, cumulative_growth['Mon Portefeuille'], label='Mon Portefeuille', color='#2ecc71', linewidth=2)
ax1.plot(cumulative_growth.index, cumulative_growth['S&P 500'], label='S&P 500 Benchmark', color='#34495e', linewidth=2, linestyle='--')
ax1.set_title('Évolution de la Valeur (Base 100 au 22/10/2022)', fontsize=14, fontweight='bold')
ax1.set_ylabel('Valeur du Portefeuille')
ax1.legend()
ax1.grid(True, which='major', linestyle='--', alpha=0.6)

if results_list:
    plot_df = res_df[['Année', 'Mon Portefeuille', 'S&P 500']]
    res_melted = plot_df.melt(id_vars='Année', var_name='Actif', value_name='Rendement')
    ax2 = axes[1]
    sns.barplot(data=res_melted, x='Année', y='Rendement', hue='Actif', ax=ax2, palette=['#2ecc71', '#34495e'])
    ax2.set_title('Performance Annuelle Comparée (Year-over-Year)', fontsize=14, fontweight='bold')
    ax2.set_ylabel('Rendement Total sur la période')

    for container in ax2.containers:
        ax2.bar_label(container, fmt='%.1f%%', padding=3)

plt.tight_layout()
image_filename = "Comparaison_Benchmark.png"
plt.savefig(image_filename, dpi=300, bbox_inches='tight')
print(f"Image sauvegardée : {image_filename}")
plt.show()
