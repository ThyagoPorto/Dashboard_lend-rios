import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

# Configuração de fontes para suportar CJK
plt.rcParams["font.family"] = ["sans-serif"]
plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial Unicode MS", "Noto Sans CJK SC"] # Adicionado Noto Sans CJK SC como fallback

# Cores da paleta
palette = ["#4080FF", "#57A9FB", "#37D4CF", "#23C343", "#FBE842", "#FF9A2E", "#A9AEB8"]

df = pd.read_csv("processed_metrics.csv")

for index, row in df.iterrows():
    metric = row["Metrica"]
    total = row["Total"]
    cumprido = row["Cumprido"]
    percentual_meta = row["Percentual_Meta"]

    fig, ax = plt.subplots(figsize=(8, 6))

    # Configurações de estilo
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")
    ax.grid(True, linestyle="--", alpha=0.7, color="lightgray")

    # Dados para o gráfico de barras
    labels = ["Total", "Cumprido"]
    values = [total, cumprido]

    bars = ax.bar(labels, values, color=[palette[0], palette[1]])

    # Adicionar rótulos de valor nas barras
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval + 0.5, round(yval, 2), ha="center", va="bottom")

    # Adicionar o percentual da meta como texto no gráfico
    ax.text(0.5, 0.95, f"Percentual da Meta: {percentual_meta:.2%}" if isinstance(percentual_meta, (int, float)) else f"Percentual da Meta: {percentual_meta}" , transform=ax.transAxes, ha="center", va="top", fontsize=12, bbox=dict(facecolor="white", alpha=0.8, edgecolor="none"))

    ax.set_ylabel("Quantidade")
    ax.set_title(f"Métrica: {metric} - Meta: {percentual_meta:.2%}")

    # Substituir espaços por underscores no nome do arquivo
    file_name = f"{metric.replace(' ', '_')}_chart.png"
    plt.tight_layout()
    plt.savefig(file_name)
    plt.close()

print("Gráficos gerados com sucesso.")
