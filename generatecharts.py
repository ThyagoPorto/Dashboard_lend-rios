import pandas as pd
import os

# Caminho do Excel dentro da pasta upload
excel_path = os.path.join(os.path.dirname(__file__), "upload", "dados.xlsx")

def get_metrics():
    df = pd.read_excel(excel_path, sheet_name="Métricas")

    metrics_data = []

    # CSAT
    csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
    csat_data = df.iloc[csat_row_index + 1, :].dropna().tolist()
    csat_semanas = df.columns[1:len(csat_data)+1].tolist()
    csat_meta = df.iloc[csat_row_index, 1]
    csat_status = df.iloc[csat_row_index, 2]
    metrics_data.append({
        "metrica": "CSAT",
        "meta_objetivo": f"{csat_meta}%",
        "meta_percentual": csat_meta / 100,
        "cor": "#00D9FF",
        "status": csat_status,
        "semanas": [{"periodo": p, "cumprido": v, "total": v, "percentual": v/csat_meta if csat_meta else 0}
                    for p, v in zip(csat_semanas, csat_data)],
        "total_mes": {
            "cumprido": sum(csat_data),
            "total": sum(csat_data),
            "percentual": (sum(csat_data)/csat_meta) if csat_meta else 0
        }
    })

    # SLA dos DS
    sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
    sla_data = df.iloc[sla_row_index + 1, :].dropna().tolist()
    sla_semanas = df.columns[1:len(sla_data)+1].tolist()
    sla_meta = df.iloc[sla_row_index, 1]
    sla_status = df.iloc[sla_row_index, 2]
    metrics_data.append({
        "metrica": "SLA dos DS",
        "meta_objetivo": f"{sla_meta}%",
        "meta_percentual": sla_meta / 100,
        "cor": "#57A9FB",
        "status": sla_status,
        "semanas": [{"periodo": p, "cumprido": v, "total": v, "percentual": v/sla_meta if sla_meta else 0}
                    for p, v in zip(sla_semanas, sla_data)],
        "total_mes": {
            "cumprido": sum(sla_data),
            "total": sum(sla_data),
            "percentual": (sum(sla_data)/sla_meta) if sla_meta else 0
        }
    })

    # Cobertura de Carteira
    cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
    cobertura_data = df.iloc[cobertura_row_index + 1, :].dropna().tolist()
    cobertura_semanas = df.columns[1:len(cobertura_data)+1].tolist()
    cobertura_meta = df.iloc[cobertura_row_index, 1]
    cobertura_status = df.iloc[cobertura_row_index, 2]
    metrics_data.append({
        "metrica": "Cobertura de Carteira",
        "meta_objetivo": f"{cobertura_meta}%",
        "meta_percentual": cobertura_meta / 100,
        "cor": "#8B5CF6",
        "status": cobertura_status,
        "semanas": [{"periodo": p, "cumprido": v, "total": v, "percentual": v/cobertura_meta if cobertura_meta else 0}
                    for p, v in zip(cobertura_semanas, cobertura_data)],
        "total_mes": {
            "cumprido": sum(cobertura_data),
            "total": sum(cobertura_data),
            "percentual": (sum(cobertura_data)/cobertura_meta) if cobertura_meta else 0
        }
    })

    # Churn
    churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]
    churn_data = df.iloc[churn_row_index + 1, :].dropna().tolist()
    churn_semanas = df.columns[1:len(churn_data)+1].tolist()
    churn_meta = df.iloc[churn_row_index, 1]
    churn_status = df.iloc[churn_row_index, 2]
    metrics_data.append({
        "metrica": "Cancelamento - Churn",
        "meta_objetivo": f"{churn_meta}%",
        "meta_percentual": churn_meta / 100,
        "cor": "#FF6B35",
        "status": churn_status,
        "semanas": [{"periodo": p, "cumprido": v, "total": v, "percentual": v/churn_meta if churn_meta else 0}
                    for p, v in zip(churn_semanas, churn_data)],
        "total_mes": {
            "cumprido": sum(churn_data),
            "total": sum(churn_data),
            "percentual": (sum(churn_data)/churn_meta) if churn_meta else 0
        }
    })

    return metrics_data


def get_summary():
    df = pd.read_excel(excel_path, sheet_name="Métricas")

    total_metrics = 4  # Agora são 4 métricas
    atingidas = 0

    # Para cada métrica, verificamos se cumpriu a meta
    for metrica in ["CSAT", "SLA dos DS", "Cobertura de Carteira", "Cancelamento - Churn"]:
        row_index = df[df.iloc[:, 0] == metrica].index[0]
        meta_percent = df.iloc[row_index, 1] / 100
        data = df.iloc[row_index + 1, :].dropna().tolist()
        if (sum(data)/sum(data)) >= meta_percent:  # condição simplificada
            atingidas += 1

    nao_atingidas = total_metrics - atingidas
    performance = (atingidas / total_metrics) * 100

    return {
        "total_metrics": total_metrics,
        "atingidas": atingidas,
        "nao_atingidas": nao_atingidas,
        "performance": performance
    }


def get_weekly_data():
    df = pd.read_excel(excel_path, sheet_name="Métricas")

    semanas = df.columns[1:].tolist()

    csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
    sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
    cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
    churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]

    weekly_data = {
        "semanas": semanas[1:],  # remove coluna de título
        "csat": df.iloc[csat_row_index + 1, 1:].dropna().tolist(),
        "sla": df.iloc[sla_row_index + 1, 1:].dropna().tolist(),
        "cobertura": df.iloc[cobertura_row_index + 1, 1:].dropna().tolist(),
        "churn": df.iloc[churn_row_index + 1, 1:].dropna().tolist()
    }

    return weekly_data
