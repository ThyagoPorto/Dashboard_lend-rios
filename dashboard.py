import os
import pandas as pd
from flask import Blueprint, jsonify

dashboard_bp = Blueprint("dashboard", __name__, url_prefix="/api/dashboard")

# Caminho do Excel
excel_path = os.path.join(os.path.dirname(__file__), "upload", "dados.xlsx")


def load_excel():
    try:
        return pd.ExcelFile(excel_path)
    except Exception as e:
        return None


@dashboard_bp.route("/metrics", methods=["GET"])
def get_metrics():
    try:
        xls = load_excel()
        if not xls:
            return jsonify({"success": False, "error": "Erro ao carregar Excel"})

        df = pd.read_excel(xls, sheet_name="Métricas")

        metrics_data = []

        # --- CSAT ---
        csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
        csat_data = df.iloc[csat_row_index + 1, :].dropna().tolist()
        metrics_data.append(_build_metric("CSAT", "#00D9FF", csat_data))

        # --- SLA dos DS ---
        sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
        sla_data = df.iloc[sla_row_index + 1, :].dropna().tolist()
        metrics_data.append(_build_metric("SLA dos DS", "#00D9FF", sla_data))

        # --- Cobertura de Carteira ---
        cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
        cobertura_data = df.iloc[cobertura_row_index + 1, :].dropna().tolist()
        metrics_data.append(_build_metric("Cobertura de Carteira", "#8B5CF6", cobertura_data))

        # --- Cancelamento - Churn ---
        churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]
        churn_data = df.iloc[churn_row_index + 1, :].dropna().tolist()
        metrics_data.append(_build_metric("Cancelamento - Churn", "#FF6B35", churn_data))

        return jsonify({"success": True, "data": metrics_data})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


def _build_metric(name, color, data):
    meta_percentual = data[0]
    meta_objetivo = f"{int(meta_percentual*100)}%"
    semanas = []

    for i in range(1, len(data) - 2, 3):  # total, cumprido, percentual
        periodo = data[i]
        total = data[i + 1]
        cumprido = data[i + 2]
        percentual = round(cumprido / total, 4) if total else 0
        semanas.append({
            "periodo": periodo,
            "total": total,
            "cumprido": cumprido,
            "percentual": percentual
        })

    total_mes = semanas[-1]
    status = _get_status(total_mes["percentual"], meta_percentual)

    return {
        "metrica": name,
        "cor": color,
        "meta_objetivo": meta_objetivo,
        "meta_percentual": meta_percentual,
        "semanas": semanas,
        "total_mes": total_mes,
        "status": status
    }


def _get_status(percentual, meta):
    if percentual >= meta:
        return "Excelente"
    elif percentual >= meta * 0.9:
        return "Atenção"
    else:
        return "Crítico"


@dashboard_bp.route("/weekly-data", methods=["GET"])
def get_weekly_data():
    try:
        xls = load_excel()
        if not xls:
            return jsonify({"success": False, "error": "Erro ao carregar Excel"})

        df = pd.read_excel(xls, sheet_name="Métricas")

        # CSAT
        csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
        csat_data = df.iloc[csat_row_index + 1, :].dropna().tolist()
        csat = [round(cumprido / total * 100, 2) if total else 0
                for _, total, cumprido in zip(csat_data[1::3], csat_data[2::3], csat_data[3::3])]

        # SLA
        sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
        sla_data = df.iloc[sla_row_index + 1, :].dropna().tolist()
        sla = [round(cumprido / total * 100, 2) if total else 0
               for _, total, cumprido in zip(sla_data[1::3], sla_data[2::3], sla_data[3::3])]

        # Cobertura
        cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
        cobertura_data = df.iloc[cobertura_row_index + 1, :].dropna().tolist()
        cobertura = [round(cumprido / total * 100, 2) if total else 0
                     for _, total, cumprido in zip(cobertura_data[1::3], cobertura_data[2::3], cobertura_data[3::3])]

        # Churn
        churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]
        churn_data = df.iloc[churn_row_index + 1, :].dropna().tolist()
        churn = [round(cumprido / total * 100, 2) if total else 0
                 for _, total, cumprido in zip(churn_data[1::3], churn_data[2::3], churn_data[3::3])]

        semanas = [p for p in csat_data[1::3]]  # pega só os períodos

        weekly_data = {
            "semanas": semanas,
            "csat": csat,
            "sla": sla,
            "cobertura": cobertura,
            "churn": churn
        }

        return jsonify({"success": True, "data": weekly_data})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})
