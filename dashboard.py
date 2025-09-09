import os
import pandas as pd
from flask import Blueprint, jsonify

dashboard_bp = Blueprint("dashboard", __name__, url_prefix="/api/dashboard")

# Caminho do Excel
excel_path = os.path.join(os.path.dirname(__file__), "upload", "dados.xlsx")


def load_excel():
    try:
        return pd.ExcelFile(excel_path)
    except Exception:
        return None


@dashboard_bp.route("/metrics", methods=["GET"])
def get_metrics():
    try:
        xls = load_excel()
        if not xls:
            return jsonify({"success": False, "error": "Erro ao carregar Excel"})

        df = pd.read_excel(xls, sheet_name="Métricas")

        metrics_data = []
        metricas = ["CSAT", "SLA dos DS", "Cobertura de Carteira", "Cancelamento - Churn"]
        cores = {
            "CSAT": "#00D9FF",
            "SLA dos DS": "#57A9FB",
            "Cobertura de Carteira": "#8B5CF6",
            "Cancelamento - Churn": "#FF6B35"
        }

        for metrica in metricas:
            row_index = df[df.iloc[:, 0] == metrica].index[0]

            # Meta (%)
            meta_val = str(df.iloc[row_index, 1]).replace("%", "").strip()
            try:
                meta_percentual = float(meta_val) / 100
            except:
                meta_percentual = 0
            meta_objetivo = f"{meta_val}%"

            # Status
            status = str(df.iloc[row_index, 2])

            # Linha de dados logo abaixo
            data_row = df.iloc[row_index + 1, 1:].dropna().tolist()

            semanas = []
            for i in range(0, len(data_row), 3):
                try:
                    periodo = data_row[i]
                    total = _to_number(data_row[i + 1])
                    cumprido = _to_number(data_row[i + 2])
                    percentual = round(cumprido / total, 4) if total else 0
                    semanas.append({
                        "periodo": periodo,
                        "total": total,
                        "cumprido": cumprido,
                        "percentual": percentual
                    })
                except Exception:
                    break

            total_mes = semanas[-1] if semanas else {"total": 0, "cumprido": 0, "percentual": 0}
            status_final = _get_status(total_mes["percentual"], meta_percentual)

            metrics_data.append({
                "metrica": metrica,
                "cor": cores.get(metrica, "#000"),
                "meta_objetivo": meta_objetivo,
                "meta_percentual": meta_percentual,
                "semanas": semanas,
                "total_mes": total_mes,
                "status": status_final
            })

        return jsonify({"success": True, "data": metrics_data})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


def _to_number(val):
    """Converte strings como '95%' ou '131' em float/int."""
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return val
    try:
        return float(str(val).replace("%", "").replace(",", ".").strip())
    except:
        return 0


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
        metricas = ["CSAT", "SLA dos DS", "Cobertura de Carteira", "Cancelamento - Churn"]

        semanas_labels = None
        weekly_data = {"semanas": []}

        for metrica in metricas:
            row_index = df[df.iloc[:, 0] == metrica].index[0]
            data_row = df.iloc[row_index + 1, 1:].dropna().tolist()

            periodos, valores = [], []
            for i in range(0, len(data_row), 3):
                try:
                    periodo = data_row[i]
                    total = _to_number(data_row[i + 1])
                    cumprido = _to_number(data_row[i + 2])
                    percentual = round((cumprido / total) * 100, 2) if total else 0
                    periodos.append(periodo)
                    valores.append(percentual)
                except Exception:
                    break

            if semanas_labels is None:
                semanas_labels = periodos

            weekly_data[metrica.lower().replace(" ", "_")] = valores

        weekly_data["semanas"] = semanas_labels
        return jsonify({"success": True, "data": weekly_data})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})
