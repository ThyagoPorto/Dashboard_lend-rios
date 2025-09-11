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

            # === META ===
            raw_meta = str(df.iloc[row_index, 1]).replace("%", "").strip()
            meta_val = _to_number(raw_meta)
            if meta_val == 0:
                # se não encontrou número válido na célula da meta,
                # procura na linha inteira o primeiro valor numérico > 0
                row_values = df.iloc[row_index].tolist()
                for v in row_values:
                    num = _to_number(v)
                    if num > 0:
                        meta_val = num
                        break
            meta_percentual = meta_val / 100 if meta_val else 0
            meta_objetivo = f"{meta_val}%" if meta_val else "0%"

            # === STATUS ===
            status = str(df.iloc[row_index, 2]) if len(df.columns) > 2 else "-"

            # === DADOS SEMANAIS ===
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
    """Converte qualquer valor em número (95% -> 95.0, '131' -> 131.0)."""
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return float(val)
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


@dashboard_bp.route("/summary", methods=["GET"])
def get_summary():
    try:
        xls = load_excel()
        if not xls:
            return jsonify({"success": False, "error": "Erro ao carregar Excel"})

        df = pd.read_excel(xls, sheet_name="Métricas")
        metricas = ["CSAT", "SLA dos DS", "Cobertura de Carteira", "Cancelamento - Churn"]

        total_metrics = len(metricas)
        metrics_above_target = 0
        metrics_below_target = 0

        for metrica in metricas:
            try:
                row_index = df[df.iloc[:, 0] == metrica].index[0]
                
                # Obter meta
                raw_meta = str(df.iloc[row_index, 1]).replace("%", "").strip()
                meta_val = _to_number(raw_meta)
                meta_percentual = meta_val / 100 if meta_val else 0
                
                # Obter dados do mês
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
                
                if semanas:
                    total_mes = semanas[-1]
                    # Verificar se atingiu a meta (lógica invertida para Churn)
                    if "Churn" in metrica:
                        if total_mes["percentual"] <= meta_percentual:
                            metrics_above_target += 1
                        else:
                            metrics_below_target += 1
                    else:
                        if total_mes["percentual"] >= meta_percentual:
                            metrics_above_target += 1
                        else:
                            metrics_below_target += 1
                            
            except Exception:
                continue

        performance_percentage = (metrics_above_target / total_metrics) * 100 if total_metrics > 0 else 0

        return jsonify({
            "success": True,
            "data": {
                "total_metrics": total_metrics,
                "metrics_above_target": metrics_above_target,
                "metrics_below_target": metrics_below_target,
                "performance_percentage": performance_percentage
            }
        })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})
