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

        # METAS FIXAS DO MÊS (conforme informado)
        metas_fixas = {
            "CSAT": 95,
            "SLA dos DS": 82,
            "Cobertura de Carteira": 100,
            "Cancelamento - Churn": 1
        }

        for metrica in metricas:
            row_index = df[df.iloc[:, 0] == metrica].index[0]

            # USAR META FIXA DO MÊS
            meta_val = metas_fixas.get(metrica, 0)
            meta_percentual = meta_val / 100
            meta_objetivo = f"{meta_val}%"

            # Ler dados do mês das colunas O, P (índices 14, 15)
            total_mes_total = _to_number(df.iloc[row_index + 1, 14])  # Coluna O
            total_mes_cumprido = _to_number(df.iloc[row_index + 1, 15])  # Coluna P
            
            # Calcular percentual manualmente
            if total_mes_total > 0:
                total_mes_percentual = total_mes_cumprido / total_mes_total
            else:
                total_mes_percentual = 0

            # Ler dados semanais das colunas B, E, H, K
            semanas = []
            colunas_semanas = [1, 4, 7, 10]  # Colunas B, E, H, K
            
            for col_index in colunas_semanas:
                try:
                    periodo = df.iloc[row_index, col_index]
                    if pd.isna(periodo) or periodo == "":
                        continue
                        
                    total = _to_number(df.iloc[row_index + 1, col_index])
                    cumprido = _to_number(df.iloc[row_index + 1, col_index + 1])
                    
                    if total > 0:
                        percentual = cumprido / total
                    else:
                        percentual = 0
                        
                    semanas.append({
                        "periodo": str(periodo),
                        "total": total,
                        "cumprido": cumprido,
                        "percentual": percentual
                    })
                except Exception as e:
                    print(f"Erro ao processar semana: {e}")
                    continue

            # Determinar status baseado no percentual e meta
            status_final = _get_status(total_mes_percentual, meta_percentual)

            metrics_data.append({
                "metrica": metrica,
                "cor": cores.get(metrica, "#000"),
                "meta_objetivo": meta_objetivo,
                "meta_percentual": meta_percentual,
                "meta_valor": meta_val,  # Adicionando o valor numérico da meta
                "semanas": semanas,
                "total_mes": {
                    "total": total_mes_total,
                    "cumprido": total_mes_cumprido,
                    "percentual": total_mes_percentual
                },
                "status": status_final
            })

        return jsonify({"success": True, "data": metrics_data})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


def _to_number(val):
    """Converte qualquer valor em número."""
    if pd.isna(val):
        return 0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        # Remove porcentagens, vírgulas, etc.
        cleaned = str(val).replace("%", "").replace(",", ".").strip()
        return float(cleaned)
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

        weekly_data = {"semanas": []}
        semanas_coletadas = False

        for metrica in metricas:
            row_index = df[df.iloc[:, 0] == metrica].index[0]
            
            # Coletar dados das semanas
            valores = []
            colunas_semanas = [1, 4, 7, 10]  # Colunas B, E, H, K
            
            for col_index in colunas_semanas:
                try:
                    periodo = df.iloc[row_index, col_index]
                    if pd.isna(periodo) or periodo == "":
                        continue
                        
                    total = _to_number(df.iloc[row_index + 1, col_index])
                    cumprido = _to_number(df.iloc[row_index + 1, col_index + 1])
                    
                    if total > 0:
                        percentual = (cumprido / total) * 100
                    else:
                        percentual = 0
                        
                    valores.append(percentual)
                    
                    # Coletar labels das semanas apenas uma vez
                    if not semanas_coletadas:
                        weekly_data["semanas"].append(str(periodo))
                        
                except Exception:
                    continue
            
            semanas_coletadas = True
            weekly_data[metrica.lower().replace(" ", "_")] = valores

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

        # METAS FIXAS DO MÊS (conforme informado)
        metas_fixas = {
            "CSAT": 95,
            "SLA dos DS": 82,
            "Cobertura de Carteira": 100,
            "Cancelamento - Churn": 1
        }

        for metrica in metricas:
            try:
                row_index = df[df.iloc[:, 0] == metrica].index[0]
                
                # USAR META FIXA DO MÊS
                meta_val = metas_fixas.get(metrica, 0)
                meta_percentual = meta_val / 100
                
                # Ler dados do mês das colunas O, P
                total_mes_total = _to_number(df.iloc[row_index + 1, 14])
                total_mes_cumprido = _to_number(df.iloc[row_index + 1, 15])
                
                if total_mes_total > 0:
                    percentual = total_mes_cumprido / total_mes_total
                else:
                    percentual = 0
                
                # Verificar se atingiu a meta
                if "Churn" in metrica:
                    if percentual <= meta_percentual:
                        metrics_above_target += 1
                    else:
                        metrics_below_target += 1
                else:
                    if percentual >= meta_percentual:
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
