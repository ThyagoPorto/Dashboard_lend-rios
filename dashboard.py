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

        # PERÍODOS DAS SEMANAS FIXOS (conforme informado)
        periodos_semanas = [
            "01 a 05",
            "08 a 12", 
            "15 a 19",
            "22 a 30"
        ]

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

            # BUSCAR DADOS DAS SEMANAS - USANDO PERÍODOS FIXOS
            semanas = []
            # Colunas para cada semana: [B,C], [E,F], [H,I], [K,L] (Total, Cumprido)
            colunas_semanas = [
                (1, 2),   # B, C - Semana 01 a 05 (Total, Cumprido)
                (4, 5),   # E, F - Semana 08 a 12 (Total, Cumprido)
                (7, 8),   # H, I - Semana 15 a 19 (Total, Cumprido)
                (10, 11)  # K, L - Semana 22 a 30 (Total, Cumprido)
            ]
            
            for i, (col_total, col_cumprido) in enumerate(colunas_semanas):
                try:
                    periodo = periodos_semanas[i]  # Usar período fixo
                    
                    total = _to_number(df.iloc[row_index + 1, col_total])
                    cumprido = _to_number(df.iloc[row_index + 1, col_cumprido])
                    
                    # Só criar card se tiver dados válidos
                    if total > 0 or cumprido >= 0:
                        percentual = cumprido / total if total > 0 else 0
                        
                        semanas.append({
                            "periodo": periodo,
                            "total": total,
                            "cumprido": cumprido,
                            "percentual": percentual
                        })
                        
                except Exception as e:
                    print(f"Erro ao processar semana {periodos_semanas[i]}: {e}")
                    continue

            # Determinar status baseado no percentual e meta
            status_final = _get_status(total_mes_percentual, meta_percentual)

            metrics_data.append({
                "metrica": metrica,
                "cor": cores.get(metrica, "#000"),
                "meta_objetivo": meta_objetivo,
                "meta_percentual": meta_percentual,
                "meta_valor": meta_val,
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

        # PERÍODOS DAS SEMANAS FIXOS
        periodos_semanas = [
            "01 a 05",
            "08 a 12", 
            "15 a 19",
            "22 a 30"
        ]

        weekly_data = {"semanas": periodos_semanas}  # Usar períodos fixos

        # Colunas para cada semana: [B,C], [E,F], [H,I], [K,L] (Total, Cumprido)
        colunas_semanas = [
            (1, 2),   # B, C - Semana 01 a 05
            (4, 5),   # E, F - Semana 08 a 12
            (7, 8),   # H, I - Semana 15 a 19
            (10, 11)  # K, L - Semana 22 a 30
        ]

        for metrica in metricas:
            row_index = df[df.iloc[:, 0] == metrica].index[0]
            
            # Coletar dados das semanas
            valores = []
            
            for col_total, col_cumprido in colunas_semanas:
                try:
                    total = _to_number(df.iloc[row_index + 1, col_total])
                    cumprido = _to_number(df.iloc[row_index + 1, col_cumprido])
                    
                    if total > 0:
                        percentual = (cumprido / total) * 100
                    else:
                        percentual = 0
                        
                    valores.append(percentual)
                        
                except Exception:
                    continue
            
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
