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


def _get_status(percentual, meta, metrica=""):
    """
    Determina o status baseado no percentual e meta.
    Para Churn: valores BAIXOS são bons (≤ meta)
    Para outras métricas: valores ALTOS são bons (≥ meta)
    """
    if "Churn" in metrica:
        # Lógica INVERTIDA para Churn - quanto MENOR, melhor
        if percentual <= meta:
            return "Excelente"
        elif percentual <= meta * 1.1:  # Até 10% acima da meta ainda é Atenção
            return "Atenção"
        else:
            return "Crítico"
    else:
        # Lógica NORMAL para outras métricas - quanto MAIOR, melhor
        if percentual >= meta:
            return "Excelente"
        elif percentual >= meta * 0.9:  # Pelo menos 90% da meta é Atenção
            return "Atenção"
        else:
            return "Crítico"


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
                        
                        # Determinar status para a semana (passando o nome da métrica)
                        status_semana = _get_status(percentual, meta_percentual, metrica)
                        
                        semanas.append({
                            "periodo": periodo,
                            "total": total,
                            "cumprido": cumprido,
                            "percentual": percentual,
                            "status": status_semana  # Adicionar status para cada semana
                        })
                        
                except Exception as e:
                    print(f"Erro ao processar semana {periodos_semanas[i]}: {e}")
                    continue

            # Determinar status baseado no percentual e meta (passando o nome da métrica)
            status_final = _get_status(total_mes_percentual, meta_percentual, metrica)

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

@dashboard_bp.route("/weekly-data", methods=["GET"])
def get_weekly_data():
    try:
        xls = load_excel()
        if not xls:
            return jsonify({"success": False, "error": "Erro ao carregar Excel"})

        df = pd.read_excel(xls, sheet_name="Métricas")

        # PERÍODOS DAS SEMANAS FIXOS
        periodos_semanas = [
            "01 a 05",
            "08 a 12", 
            "15 a 19",
            "22 a 30"
        ]

        weekly_data = {"semanas": periodos_semanas}

        # DEBUG: Verificar a estrutura do DataFrame
        print("Colunas do DataFrame:", df.columns.tolist())
        print("Primeiras linhas do DataFrame:")
        print(df.head(20))

        # Mapeamento CORRETO das células onde estão os percentuais
        metric_cells = {
            "CSAT": [
                df.iloc[3, 3],   # D3 - Semana 01 a 05
                df.iloc[3, 6],   # G3 - Semana 08 a 12  
                df.iloc[3, 9],   # J3 - Semana 15 a 19
                df.iloc[3, 12]   # M3 - Semana 22 a 30
            ],
            "SLA dos DS": [
                df.iloc[7, 3],   # D7 - Semana 01 a 05
                df.iloc[7, 6],   # G7 - Semana 08 a 12
                df.iloc[7, 9],   # J7 - Semana 15 a 19
                df.iloc[7, 12]   # M7 - Semana 22 a 30
            ],
            "Cobertura de Carteira": [
                df.iloc[11, 3],  # D11 - Semana 01 a 05
                df.iloc[11, 6],  # G11 - Semana 08 a 12
                df.iloc[11, 9],  # J11 - Semana 15 a 19
                df.iloc[11, 12]  # M11 - Semana 22 a 30
            ],
            "Cancelamento - Churn": [
                df.iloc[15, 3],  # D15 - Semana 01 a 05
                df.iloc[15, 6],  # G15 - Semana 08 a 12
                df.iloc[15, 9],  # J15 - Semana 15 a 19
                df.iloc[15, 12]  # M15 - Semana 22 a 30
            ]
        }

        for metrica, valores in metric_cells.items():
            percentuais = []
            
            for valor in valores:
                try:
                    # Converter o valor para número
                    valor_num = _to_number(valor)
                    print(f"{metrica} - valor extraído: {valor} -> convertido: {valor_num}")
                    
                    # Para Churn, a lógica é invertida
                    if "Churn" in metrica:
                        # Se for churn, valores menores são melhores
                        # Ex: 0.0057 significa 0.57% de churn → 99.43% de atingimento
                        percentual = (1 - valor_num) * 100
                    else:
                        # Para outras métricas, valores maiores são melhores
                        percentual = valor_num * 100
                    
                    percentuais.append(percentual)
                        
                except Exception as e:
                    print(f"Erro em {metrica}: {e}")
                    percentuais.append(0)
            
            print(f"{metrica} - percentuais: {percentuais}")
            
            # Adicionar ao weekly_data
            metric_key = metrica.lower().replace(" ", "_").replace("-", "_")
            weekly_data[metric_key] = percentuais

        print("Dados finais:", weekly_data)
        return jsonify({"success": True, "data": weekly_data})

    except Exception as e:
        print(f"Erro geral: {e}")
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
                
                # Verificar se atingiu a meta (lógica invertida para Churn)
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

