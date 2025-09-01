from flask import Blueprint, jsonify, redirect, url_for
import pandas as pd
import os

dashboard_bp = Blueprint("dashboard", __name__)

@dashboard_bp.route("/", methods=["GET"])
def index():
    return redirect(url_for('dashboard.get_metrics'))

@dashboard_bp.route("/metrics", methods=["GET"])
def get_metrics():
    """Retorna os dados das métricas da equipe Lendários com dados semanais completos"""
    try:
        excel_path = os.path.join(os.getcwd(), "dados.xlsx")
        df = pd.read_excel(excel_path, header=None)

        csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
        csat_data = df.iloc[csat_row_index + 1, :].dropna().tolist()

        sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
        sla_data = df.iloc[sla_row_index + 1, :].dropna().tolist()

        cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
        cobertura_data = df.iloc[cobertura_row_index + 1, :].dropna().tolist()

        churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]
        churn_data = df.iloc[churn_row_index + 1, :].dropna().tolist()

        rnr_row_index = df[df.iloc[:, 0] == "RNR"].index[0]
        rnr_data = df.iloc[rnr_row_index + 1, :].dropna().tolist()

        csat_semanas = [
            {"periodo": "04 a 08", "total": csat_data[0], "cumprido": csat_data[1], "percentual": csat_data[2]},
            {"periodo": "11 a 15", "total": csat_data[3], "cumprido": csat_data[4], "percentual": csat_data[5]},
            {"periodo": "18 a 22", "total": csat_data[6], "cumprido": csat_data[7], "percentual": csat_data[8]},
            {"periodo": "25 a 29", "total": csat_data[9], "cumprido": csat_data[10], "percentual": csat_data[11]},
            {"periodo": "Total Mês", "total": csat_data[12], "cumprido": csat_data[13], "percentual": csat_data[14]}
        ]

        sla_semanas = [
            {"periodo": "04 a 08", "total": sla_data[0], "cumprido": sla_data[1], "percentual": sla_data[2]},
            {"periodo": "11 a 15", "total": sla_data[3], "cumprido": sla_data[4], "percentual": sla_data[5]},
            {"periodo": "18 a 22", "total": sla_data[6], "cumprido": sla_data[7], "percentual": sla_data[8]},
            {"periodo": "25 a 29", "total": sla_data[9], "cumprido": sla_data[10], "percentual": sla_data[11]},
            {"periodo": "Total Mês", "total": sla_data[12], "cumprido": sla_data[13], "percentual": sla_data[14]}
        ]

        cobertura_semanas = [
            {"periodo": "04 a 08", "total": cobertura_data[0], "cumprido": cobertura_data[1], "percentual": cobertura_data[2]},
            {"periodo": "11 a 15", "total": cobertura_data[3], "cumprido": cobertura_data[4], "percentual": cobertura_data[5]},
            {"periodo": "18 a 22", "total": cobertura_data[6], "cumprido": cobertura_data[7], "percentual": cobertura_data[8]},
            {"periodo": "25 a 29", "total": cobertura_data[9], "cumprido": cobertura_data[10], "percentual": cobertura_data[11]},
            {"periodo": "Total Mês", "total": cobertura_data[12], "cumprido": cobertura_data[13], "percentual": cobertura_data[14]}
        ]

        churn_semanas = [
            {"periodo": "04 a 08", "total": churn_data[0], "cumprido": churn_data[1], "percentual": churn_data[2]},
            {"periodo": "11 a 15", "total": churn_data[3], "cumprido": churn_data[4], "percentual": churn_data[5]},
            {"periodo": "18 a 22", "total": churn_data[6], "cumprido": churn_data[7], "percentual": churn_data[8]},
            {"periodo": "25 a 29", "total": churn_data[9], "cumprido": churn_data[10], "percentual": churn_data[11]},
            {"periodo": "Total Mês", "total": churn_data[12], "cumprido": churn_data[13], "percentual": churn_data[14]}
        ]

        rnr_semanas = [
            {"periodo": "04 a 08", "total": rnr_data[0], "cumprido": rnr_data[1], "percentual": rnr_data[2]},
            {"periodo": "11 a 15", "total": rnr_data[3], "cumprido": rnr_data[4], "percentual": rnr_data[5]},
            {"periodo": "18 a 22", "total": rnr_data[6], "cumprido": rnr_data[7], "percentual": rnr_data[8]},
            {"periodo": "25 a 29", "total": rnr_data[9], "cumprido": rnr_data[10], "percentual": rnr_data[11]},
            {"periodo": "Total Mês", "total": rnr_data[12], "cumprido": rnr_data[13], "percentual": rnr_data[14]}
        ]

        metrics_data = [
            {"metrica": "CSAT", "meta_objetivo": "95%", "meta_percentual": 0.95, "semanas": csat_semanas, "total_mes": {"total": csat_data[12], "cumprido": csat_data[13], "percentual": csat_data[14]}, "status": "Excelente" if float(csat_data[14]) >= 0.95 else "Atenção", "cor": "#00D9FF" if float(csat_data[14]) >= 0.95 else "#FF6B35"},
            {"metrica": "SLA dos DS", "meta_objetivo": "82%", "meta_percentual": 0.82, "semanas": sla_semanas, "total_mes": {"total": sla_data[12], "cumprido": sla_data[13], "percentual": sla_data[14]}, "status": "Excelente" if float(sla_data[14]) >= 0.82 else "Atenção", "cor": "#00D9FF" if float(sla_data[14]) >= 0.82 else "#FF6B35"},
            {"metrica": "Cobertura de Carteira", "meta_objetivo": "100%", "meta_percentual": 1.0, "semanas": cobertura_semanas, "total_mes": {"total": cobertura_data[12], "cumprido": cobertura_data[13], "percentual": cobertura_data[14]}, "status": "Excelente" if float(cobertura_data[14]) >= 1.0 else "Crítico" if float(cobertura_data[14]) < 0.6 else "Atenção", "cor": "#00D9FF" if float(cobertura_data[14]) >= 1.0 else "#8B5CF6" if float(cobertura_data[14]) < 0.6 else "#FF6B35"},
            {"metrica": "Cancelamento - Churn", "meta_objetivo": "1%", "meta_percentual": 0.01, "semanas": churn_semanas, "total_mes": {"total": churn_data[12], "cumprido": churn_data[13], "percentual": churn_data[14]}, "status": "Excelente" if float(churn_data[14]) <= 0.01 else "Atenção", "cor": "#00D9FF" if float(churn_data[14]) <= 0.01 else "#FF6B35"},
            {"metrica": "RNR", "meta_objetivo": "22.5%", "meta_percentual": 0.225, "semanas": rnr_semanas, "total_mes": {"total": rnr_data[12], "cumprido": rnr_data[13], "percentual": rnr_data[14]}, "status": "Excelente" if float(rnr_data[14]) <= 0.225 else "Atenção", "cor": "#00D9FF" if float(rnr_data[14]) <= 0.225 else "#FF6B35"}
        ]
        
        return jsonify({"success": True, "data": metrics_data})
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@dashboard_bp.route("/summary", methods=["GET"])
def get_summary():
    """Retorna um resumo geral das métricas"""
    try:
        excel_path = os.path.join(os.getcwd(), "dados.xlsx")
        df = pd.read_excel(excel_path, header=None)

        csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
        csat_data = df.iloc[csat_row_index + 1, :].dropna().tolist()

        sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
        sla_data = df.iloc[sla_row_index + 1, :].dropna().tolist()

        cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
        cobertura_data = df.iloc[cobertura_row_index + 1, :].dropna().tolist()

        churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]
        churn_data = df.iloc[churn_row_index + 1, :].dropna().tolist()

        rnr_row_index = df[df.iloc[:, 0] == "RNR"].index[0]
        rnr_data = df.iloc[rnr_row_index + 1, :].dropna().tolist()

        csat_meta_percent = csat_data[14]
        sla_meta_percent = sla_data[14]
        cobertura_meta_percent = cobertura_data[14]
        churn_meta_percent = churn_data[14]
        rnr_meta_percent = rnr_data[14]
        
        total_metrics = 5
        metrics_above_target = 0
        
        if float(csat_meta_percent) >= 0.95: metrics_above_target += 1
        if float(sla_meta_percent) >= 0.82: metrics_above_target += 1
        if float(cobertura_meta_percent) >= 1.0: metrics_above_target += 1
        if float(churn_meta_percent) <= 0.01: metrics_above_target += 1
        if float(rnr_meta_percent) <= 0.225: metrics_above_target += 1
        
        metrics_below_target = total_metrics - metrics_above_target
        
        return jsonify({"success": True, "data": {"total_metrics": total_metrics, "metrics_above_target": metrics_above_target, "metrics_below_target": metrics_below_target, "performance_percentage": (metrics_above_target / total_metrics) * 100, "team_name": "Lendários"}})
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@dashboard_bp.route("/weekly-data", methods=["GET"])
def get_weekly_data():
    """Retorna dados semanais para gráficos de tendência"""
    try:
        excel_path = os.path.join(os.getcwd(), "dados.xlsx")
        df = pd.read_excel(excel_path, header=None)

        csat_row_index = df[df.iloc[:, 0] == "CSAT"].index[0]
        csat_data = df.iloc[csat_row_index + 1, :].dropna().tolist()

        sla_row_index = df[df.iloc[:, 0] == "SLA dos DS"].index[0]
        sla_data = df.iloc[sla_row_index + 1, :].dropna().tolist()

        cobertura_row_index = df[df.iloc[:, 0] == "Cobertura de Carteira"].index[0]
        cobertura_data = df.iloc[cobertura_row_index + 1, :].dropna().tolist()

        churn_row_index = df[df.iloc[:, 0] == "Cancelamento - Churn"].index[0]
        churn_data = df.iloc[churn_row_index + 1, :].dropna().tolist()

        rnr_row_index = df[df.iloc[:, 0] == "RNR"].index[0]
        rnr_data = df.iloc[rnr_row_index + 1, :].dropna().tolist()

        weekly_data = {
            "semanas": ["04-08", "11-15", "18-22", "25-29"],
            "csat": [float(csat_data[2]) * 100, float(csat_data[5]) * 100, float(csat_data[8]) * 100, float(csat_data[11]) * 100 if csat_data[11] != 0 else 0],
            "sla": [float(sla_data[2]) * 100, float(sla_data[5]) * 100, float(sla_data[8]) * 100, float(sla_data[11]) * 100 if sla_data[11] != 0 else 0],
            "cobertura": [float(cobertura_data[2]) * 100, float(cobertura_data[5]) * 100, float(cobertura_data[8]) * 100, float(cobertura_data[11]) * 100 if cobertura_data[11] != 0 else 0],
            "churn": [float(churn_data[2]) * 100, float(churn_data[5]) * 100, float(churn_data[8]) * 100, float(churn_data[11]) * 100 if churn_data[11] != 0 else 0],
            "rnr": [float(rnr_data[2]) * 100, float(rnr_data[5]) * 100, float(rnr_data[8]) * 100, float(rnr_data[11]) * 100 if rnr_data[11] != 0 else 0]
        }
        
        return jsonify({"success": True, "data": weekly_data})
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
