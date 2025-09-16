from flask import Flask, jsonify
import random

app = Flask(__name__)

@app.route('/api/dashboard/metrics')
def get_metrics():
    # Exemplo de dados estáticos (substituir pela lógica real)
    metrics = [
        {
            "metrica": "CSAT",
            "meta_percentual": 0.9,
            "total_mes": {"total": 100, "cumprido": 90, "percentual": 0.9},
            "status": "Excelente",
            "cor": "#00D9FF"
        },
        {
            "metrica": "SLA dos DS",
            "meta_percentual": 0.95,
            "total_mes": {"total": 200, "cumprido": 180, "percentual": 0.9},
            "status": "Atenção",
            "cor": "#FF6B35"
        },
        {
            "metrica": "Cobertura de Carteira",
            "meta_percentual": 0.85,
            "total_mes": {"total": 150, "cumprido": 120, "percentual": 0.8},
            "status": "Crítico",
            "cor": "#8B5CF6"
        },
        {
            "metrica": "Cancelamento - Churn",
            "meta_percentual": 0.05,
            "total_mes": {"total": 50, "cumprido": 3, "percentual": 0.06},
            "status": "Atenção",
            "cor": "#FBE842"
        }
    ]
    return jsonify({"success": True, "data": metrics})

@app.route('/api/dashboard/weekly-data')
def get_weekly_data():
    # Simulação de dados semanais (substituir pela lógica real)
    semanas = ["01 a 05", "08 a 12", "15 a 19", "22 a 30"]
    csat = []
    sla = []
    cobertura = []
    churn = []

    for _ in semanas:
        total = random.randint(50, 100)
        cumprido = random.randint(0, total)
        csat.append(cumprido / total if total > 0 else 0)

        total = random.randint(50, 100)
        cumprido = random.randint(0, total)
        sla.append(cumprido / total if total > 0 else 0)

        total = random.randint(50, 100)
        cumprido = random.randint(0, total)
        cobertura.append(cumprido / total if total > 0 else 0)

        total = random.randint(50, 100)
        cumprido = random.randint(0, total)
        churn.append(cumprido / total if total > 0 else 0)

    return jsonify({
        "success": True,
        "data": {
            "semanas": semanas,
            "csat": csat,
            "sla": sla,
            "cobertura": cobertura,
            "churn": churn
        }
    })

if __name__ == '__main__':
    app.run(debug=True)
