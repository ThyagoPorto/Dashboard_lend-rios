import os
from flask import Flask, send_from_directory
from dashboard import dashboard_bp

# Configuração do Flask
app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'asdf#FGSgvasgf$5$WGT'

# Registrar o blueprint do dashboard (API) com prefixo
app.register_blueprint(dashboard_bp, url_prefix="/api/dashboard")

# Rota raiz → serve o index.html
@app.route("/")
def serve_index():
    return send_from_directory(app.static_folder, "index.html")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
