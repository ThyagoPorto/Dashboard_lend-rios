import os
from flask import Flask, send_from_directory
from dashboard import dashboard_bp # Esta linha fica, pois você tem o dashboard.py

app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), 'static'))
app.config['SECRET_KEY'] = 'asdf#FGSgvasgf$5$WGT'

# A linha que regista o user_bp foi removida.
app.register_blueprint(dashboard_bp)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
