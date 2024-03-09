from flask import Flask, render_template, send_from_directory
import os
import excel_creater
from datetime import datetime
app = Flask(__name__, static_folder='download')

global dt, exist
dt=datetime.now()
exist=False

@app.route('/',methods=['POST', 'GET'])
def index():
    global dt, exist
    time_difference = datetime.now() - dt
    if time_difference.total_seconds() > 86400:
        excel_creater.excel_create(app.root_path)
        dt = datetime.now()
        exist = True
    if not exist:
        excel_creater.excel_create(app.root_path)
        dt = datetime.now()
        exist = True
    return render_template("index.html")

@app.route('/download/<path:filename>', methods=['GET', 'POST'])
def download(filename):
    downloads = os.path.join(app.root_path, 'download')
    return send_from_directory(directory=downloads, path='/', filename=filename)

app.run("127.0.0.1", 500, debug=True)
