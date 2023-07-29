# coding:utf-8
from flask import Flask, render_template, request, redirect, url_for
from flask import send_from_directory
from werkzeug.utils import secure_filename
import os
 
app = Flask(__name__)
 
@app.route('/', methods=['POST', 'GET'])
def upload_run():
    if request.method == 'POST':
        f = request.files['file']
        basepath = os.path.dirname(__file__)
        upload_path = os.path.join(basepath, './', f.filename)
        f.save(upload_path)
        print('uploading ...')
        import subprocess
        subprocess.call(['python', 'calculate.py'])
        return '完成'
    return render_template('index.html')

# @app.route('/run', methods=['POST'])
# def run():
#     import subprocess
#     subprocess.call(['python', 'calculate.py'])
#     return render_template('index.html')

 
# @app.route('/download')
# def download():
#     print('downloading ...')
#     return send_from_directory(r"/root/", "test.sh", as_attachment=True)
 
 
if __name__ == '__main__':
    app.run(host="0.0.0.0", debug=True)