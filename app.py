# coding:utf-8
from flask import Flask, render_template, request, redirect, url_for
from flask import send_from_directory
from werkzeug.utils import secure_filename
import os
 
app = Flask(__name__)
 
@app.route('/', methods=['POST', 'GET'])
def upload_run():
    if request.method == 'POST':
        # 获取文件信息
        file1 = request.files['file1']
        file2 = request.files['file2']
        # 获取应用程序所在目录路径
        basepath = os.path.dirname(__file__)

        # 保存第一个文件
        if file1:
            upload_path1 = os.path.join(basepath, file1.filename)
            file1.save(upload_path1)
            print('File 1 uploaded.')

        # 保存第二个文件
        if file2:
            upload_path2 = os.path.join(basepath, file2.filename)
            file2.save(upload_path2)
            print('File 2 uploaded.')

        print('uploading ...')
        import subprocess
        subprocess.call(['python', 'calculate.py'])
        return '完成'
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", debug=True)
