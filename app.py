# coding:utf-8
from flask import Flask, render_template, request
import os
 
app = Flask(__name__)

# 设置允许上传的文件类型
ALLOWED_EXTENSIONS = {'xlsx'}

# 检查文件类型是否允许上传
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
 
@app.route('/', methods=['POST', 'GET'])
def upload_files():
    if request.method == 'POST':
        files = request.files.getlist('file')
        # 获取应用程序所在目录路径
        basepath = os.path.dirname(__file__)
        upload_folder = os.path.join(basepath, 'uploads')  # 指定上传文件夹的路径

        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)  # 如果上传文件夹不存在，创建它

        for file in files:
            if file and allowed_file(file.filename):
                filename = file.filename
                upload_path = os.path.join(upload_folder, filename)  # 确定文件保存的路径
                file.save(upload_path)
                print(f'File "{filename}" uploaded to "{upload_path}".')

        print('uploading ...')
        import subprocess
        subprocess.call(['python', 'calculate.py'])
        return '完成'
    return render_template('index.html')

if __name__ == '__main__':
    app.run(host="0.0.0.0", debug=True)
