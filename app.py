import os
import csv
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for
from wordcloud import WordCloud
import jieba.posseg as pseg  # 引入词性标注模块
from werkzeug.utils import secure_filename
from docx import Document
import io

app = Flask(__name__)

# 设置文件上传的路径和允许上传的文件类型
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'txt', 'csv', 'tsv', 'doc', 'docx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


# 检查文件是否符合上传的格式
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# 生成仅包含形容词的词云图像
def generate_word_cloud(text):
    try:
        # 使用 jieba 分词并标注词性
        words = pseg.cut(text)

        # 过滤出词性为形容词（'a'）的词语
        adjective_words = [word.word for word in words if word.flag == 'a']
        filtered_text = " ".join(adjective_words)

        # 生成词云
        wordcloud = WordCloud(font_path="SimHei.ttf", background_color="white", max_words=100, width=800,
                              height=400).generate(filtered_text)

        # 确保 static/uploads 目录存在
        static_upload_folder = os.path.join(app.root_path, 'static', 'uploads')
        if not os.path.exists(static_upload_folder):
            os.makedirs(static_upload_folder)

        # 保存词云图像到 static/uploads 文件夹
        image_path = os.path.join(static_upload_folder, 'wordcloud.png')
        wordcloud.to_file(image_path)

        return 'uploads/wordcloud.png'
    except Exception as e:
        print(f"Error generating wordcloud: {e}")
        return None


# 读取文件内容并根据不同类型进行解析
def read_file_content(file_path, file_extension):
    text = ""
    if file_extension == 'txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
    elif file_extension in ['csv', 'tsv']:
        delimiter = ',' if file_extension == 'csv' else '\t'
        df = pd.read_csv(file_path, delimiter=delimiter)
        text = " ".join(df.astype(str).stack().tolist())  # 将所有单元格内容转换为文本
    elif file_extension == 'docx':
        doc = Document(file_path)
        text = " ".join([para.text for para in doc.paragraphs])
    elif file_extension == 'doc':
        # 对于 .doc 文件，可以使用 pypandoc 或者其他转换工具将 .doc 文件转换为 .docx，再解析
        # pypandoc 例子：需安装 pypandoc 和 pandoc
        # import pypandoc
        # text = pypandoc.convert_file(file_path, 'plain')
        pass
    return text


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 检查是否有上传文件
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']

        # 检查文件名是否为空并符合格式
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # 获取文件扩展名
            file_extension = file.filename.rsplit('.', 1)[1].lower()

            # 读取文件内容
            text = read_file_content(file_path, file_extension)

            if not text:
                return "文件格式不支持或解析失败"

            # 生成词云
            image_path = generate_word_cloud(text)

            # 显示生成的词云
            return render_template('index.html', wordcloud_image=image_path)

    return render_template('index.html')


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
