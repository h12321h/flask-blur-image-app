from flask import Flask, render_template, request, send_file
from docx import Document
from PIL import Image, ImageFilter
import io
import os

app = Flask(__name__)


# 模糊处理函数
def blur_images_in_docx(docx_stream):
    doc = Document(docx_stream)

    for rel in doc.part.rels:
        if "image" in doc.part.rels[rel].target_ref:  # 找到所有图片
            image = doc.part.rels[rel].target_part.blob  # 读取图片内容
            image_io = io.BytesIO(image)  # 转换为二进制流
            img = Image.open(image_io)  # 使用Pillow打开图片

            # 对图片应用模糊滤镜
            blurred_img = img.filter(ImageFilter.GaussianBlur(radius=5))  # radius调整模糊程度

            # 保存模糊后的图片到二进制流
            img_byte_arr = io.BytesIO()
            blurred_img.save(img_byte_arr, format=img.format)

            # 替换原图片
            doc.part.rels[rel].target_part._blob = img_byte_arr.getvalue()

    # 将处理后的文档保存到内存流
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)

    return output_stream


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'

    file = request.files['file']

    if file.filename == '':
        return 'No selected file'

    if file:
        # 获取原文件名和扩展名
        original_filename = os.path.splitext(file.filename)[0]
        file_extension = os.path.splitext(file.filename)[1]

        # 处理文件并返回处理后的文件
        output_stream = blur_images_in_docx(file)

        # 构造处理后的文件名
        new_filename = f"{original_filename}（模糊图片）{file_extension}"

        return send_file(output_stream, as_attachment=True, download_name=new_filename,
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    app.run(debug=True)
