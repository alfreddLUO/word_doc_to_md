import mammoth
from markdownify import markdownify
import time

# 转存 Word 文档内的图片
def convert_img(image):
    with image.open() as image_bytes:
        file_suffix = image.content_type.split("/")[1]
        path_file = "./img/{}.{}".format(str(time.time()),file_suffix)
        with open(path_file, 'wb') as f:
            f.write(image_bytes.read())

    return {"src":path_file}

# 读取 Word 文件
def convert_word_doc_to_md(word_filename,html_filename, md_filename):
    with open(word_filename ,"rb") as docx_file:
        # 转化 Word 文档为 HTML
        result = mammoth.convert_to_html(docx_file,convert_image=mammoth.images.img_element(convert_img))
        # 获取 HTML 内容
        html = result.value
        # 转化 HTML 为 Markdown
        md = markdownify(html,heading_style="ATX")
        print(md)
        with open(html_filename,'w',encoding='utf-8') as html_file,open(md_filename,"w",encoding='utf-8') as md_file:
            html_file.write(html)
            md_file.write(md)
        messages = result.messages

convert_word_doc_to_md("./WordDocs/4486_HW3_GP14.docx","./HtmlDocs/4486_HW3_GP14.html", "./MdDocs/4486_HW3_GP14.md")
