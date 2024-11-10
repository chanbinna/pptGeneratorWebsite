from flask import Flask, render_template, request, send_file, redirect, url_for
from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from pptx.dml.color import RGBColor
import os
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        content = request.form.get("content")

        if file and file.filename:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            pptx_filepath = create_ppt_from_text(filepath)
            os.remove(filepath)
            return send_file(pptx_filepath, as_attachment=True)

        elif content:
            pptx_data = create_ppt_from_textarea(content)
            return send_file(pptx_data, as_attachment=True, download_name="generated_presentation.pptx")

        return redirect(url_for("index"))

    return render_template("index.html")

def create_ppt_from_text(text_file):
    prs = Presentation()
    prs.slide_width = Cm(25.4)
    prs.slide_height = Cm(14.29)

    with open(text_file, "r", encoding="utf-8") as file:
        paragraphs = get_paragraphs(file.readlines())

    for paragraph in paragraphs:
        create_slide(prs, paragraph)

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], "generated_presentation.pptx")
    prs.save(output_path)
    return output_path

def create_ppt_from_textarea(content):
    prs = Presentation()
    prs.slide_width = Cm(25.4)
    prs.slide_height = Cm(14.29)

    lines = content.splitlines()
    paragraphs = get_paragraphs(lines)

    for paragraph in paragraphs:
        create_slide(prs, paragraph)

    pptx_data = io.BytesIO()
    prs.save(pptx_data)
    pptx_data.seek(0)
    return pptx_data

def get_paragraphs(lines):
    paragraphs = []
    paragraph = []

    for line in lines:
        line = line.strip()
        if line:
            paragraph.append(line)
        else:
            if paragraph:
                paragraphs.append(paragraph)
                paragraph = []
    if paragraph:
        paragraphs.append(paragraph)

    return paragraphs

def create_slide(prs, content):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    img_path = "static/background.jpg"  # Replace with your image file path

    background = slide.shapes.add_picture(img_path, Cm(0), Cm(0), width=prs.slide_width, height=prs.slide_height)
    slide.shapes._spTree.remove(background._element)
    slide.shapes._spTree.insert(2, background._element)

    title_box = slide.shapes.add_textbox(Cm(0), Cm((14.29 - 3) / 2), prs.slide_width, Cm(3))
    title_frame = title_box.text_frame
    title_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    for line in content:
        p = title_frame.add_paragraph() if title_frame.text else title_frame.paragraphs[0]
        p.text = line
        p.font.name = "Malgun Gothic"
        p.font.size = Pt(40)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        p.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

if __name__ == "__main__":
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)