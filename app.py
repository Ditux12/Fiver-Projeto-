from flask import Flask, request, send_file
import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import io

app = Flask(__name__)

# Funções do seu código adaptadas para receber dataframe e gerar pptx na memória
def criar_pptx(dataframes, logo_path=None):
    BACKGROUND_COLOR = RGBColor(230, 230, 230)
    TITLE_FONT = 'Barlow'
    BODY_FONT = 'Calibri'
    TITLE_FONT_SIZE = Pt(28)
    BODY_FONT_SIZE = Pt(18)
    TEXT_COLOR = RGBColor(50, 50, 50)

    prs = Presentation()

    def set_slide_background_gray(slide):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = BACKGROUND_COLOR

    def add_logo(slide):
        if logo_path:
            left = prs.slide_width - Inches(1.5)
            top = Inches(0.2)
            height = Inches(1)
            slide.shapes.add_picture(logo_path, left, top, height=height)

    def add_title_slide(prs, title, subtitle):
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        set_slide_background_gray(slide)
        title_box = slide.shapes.add_textbox(Inches(1), Inches(1.5), prs.slide_width - Inches(2), Inches(1))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title
        p.font.name = TITLE_FONT
        p.font.size = TITLE_FONT_SIZE
        p.font.bold = True
        p.font.color.rgb = TEXT_COLOR

        subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(3), prs.slide_width - Inches(2), Inches(0.7))
        tf2 = subtitle_box.text_frame
        p2 = tf2.paragraphs[0]
        run2 = p2.add_run()
        run2.text = subtitle
        p2.font.name = BODY_FONT
        p2.font.size = Pt(16)
        p2.font.italic = True
        p2.font.color.rgb = TEXT_COLOR

        add_logo(slide)

    def add_category_summary_slide(prs, category, df):
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        set_slide_background_gray(slide)
        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.5), prs.slide_width - Inches(1.4), Inches(1))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = category.upper()
        p.font.name = TITLE_FONT
        p.font.size = TITLE_FONT_SIZE
        p.font.bold = True
        p.font.color.rgb = TEXT_COLOR

        total_news = len(df)
        total_circulation = df['Circulação'].sum() if 'Circulação' in df.columns else 0

        content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.7), prs.slide_width - Inches(1.4), Inches(1))
        tf2 = content_box.text_frame
        p2 = tf2.paragraphs[0]
        run2 = p2.add_run()
        run2.text = f"Total de Notícias: {total_news}\nTotal de Circulação: {total_circulation:,}".replace(",", ".")
        p2.font.name = BODY_FONT
        p2.font.size = BODY_FONT_SIZE
        p2.font.color.rgb = TEXT_COLOR

        add_logo(slide)

    def add_news_slide(prs, title_text, circulation):
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        set_slide_background_gray(slide)

        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.3), prs.slide_width - Inches(1.4), Inches(0.8))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "NOTÍCIA"
        p.font.name = TITLE_FONT
        p.font.size = TITLE_FONT_SIZE
        p.font.bold = True
        p.font.color.rgb = TEXT_COLOR

        content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.3), prs.slide_width - Inches(1.4), Inches(3))
        tf2 = content_box.text_frame
        p2 = tf2.paragraphs[0]
        run2 = p2.add_run()
        run2.text = title_text
        p2.font.name = BODY_FONT
        p2.font.size = BODY_FONT_SIZE
        p2.font.color.rgb = TEXT_COLOR

        circ_box = slide.shapes.add_textbox(Inches(0.7), Inches(4.6), prs.slide_width - Inches(1.4), Inches(0.7))
        tf3 = circ_box.text_frame
        p3 = tf3.paragraphs[0]
        run3 = p3.add_run()
        run3.text = f"Circulação: {circulation:,}".replace(",", ".")
        p3.font.name = BODY_FONT
        p3.font.size = BODY_FONT_SIZE
        p3.font.italic = True
        p3.font.color.rgb = TEXT_COLOR

        add_logo(slide)

    # Começa a montar o pptx
    add_title_slide(prs, "Relatório de Notícias", "Gerado automaticamente a partir do Excel")

    for category, df in dataframes.items():
        if df.empty or 'Título' not in df.columns:
            continue
        df = df.dropna(subset=['Título'])
        if df.empty:
            continue

        add_category_summary_slide(prs, category, df)

        for _, row in df.iterrows():
            title = str(row['Título'])
            circulation = row.get('Circulação', 0)
            add_news_slide(prs, title, circulation)

    return prs

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return {"error": "Arquivo Excel não enviado"}, 400
    file = request.files['file']

    # Lê o Excel em memória
    xls = pd.ExcelFile(file)
    sheets = [s for s in xls.sheet_names if not s.startswith("Sheet")]
    dataframes = {sheet: xls.parse(sheet) for sheet in sheets}

    prs = criar_pptx(dataframes)

    # Salvar em buffer para enviar sem gravar no disco
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)

    return send_file(
        pptx_io,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        download_name='Clipping_Automatico_Design.pptx',
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
