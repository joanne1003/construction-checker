import streamlit as st
import json
from io import BytesIO
import google.generativeai as genai
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR
from pptx.dml.color import RGBColor

# 設定 API
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel('gemini-3.6-flash')

# 配色定義
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)

st.set_page_config(page_title="專業工地檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業工程檢核圖生成器")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成參考圖風格檢核表"):
        with st.spinner('正在繪製專業圖表...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                prompt = """
                分析施工照片，判斷工程類型（例如：橋梁墩柱、瀝青鋪設等）。
                找出 4-6 個關鍵檢核點。
                回傳嚴格的 JSON (不含其他文字): 
                {
                  "project_type": "工程名稱",
                  "checkpoints": [{"label": "A", "box_2d": [ymin, xmin, ymax, xmax], "title": "標題", "items": "項目一\n項目二"}]
                }
                """
                
                response = model.generate_content([img_pil, prompt])
                text = response.text.replace('```json', '').replace('```', '').strip()
                data = json.loads(text)
                
                proj_type = data.get("project_type", "工程")
                boxes_data = data.get("checkpoints", [])

                prs = Presentation()
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * (width_px / height_px))
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # 插入背景圖
                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 1. 標題區 (左上角)
                title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.3), Inches(5.0), Inches(0.6))
                title_box.fill.solid()
                title_box.fill.fore_color.rgb = COLOR_BLUE
                tf = title_box.text_frame
                tf.text = f"{proj_type}施工簡易檢核圖"
                tf.paragraphs[0].font.size = Pt(20)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                
                sub_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.9), Inches(5.0), Inches(0.4))
                sub_box.fill.solid()
                sub_box.fill.fore_color.rgb = RGBColor(50, 50, 50)
                tf_sub = sub_box.text_frame
                tf_sub.text = f"{proj_type}施工檢核要點"
                tf_sub.paragraphs[0].font.size = Pt(14)
                tf_sub.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                # 2. 檢核卡片 (網格佈局)
                for idx, item in enumerate(boxes_data):
                    is_left = (idx % 2 == 0)
                    base_color = COLOR_BLUE if is_left else COLOR_GREEN
                    card_x = Inches(0.4) if is_left else prs.slide_width - Inches(3.9)
                    card_y = Inches(1.5 + (idx // 2) * 1.5)

                    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, card_y, Inches(3.5), Inches(1.3))
                    card.fill.solid()
                    card.fill.fore_color.rgb = base_color
                    card.line.width = Pt(1.5)

                    tf = card.text_frame
                    tf.text = f"{item.get('label', '')}. {item['title']}"
                    tf.paragraphs[0].font.bold = True
                    tf.paragraphs[0].font.size = Pt(14)
                    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                    for it in item["items"].split('\n'):
                        p = tf.add_paragraph()
                        p.text = f"☑ {it.strip()}"
                        p.font.size = Pt(11)
                        p.font.color.rgb = RGBColor(255, 255, 255)

                    # 引導線
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    tx = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    ty = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    lx = int(card_x + Inches(3.5)) if is_left else int(card_x)
                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, lx, int(card_y + Inches(0.6)), tx, ty)
                    line.line.color.rgb = RGBColor(255, 255, 255)
                    line.line.width = Pt(2)

                    # 標籤圈
                    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, tx-10, ty-10, 20, 20)
                    circle.fill.solid()
                    circle.fill.fore_color.rgb = base_color
                    circle.text_frame.text = item.get('label', '')

                # 3. 現場檢核結果記錄 (右下角固定表格)
                table_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, prs.slide_width - Inches(3.9), prs.slide_height - Inches(1.8), Inches(3.5), Inches(1.5))
                table_box.fill.solid()
                table_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                table_box.line.width = Pt(1.5)
                tf_table = table_box.text_frame
                tf_table.text = "現場檢核結果記錄\n(項目 | 檢核結果 | 備註)\n\n\n\n"
                tf_table.paragraphs[0].font.bold = True

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 專業工程檢核圖生成成功！")
                st.download_button("📥 下載專業檢核報告", pptx_io, "專業工安檢核圖.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
