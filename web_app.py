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

# 使用錯誤訊息指定的最新可用模型
model = genai.GenerativeModel('gemini-3.6-flash')

COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)

st.set_page_config(page_title="專業工地檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業工程檢核圖生成器")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成專業檢核圖"):
        with st.spinner('正在與 AI 連線並繪製圖表...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                prompt = """
                分析圖片，找出 4-6 個關鍵檢核點。
                回傳嚴格的 JSON 陣列: [{"label": "A", "box_2d": [ymin, xmin, ymax, xmax], "title": "標題", "items": "項目一\n項目二"}]
                items 內的文字請簡短，不需額外符號。只輸出 JSON，不加其他文字。
                """
                
                response = model.generate_content([img_pil, prompt])
                
                text = response.text.replace('```json', '').replace('```', '').strip()
                boxes_data = json.loads(text)

                prs = Presentation()
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * (width_px / height_px))
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                for idx, item in enumerate(boxes_data):
                    is_left = (idx % 2 == 0)
                    base_color = COLOR_BLUE if is_left else COLOR_GREEN
                    card_x = Inches(0.2) if is_left else prs.slide_width - Inches(3.7)
                    card_y = Inches(1.0 + (idx // 2) * 1.5)

                    title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, card_y, Inches(3.5), Inches(0.5))
                    title_box.fill.solid()
                    title_box.fill.fore_color.rgb = base_color
                    tf = title_box.text_frame
                    tf.text = f"{item.get('label', '')}. {item['title']}"
                    tf.paragraphs[0].font.size = Pt(16)
                    tf.paragraphs[0].font.bold = True
                    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                    items_list = item["items"].split('\n')
                    items_height = max(len(items_list) * 0.35 + 0.2, 0.7)
                    items_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, card_y + Inches(0.5), Inches(3.5), Inches(items_height))
                    items_box.fill.solid()
                    items_box.fill.fore_color.rgb = RGBColor(240, 240, 240)
                    
                    tf_items = items_box.text_frame
                    tf_items.vertical_anchor = MSO_ANCHOR.TOP
                    for it in items_list:
                        p = tf_items.add_paragraph()
                        p.text = f"☑ {it.strip()}"
                        p.font.size = Pt(12)
                        p.font.color.rgb = RGBColor(50, 50, 50)

                    ymin, xmin, ymax, xmax = item["box_2d"]
                    target_x = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    target_y = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    
                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, int(card_x + Inches(3.5)) if is_left else int(card_x), int(card_y + Inches(0.4)), target_x, target_y)
                    line.line.color.rgb = RGBColor(255, 255, 255)
                    line.line.width = Pt(3) 

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 報告生成成功！")
                st.download_button("📥 下載專業檢核報告", pptx_io, "專業工安檢核圖.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
