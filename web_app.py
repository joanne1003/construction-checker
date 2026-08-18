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

# 定義專業工程配色
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)

st.set_page_config(page_title="專業工地檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業工程檢核圖生成器")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成專業檢核圖"):
        with st.spinner('正在分析與繪製專業檢核圖表...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                prompt = """
                分析施工照片，判斷工程類型（例如：橋梁墩柱、鋼筋綁紮、模板支撐、基礎開挖等），並找出 4-6 個關鍵檢核點。
                必須回傳嚴格的 JSON 格式（不要包覆在 markdown 以外的其他說明文字中），結構如下：
                {
                  "project_type": "工程類型名稱（例如：橋梁墩柱）",
                  "checkpoints": [
                    {
                      "label": "A",
                      "box_2d": [ymin, xmin, ymax, xmax],
                      "title": "檢核標題",
                      "items": "項目一\n項目二"
                    }
                  ]
                }
                """
                
                response = model.generate_content([img_pil, prompt])
                text = response.text.replace('```json', '').replace('```', '').strip()
                data = json.loads(text)
                
                proj_type = data.get("project_type", "施工")
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

                # --- 頂部標題區 (左上角) ---
                # 主標題: XXX施工簡易檢核圖
                title_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.3), Inches(5.0), Inches(0.6))
                title_box.fill.solid()
                title_box.fill.fore_color.rgb = COLOR_BLUE
                title_box.line.color.rgb = RGBColor(255, 255, 255)
                title_box.line.width = Pt(1.5)
                tf = title_box.text_frame
                tf.text = f"{proj_type}施工簡易檢核圖"
                tf.paragraphs[0].font.size = Pt(20)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE

                # 副標題: XXX施工檢核要點
                sub_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.95), Inches(5.0), Inches(0.45))
                sub_box.fill.solid()
                sub_box.fill.fore_color.rgb = RGBColor(40, 40, 40)
                sub_box.line.color.rgb = RGBColor(255, 255, 255)
                sub_box.line.width = Pt(1.5)
                tf_sub = sub_box.text_frame
                tf_sub.text = f"{proj_type}施工檢核要點"
                tf_sub.paragraphs[0].font.size = Pt(14)
                tf_sub.paragraphs[0].font.bold = True
                tf_sub.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                tf_sub.vertical_anchor = MSO_ANCHOR.MIDDLE

                # --- 繪製檢核卡片 (圓角矩形、整潔對齊) ---
                card_width = Inches(3.4)
                card_height = Inches(1.5)
                
                for idx, item in enumerate(boxes_data):
                    is_left = (idx % 2 == 0)
                    base_color = COLOR_BLUE if is_left else COLOR_GREEN
                    
                    card_x = Inches(0.4) if is_left else prs.slide_width - Inches(3.8)
                    card_y = Inches(1.5 + (idx // 2) * 1.65)

                    # 圓角矩形卡片
                    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, card_y, card_width, card_height)
                    card.fill.solid()
                    card.fill.fore_color.rgb = base_color
                    card.line.color.rgb = RGBColor(255, 255, 255)
                    card.line.width = Pt(2)

                    tf_card = card.text_frame
                    tf_card.word_wrap = True
                    tf_card.vertical_anchor = MSO_ANCHOR.TOP

                    # 卡片標題
                    p_t = tf_card.paragraphs[0]
                    p_t.text = f"{item.get('label', '')}. {item['title']}"
                    p_t.font.size = Pt(15)
                    p_t.font.bold = True
                    p_t.font.color.rgb = RGBColor(255, 255, 255)
                    p_t.space_after = Pt(4)

                    # 檢核項目
                    items_list = item["items"].split('\n')
                    for it in items_list:
                        if it.strip():
                            p_i = tf_card.add_paragraph()
                            p_i.text = f"☑ {it.strip()}"
                            p_i.font.size = Pt(11.5)
                            p_i.font.color.rgb = RGBColor(255, 255, 255)
                            p_i.space_after = Pt(2)

                    # 引導線與座標標籤
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    target_x = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    target_y = int((ymin + ymax) / 2 / 1000 * prs.slide_height)

                    line_start_x = int(card_x + card_width) if is_left else int(card_x)
                    line_start_y = int(card_y + card_height / 2)

                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, line_start_x, line_start_y, target_x, target_y)
                    line.line.color.rgb = RGBColor(255, 255, 255)
                    line.line.width = Pt(2.5)

                    # 圓形標籤 (A, B, C...)
                    circle_size = Inches(0.28)
                    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, target_x - circle_size/2, target_y - circle_size/2, circle_size, circle_size)
                    circle.fill.solid()
                    circle.fill.fore_color.rgb = base_color
                    circle.line.color.rgb = RGBColor(255, 255, 255)
                    circle.line.width = Pt(1.5)
                    tf_c = circle.text_frame
                    tf_c.text = item.get('label', '')
                    tf_c.paragraphs[0].font.size = Pt(11)
                    tf_c.paragraphs[0].font.bold = True
                    tf_c.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                    tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 專業工程檢核圖生成成功！")
                st.download_button("📥 下載專業檢核報告", pptx_io, "專業工安檢核圖.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
