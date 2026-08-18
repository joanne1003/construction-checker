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

st.set_page_config(page_title="專業職安檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業職安人員檢核圖生成器")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成專業職安檢核表"):
        with st.spinner('職安工程師正在進行風險評估...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                # 【強化職安 persona 的提示詞】
                prompt = """
                您是一位擁有豐富營造工程資歷的專業職業安全衛生管理員 (Safety and Health Engineer)。
                請針對這張施工照片進行深入的風險分析與職安稽核。
                
                目標：
                1. 判斷工程類型（例如：橋梁墩柱、模板支撐、鋼筋綁紮、瀝青鋪設等）。
                2. 找出 4-6 個關鍵的工安風險點或品質檢核點，須符合職業安全衛生法規的精神。
                3. 用語必須專業、精煉，具備工地現場稽核的實務指導意義。
                
                回傳嚴格的 JSON (不含其他文字): 
                {
                  "project_type": "工程名稱",
                  "checkpoints": [{"label": "A", "box_2d": [ymin, xmin, ymax, xmax], "title": "稽核重點項目", "items": "稽核要點一\n稽核要點二"}]
                }
                """
                
                response = model.generate_content([img_pil, prompt])
                text = response.text.replace('```json', '').replace('```', '').strip()
                data = json.loads(text)
                
                proj_type = data.get("project_type", "施工")
                checkpoints = data.get("checkpoints", [])

                # --- 排序與分邊邏輯 (保持線條整潔) ---
                left_items = []
                right_items = []
                for item in checkpoints:
                    if (item["box_2d"][1] + item["box_2d"][3]) / 2 < 500:
                        left_items.append(item)
                    else:
                        right_items.append(item)
                
                left_items.sort(key=lambda x: x["box_2d"][0])
                right_items.sort(key=lambda x: x["box_2d"][0])

                prs = Presentation()
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * (width_px / height_px))
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # 插入背景圖
                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 標題區
                title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.3), Inches(5.2), Inches(0.6))
                title_box.fill.solid()
                title_box.fill.fore_color.rgb = COLOR_BLUE
                tf = title_box.text_frame
                tf.text = f"{proj_type}施工簡易檢核圖"
                tf.paragraphs[0].font.size = Pt(18)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                
                sub_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.95), Inches(5.2), Inches(0.45))
                sub_box.fill.solid()
                sub_box.fill.fore_color.rgb = RGBColor(50, 50, 50)
                tf_sub = sub_box.text_frame
                tf_sub.text = f"{proj_type}職業安全衛生檢核要點"
                tf_sub.paragraphs[0].font.size = Pt(13)
                tf_sub.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                # 繪製卡片函數
                def draw_cards(items, is_left):
                    base_color = COLOR_BLUE if is_left else COLOR_GREEN
                    card_x = Inches(0.4) if is_left else prs.slide_width - Inches(3.9)
                    
                    for i, item in enumerate(items):
                        card_y = Inches(1.5 + i * 1.5)
                        card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, card_y, Inches(3.5), Inches(1.3))
                        card.fill.solid()
                        card.fill.fore_color.rgb = base_color
                        card.line.color.rgb = RGBColor(255, 255, 255)
                        card.line.width = Pt(1.5)
                        
                        tf = card.text_frame
                        tf.text = f"{item.get('label', '')}. {item['title']}"
                        tf.paragraphs[0].font.bold = True
                        for it in item["items"].split('\n'):
                            if it.strip():
                                p = tf.add_paragraph()
                                p.text = f"☑ {it.strip()}"
                                p.font.size = Pt(10.5)
                        
                        # 引導線
                        ymin, xmin, ymax, xmax = item["box_2d"]
                        tx = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                        ty = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                        lx = int(card_x + Inches(3.5)) if is_left else int(card_x)
                        
                        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, lx, int(card_y + Inches(0.65)), tx, ty)
                        line.line.color.rgb = RGBColor(255, 255, 255)
                        line.line.width = Pt(2)

                        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, tx-12, ty-12, 24, 24)
                        circle.fill.solid()
                        circle.fill.fore_color.rgb = base_color
                        circle.text_frame.text = item.get('label', '')

                draw_cards(left_items, True)
                draw_cards(right_items, False)

                # 現場檢核結果記錄
                table_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, prs.slide_width - Inches(3.9), prs.slide_height - Inches(1.8), Inches(3.5), Inches(1.5))
                table_box.fill.solid()
                table_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                table_box.line.color.rgb = RGBColor(100, 100, 100)
                tf_table = table_box.text_frame
                tf_table.text = "現場檢核結果記錄\n(項目 | 檢核結果 | 備註)"
                tf_table.paragraphs[0].font.bold = True

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 專業職安檢核報告已生成！")
                st.download_button("📥 下載專業檢核報告", pptx_io, "職安檢核報告.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
