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

# 使用指定的最新可用模型
model = genai.GenerativeModel('gemini-3.6-flash')

# 專業工程配色定義
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)
COLOR_DARK = RGBColor(40, 40, 40)

st.set_page_config(page_title="專業職安工程檢核圖生成器", page_icon="🏗️")
st.title("🏗️ 專業職安工程檢核圖生成器")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成專業檢核圖"):
        with st.spinner('專業職安工程師正在進行圖說繪製...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                prompt = """
                身為資深職業安全衛生工程師與營造工程專家，請分析此施工照片。
                辨識工程類型（例如：橋梁墩柱、鋼筋模板、基礎開挖等），並找出 4-6 個關鍵的工安風險與品質檢核點。
                回傳嚴格的 JSON 格式（絕對不可包含任何 markdown 以外的說明文字）：
                {
                  "project_type": "工程名稱",
                  "checkpoints": [
                    {
                      "label": "1",
                      "box_2d": [ymin, xmin, ymax, xmax],
                      "title": "檢核項目名稱",
                      "items": "檢核要點一\n檢核要點二"
                    }
                  ]
                }
                """
                
                response = model.generate_content([img_pil, prompt])
                text = response.text.replace('```json', '').replace('```', '').strip()
                data = json.loads(text)
                
                proj_type = data.get("project_type", "施工")
                checkpoints = data.get("checkpoints", [])

                # 依據物件實際位置動態分邊與排序，確保線條整潔不交錯
                left_items = []
                right_items = []
                for item in checkpoints:
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    if (xmin + xmax) / 2 < 500:
                        left_items.append(item)
                    else:
                        right_items.append(item)
                
                left_items.sort(key=lambda x: x["box_2d"][0])
                right_items.sort(key=lambda x: x["box_2d"][0])

                prs = Presentation()
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * (width_px / height_px))
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 左上角主標題
                title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.3), Inches(5.2), Inches(0.6))
                title_box.fill.solid()
                title_box.fill.fore_color.rgb = COLOR_BLUE
                title_box.line.color.rgb = RGBColor(255, 255, 255)
                tf = title_box.text_frame
                tf.text = f"{proj_type}施工簡易檢核圖"
                tf.paragraphs[0].font.size = Pt(20)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                tf.vertical_anchor = MSO_ANCHOR.MIDDLE

                # 左上角副標題
                sub_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.92), Inches(5.2), Inches(0.45))
                sub_box.fill.solid()
                sub_box.fill.fore_color.rgb = COLOR_DARK
                sub_box.line.color.rgb = RGBColor(255, 255, 255)
                tf_sub = sub_box.text_frame
                tf_sub.text = f"{proj_type}職業安全衛生與品質檢核要點"
                tf_sub.paragraphs[0].font.size = Pt(13)
                tf_sub.paragraphs[0].font.bold = True
                tf_sub.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                tf_sub.vertical_anchor = MSO_ANCHOR.MIDDLE

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
                        tf.word_wrap = True
                        tf.vertical_anchor = MSO_ANCHOR.TOP
                        
                        p_title = tf.paragraphs[0]
                        p_title.text = f"{item.get('label', '')}. {item['title']}"
                        p_title.font.bold = True
                        p_title.font.size = Pt(13)
                        p_title.font.color.rgb = RGBColor(255, 255, 255)
                        p_title.space_after = Pt(3)

                        for it in item["items"].split('\n'):
                            if it.strip():
                                p = tf.add_paragraph()
                                p.text = f"☑ {it.strip()}"
                                p.font.size = Pt(10.5)
                                p.font.color.rgb = RGBColor(255, 255, 255)
                                p.space_after = Pt(1)

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
                        circle.line.color.rgb = RGBColor(255, 255, 255)
                        circle.line.width = Pt(1.5)
                        tf_c = circle.text_frame
                        tf_c.text = item.get('label', '')
                        tf_c.paragraphs[0].font.size = Pt(11)
                        tf_c.paragraphs[0].font.bold = True
                        tf_c.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                        tf_c.vertical_anchor = MSO_ANCHOR.MIDDLE

                draw_cards(left_items, True)
                draw_cards(right_items, False)

                # 右下角單一檢核結果記錄表格
                table_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, prs.slide_width - Inches(3.9), prs.slide_height - Inches(1.8), Inches(3.5), Inches(1.5))
                table_box.fill.solid()
                table_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                table_box.line.color.rgb = RGBColor(100, 100, 100)
                table_box.line.width = Pt(1.5)
                
                tf_table = table_box.text_frame
                tf_table.word_wrap = True
                tf_table.vertical_anchor = MSO_ANCHOR.TOP
                
                p_th = tf_table.paragraphs[0]
                p_th.text = "現場檢核結果記錄"
                p_th.font.bold = True
                p_th.font.size = Pt(13)
                p_th.font.color.rgb = RGBColor(0, 0, 0)
                p_th.space_after = Pt(2)

                p_th2 = tf_table.add_paragraph()
                p_th2.text = "(項目 | 檢核結果 | 備註)"
                p_th2.font.size = Pt(10)
                p_th2.font.color.rgb = RGBColor(100, 100, 100)

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 專業職安檢核圖生成成功！")
                st.download_button("📥 下載專業檢核報告", pptx_io, "職安專業檢核圖.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
