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

# API 初始化
genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel('gemini-1.5-flash')

# 專業工程配色 (仿參考圖)
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)
COLOR_DARK = RGBColor(40, 40, 40)

st.set_page_config(page_title="專業職安工程檢核圖", page_icon="🏗️")
st.title("🏗️ 專業職安檢核圖生成器")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成與參考圖風格一致的專業檢核圖"):
        with st.spinner('專業職安工程師正在規劃圖說...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size
                
                # 專業職安人員視角的 Prompt
                prompt = """
                身為資深職業安全衛生工程師，請分析此施工照片。
                辨識工程類型，並找出 4-6 個關鍵的工安風險與品質檢核點。
                請確保檢核項目符合營造安全衛生設施標準。
                回傳 JSON (不可包含markdown代碼區塊): 
                {
                  "project_type": "工程名稱",
                  "checkpoints": [{"label": "A", "box_2d": [ymin, xmin, ymax, xmax], "title": "檢核項目", "items": "要點一\n要點二"}]
                }
                """
                
                response = model.generate_content([img_pil, prompt])
                text = response.text.replace('```json', '').replace('```', '').strip()
                data = json.loads(text)
                
                proj_type = data.get("project_type", "施工")
                checkpoints = data.get("checkpoints", [])

                # 排版邏輯
                left_items = []
                right_items = []
                for item in checkpoints:
                    if (item["box_2d"][1] + item["box_2d"][3]) / 2 < 500:
                        left_items.append(item)
                    else:
                        right_items.append(item)
                
                prs = Presentation()
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * (width_px / height_px))
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # 背景
                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 標題區
                t_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.3), Inches(5.2), Inches(0.6))
                t_box.fill.solid()
                t_box.fill.fore_color.rgb = COLOR_BLUE
                t_box.text_frame.text = f"{proj_type}施工簡易檢核圖"
                t_box.text_frame.paragraphs[0].font.bold = True
                
                s_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.95), Inches(5.2), Inches(0.45))
                s_box.fill.solid()
                s_box.fill.fore_color.rgb = COLOR_DARK
                s_box.text_frame.text = f"{proj_type}職業安全衛生檢核要點"
                s_box.text_frame.paragraphs[0].font.size = Pt(13)

                # 繪製圓角卡片
                def draw_card(item, x, y, is_left):
                    base_color = COLOR_BLUE if is_left else COLOR_GREEN
                    # 卡片本體 (圓角)
                    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.5), Inches(1.5))
                    card.fill.solid()
                    card.fill.fore_color.rgb = base_color
                    card.line.color.rgb = RGBColor(255, 255, 255)
                    
                    # 標題區 (用文字段落區隔)
                    tf = card.text_frame
                    p1 = tf.paragraphs[0]
                    p1.text = f"{item.get('label', '')}. {item['title']}"
                    p1.font.bold = True
                    p1.font.size = Pt(13)
                    p1.font.color.rgb = RGBColor(255, 255, 255)
                    
                    # 加入一條線區隔標題與內容
                    tf.add_paragraph().text = "─────────────────"
                    
                    # 內容區
                    for it in item["items"].split('\n'):
                        p = tf.add_paragraph()
                        p.text = f"☑ {it.strip()}"
                        p.font.size = Pt(10)
                        p.font.color.rgb = RGBColor(255, 255, 255)
                    
                    # 連線
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    tx = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    ty = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    lx = int(x + Inches(3.5)) if is_left else int(x)
                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, lx, int(y + Inches(0.75)), tx, ty)
                    line.line.color.rgb = RGBColor(255, 255, 255)
                    line.line.width = Pt(2)

                for i, item in enumerate(left_items): draw_card(item, Inches(0.4), Inches(1.6 + i * 1.7), True)
                for i, item in enumerate(right_items): draw_card(item, prs.slide_width - Inches(3.9), Inches(1.6 + i * 1.7), False)

                # 表格
                t_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, prs.slide_width - Inches(3.9), prs.slide_height - Inches(1.8), Inches(3.5), Inches(1.5))
                t_box.fill.solid()
                t_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                t_box.text_frame.text = "現場檢核結果記錄\n(項目 | 檢核結果 | 備註)"

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 專業工安檢核報告已生成！")
                st.download_button("📥 下載專業工安檢核圖.pptx", pptx_io, "檢核圖.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
