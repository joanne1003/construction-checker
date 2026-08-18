import streamlit as st
import json
import time
from io import BytesIO
from google import genai
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR, MSO_TEXT_FRAME_ANCHOR
from pptx.dml.color import RGBColor

# ==========================================
# 🔑 設定
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# 定義顏色配置 (工程藍與工程綠)
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)

st.set_page_config(page_title="工地檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業工程檢核圖生成器")
uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, use_container_width=True)
    if st.button("🚀 生成工程級檢核圖"):
        with st.spinner('正在繪製專業圖表...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                # --- AI 提示詞優化 ---
                prompt = """
                擔任資深營造工程師。分析圖片並找出 4-6 個關鍵檢核點。
                請依照以下 JSON 格式回傳，務必包含一個標籤(label) 'A', 'B', 'C' 等順序編號。
                [
                  {
                    "label": "A",
                    "box_2d": [ymin, xmin, ymax, xmax],
                    "title": "簡短標題",
                    "items": "項目一\n項目二"
                  }
                ]
                """
                
                # ... (API 呼叫與重試機制同前) ...
                response = client.models.generate_content(
                    model="models/gemini-flash-latest", 
                    contents=[img_pil, prompt]
                )
                boxes_data = json.loads(response.text.replace('```json', '').replace('```', '').strip())

                # --- 建立 PPTX ---
                prs = Presentation()
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * (width_px / height_px))
                slide = prs.slides.add_slide(prs.slide_layouts[6])

                # 插入背景
                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 繪製標題列
                def draw_header(text, x, y, width, color):
                    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, width, Inches(0.8))
                    box.fill.solid()
                    box.fill.fore_color.rgb = color
                    tf = box.text_frame
                    tf.text = text
                    tf.paragraphs[0].font.size = Pt(24)
                    tf.paragraphs[0].font.bold = True
                    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                    tf.vertical_anchor = MSO_TEXT_FRAME_ANCHOR.MIDDLE

                draw_header("施工簡易檢核圖", Inches(0.2), Inches(0.2), Inches(6), COLOR_BLUE)

                # 繪製檢核框
                for idx, item in enumerate(boxes_data):
                    is_left = (idx % 2 == 0) # 左右交替
                    color = COLOR_BLUE if is_left else COLOR_GREEN
                    
                    x_pos = Inches(0.2) if is_left else prs.slide_width - Inches(3.7)
                    y_pos = Inches(1.2 + idx * 1.5)
                    
                    # 繪製方框
                    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_pos, y_pos, Inches(3.5), Inches(1.3))
                    box.fill.solid()
                    box.fill.fore_color.rgb = color
                    box.line.color.rgb = RGBColor(255, 255, 255)
                    
                    tf = box.text_frame
                    tf.text = f"{item['label']}. {item['title']}\n{item['items']}"
                    tf.paragraphs[0].font.size = Pt(16)
                    tf.paragraphs[0].font.bold = True
                    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                    tf.paragraphs[1].font.size = Pt(12)
                    tf.paragraphs[1].font.color.rgb = RGBColor(255, 255, 255)

                    # 連接線 (Connector)
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    target_x = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    target_y = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    start_x = int(x_pos + Inches(3.5)) if is_left else int(x_pos)
                    start_y = int(y_pos + Inches(0.65))
                    
                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, start_x, start_y, target_x, target_y)
                    line.line.color.rgb = RGBColor(255, 255, 255)
                    line.line.width = Pt(3)

                # 下載按鈕 (同前)
                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.download_button("📥 下載專業檢核報告", pptx_io, "檢核圖.pptx")
            except Exception as e:
                st.error(f"錯誤: {e}")
                )
                
            except Exception as e:
                st.error(f"發生錯誤：{e} (若持續發生 503 錯誤，代表目前伺服器流量較大，請稍候再按一次按鈕)")
