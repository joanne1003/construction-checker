import streamlit as st
import json
import time
from io import BytesIO
from google import genai
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR
from pptx.dml.color import RGBColor

# ==========================================
# 🔑 設定 API Key (從 Streamlit Cloud Secrets 讀取)
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# 定義顏色配置 (工程藍與工程綠)
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)

# ==========================================
# 🌐 網頁介面設計
# ==========================================
st.set_page_config(page_title="工地檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業工程檢核圖生成器")
st.markdown("上傳現場施工照片，AI 將自動辨識風險點並生成**專業工程圖表風格**的 PPTX 簡報。")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, caption="已上傳的現場照片", use_container_width=True)
    if st.button("🚀 生成工程級檢核圖"):
        with st.spinner('正在繪製專業圖表 (含自動重試機制)...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                # --- AI 提示詞優化 (加入 A, B, C 標籤) ---
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
                注意：items 請合併成一個字串，用 \\n 換行。只輸出 JSON，不加其他文字。
                """
                
                # --- API 呼叫與自動重試 ---
                response = None
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model="models/gemini-flash-latest", 
                            contents=[img_pil, prompt]
                        )
                        break
                    except Exception as api_err:
                        if "503" in str(api_err) and attempt < max_retries - 1:
                            time.sleep(2)
                            continue
                        else:
                            raise api_err

                text = response.text.replace('```json', '').replace('```', '').strip()
                boxes_data = json.loads(text)

                # --- 建立 PPTX ---
                prs = Presentation()
                aspect_ratio = width_px / height_px
                prs.slide_height = Inches(7.5) 
                prs.slide_width = Inches(7.5 * aspect_ratio)
                blank_slide_layout = prs.slide_layouts[6] 
                slide = prs.slides.add_slide(blank_slide_layout)

                # 插入背景圖
                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 繪製標題列功能函式
                def draw_header(text, x, y, width, color):
                    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, width, Inches(0.8))
                    box.fill.solid()
                    box.fill.fore_color.rgb = color
                    box.line.color.rgb = RGBColor(255, 255, 255)
                    box.line.width = Pt(2)
                    tf = box.text_frame
                    tf.text = text
                    tf.paragraphs[0].font.size = Pt(24)
                    tf.paragraphs[0].font.bold = True
                    tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

                # 繪製主標題
                draw_header("施工簡易檢核圖", Inches(0.2), Inches(0.2), Inches(5.5), COLOR_BLUE)
                
                # 繪製右上方簽核記錄提示
                sign_box_width = Inches(3.5)
                sign_box_x = prs.slide_width - Inches(0.2) - sign_box_width
                sign_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, sign_box_x, Inches(0.2), sign_box_width, Inches(0.8))
                sign_box.fill.solid()
                sign_box.fill.fore_color.rgb = RGBColor(80, 80, 80)
                sign_box.line.color.rgb = RGBColor(255, 255, 255)
                sign_box.line.width = Pt(2)
                tf_sign = sign_box.text_frame
                tf_sign.text = "現場檢核結果記錄\n(項目 | 檢核結果 | 備註)"
                tf_sign.paragraphs[0].font.size = Pt(14)
                tf_sign.paragraphs[0].font.bold = True
                tf_sign.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                tf_sign.paragraphs[1].font.size = Pt(10)
                tf_sign.paragraphs[1].font.color.rgb = RGBColor(200, 200, 200)

                # 繪製檢核框與連接線
                card_width = Inches(3.5)
                card_height = Inches(1.3)
                
                for idx, item in enumerate(boxes_data):
                    is_left = (idx % 2 == 0)
                    color = COLOR_BLUE if is_left else COLOR_GREEN
                    
                    card_x = Inches(0.2) if is_left else prs.slide_width - Inches(0.2) - card_width
                    card_y = Inches(1.2 + (idx // 2) * 1.5)
                    
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    target_x = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    target_y = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    
                    line_start_x = int(card_x + card_width) if is_left else int(card_x)
                    line_start_y = int(card_y + card_height / 2)
                    
                    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, line_start_x, line_start_y, target_x, target_y)
                    connector.line.color.rgb = RGBColor(255, 255, 255)
                    connector.line.width = Pt(3) 

                    circle_size = Inches(0.3)
                    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, target_x - circle_size/2, target_y - circle_size/2, circle_size, circle_size)
                    circle.fill.solid()
                    circle.fill.fore_color.rgb = color
                    circle.line.color.rgb = RGBColor(255, 255, 255)
                    circle.line.width = Pt(1.5)
                    tf_circle = circle.text_frame
                    tf_circle.text = item.get('label', '')
                    tf_circle.paragraphs[0].font.size = Pt(12)
                    tf_circle.paragraphs[0].font.bold = True
                    tf_circle.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                    tf_circle.vertical_anchor = MSO_ANCHOR.MIDDLE

                    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, card_y, card_width, card_height)
                    box.fill.solid()
                    box.fill.fore_color.rgb = color
                    box.line.color.rgb = RGBColor(255, 255, 255)
                    box.line.width = Pt(2)
                    
                    tf = box.text_frame
                    tf.word_wrap = True
                    
                    title_text = f"{item.get('label', '')}. {item['title']}"
                    p_title = tf.paragraphs[0]
                    p_title.text = title_text
                    p_title.font.bold = True
                    p_title.font.size = Pt(16)
                    p_title.font.color.rgb = RGBColor(255, 255, 255)
                    
                    items_text = item["items"].replace('☑', '').replace('- ', '')
                    items_list = items_text.split('\n')
                    formatted_items = '\n'.join([f"☑ {it.strip()}" for it in items_list if it.strip()])
                    
                    p_items = tf.add_paragraph()
                    p_items.text = formatted_items
                    p_items.font.size = Pt(12)
                    p_items.font.color.rgb = RGBColor(255, 255, 255)

                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                
                st.success("✅ 報告生成成功！請點擊下方按鈕下載。")
                st.download_button(
                    label="📥 下載專業檢核報告",
                    data=pptx_io,
                    file_name="專業工安檢核圖.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"錯誤: {e}")
