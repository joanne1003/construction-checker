import streamlit as st
import json
import time
from io import BytesIO
from google import genai
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor

# ==========================================
# 🔑 設定 API Key
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# ==========================================
# 🌐 網頁介面設計
# ==========================================
st.set_page_config(page_title="工安檢核圖生成器", page_icon="👷", layout="centered")
st.title("🚧 施工簡易檢核圖 AI 生成器")
st.markdown("上傳現場施工照片，AI 將自動辨識風險點並生成**高畫質、可編輯**的 PPTX 簡報。")

# 建立檔案上傳區塊
uploaded_file = st.file_uploader("📂 請上傳施工照片 (支援 JPG, PNG)", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, caption="已上傳的現場照片", use_container_width=True)
    
    if st.button("🚀 開始生成檢核報告 (含自動重試機制)"):
        with st.spinner('AI 正在深度分析照片並排版中（若遇塞車將自動重試）...'):
            try:
                # 讀取圖片
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                # --- AI 分析 (加入自動重試機制，最多重試 3 次) ---
                prompt = """
                一位專業營造工程師的視角，找出照片中 3 個「需要進行品質或安全檢核的施工部位」。
                請回傳嚴格的 JSON 陣列，標示座標(0-1000)與檢核細項，格式如下：
                [
                  {
                    "box_2d": [ymin, xmin, ymax, xmax],
                    "title": "1. 模板支撐與強度",
                    "items": "☑ 模板是否漏漿、變形\n☑ 支撐架牢固性檢查"
                  }
                ]
                注意：items 請合併成一個字串，用 \\n 換行。title 限 12 字內。只輸出 JSON，不加其他文字。
                """

                response = None
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model="models/gemini-flash-latest", 
                            contents=[img_pil, prompt]
                        )
                        break # 成功就跳出迴圈
                    except Exception as api_err:
                        if "503" in str(api_err) and attempt < max_retries - 1:
                            time.sleep(2) # 遇到 503 等待 2 秒後重試
                            continue
                        else:
                            raise api_err # 其他錯誤或超過次數則拋出

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

                # 標題列
                title_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(Inches(0.5)), int(Inches(0.5)), int(Inches(5.5)), int(Inches(1)))
                title_box.fill.solid()
                title_box.fill.fore_color.rgb = RGBColor(0, 80, 160) 
                title_box.line.color.rgb = RGBColor(255, 255, 255)
                tf = title_box.text_frame
                tf.text = "施工簡易檢核圖"
                tf.paragraphs[0].font.size = Pt(28)
                tf.paragraphs[0].font.bold = True
                tf.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)

                # 簽核欄位
                sign_box_width = int(Inches(3.5))
                sign_box_x = int(prs.slide_width - Inches(0.5) - sign_box_width)
                sign_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sign_box_x, int(Inches(0.5)), sign_box_width, int(Inches(2)))
                sign_box.fill.solid()
                sign_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
                sign_box.line.color.rgb = RGBColor(0, 80, 160)
                sign_box.line.width = Pt(2)
                tf_sign = sign_box.text_frame
                tf_sign.text = "現場檢核結果記錄\n(項目 | 檢核結果 | 備註)"
                tf_sign.paragraphs[0].font.size = Pt(16)
                tf_sign.paragraphs[0].font.bold = True
                tf_sign.paragraphs[0].font.color.rgb = RGBColor(0, 80, 160)
                tf_sign.paragraphs[1].font.size = Pt(12)
                tf_sign.paragraphs[1].font.color.rgb = RGBColor(100, 100, 100)

                card_width = int(Inches(3.5))
                card_height = int(Inches(1.8))
                
                for idx, item in enumerate(boxes_data):
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    target_x = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    target_y = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    
                    is_left = (idx % 2 == 0)
                    card_x = int(Inches(0.5)) if is_left else int(prs.slide_width - Inches(0.5) - card_width)
                    card_y = int(Inches(2.5 + idx * 1.9)) 
                    
                    pixel_x = int((xmin + xmax) / 2 / 1000 * width_px)
                    pixel_y = int((ymin + ymax) / 2 / 1000 * height_px)
                    try:
                        r, g, b = img_pil.getpixel((pixel_x, pixel_y))
                        luminance = 0.299*r + 0.587*g + 0.114*b
                        line_color = RGBColor(255, 255, 255) if luminance < 130 else RGBColor(0, 80, 160)
                    except:
                        line_color = RGBColor(255, 255, 0) 

                    line_start_x = int(card_x + card_width) if is_left else int(card_x)
                    line_start_y = int(card_y + card_height / 2)
                    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, line_start_x, line_start_y, target_x, target_y)
                    connector.line.color.rgb = line_color
                    connector.line.width = Pt(4.5) 

                    circle_size = int(Inches(0.2))
                    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, int(target_x - circle_size/2), int(target_y - circle_size/2), circle_size, circle_size)
                    circle.fill.solid()
                    circle.fill.fore_color.rgb = line_color
                    circle.line.color.rgb = RGBColor(255, 255, 255)
                    circle.line.width = Pt(1.5)

                    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, card_x, card_y, card_width, card_height)
                    card.fill.solid()
                    card.fill.fore_color.rgb = RGBColor(255, 255, 255) 
                    card.line.color.rgb = RGBColor(0, 80, 160)
                    card.line.width = Pt(2)
                    
                    tf_card = card.text_frame
                    tf_card.word_wrap = True
                    
                    p_title = tf_card.paragraphs[0]
                    p_title.text = item["title"]
                    p_title.font.bold = True
                    p_title.font.size = Pt(14)
                    p_title.font.color.rgb = RGBColor(0, 80, 160)
                    
                    p_items = tf_card.add_paragraph()
                    p_items.text = item["items"]
                    p_items.font.size = Pt(12)
                    p_items.font.color.rgb = RGBColor(50, 50, 50)

                # 提供下載
                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                
                st.success("✅ 報告生成成功！請點擊下方按鈕下載。")
                st.download_button(
                    label="📥 下載 PPTX 檢核報告",
                    data=pptx_io,
                    file_name="工安簡易檢核圖.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"發生錯誤：{e} (若持續發生 503 錯誤，代表目前伺服器流量較大，請稍候再按一次按鈕)")