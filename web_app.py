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
# 🔑 設定 API Key
# ==========================================
API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

# 定義專業工程配色
COLOR_BLUE = RGBColor(0, 80, 160)
COLOR_GREEN = RGBColor(0, 128, 64)

# ==========================================
# 🌐 網頁介面設計
# ==========================================
st.set_page_config(page_title="專業工地檢核生成器", page_icon="🏗️")
st.title("🏗️ 專業工程檢核圖生成器")
st.markdown("上傳照片，AI 將自動產出符合工程規範的專業檢核圖。")

uploaded_file = st.file_uploader("📂 上傳施工照片", type=["jpg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, caption="已上傳現場照片", use_container_width=True)
    if st.button("🚀 生成檢核圖表"):
        with st.spinner('正在繪製專業圖表...'):
            try:
                img_pil = Image.open(uploaded_file).convert("RGB")
                width_px, height_px = img_pil.size

                prompt = """
                擔任資深營造工程師。分析圖片找出 4-6 個關鍵檢核點。
                回傳嚴格的 JSON 陣列，務必包含一個標籤(label) 'A', 'B', 'C' 等順序編號。
                格式範例: [{"label": "A", "box_2d": [ymin, xmin, ymax, xmax], "title": "支座安裝", "items": "支座正位、平整\n固定錨栓鎖緊"}]
                注意：items 為字串，用 \\n 隔開，不需額外加勾選符號。
                """
                
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

                # 插入背景圖
                img_io = BytesIO()
                img_pil.save(img_io, format='JPEG')
                img_io.seek(0)
                slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 繪製檢核框
                for idx, item in enumerate(boxes_data):
                    is_left = (idx % 2 == 0)
                    color = COLOR_BLUE if is_left else COLOR_GREEN
                    
                    card_x = Inches(0.2) if is_left else prs.slide_width - Inches(3.7)
                    card_y = Inches(1.0 + (idx // 2) * 1.6)
                    
                    # 繪製方框
                    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, card_x, card_y, Inches(3.5), Inches(1.4))
                    box.fill.solid()
                    box.fill.fore_color.rgb = color
                    box.line.color.rgb = RGBColor(255, 255, 255)
                    box.line.width = Pt(2)
                    
                    # 設定文字區塊
                    tf = box.text_frame
                    tf.clear()
                    tf.word_wrap = True
                    tf.vertical_anchor = MSO_ANCHOR.TOP
                    
                    # 標題 (段落 1)
                    p_title = tf.add_paragraph()
                    p_title.text = f"{item.get('label', '')}. {item['title']}"
                    p_title.font.bold = True
                    p_title.font.size = Pt(18)
                    p_title.font.color.rgb = RGBColor(255, 255, 255)
                    p_title.space_after = Pt(6) # 標題與清單的間距
                    
                    # 清單項目 (段落 2+)
                    items_list = item["items"].split('\n')
                    for it in items_list:
                        p = tf.add_paragraph()
                        p.text = f"☑ {it.strip()}"
                        p.font.size = Pt(14)
                        p.font.bold = True # 清單字體也加粗，符合參考圖風格
                        p.font.color.rgb = RGBColor(255, 255, 255)
                        p.space_after = Pt(2)

                    # 連接線與小圓圈標籤
                    ymin, xmin, ymax, xmax = item["box_2d"]
                    target_x = int((xmin + xmax) / 2 / 1000 * prs.slide_width)
                    target_y = int((ymin + ymax) / 2 / 1000 * prs.slide_height)
                    
                    line_start_x = int(card_x + Inches(3.5)) if is_left else int(card_x)
                    line_start_y = int(card_y + Inches(0.7))
                    
                    # 連接線
                    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, line_start_x, line_start_y, target_x, target_y)
                    line.line.color.rgb = RGBColor(255, 255, 255)
                    line.line.width = Pt(3) 

                    # 標籤圈 (A, B, C...)
                    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, target_x - Inches(0.15), target_y - Inches(0.15), Inches(0.3), Inches(0.3))
                    circle.fill.solid()
                    circle.fill.fore_color.rgb = color
                    circle.line.color.rgb = RGBColor(255, 255, 255)
                    tf_circle = circle.text_frame
                    tf_circle.text = item.get('label', '')
                    tf_circle.paragraphs[0].font.size = Pt(12)
                    tf_circle.paragraphs[0].font.bold = True
                    tf_circle.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
                    tf_circle.vertical_anchor = MSO_ANCHOR.MIDDLE

                # 下載
                pptx_io = BytesIO()
                prs.save(pptx_io)
                pptx_io.seek(0)
                st.success("✅ 報告生成成功！")
                st.download_button("📥 下載專業檢核報告", pptx_io, "專業工安檢核圖.pptx")
                
            except Exception as e:
                st.error(f"錯誤: {e}")
                
            except Exception as e:
                st.error(f"錯誤: {e}")
