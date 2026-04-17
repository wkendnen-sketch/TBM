import os
import io
import json
import time
import tempfile
from dataclasses import dataclass
from typing import List
import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

# --- 설정 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
BASE_FONT_SIZE_PT = 30
TITLE_TEXT = "설명/说明"

@dataclass
class SlideData:
    image_path: str
    ko: str
    zh: str
    vi: str
    my: str

def translate_batch_with_gemini(api_key: str, korean_list: List[str]):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={api_key}"
    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])
    
    prompt = f"Translate to JSON array (keys: zh, vi, my). No markdown. Input:\n{joined_text}"

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1}
    }

    resp = requests.post(url, headers={"Content-Type": "application/json"}, json=payload, timeout=90)
    if resp.status_code != 200:
        st.error(f"API Error: {resp.text}")
        resp.raise_for_status()
    
    raw_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
    # 마크다운 제거 로직 추가
    clean_json = raw_text.replace("```json", "").replace("```", "").strip()
    return json.loads(clean_json)

def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes): yield sub_shape

def find_picture_area(slide):
    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        if "사진대지" in (shape.text.strip() if hasattr(shape, "has_text_frame") and shape.has_text_frame else ""):
            return shape
        try: candidates.append((shape.top, -(shape.width * shape.height), shape))
        except: pass
    return sorted(candidates)[0][2]

def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    prs = Presentation(TEMPLATE_PPT)
    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides): break
        slide = prs.slides[i]
        pic_area = find_picture_area(slide)
        
        # 이미지 삽입
        slide.shapes.add_picture(item.image_path, pic_area.left, pic_area.top, width=pic_area.width, height=pic_area.height)
        
        # 텍스트 삽입 (단순화된 로직)
        txt_shapes = [s for s in iter_all_shapes(slide.shapes) if hasattr(s, "has_text_frame") and s.has_text_frame and s.top > pic_area.top + pic_area.height - 100]
        txt_shapes.sort(key=lambda s: (s.top, s.left))
        
        for shape, txt in zip(txt_shapes, [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]):
            tf = shape.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = txt
            run.font.size = Pt(BASE_FONT_SIZE_PT)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def main():
    st.title("🏗️ TBM 자동 생성기 (에러 수정판)")
    
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("Secrets에 API 키를 설정해주세요.")
        st.stop()

    uploaded_files = st.file_uploader("이미지 업로드", accept_multiple_files=True)
    if uploaded_files:
        slide_inputs = []
        temp_paths = []
        for idx, f in enumerate(uploaded_files):
            ko = st.text_input(f"설명 #{idx+1}", value="안전모 착용", key=f"ko_{idx}")
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(f.getbuffer())
                temp_paths.append(tmp.name)
                slide_inputs.append(SlideData(tmp.name, ko, "", "", ""))

        if st.button("PPT 생성"):
            try:
                translations = translate_batch_with_gemini(st.secrets["GEMINI_API_KEY"], [s.ko for s in slide_inputs])
                for s, tr in zip(slide_inputs, translations):
                    s.zh, s.vi, s.my = tr['zh'], tr['vi'], tr['my']
                
                ppt_file = build_ppt(slide_inputs)
                st.download_button("PPT 다운로드", data=ppt_file, file_name="TBM_Result.pptx")
            except Exception as e:
                st.error(f"실행 에러: {e}")
            finally:
                for p in temp_paths: 
                    if os.path.exists(p): os.remove(p)

if __name__ == "__main__":
    main()