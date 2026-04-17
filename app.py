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
    # v1에서 찾을 수 없다는 에러가 나면 v1beta를 사용해야 합니다.
    url = f"[https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=){api_key}"
    
    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])
    
    # 프롬프트를 아주 단순하고 명확하게 작성합니다.
    prompt = f"""Translate the following Korean phrases into Chinese (zh), Vietnamese (vi), and Burmese (my).
Return ONLY a JSON array of objects with keys "zh", "vi", "my".
No conversation, no markdown code blocks.

Input:
{joined_text}"""

    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "temperature": 0.2
        }
    }

    resp = requests.post(
        url, 
        headers={"Content-Type": "application/json"}, 
        json=payload, 
        timeout=90
    )
    
    if resp.status_code != 200:
        st.error(f"API 상세 에러: {resp.text}")
        resp.raise_for_status()
    
    data = resp.json()
    raw_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
    
    # 모델이 마크다운(```json)을 섞어 보낼 경우를 대비한 정제 로직
    clean_json = raw_text
    if "```" in raw_text:
        clean_json = raw_text.split("```")[1]
        if clean_json.startswith("json"):
            clean_json = clean_json[4:].strip()
    
    try:
        return json.loads(clean_json)
    except json.JSONDecodeError:
        st.error(f"JSON 파싱 실패. 응답 내용: {raw_text}")
        raise

# --- PPT 빌드 및 UI 로직 (기존과 동일하되 안정화) ---

def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes):
                yield sub_shape

def find_picture_area(slide):
    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        # '사진대지' 텍스트 포함 여부 확인
        txt = ""
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            txt = shape.text.strip()
        if "사진대지" in txt:
            return shape
        try:
            candidates.append((shape.top, -(shape.width * shape.height), shape))
        except:
            pass
    return sorted(candidates)[0][2]

def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    if not os.path.exists(TEMPLATE_PPT):
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {TEMPLATE_PPT}")
        
    prs = Presentation(TEMPLATE_PPT)
    
    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides): break
        slide = prs.slides[i]
        
        # 1. 사진 영역 찾기 및 삽입
        pic_area = find_picture_area(slide)
        slide.shapes.add_picture(item.image_path, pic_area.left, pic_area.top, width=pic_area.width, height=pic_area.height)
        
        # 2. 텍스트 영역 순차 채우기 (사진 아래 영역 대상)
        pic_bottom = pic_area.top + pic_area.height
        txt_shapes = [s for s in iter_all_shapes(slide.shapes) 
                      if hasattr(s, "has_text_frame") and s.has_text_frame and s.top >= pic_bottom - 50000] # 약간의 오차 허용
        txt_shapes.sort(key=lambda s: (s.top, s.left))
        
        fill_texts = [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]
        for shape, txt in zip(txt_shapes, fill_texts):
            tf = shape.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = txt
            run.font.size = Pt(BASE_FONT_SIZE_PT)

    # 사용하지 않는 슬라이드 삭제
    for idx in range(len(prs.slides) - 1, len(slide_data_list) - 1, -1):
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def main():
    st.set_page_config(page_title="TBM PPT", layout="wide")
    st.title("🏗️ TBM 교육자료 생성기 (v1beta 대응)")
    
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("Streamlit Secrets에 API 키가 없습니다.")
        st.stop()

    files = st.file_uploader("사진 업로드", accept_multiple_files=True, type=['jpg', 'jpeg', 'png'])
    
    if files:
        slide_inputs = []
        temp_paths = []
        
        for idx, f in enumerate(files):
            col1, col2 = st.columns([1, 4])
            col1.image(f, width=150)
            ko = col2.text_input(f"설명 입력 #{idx+1}", value="지정된 통로로 이동", key=f"ko_{idx}")
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(f.getbuffer())
                temp_paths.append(tmp.name)
                slide_inputs.append(SlideData(tmp.name, ko, "", "", ""))

        if st.button("PPT 생성 및 AI 번역"):
            try:
                with st.spinner("AI 번역 중... (v1beta)"):
                    kos = [s.ko for s in slide_inputs]
                    results = translate_batch_with_gemini(st.secrets["GEMINI_API_KEY"], kos)
                    
                    for s, tr in zip(slide_inputs, results):
                        s.zh, s.vi, s.my = tr['zh'], tr['vi'], tr['my']
                
                with st.spinner("PPT 파일 제작 중..."):
                    ppt_file = build_ppt(slide_inputs)
                    st.success("완료!")
                    st.download_button("PPT 다운로드", data=ppt_file, file_name="TBM_Education.pptx")
                    
            except Exception as e:
                st.error(f"실행 에러: {e}")
            finally:
                for p in temp_paths:
                    if os.path.exists(p): os.remove(p)

if __name__ == "__main__":
    main()