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

# --- 설정 및 상수 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
BASE_FONT_SIZE_PT = 35
TITLE_TEXT = "설명/说明"
OUTPUT_PPT_NAME = "TBM_완성본.pptx"

@dataclass
class SlideData:
    image_path: str
    ko: str
    zh: str
    vi: str
    my: str

def translate_batch_with_gemini(api_key: str, korean_list: List[str]):
    # v1 엔드포인트 사용
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={api_key}"
    
    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])

    # 프롬프트를 명확하게 구조화
    prompt = f"""Translate the following construction safety phrases into Chinese (zh), Vietnamese (vi), and Burmese (my).
Provide the output ONLY as a valid JSON array of objects.

Input:
{joined_text}

Output format example:
[
  {{"zh": "중국어", "vi": "베트남어", "my": "미얀마어"}}
]"""

    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "response_mime_type": "application/json",
            "temperature": 0.1 # 일관된 출력을 위해 온도를 낮춤
        }
    }

    try:
        resp = requests.post(
            url,
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=90
        )
        
        # 400 에러 등이 발생했을 때 구체적인 이유를 화면에 표시
        if resp.status_code != 200:
            error_detail = resp.json().get('error', {}).get('message', 'Unknown error')
            st.error(f"API 요청 실패 ({resp.status_code}): {error_detail}")
            if "key" in error_detail.lower():
                st.error("💡 API 키가 유효하지 않거나 차단된 것 같습니다. 새 키를 발급받으세요.")
            resp.raise_for_status()

        data = resp.json()
        text_content = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        parsed_result = json.loads(text_content)
        
        if len(parsed_result) != len(korean_list):
            raise ValueError("번역된 문구 개수가 입력과 다릅니다.")
            
        return parsed_result

    except Exception as e:
        st.error(f"번역 로직 에러: {str(e)}")
        raise e

# --- PPT 편집 로직 (안정화 버전) ---
def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes):
                yield sub_shape

def has_text(shape) -> bool:
    return hasattr(shape, "has_text_frame") and shape.has_text_frame

def get_text(shape) -> str:
    return shape.text.strip() if has_text(shape) else ""

def find_picture_area(slide):
    # '사진대지' 텍스트가 있는 도형 우선 검색
    for shape in iter_all_shapes(slide.shapes):
        if "사진대지" in get_text(shape):
            return shape
    # 없으면 가장 큰 상단 도형 반환
    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        try:
            candidates.append((shape.top, -(shape.width * shape.height), shape))
        except: pass
    candidates.sort()
    return candidates[0][2]

def insert_image(slide, image_path, area_shape):
    left, top, width, height = area_shape.left, area_shape.top, area_shape.width, area_shape.height
    if has_text(area_shape): area_shape.text_frame.clear()
    
    with Image.open(image_path) as img:
        img_w, img_h = img.size
    
    # 비율에 맞춰 꽉 채우기 (Aspect Fill)
    box_ratio, img_ratio = width / height, img_w / img_h
    if img_ratio > box_ratio:
        new_h = height
        new_w = int(height * img_ratio)
    else:
        new_w = width
        new_h = int(width / img_ratio)

    slide.shapes.add_picture(image_path, left + (width - new_w)//2, top + (height - new_h)//2, width=new_w, height=new_h)

def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    prs = Presentation(TEMPLATE_PPT)
    
    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides): break
        slide = prs.slides[i]
        
        pic_area = find_picture_area(slide)
        insert_image(slide, item.image_path, pic_area)
        
        # 사진 아래에 있는 텍스트 상자들 찾기
        pic_bottom = pic_area.top + pic_area.height
        txt_shapes = [s for s in iter_all_shapes(slide.shapes) if has_text(s) and s.top >= pic_bottom - Pt(5)]
        txt_shapes.sort(key=lambda s: (s.top, s.left))
        
        texts = [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]
        for shape, txt in zip(txt_shapes, texts):
            tf = shape.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = txt
            run.font.size = Pt(BASE_FONT_SIZE_PT)
            
    # 남는 슬라이드 삭제
    for idx in range(len(prs.slides) - 1, len(slide_data_list) - 1, -1):
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# --- UI 메인 ---
def main():
    st.set_page_config(page_title="TBM Generator", layout="wide")
    st.title("🏗️ 안전 TBM 다국어 자료 생성기")

    if "GEMINI_API_KEY" not in st.secrets:
        st.error("Secrets 설정에 'GEMINI_API_KEY'를 입력해주세요.")
        st.stop()

    uploaded_files = st.file_uploader("사진 업로드", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

    if uploaded_files:
        slide_data = []
        temp_paths = []
        
        for idx, file in enumerate(uploaded_files):
            with st.expander(f"슬라이드 {idx+1} 설정", expanded=True):
                c1, c2 = st.columns([1, 3])
                c1.image(file, width=150)
                ko = c2.text_input(f"설명 (한국어)", value="안전통로를 이용합시다", key=f"ko_{idx}")
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp.write(file.getbuffer())
                    temp_paths.append(tmp.name)
                    slide_data.append(SlideData(tmp.name, ko, "", "", ""))

        if st.button("PPT 생성하기"):
            try:
                with st.spinner("AI 번역 및 슬라이드 구성 중..."):
                    # 번역
                    ko_texts = [s.ko for s in slide_data]
                    translations = translate_batch_with_gemini(st.secrets["GEMINI_API_KEY"], ko_texts)
                    
                    for s, tr in zip(slide_data, translations):
                        s.zh, s.vi, s.my = tr['zh'], tr['vi'], tr['my']
                    
                    # PPT 생성
                    ppt_file = build_ppt(slide_data)
                    st.success("PPT 생성 성공!")
                    st.download_button("PPT 다운로드", data=ppt_file, file_name=OUTPUT_PPT_NAME)
            except Exception as e:
                st.error(f"처리 중 오류가 발생했습니다: {e}")
            finally:
                for p in temp_paths:
                    if os.path.exists(p): os.remove(p)

if __name__ == "__main__":
    main()