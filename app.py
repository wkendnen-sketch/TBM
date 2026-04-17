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
# 템플릿 경로를 실제 환경에 맞게 조정하세요.
TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
BASE_FONT_SIZE_PT = 35
TITLE_FONT_SIZE_PT = 35
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
    # 1. 경로를 v1으로 변경하여 안정성 확보
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={api_key}"

    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])

    prompt = f"""
    Translate the following Korean safety phrases for construction TBM education into Chinese (zh), Vietnamese (vi), and Burmese (my).
    Return ONLY a JSON array. Do not include markdown code blocks.
    
    Phrases:
    {joined_text}

    JSON Output Format:
    [
      {{ "zh": "번역", "vi": "번역", "my": "번역" }}
    ]
    """

    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "response_mime_type": "application/json" # JSON 응답 강제
        }
    }

    max_retries = 3
    for attempt in range(max_retries):
        try:
            resp = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=90
            )
            
            # 상세 에러 메시지 처리
            if resp.status_code == 404:
                raise Exception("API 엔드포인트를 찾을 수 없습니다 (404). 모델명이나 URL을 확인하세요.")
            if resp.status_code == 403:
                raise Exception("API 키 권한 오류 (403). 키가 올바른지, 혹은 노출되어 차단되지 않았는지 확인하세요.")
            if resp.status_code == 429:
                wait_time = (attempt + 1) * 15
                st.warning(f"요청 한도 초과. {wait_time}초 후 재시도... ({attempt+1}/{max_retries})")
                time.sleep(wait_time)
                continue

            resp.raise_for_status()
            data = resp.json()
            
            # 응답 텍스트 파싱
            text_content = data["candidates"][0]["content"]["parts"][0]["text"].strip()
            parsed = json.loads(text_content)

            if len(parsed) != len(korean_list):
                raise ValueError(f"개수 불일치: 입력 {len(korean_list)} / 결과 {len(parsed)}")

            return parsed

        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(5)
                continue
            raise e

# --- PPT 처리 함수 (최적화 버전) ---
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

def clear_and_set_text(shape, text: str, font_size: int):
    if not has_text(shape): return
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)

def find_picture_area(slide):
    for shape in iter_all_shapes(slide.shapes):
        if "사진대지" in get_text(shape): return shape
    
    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            try:
                candidates.append((shape.top, -(shape.width * shape.height), shape))
            except: pass
    if not candidates: raise ValueError("사진 영역을 찾지 못했습니다.")
    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[0][2]

def insert_image_cover(slide, image_path, area_shape):
    left, top, width, height = area_shape.left, area_shape.top, area_shape.width, area_shape.height
    if has_text(area_shape): area_shape.text_frame.clear()
    
    with Image.open(image_path) as img:
        img_w, img_h = img.size
    
    box_ratio, img_ratio = width / height, img_w / img_h
    if img_ratio > box_ratio:
        new_h = height
        new_w = int(height * img_ratio)
    else:
        new_w = width
        new_h = int(width / img_ratio)

    pic_left = left + (width - new_w) // 2
    pic_top = top + (height - new_h) // 2
    slide.shapes.add_picture(image_path, pic_left, pic_top, width=new_w, height=new_h)

def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    prs = Presentation(TEMPLATE_PPT)
    if len(prs.slides) < len(slide_data_list):
        raise ValueError("템플릿 슬라이드 수가 부족합니다.")
    
    for i, item in enumerate(slide_data_list):
        slide = prs.slides[i]
        pic_area = find_picture_area(slide)
        insert_image_cover(slide, item.image_path, pic_area)
        
        # 텍스트 교체 (상단 사진 제외하고 텍스트 박스 순서대로)
        pic_bottom = pic_area.top + pic_area.height
        txt_shapes = [s for s in iter_all_shapes(slide.shapes) if has_text(s) and s.top >= pic_bottom - Pt(5)]
        txt_shapes.sort(key=lambda s: (s.top, s.left))
        
        new_texts = [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]
        for shape, txt in zip(txt_shapes[:5], new_texts):
            clear_and_set_text(shape, txt, BASE_FONT_SIZE_PT)
    
    # 남는 슬라이드 삭제
    for idx in range(len(prs.slides) - 1, len(slide_data_list) - 1, -1):
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- Streamlit 앱 메인 ---
def main():
    st.set_page_config(page_title="TBM PPT Generator", layout="wide")
    st.title("🚧 TBM 다국어 PPT 자동 생성기")

    if not os.path.exists(TEMPLATE_PPT):
        st.error(f"템플릿을 찾을 수 없습니다: {TEMPLATE_PPT}")
        st.stop()

    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except:
        st.error("Streamlit Secrets에 API 키를 설정해주세요.")
        st.stop()

    files = st.file_uploader("사진을 선택하세요", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
    if files:
        slide_inputs = []
        temp_files = []
        
        for idx, f in enumerate(files):
            with st.expander(f"슬라이드 {idx+1} 설정", expanded=True):
                col1, col2 = st.columns([1, 4])
                col1.image(f, width=150)
                ko_txt = col2.text_input(f"한국어 설명", value="안전모 착용 철저", key=f"ko_{idx}")
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp.write(f.getbuffer())
                    temp_files.append(tmp.name)
                    slide_inputs.append(SlideData(image_path=tmp.name, ko=ko_txt, zh="", vi="", my=""))

        if st.button("PPT 생성 (AI 번역 포함)"):
            try:
                with st.spinner("AI 번역 및 PPT 제작 중..."):
                    # 번역 실행
                    ko_list = [s.ko for s in slide_inputs]
                    results = translate_batch_with_gemini(api_key, ko_list)
                    
                    for item, res in zip(slide_inputs, results):
                        item.zh, item.vi, item.my = res['zh'], res['vi'], res['my']
                    
                    # PPT 빌드
                    ppt_out = build_ppt(slide_inputs)
                    
                st.success("완료!")
                st.download_button("PPT 다운로드", data=ppt_out, file_name=OUTPUT_PPT_NAME)
                
            except Exception as e:
                st.error(f"오류 발생: {e}")
            finally:
                for p in temp_files:
                    if os.path.exists(p): os.remove(p)

if __name__ == "__main__":
    main()