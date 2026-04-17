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
    # 1. 모델명을 안정적인 1.5 버전으로 변경 제안
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"gemini-1.5-flash:generateContent?key={api_key}"
    )

    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])

    prompt = f"""
다음 한국어 안전 문구들을 건설현장 TBM 교육자료용으로 자연스럽고 짧게 번역하라.
결과는 반드시 JSON 배열 형태로만 출력하라.

입력:
{joined_text}

출력형식:
[
  {{ "zh": "중국어", "vi": "베트남어", "my": "미얀마어" }}
]
"""

    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        # 2. JSON 응답을 강제하는 설정 추가 (파싱 에러 방지)
        "generationConfig": {
            "response_mime_type": "application/json"
        }
    }

    last_error = None
    # 3. 재시도 횟수 증가 및 대기 시간 전략 수정
    max_retries = 3
    for attempt in range(max_retries):
        try:
            resp = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=90
            )
            
            # 429 에러(Quota Exceeded) 처리 강화
            if resp.status_code == 429:
                wait_time = (attempt + 1) * 12  # 시도할수록 대기 시간 증가 (12초, 24초...)
                st.warning(f"API 요청 한도 초과. {wait_time}초 후 다시 시도합니다... ({attempt + 1}/{max_retries})")
                time.sleep(wait_time)
                continue

            resp.raise_for_status()
            data = resp.json()
            
            # 응답 데이터 추출
            text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
            parsed = json.loads(text)

            if not isinstance(parsed, list) or len(parsed) != len(korean_list):
                raise ValueError("번역 결과 개수가 일치하지 않습니다.")

            return parsed

        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                time.sleep(5)
                continue
            raise e

    raise last_error

# --- PPT 조작 관련 함수들 (기존과 동일) ---
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

def find_text_shapes_for_content(slide, picture_area):
    pic_bottom = picture_area.top + picture_area.height
    candidates = [s for s in iter_all_shapes(slide.shapes) if has_text(s) and s.top >= pic_bottom - Pt(5)]
    candidates.sort(key=lambda s: (s.top, s.left))
    return candidates

def rewrite_slide_texts(slide, picture_area, item: SlideData):
    text_shapes = find_text_shapes_for_content(slide, picture_area)
    if len(text_shapes) < 5: raise ValueError(f"텍스트 상자 부족 (필요 5개, 현재 {len(text_shapes)}개)")
    
    targets = text_shapes[:5]
    new_texts = [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]
    for shape, txt in zip(targets, new_texts):
        clear_and_set_text(shape, txt, BASE_FONT_SIZE_PT)

def delete_slide(prs, index: int):
    xml_slides = prs.slides._sldIdLst
    slide_id = xml_slides[index]
    prs.part.drop_rel(slide_id.rId)
    xml_slides.remove(slide_id)

def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    prs = Presentation(TEMPLATE_PPT)
    if len(prs.slides) < len(slide_data_list):
        raise ValueError(f"템플릿 슬라이드 부족 (사진 {len(slide_data_list)}장 / 슬라이드 {len(prs.slides)}장)")
    
    for i, item in enumerate(slide_data_list):
        slide = prs.slides[i]
        pic_area = find_picture_area(slide)
        insert_image_cover(slide, item.image_path, pic_area)
        rewrite_slide_texts(slide, pic_area, item)
    
    for idx in range(len(prs.slides) - 1, len(slide_data_list) - 1, -1):
        delete_slide(prs, idx)
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- 메인 실행 로직 ---
def main():
    st.set_page_config(page_title="TBM PPT 생성기", layout="wide")
    st.title("🏗️ TBM 교육자료 자동 생성기")

    if not os.path.exists(TEMPLATE_PPT):
        st.error(f"템플릿 파일 없음: {TEMPLATE_PPT}")
        st.stop()

    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except:
        st.error("Streamlit Secrets에 'GEMINI_API_KEY'가 설정되지 않았습니다.")
        st.stop()

    uploaded_files = st.file_uploader("사진 업로드", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True)
    if not uploaded_files:
        st.info("번역할 사진들을 먼저 업로드해주세요.")
        st.stop()

    slide_data_list = []
    temp_paths = []

    # 입력 UI
    for idx, file in enumerate(uploaded_files, start=1):
        with st.expander(f"슬라이드 {idx} 내용 입력", expanded=True):
            col1, col2 = st.columns([1, 3])
            with col1:
                st.image(file, width=150)
            with col2:
                ko_text = st.text_input(f"한국어 문구 #{idx}", value="안전모 미착용 금지", key=f"ko_{idx}")
            
            suffix = os.path.splitext(file.name)[1] or ".jpg"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(file.getbuffer())
                temp_paths.append(tmp.name)
                slide_data_list.append(SlideData(image_path=tmp.name, ko=ko_text, zh="", vi="", my=""))

    if st.button("PPT 생성 및 번역 시작"):
        try:
            with st.spinner("Gemini AI가 실시간 번역 중입니다..."):
                ko_list = [d.ko for d in slide_data_list]
                translations = translate_batch_with_gemini(api_key, ko_list)

                for item, tr in zip(slide_data_list, translations):
                    item.zh, item.vi, item.my = tr["zh"], tr["vi"], tr["my"]

                ppt_data = build_ppt(slide_data_list)
                
            st.success("✅ PPT 생성이 완료되었습니다!")
            st.download_button(
                label="📁 완성된 PPT 다운로드",
                data=ppt_data,
                file_name=OUTPUT_PPT_NAME,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        except Exception as e:
            st.error(f"오류 발생: {e}")
        finally:
            for p in temp_paths:
                if os.path.exists(p): os.remove(p)

if __name__ == "__main__":
    main()