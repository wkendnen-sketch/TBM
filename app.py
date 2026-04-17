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

# --- 상수 설정 ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
BASE_FONT_SIZE_PT = 30
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
    # URL 및 API 키 공백 제거 (Connection Adapter 에러 방지 핵심)
    api_key = api_key.strip()
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}".strip()
    
    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])
    
    prompt = f"""Translate the following construction safety phrases into Chinese (zh), Vietnamese (vi), and Burmese (my).
Return ONLY a JSON array of objects with keys "zh", "vi", "my".
No preamble, no markdown code blocks.

Input:
{joined_text}"""

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.1}
    }

    try:
        resp = requests.post(
            url,
            headers={"Content-Type": "application/json"},
            json=payload,
            timeout=90
        )
        
        if resp.status_code != 200:
            st.error(f"API 응답 오류 ({resp.status_code}): {resp.text}")
            resp.raise_for_status()
            
        data = resp.json()
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        
        # 마크다운 코드 블록(```json)이 섞여 나올 경우 정제
        clean_json = raw_text
        if "```" in raw_text:
            parts = raw_text.split("```")
            for part in parts:
                if "zh" in part and "vi" in part: # JSON 내용이 포함된 부분 찾기
                    clean_json = part.replace("json", "").strip()
                    break
        
        return json.loads(clean_json)

    except Exception as e:
        st.error(f"번역 프로세스 중 오류 발생: {str(e)}")
        raise e

# --- PPT 처리 로직 ---
def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes):
                yield sub_shape

def find_picture_area(slide):
    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        # '사진대지' 텍스트 박스 우선 찾기
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            if "사진대지" in shape.text:
                return shape
        try:
            # 대체 수단: 상단에 위치한 큰 도형
            candidates.append((shape.top, -(shape.width * shape.height), shape))
        except: pass
    candidates.sort()
    return candidates[0][2]

def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    if not os.path.exists(TEMPLATE_PPT):
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {TEMPLATE_PPT}")
        
    prs = Presentation(TEMPLATE_PPT)
    
    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides): break
        slide = prs.slides[i]
        
        # 1. 사진 삽입
        pic_area = find_picture_area(slide)
        slide.shapes.add_picture(item.image_path, pic_area.left, pic_area.top, width=pic_area.width, height=pic_area.height)
        
        # 2. 텍스트 삽입 (사진 하단 텍스트 박스 순서대로)
        pic_bottom = pic_area.top + pic_area.height
        txt_shapes = [s for s in iter_all_shapes(slide.shapes) 
                      if hasattr(s, "has_text_frame") and s.has_text_frame and s.top >= pic_bottom - 50000]
        txt_shapes.sort(key=lambda s: (s.top, s.left))
        
        contents = [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]
        for shape, txt in zip(txt_shapes, contents):
            tf = shape.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = txt
            run.font.size = Pt(BASE_FONT_SIZE_PT)

    # 남은 빈 슬라이드 삭제
    for idx in range(len(prs.slides) - 1, len(slide_data_list) - 1, -1):
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# --- 메인 앱 ---
def main():
    st.set_page_config(page_title="TBM PPT Maker", layout="wide")
    st.title("🚧 TBM 교육자료 자동 번역 생성기")

    # API 키 확인
    if "GEMINI_API_KEY" not in st.secrets:
        st.warning("Streamlit Secrets에 'GEMINI_API_KEY'를 설정해주세요.")
        st.stop()

    files = st.file_uploader("사진을 업로드하세요", accept_multiple_files=True, type=['jpg', 'png', 'jpeg'])

    if files:
        slide_inputs = []
        temp_paths = []
        
        for idx, f in enumerate(files):
            with st.expander(f"슬라이드 #{idx+1} 설정", expanded=True):
                c1, c2 = st.columns([1, 4])
                c1.image(f, width=150)
                ko_input = c2.text_input(f"한국어 설명", value="안전모를 반드시 착용합시다", key=f"ko_{idx}")
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp.write(f.getbuffer())
                    temp_paths.append(tmp.name)
                    slide_inputs.append(SlideData(tmp.name, ko_input, "", "", ""))

        if st.button("AI 번역 및 PPT 다운로드"):
            try:
                with st.spinner("AI가 다국어로 번역 중입니다..."):
                    ko_list = [s.ko for s in slide_inputs]
                    translations = translate_batch_with_gemini(st.secrets["GEMINI_API_KEY"], ko_list)
                    
                    for s, tr in zip(slide_inputs, translations):
                        s.zh, s.vi, s.my = tr['zh'], tr['vi'], tr['my']
                
                with st.spinner("PPT 파일을 생성 중입니다..."):
                    final_ppt = build_ppt(slide_inputs)
                    st.success("모든 작업이 완료되었습니다!")
                    st.download_button(
                        label="📁 완성된 PPT 다운로드",
                        data=final_ppt,
                        file_name=OUTPUT_PPT_NAME,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            except Exception as e:
                st.error(f"오류 발생: {e}")
            finally:
                # 임시 파일 삭제
                for p in temp_paths:
                    if os.path.exists(p): os.remove(p)

if __name__ == "__main__":
    main()