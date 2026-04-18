import os
import io
import json
import tempfile
from dataclasses import dataclass
from typing import List

import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
BASE_FONT_SIZE_PT = 30
TITLE_TEXT = "설명/说明"
OUTPUT_PPT_NAME = "TBM_완성본.pptx"
APP_VERSION = "GPT-DEBUG-2026-04-18-01"


@dataclass
class SlideData:
    image_path: str
    ko: str
    zh: str
    vi: str
    my: str


def translate_batch_with_gpt(api_key: str, korean_list: List[str]):
    url = "https://api.openai.com/v1/responses"

    joined_text = "\n".join([f"{i+1}. {txt}" for i, txt in enumerate(korean_list)])

    prompt = f"""
다음 한국어 안전 문구들을 건설현장 TBM용으로 짧고 명확하게 번역하라.

조건:
- 반드시 JSON 배열만 출력
- 설명 금지
- 코드블록 금지
- 각 항목은 zh, vi, my 포함
- 입력 개수와 출력 개수는 반드시 같아야 함

입력:
{joined_text}

출력:
[
 {{
  "zh":"중국어",
  "vi":"베트남어",
  "my":"미얀마어"
 }}
]
"""

    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": "gpt-4o-mini",
        "input": prompt
    }

    resp = requests.post(url, headers=headers, json=payload, timeout=60)

    if resp.status_code != 200:
        raise Exception(f"API Error: {resp.text}")

    data = resp.json()

    text = ""
    if "output_text" in data and data["output_text"]:
        text = data["output_text"]
    else:
        for item in data.get("output", []):
            for c in item.get("content", []):
                if c.get("type") == "output_text":
                    text += c.get("text", "")

    text = text.replace("```json", "").replace("```", "").strip()
    parsed = json.loads(text)

    if not isinstance(parsed, list):
        raise ValueError("GPT 응답이 배열이 아닙니다.")

    if len(parsed) != len(korean_list):
        raise ValueError(
            f"번역 개수 불일치: 입력 {len(korean_list)} / 출력 {len(parsed)}"
        )

    return parsed


def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes):
                yield sub_shape


def has_text(shape):
    return hasattr(shape, "has_text_frame") and shape.has_text_frame


def get_text(shape):
    if has_text(shape):
        return shape.text.strip()
    return ""


def is_red_fill(shape):
    try:
        if not hasattr(shape, "fill"):
            return False
        if shape.fill is None:
            return False
        fore = shape.fill.fore_color
        if not hasattr(fore, "rgb") or fore.rgb is None:
            return False
        rgb = str(fore.rgb).upper()
        return rgb in ["FF0000", "C00000", "FF1F1F", "D32F2F"]
    except Exception:
        return False


def find_picture_area(slide):
    red_shapes = []
    for shape in iter_all_shapes(slide.shapes):
        try:
            if is_red_fill(shape):
                area = shape.width * shape.height
                red_shapes.append((shape.top, -area, shape))
        except Exception:
            pass

    if red_shapes:
        red_shapes.sort(key=lambda x: (x[0], x[1]))
        return red_shapes[0][2]

    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape) and "사진대지" in get_text(shape):
            return shape

    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        try:
            area = shape.width * shape.height
            candidates.append((shape.top, -area, shape))
        except Exception:
            pass

    if not candidates:
        raise ValueError("사진 영역을 찾지 못했습니다.")

    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[0][2]


def add_picture_cover(slide, image_path, target_shape):
    left = target_shape.left
    top = target_shape.top
    width = target_shape.width
    height = target_shape.height

    with Image.open(image_path) as img:
        img_w, img_h = img.size

    img_ratio = img_w / img_h
    box_ratio = width / height

    pic = slide.shapes.add_picture(
        image_path,
        left,
        top,
        width=width,
        height=height
    )

    if img_ratio > box_ratio:
        crop = (1 - (box_ratio / img_ratio)) / 2
        pic.crop_left = crop
        pic.crop_right = crop
        pic.crop_top = 0
        pic.crop_bottom = 0
    else:
        crop = (1 - (img_ratio / box_ratio)) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
        pic.crop_left = 0
        pic.crop_right = 0


def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    prs = Presentation(TEMPLATE_PPT)

    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides):
            break

        slide = prs.slides[i]

        pic_area = find_picture_area(slide)
        add_picture_cover(slide, item.image_path, pic_area)

        pic_bottom = pic_area.top + pic_area.height
        txt_shapes = [
            s for s in iter_all_shapes(slide.shapes)
            if has_text(s) and s.top >= pic_bottom - 50000
        ]
        txt_shapes.sort(key=lambda s: (s.top, s.left))

        contents = [TITLE_TEXT, item.ko, item.zh, item.vi, item.my]

        for shape, txt in zip(txt_shapes, contents):
            tf = shape.text_frame
            tf.clear()
            run = tf.paragraphs[0].add_run()
            run.text = txt
            run.font.size = Pt(BASE_FONT_SIZE_PT)

    for idx in range(len(prs.slides) - 1, len(slide_data_list) - 1, -1):
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out


def main():
    st.set_page_config(page_title="TBM PPT Maker", layout="wide")
    st.title(f"🚧 TBM 교육자료 자동 번역 생성기 [{APP_VERSION}]")

    st.caption(f"템플릿 경로: {TEMPLATE_PPT}")
    st.caption(f"템플릿 존재 여부: {os.path.exists(TEMPLATE_PPT)}")

    if "GPT_API_KEY" not in st.secrets:
        st.warning("Secrets에 GPT_API_KEY 설정 필요")
        st.stop()

    files = st.file_uploader(
        "사진 업로드",
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg"]
    )

    if files:
        st.write("업로드 파일 개수:", len(files))

        slide_inputs = []
        temp_paths = []

        for idx, f in enumerate(files):
            with st.expander(f"슬라이드 #{idx+1}", expanded=True):
                c1, c2 = st.columns([1, 4])
                c1.image(f, width=150)

                ko_input = c2.text_input(
                    "한국어 문구",
                    value="",
                    placeholder="예: 지정된 이동통로 통행",
                    key=f"ko_{idx}"
                )

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    tmp.write(f.getbuffer())
                    temp_paths.append(tmp.name)

                    slide_inputs.append(
                        SlideData(tmp.name, ko_input, "", "", "")
                    )

        if st.button("PPT 생성"):
            try:
                with st.spinner("번역 중..."):
                    ko_list = [s.ko for s in slide_inputs]

                    st.subheader("디버그: 입력 한국어 목록")
                    st.write(ko_list)

                    translations = translate_batch_with_gpt(
                        st.secrets["GPT_API_KEY"], ko_list
                    )

                    st.subheader("디버그: GPT 번역 결과")
                    st.json(translations)

                    for s, tr in zip(slide_inputs, translations):
                        s.zh = tr["zh"]
                        s.vi = tr["vi"]
                        s.my = tr["my"]

                    debug_slide_data = [
                        {
                            "ko": s.ko,
                            "zh": s.zh,
                            "vi": s.vi,
                            "my": s.my,
                            "image_path": s.image_path
                        }
                        for s in slide_inputs
                    ]

                    st.subheader("디버그: PPT 반영 직전 데이터")
                    st.json(debug_slide_data)

                with st.spinner("PPT 생성 중..."):
                    ppt = build_ppt(slide_inputs)

                st.success("완료!")
                st.download_button(
                    "PPT 다운로드",
                    ppt,
                    file_name=OUTPUT_PPT_NAME,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except Exception as e:
                st.error(f"오류 발생: {e}")

            finally:
                for p in temp_paths:
                    if os.path.exists(p):
                        os.remove(p)


if __name__ == "__main__":
    main()