import os
import io
import json
import tempfile
from dataclasses import dataclass
from typing import List

import requests
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
BASE_FONT_SIZE_PT = 35
TITLE_TEXT = "설명/说明"
OUTPUT_PPT_NAME = "TBM_완성본.pptx"
APP_VERSION = "GPT-PLACEHOLDER-PHOTOBOX-01"

PHOTO_BOX_TEXT = "PHOTO_BOX"
KO_BOX_TEXT = "1"
ZH_BOX_TEXT = "2"
VI_BOX_TEXT = "3"
MY_BOX_TEXT = "4"


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


def clear_and_set_text(shape, text: str, size_pt: int):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)


def find_shape_by_exact_text(slide, target_text: str):
    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape) and get_text(shape) == target_text:
            return shape
    return None


def add_picture_cover(slide, image_path, target_shape):
    left = target_shape.left
    top = target_shape.top
    width = target_shape.width
    height = target_shape.height

    pic = slide.shapes.add_picture(
        image_path,
        left,
        top,
        width=width,
        height=height
    )

    image_w = 1.0
    image_h = 1.0
    box_ratio = float(width) / float(height)

    from PIL import Image
    with Image.open(image_path) as img:
        image_w, image_h = img.size

    image_ratio = float(image_w) / float(image_h)

    if image_ratio > box_ratio:
        crop = (1.0 - (box_ratio / image_ratio)) / 2.0
        pic.crop_left = crop
        pic.crop_right = crop
        pic.crop_top = 0
        pic.crop_bottom = 0
    else:
        crop = (1.0 - (image_ratio / box_ratio)) / 2.0
        pic.crop_top = crop
        pic.crop_bottom = crop
        pic.crop_left = 0
        pic.crop_right = 0


def fill_slide_by_placeholders(slide, item: SlideData):
    photo_shape = find_shape_by_exact_text(slide, PHOTO_BOX_TEXT)
    ko_shape = find_shape_by_exact_text(slide, KO_BOX_TEXT)
    zh_shape = find_shape_by_exact_text(slide, ZH_BOX_TEXT)
    vi_shape = find_shape_by_exact_text(slide, VI_BOX_TEXT)
    my_shape = find_shape_by_exact_text(slide, MY_BOX_TEXT)

    missing = []
    for name, shp in [
        ("PHOTO_BOX", photo_shape),
        ("1", ko_shape),
        ("2", zh_shape),
        ("3", vi_shape),
        ("4", my_shape),
    ]:
        if shp is None:
            missing.append(name)

    if missing:
        raise ValueError(f"슬라이드에서 플레이스홀더를 찾지 못했습니다: {', '.join(missing)}")

    add_picture_cover(slide, item.image_path, photo_shape)

    clear_and_set_text(ko_shape, item.ko, BASE_FONT_SIZE_PT)
    clear_and_set_text(zh_shape, item.zh, BASE_FONT_SIZE_PT)
    clear_and_set_text(vi_shape, item.vi, BASE_FONT_SIZE_PT)
    clear_and_set_text(my_shape, item.my, BASE_FONT_SIZE_PT)


def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    if not os.path.exists(TEMPLATE_PPT):
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {TEMPLATE_PPT}")

    prs = Presentation(TEMPLATE_PPT)

    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides):
            break
        slide = prs.slides[i]
        fill_slide_by_placeholders(slide, item)

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

    if "GPT_API_KEY" not in st.secrets:
        st.warning("Secrets에 GPT_API_KEY 설정 필요")
        st.stop()

    files = st.file_uploader(
        "사진 업로드",
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg", "webp"]
    )

    if files:
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

                suffix = os.path.splitext(f.name)[1].lower() or ".jpg"

                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(f.getbuffer())
                    temp_paths.append(tmp.name)
                    slide_inputs.append(SlideData(tmp.name, ko_input, "", "", ""))

        if st.button("PPT 생성"):
            try:
                with st.spinner("번역 중..."):
                    ko_list = [s.ko for s in slide_inputs]

                    if any(not x.strip() for x in ko_list):
                        raise ValueError("빈 한국어 문구가 있습니다. 모든 슬라이드 문구를 입력하세요.")

                    translations = translate_batch_with_gpt(
                        st.secrets["GPT_API_KEY"],
                        ko_list
                    )

                    for s, tr in zip(slide_inputs, translations):
                        s.zh = tr["zh"]
                        s.vi = tr["vi"]
                        s.my = tr["my"]

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