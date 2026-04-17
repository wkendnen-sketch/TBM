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


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PPT = os.path.join(
    BASE_DIR,
    "templates",
    "sample_template.pptx"
)

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
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"gemini-2.0-flash:generateContent?key={api_key}"
    )

    joined_text = "\n".join(
        [f"{i+1}. {txt}" for i, txt in enumerate(korean_list)]
    )

    prompt = f"""
다음 한국어 안전 문구들을 건설현장 TBM 교육자료용으로 자연스럽고 짧게 번역하라.

반드시 아래 형식의 JSON 배열만 출력하라.
설명 금지.
코드블록 금지.
배열 개수는 입력 개수와 정확히 같아야 한다.

입력:
{joined_text}

출력형식:
[
  {{
    "zh": "중국어 번역",
    "vi": "베트남어 번역",
    "my": "미얀마어 번역"
  }}
]
"""

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ]
    }

    last_error = None

    for attempt in range(1):
        try:
            resp = requests.post(
                url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=90
            )
            resp.raise_for_status()

            data = resp.json()
            text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
            text = text.replace("```json", "").replace("```", "").strip()

            parsed = json.loads(text)

            if not isinstance(parsed, list):
                raise ValueError("Gemini 응답이 JSON 배열 형식이 아닙니다.")

            if len(parsed) != len(korean_list):
                raise ValueError(
                    f"번역 결과 개수 불일치: 입력 {len(korean_list)}개 / 출력 {len(parsed)}개"
                )

            for item in parsed:
                if not all(k in item for k in ("zh", "vi", "my")):
                    raise ValueError("번역 결과에 zh/vi/my 키가 없습니다.")

            return parsed

        except requests.HTTPError as e:
            last_error = e
            status_code = e.response.status_code if e.response is not None else None

            if status_code == 429 and attempt < 2:
                time.sleep(5 * (attempt + 1))
                continue
            raise

        except Exception as e:
            last_error = e
            if attempt < 2:
                time.sleep(3)
                continue
            raise

    raise last_error


def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes):
                yield sub_shape


def has_text(shape) -> bool:
    return hasattr(shape, "has_text_frame") and shape.has_text_frame


def get_text(shape) -> str:
    if not has_text(shape):
        return ""
    return shape.text.strip()


def clear_and_set_text(shape, text: str, font_size: int):
    if not has_text(shape):
        return

    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)


def find_picture_area(slide):
    for shape in iter_all_shapes(slide.shapes):
        txt = get_text(shape)
        if "사진대지" in txt:
            return shape

    candidates = []
    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            try:
                area = shape.width * shape.height
                candidates.append((shape.top, -area, shape))
            except Exception:
                pass

    if not candidates:
        raise ValueError("사진 영역을 찾지 못했습니다.")

    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[0][2]


def insert_image_cover(slide, image_path, area_shape):
    left = area_shape.left
    top = area_shape.top
    width = area_shape.width
    height = area_shape.height

    if has_text(area_shape):
        area_shape.text_frame.clear()

    with Image.open(image_path) as img:
        img_w, img_h = img.size

    box_ratio = width / height
    img_ratio = img_w / img_h

    if img_ratio > box_ratio:
        new_height = height
        new_width = int(height * img_ratio)
    else:
        new_width = width
        new_height = int(width / img_ratio)

    pic_left = left + int((width - new_width) / 2)
    pic_top = top + int((height - new_height) / 2)

    slide.shapes.add_picture(
        image_path,
        pic_left,
        pic_top,
        width=new_width,
        height=new_height,
    )


def find_text_shapes_for_content(slide, picture_area):
    candidates = []
    pic_bottom = picture_area.top + picture_area.height

    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            if shape.top >= pic_bottom - Pt(5):
                candidates.append(shape)

    candidates.sort(key=lambda s: (s.top, s.left))
    return candidates


def rewrite_slide_texts(slide, picture_area, item: SlideData):
    text_shapes = find_text_shapes_for_content(slide, picture_area)

    if len(text_shapes) < 5:
        raise ValueError(f"설명 영역 텍스트 상자 5개 필요 / 현재: {len(text_shapes)}개")

    target_shapes = text_shapes[:5]

    new_texts = [
        TITLE_TEXT,
        item.ko,
        item.zh,
        item.vi,
        item.my,
    ]

    font_sizes = [
        TITLE_FONT_SIZE_PT,
        BASE_FONT_SIZE_PT,
        BASE_FONT_SIZE_PT,
        BASE_FONT_SIZE_PT,
        BASE_FONT_SIZE_PT,
    ]

    for shape, text, size in zip(target_shapes, new_texts, font_sizes):
        clear_and_set_text(shape, text, size)


def delete_slide(prs, index: int):
    slide_id_list = prs.slides._sldIdLst
    slide = slide_id_list[index]
    r_id = slide.rId
    prs.part.drop_rel(r_id)
    del slide_id_list[index]


def remove_unused_slides(prs: Presentation, used_count: int):
    total_count = len(prs.slides)
    for idx in range(total_count - 1, used_count - 1, -1):
        delete_slide(prs, idx)


def fill_slide(slide, item: SlideData):
    picture_area = find_picture_area(slide)
    insert_image_cover(slide, item.image_path, picture_area)
    rewrite_slide_texts(slide, picture_area, item)


def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    prs = Presentation(TEMPLATE_PPT)

    if len(prs.slides) < len(slide_data_list):
        raise ValueError(
            f"템플릿 슬라이드 수({len(prs.slides)})보다 사진 수({len(slide_data_list)})가 많습니다."
        )

    for i, item in enumerate(slide_data_list):
        fill_slide(prs.slides[i], item)

    remove_unused_slides(prs, len(slide_data_list))

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output


def main():
    st.set_page_config(page_title="TBM PPT 생성기", layout="wide")
    st.title("TBM 조회자료 PPT 생성기")

    if not os.path.exists(TEMPLATE_PPT):
        st.error("templates/sample_template.pptx 파일을 찾을 수 없습니다.")
        st.stop()

    try:
        gemini_api_key = st.secrets["GEMINI_API_KEY"]
    except Exception as e:
        st.error(f"Streamlit Secrets 설정 오류: {e}")
        st.stop()

    uploaded_files = st.file_uploader(
        "사진 업로드",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True
    )

    if not uploaded_files:
        st.info("사진을 업로드하세요.")
        st.stop()

    slide_data_list: List[SlideData] = []
    temp_paths = []

    st.subheader("사진별 한국어 문구 입력")

    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        st.markdown(f"### 슬라이드 {idx}")

        col1, col2 = st.columns([1, 2])

        with col1:
            st.image(
    uploaded_file,
    caption=uploaded_file.name,
    width=180
)

        with col2:
            ko = st.text_input(
                f"한국어 #{idx}",
                value="지정된 이동통로 통행",
                key=f"ko_{idx}"
            )

        suffix = os.path.splitext(uploaded_file.name)[1].lower() or ".jpg"

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(uploaded_file.getbuffer())
            temp_path = tmp.name
            temp_paths.append(temp_path)

        slide_data_list.append(
            SlideData(
                image_path=temp_path,
                ko=ko,
                zh="",
                vi="",
                my="",
            )
        )

    if st.button("PPT 생성"):
        try:
            with st.spinner("번역 및 PPT 생성 중..."):
                ko_list = [item.ko for item in slide_data_list]
                translations = translate_batch_with_gemini(gemini_api_key, ko_list)

                translated_list: List[SlideData] = []
                for item, tr in zip(slide_data_list, translations):
                    translated_list.append(
                        SlideData(
                            image_path=item.image_path,
                            ko=item.ko,
                            zh=tr["zh"],
                            vi=tr["vi"],
                            my=tr["my"],
                        )
                    )

                ppt_data = build_ppt(translated_list)

            st.success(f"PPT 생성 완료 ({len(translated_list)}장)")
            st.download_button(
                label="완성 PPT 다운로드",
                data=ppt_data,
                file_name=OUTPUT_PPT_NAME,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        except requests.HTTPError as e:
            status_code = e.response.status_code if e.response is not None else None
            if status_code == 429:
                st.error("번역 API 요청 한도를 초과했습니다. 잠시 후 다시 시도하세요.")
            else:
                st.error(f"HTTP 오류 발생: {e}")

        except Exception as e:
            st.error(f"오류 발생: {e}")

        finally:
            for path in temp_paths:
                try:
                    if os.path.exists(path):
                        os.remove(path)
                except Exception:
                    pass


if __name__ == "__main__":
    main()