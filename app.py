import os
import io
import re
import json
import time
import tempfile
from dataclasses import dataclass
from typing import List
from datetime import datetime

import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
DAILY_TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template2.pptx")

PUBLIC_DRIVE_DIR = os.path.join(BASE_DIR, "public_drive")

BASE_FONT_SIZE_PT = 35
OUTPUT_PPT_NAME = "TBM_완성본.pptx"
DAILY_OUTPUT_PPT_NAME = "일일안전회의_완성본.pptx"
APP_VERSION = "26년 5월 버전"

PHOTO_BOX_TEXT = "PHOTO_BOX"
KO_BOX_TEXT = "1"
ZH_BOX_TEXT = "2"
VI_BOX_TEXT = "3"
MY_BOX_TEXT = "4"
DATE_BOX_TEXT = "DATE_BOX"
WEATHER_BOX_1_TEXT = "WEATHER_BOX_1"
WEATHER_BOX_2_TEXT = "WEATHER_BOX_2"

PUBLIC_DRIVE_LIMIT_MB = 100
PUBLIC_DRIVE_LIMIT_BYTES = PUBLIC_DRIVE_LIMIT_MB * 1024 * 1024
PUBLIC_DRIVE_EXPIRE_SECONDS = 24 * 60 * 60

NAVER_WEATHER_URL = "https://weather.naver.com/"

OSAN_YANGSAN_LAT = 37.196790422777
OSAN_YANGSAN_LON = 127.02460549856

BROWSER_VIEWPORT = {
    "width": 1280,
    "height": 1600,
}

# 여기 좌표는 직접 수정하면 됨
WEATHER_CAPTURE_1 = {
    "scroll_y": 0,
    "clip": {
        "x": 0,
        "y": 120,
        "width": 1280,
        "height": 550,
    }
}

WEATHER_CAPTURE_2 = {
    "scroll_y": 900,
    "clip": {
        "x": 0,
        "y": 100,
        "width": 1280,
        "height": 650,
    }
}


@dataclass
class SlideData:
    image_path: str
    ko: str
    zh: str
    vi: str
    my: str


def hide_streamlit_ui():
    st.markdown(
        """
        <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .stDeployButton {display: none;}
        [data-testid="stToolbar"] {display: none;}
        [data-testid="stDecoration"] {display: none;}
        [data-testid="stStatusWidget"] {display: none;}
        [data-testid="manage-app-button"] {display: none;}
        </style>
        """,
        unsafe_allow_html=True
    )


def safe_filename(filename: str) -> str:
    name = os.path.basename(filename)
    name = re.sub(r"[^a-zA-Z0-9가-힣._ -]", "_", name)
    return name[:120]


def format_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes}B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f}KB"
    return f"{size_bytes / 1024 / 1024:.1f}MB"


def ensure_public_drive():
    os.makedirs(PUBLIC_DRIVE_DIR, exist_ok=True)


def cleanup_old_drive_files():
    ensure_public_drive()
    now = time.time()

    for name in os.listdir(PUBLIC_DRIVE_DIR):
        path = os.path.join(PUBLIC_DRIVE_DIR, name)
        if os.path.isfile(path):
            if now - os.path.getmtime(path) > PUBLIC_DRIVE_EXPIRE_SECONDS:
                try:
                    os.remove(path)
                except Exception:
                    pass


def get_drive_size() -> int:
    ensure_public_drive()
    total = 0

    for name in os.listdir(PUBLIC_DRIVE_DIR):
        path = os.path.join(PUBLIC_DRIVE_DIR, name)
        if os.path.isfile(path):
            total += os.path.getsize(path)

    return total


def save_drive_file(uploaded_file):
    ensure_public_drive()

    current_size = get_drive_size()
    file_bytes = uploaded_file.getvalue()
    file_size = len(file_bytes)

    if current_size + file_size > PUBLIC_DRIVE_LIMIT_BYTES:
        raise ValueError(
            f"부적합 사진 용량 초과: 현재 {format_size(current_size)} / "
            f"추가 {format_size(file_size)} / 최대 {PUBLIC_DRIVE_LIMIT_MB}MB"
        )

    filename = safe_filename(uploaded_file.name)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    save_name = f"{timestamp}_{filename}"
    save_path = os.path.join(PUBLIC_DRIVE_DIR, save_name)

    with open(save_path, "wb") as f:
        f.write(file_bytes)

    return save_path


def make_thumbnail_bytes(path: str, max_size: int = 180):
    try:
        img = Image.open(path)

        try:
            img.seek(0)
        except Exception:
            pass

        if img.mode != "RGB":
            img = img.convert("RGB")

        img.thumbnail((max_size, max_size))

        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        buf.seek(0)
        return buf

    except Exception:
        return None


def delete_drive_file(path: str):
    try:
        if os.path.exists(path) and os.path.isfile(path):
            os.remove(path)
            return True
    except Exception:
        pass
    return False


def render_public_drive():
    cleanup_old_drive_files()

    with st.expander("부적합 사진", expanded=False):
        used = get_drive_size()
        st.caption(f"용량 {format_size(used)} / {PUBLIC_DRIVE_LIMIT_MB}MB")

        drive_uploads = st.file_uploader(
            "부적합 등록",
            accept_multiple_files=True,
            type=["jpg", "jpeg", "png", "webp", "heic", "heif", "mpo"],
            key="public_drive_uploader"
        )

        if drive_uploads:
            uploaded_count = 0

            for file in drive_uploads:
                try:
                    save_drive_file(file)
                    uploaded_count += 1
                except Exception as e:
                    st.error(str(e))

            if uploaded_count > 0:
                st.success(f"{uploaded_count}개 등록 완료")
                st.rerun()

        files = []
        ensure_public_drive()

        for name in sorted(os.listdir(PUBLIC_DRIVE_DIR), reverse=True):
            path = os.path.join(PUBLIC_DRIVE_DIR, name)
            if os.path.isfile(path):
                files.append((name, path, os.path.getsize(path)))

        if files:
            st.markdown("#### 등록 목록")

            for idx, (name, path, size) in enumerate(files, start=1):
                col1, col2, col3 = st.columns([1, 2.2, 0.55])

                thumb = make_thumbnail_bytes(path)

                with col1:
                    if thumb:
                        st.image(thumb, width=110)
                    else:
                        st.write("파일")

                with col2:
                    st.caption(name)
                    st.caption(format_size(size))

                    with open(path, "rb") as f:
                        st.download_button(
                            label="다운로드",
                            data=f,
                            file_name=name,
                            use_container_width=True,
                            key=f"download_{name}_{idx}"
                        )

                with col3:
                    if st.button("X", key=f"delete_{name}_{idx}", use_container_width=True):
                        if delete_drive_file(path):
                            st.rerun()
                        else:
                            st.error("삭제 실패")

        else:
            st.info("부적합 사진 없음.")


def convert_to_jpg(input_path: str, max_size: int = 1600, quality: int = 88) -> str:
    try:
        img = Image.open(input_path)

        try:
            img.seek(0)
        except Exception:
            pass

        if img.mode != "RGB":
            img = img.convert("RGB")

        width, height = img.size
        longest = max(width, height)

        if longest > max_size:
            ratio = max_size / longest
            new_width = int(width * ratio)
            new_height = int(height * ratio)
            img = img.resize((new_width, new_height), Image.LANCZOS)

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        output_path = tmp.name
        tmp.close()

        img.save(
            output_path,
            format="JPEG",
            quality=quality,
            optimize=True
        )

        return output_path

    except Exception as e:
        raise ValueError(f"이미지 변환 실패: {e}")


def get_korean_date_text():
    weekdays = ["월요일", "화요일", "수요일", "목요일", "금요일", "토요일", "일요일"]
    now = datetime.now()
    return f"{now.year}년 {now.month:02d}월 {now.day:02d}일 {weekdays[now.weekday()]}"


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

    try:
        parsed = json.loads(text)
    except Exception:
        text = re.sub(r'[\x00-\x1F]+', ' ', text)
        parsed = json.loads(text)

    if not isinstance(parsed, list):
        raise ValueError("GPT 응답이 배열이 아닙니다.")

    if len(parsed) != len(korean_list):
        raise ValueError(
            f"번역 개수 불일치: 입력 {len(korean_list)} / 출력 {len(parsed)}"
        )

    for item in parsed:
        if not all(k in item for k in ("zh", "vi", "my")):
            raise ValueError("번역 결과에 zh, vi, my 키가 없습니다.")

    return parsed


def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub_shape in iter_all_shapes(shape.shapes):
                yield sub_shape


def has_text(shape):
    return hasattr(shape, "has_text_frame") and shape.has_text_frame


def normalize_text(text: str) -> str:
    return str(text).strip().replace("\n", "").replace("\r", "")


def find_text_target(slide, target_text: str):
    target = normalize_text(target_text)

    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            if normalize_text(shape.text) == target:
                return ("shape", shape)

    for shape in iter_all_shapes(slide.shapes):
        if hasattr(shape, "has_table") and shape.has_table:
            table = shape.table
            for row in table.rows:
                for cell in row.cells:
                    if normalize_text(cell.text) == target:
                        return ("cell", cell)

    return None


def set_target_text(target_obj, text: str, size_pt: int, font_name: str = None):
    kind, obj = target_obj

    tf = obj.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)

    if font_name:
        run.font.name = font_name


def add_picture_to_shape(slide, image_path, target_shape):
    slide.shapes.add_picture(
        image_path,
        target_shape.left,
        target_shape.top,
        width=target_shape.width,
        height=target_shape.height
    )


def fill_slide_by_placeholders(slide, item: SlideData):
    photo_target = find_text_target(slide, PHOTO_BOX_TEXT)
    ko_target = find_text_target(slide, KO_BOX_TEXT)
    zh_target = find_text_target(slide, ZH_BOX_TEXT)
    vi_target = find_text_target(slide, VI_BOX_TEXT)
    my_target = find_text_target(slide, MY_BOX_TEXT)

    missing = []
    for name, obj in [
        ("PHOTO_BOX", photo_target),
        ("1", ko_target),
        ("2", zh_target),
        ("3", vi_target),
        ("4", my_target),
    ]:
        if obj is None:
            missing.append(name)

    if missing:
        raise ValueError(f"슬라이드에서 플레이스홀더를 찾지 못했습니다: {', '.join(missing)}")

    photo_kind, photo_obj = photo_target
    if photo_kind != "shape":
        raise ValueError("PHOTO_BOX는 텍스트 상자/도형이어야 합니다.")

    add_picture_to_shape(slide, item.image_path, photo_obj)

    set_target_text(ko_target, item.ko, BASE_FONT_SIZE_PT)
    set_target_text(zh_target, item.zh, BASE_FONT_SIZE_PT)
    set_target_text(vi_target, item.vi, BASE_FONT_SIZE_PT)
    set_target_text(my_target, item.my, BASE_FONT_SIZE_PT)


def insert_image_to_placeholder(slide, placeholder_text: str, image_path: str):
    target = find_text_target(slide, placeholder_text)

    if target is None:
        raise ValueError(f"{placeholder_text} 플레이스홀더를 찾지 못했습니다.")

    kind, obj = target

    if kind != "shape":
        raise ValueError(f"{placeholder_text}는 도형/텍스트박스여야 합니다.")

    add_picture_to_shape(slide, image_path, obj)


def fill_date_box(slide):
    target = find_text_target(slide, DATE_BOX_TEXT)

    if target:
        set_target_text(
            target,
            get_korean_date_text(),
            30,
            font_name="맑은 고딕"
        )


def capture_naver_weather_region():
    if sync_playwright is None:
        raise ValueError("playwright가 설치되지 않았습니다.")

    weather_1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
    weather_2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        context = browser.new_context(
            viewport=BROWSER_VIEWPORT,
            locale="ko-KR",
            geolocation={
                "latitude": OSAN_YANGSAN_LAT,
                "longitude": OSAN_YANGSAN_LON,
            },
            permissions=["geolocation"],
        )

        page = context.new_page()
        page.goto(NAVER_WEATHER_URL, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(2500)

        page.evaluate(f"window.scrollTo(0, {WEATHER_CAPTURE_1['scroll_y']})")
        page.wait_for_timeout(1000)
        page.screenshot(
            path=weather_1,
            clip=WEATHER_CAPTURE_1["clip"]
        )

        page.evaluate(f"window.scrollTo(0, {WEATHER_CAPTURE_2['scroll_y']})")
        page.wait_for_timeout(1000)
        page.screenshot(
            path=weather_2,
            clip=WEATHER_CAPTURE_2["clip"]
        )

        browser.close()

    return weather_1, weather_2


def build_ppt_from_template(
    slide_data_list: List[SlideData],
    template_path: str,
    include_daily_options: bool = False
) -> io.BytesIO:
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")

    prs = Presentation(template_path)
    temp_extra_paths = []

    if include_daily_options:
        if len(prs.slides) >= 1:
            fill_date_box(prs.slides[0])

        try:
            weather_1, weather_2 = capture_naver_weather_region()
            temp_extra_paths.extend([weather_1, weather_2])

            if len(prs.slides) >= 2:
                insert_image_to_placeholder(prs.slides[1], WEATHER_BOX_1_TEXT, weather_1)

            if len(prs.slides) >= 3:
                insert_image_to_placeholder(prs.slides[2], WEATHER_BOX_2_TEXT, weather_2)

        except Exception as e:
            raise ValueError(f"날씨 캡쳐 실패: {e}")

    start_slide_index = 3 if include_daily_options else 0

    for i, item in enumerate(slide_data_list):
        target_index = start_slide_index + i

        if target_index >= len(prs.slides):
            break

        slide = prs.slides[target_index]
        fill_slide_by_placeholders(slide, item)

    keep_slide_count = start_slide_index + len(slide_data_list)

    if include_daily_options:
        keep_slide_count = max(keep_slide_count, 3)

    for idx in range(len(prs.slides) - 1, keep_slide_count - 1, -1):
        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)

    for p in temp_extra_paths:
        if os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass

    return out


def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    return build_ppt_from_template(
        slide_data_list,
        TEMPLATE_PPT,
        include_daily_options=False
    )


def build_daily_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    return build_ppt_from_template(
        slide_data_list,
        DAILY_TEMPLATE_PPT,
        include_daily_options=True
    )


def render_slide_input_area(
    uploader_label: str,
    button_label: str,
    output_name: str,
    build_func,
    uploader_key: str,
    button_key: str,
    download_key: str
):
    files = st.file_uploader(
        uploader_label,
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg", "webp", "heic", "heif", "mpo"],
        key=uploader_key
    )

    if files:
        slide_inputs = []
        temp_paths = []

        for idx, f in enumerate(files):
            with st.expander(f"슬라이드 #{idx+1}", expanded=True):
                c1, c2 = st.columns([1, 4])

                suffix = os.path.splitext(f.name)[1].lower() or ".jpg"

                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(f.getbuffer())
                    original_path = tmp.name
                    temp_paths.append(original_path)

                jpg_path = convert_to_jpg(original_path)
                temp_paths.append(jpg_path)

                c1.image(jpg_path, width=150)

                ko_input = c2.text_input(
                    "한국어 문구",
                    value="",
                    placeholder="예: 지정된 이동통로 통행",
                    key=f"{uploader_key}_ko_{idx}"
                )

                slide_inputs.append(SlideData(jpg_path, ko_input, "", "", ""))

        if st.button(button_label, key=button_key):
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
                    ppt = build_func(slide_inputs)

                st.success("완료!")
                st.download_button(
                    "PPT 다운로드",
                    ppt,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key=download_key
                )

            except Exception as e:
                st.error(f"오류 발생: {e}")

            finally:
                for p in temp_paths:
                    if os.path.exists(p):
                        try:
                            os.remove(p)
                        except Exception:
                            pass


def render_daily_safety_meeting():
    with st.expander("일일안전회의", expanded=False):
        render_slide_input_area(
            uploader_label="사진 업로드",
            button_label="일일안전회의 PPT 생성",
            output_name=DAILY_OUTPUT_PPT_NAME,
            build_func=build_daily_ppt,
            uploader_key="daily_meeting_uploader",
            button_key="daily_create_btn",
            download_key="daily_download_btn"
        )


def main():
    st.set_page_config(page_title="TBM PPT Maker", layout="wide")
    hide_streamlit_ui()

    top_left, top_right = st.columns([3, 1])
    with top_left:
        st.title(f"🚧 TBM 교육자료 자동 번역 생성기 [{APP_VERSION}]")
    with top_right:
        render_public_drive()
        render_daily_safety_meeting()

    if "GPT_API_KEY" not in st.secrets:
        st.warning("Secrets에 GPT_API_KEY 설정 필요")
        st.stop()

    render_slide_input_area(
        uploader_label="사진 업로드",
        button_label="PPT 생성",
        output_name=OUTPUT_PPT_NAME,
        build_func=build_ppt,
        uploader_key="main_tbm_uploader",
        button_key="main_create_btn",
        download_key="main_download_btn"
    )


if __name__ == "__main__":
    main()