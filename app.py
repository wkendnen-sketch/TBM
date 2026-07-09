import os
import io
import re
import json
import time
import zipfile
import gc
import tempfile
import subprocess
from copy import deepcopy
from dataclasses import dataclass
from typing import List, Tuple, Optional
from datetime import datetime
from difflib import SequenceMatcher

import requests
import streamlit as st
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Pt
from pptx.dml.color import RGBColor

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

try:
    from playwright.sync_api import sync_playwright
except Exception:
    sync_playwright = None

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    from paddleocr import PaddleOCR
except Exception:
    PaddleOCR = None

_PADDLE_OCR = None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template.pptx")
DAILY_TEMPLATE_PPT = os.path.join(BASE_DIR, "templates", "sample_template2.pptx")
BAD_PHOTO_DIR = os.path.join(BASE_DIR, "bad_photo_upload")
SHARED_NOTICE_FILE = os.path.join(BASE_DIR, "shared_notice.md")
SHARED_NOTICE_META_FILE = os.path.join(BASE_DIR, "shared_notice_meta.json")

BASE_FONT_SIZE_PT = 35
OUTPUT_PPT_NAME = "TBM_완성본.pptx"
DAILY_OUTPUT_PPT_NAME = "일일안전회의_완성본.pptx"
APP_VERSION = "26년 5월 버전"

# 대량 업로드/고용량 사진 안정화 설정
TBM_IMAGE_MAX_SIZE = 1200
TBM_IMAGE_QUALITY = 82
DAILY_IMAGE_MAX_SIZE = 1400
DAILY_IMAGE_QUALITY = 84
MAX_TBM_FILES_SOFT_WARN = 35
MAX_DAILY_FILES_SOFT_WARN = 35
TRANSLATION_BATCH_SIZE = 15

PHOTO_BOX_TEXT = "PHOTO_BOX"
KO_BOX_TEXT = "1"
ZH_BOX_TEXT = "2"
VI_BOX_TEXT = "3"
MY_BOX_TEXT = "4"

DATE_BOX_TEXT = "DATE_BOX"
WEATHER_BOX_1_TEXT = "WEATHER_BOX_1"
WEATHER_BOX_2_TEXT = "WEATHER_BOX_2"

DAILY_PHOTO_BOX_TEXT = "PHOTO_BOX_1"
DAILY_TEXT_BOX_TEXT = "TEXT_BOX_1"
TIME_BOX_TEXT = "TIME_BOX_1"
HOLD_POINT_TEXT = "HOLD POINT"

BAD_PHOTO_LIMIT_MB = 300
BAD_PHOTO_LIMIT_BYTES = BAD_PHOTO_LIMIT_MB * 1024 * 1024
BAD_PHOTO_EXPIRE_SECONDS = 24 * 60 * 60

NAVER_WEATHER_URL = "https://weather.naver.com/today/02370550"

BROWSER_VIEWPORT = {
    "width": 1920,
    "height": 2400,
}

WEATHER_CAPTURE_1 = {
    "scroll_y": 0,
    "clip": {
        "x": 350,
        "y": 200,
        "width": 780,
        "height": 428,
    }
}

WEATHER_CAPTURE_2 = {
    "scroll_y": 300,
    "clip": {
        "x": 350,
        "y": 399,
        "width": 780,
        "height": 290,
    }
}


@dataclass
class SlideData:
    image_path: str
    ko: str = ""
    zh: str = ""
    vi: str = ""
    my: str = ""
    orientation: str = ""


@dataclass
class DailySlideData:
    image_path: str
    text: str = ""
    orientation: str = ""


@dataclass
class MaterialWorkItem:
    image_path: str
    original_name: str = ""
    upload_index: int = 0
    ocr_text: str = ""
    work_type: str = "material"
    company: str = "기타업체"
    number: int = 0


COMPANY_ORDER = [
    "원영건업",
    "청암기업",
    "유셀네트웍스",
    "엠케이지",
    "KEC",
    "우신에이스",
    "진솔",
    "장한건설",
]

COMPANY_ALIAS = {
    "원영건업": ["원영건업", "원영"],
    "청암기업": ["청암기업", "청암"],
    "유셀네트웍스": ["유셀네트웍스", "유셀네트윅스", "유셀네트", "유셀"],
    "엠케이지": ["엠케이지", "MKG", "mkg"],
    "KEC": ["KEC", "kec", "케이이씨", "케이씨", "케이"],
    "우신에이스": ["우신에이스", "우신"],
    "진솔": ["진솔"],
    "장한건설": ["장한건설", "장한"],
}

HIGH_RISK_KEYWORDS = [
    "25대 고위험",
    "25대고위험",
    "25 대 고위험",
    "25대",
    "고위험",
    "고 위험",
]


def install_playwright_browser():
    try:
        subprocess.run(
            ["playwright", "install", "chromium"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except Exception:
        pass


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

        .block-container {
            padding-top: 1.2rem !important;
        }

        hr {
            margin-top: 0.35rem !important;
            margin-bottom: 0.35rem !important;
        }

        h2, h3, h4 {
            margin-top: 0.35rem !important;
            margin-bottom: 0.35rem !important;
        }

        div[data-testid="stMarkdownContainer"] p {
            margin-bottom: 0.25rem !important;
        }

        div[data-testid="stFileUploader"] {
            margin-top: -0.2rem !important;
            margin-bottom: 0.35rem !important;
        }

        div[data-testid="stVerticalBlock"] {
            gap: 0.35rem !important;
        }

        .bad-photo-small-button button {
            min-height: 34px !important;
            padding: 0.25rem 0.4rem !important;
            font-size: 12px !important;
        }
        </style>
        """,
        unsafe_allow_html=True
    )


def render_app_title():
    st.markdown(
        f"""
        <div style="margin-top:-30px; margin-bottom:8px;">
            <h2 style="
                font-size:24px;
                font-weight:700;
                margin:0;
                padding:0;
                line-height:1.2;
            ">
                🚧 TBM 교육자료 자동 번역 생성기 [{APP_VERSION}]
            </h2>
        </div>
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


def ensure_bad_photo_dir():
    os.makedirs(BAD_PHOTO_DIR, exist_ok=True)


def cleanup_old_bad_photo_files():
    ensure_bad_photo_dir()
    now = time.time()

    for name in os.listdir(BAD_PHOTO_DIR):
        path = os.path.join(BAD_PHOTO_DIR, name)
        if os.path.isfile(path):
            if now - os.path.getmtime(path) > BAD_PHOTO_EXPIRE_SECONDS:
                try:
                    os.remove(path)
                except Exception:
                    pass


def get_bad_photo_size() -> int:
    ensure_bad_photo_dir()
    total = 0

    for name in os.listdir(BAD_PHOTO_DIR):
        path = os.path.join(BAD_PHOTO_DIR, name)
        if os.path.isfile(path):
            total += os.path.getsize(path)

    return total


def save_bad_photo_file(uploaded_file):
    ensure_bad_photo_dir()

    file_bytes = uploaded_file.getvalue()
    file_size = len(file_bytes)
    original_filename = safe_filename(uploaded_file.name)

    for existing_name in os.listdir(BAD_PHOTO_DIR):
        if existing_name.endswith(original_filename):
            existing_path = os.path.join(BAD_PHOTO_DIR, existing_name)
            if os.path.isfile(existing_path) and os.path.getsize(existing_path) == file_size:
                return existing_path

    current_size = get_bad_photo_size()

    if current_size + file_size > BAD_PHOTO_LIMIT_BYTES:
        raise ValueError(
            f"부적합사진 용량 초과: 현재 {format_size(current_size)} / "
            f"추가 {format_size(file_size)} / 최대 {BAD_PHOTO_LIMIT_MB}MB"
        )

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    save_name = f"{timestamp}_{original_filename}"
    save_path = os.path.join(BAD_PHOTO_DIR, save_name)

    with open(save_path, "wb") as f:
        f.write(file_bytes)

    return save_path


def save_generated_ppt_to_bad_photo_storage(ppt_bytes: io.BytesIO, filename: str):
    ensure_bad_photo_dir()

    ppt_bytes.seek(0)
    data = ppt_bytes.getvalue()
    file_size = len(data)

    current_size = get_bad_photo_size()

    if current_size + file_size > BAD_PHOTO_LIMIT_BYTES:
        ppt_bytes.seek(0)
        return False

    safe_name = safe_filename(filename)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    save_name = f"{timestamp}_{safe_name}"
    save_path = os.path.join(BAD_PHOTO_DIR, save_name)

    with open(save_path, "wb") as f:
        f.write(data)

    ppt_bytes.seek(0)
    return True


def make_bad_photo_zip():
    ensure_bad_photo_dir()

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for name in sorted(os.listdir(BAD_PHOTO_DIR)):
            path = os.path.join(BAD_PHOTO_DIR, name)
            if os.path.isfile(path):
                zip_file.write(path, arcname=name)

    zip_buffer.seek(0)
    return zip_buffer


def delete_all_bad_photo_files():
    ensure_bad_photo_dir()
    deleted_count = 0

    for name in os.listdir(BAD_PHOTO_DIR):
        path = os.path.join(BAD_PHOTO_DIR, name)
        if os.path.isfile(path):
            try:
                os.remove(path)
                deleted_count += 1
            except Exception:
                pass

    return deleted_count


def make_thumbnail_bytes(path: str, max_size: int = 180):
    try:
        img = Image.open(path)

        try:
            img.seek(0)
        except Exception:
            pass

        try:
            img = ImageOps.exif_transpose(img)
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


def delete_file(path: str):
    try:
        if os.path.exists(path) and os.path.isfile(path):
            os.remove(path)
            return True
    except Exception:
        pass
    return False


def render_bad_photo_storage():
    cleanup_old_bad_photo_files()

    files = []
    ensure_bad_photo_dir()

    for name in sorted(os.listdir(BAD_PHOTO_DIR), reverse=True):
        path = os.path.join(BAD_PHOTO_DIR, name)
        if os.path.isfile(path):
            files.append((name, path, os.path.getsize(path)))

    col_title, col_zip, col_delete = st.columns([5.2, 0.9, 0.9])

    with col_title:
        used = get_bad_photo_size()
        st.markdown(
            f"""
            <div style="
                display:flex;
                align-items:center;
                gap:12px;
                margin:0;
                padding:0;
                line-height:1.1;
                min-height:34px;
            ">
                <span style="font-size:48px; font-weight:800; color:#006400;">부적합사진</span>
                <span style="font-size:13px; color:#666;">
                    용량 {format_size(used)} / {BAD_PHOTO_LIMIT_MB}MB
                </span>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col_zip:
        st.markdown("<div class='bad-photo-small-button'>", unsafe_allow_html=True)
        if files:
            st.download_button(
                "ZIP",
                data=make_bad_photo_zip(),
                file_name="부적합사진_전체.zip",
                mime="application/zip",
                use_container_width=True,
                key="bad_photo_download_all_zip"
            )
        else:
            st.button("ZIP", disabled=True, use_container_width=True, key="bad_photo_download_all_zip_disabled")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_delete:
        st.markdown("<div class='bad-photo-small-button'>", unsafe_allow_html=True)
        if st.button("전체삭제", use_container_width=True, key="bad_photo_delete_all"):
            delete_all_bad_photo_files()
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    upload_files = st.file_uploader(
        "부적합사진 파일 업로드",
        accept_multiple_files=True,
        type=None,  # 확장자 제한 없음: 사진, 영상, MP3, 압축파일, 문서 등 업로드 가능
        key="daily_bad_photo_upload_uploader",
        help="확장자 제한 없음: 사진, 영상, MP3, 압축파일, 문서 등 거의 모든 파일 업로드 가능"
    )

    if upload_files:
        uploaded_count = 0

        for file in upload_files:
            try:
                save_bad_photo_file(file)
                uploaded_count += 1
            except Exception as e:
                st.error(str(e))

        if uploaded_count > 0:
            st.success(f"{uploaded_count}개 등록 완료")
            st.rerun()

    if files:
        with st.expander(f"파일 목록 보기 ({len(files)}개)", expanded=False):
            for idx, (name, path, size) in enumerate(files, start=1):
                col1, col2, col3 = st.columns([1, 3, 0.5])

                thumb = make_thumbnail_bytes(path)

                with col1:
                    if thumb:
                        st.image(thumb, width=110)
                    else:
                        st.write("파일")
                        st.caption("미리보기 불가")

                with col2:
                    st.caption(name)
                    st.caption(format_size(size))

                    with open(path, "rb") as f:
                        st.download_button(
                            label="다운로드",
                            data=f,
                            file_name=name,
                            use_container_width=True,
                            key=f"bad_photo_download_{name}_{idx}"
                        )

                with col3:
                    if st.button("X", key=f"bad_photo_delete_{name}_{idx}", use_container_width=True):
                        if delete_file(path):
                            st.rerun()
                        else:
                            st.error("삭제 실패")
    else:
        st.info("부적합사진 파일 없음.")


def get_image_orientation(path: str) -> str:
    """EXIF 방향 보정 후 실제 이미지 방향을 반환."""
    try:
        img = Image.open(path)
        try:
            img.seek(0)
        except Exception:
            pass
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass
        w, h = img.size
        if w > h:
            return "landscape"
        if h > w:
            return "portrait"
        return "square"
    except Exception:
        return "unknown"


def convert_to_jpg(
    input_path: str,
    max_size: int = TBM_IMAGE_MAX_SIZE,
    quality: int = TBM_IMAGE_QUALITY
) -> str:
    """
    PPT 삽입용 JPG 변환.
    - EXIF Orientation을 실제 픽셀 방향으로 반영한다.
    - 가로/세로 방향을 임의 변경하지 않는다.
    - 긴 변 기준으로 축소해 대량 사진/고용량 사진 생성 실패를 줄인다.
    """
    try:
        img = Image.open(input_path)

        try:
            img.seek(0)
        except Exception:
            pass

        original_orientation = get_image_orientation(input_path)

        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass

        if img.mode not in ("RGB", "L"):
            # 투명 PNG/WEBP는 흰 배경으로 합성해서 PPT 호환성 확보
            if "A" in img.getbands():
                bg = Image.new("RGB", img.size, (255, 255, 255))
                bg.paste(img, mask=img.getchannel("A"))
                img = bg
            else:
                img = img.convert("RGB")
        elif img.mode == "L":
            img = img.convert("RGB")

        width, height = img.size
        longest = max(width, height)

        if longest > max_size:
            ratio = max_size / longest
            new_width = max(1, int(width * ratio))
            new_height = max(1, int(height * ratio))
            img = img.resize((new_width, new_height), Image.LANCZOS)

        converted_orientation = "landscape" if img.width > img.height else "portrait" if img.height > img.width else "square"
        if original_orientation in ("landscape", "portrait") and converted_orientation != original_orientation:
            raise ValueError("이미지 변환 중 가로/세로 방향이 변경되었습니다.")

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        output_path = tmp.name
        tmp.close()

        img.save(output_path, format="JPEG", quality=quality, optimize=True, progressive=True)
        img.close()
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


def translate_all_with_gpt(api_key: str, korean_list: List[str]):
    """사진 수가 많을 때 API 응답 길이/시간초과를 줄이기 위해 나눠 번역."""
    results = []
    for start_idx in range(0, len(korean_list), TRANSLATION_BATCH_SIZE):
        chunk = korean_list[start_idx:start_idx + TRANSLATION_BATCH_SIZE]
        results.extend(translate_batch_with_gpt(api_key, chunk))
    return results


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


def normalize_for_match(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).upper()


def get_paddle_ocr():
    """PaddleOCR 지연 로딩. 여러 버전 호환을 위해 옵션을 단계적으로 시도."""
    global _PADDLE_OCR

    if PaddleOCR is None:
        return None

    if _PADDLE_OCR is not None:
        return _PADDLE_OCR

    option_list = [
        dict(lang="korean", use_angle_cls=True, show_log=False, det_db_box_thresh=0.25, det_db_unclip_ratio=2.0),
        dict(lang="korean", use_angle_cls=True, show_log=False),
        dict(lang="korean", use_angle_cls=True),
        dict(lang="korean"),
    ]

    for opts in option_list:
        try:
            _PADDLE_OCR = PaddleOCR(**opts)
            return _PADDLE_OCR
        except Exception:
            continue

    _PADDLE_OCR = None
    return None


def normalize_ocr_text(text: str) -> str:
    text = str(text or "")
    text = text.replace("|", "I").replace("﹣", "-").replace("–", "-").replace("—", "-")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def resize_for_ocr(img: Image.Image, min_longest: int = 2600, max_longest: int = 4600) -> Image.Image:
    """OCR용 크기 보정. 작으면 키우고, 너무 크면 메모리 보호 수준에서만 줄임."""
    width, height = img.size
    longest = max(width, height)

    if longest < min_longest:
        ratio = min_longest / max(1, longest)
        img = img.resize((int(width * ratio), int(height * ratio)), Image.LANCZOS)
    elif longest > max_longest:
        ratio = max_longest / longest
        img = img.resize((int(width * ratio), int(height * ratio)), Image.LANCZOS)

    return img


def make_ocr_preprocess_variants(img: Image.Image) -> List[Image.Image]:
    """
    현장 문서용 강화 OCR 전처리.
    같은 영역에서 여러 이미지 버전을 만들어 OCR 결과를 합친다.
    느려지지만 작은 글씨/그림자/흐림 대응력이 올라간다.
    """
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass

    if img.mode != "RGB":
        img = img.convert("RGB")

    img = resize_for_ocr(img)

    variants = []

    # 1) 원본 RGB 확대본
    variants.append(img)

    gray = img.convert("L")
    gray = ImageOps.autocontrast(gray)

    # 2) 기본 흑백 대비 강화
    v1 = ImageEnhance.Contrast(gray).enhance(1.8)
    v1 = ImageEnhance.Sharpness(v1).enhance(2.2)
    v1 = v1.filter(ImageFilter.SHARPEN)
    variants.append(v1)

    # 3) 더 강한 대비/샤프닝
    v2 = ImageEnhance.Contrast(gray).enhance(2.6)
    v2 = ImageEnhance.Sharpness(v2).enhance(3.0)
    v2 = v2.filter(ImageFilter.SHARPEN)
    variants.append(v2)

    # 4) 밝기 보정 + 대비
    v3 = ImageEnhance.Brightness(gray).enhance(1.12)
    v3 = ImageOps.autocontrast(v3)
    v3 = ImageEnhance.Contrast(v3).enhance(2.1)
    variants.append(v3)

    # 5) 이진화 1 - 일반 문서
    try:
        v4 = gray.point(lambda p: 255 if p > 165 else 0)
        variants.append(v4)
    except Exception:
        pass

    # 6) 이진화 2 - 어두운 사진/그림자 대응
    try:
        v5 = gray.point(lambda p: 255 if p > 135 else 0)
        variants.append(v5)
    except Exception:
        pass

    # 7) 가장 강한 확대 + 샤프닝
    try:
        w, h = img.size
        if max(w, h) < 4200:
            v6 = img.resize((int(w * 1.35), int(h * 1.35)), Image.LANCZOS).convert("L")
            v6 = ImageOps.autocontrast(v6)
            v6 = ImageEnhance.Contrast(v6).enhance(2.4)
            v6 = ImageEnhance.Sharpness(v6).enhance(3.5)
            variants.append(v6)
    except Exception:
        pass

    return variants


def ocr_with_paddle(img: Image.Image) -> str:
    ocr = get_paddle_ocr()
    if ocr is None:
        return ""

    temp_path = None
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_path = tmp.name
        tmp.close()

        save_img = img.convert("RGB") if img.mode != "RGB" else img
        save_img.save(temp_path, format="PNG")

        try:
            result = ocr.ocr(temp_path, cls=True)
        except TypeError:
            result = ocr.ocr(temp_path)

        lines = []
        if result:
            for page in result:
                if not page:
                    continue
                for line in page:
                    try:
                        text = line[1][0]
                        score = 1.0
                        try:
                            score = float(line[1][1])
                        except Exception:
                            pass
                        if text and score >= 0.25:
                            lines.append(str(text))
                    except Exception:
                        continue

        return normalize_ocr_text("\n".join(lines))

    except Exception:
        return ""

    finally:
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def ocr_with_tesseract(img: Image.Image) -> str:
    if pytesseract is None:
        return ""

    configs = [
        "--oem 3 --psm 6",
        "--oem 3 --psm 11",
        "--oem 3 --psm 12",
        "--oem 3 --psm 4",
    ]
    langs = ["kor+eng", "kor", "eng", None]

    best_text = ""
    for lang in langs:
        for config in configs:
            try:
                if lang:
                    text = pytesseract.image_to_string(img, lang=lang, config=config)
                else:
                    text = pytesseract.image_to_string(img, config=config)
                text = normalize_ocr_text(text)
                if len(normalize_for_match(text)) > len(normalize_for_match(best_text)):
                    best_text = text
            except Exception:
                continue

    return normalize_ocr_text(best_text)


def get_ocr_regions(img: Image.Image, top_ratio: float = 0.65) -> List[Tuple[str, Image.Image]]:
    """전체/상단/좌측/중앙 등 여러 영역을 읽어 업체명·25대고위험 누락을 줄임."""
    width, height = img.size
    top_h = max(1, int(height * top_ratio))
    mid_y1 = int(height * 0.18)
    mid_y2 = int(height * 0.78)

    regions = [
        ("full", img),
        ("top", img.crop((0, 0, width, top_h))),
        ("top_left", img.crop((0, 0, int(width * 0.72), top_h))),
        ("top_right", img.crop((int(width * 0.28), 0, width, top_h))),
        ("left", img.crop((0, 0, int(width * 0.72), height))),
        ("center", img.crop((int(width * 0.10), mid_y1, int(width * 0.90), mid_y2))),
    ]
    return regions


def append_unique_text(texts: List[str], seen: set, text: str):
    text = normalize_ocr_text(text)
    key = normalize_for_match(text)
    if text and key and key not in seen:
        texts.append(text)
        seen.add(key)


def extract_ocr_text(image_path: str, top_ratio: float = 0.65) -> str:
    """
    정확도 우선 강화 OCR.
    - 원본 이미지 기준 OCR
    - 전체/상단/좌측/중앙 영역 OCR
    - 영역별 전처리 5~7종 OCR
    - PaddleOCR 우선 + Tesseract 보조
    - 결과를 합쳐 분류 판단에 사용
    """
    try:
        img = Image.open(image_path)
        try:
            img.seek(0)
        except Exception:
            pass

        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass

        if img.mode != "RGB":
            img = img.convert("RGB")

        all_texts = []
        seen = set()

        paddle_available = get_paddle_ocr() is not None

        for region_name, region in get_ocr_regions(img, top_ratio=top_ratio):
            variants = make_ocr_preprocess_variants(region)

            for variant_idx, variant in enumerate(variants):
                # PaddleOCR은 정확도가 좋으므로 모든 주요 전처리 버전에 실행
                if paddle_available:
                    append_unique_text(all_texts, seen, ocr_with_paddle(variant))

                # Tesseract는 느리고 중복이 많으므로 원본/강화/이진화 일부만 보조 실행
                if variant_idx in (1, 2, 4):
                    append_unique_text(all_texts, seen, ocr_with_tesseract(variant))

        return normalize_ocr_text("\n".join(all_texts))

    except Exception:
        return ""


def extract_top_ocr_text(image_path: str, top_ratio: float = 0.65) -> str:
    """기존 함수명 호환용. 내부는 강화 OCR 사용."""
    return extract_ocr_text(image_path, top_ratio=top_ratio)


def fuzzy_contains(text: str, candidates: List[str], threshold: float = 0.72) -> bool:
    compact = normalize_for_match(text)
    if not compact:
        return False

    for cand in candidates:
        cand_norm = normalize_for_match(cand)
        if not cand_norm:
            continue
        if cand_norm in compact:
            return True

        # 긴 OCR 문자열 안에서 후보 길이만큼 잘라 유사도 검사
        n = len(cand_norm)
        if n <= 1:
            continue
        for i in range(0, max(1, len(compact) - n + 1)):
            part = compact[i:i + n]
            if SequenceMatcher(None, cand_norm, part).ratio() >= threshold:
                return True

    return False


def detect_high_risk(text: str) -> bool:
    compact = normalize_for_match(text)
    loose = str(text or "")

    if "고위험" in compact or "고 위험" in loose:
        return True
    if "25대" in compact or "25 대" in loose:
        return True

    # OCR에서 숫자/한글이 일부 깨지는 경우 보정
    high_risk_candidates = HIGH_RISK_KEYWORDS + [
        "이십오대고위험",
        "25고위험",
        "25대위험",
        "고위헙",
        "고위힘",
        "고위혐",
        "고위험작업",
        "고위험 작업",
    ]

    if fuzzy_contains(compact, high_risk_candidates, threshold=0.70):
        return True

    # 25와 위험류 단어가 따로 읽힌 경우
    has_25 = bool(re.search(r"2\s*5|25|이십오", compact))
    has_risk = fuzzy_contains(compact, ["고위험", "위험", "위헙", "위혐"], threshold=0.68)
    return has_25 and has_risk


def detect_company(text: str) -> str:
    compact = normalize_for_match(text)

    # 자주 틀리는 OCR 후보 추가
    extra_alias = {
        "원영건업": ["원영건업", "원영", "원영건", "원명건업"],
        "청암기업": ["청암기업", "청암", "청암기엽", "청암업"],
        "유셀네트웍스": ["유셀네트웍스", "유셀네트윅스", "유셀네트", "유셀", "유셀네트워크", "유셀네트웍"],
        "엠케이지": ["엠케이지", "MKG", "mkg", "엠케이", "엠케", "MK G"],
        "KEC": ["KEC", "kec", "케이이씨", "케이씨", "케이", "K E C"],
        "우신에이스": ["우신에이스", "우신", "우신에이", "우신이스"],
        "진솔": ["진솔", "진술", "진슬", "진솔건", "진솔건설"],
        "장한건설": ["장한건설", "장한", "장한전설", "장한건", "장한건썰"],
    }

    for company in COMPANY_ORDER:
        aliases = extra_alias.get(company, COMPANY_ALIAS.get(company, [company]))
        for alias in aliases:
            if normalize_for_match(alias) in compact:
                return company

    # OCR 오인식 보정: 토큰 및 슬라이딩 유사도 기반 판단
    tokens = re.findall(r"[가-힣A-Za-z0-9]{2,}", str(text or ""))
    compact_tokens = [normalize_for_match(t) for t in tokens]
    compact_tokens.append(compact)

    best_company = "기타업체"
    best_score = 0.0

    for company in COMPANY_ORDER:
        aliases = extra_alias.get(company, COMPANY_ALIAS.get(company, [company]))
        for alias in aliases:
            alias_norm = normalize_for_match(alias)
            if not alias_norm:
                continue
            for token in compact_tokens:
                if len(token) < 2:
                    continue

                # 토큰 전체 비교
                score = SequenceMatcher(None, alias_norm, token).ratio()
                if score > best_score:
                    best_score = score
                    best_company = company

                # 긴 토큰 안에 업체명이 섞여 있는 경우 부분 비교
                n = len(alias_norm)
                if len(token) >= n >= 2:
                    for i in range(0, len(token) - n + 1):
                        part = token[i:i + n]
                        score = SequenceMatcher(None, alias_norm, part).ratio()
                        if score > best_score:
                            best_score = score
                            best_company = company

    if best_score >= 0.68:
        return best_company

    return "기타업체"


def extract_sort_number(text: str) -> int:
    raw = str(text or "")

    # 엠케이지 1. 지게차 / 진솔-2 / KEC 지하 2층 등 대부분 대응
    patterns = [
        r"(?:지하|B)\s*[-]?\s*(\d+)\s*층",
        r"[-–_]\s*(\d+)",
        r"\b(\d+)\s*[.)]",
        r"(?:^|\s)(\d+)(?:\s|$)",
    ]

    for pattern in patterns:
        m = re.search(pattern, raw, flags=re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                pass

    return 0


def classify_material_work_image(
    image_path: str,
    original_name: str,
    upload_index: int,
    ocr_image_path: str = None
) -> MaterialWorkItem:
    # OCR은 원본 이미지로, PPT 삽입은 변환/축소된 JPG로 분리
    ocr_source = ocr_image_path or image_path
    ocr_text = extract_ocr_text(ocr_source, top_ratio=0.60)
    combined_text = f"{ocr_text} {original_name}"

    work_type = "high_risk" if detect_high_risk(combined_text) else "material"
    company = detect_company(combined_text)
    number = extract_sort_number(combined_text)

    return MaterialWorkItem(
        image_path=image_path,
        original_name=original_name,
        upload_index=upload_index,
        ocr_text=ocr_text,
        work_type=work_type,
        company=company,
        number=number,
    )


def company_order_index(company: str) -> int:
    try:
        return COMPANY_ORDER.index(company)
    except ValueError:
        return 999


def sort_material_work_items(items: List[MaterialWorkItem]) -> List[MaterialWorkItem]:
    # 자재입고현황 전체 → 25대 고위험작업 전체
    # 각 그룹 내부는 업체순 → 숫자순 → 업로드순
    type_order = {"material": 0, "high_risk": 1}

    return sorted(
        items,
        key=lambda x: (
            type_order.get(x.work_type, 99),
            company_order_index(x.company),
            x.number,
            x.upload_index,
        )
    )


def slide_has_text(slide, target_text: str) -> bool:
    target = normalize_text(target_text)

    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            if target in normalize_text(shape.text):
                return True

    for shape in iter_all_shapes(slide.shapes):
        if hasattr(shape, "has_table") and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if target in normalize_text(cell.text):
                        return True

    return False


def find_text_target(slide, target_text: str):
    """
    PPT 플레이스홀더 탐색 강화 버전.
    1) 도형 안 텍스트가 target_text와 정확히 일치하면 반환
    2) 도형 이름(shape.name)이 target_text와 정확히 일치하면 반환
    3) 그룹도형 내부 도형까지 탐색
    4) 표 셀도 탐색

    PowerPoint에서 텍스트박스를 복사/수정하면
    화면에는 1,2,3,4가 보여도 shape.name은 TextBox 7처럼 바뀔 수 있어서
    텍스트와 도형 이름을 둘 다 확인한다.
    """
    target = normalize_text(target_text)
    target_match = normalize_for_match(target_text)

    # 1차: 텍스트 정확 일치
    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            if normalize_text(shape.text) == target:
                return ("shape", shape)

    # 2차: 도형 이름 정확 일치
    for shape in iter_all_shapes(slide.shapes):
        shape_name = normalize_text(getattr(shape, "name", ""))
        if shape_name == target:
            return ("shape", shape)

    # 3차: 공백/줄바꿈 제거 후 텍스트 일치
    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            if normalize_for_match(shape.text) == target_match:
                return ("shape", shape)

    # 4차: 공백/줄바꿈 제거 후 도형 이름 일치
    for shape in iter_all_shapes(slide.shapes):
        shape_name = getattr(shape, "name", "")
        if normalize_for_match(shape_name) == target_match:
            return ("shape", shape)

    # 5차: 표 셀 탐색
    for shape in iter_all_shapes(slide.shapes):
        if hasattr(shape, "has_table") and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if normalize_text(cell.text) == target:
                        return ("cell", cell)
                    if normalize_for_match(cell.text) == target_match:
                        return ("cell", cell)

    return None


def debug_slide_placeholders(slide):
    """Streamlit 화면에서 PPT 도형 이름/텍스트 확인용."""
    rows = []
    for idx, shape in enumerate(iter_all_shapes(slide.shapes), start=1):
        rows.append({
            "순번": idx,
            "도형이름": getattr(shape, "name", ""),
            "텍스트": normalize_text(shape.text) if has_text(shape) else "",
            "타입": str(getattr(shape, "shape_type", "")),
        })
    return rows



def find_tbm_translation_cells(slide):
    """TBM 템플릿 표 구조용 안전장치.
    sample_template.pptx는 1,2,3,4가 표 셀에 들어있으므로,
    텍스트 탐색이 실패해도 표의 우측 2~5행 셀을 직접 사용한다.
    """
    for shape in iter_all_shapes(slide.shapes):
        try:
            if hasattr(shape, "has_table") and shape.has_table:
                table = shape.table
                if len(table.rows) >= 6 and len(table.columns) >= 2:
                    # 기본 템플릿 구조: 2~5행, 우측 열이 1~4 번역칸
                    return {
                        "1": ("cell", table.cell(2, 1)),
                        "2": ("cell", table.cell(3, 1)),
                        "3": ("cell", table.cell(4, 1)),
                        "4": ("cell", table.cell(5, 1)),
                    }
        except Exception:
            continue
    return {}


def find_slide_index_by_text(prs, target_text: str):
    for idx, slide in enumerate(prs.slides):
        if slide_has_text(slide, target_text):
            return idx
    return None


def set_target_text(
    target_obj,
    text: str,
    size_pt: int,
    font_name: str = None,
    bold: bool = False,
    font_color=None
):
    kind, obj = target_obj

    tf = obj.text_frame
    tf.clear()

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold

    if font_name:
        run.font.name = font_name

    if font_color:
        run.font.color.rgb = font_color


def add_picture_to_shape(slide, image_path, target_shape):
    """
    PHOTO_BOX에 사진을 넣을 때 회전/왜곡 금지.
    - 사진 비율 유지
    - 박스는 꽉 채움(cover crop)
    - crop만 적용하고 rotation은 항상 0
    """
    try:
        img = Image.open(image_path)
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass
        img_w, img_h = img.size
        img.close()
    except Exception:
        slide.shapes.add_picture(
            image_path,
            target_shape.left,
            target_shape.top,
            width=target_shape.width,
            height=target_shape.height
        )
        return

    if img_w <= 0 or img_h <= 0:
        return

    box_w = float(target_shape.width)
    box_h = float(target_shape.height)
    img_ratio = img_w / img_h
    box_ratio = box_w / box_h

    pic = slide.shapes.add_picture(
        image_path,
        target_shape.left,
        target_shape.top,
        width=target_shape.width,
        height=target_shape.height
    )

    try:
        pic.rotation = 0
    except Exception:
        pass

    # python-pptx crop 값은 0~1 비율. 가로/세로를 바꾸지 않고 초과분만 잘라낸다.
    try:
        if img_ratio > box_ratio:
            crop = max(0.0, min(0.49, (1 - (box_ratio / img_ratio)) / 2))
            pic.crop_left = crop
            pic.crop_right = crop
            pic.crop_top = 0
            pic.crop_bottom = 0
        elif img_ratio < box_ratio:
            crop = max(0.0, min(0.49, (1 - (img_ratio / box_ratio)) / 2))
            pic.crop_top = crop
            pic.crop_bottom = crop
            pic.crop_left = 0
            pic.crop_right = 0
    except Exception:
        pass

    return pic


def duplicate_slide(prs, source_slide):
    """
    슬라이드 복제 시 도형 XML뿐 아니라 이미지 relationship도 같이 복사.
    기존 deepcopy만 쓰면 국기 같은 그림이 2번째 슬라이드부터 깨질 수 있다.
    """
    blank_slide_layout = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank_slide_layout)

    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(
            new_el,
            "p:extLst"
        )

    # 이미지/차트 등 외부 관계 복사. notesSlide는 복사하지 않음.
    try:
        for rel in source_slide.part.rels.values():
            if "notesSlide" in rel.reltype:
                continue
            try:
                new_slide.part.rels.add_relationship(rel.reltype, rel._target, rel.rId)
            except Exception:
                try:
                    new_slide.part.rels.add_relationship(rel.reltype, rel.target_part, rel.rId)
                except Exception:
                    pass
    except Exception:
        pass

    return new_slide


def get_slide_index_by_id(prs, slide_id: int):
    for idx, slide in enumerate(prs.slides):
        if slide.slide_id == slide_id:
            return idx
    return None


def delete_slide_by_index(prs, idx: int):
    slide_id = prs.slides._sldIdLst[idx]
    prs.part.drop_rel(slide_id.rId)
    del prs.slides._sldIdLst[idx]


def delete_slide_by_id(prs, slide_id: int):
    idx = get_slide_index_by_id(prs, slide_id)
    if idx is not None:
        delete_slide_by_index(prs, idx)


def move_slide_by_id(prs, slide_id: int, target_index: int):
    idx = get_slide_index_by_id(prs, slide_id)
    if idx is None:
        return

    sld_id = prs.slides._sldIdLst[idx]
    prs.slides._sldIdLst.remove(sld_id)

    target_index = max(0, min(target_index, len(prs.slides._sldIdLst)))
    prs.slides._sldIdLst.insert(target_index, sld_id)


def find_slide_index_by_any_text(prs, texts: List[str]):
    for text in texts:
        idx = find_slide_index_by_text(prs, text)
        if idx is not None:
            return idx
    return None


def clone_and_fill_bad_slides(prs, template_slide, bad_items: List[DailySlideData]) -> List[int]:
    created_ids = []
    for item in bad_items:
        new_slide = duplicate_slide(prs, template_slide)
        fill_daily_slide(new_slide, item, strict=False)
        created_ids.append(new_slide.slide_id)
    return created_ids


def clone_and_fill_material_slides(prs, template_slide, material_items: List[MaterialWorkItem]) -> List[int]:
    created_ids = []
    for item in material_items:
        new_slide = duplicate_slide(prs, template_slide)
        insert_image_to_placeholder(new_slide, TIME_BOX_TEXT, item.image_path)
        created_ids.append(new_slide.slide_id)
    return created_ids


def fill_slide_by_placeholders(slide, item: SlideData, strict: bool = True):
    photo_target = find_text_target(slide, PHOTO_BOX_TEXT)
    ko_target = find_text_target(slide, KO_BOX_TEXT)
    zh_target = find_text_target(slide, ZH_BOX_TEXT)
    vi_target = find_text_target(slide, VI_BOX_TEXT)
    my_target = find_text_target(slide, MY_BOX_TEXT)

    # 안전장치: sample_template.pptx의 1,2,3,4가 표 셀로 들어간 경우 직접 매칭
    table_cells = find_tbm_translation_cells(slide)
    if ko_target is None:
        ko_target = table_cells.get("1")
    if zh_target is None:
        zh_target = table_cells.get("2")
    if vi_target is None:
        vi_target = table_cells.get("3")
    if my_target is None:
        my_target = table_cells.get("4")

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
        if strict:
            debug = []
            for shape in iter_all_shapes(slide.shapes):
                try:
                    debug.append(f"name=[{getattr(shape, 'name', '')}] text=[{getattr(shape, 'text', '')}]")
                    if hasattr(shape, "has_table") and shape.has_table:
                        for r, row in enumerate(shape.table.rows):
                            for c, cell in enumerate(row.cells):
                                debug.append(f"table[{r},{c}]=[{cell.text}]")
                except Exception:
                    pass
            raise ValueError(
                "슬라이드에서 플레이스홀더를 찾지 못했습니다: "
                + ", ".join(missing)
                + "\n\n"
                + "\n".join(debug[:80])
            )
        return

    photo_kind, photo_obj = photo_target
    if photo_kind != "shape":
        if strict:
            raise ValueError("PHOTO_BOX는 텍스트 상자/도형이어야 합니다.")
        return

    add_picture_to_shape(slide, item.image_path, photo_obj)

    set_target_text(ko_target, item.ko, BASE_FONT_SIZE_PT)
    set_target_text(zh_target, item.zh, BASE_FONT_SIZE_PT)
    set_target_text(vi_target, item.vi, BASE_FONT_SIZE_PT)
    set_target_text(my_target, item.my, BASE_FONT_SIZE_PT)


def fill_daily_slide(slide, item: DailySlideData, strict: bool = False):
    photo_target = find_text_target(slide, DAILY_PHOTO_BOX_TEXT)
    text_target = find_text_target(slide, DAILY_TEXT_BOX_TEXT)

    if photo_target:
        kind, obj = photo_target
        if kind == "shape":
            add_picture_to_shape(slide, item.image_path, obj)
    elif strict:
        raise ValueError("PHOTO_BOX_1 플레이스홀더를 찾지 못했습니다.")

    if text_target:
        set_target_text(
            text_target,
            item.text,
            28,
            font_name="맑은 고딕",
            bold=True,
            font_color=RGBColor(0, 0, 0)
        )
    elif strict:
        raise ValueError("TEXT_BOX_1 플레이스홀더를 찾지 못했습니다.")


def insert_image_to_placeholder(slide, placeholder_text: str, image_path: str):
    target = find_text_target(slide, placeholder_text)

    if target is None:
        return

    kind, obj = target

    if kind != "shape":
        return

    add_picture_to_shape(slide, image_path, obj)


def fill_material_slides(prs, material_items: List[MaterialWorkItem]):
    if not material_items:
        return

    base_idx = find_slide_index_by_text(prs, TIME_BOX_TEXT)

    if base_idx is None:
        return

    base_slide = prs.slides[base_idx]

    for i, item in enumerate(material_items):
        if i == 0:
            target_slide = base_slide
        else:
            target_slide = duplicate_slide(prs, base_slide)

        insert_image_to_placeholder(
            target_slide,
            TIME_BOX_TEXT,
            item.image_path
        )


def fill_date_box(slide):
    target = find_text_target(slide, DATE_BOX_TEXT)

    if target:
        set_target_text(
            target,
            get_korean_date_text(),
            30,
            font_name="맑은 고딕",
            bold=True,
            font_color=RGBColor(0, 0, 0)
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
            device_scale_factor=2,
        )

        page = context.new_page()

        page.goto(
            NAVER_WEATHER_URL,
            wait_until="networkidle",
            timeout=60000
        )

        page.wait_for_timeout(3000)

        page.evaluate(f"window.scrollTo(0, {WEATHER_CAPTURE_1['scroll_y']})")
        page.wait_for_timeout(1000)
        page.screenshot(path=weather_1, clip=WEATHER_CAPTURE_1["clip"])

        page.evaluate(f"window.scrollTo(0, {WEATHER_CAPTURE_2['scroll_y']})")
        page.wait_for_timeout(1000)
        page.screenshot(path=weather_2, clip=WEATHER_CAPTURE_2["clip"])

        browser.close()

    return weather_1, weather_2


def delete_extra_slides(prs, keep_slide_count: int):
    for idx in range(len(prs.slides) - 1, keep_slide_count - 1, -1):
        slide = prs.slides[idx]

        if slide_has_text(slide, HOLD_POINT_TEXT):
            continue

        slide_id = prs.slides._sldIdLst[idx]
        prs.part.drop_rel(slide_id.rId)
        del prs.slides._sldIdLst[idx]


def build_ppt(slide_data_list: List[SlideData]) -> io.BytesIO:
    if not os.path.exists(TEMPLATE_PPT):
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {TEMPLATE_PPT}")

    prs = Presentation(TEMPLATE_PPT)

    if not slide_data_list:
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        return out

    if len(prs.slides) < 1:
        raise ValueError("TBM 템플릿에 기준 슬라이드가 없습니다.")

    # 국기 이미지 깨짐 방지: 템플릿에 이미 들어있는 슬라이드를 먼저 사용한다.
    # 템플릿 수를 초과한 사진만 relationship 복사 방식으로 추가 복제한다.
    filled_ids = []
    template_slide_count = len(prs.slides)
    base_slide = prs.slides[0]

    for i, item in enumerate(slide_data_list):
        if i < template_slide_count:
            target_slide = prs.slides[i]
        else:
            target_slide = duplicate_slide(prs, base_slide)

        fill_slide_by_placeholders(target_slide, item, strict=True)
        filled_ids.append(target_slide.slide_id)

        if i % 10 == 0:
            gc.collect()

    keep_ids = set(filled_ids)
    for idx in range(len(prs.slides) - 1, -1, -1):
        slide = prs.slides[idx]
        if slide.slide_id in keep_ids:
            continue
        if slide_has_text(slide, HOLD_POINT_TEXT):
            continue
        delete_slide_by_index(prs, idx)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    gc.collect()
    return out


def build_daily_ppt(
    bad_items: List[DailySlideData],
    material_items: List[MaterialWorkItem]
) -> io.BytesIO:
    if not os.path.exists(DAILY_TEMPLATE_PPT):
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {DAILY_TEMPLATE_PPT}")

    prs = Presentation(DAILY_TEMPLATE_PPT)
    temp_extra_paths = []

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

    # 템플릿 기준 슬라이드 찾기
    bad_template_idx = find_slide_index_by_text(prs, DAILY_PHOTO_BOX_TEXT)
    material_template_idx = find_slide_index_by_text(prs, TIME_BOX_TEXT)

    if bad_items and bad_template_idx is None:
        raise ValueError("부적합사진 기준 슬라이드(PHOTO_BOX_1)를 찾지 못했습니다.")

    if material_items and material_template_idx is None:
        raise ValueError("자재입고/고위험 기준 슬라이드(TIME_BOX_1)를 찾지 못했습니다.")

    bad_template_slide = prs.slides[bad_template_idx] if bad_template_idx is not None else None
    material_template_slide = prs.slides[material_template_idx] if material_template_idx is not None else None

    bad_template_id = bad_template_slide.slide_id if bad_template_slide is not None else None
    material_template_id = material_template_slide.slide_id if material_template_slide is not None else None

    created_bad_ids = []
    created_material_ids = []

    if bad_template_slide is not None:
        created_bad_ids = clone_and_fill_bad_slides(prs, bad_template_slide, bad_items)

    if material_template_slide is not None:
        created_material_ids = clone_and_fill_material_slides(prs, material_template_slide, material_items)

    # 원본 기준 슬라이드는 빈 템플릿이므로 제거. 같은 슬라이드 중복 제거 방지.
    for sid in sorted({x for x in [bad_template_id, material_template_id] if x is not None}, reverse=True):
        delete_slide_by_id(prs, sid)

    # 동적 슬라이드를 명일 작업내용 발표 앞에 배치.
    # 명일 문구를 못 찾으면 HOLD POINT 앞, 그것도 못 찾으면 맨 뒤에 배치.
    anchor_idx = find_slide_index_by_any_text(
        prs,
        [
            "명일 작업내용 발표",
            "명일작업내용발표",
            "명일 작업내용",
            "명일작업내용",
        ]
    )

    if anchor_idx is None:
        anchor_idx = find_slide_index_by_text(prs, HOLD_POINT_TEXT)

    if anchor_idx is None:
        anchor_idx = len(prs.slides)

    desired_dynamic_ids = created_bad_ids + created_material_ids

    insert_at = anchor_idx
    for sid in desired_dynamic_ids:
        move_slide_by_id(prs, sid, insert_at)
        insert_at += 1

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


def render_tbm_input_area():
    files = st.file_uploader(
        "사진 업로드",
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg", "webp", "heic", "heif", "mpo"],
        key="main_tbm_uploader"
    )

    if files:
        if len(files) > MAX_TBM_FILES_SOFT_WARN:
            st.warning(f"사진이 {len(files)}장입니다. 고용량 사진은 자동 축소해서 처리하지만 생성 시간이 길어질 수 있습니다.")

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

                jpg_path = convert_to_jpg(original_path, max_size=TBM_IMAGE_MAX_SIZE, quality=TBM_IMAGE_QUALITY)
                temp_paths.append(jpg_path)

                c1.image(jpg_path, width=150)

                ko_input = c2.text_input(
                    "한국어 문구",
                    value="",
                    placeholder="예: 지정된 이동통로 통행",
                    key=f"main_tbm_ko_{idx}"
                )

                slide_inputs.append(SlideData(jpg_path, ko_input, "", "", "", get_image_orientation(jpg_path)))

        if st.button("PPT 생성", key="main_create_btn"):
            try:
                if "GPT_API_KEY" not in st.secrets:
                    raise ValueError("Secrets에 GPT_API_KEY 설정 필요")

                with st.spinner("번역 중..."):
                    ko_list = [s.ko for s in slide_inputs]

                    if any(not x.strip() for x in ko_list):
                        raise ValueError("빈 한국어 문구가 있습니다. 모든 슬라이드 문구를 입력하세요.")

                    translations = translate_all_with_gpt(
                        st.secrets["GPT_API_KEY"],
                        ko_list
                    )

                    for s, tr in zip(slide_inputs, translations):
                        s.zh = tr["zh"]
                        s.vi = tr["vi"]
                        s.my = tr["my"]

                with st.spinner("PPT 생성 중..."):
                    ppt = build_ppt(slide_inputs)

                save_generated_ppt_to_bad_photo_storage(ppt, OUTPUT_PPT_NAME)

                st.success("완료!")
                st.download_button(
                    "PPT 다운로드",
                    ppt,
                    file_name=OUTPUT_PPT_NAME,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="main_download_btn"
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
    st.markdown(
        """
        <h2 style="color:#d00000; font-weight:800; font-size:48px; margin-top:0.35rem; margin-bottom:0.35rem;">
            일일안전회의
        </h2>
        """,
        unsafe_allow_html=True
    )

    bad_files = st.file_uploader(
        "부적합사진",
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg", "webp", "heic", "heif", "mpo"],
        key="daily_bad_uploader"
    )

    bad_items = []
    material_items = []
    temp_paths = []

    if bad_files:
        if len(bad_files) > MAX_DAILY_FILES_SOFT_WARN:
            st.warning(f"부적합사진이 {len(bad_files)}장입니다. 자동 축소 처리하지만 생성 시간이 길어질 수 있습니다.")
        st.markdown("#### 부적합사진")

        for idx, f in enumerate(bad_files):
            suffix = os.path.splitext(f.name)[1].lower() or ".jpg"

            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(f.getbuffer())
                original_path = tmp.name
                temp_paths.append(original_path)

            jpg_path = convert_to_jpg(original_path, max_size=DAILY_IMAGE_MAX_SIZE, quality=DAILY_IMAGE_QUALITY)
            temp_paths.append(jpg_path)

            with st.expander(f"부적합사진 #{idx + 1}", expanded=True):
                c1, c2 = st.columns([1, 4])

                with c1:
                    st.image(jpg_path, width=130)

                with c2:
                    text_value = st.text_input(
                        "문구 입력",
                        value="",
                        placeholder="예: 자재 반입 확인",
                        key=f"daily_bad_text_{idx}"
                    )

            bad_items.append(DailySlideData(jpg_path, text_value, get_image_orientation(jpg_path)))

    material_files = st.file_uploader(
        "자재입고 및 고위험작업",
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg", "webp", "heic", "heif", "mpo"],
        key="daily_material_uploader"
    )

    if material_files:
        if len(material_files) > MAX_DAILY_FILES_SOFT_WARN:
            st.warning(f"자재입고 및 고위험작업 사진이 {len(material_files)}장입니다. OCR 때문에 시간이 오래 걸릴 수 있습니다.")
        st.markdown("#### 자재입고 및 고위험작업")

        for idx, f in enumerate(material_files):
            suffix = os.path.splitext(f.name)[1].lower() or ".jpg"

            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(f.getbuffer())
                material_original_path = tmp.name
                temp_paths.append(material_original_path)

            material_jpg_path = convert_to_jpg(material_original_path, max_size=DAILY_IMAGE_MAX_SIZE, quality=DAILY_IMAGE_QUALITY)
            temp_paths.append(material_jpg_path)

            item = classify_material_work_image(
                material_jpg_path,
                original_name=f.name,
                upload_index=idx,
                ocr_image_path=material_original_path
            )
            material_items.append(item)

            c1, c2 = st.columns([1, 4])
            with c1:
                st.image(material_jpg_path, width=130)
            with c2:
                kind_label = "25대 고위험작업" if item.work_type == "high_risk" else "자재입고현황"
                st.caption(f"{idx + 1}번 / {kind_label} / {item.company} / 번호 {item.number}")
                st.caption(f.name)
                # OCR 원문/실패 문구는 화면에 표시하지 않음.

    if material_items:
        sorted_preview = sort_material_work_items(material_items)
        with st.expander("자재입고 및 고위험작업 정렬 결과", expanded=False):
            for order_idx, item in enumerate(sorted_preview, start=1):
                kind_label = "25대 고위험작업" if item.work_type == "high_risk" else "자재입고현황"
                st.caption(
                    f"{order_idx}. {kind_label} / {item.company} / 번호 {item.number} / {item.original_name}"
                )

    if st.button("일일안전회의 PPT 생성", key="daily_create_btn"):
        try:
            with st.spinner("PPT 생성 중..."):
                sorted_material_items = sort_material_work_items(material_items)
                ppt = build_daily_ppt(bad_items, sorted_material_items)

            save_generated_ppt_to_bad_photo_storage(ppt, DAILY_OUTPUT_PPT_NAME)

            st.success("완료!")
            st.download_button(
                "일일안전회의 PPT 다운로드",
                ppt,
                file_name=DAILY_OUTPUT_PPT_NAME,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="daily_download_btn"
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

    st.markdown("---")
    render_bad_photo_storage()
    st.markdown("---")


def load_shared_notice() -> str:
    try:
        if os.path.exists(SHARED_NOTICE_FILE):
            with open(SHARED_NOTICE_FILE, "r", encoding="utf-8") as f:
                return f.read()
    except Exception:
        pass
    return ""


def save_shared_notice(text: str):
    try:
        with open(SHARED_NOTICE_FILE, "w", encoding="utf-8") as f:
            f.write(text or "")

        meta = {
            "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(SHARED_NOTICE_META_FILE, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        return True
    except Exception:
        return False


def clear_shared_notice():
    ok = save_shared_notice("")
    return ok


def load_shared_notice_saved_at() -> str:
    try:
        if os.path.exists(SHARED_NOTICE_META_FILE):
            with open(SHARED_NOTICE_META_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return str(data.get("saved_at", "") or "")
    except Exception:
        pass
    return ""


def render_shared_notice_board():
    st.markdown("---")
    st.markdown(
        """
        <h2 style="font-size:1.35em; font-weight:700; margin-top:0.35rem; margin-bottom:0.35rem;">
            공지사항 / 메모
        </h2>
        """,
        unsafe_allow_html=True
    )

    current_text = load_shared_notice()

    notice_text = st.text_area(
        "",
        value=current_text,
        height=260,
        key="shared_notice_text_area",
        placeholder="공지사항/메모/안전사항 등."
    )

    col_save, col_delete, col_space = st.columns([0.8, 0.8, 5.4])

    with col_save:
        if st.button("저장", use_container_width=True, key="shared_notice_save_btn"):
            if save_shared_notice(notice_text):
                st.success("저장 완료")
                st.rerun()
            else:
                st.error("저장 실패")

    with col_delete:
        if st.button("삭제", use_container_width=True, key="shared_notice_delete_btn"):
            if clear_shared_notice():
                st.success("삭제 완료")
                st.rerun()
            else:
                st.error("삭제 실패")

    saved_at = load_shared_notice_saved_at()
    if saved_at:
        st.caption(f"최종 저장: {saved_at}")
    else:
        st.caption("최종 저장: 없음")


def main():
    install_playwright_browser()

    st.set_page_config(page_title="TBM PPT Maker", layout="wide")
    hide_streamlit_ui()

    render_app_title()

    render_daily_safety_meeting()

    st.markdown("---")
    st.markdown(
        """
        <h2 style="color:#0057d9; font-weight:800; font-size:48px; margin-top:0.35rem; margin-bottom:0.35rem;">
            TBM 번역 PPT
        </h2>
        """,
        unsafe_allow_html=True
    )

    render_tbm_input_area()

    render_shared_notice_board()


if __name__ == "__main__":
    main()