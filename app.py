import os
import io
import re
import json
import time
import zipfile
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
TEMP_UPLOAD_DIR = os.path.join(BASE_DIR, "temp_upload")

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

DAILY_PHOTO_BOX_TEXT = "PHOTO_BOX_1"
DAILY_TEXT_BOX_TEXT = "TEXT_BOX_1"
TIME_BOX_TEXT = "TIME_BOX_1"
HOLD_POINT_TEXT = "HOLD POINT"

TEMP_UPLOAD_LIMIT_MB = 100
TEMP_UPLOAD_LIMIT_BYTES = TEMP_UPLOAD_LIMIT_MB * 1024 * 1024
TEMP_UPLOAD_EXPIRE_SECONDS = 24 * 60 * 60

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


@dataclass
class DailySlideData:
    image_path: str
    text: str = ""


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

        .temp-small-button button {
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


def ensure_temp_upload_dir():
    os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)


def cleanup_old_temp_files():
    ensure_temp_upload_dir()
    now = time.time()

    for name in os.listdir(TEMP_UPLOAD_DIR):
        path = os.path.join(TEMP_UPLOAD_DIR, name)
        if os.path.isfile(path):
            if now - os.path.getmtime(path) > TEMP_UPLOAD_EXPIRE_SECONDS:
                try:
                    os.remove(path)
                except Exception:
                    pass


def get_temp_upload_size() -> int:
    ensure_temp_upload_dir()
    total = 0

    for name in os.listdir(TEMP_UPLOAD_DIR):
        path = os.path.join(TEMP_UPLOAD_DIR, name)
        if os.path.isfile(path):
            total += os.path.getsize(path)

    return total


def save_temp_upload_file(uploaded_file):
    ensure_temp_upload_dir()

    file_bytes = uploaded_file.getvalue()
    file_size = len(file_bytes)
    original_filename = safe_filename(uploaded_file.name)

    for existing_name in os.listdir(TEMP_UPLOAD_DIR):
        if existing_name.endswith(original_filename):
            existing_path = os.path.join(TEMP_UPLOAD_DIR, existing_name)
            if os.path.isfile(existing_path) and os.path.getsize(existing_path) == file_size:
                return existing_path

    current_size = get_temp_upload_size()

    if current_size + file_size > TEMP_UPLOAD_LIMIT_BYTES:
        raise ValueError(
            f"임시업로드 용량 초과: 현재 {format_size(current_size)} / "
            f"추가 {format_size(file_size)} / 최대 {TEMP_UPLOAD_LIMIT_MB}MB"
        )

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    save_name = f"{timestamp}_{original_filename}"
    save_path = os.path.join(TEMP_UPLOAD_DIR, save_name)

    with open(save_path, "wb") as f:
        f.write(file_bytes)

    return save_path


def save_generated_ppt_to_temp_upload(ppt_bytes: io.BytesIO, filename: str):
    ensure_temp_upload_dir()

    ppt_bytes.seek(0)
    data = ppt_bytes.getvalue()
    file_size = len(data)

    current_size = get_temp_upload_size()

    if current_size + file_size > TEMP_UPLOAD_LIMIT_BYTES:
        ppt_bytes.seek(0)
        return False

    safe_name = safe_filename(filename)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    save_name = f"{timestamp}_{safe_name}"
    save_path = os.path.join(TEMP_UPLOAD_DIR, save_name)

    with open(save_path, "wb") as f:
        f.write(data)

    ppt_bytes.seek(0)
    return True


def make_temp_upload_zip():
    ensure_temp_upload_dir()

    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for name in sorted(os.listdir(TEMP_UPLOAD_DIR)):
            path = os.path.join(TEMP_UPLOAD_DIR, name)
            if os.path.isfile(path):
                zip_file.write(path, arcname=name)

    zip_buffer.seek(0)
    return zip_buffer


def delete_all_temp_uploads():
    ensure_temp_upload_dir()
    deleted_count = 0

    for name in os.listdir(TEMP_UPLOAD_DIR):
        path = os.path.join(TEMP_UPLOAD_DIR, name)
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


def render_temp_upload():
    cleanup_old_temp_files()

    files = []
    ensure_temp_upload_dir()

    for name in sorted(os.listdir(TEMP_UPLOAD_DIR), reverse=True):
        path = os.path.join(TEMP_UPLOAD_DIR, name)
        if os.path.isfile(path):
            files.append((name, path, os.path.getsize(path)))

    col_title, col_zip, col_delete = st.columns([5.2, 0.9, 0.9])

    with col_title:
        used = get_temp_upload_size()
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
                <span style="font-size:18px; font-weight:700;">임시업로드</span>
                <span style="font-size:13px; color:#666;">
                    용량 {format_size(used)} / {TEMP_UPLOAD_LIMIT_MB}MB
                </span>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col_zip:
        st.markdown("<div class='temp-small-button'>", unsafe_allow_html=True)
        if files:
            st.download_button(
                "ZIP",
                data=make_temp_upload_zip(),
                file_name="임시업로드_전체.zip",
                mime="application/zip",
                use_container_width=True,
                key="temp_download_all_zip"
            )
        else:
            st.button("ZIP", disabled=True, use_container_width=True, key="temp_download_all_zip_disabled")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_delete:
        st.markdown("<div class='temp-small-button'>", unsafe_allow_html=True)
        if st.button("전체삭제", use_container_width=True, key="temp_delete_all"):
            delete_all_temp_uploads()
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    upload_files = st.file_uploader(
        "임시업로드",
        accept_multiple_files=True,
        type=["jpg", "jpeg", "png", "webp", "heic", "heif", "mpo", "pdf", "pptx", "xlsx", "docx", "txt"],
        key="daily_temp_upload_uploader"
    )

    if upload_files:
        uploaded_count = 0

        for file in upload_files:
            try:
                save_temp_upload_file(file)
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

                with col2:
                    st.caption(name)
                    st.caption(format_size(size))

                    with open(path, "rb") as f:
                        st.download_button(
                            label="다운로드",
                            data=f,
                            file_name=name,
                            use_container_width=True,
                            key=f"temp_download_{name}_{idx}"
                        )

                with col3:
                    if st.button("X", key=f"temp_delete_{name}_{idx}", use_container_width=True):
                        if delete_file(path):
                            st.rerun()
                        else:
                            st.error("삭제 실패")
    else:
        st.info("임시업로드 파일 없음.")


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

        img.save(output_path, format="JPEG", quality=quality, optimize=True)
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


def normalize_for_match(text: str) -> str:
    return re.sub(r"\s+", "", str(text or "")).upper()


def get_paddle_ocr():
    """PaddleOCR 지연 로딩. 설치되지 않았거나 초기화 실패 시 None 반환."""
    global _PADDLE_OCR

    if PaddleOCR is None:
        return None

    if _PADDLE_OCR is not None:
        return _PADDLE_OCR

    try:
        _PADDLE_OCR = PaddleOCR(
            lang="korean",
            use_angle_cls=True,
            show_log=False
        )
        return _PADDLE_OCR
    except Exception:
        _PADDLE_OCR = None
        return None


def prepare_ocr_image(img: Image.Image, upscale: int = 2) -> Image.Image:
    """현장 사진 OCR용 전처리: 원본 해상도 유지 + 대비/선명도 보정."""
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass

    if img.mode != "RGB":
        img = img.convert("RGB")

    width, height = img.size

    # 너무 큰 원본은 메모리 보호용으로만 제한. PPT용 축소와 별개로 OCR은 최대한 크게 유지.
    longest = max(width, height)
    if longest > 4200:
        ratio = 4200 / longest
        img = img.resize((int(width * ratio), int(height * ratio)), Image.LANCZOS)
        width, height = img.size

    if upscale > 1 and max(width, height) < 2600:
        img = img.resize((width * upscale, height * upscale), Image.LANCZOS)

    gray = img.convert("L")
    gray = ImageOps.autocontrast(gray)
    gray = ImageEnhance.Contrast(gray).enhance(1.8)
    gray = ImageEnhance.Sharpness(gray).enhance(2.0)
    gray = gray.filter(ImageFilter.SHARPEN)

    return gray


def ocr_with_paddle(img: Image.Image) -> str:
    ocr = get_paddle_ocr()
    if ocr is None:
        return ""

    temp_path = None
    try:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_path = tmp.name
        tmp.close()

        if img.mode != "RGB":
            save_img = img.convert("RGB")
        else:
            save_img = img

        save_img.save(temp_path, format="PNG")

        result = ocr.ocr(temp_path, cls=True)
        lines = []

        if result:
            for page in result:
                if not page:
                    continue
                for line in page:
                    try:
                        text = line[1][0]
                        if text:
                            lines.append(str(text))
                    except Exception:
                        continue

        return "\n".join(lines).strip()

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
    ]
    langs = ["kor+eng", "eng", None]

    best_text = ""

    for lang in langs:
        for config in configs:
            try:
                if lang:
                    text = pytesseract.image_to_string(img, lang=lang, config=config)
                else:
                    text = pytesseract.image_to_string(img, config=config)
                text = str(text or "").strip()
                if len(text) > len(best_text):
                    best_text = text
            except Exception:
                continue

    return best_text.strip()


def extract_ocr_text(image_path: str, top_ratio: float = 0.55) -> str:
    """
    OCR 강화 버전.
    - PPT용 축소본이 아니라 원본 이미지 경로를 넣는 것이 핵심.
    - 전체/상단/좌측 영역을 함께 읽어 회사명, 25대 고위험 문구 누락을 줄임.
    - PaddleOCR이 있으면 우선 사용하고, 없으면 pytesseract로 fallback.
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

        width, height = img.size
        crop_h = max(1, int(height * top_ratio))

        regions = [
            img,
            img.crop((0, 0, width, crop_h)),
            img.crop((0, 0, int(width * 0.68), height)),
            img.crop((0, 0, int(width * 0.68), crop_h)),
        ]

        texts = []
        seen = set()

        for region in regions:
            prepared = prepare_ocr_image(region)

            text = ocr_with_paddle(prepared)
            if not text:
                text = ocr_with_tesseract(prepared)

            text = str(text or "").strip()
            key = normalize_for_match(text)

            if text and key not in seen:
                texts.append(text)
                seen.add(key)

        return "\n".join(texts).strip()

    except Exception:
        return ""


def extract_top_ocr_text(image_path: str, top_ratio: float = 0.55) -> str:
    """기존 함수명 호환용. 내부는 강화 OCR 사용."""
    return extract_ocr_text(image_path, top_ratio=top_ratio)


def detect_high_risk(text: str) -> bool:
    compact = normalize_for_match(text)
    loose = str(text or "")

    if "고위험" in compact:
        return True
    if "25대" in compact:
        return True

    for keyword in HIGH_RISK_KEYWORDS:
        if normalize_for_match(keyword) in compact or keyword in loose:
            return True

    return False


def detect_company(text: str) -> str:
    compact = normalize_for_match(text)

    for company in COMPANY_ORDER:
        aliases = COMPANY_ALIAS.get(company, [company])
        for alias in aliases:
            if normalize_for_match(alias) in compact:
                return company

    # OCR 오인식 보정: 공백/기호 제거 후 유사도 기반 보조 판단
    tokens = re.findall(r"[가-힣A-Za-z0-9]{2,}", str(text or ""))
    compact_tokens = [normalize_for_match(t) for t in tokens]

    best_company = "기타업체"
    best_score = 0.0

    for company in COMPANY_ORDER:
        aliases = COMPANY_ALIAS.get(company, [company])
        for alias in aliases:
            alias_norm = normalize_for_match(alias)
            if not alias_norm:
                continue
            for token in compact_tokens:
                if len(token) < 2:
                    continue
                score = SequenceMatcher(None, alias_norm, token).ratio()
                if score > best_score:
                    best_score = score
                    best_company = company

    if best_score >= 0.72:
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
    target = normalize_text(target_text)

    for shape in iter_all_shapes(slide.shapes):
        if has_text(shape):
            if normalize_text(shape.text) == target:
                return ("shape", shape)

    for shape in iter_all_shapes(slide.shapes):
        if hasattr(shape, "has_table") and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if normalize_text(cell.text) == target:
                        return ("cell", cell)

    return None


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
    slide.shapes.add_picture(
        image_path,
        target_shape.left,
        target_shape.top,
        width=target_shape.width,
        height=target_shape.height
    )


def duplicate_slide(prs, source_slide):
    blank_slide_layout = prs.slide_layouts[6]
    new_slide = prs.slides.add_slide(blank_slide_layout)

    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(
            new_el,
            "p:extLst"
        )

    return new_slide


def fill_slide_by_placeholders(slide, item: SlideData, strict: bool = True):
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
        if strict:
            raise ValueError(f"슬라이드에서 플레이스홀더를 찾지 못했습니다: {', '.join(missing)}")
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

    for i, item in enumerate(slide_data_list):
        if i >= len(prs.slides):
            break

        fill_slide_by_placeholders(prs.slides[i], item, strict=True)

    delete_extra_slides(prs, len(slide_data_list))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
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

    fill_material_slides(prs, material_items)

    start_slide_index = 3

    for i, item in enumerate(bad_items):
        target_index = start_slide_index + i

        if target_index >= len(prs.slides):
            break

        fill_daily_slide(prs.slides[target_index], item, strict=False)

    # 일일안전회의 PPT 회색화면 방지를 위해 슬라이드 삭제를 하지 않음.

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
                    key=f"main_tbm_ko_{idx}"
                )

                slide_inputs.append(SlideData(jpg_path, ko_input, "", "", ""))

        if st.button("PPT 생성", key="main_create_btn"):
            try:
                if "GPT_API_KEY" not in st.secrets:
                    raise ValueError("Secrets에 GPT_API_KEY 설정 필요")

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

                save_generated_ppt_to_temp_upload(ppt, OUTPUT_PPT_NAME)

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
    st.markdown("## 일일안전회의")

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
        st.markdown("#### 부적합사진")

        for idx, f in enumerate(bad_files):
            suffix = os.path.splitext(f.name)[1].lower() or ".jpg"

            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(f.getbuffer())
                original_path = tmp.name
                temp_paths.append(original_path)

            jpg_path = convert_to_jpg(original_path)
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

            bad_items.append(DailySlideData(jpg_path, text_value))

    material_files = st.file_uploader(
        "자재입고 및 고위험작업",
        accept_multiple_files=True,
        type=["jpg", "png", "jpeg", "webp", "heic", "heif", "mpo"],
        key="daily_material_uploader"
    )

    if material_files:
        st.markdown("#### 자재입고 및 고위험작업")

        for idx, f in enumerate(material_files):
            suffix = os.path.splitext(f.name)[1].lower() or ".jpg"

            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(f.getbuffer())
                material_original_path = tmp.name
                temp_paths.append(material_original_path)

            material_jpg_path = convert_to_jpg(material_original_path)
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

            save_generated_ppt_to_temp_upload(ppt, DAILY_OUTPUT_PPT_NAME)

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
    render_temp_upload()
    st.markdown("---")


def main():
    install_playwright_browser()

    st.set_page_config(page_title="TBM PPT Maker", layout="wide")
    hide_streamlit_ui()

    render_app_title()

    render_daily_safety_meeting()

    st.markdown("---")
    st.markdown("## TBM 번역 PPT")

    render_tbm_input_area()


if __name__ == "__main__":
    main()