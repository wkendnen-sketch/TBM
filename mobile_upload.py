"""Mobile XLSX receiving view for the existing Streamlit file store.

Uses Streamlit's normal uploader, including its XSRF/session handling. No separate
public API, extra credentials or external storage is introduced.
"""
import hashlib
import html
import io
import os
import re
import tempfile
import threading
import time
import zipfile
from pathlib import Path

_LOCK = threading.Lock()
MAX_UPLOAD_BYTES = 50 * 1024 * 1024


def validate_xlsx(data: bytes):
    if not data or len(data) > MAX_UPLOAD_BYTES:
        raise ValueError("파일은 50MB 이하의 엑셀 사진대지여야 합니다.")
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            names = set(archive.namelist())
            if not {"[Content_Types].xml", "xl/workbook.xml"}.issubset(names):
                raise ValueError("정상적인 XLSX 파일이 아닙니다.")
            if len(names) > 3000 or sum(i.file_size for i in archive.infolist()) > 150 * 1024 * 1024:
                raise ValueError("엑셀 내부 데이터가 너무 큽니다.")
            if "xl/vbaProject.bin" in names:
                raise ValueError("매크로가 포함된 파일은 사용할 수 없습니다.")
    except zipfile.BadZipFile as exc:
        raise ValueError("정상적인 XLSX 파일이 아닙니다.") from exc


def save_mobile_report(data: bytes, filename: str, folder: str, limit_bytes: int):
    validate_xlsx(data)
    if not filename.lower().endswith(".xlsx"):
        raise ValueError("XLSX 파일만 전송할 수 있습니다.")
    digest = hashlib.sha256(data).hexdigest()
    safe = re.sub(r"[^a-zA-Z0-9가-힣._ -]", "_", os.path.basename(filename))[:100]
    directory = Path(folder)
    directory.mkdir(parents=True, exist_ok=True)
    # Stable digest in the name makes retries idempotent across app/server sessions.
    suffix = f"_mobile_{digest[:24]}_{safe}"
    with _LOCK:
        for existing in directory.glob("*" + suffix):
            if existing.is_file() and hashlib.sha256(existing.read_bytes()).hexdigest() == digest:
                return existing.name, digest, False
        used = sum(p.stat().st_size for p in directory.iterdir() if p.is_file())
        if used + len(data) > limit_bytes:
            raise ValueError("부적합사진 보관함 용량이 부족합니다. 기존 파일을 정리한 뒤 다시 전송하세요.")
        name = time.strftime("%Y%m%d_%H%M%S") + suffix
        fd, temporary = tempfile.mkstemp(prefix=".mobile-", dir=directory)
        try:
            with os.fdopen(fd, "wb") as output:
                output.write(data)
                output.flush()
                os.fsync(output.fileno())
            os.replace(temporary, directory / name)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
    return name, digest, True


def render_mobile_upload(folder, limit_bytes, cleanup):
    import streamlit as st
    cleanup()
    st.markdown('<div id="mobile-upload-ready" data-version="1"></div>', unsafe_allow_html=True)
    st.title("사진대지 자동 업로드")
    st.caption("완성된 엑셀을 부적합사진 보관함으로 전달합니다.")
    uploaded = st.file_uploader("엑셀 사진대지 전송", type=["xlsx"], accept_multiple_files=False, key="mobile_report_xlsx")
    if uploaded is None:
        st.info("앱에서 파일을 연결하고 있습니다. 진행되지 않으면 위 업로드 버튼을 누르세요.")
        return
    try:
        name, digest, created = save_mobile_report(uploaded.getvalue(), uploaded.name, folder, limit_bytes)
        st.success("업로드 완료" if created else "이미 업로드된 파일입니다.")
        st.markdown(f'<div id="mobile-upload-result" data-status="saved" data-sha256="{digest}">{html.escape(name)}</div>', unsafe_allow_html=True)
        st.caption("사이트 보관기간은 24시간입니다. 휴대폰에 저장한 원본은 유지됩니다.")
    except Exception as exc:
        st.error(str(exc))
        st.markdown(f'<div id="mobile-upload-error">{html.escape(str(exc))}</div>', unsafe_allow_html=True)
