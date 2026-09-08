# Streamlit 최종 최적화본

## 기준
- GitHub 원본 `app.py` 4,161줄 기준 → 최종 `app.py` 3,705줄
- 현재 사용 중인 Streamlit Community Cloud 구조 유지
- 기존 PPT/Excel 템플릿 3개 포함
- `packages.txt`는 포함하지 않음

## 제거
- PaddleOCR / Tesseract 구형 OCR 경로 전체
- 해당 OCR 전처리/업체판독/고위험판독 보조 함수
- 실제 실행 경로에서 호출되지 않는 함수
- 미사용 dataclass 필드
- 미사용 import/상수

## 개선
- TBM 업로드 시 문구를 수정할 때마다 임시 JPG를 계속 생성하던 구조 제거
  - 실제 이미지 변환은 `PPT 생성` 버튼을 누른 뒤에만 실행
- 부적합사진 업로더의 동일 파일 중복 저장/반복 rerun 방지
- TBM 번역 API에 429 rate-limit 자동 재시도 적용
- 일일안전회의 안내 문구를 실제 GPT 이미지 분류 방식에 맞게 정리
- `.streamlit/config.toml`에 300MB 업로드 한도 설정

## 유지
- TBM 다국어 번역 PPT
- 일일안전회의 PPT
- GPT 이미지 분류(자재입고/25대 고위험/업체/순번)
- 부적합사진 저장/ZIP/삭제
- 공유 공지
- 체감온도 GPT 판독/엑셀 누적/중복방지/수정/완료파일
- 기존 템플릿 구조

## 배포
GitHub 저장소 최상단에 이 폴더의 내용물을 그대로 업로드하세요.

필수 구조:
- app.py
- requirements.txt
- .streamlit/config.toml
- templates/sample_template.pptx
- templates/sample_template2.pptx
- templates/heat_index_template.xlsx

Streamlit Secrets의 `GPT_API_KEY`는 기존 값을 그대로 사용합니다.

중요: `packages.txt`는 만들지 마세요.

## 검증 결과
- Python 문법 컴파일: 통과
- 정적 실행경로 검사: `main()` 기준 미사용 최상위 함수 0개
- TBM 템플릿 로딩 + 샘플 PPT 생성: 통과
- 일일안전회의 템플릿 로딩 + 샘플 PPT 생성: 통과
- 체감온도 Excel 템플릿 로딩: 통과
- `pytesseract`, `PaddleOCR`, `paddleocr` 참조: 제거 완료
