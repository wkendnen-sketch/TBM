from pptx import Presentation
from pathlib import Path

PPT_PATH = "sample_template.pptx"
REQUIRED_TEXTS = ["1", "2", "3", "4"]
REQUIRED_NAMES = ["PHOTO_BOX"]


def safe_text(shape):
    try:
        if hasattr(shape, "text"):
            return shape.text.strip()
    except Exception:
        pass
    return ""


def inspect_ppt(ppt_path):
    ppt_path = Path(ppt_path)
    if not ppt_path.exists():
        print(f"❌ 파일 없음: {ppt_path}")
        return

    prs = Presentation(str(ppt_path))
    print("=" * 80)
    print(f"PPT 점검 시작: {ppt_path}")
    print(f"총 슬라이드 수: {len(prs.slides)}")
    print("=" * 80)

    all_ok = True

    for slide_idx, slide in enumerate(prs.slides, start=1):
        print(f"\n[슬라이드 {slide_idx}]")
        found_names = set()
        found_texts = set()

        for shape_idx, shape in enumerate(slide.shapes, start=1):
            name = getattr(shape, "name", "")
            text = safe_text(shape)
            found_names.add(name)
            if text:
                found_texts.add(text)

            print(f"{shape_idx:02d}. name=[{name}] text=[{text}] type={shape.shape_type}")

        missing = []
        for key in REQUIRED_NAMES:
            if key not in found_names:
                missing.append(f"이름 없음: {key}")
        for key in REQUIRED_TEXTS:
            if key not in found_names and key not in found_texts:
                missing.append(f"이름/텍스트 없음: {key}")

        if missing:
            all_ok = False
            print("❌ 문제 있음:", ", ".join(missing))
        else:
            print("✅ 정상")

    print("\n" + "=" * 80)
    print("✅ 전체 정상" if all_ok else "❌ 일부 슬라이드 문제 있음")
    print("=" * 80)


if __name__ == "__main__":
    inspect_ppt(PPT_PATH)
