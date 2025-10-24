import io
import re
import regex as rxx
from typing import List, Tuple

import streamlit as st
from PIL import Image
from docx import Document
from docx.shared import Pt
import pypdfium2 as pdfium
from rapidocr_onnxruntime import RapidOCR


@st.cache_resource(show_spinner=False)
def get_ocr():
    return RapidOCR()


def pdf_to_images(file_bytes: bytes, scale: float = 2.0) -> List[Image.Image]:
    pdf = pdfium.PdfDocument(io.BytesIO(file_bytes))
    images = []
    for page in pdf:
        pil = page.render(scale=scale).to_pil()
        images.append(pil)
    return images


def ocr_image(pil_img: Image.Image, ocr: RapidOCR) -> str:
    img = pil_img.convert("RGB")
    result, _ = ocr(img)
    if not result:
        return ""
    lines = [item[1] for item in result]
    return "\n".join(lines)


def extract_text(uploaded_file, ocr: RapidOCR) -> str:
    filename = uploaded_file.name.lower()
    data = uploaded_file.read()
    if filename.endswith(".pdf"):
        images = pdf_to_images(data, scale=2.0)
        texts = [ocr_image(im, ocr) for im in images]
        return "\n".join(texts).strip()
    try:
        pil = Image.open(io.BytesIO(data))
        return ocr_image(pil, ocr)
    except Exception:
        return ""


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\ufeff", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"(출처|자료|작성|페이지)\s*[:：].*", "", text, flags=re.IGNORECASE)
    return text.strip()


KW_OVERVIEW = ["개요", "사례", "사고", "배경", "요약", "현황"]
KW_CAUSE = ["원인", "이유", "문제점", "부적정", "미비", "위험요인"]
KW_RULES = ["예방", "대책", "수칙", "점검", "조치", "확인", "준수", "관리"]

def split_sentences_ko(text: str) -> List[str]:
    sents = rxx.split(r"(?<=[\.!?…]|다\.|다!|다\?)\s+|\n", text)
    sents = [s.strip(" -•\t") for s in sents if len(s.strip()) > 3]
    return sents

def pick_sentences(sents: List[str], keywords: List[str], limit: int) -> List[str]:
    scored = []
    for s in sents:
        score = sum(1 for k in keywords if k in s)
        if score > 0:
            scored.append((score, len(s), s))
    scored.sort(key=lambda x: (-x[0], x[1]))
    out = [s for _, _, s in scored[:limit]]
    return out

def bulletize(lines: List[str]) -> List[str]:
    blt = []
    for s in lines:
        s = re.sub(r"\s{2,}", " ", s).strip(" -•\t")
        if len(s) > 120:
            s = s[:117] + "…"
        blt.append(s)
    return blt


SAFETY_CONSTANTS = [
    "선조치 후작업(안전설비·난간·라이프라인 설치 후 작업)",
    "감시자 배치 및 위험구역 출입 통제",
    "개인보호구 착용(안전모·안전벨트·안전화 등) 철저",
    "작업계획서·위험성평가 사전 검토 및 TBM 공유",
    "추락·협착 등 고위험 작업 시 작업중지 기준 숙지",
]

ROOF_EXTRAS = [
    "투광판(썬라이트) 위 절대 밟지 않기(취약부 표시)",
    "지붕 작업 시 안전발판·난간·라이프라인·추락방지망 설치",
    "기상(강풍·우천) 불량 시 작업 중지",
]


def make_tbm_script(raw_text: str) -> Tuple[str, dict]:
    text = clean_text(raw_text)
    sents = split_sentences_ko(text)
    overview = pick_sentences(sents, KW_OVERVIEW, 3)
    causes = pick_sentences(sents, KW_CAUSE, 4)
    rules = pick_sentences(sents, KW_RULES, 6)
    if len(rules) < 4:
        rules = rules + SAFETY_CONSTANTS[: (4 - len(rules)) + 1]
    rules = rules[:6]
    title = "OPS 기반 안전 TBM"
    for cand in sents[:5]:
        if any(k in cand for k in ["지붕", "추락", "질식", "화재", "협착", "감전"]):
            title = cand.strip(" .")
            if len(title) > 22:
                title = title[:20] + "…"
            break
    closing = [
        "‘잠깐이면 돼’는 가장 위험한 말입니다. 애매하면 멈추고 점검합시다.",
        "오늘 작업 전, 취약부·안전설비·PPE·감시자 여부를 다시 확인합시다.",
    ]
    chant = "한 번 더 점검! 모두가 안전!"
    lines = []
    lines.append(f"🦺 TBM 대본 – 「{title}」\n")
    lines.append("◎ 인사 및 도입\n- 오늘 TBM에서는 OPS 자료를 바탕으로 최근 재해 위험요인을 짚고, 우리 현장에서 바로 적용할 수 있는 안전수칙을 공유합니다.\n")
    lines.append("◎ 1. 사고 개요")
    for s in overview or ["(OPS에서 개요를 찾지 못했습니다. 일반 개요로 대체합니다.)"]:
        lines.append(f"- {s}")
    lines.append("\n◎ 2. 사고 원인")
    for s in bulletize(causes or ["작업계획 부재", "보호구 미착용", "감시자 부재"]):
        lines.append(f"- {s}")
    lines.append("\n◎ 3. 주요 안전수칙(우리 현장 적용)")
    for s in bulletize(rules):
        lines.append(f"- {s}")
    if any(k in text for k in ["지붕", "썬라이트", "투광판", "추락"]):
        for s in ROOF_EXTRAS:
            lines.append(f"- {s}")
    lines.append("\n◎ 4. 마무리 당부")
    for s in closing:
        lines.append(f"- {s}")
    lines.append("\n◎ 마무리 구호")
    lines.append(f"- {chant}")
    script = "\n".join(lines).strip()
    parts = {
        "title": title,
        "overview": overview,
        "causes": causes,
        "rules": rules,
        "closing": closing,
        "chant": chant,
    }
    return script, parts


def to_txt_bytes(text: str) -> bytes:
    return text.encode("utf-8")

def to_docx_bytes(text: str) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style.font.size = Pt(11)
    for line in text.split("\n"):
        doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")
st.title("🦺 OPS2TBM — OPS/포스터를 TBM 대본으로 자동 변환")

st.markdown("""
**사용법**
1) OPS 파일 업로드(PDF/PNG/JPG) 또는 텍스트 붙여넣기  
2) 대본 생성 버튼 클릭  
3) .txt / .docx 다운로드
""")

col1, col2 = st.columns([1, 1])
with col1:
    uploaded = st.file_uploader("OPS 파일 업로드 (PDF/PNG/JPG)", type=["pdf", "png", "jpg", "jpeg"])
    manual_text = st.text_area("또는 OPS 텍스트 직접 붙여넣기", height=220)
    if "ocr_model" not in st.session_state:
        st.session_state.ocr_model = get_ocr()
    extracted = ""
    if uploaded:
        with st.spinner("파일에서 텍스트 추출 중..."):
            extracted = extract_text(uploaded, st.session_state.ocr_model)
    base_text = manual_text.strip() if manual_text.strip() else extracted.strip()
    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=260)

with col2:
    if st.button("🛠️ TBM 대본 생성", type="primary", use_container_width=True):
        if not edited_text.strip():
            st.warning("텍스트가 비어 있습니다.")
        else:
            script, parts = make_tbm_script(edited_text)
            st.success("대본 생성 완료!")
            st.text_area("TBM 대본", value=script, height=400)
            st.download_button("⬇️ .txt 다운로드", data=to_txt_bytes(script), file_name="tbm_script.txt")
            st.download_button("⬇️ .docx 다운로드", data=to_docx_bytes(script), file_name="tbm_script.docx")

st.caption("⚙️ OCR: RapidOCR • PDF: pypdfium2 • 배포: Streamlit Cloud • Made for AI 해커톤")
