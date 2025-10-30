# ==========================================================
# OPS2TBM — 안전 교육 및 점검 일지 반영 (TBM 및 일일안전교육 일지 확장)
#  - 기존 UI/레이아웃 유지, 문서에 추가 필드 포함
# ==========================================================

import io, zipfile, re
from collections import Counter
from typing import List, Dict
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text  # 안정 경로
import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ----------------------------
# 세션 상태
# ----------------------------
# 세션 상태 초기화 (업로드한 파일의 키값, 용어/행동/질문 용어 집합을 저장)
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "kb_terms" not in st.session_state:
    st.session_state.kb_terms: Counter = Counter()
if "kb_actions" not in st.session_state:
    st.session_state.kb_actions: List[str] = []
if "kb_questions" not in st.session_state:
    st.session_state.kb_questions: List[str] = []

# ==========================================================
# 전처리 및 텍스트 정리
# ==========================================================
# 잡음 패턴(불필요한 헤더, 링크 등을 제거하기 위한 정규식)
NOISE_PATTERNS = [
    r"^제?\s?\d{4}\s?[-.]?\s?\d+\s?호$",
    r"^(동절기\s*주요사고|안전작업방법|콘텐츠링크|책자\s*OPS|숏폼\s*OPS)$",
    r"^(포스터|책자|스티커|콘텐츠 링크)$",
    r"^(스마트폰\s*APP|중대재해\s*사이렌|산업안전포털|고용노동부)$",
    r"^https?://\S+$",
    r"^\(?\s*PowerPoint\s*프레젠테이션\s*\)?$",
    r"^안전보건자료실.*$",
]
# 불릿포인트 문자 패턴
BULLET_PREFIX = r"^[\s\-\•\●\▪\▶\▷\·\*\u25CF\u25A0\u25B6\u25C6\u2022\u00B7\u279C\u27A4\u25BA\u25AA\u25AB\u2611\u2713\u2714\u2716\u2794\u27A2]+"
# 날짜 패턴
DATE_PAT = r"(\d{4})\.\s?(\d{1,2})\.\s?(\d{1,2})\.?"
# 사고사례에서 나타날 수 있는 패턴
META_PATTERNS = [r"<\s*\d+\s*명\s*사망\s*>", r"<\s*\d+\s*명\s*사상\s*>", r"<\s*\d+\s*명\s*의식불명\s*>"]
# 불용어(구분된 용어로부터 제거할 단어들)
STOP_TERMS = set([
    "및","등","관련","사항","내용","예방","안전","작업","현장","교육","방법","기준","조치",
    "실시","확인","필요","경우","대상","사용","관리","점검","적용","정도","주의","중","전","후",
    "주요","사례","안전작업방법","포스터","동영상","리플릿","가이드","자료실","검색",
])

def normalize_text(t: str) -> str:
    t = t.replace("\x0c","\n")
    t = re.sub(r"[ \t]+\n","\n",t)
    t = re.sub(r"\n{3,}","\n\n",t)
    return t.strip()

def strip_noise_line(line: str) -> str:
    """라인에서 불필요한 문자를 제거하고 필요한 문장만 반환"""
    s = (line or "").strip()
    if not s: return ""
    s = re.sub(BULLET_PREFIX,"",s).strip()
    for pat in NOISE_PATTERNS:
        if re.match(pat,s,re.IGNORECASE):
            return ""
    s = re.sub(r"https?://\S+","",s).strip()
    s = s.strip("•●▪▶▷·-—–")
    return s

def merge_broken_lines(lines: List[str]) -> List[str]:
    """줄 끝 문장부호/머리표시에 따라 자연 병합"""
    out=[]; buf=""
    for raw in lines:
        s = strip_noise_line(raw)
        if not s: continue
        s = s.lstrip("↳").strip()
        if buf and (buf.endswith((":", "·", "•", "▪", "▶", "▷")) or re.match(r"^\(.*\)$", buf)):
            buf += " " + s
        else:
            if buf: out.append(buf)
            buf = s
    if buf: out.append(buf)
    return out

def combine_date_with_next(lines: List[str]) -> List[str]:
    """YYYY.MM.DD 다음 줄의 '사고/사망/중독/무너짐/부딪힘...'과 결합해 사건문 형태로"""
    out=[]; i=0
    while i<len(lines):
        cur=strip_noise_line(lines[i])
        if re.search(DATE_PAT,cur) and (i+1)<len(lines):
            nxt=strip_noise_line(lines[i+1])
            if re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|무너짐)",nxt):
                m=re.search(DATE_PAT,cur); y,mo,d=m.groups()
                out.append(f"{int(y)}년 {int(mo)}월 {int(d)}일, {nxt}")
                i+=2; continue
        out.append(cur); i+=1
    return out

def preprocess_text_to_sentences(text: str) -> List[str]:
    """문서에서 클린 문장 리스트로 변환"""
    text = normalize_text(text)
    raw_lines = [ln for ln in text.splitlines() if ln.strip()]
    lines = merge_broken_lines(raw_lines)
    lines = combine_date_with_next(lines)
    joined = "\n".join(lines)
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", joined)
    sents=[]
    for s in raw:
        s2 = strip_noise_line(s)
        if not s2: continue
        if re.search(r"(주요사고|안전작업방법|콘텐츠링크)$",s2): continue
        if len(s2)<6: continue
        sents.append(s2)
    # 중복 제거
    seen=set(); dedup=[]
    for s in sents:
        k=re.sub(r"\s+","",s)
        if k not in seen:
            seen.add(k); dedup.append(s)
    return dedup

# ==========================================================
# PDF 텍스트 추출 후 처리 - 빈 텍스트 체크 추가
# ==========================================================
def read_pdf_text(b: bytes) -> str:
    """PDF에서 텍스트를 추출하는 함수"""
    try:
        with io.BytesIO(b) as bio:
            t = pdf_extract_text(bio) or ""
    except Exception:
        t = ""
    t = normalize_text(t)
    if len(t.strip()) < 10:  # 텍스트가 너무 짧으면, 이미지/스캔 PDF로 판단
        try:
            with io.BytesIO(b) as bio:
                pdf = pdfium.PdfDocument(bio)
                if len(pdf) > 0 and not t.strip():
                    st.warning("⚠️ 이미지/스캔 PDF로 보입니다. 현재 OCR 미지원.")
        except Exception:
            pass
    return t

# ==========================================================
# 업로드 시 텍스트 처리 (빈 값 체크 추가)
# ==========================================================
extracted = ""

# PDF 파일을 업로드했을 때, 텍스트 추출 및 처리
if uploaded and uploaded.name.lower().endswith(".pdf"):
    with st.spinner("PDF 텍스트 추출 중..."):
        data = uploaded.read()
        extracted = read_pdf_text(data)
        if extracted.strip():  # 추출된 텍스트가 비어 있지 않은 경우
            kb_ingest_text(extracted)   # 🔹 단일 PDF도 즉시 학습
            kb_prune()
        else:
            st.warning("⚠️ PDF에서 유효한 텍스트를 추출할 수 없습니다.")

elif selected_zip_pdf:
    with st.spinner("ZIP 내부 PDF 텍스트 추출 중..."):
        extracted = read_pdf_text(zip_pdfs[selected_zip_pdf])
        if extracted.strip():  # 추출된 텍스트가 비어 있지 않은 경우
            kb_ingest_text(extracted)  # 🔹 단일 PDF도 즉시 학습
            kb_prune()
        else:
            st.warning("⚠️ ZIP에서 PDF를 처리하는 동안 오류가 발생했습니다.")

# 수동 텍스트 입력이 있을 경우, 그 또한 학습 반영
pasted = (st.session_state.get("manual_text") or "").strip()
if pasted:
    kb_ingest_text(pasted)
    kb_prune()

base_text = pasted or extracted.strip()

if not base_text:
    st.warning("⚠️ 텍스트가 비어 있습니다. PDF 업로드 또는 텍스트 입력을 확인해주세요.")

# 텍스트 미리보기 표시
st.markdown("**추출/입력 텍스트 미리보기**")
edited_text = st.text_area("텍스트", value=base_text, height=240, key="edited_text")

# ==========================================================
# TBM 대본 생성 및 다운로드
# ==========================================================
use_ai = st.toggle("🔹 AI 요약(TextRank+MMR) 사용", value=True)
tmpl_choice = st.selectbox("🧩 템플릿", ["자동 선택","사고사례형","가이드형"])  # 표시만 유지
gen_mode = st.selectbox("🧠 생성 모드", ["TBM 기본(현행)","자연스러운 교육대본(무료)"])
max_points = st.slider("요약 강도(핵심문장 개수)", 3, 10, 6)

if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
    text_for_gen = (st.session_state.get("edited_text") or "").strip()
    if not text_for_gen:
        st.warning("텍스트가 비어 있습니다. PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
    else:
        with st.spinner("대본 생성 중..."):
            if gen_mode == "자연스러운 교육대본(무료)":
                script = make_structured_script(text_for_gen, max_points=max_points)
                subtitle = "자연스러운 교육대본(무료)"
            else:
                sents = ai_extract_summary(text_for_gen, max_points if use_ai else 6)
                sents = [soften(s) for s in sents]
                script = "\n".join([f"- {s}" for s in sents]) if sents else "텍스트에서 핵심 문장을 찾지 못했습니다."
                subtitle = "TBM 기본(현행)"

        st.success(f"대본 생성 완료! ({subtitle})")
        st.text_area("대본 미리보기", value=script, height=420)
        c3,c4 = st.columns(2)
        with c3:
            st.download_button("⬇️ TXT 다운로드", data=_xml_safe(script).encode("utf-8"),
                               file_name="tbm_script.txt", use_container_width=True)
        with c4:
            st.download_button("⬇️ DOCX 다운로드", data=to_docx_bytes(script),
                               file_name="tbm_script.docx", use_container_width=True)

st.caption("완전 무료. 업로드/붙여넣기만 해도 누적 학습 → 요약 가중/행동/질문 보강. 동적 주제 라벨. UI 변경 없음.")
