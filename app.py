# =========================
# OPS2TBM (Stable UI + Pure NumPy TextRank)
# - 텍스트 PDF / 텍스트 입력 지원
# - 이미지/스캔 PDF는 미지원(안내)
# - Sidebar에 소개/사용법/시연 멘트 표시
# - 템플릿: 자동/사고사례형/가이드형
# - AI 요약: 순수 NumPy TextRank (설치 이슈 無)
# =========================

import io
from typing import List, Tuple, Dict

import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text
import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ----------------------------
# 공통 유틸
# ----------------------------
def normalize_text(text: str) -> str:
    text = text.replace("\x0c", "\n")
    text = rxx.sub(r"[ \t]+\n", "\n", text)
    text = rxx.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def read_pdf_text(file_bytes: bytes) -> str:
    # 1) pdfminer로 텍스트 추출
    try:
        with io.BytesIO(file_bytes) as bio:
            text = pdf_extract_text(bio) or ""
    except Exception:
        text = ""
    text = normalize_text(text)

    # 2) 텍스트가 거의 없으면(이미지 PDF 추정) 페이지 수 체크 후 안내
    if len(text.strip()) < 10:
        try:
            with io.BytesIO(file_bytes) as bio:
                pdf = pdfium.PdfDocument(bio)
                pages = len(pdf)
            if pages > 0 and not text.strip():
                st.warning("이 PDF는 이미지/스캔 기반으로 보여요. 현재 버전은 OCR 없이 '텍스트'만 처리합니다.")
        except Exception:
            pass
    return text

def split_sentences_ko(text: str) -> List[str]:
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", text)
    return [s.strip() for s in raw if s and len(s.strip()) > 1]

def simple_tokens(s: str) -> List[str]:
    s = s.lower()
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s)

# ----------------------------
# 순수 NumPy TextRank
# ----------------------------
def sentence_tfidf_vectors(sents: List[str]) -> Tuple[np.ndarray, List[str]]:
    toks_list = [simple_tokens(s) for s in sents]
    vocab: Dict[str, int] = {}
    for toks in toks_list:
        for t in toks:
            if t not in vocab:
                vocab[t] = len(vocab)
    if not vocab:
        return np.zeros((len(sents), 0), dtype=np.float32), []
    mat = np.zeros((len(sents), len(vocab)), dtype=np.float32)
    df = np.zeros((len(vocab),), dtype=np.float32)
    for i, toks in enumerate(toks_list):
        for t in toks:
            j = vocab[t]
            mat[i, j] += 1.0
        for t in set(toks):
            df[vocab[t]] += 1.0
    N = float(len(sents))
    idf = np.log((N + 1.0) / (df + 1.0)) + 1.0
    mat *= idf
    norms = np.linalg.norm(mat, axis=1, keepdims=True) + 1e-8
    mat = mat / norms
    return mat, list(vocab.keys())

def cosine_sim_matrix(X: np.ndarray) -> np.ndarray:
    if X.size == 0:
        return np.zeros((X.shape[0], X.shape[0]), dtype=np.float32)
    sim = np.clip(X @ X.T, 0.0, 1.0)
    np.fill_diagonal(sim, 0.0)
    return sim

def textrank_scores(sents: List[str], d: float = 0.85, max_iter: int = 50, tol: float = 1e-4) -> List[float]:
    if len(sents) == 0:
        return []
    if len(sents) == 1:
        return [1.0]
    X, _ = sentence_tfidf_vectors(sents)
    W = cosine_sim_matrix(X)
    row_sums = W.sum(axis=1, keepdims=True)
    P = np.divide(W, row_sums, out=np.zeros_like(W), where=row_sums > 0)

    n = W.shape[0]
    r = np.ones((n, 1), dtype=np.float32) / n
    teleport = np.ones((n, 1), dtype=np.float32) / n

    for _ in range(max_iter):
        r_new = d * (P.T @ r) + (1 - d) * teleport
        if np.linalg.norm(r_new - r, ord=1) < tol:
            r = r_new
            break
        r = r_new
    return [float(v) for v in r.flatten()]

# ----------------------------
# 선택 로직 (규칙/AI)
# ----------------------------
def pick_rule(sents: List[str], keywords: List[str], limit: int) -> List[str]:
    hits = [s for s in sents if any(k in s for k in keywords)]
    if len(hits) >= limit:
        return hits[:limit]
    remain = [s for s in sents if s not in hits]
    remain = sorted(remain, key=lambda x: (-len(x), sents.index(x)))
    return hits + remain[: max(0, limit - len(hits))]

def pick_tr(sents: List[str], keywords: List[str], limit: int, scores: List[float]) -> List[str]:
    if not sents:
        return []
    w = np.array(scores, dtype=np.float32)
    if keywords:
        for i, s in enumerate(sents):
            if any(k in s for k in keywords):
                w[i] += 0.2
    idx = np.argsort(-w)[:limit]
    return [sents[i] for i in idx]

# ----------------------------
# 템플릿/키워드
# ----------------------------
# 가이드형
KW_GUIDE_CORE = ["가이드", "안내", "보호", "건강", "대응", "절차", "지침", "매뉴얼", "예방", "상담", "지원"]
KW_GUIDE_STEP = ["절차", "순서", "방법", "점검", "확인", "보고", "조치"]
KW_GUIDE_QA   = ["질문", "왜", "어떻게", "무엇", "주의"]

# 사고사례형
KW_ACC_CORE = ["사고", "재해", "위험", "원인", "예방", "대책", "노후", "추락", "협착", "감전", "화재"]
KW_ACC_STEP = ["발생", "경위", "조치", "개선", "교육", "설치", "배치"]
KW_ACC_QA   = ["원인은", "다음에는", "예방하려면", "확인할 점", "체크"]

def detect_template(text: str) -> str:
    g_hits = sum(text.count(k) for k in (KW_GUIDE_CORE + KW_GUIDE_STEP))
    a_hits = sum(text.count(k) for k in (KW_ACC_CORE + KW_ACC_STEP))
    return "가이드형" if g_hits >= a_hits else "사고사례형"

# ----------------------------
# TBM 생성
# ----------------------------
def make_tbm_guide(text: str, use_ai: bool) -> Tuple[str, Dict[str, List[str]]]:
    sents = split_sentences_ko(text)
    if use_ai:
        scores = textrank_scores(sents)
        core = pick_tr(sents, KW_GUIDE_CORE, 3, scores)
    else:
        core = pick_rule(sents, KW_GUIDE_CORE, 3)
    steps = pick_rule(sents, KW_GUIDE_STEP, 5)
    qa    = pick_rule(sents, KW_GUIDE_QA,   3)

    parts = {"핵심": core, "절차": steps, "질문": qa}
    lines = []
    lines.append("🦺 TBM 대본 – 가이드형")
    lines.append("")
    lines.append("◎ 오늘의 핵심 포인트")
    for s in core: lines.append(f"- {s}")
    lines.append("")
    lines.append("◎ 작업 전 절차/점검")
    for s in steps: lines.append(f"- {s}")
    lines.append("")
    lines.append("◎ 현장 토의 질문")
    for s in qa: lines.append(f"- {s}")
    lines.append("")
    lines.append("◎ 마무리 멘트")
    lines.append("- “오늘 작업의 핵심은 위 세 가지입니다. 다 같이 확인하고 시작합시다.”")
    lines.append("- “잠깐이라도 이상하면 바로 중지하고, 관리자에게 알립니다.”")
    script = "\n".join(lines)
    return script, parts

def make_tbm_accident(text: str, use_ai: bool) -> Tuple[str, Dict[str, List[str]]]:
    sents = split_sentences_ko(text)
    if use_ai:
        scores = textrank_scores(sents)
        core = pick_tr(sents, KW_ACC_CORE, 3, scores)
    else:
        core = pick_rule(sents, KW_ACC_CORE, 3)
    steps = pick_rule(sents, KW_ACC_STEP, 5)
    qa    = pick_rule(sents, KW_ACC_QA,   3)

    parts = {"핵심": core, "사고/조치": steps, "질문": qa}
    lines = []
    lines.append("🦺 TBM 대본 – 사고사례형")
    lines.append("")
    lines.append("◎ 사고/위험 요인 요약")
    for s in core:  lines.append(f"- {s}")
    lines.append("")
    lines.append("◎ 발생 경위/조치/개선")
    for s in steps: lines.append(f"- {s}")
    lines.append("")
    lines.append("◎ 재발 방지 토의 질문")
    for s in qa:    lines.append(f"- {s}")
    lines.append("")
    lines.append("◎ 마무리 멘트")
    lines.append("- “이 사례에서 배운 예방 포인트를 오늘 작업에 바로 적용합시다.”")
    lines.append("- “각자 맡은 공정에서 동일 위험이 없는지 다시 점검해 주세요.”")
    script = "\n".join(lines)
    return script, parts

# ----------------------------
# 내보내기
# ----------------------------
def to_docx_bytes(script: str) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Malgun Gothic"
    style.font.size = Pt(11)
    for line in script.split("\n"):
        p = doc.add_paragraph(line)
        for run in p.runs:
            run.font.name = "Malgun Gothic"
            run.font.size = Pt(11)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# ----------------------------
# UI 구성 (예전 느낌으로 복원)
# ----------------------------
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**이 앱은 무엇을 하나요?**  
OPS 문서(텍스트 PDF/본문 텍스트)를 넣으면,  
- **사고사례형** 또는 **가이드형** 템플릿으로  
- 현장에서 바로 읽을 수 있는 **TBM 대본**을 자동 생성합니다.

**사용법**  
1) 좌측에 **파일 업로드**(텍스트 PDF) 또는 **텍스트 붙여넣기**  
2) 우측에서 **AI 요약 (TextRank) 토글**과 **템플릿(자동/수동)** 선택  
3) **대본 생성** → 미리보기 확인 → **TXT/DOCX 다운로드**

**시연 멘트**  
- “오늘 작업의 핵심은 위 세 가지입니다. 다 같이 확인하고 시작합시다.”  
- “잠깐이라도 이상하면 바로 중지하고, 관리자에게 알립니다.”  
- “이 사례에서 배운 예방 포인트를 오늘 작업에 바로 적용합시다.”  
- “각자 맡은 공정에서 동일 위험이 없는지 다시 점검해 주세요.”
""")

st.title("🦺 OPS2TBM — OPS/포스터를 TBM 대본으로 자동 변환")

st.markdown("""
**안내**  
- 텍스트가 포함된 PDF 또는 본문 텍스트를 사용하세요.  
- 이미지/스캔형 PDF는 현재 OCR 미지원입니다(텍스트가 없으면 경고가 뜹니다).
""")

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    uploaded = st.file_uploader("OPS 파일 업로드 (텍스트 PDF 권장)", type=["pdf"])
    manual_text = st.text_area("또는 OPS 텍스트 직접 붙여넣기", height=220, placeholder="예: 지붕 작업 중 추락사고 예방을 위해...")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("사고 샘플 넣기", use_container_width=True):
            manual_text = ("2020년 2월, 지붕 썬라이트 위 작업 중 추락 재해 발생. "
                           "작업계획 미흡, 추락방지설비 미설치, 감시자 부재가 주요 원인. "
                           "예방: 안전발판/난간/라이프라인 설치, 취약부 표시, 감시자 배치.")
    with c2:
        if st.button("가이드 샘플 넣기", use_container_width=True):
            manual_text = ("감정노동 근로자 건강보호 안내. 고객의 폭언·폭행 등으로 인한 건강장해 예방과 대응절차 제시. "
                           "사업주는 지침 마련·예방교육·상담 지원, 근로자는 조치 요구 가능, 고객은 존중 의무. "
                           "폭언 발생 시: 중지 요청 → 보고 → 기록 → 휴식/상담·치료 → 재발방지 대책.")

    extracted = ""
    if uploaded:
        with st.spinner("PDF 텍스트 추출 중..."):
            data = uploaded.read()
            extracted = read_pdf_text(data)

    base_text = manual_text.strip() if manual_text.strip() else extracted.strip()
    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=240)

with col2:
    # 모드/템플릿 선택
    use_ai = st.toggle("🔹 AI 요약(TextRank) 사용", value=True)
    tmpl_choice = st.selectbox("🧩 템플릿", ["자동 선택", "사고사례형", "가이드형"])

    if st.button("🛠️ TBM 대본 생성", type="primary", use_container_width=True):
        if not edited_text.strip():
            st.warning("텍스트가 비어 있습니다. PDF 업로드(텍스트 PDF) 또는 텍스트 입력 후 시도하세요.")
        else:
            # 템플릿 결정
            if tmpl_choice == "자동 선택":
                detected = detect_template(edited_text)
            else:
                detected = tmpl_choice

            with st.spinner("대본 생성 중..."):
                if detected == "사고사례형":
                    script, parts = make_tbm_accident(edited_text, use_ai=use_ai)
                    subtitle = "사고사례형 템플릿 적용"
                else:
                    script, parts = make_tbm_guide(edited_text, use_ai=use_ai)
                    subtitle = "가이드형 템플릿 적용"

            st.success(f"대본 생성 완료! ({subtitle}{' · AI 요약' if use_ai else ''})")
            st.text_area("TBM 대본 미리보기", value=script, height=420)

            c3, c4 = st.columns(2)
            with c3:
                st.download_button("⬇️ TXT 다운로드", data=script.encode("utf-8"),
                                   file_name="tbm_script.txt", use_container_width=True)
            with c4:
                st.download_button("⬇️ DOCX 다운로드", data=to_docx_bytes(script),
                                   file_name="tbm_script.docx", use_container_width=True)

st.caption("현재: 규칙 + 순수 NumPy TextRank(경량 AI). 템플릿 자동/수동. OCR 미지원(텍스트 PDF 권장).")
