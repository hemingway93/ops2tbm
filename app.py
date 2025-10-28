# ==========================================================
# OPS2TBM — 완전 무료 AI 버전 (규칙 준수)
# 작성 목적:
#   - 기존 UI/레이아웃 유지
#   - AI 요약(TextRank+MMR) + 문체 변환 + 3분형 교육대본 자동 구성
# ==========================================================

import io, zipfile
from typing import List, Tuple
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text
import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ==========================================================
# 1️⃣ 기본 세팅 및 공통 함수
# ==========================================================

if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ──────────────────────────────────────────────
# 텍스트 정리 (공백/줄바꿈 정리)
# ──────────────────────────────────────────────
def normalize_text(t: str) -> str:
    t = t.replace("\x0c", "\n")
    t = rxx.sub(r"[ \t]+\n", "\n", t)
    t = rxx.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

# ──────────────────────────────────────────────
# PDF 텍스트 추출 (텍스트형 PDF만 가능)
# ──────────────────────────────────────────────
def read_pdf_text(b: bytes) -> str:
    try:
        with io.BytesIO(b) as bio:
            text = pdf_extract_text(bio) or ""
    except Exception:
        text = ""
    text = normalize_text(text)
    if len(text.strip()) < 10:
        try:
            pdf = pdfium.PdfDocument(io.BytesIO(b))
            if len(pdf) > 0:
                st.warning("⚠️ OCR 미지원 PDF입니다. 텍스트가 없습니다.")
        except Exception:
            pass
    return text

# ==========================================================
# 2️⃣ AI 요약(TextRank + MMR)
# ==========================================================
# OPS 문서에서 문장 간 의미 유사도를 계산해 핵심 문장만 추출하는 비지도 AI 방식

def split_sentences_ko(text: str) -> List[str]:
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", text)
    return [s.strip(" -•●▪▶▷\t") for s in raw if len(s.strip()) > 5]

def simple_tokens(s: str) -> List[str]:
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

def sentence_tfidf_vectors(sents: List[str]):
    toks = [simple_tokens(s) for s in sents]
    vocab = {}
    for ts in toks:
        for t in ts:
            if t not in vocab:
                vocab[t] = len(vocab)
    if not vocab:
        return np.zeros((len(sents), 0)), []
    M = np.zeros((len(sents), len(vocab)))
    df = np.zeros(len(vocab))
    for i, ts in enumerate(toks):
        for t in ts:
            M[i, vocab[t]] += 1
        for t in set(ts):
            df[vocab[t]] += 1
    idf = np.log((len(sents) + 1) / (df + 1)) + 1
    M *= idf
    M /= np.linalg.norm(M, axis=1, keepdims=True) + 1e-8
    return M, list(vocab.keys())

def cosine_sim_matrix(X):
    if X.size == 0:
        return np.zeros((X.shape[0], X.shape[0]))
    sim = np.clip(X @ X.T, 0, 1)
    np.fill_diagonal(sim, 0)
    return sim

def textrank_scores(sents: List[str]) -> List[float]:
    n = len(sents)
    if n == 0:
        return []
    X, _ = sentence_tfidf_vectors(sents)
    W = cosine_sim_matrix(X)
    row = W.sum(1, keepdims=True)
    P = np.divide(W, row, out=np.zeros_like(W), where=row > 0)
    r = np.ones((n, 1)) / n
    for _ in range(50):
        r_new = 0.85 * P.T @ r + 0.15 / n
        if np.linalg.norm(r_new - r) < 1e-4:
            break
        r = r_new
    return [float(x) for x in r.flatten()]

def mmr_select(sents, scores, X, k=6):
    sim = cosine_sim_matrix(X)
    sel = []
    cand = set(range(len(sents)))
    while cand and len(sel) < k:
        best, bestv = None, -1e9
        for i in cand:
            rel = scores[i]
            div = max((sim[i, j] for j in sel), default=0)
            mmr = 0.7 * rel - 0.3 * div
            if mmr > bestv:
                bestv, best = mmr, i
        sel.append(best)
        cand.remove(best)
    return sel

def ai_extract_summary(text: str, limit: int = 8) -> List[str]:
    sents = split_sentences_ko(text)
    if not sents:
        return []
    X, _ = sentence_tfidf_vectors(sents)
    scores = textrank_scores(sents)
    idx = mmr_select(sents, scores, X, limit)
    return [sents[i] for i in idx]

# ==========================================================
# 3️⃣ 규칙 기반 문체 변환 + 3분 교육대본 자동 구성
# ==========================================================

# ────────── 주요 단어별 주제 감지 ──────────
def detect_topic(t: str) -> str:
    if "온열" in t or "폭염" in t:
        return "온열질환 예방"
    if "질식" in t or "밀폐" in t:
        return "질식재해 예방"
    if "감전" in t:
        return "감전사고 예방"
    if "지붕" in t or "썬라이트" in t:
        return "지붕 작업 추락사고 예방"
    if "컨베이어" in t or "끼임" in t:
        return "컨베이어 끼임사고 예방"
    return "안전보건 교육"

# ────────── 문체 완화 ──────────
def soften(s: str) -> str:
    s = s.replace("한다", "합니다").replace("하여야", "해야 합니다")
    s = s.replace("바랍니다", "해주세요").replace("확인 바람", "확인해주세요")
    return s.strip()

# ────────── 3분 교육대본 생성 ──────────
def make_structured_script(text: str) -> str:
    topic = detect_topic(text)
    core = ai_extract_summary(text, 8)
    if not core:
        return "본문이 충분하지 않아 대본을 생성할 수 없습니다."

    intro, case, risk, act, ask = [], [], [], [], []
    for s in core:
        st = soften(s)
        if any(k in s for k in ["사고", "발생", "사례", "사망"]):
            case.append(st)
        elif any(k in s for k in ["위험", "요인", "원인"]):
            risk.append(st)
        elif any(k in s for k in ["조치", "예방", "착용", "설치", "점검", "휴식", "공급"]):
            act.append(st)
        elif "?" in s or "확인" in s:
            ask.append(st)
        else:
            intro.append(st)

    lines = []
    lines.append(f"🦺 TBM 교육대본 – {topic}\n")
    lines.append("◎ 도입")
    lines.append(f"오늘은 {topic}에 대해 이야기하겠습니다. 현장에서 자주 발생하지만, 예방만으로 충분히 막을 수 있는 부분입니다.\n")

    if case:
        lines.append("◎ 사고 사례")
        for c in case:
            lines.append(f"- {c}")
        lines.append("")
    if risk:
        lines.append("◎ 주요 위험요인")
        for r in risk:
            lines.append(f"- {r}")
        lines.append("")
    if act:
        lines.append("◎ 예방조치 / 실천 수칙")
        for i, a in enumerate(act[:5], 1):
            lines.append(f"{i}️⃣ {a}")
        lines.append("")
    if ask:
        lines.append("◎ 현장 점검 질문")
        for q in ask:
            lines.append(f"- {q}")
        lines.append("")
    lines.append("◎ 마무리 당부")
    lines.append("안전은 한순간의 관심에서 시작됩니다. 오늘 작업 전 서로 한 번 더 확인합시다.")
    lines.append("◎ 구호")
    lines.append("“한 번 더 확인! 한 번 더 점검!”")

    return "\n".join(lines)

# ==========================================================
# 4️⃣ DOCX 저장
# ==========================================================
def to_docx_bytes(script: str) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Malgun Gothic"
    style.font.size = Pt(11)
    for line in script.split("\n"):
        doc.add_paragraph(line)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# ==========================================================
# 5️⃣ Streamlit UI (변경 없음)
# ==========================================================
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 기능(완전 무료)**  
- TextRank + MMR 기반 요약 (비지도 AI)
- 규칙 기반 문체 변환 (생성형 AI 대체)
- OPS 문서에서 자동으로 3분 교육대본 구성
""")

st.title("🦺 OPS2TBM — OPS 문서를 AI로 교육대본 자동 변환 (완전 무료)")

def reset_all():
    st.session_state.pop("manual_text", None)
    st.session_state.pop("edited_text", None)
    st.session_state.uploader_key += 1
    st.rerun()

col1, col2 = st.columns([4, 1])
with col2:
    st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

uploaded = st.file_uploader("OPS 업로드 (PDF 또는 ZIP)", type=["pdf", "zip"])
manual = st.text_area("또는 텍스트 직접 붙여넣기", height=200)
text = ""

if uploaded and uploaded.name.lower().endswith(".pdf"):
    text = read_pdf_text(uploaded.read())
elif manual.strip():
    text = manual

colL, colR = st.columns(2)
with colL:
    st.text_area("📄 추출된 텍스트", text, height=300, key="edited_text")
with colR:
    mode = st.selectbox("🧠 생성 모드", ["TBM 기본(현행)", "자연스러운 교육대본(무료·3분형)"])
    if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
        if not text.strip():
            st.warning("텍스트를 입력하거나 PDF를 업로드하세요.")
        else:
            with st.spinner("AI 요약 + 대본 생성 중..."):
                if mode.startswith("자연스러운"):
                    script = make_structured_script(text)
                else:
                    sents = ai_extract_summary(text, 6)
                    script = "\n".join([f"- {soften(s)}" for s in sents])
            st.success("✅ 대본 생성 완료!")
            st.text_area("결과 미리보기", script, height=420)
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇️ TXT", data=script.encode("utf-8"), file_name="tbm_script.txt")
            with c2:
                st.download_button("⬇️ DOCX", data=to_docx_bytes(script), file_name="tbm_script.docx")
