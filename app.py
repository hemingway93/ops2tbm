# ==========================================================
# OPS2TBM — 완전 무료 안정판
# - 기존 UI/레이아웃 그대로 유지
# - AI 요약(TextRank+MMR) + 규칙 기반 문체 변환
# - "자연스러운 교육대본(무료)"에서 3분형 구조(도입~구호) 자동 구성
# - ZIP/PDF/텍스트 입력 모두 지원
# - DOCX 내보내기 특수문자 Sanitizer 적용(이전 ValueError 방지)
# ==========================================================

import io, zipfile
from typing import List, Tuple, Dict
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text
import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ----------------------------
# 세션 상태 (업로더 초기화용 키)
# ----------------------------
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ----------------------------
# 유틸: 텍스트 정리/분리
# ----------------------------
def normalize_text(t: str) -> str:
    t = t.replace("\x0c", "\n")
    t = rxx.sub(r"[ \t]+\n", "\n", t)
    t = rxx.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

def split_sentences_ko(text: str) -> List[str]:
    # 한국어/마침표/줄바꿈 기준 문장 분리
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", text)
    sents = [s.strip(" -•●▪▶▷\t") for s in raw if s and len(s.strip()) > 5]
    return sents

def simple_tokens(s: str) -> List[str]:
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

# ----------------------------
# PDF 텍스트 추출 (텍스트 PDF 권장)
# ----------------------------
def read_pdf_text(b: bytes) -> str:
    try:
        with io.BytesIO(b) as bio:
            t = pdf_extract_text(bio) or ""
    except Exception:
        t = ""
    t = normalize_text(t)

    if len(t.strip()) < 10:
        # 스캔/이미지형 PDF 추정 — 안내만
        try:
            with io.BytesIO(b) as bio:
                pdf = pdfium.PdfDocument(bio)
                pages = len(pdf)
            if pages > 0 and not t.strip():
                st.warning("⚠️ 이 PDF는 이미지/스캔 기반으로 보입니다. 현재 버전은 OCR 없이 텍스트만 처리합니다.")
        except Exception:
            pass
    return t

# ----------------------------
# AI 요약: TextRank + MMR (완전 무료)
# ----------------------------
def sentence_tfidf_vectors(sents: List[str]):
    toks = [simple_tokens(s) for s in sents]
    vocab: Dict[str, int] = {}
    for ts in toks:
        for t in ts:
            if t not in vocab:
                vocab[t] = len(vocab)
    if not vocab:
        return np.zeros((len(sents), 0), dtype=np.float32), []
    M = np.zeros((len(sents), len(vocab)), dtype=np.float32)
    df = np.zeros((len(vocab),), dtype=np.float32)
    for i, ts in enumerate(toks):
        for t in ts:
            M[i, vocab[t]] += 1.0
        for t in set(ts):
            df[vocab[t]] += 1.0
    N = float(len(sents))
    idf = np.log((N + 1.0) / (df + 1.0)) + 1.0
    M *= idf
    M /= (np.linalg.norm(M, axis=1, keepdims=True) + 1e-8)
    return M, list(vocab.keys())

def cosine_sim_matrix(X: np.ndarray) -> np.ndarray:
    if X.size == 0:
        return np.zeros((X.shape[0], X.shape[0]), dtype=np.float32)
    sim = np.clip(X @ X.T, 0.0, 1.0)
    np.fill_diagonal(sim, 0.0)
    return sim

def textrank_scores(sents: List[str], d: float = 0.85, max_iter: int = 60, tol: float = 1e-4) -> List[float]:
    n = len(sents)
    if n == 0: return []
    X, _ = sentence_tfidf_vectors(sents)
    W = cosine_sim_matrix(X)
    row_sums = W.sum(axis=1, keepdims=True)
    P = np.divide(W, row_sums, out=np.zeros_like(W), where=row_sums > 0)
    r = np.ones((n, 1), dtype=np.float32) / n
    teleport = np.ones((n, 1), dtype=np.float32) / n
    for _ in range(max_iter):
        r_new = d * (P.T @ r) + (1 - d) * teleport
        if np.linalg.norm(r_new - r, ord=1) < tol:
            r = r_new; break
        r = r_new
    return [float(v) for v in r.flatten()]

def mmr_select(cands: List[str], scores: List[float], X: np.ndarray, k: int, lambda_: float = 0.7) -> List[int]:
    if not cands: return []
    selected: List[int] = []
    remaining = set(range(len(cands)))
    sim = cosine_sim_matrix(X)
    while remaining and len(selected) < k:
        best_idx, best_val = None, -1e9
        for i in remaining:
            rel = scores[i]
            div = max((sim[i, j] for j in selected), default=0.0)
            mmr = lambda_ * rel - (1 - lambda_) * div
            if mmr > best_val:
                best_val, best_idx = mmr, i
        selected.append(best_idx)
        remaining.remove(best_idx)
    return selected

def ai_extract_summary(text: str, limit: int = 8) -> List[str]:
    sents = split_sentences_ko(text)
    if not sents: return []
    X, _ = sentence_tfidf_vectors(sents)
    scores = textrank_scores(sents)
    idx = mmr_select(sents, scores, X, limit, lambda_=0.7)
    return [sents[i] for i in idx]

# ----------------------------
# 규칙 기반 문체 변환 + 3분형 자동 구성
# ----------------------------
ACTION_VERBS = [
    "설치","배치","착용","점검","확인","측정","기록","표시","제공","비치",
    "보고","신고","교육","주지","중지","통제","휴식","환기","차단","교대","배제","배려"
]

def soften(s: str) -> str:
    # 딱딱한 표현을 말하기 톤으로 완화
    s = s.replace("하여야", "해야 합니다")
    s = s.replace("한다", "합니다").replace("한다.", "합니다.")
    s = s.replace("바랍니다", "해주세요").replace("확인 바람", "확인해주세요")
    s = s.replace("조치한다", "조치합니다").replace("착용한다", "착용합니다")
    s = s.replace("필요하다", "필요합니다").replace("금지한다", "금지합니다")
    return s.strip(" -•●\t")

def detect_topic(t: str) -> str:
    if "온열" in t or "폭염" in t: return "온열질환 예방"
    if "질식" in t or "밀폐" in t: return "질식재해 예방"
    if "감전" in t: return "감전사고 예방"
    if "지붕" in t or "썬라이트" in t: return "지붕 작업 추락사고 예방"
    if "컨베이어" in t or "끼임" in t: return "컨베이어 끼임사고 예방"
    return "안전보건 교육"

def make_structured_script(text: str, max_points: int = 6) -> str:
    """
    3분형 교육대본 자동 구성:
    - AI 요약(TextRank+MMR)으로 핵심 문장 추출
    - '사고/위험/조치/질문' 성격으로 분류
    - 도입~구호 포맷으로 조립
    """
    topic = detect_topic(text)
    core = [soften(s) for s in ai_extract_summary(text, max_points)]
    if not core:
        return "본문이 충분하지 않아 대본을 생성할 수 없습니다."

    case, risk, act, ask, misc = [], [], [], [], []
    for s in core:
        if any(k in s for k in ["사고", "발생", "사례", "사망"]):
            case.append(s)
        elif any(k in s for k in ["위험", "요인", "원인"]):
            risk.append(s)
        elif any(k in s for k in ["조치", "예방"]) or any(v in s for v in ACTION_VERBS):
            act.append(s)
        elif "?" in s or "확인" in s:
            ask.append(s if s.endswith("?") else (s + " 맞습니까?"))
        else:
            misc.append(s)

    # 현장 적용 포인트 부족 시 ACTION 보강
    if len(act) < 3:
        extra = [s for s in split_sentences_ko(text) if any(v in s for v in ACTION_VERBS)]
        for s in extra:
            s2 = soften(s)
            if s2 not in act:
                act.append(s2)
            if len(act) >= 5:
                break
    act = act[:5]

    # 조립
    lines = []
    lines.append(f"🦺 TBM 교육대본 – {topic}\n")
    lines.append("◎ 도입")
    lines.append(f"오늘은 {topic}에 대해 이야기하겠습니다. 현장에서 자주 발생하지만, 올바른 절차만 지키면 예방할 수 있습니다.\n")

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
        for i, a in enumerate(act, 1):
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

# ----------------------------
# DOCX 내보내기 (XML 금지문자 필터)
# ----------------------------
_XML_FORBIDDEN = r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
def _xml_safe(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    return rxx.sub(_XML_FORBIDDEN, "", s)

def to_docx_bytes(script: str) -> bytes:
    doc = Document()
    # 글꼴은 환경에 따라 기본체가 적용될 수 있음(설정 실패시 예외 무시)
    try:
        style = doc.styles["Normal"]
        style.font.name = "Malgun Gothic"
        style.font.size = Pt(11)
    except Exception:
        pass
    for raw in script.split("\n"):
        line = _xml_safe(raw)
        p = doc.add_paragraph(line)
        for run in p.runs:
            try:
                run.font.name = "Malgun Gothic"
                run.font.size = Pt(11)
            except Exception:
                pass
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# ----------------------------
# UI (기존 구성 유지)
# ----------------------------
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 기능(완전 무료)**  
- 경량 TextRank(NumPy) + **MMR 다양성**으로 핵심 문장 추출  
- **자연스러운 교육대본(무료)**: 규칙 기반 문체 변환/3분형 구성 → 담당자가 읽기 좋은 말하기 톤  
- **짧은 헤더/슬로건 제거**, **행동 동사 문장 우선**

**초기화**  
- 우상단 **🧹 초기화** 버튼 → 업로드/선택/텍스트 리셋
""")

st.title("🦺 OPS2TBM — OPS/포스터를 교육 대본으로 자동 변환 (완전 무료)")

def reset_all():
    st.session_state.pop("manual_text", None)
    st.session_state.pop("edited_text", None)
    st.session_state.pop("zip_choice", None)
    st.session_state.uploader_key += 1
    st.rerun()

col_top1, col_top2 = st.columns([4, 1])
with col_top2:
    st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

st.markdown("""
**안내**  
- 텍스트가 포함된 PDF 또는 본문 텍스트를 권장합니다.  
- ZIP 업로드 시 내부의 PDF들을 자동 인식하여 선택할 수 있습니다.  
- 이미지/스캔형 PDF는 현재 OCR 미지원입니다.
""")

col1, col2 = st.columns([1, 1], gap="large")

# 좌측: 입력 (PDF/ZIP/텍스트)
with col1:
    uploaded = st.file_uploader("OPS 업로드 (PDF 또는 ZIP) • 텍스트 PDF 권장",
                                type=["pdf", "zip"], key=f"uploader_{st.session_state.uploader_key}")
    manual_text = st.text_area("또는 OPS 텍스트 직접 붙여넣기", key="manual_text",
                               height=220, placeholder="예: 폭염 시 온열질환 예방을 위해…")

    # ZIP 내부 PDF 선택
    zip_pdfs: Dict[str, bytes] = {}
    selected_zip_pdf = None
    if uploaded and uploaded.name.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(uploaded, "r") as zf:
                for name in zf.namelist():
                    if name.lower().endswith(".pdf"):
                        zip_pdfs[name] = zf.read(name)
        except Exception:
            st.error("ZIP을 해제하는 중 오류가 발생했습니다. 파일을 확인해 주세요.")
        if zip_pdfs:
            selected_zip_pdf = st.selectbox("ZIP 내 PDF 선택", list(zip_pdfs.keys()), key="zip_choice")
        else:
            st.warning("ZIP 안에서 PDF를 찾지 못했습니다.")

    # 텍스트 추출
    extracted = ""
    if uploaded and uploaded.name.lower().endswith(".pdf"):
        with st.spinner("PDF 텍스트 추출 중..."):
            extracted = read_pdf_text(uploaded.read())
    elif selected_zip_pdf:
        with st.spinner("ZIP 내부 PDF 텍스트 추출 중..."):
            extracted = read_pdf_text(zip_pdfs[selected_zip_pdf])

    base_text = st.session_state.get("manual_text", "").strip() or extracted.strip()
    st.markdown("**추출/입력 텍스트 미리보기**")
    # 🔴 중요: 이후 생성 로직은 반드시 이 값(st.session_state['edited_text'])을 사용
    edited_text = st.text_area("텍스트", value=base_text, height=240, key="edited_text")

# 우측: 옵션/생성/다운로드 (UI 유지)
with col2:
    use_ai = st.toggle("🔹 AI 요약(TextRank+MMR) 사용", value=True)
    tmpl_choice = st.selectbox("🧩 템플릿", ["자동 선택", "사고사례형", "가이드형"])
    gen_mode = st.selectbox("🧠 생성 모드", ["TBM 기본(현행)", "자연스러운 교육대본(무료)"])
    max_points = st.slider("요약 강도(핵심문장 개수)", 3, 10, 6)

    if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
        # 🔴 입력은 반드시 세션의 edited_text를 기준으로 사용 (UI와 생성 로직 일관성 보장)
        text_for_gen = (st.session_state.get("edited_text") or "").strip()
        if not text_for_gen:
            st.warning("텍스트가 비어 있습니다. PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
        else:
            # 템플릿 자동/수동 결정 (현재는 3분형 구조와 무관하지만, 기본 모드에서 사용)
            if tmpl_choice == "자동 선택":
                detected = "가이드형" if ("가이드" in text_for_gen or "안내" in text_for_gen) else "사고사례형"
            else:
                detected = tmpl_choice

            with st.spinner("대본 생성 중..."):
                if gen_mode == "자연스러운 교육대본(무료)":
                    # ✅ 3분형 구조 자동 생성 (도입~구호)
                    script = make_structured_script(text_for_gen, max_points=max_points)
                    subtitle = f"{detected} · 자연스러운 교육대본(무료)"
                else:
                    # ✅ TBM 기본(현행): 핵심 문장 나열 (간단/안정)
                    sents = ai_extract_summary(text_for_gen, max_points if use_ai else 6)
                    script = "\n".join([f"- {soften(s)}" for s in sents]) if sents else "텍스트에서 핵심 문장을 찾지 못했습니다."
                    subtitle = f"{detected} · TBM 기본"

            st.success(f"대본 생성 완료! ({subtitle})")
            st.text_area("대본 미리보기", value=script, height=420)

            c3, c4 = st.columns(2)
            with c3:
                st.download_button("⬇️ TXT 다운로드", data=_xml_safe(script).encode("utf-8"),
                                   file_name="tbm_script.txt", use_container_width=True)
            with c4:
                st.download_button("⬇️ DOCX 다운로드", data=to_docx_bytes(script),
                                   file_name="tbm_script.docx", use_container_width=True)

st.caption("현재: 완전 무료. TextRank(경량 AI) + 규칙 기반 NLG. 템플릿 자동/수동. ZIP/PDF/텍스트 지원. OCR 미지원. ‘자연스러운 교육대본(무료)’에서 3분형 구조 자동 생성. DOCX 특수문자 필터 적용.")
