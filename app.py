# app.py
import io
from typing import List, Tuple, Dict
import streamlit as st
from docx import Document
from docx.shared import Pt
import pypdfium2 as pdfium
from pdfminer.high_level import extract_text as pdf_extract_text
import numpy as np
import regex as rxx

# ----------------------------
# 전처리 유틸
# ----------------------------
def read_pdf_text(file_bytes: bytes) -> str:
    """텍스트가 포함된 일반 PDF의 텍스트 추출(pdfminer.six 우선).
    이미지 기반 페이지라도 pypdfium2로 1차 확인 후, 텍스트가 없으면 빈 문자열 유지."""
    # pdfminer로 시도 (가장 안정적)
    with io.BytesIO(file_bytes) as bio:
        text = pdf_extract_text(bio) or ""

    # 텍스트가 너무 짧으면(이미지 기반 추정) 페이지 수 정도는 확인
    if len(text.strip()) < 10:
        try:
            with io.BytesIO(file_bytes) as bio:
                pdf = pdfium.PdfDocument(bio)
                pages = len(pdf)
            # 이미지 기반일 가능성 안내
            if pages > 0 and not text.strip():
                st.warning("이 PDF는 이미지 기반으로 보여요. 현재 버전은 OCR 없이 텍스트만 처리합니다.")
        except Exception:
            pass
    return normalize_text(text)

def normalize_text(text: str) -> str:
    # 페이지 넘버/머리글 꼬리글 제거에 가까운 간단 정리
    text = text.replace("\x0c", "\n")
    text = rxx.sub(r"[ \t]+\n", "\n", text)
    text = rxx.sub(r"\n{3,}", "\n\n", text)
    text = text.strip()
    return text

def split_sentences_ko(text: str) -> List[str]:
    # 마침표/물음표/느낌표/종결어미 등을 기준으로 러프하게 분리
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", text)
    sents = [s.strip() for s in raw if s and len(s.strip()) > 1]
    return sents

# ----------------------------
# 순수 NumPy TextRank
# ----------------------------
def sentence_tfidf_vectors(sents: List[str]) -> Tuple[np.ndarray, List[str]]:
    # 아주 단순한 BoW 기반 TF-IDF 흉내 (빈도/문장수 보정)
    tokenized = [simple_tokens(s) for s in sents]
    vocab = {}
    for toks in tokenized:
        for t in toks:
            if t not in vocab:
                vocab[t] = len(vocab)
    if not vocab:
        return np.zeros((len(sents), 0)), []
    mat = np.zeros((len(sents), len(vocab)), dtype=np.float32)
    df = np.zeros((len(vocab),), dtype=np.float32)
    for i, toks in enumerate(tokenized):
        for t in toks:
            j = vocab[t]
            mat[i, j] += 1.0
        unique = set(toks)
        for t in unique:
            df[vocab[t]] += 1.0
    # idf
    N = float(len(sents))
    idf = np.log((N + 1.0) / (df + 1.0)) + 1.0
    mat = mat * idf
    # L2 정규화
    norms = np.linalg.norm(mat, axis=1, keepdims=True) + 1e-8
    mat = mat / norms
    return mat, list(vocab.keys())

def simple_tokens(s: str) -> List[str]:
    # 한글/숫자/영문만 남기고 2글자 이상만 토큰으로 사용
    s = s.lower()
    toks = rxx.findall(r"[가-힣a-z0-9]{2,}", s)
    return toks

def cosine_sim_matrix(X: np.ndarray) -> np.ndarray:
    if X.size == 0:
        return np.zeros((len(X), len(X)), dtype=np.float32)
    # 코사인 유사도 = X @ X^T (이미 L2 normalize 되어 있음)
    sim = np.clip(X @ X.T, 0.0, 1.0)
    np.fill_diagonal(sim, 0.0)  # 자기 자신 제외
    return sim

def textrank_scores(sents: List[str], d: float = 0.85, max_iter: int = 50, tol: float = 1e-4) -> List[float]:
    """SciPy/NetworkX 없이 순수 NumPy로 TextRank 점수 계산."""
    if len(sents) == 0:
        return []
    if len(sents) == 1:
        return [1.0]

    X, _ = sentence_tfidf_vectors(sents)
    W = cosine_sim_matrix(X)
    # 각 행을 확률 분포가 되도록 정규화
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

def pick_sentences_tr(sents: List[str], kw: List[str], limit: int, scores: List[float]) -> List[str]:
    """키워드 가중치 + TextRank 점수로 상위 문장 선택."""
    if not sents:
        return []
    # 키워드 가중
    w = np.array(scores, dtype=np.float32)
    if kw:
        for i, s in enumerate(sents):
            if any(k in s for k in kw):
                w[i] += 0.2
    idx = np.argsort(-w)[:limit]
    return [sents[i] for i in idx]

def pick_sentences_rule(sents: List[str], kw: List[str], limit: int) -> List[str]:
    hits = [s for s in sents if any(k in s for k in kw)]
    if len(hits) >= limit:
        return hits[:limit]
    # 부족하면 길이/위치로 보충
    remain = [s for s in sents if s not in hits]
    remain = sorted(remain, key=lambda x: (-len(x), sents.index(x)))
    return hits + remain[: max(0, limit - len(hits))]

# ----------------------------
# 템플릿 키워드
# ----------------------------
KW_GUIDE_CORE = ["목표", "중요", "중점", "필수", "주의", "준수"]
KW_GUIDE_STEP = ["절차", "순서", "방법", "점검", "확인"]
KW_GUIDE_QA   = ["질문", "왜", "어떻게", "무엇"]

KW_ACC_CORE = ["사고", "재해", "위험", "원인", "예방", "대책"]
KW_ACC_STEP = ["발생", "경위", "조치", "개선", "교육"]
KW_ACC_QA   = ["원인은", "다음에는", "예방하려면", "확인할 점"]

# ----------------------------
# 대본 생성 로직
# ----------------------------
def make_tbm_script_guide(text: str, use_ai: bool) -> Tuple[str, Dict[str, List[str]]]:
    sents = split_sentences_ko(text)
    # 핵심 메시지
    if use_ai:
        scores = textrank_scores(sents)
        core = pick_sentences_tr(sents, KW_GUIDE_CORE, 3, scores)
    else:
        core = pick_sentences_rule(sents, KW_GUIDE_CORE, 3)
    # 절차/점검
    steps = pick_sentences_rule(sents, KW_GUIDE_STEP, 5)
    # 질문
    qa = pick_sentences_rule(sents, KW_GUIDE_QA, 3)

    parts = {"핵심": core, "절차": steps, "질문": qa}
    script = render_script_guide(parts)
    return script, parts

def make_tbm_script_accident(text: str, use_ai: bool) -> Tuple[str, Dict[str, List[str]]]:
    sents = split_sentences_ko(text)
    if use_ai:
        scores = textrank_scores(sents)
        core = pick_sentences_tr(sents, KW_ACC_CORE, 3, scores)
    else:
        core = pick_sentences_rule(sents, KW_ACC_CORE, 3)
    steps = pick_sentences_rule(sents, KW_ACC_STEP, 5)
    qa = pick_sentences_rule(sents, KW_ACC_QA, 3)

    parts = {"핵심": core, "사고/조치": steps, "질문": qa}
    script = render_script_accident(parts)
    return script, parts

def render_script_guide(parts: Dict[str, List[str]]) -> str:
    lines = ["[TBM 대본] 가이드형", ""]
    lines.append("1) 오늘의 핵심 포인트")
    for s in parts["핵심"]:
        lines.append(f" - {s}")
    lines.append("")
    lines.append("2) 작업 전 절차/점검")
    for s in parts["절차"]:
        lines.append(f" - {s}")
    lines.append("")
    lines.append("3) 현장 토의 질문")
    for s in parts["질문"]:
        lines.append(f" - {s}")
    lines.append("")
    lines.append("보너스(시연 멘트):")
    lines.append(" - “오늘 작업의 핵심은 위 세 가지입니다. 다 같이 확인하고 시작합시다.”")
    lines.append(" - “잠깐이라도 이상하면 바로 중지하고, 관리자에게 알립니다.”")
    return "\n".join(lines)

def render_script_accident(parts: Dict[str, List[str]]) -> str:
    lines = ["[TBM 대본] 사고사례형", ""]
    lines.append("1) 사고/위험 요인 요약")
    for s in parts["핵심"]:
        lines.append(f" - {s}")
    lines.append("")
    lines.append("2) 발생 경위/조치/개선")
    for s in parts["사고/조치"]:
        lines.append(f" - {s}")
    lines.append("")
    lines.append("3) 재발 방지 토의 질문")
    for s in parts["질문"]:
        lines.append(f" - {s}")
    lines.append("")
    lines.append("보너스(시연 멘트):")
    lines.append(" - “이 사례에서 배운 예방 포인트를 오늘 작업에 바로 적용합시다.”")
    lines.append(" - “각자 맡은 공정에서 동일 위험이 없는지 다시 점검해 주세요.”")
    return "\n".join(lines)

# ----------------------------
# DOCX 내보내기
# ----------------------------
def export_docx(script: str, filename: str = "tbm_script.docx") -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Malgun Gothic"
    style.font.size = Pt(11)

    for para in script.split("\n"):
        p = doc.add_paragraph(para)
        for run in p.runs:
            run.font.name = "Malgun Gothic"
            run.font.size = Pt(11)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.read()

# ----------------------------
# UI
# ----------------------------
st.set_page_config(page_title="OPS→TBM 대본 생성기", page_icon="🛠️", layout="wide")

st.title("OPS 문서에서 TBM 대본 뽑기 🛠️")
st.caption("텍스트 PDF 기반. 이미지 스캔본은 현재 버전에서 OCR 없이 안내만 합니다.")

col1, col2 = st.columns([1, 1])
with col1:
    file = st.file_uploader("OPS PDF 업로드", type=["pdf"])
    use_ai = st.toggle("AI 요약/가중치 사용(TextRank)", value=True)
    template = st.selectbox("템플릿 선택", ["가이드형", "사고사례형"])
    run_btn = st.button("대본 생성", type="primary")

with col2:
    st.markdown("#### 사용 팁")
    st.markdown("- 텍스트가 많은 템플릿형 OPS 문서에서 가장 성능이 좋습니다.")
    st.markdown("- 이미지/스캔본은 텍스트가 없을 수 있어요 → 변환 안내만 표시.")

result_script = ""
parts = {}

if run_btn and file is not None:
    data = file.read()
    text = read_pdf_text(data)
    if not text.strip():
        st.error("PDF에서 읽을 텍스트가 없습니다. (이미지 기반일 수 있어요)")
    else:
        with st.spinner("문서에서 핵심 내용 추리는 중..."):
            if template == "사고사례형":
                result_script, parts = make_tbm_script_accident(text, use_ai=use_ai)
                subtitle = "사고사례형 템플릿 적용"
            else:
                result_script, parts = make_tbm_script_guide(text, use_ai=use_ai)
                subtitle = "가이드형 템플릿 적용"
        st.success(f"대본 생성 완료! ({subtitle})")

        st.markdown("### 미리보기")
        st.code(result_script, language="markdown")

        docx_bytes = export_docx(result_script, "tbm_script.docx")
        st.download_button(
            "DOCX로 다운로드",
            data=docx_bytes,
            file_name="tbm_script.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

st.divider()
st.markdown("##### 보너스(시연 멘트)만 따로 보기")
st.markdown("- “오늘 작업의 핵심은 위 세 가지입니다. 다 같이 확인하고 시작합시다.”")
st.markdown("- “잠깐이라도 이상하면 바로 중지하고, 관리자에게 알립니다.”")
st.markdown("- “이 사례에서 배운 예방 포인트를 오늘 작업에 바로 적용합시다.”")
st.markdown("- “각자 맡은 공정에서 동일 위험이 없는지 다시 점검해 주세요.”")
