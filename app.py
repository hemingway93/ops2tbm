# =========================
# OPS2TBM — 안정판 + 경량 AI(TextRank) + MMR 다양성 + 초기화 버튼
# - 텍스트 PDF / 텍스트 붙여넣기 / ZIP(여러 PDF) 지원
# - 템플릿: 자동/사고사례형/가이드형 (수동 선택 가능)
# - AI 요약: 순수 NumPy TextRank (설치 이슈 無) + MMR 다양성
# - 헤더/중복 제거, 동사 우선 선별, 질문형 변환
# - AI로 뽑힌 문장 ⭐[AI] 강조 표시
# - 사이드바: 소개/사용법/시연 멘트
# - 🧹 초기화 버튼(업로더 포함)
# =========================

import io
import zipfile
from typing import List, Tuple, Dict

import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer_high_level import extract_text as pdf_extract_text  # type: ignore
# 일부 환경에서 import 경로 충돌 방지용 별칭 (pdfminer.six)
try:
    from pdfminer.high_level import extract_text as _et
    pdf_extract_text = _et
except Exception:
    pass

import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ----------------------------
# 세션 상태 키 (업로더 초기화용)
# ----------------------------
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# ----------------------------
# 공통 유틸
# ----------------------------
HEADER_HINTS = [
    # 헤더/슬로건으로 자주 등장하는 표현들(본문으로 쓰면 어색)
    "예방조치", "5대 기본수칙", "응급조치", "민감군", "체감온도",
    # 표준 캠페인 슬로건(오탈자였던 '휴휴식식' 대신 표준 표현 반영)
    "물·그늘·휴식", "물, 그늘, 휴식", "물 그늘 휴식",
    "물·바람·휴식", "물, 바람, 휴식", "물 바람 휴식",
    "위기탈출 안전보건 앱", "체감온도 계산기"
]
ACTION_VERBS = [
    "설치", "배치", "착용", "점검", "확인", "측정", "기록", "표시",
    "제공", "비치", "보고", "신고", "교육", "주지", "중지", "통제",
    "지원", "휴식", "휴게", "이동", "후송", "냉각", "공급"
]

def normalize_text(text: str) -> str:
    text = text.replace("\x0c", "\n")
    text = rxx.sub(r"[ \t]+\n", "\n", text)
    text = rxx.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def read_pdf_text(file_bytes: bytes) -> str:
    try:
        with io.BytesIO(file_bytes) as bio:
            text = pdf_extract_text(bio) or ""
    except Exception:
        text = ""
    text = normalize_text(text)

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
    sents = [s.strip(" -•●▪▶▷\t") for s in raw if s and len(s.strip()) > 1]
    # 너무 짧거나 기호성 문장 제거
    sents = [s for s in sents if len(s) >= 6]
    return sents

def simple_tokens(s: str) -> List[str]:
    s = s.lower()
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s)

def has_action_verb(s: str) -> bool:
    return any(v in s for v in ACTION_VERBS) or bool(rxx.search(r"(해야\s*합니다|하십시오|합시다|하세요)", s))

def is_header_like(s: str) -> bool:
    # 동사가 없고 짧거나(<=10자) 슬로건/헤더 패턴이면 헤더로 간주
    if len(s) <= 10 and not has_action_verb(s):
        return True
    if not has_action_verb(s) and any(h in s for h in HEADER_HINTS):
        return True
    if not rxx.search(r"[\.!\?다]$", s) and not has_action_verb(s) and len(s) < 20:
        return True
    return False

def normalize_for_dedup(s: str) -> str:
    s2 = rxx.sub(r"\s+", "", s)
    s2 = rxx.sub(r"(..)\1{1,}", r"\1", s2)  # 2글자 이상 반복 축약
    return s2

# ----------------------------
# 순수 NumPy TextRank + MMR 다양성
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

def textrank_scores(sents: List[str], d: float = 0.85, max_iter: int = 60, tol: float = 1e-4) -> List[float]:
    n = len(sents)
    if n == 0:
        return []
    if n == 1:
        return [1.0]
    X, _ = sentence_tfidf_vectors(sents)
    W = cosine_sim_matrix(X)
    row_sums = W.sum(axis=1, keepdims=True)
    P = np.divide(W, row_sums, out=np.zeros_like(W), where=row_sums > 0)

    r = np.ones((n, 1), dtype=np.float32) / n
    teleport = np.ones((n, 1), dtype=np.float32) / n

    for _ in range(max_iter):
        r_new = d * (P.T @ r) + (1 - d) * teleport
        if np.linalg.norm(r_new - r, ord=1) < tol:
            r = r_new
            break
        r = r_new
    return [float(v) for v in r.flatten()]

def mmr_select(cands: List[str], scores: List[float], X: np.ndarray, k: int, lambda_: float = 0.7) -> List[int]:
    if not cands:
        return []
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

# ----------------------------
# 선택 로직 (규칙/AI) — 헤더 제거/동사 우선/중복 제거 + AI 강조
# ----------------------------
def filter_candidates(sents: List[str]) -> List[str]:
    out = []
    seen = set()
    for s in sents:
        if is_header_like(s):
            continue
        key = normalize_for_dedup(s)
        if key in seen:
            continue
        seen.add(key)
        out.append(s.strip())
    return out

def pick_rule(sents: List[str], keywords: List[str], limit: int) -> Tuple[List[str], List[bool]]:
    base = filter_candidates(sents)
    base = sorted(base, key=lambda x: (not has_action_verb(x), len(x)))  # 동사O 먼저
    hits = [s for s in base if any(k in s for k in keywords)]
    out = hits[:limit]
    flags = [False] * len(out)
    if len(out) < limit:
        remain = [s for s in base if s not in out]
        remain = sorted(remain, key=lambda x: (not has_action_verb(x), len(x)))
        add = remain[: max(0, limit - len(out))]
        out.extend(add)
        flags.extend([False] * len(add))
    return out, flags

def pick_tr(sents: List[str], keywords: List[str], limit: int) -> Tuple[List[str], List[bool]]:
    base = filter_candidates(sents)
    if not base:
        return [], []
    scores = textrank_scores(base)
    scores = np.array(scores, dtype=np.float32)
    if keywords:
        for i, s in enumerate(base):
            if any(k in s for k in keywords):
                scores[i] += 0.2
            if has_action_verb(s):
                scores[i] += 0.1  # 행동 문장 가중
    X, _ = sentence_tfidf_vectors(base)
    idx = mmr_select(base, scores.tolist(), X, limit, lambda_=0.7)
    out = [base[i] for i in idx]
    flags = [True] * len(out)
    return out, flags

def render_with_marks(lines: List[str], ai_flags: List[bool]) -> List[str]:
    out = []
    for s, ai in zip(lines, ai_flags):
        out.append(f"- {'⭐[AI] ' if ai else ''}{s}")
    return out

# ----------------------------
# 템플릿/키워드 + 자동 분류 보정
# ----------------------------
KW_GUIDE_CORE = ["가이드","안내","보호","건강","대응","절차","지침","매뉴얼","예방","상담","지원","존중","민감군"]
KW_GUIDE_STEP = ["절차","순서","방법","점검","확인","보고","조치","기록","휴식","공급","제공","비치"]
KW_GUIDE_QA   = ["질문","왜","어떻게","무엇","주의","확인할", "토의"]

KW_ACC_CORE = ["사고","재해","위험","원인","예방","대책","노후","추락","협착","감전","화재","질식","중독"]
KW_ACC_STEP = ["발생","경위","조치","개선","교육","설치","배치","점검","관리"]

# 가이드형 강한 신호 (자동 분류 보정)
GUIDE_STRONG_HINTS = [
    "물·그늘·휴식", "물, 그늘, 휴식", "물 그늘 휴식",
    "물·바람·휴식", "물, 바람, 휴식", "물 바람 휴식",
    "보냉장구", "응급조치", "민감군", "체감온도", "사업주는"
]

def detect_template(text: str) -> str:
    g_hits = sum(text.count(k) for k in (KW_GUIDE_CORE + KW_GUIDE_STEP))
    a_hits = sum(text.count(k) for k in (KW_ACC_CORE + KW_ACC_STEP))
    g_hits += 3 * sum(text.count(k) for k in GUIDE_STRONG_HINTS)  # 강한 힌트 가중
    return "가이드형" if g_hits >= a_hits else "사고사례형"

# ----------------------------
# 질문형 변환(가이드/사고 공통)
# ----------------------------
def to_question(s: str) -> str:
    s = rxx.sub(r"\s{2,}", " ", s).strip(" -•●▪▶▷").rstrip(" .")
    if has_action_verb(s):
        return f"우리 현장에 '{s}' 하고 있나요?"
    return f"이 항목에 대해 현장 적용이 되었나요? — {s}"

# ----------------------------
# TBM 생성
# ----------------------------
def make_tbm_guide(text: str, use_ai: bool) -> Tuple[str, Dict[str, List[str]]]:
    sents = split_sentences_ko(text)
    if use_ai:
        core, core_f = pick_tr(sents, KW_GUIDE_CORE, 3)
    else:
        core, core_f = pick_rule(sents, KW_GUIDE_CORE, 3)
    steps, steps_f = pick_rule(sents, KW_GUIDE_STEP, 5)
    qa_src, _      = pick_rule(sents, KW_GUIDE_QA + KW_GUIDE_STEP, 3)
    qa = [to_question(x) for x in qa_src]

    parts = {"핵심": core, "절차": steps, "질문": qa}
    lines = []
    lines.append("🦺 TBM 대본 – 가이드형")
    lines.append("")
    lines.append("◎ 오늘의 핵심 포인트")
    lines += render_with_marks(core, core_f)
    lines.append("")
    lines.append("◎ 작업 전 절차/점검")
    lines += render_with_marks(steps, steps_f)
    lines.append("")
    lines.append("◎ 현장 토의 질문")
    for q in qa: lines.append(f"- {q}")
    lines.append("")
    lines.append("◎ 마무리 멘트")
    lines.append("- “오늘 작업의 핵심은 위 세 가지입니다. 다 같이 확인하고 시작합시다.”")
    lines.append("- “잠깐이라도 이상하면 바로 중지하고, 관리자에게 알립니다.”")
    script = "\n".join(lines)
    return script, parts

def make_tbm_accident(text: str, use_ai: bool) -> Tuple[str, Dict[str, List[str]]]:
    sents = split_sentences_ko(text)
    if use_ai:
        core, core_f = pick_tr(sents, KW_ACC_CORE, 3)
    else:
        core, core_f = pick_rule(sents, KW_ACC_CORE, 3)
    steps, steps_f = pick_rule(sents, KW_ACC_STEP, 5)
    qa_src, _      = pick_rule(sents, KW_ACC_STEP, 3)
    qa = [to_question(x) for x in qa_src]

    parts = {"핵심": core, "사고/조치": steps, "질문": qa}
    lines = []
    lines.append("🦺 TBM 대본 – 사고사례형")
    lines.append("")
    lines.append("◎ 사고/위험 요인 요약")
    lines += render_with_marks(core, core_f)
    lines.append("")
    lines.append("◎ 발생 경위/조치/개선")
    lines += render_with_marks(steps, steps_f)
    lines.append("")
    lines.append("◎ 재발 방지 토의 질문")
    for q in qa: lines.append(f"- {q}")
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
# UI (좌/우 2열 + 사이드바 + 초기화)
# ----------------------------
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 기능**  
- 경량 TextRank(NumPy) + **MMR 다양성**으로 핵심 문장을 추출  
- 대본에서 해당 문장을 **⭐[AI]** 로 강조  
- **짧은 헤더/슬로건 제거**, **행동 동사 문장 우선**

**초기화 방법**  
- 우상단 **🧹 초기화** 버튼을 누르면 업로드/선택/텍스트가 모두 리셋됩니다.
""")

st.title("🦺 OPS2TBM — OPS/포스터를 TBM 대본으로 자동 변환")

def reset_all():
    st.session_state.pop("manual_text", None)
    st.session_state.pop("edited_text", None)
    st.session_state.pop("zip_choice", None)
    st.session_state.uploader_key += 1
    st.rerun()

col_top1, col_top2 = st.columns([4,1])
with col_top2:
    st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

st.markdown("""
**안내**  
- 텍스트가 포함된 PDF 또는 본문 텍스트를 권장합니다.  
- ZIP 업로드 시 내부의 PDF들을 자동 인식하여 선택할 수 있습니다.  
- 이미지/스캔형 PDF는 현재 OCR 미지원입니다.
""")

col1, col2 = st.columns([1, 1], gap="large")

# 좌측: 입력
with col1:
    uploaded = st.file_uploader("OPS 업로드 (PDF 또는 ZIP) • 텍스트 PDF 권장",
                                type=["pdf", "zip"], key=f"uploader_{st.session_state.uploader_key}")
    manual_text = st.text_area("또는 OPS 텍스트 직접 붙여넣기", key="manual_text",
                               height=220, placeholder="예: 폭염 시 온열질환 예방을 위해…")

    # ZIP 처리
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

    extracted = ""
    if uploaded and uploaded.name.lower().endswith(".pdf"):
        with st.spinner("PDF 텍스트 추출 중..."):
            data = uploaded.read()
            extracted = read_pdf_text(data)
    elif selected_zip_pdf:
        with st.spinner("ZIP 내부 PDF 텍스트 추출 중..."):
            extracted = read_pdf_text(zip_pdfs[selected_zip_pdf])

    base_text = st.session_state.get("manual_text", "").strip() or extracted.strip()
    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=240, key="edited_text")

# 우측: 옵션/생성/다운로드
with col2:
    use_ai = st.toggle("🔹 AI 요약(TextRank+MMR) 사용", value=True)
    tmpl_choice = st.selectbox("🧩 템플릿", ["자동 선택", "사고사례형", "가이드형"])

    if st.button("🛠️ TBM 대본 생성", type="primary", use_container_width=True):
        if not edited_text.strip():
            st.warning("텍스트가 비어 있습니다. PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
        else:
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

st.caption("현재: 규칙 + NumPy TextRank(경량 AI) + MMR 다양성. 템플릿 자동/수동. ZIP 지원. OCR 미지원(텍스트 PDF 권장).")
