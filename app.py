# ==========================================================
# OPS2TBM — 완전 무료 안정판 v2 (화학물질/중독 품질 개선)
# - UI/레이아웃/컨트롤 이름 그대로 유지
# - 개선 포인트:
#   1) 텍스트 전처리 강화: 불릿/머리말/URL/코드 제거, 줄 병합
#   2) 사건문장(날짜/사망자 표기) 자연어 변환
#   3) 화학물질/중독 특화 키워드로 주제 감지 및 섹션 배치
#   4) 3분형 '자연스러운 교육대본(무료)' 품질 향상
#   5) TBM 기본(현행)도 전처리된 문장으로 요약해서 덜 어색하게
# ==========================================================

import io, zipfile, re
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

# ==========================================================
# 1) 텍스트 전처리(노이즈 제거/불릿 처리/병합)
# ==========================================================
NOISE_PATTERNS = [
    r"^제?\s?\d{4}\s?[-.]?\s?\d+\s?호$",            # 제2024-6 호
    r"^(포스터|책자|스티커|콘텐츠 링크)$",
    r"^(스마트폰\s*APP|중대재해\s*사이렌|산업안전포털|고용노동부)$",
    r"^https?://\S+$",
    r"^\(?\s*PowerPoint\s*프레젠테이션\s*\)?$",
    r"^안전보건자료실.*$",
]

BULLET_PREFIX = r"^[\-\•\●\▪\▶\▷\·\*]+[\s]*"  # 불릿 머리표 제거용
TRAILING_CODES = [
    r"<\s*\d+\s*명\s*사망\s*>", r"<\s*\d+\s*명\s*의식불명\s*>",
]

def normalize_text(t: str) -> str:
    t = t.replace("\x0c", "\n")
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

def strip_noise_line(line: str) -> str:
    s = line.strip()
    if not s:
        return ""
    # 불릿·화살표 제거
    s = re.sub(BULLET_PREFIX, "", s).strip()
    # 노이즈 라인 필터링
    for pat in NOISE_PATTERNS:
        if re.match(pat, s, re.IGNORECASE):
            return ""
    # URL/이상문자 제거
    s = re.sub(r"https?://\S+", "", s).strip()
    # 줄 끝 불필요 기호
    s = s.strip("•●▪▶▷·-").strip()
    return s

def merge_broken_lines(lines: List[str]) -> List[str]:
    """
    OPS는 한 문장을 여러 줄로 분리하는 경우가 많음:
    - 괄호 머리말(예: (호흡보호구)) 다음 줄이 본문인 경우 병합
    - '↳' 서브 항목은 뒤줄과 이어붙여 문장화
    """
    out = []
    buf = ""
    for raw in lines:
        s = strip_noise_line(raw)
        if not s:
            continue
        # '↳' 등 들여쓰기 서브포인트는 이어붙임
        s = s.lstrip("↳").strip()
        if buf and (s[0].islower() or s.startswith(("및 ", "및", "등 ", "등"))):
            buf += " " + s
        elif buf and (buf.endswith((":", "·", "•", "▪", "▶", "▷")) or re.match(r"^\(.*\)$", buf)):
            buf += " " + s
        else:
            if buf:
                out.append(buf)
            buf = s
    if buf:
        out.append(buf)
    return out

def preprocess_text_to_sentences(text: str) -> List[str]:
    text = normalize_text(text)
    # 줄 단위 처리 후 병합
    lines = [ln for ln in text.splitlines() if ln.strip()]
    lines = merge_broken_lines(lines)
    # 문장 분리 (한국어 마침표/영문부호/줄바꿈)
    joined = "\n".join(lines)
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", joined)
    sents = []
    for s in raw:
        s = strip_noise_line(s)
        if not s:
            continue
        # 너무 짧은 조각/단독 번호 제거
        if len(s) < 6:
            continue
        sents.append(s)
    # 중복 제거
    seen = set(); dedup = []
    for s in sents:
        key = re.sub(r"\s+", "", s)
        if key not in seen:
            seen.add(key); dedup.append(s)
    return dedup

# ==========================================================
# 2) PDF 텍스트 추출
# ==========================================================
def read_pdf_text(b: bytes) -> str:
    try:
        with io.BytesIO(b) as bio:
            t = pdf_extract_text(bio) or ""
    except Exception:
        t = ""
    t = normalize_text(t)

    if len(t.strip()) < 10:
        try:
            with io.BytesIO(b) as bio:
                pdf = pdfium.PdfDocument(bio)
                pages = len(pdf)
            if pages > 0 and not t.strip():
                st.warning("⚠️ 이 PDF는 이미지/스캔 기반으로 보입니다. 현재 버전은 OCR 없이 텍스트만 처리합니다.")
        except Exception:
            pass
    return t

# ==========================================================
# 3) AI 요약(TextRank + MMR) — 전처리된 문장 사용
# ==========================================================
def simple_tokens(s: str) -> List[str]:
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

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
    sents = preprocess_text_to_sentences(text)
    if not sents: return []
    X, _ = sentence_tfidf_vectors(sents)
    scores = textrank_scores(sents)
    idx = mmr_select(sents, scores, X, limit, lambda_=0.7)
    return [sents[i] for i in idx]

# ==========================================================
# 4) 말투 완화/사건문 변환/주제 감지/섹션 구성
# ==========================================================
ACTION_VERBS = [
    "설치","배치","착용","점검","확인","측정","기록","표시","제공","비치",
    "보고","신고","교육","주지","중지","통제","휴식","환기","차단","교대","배제","배려","가동"
]

CHEM_KEYS = ["화학물질","중독","유기용제","톨루엔","MSDS","방독마스크","정화통","송기마스크","공기호흡기","국소배기"]
CONFINED_KEYS = ["질식","밀폐","산소결핍","유해가스","황화수소","일산화탄소","메탄"]

def soften(s: str) -> str:
    s = s.replace("하여야", "해야 합니다")
    s = s.replace("한다", "합니다").replace("한다.", "합니다.")
    s = s.replace("바랍니다", "해주세요").replace("확인 바람", "확인해주세요")
    s = s.replace("조치한다", "조치합니다").replace("착용한다", "착용합니다")
    s = s.replace("필요하다", "필요합니다").replace("금지한다", "금지합니다")
    # "(호흡보호구)" 같은 머리말 제거
    s = re.sub(r"^\(([^)]+)\)\s*", "", s)
    return s.strip(" -•●\t")

def detect_topic(t: str) -> str:
    low = t.lower()
    if any(k in t for k in CHEM_KEYS):
        return "화학물질 중독 예방"
    if any(k in t for k in CONFINED_KEYS):
        return "밀폐공간 질식재해 예방"
    if "온열" in t or "폭염" in t:
        return "온열질환 예방"
    if "감전" in t:
        return "감전사고 예방"
    if "지붕" in t or "썬라이트" in t:
        return "지붕 작업 추락사고 예방"
    if "컨베이어" in t or "끼임" in t:
        return "컨베이어 끼임사고 예방"
    return "안전보건 교육"

def to_case_sentence(s: str) -> str:
    """
    '2024.09.26.<1명 사망> 기계식 주차장에서 도장작업 중 톨루엔 중독, 1명 사망 및 2명 의식불명'
    → '2024년 9월 26일, 기계식 주차장 도장작업 중 톨루엔 중독 사고로 1명 사망, 2명 의식불명이 발생했습니다.'
    """
    orig = s
    # 날짜
    m = re.search(r"(\d{4})\.(\d{2})\.(\d{2})", s)
    date_txt = ""
    if m:
        y, mo, d = m.groups()
        date_txt = f"{int(y)}년 {int(mo)}월 {int(d)}일, "
        s = s.replace(m.group(0), "").strip()
    # 꺽쇠 사망자
    for pat in TRAILING_CODES:
        s = re.sub(pat, "", s).strip()
    # 콤마 전처리
    s = s.replace(" ,", ",").replace("  ", " ").strip(" ,.")
    # '중'→ '중에' 자연화
    s = s.replace("중 ", "중 ").replace("중,", "중,")
    # 문장 종결
    if not s.endswith(("다", "다.", ".", "였습니다", "했습니다")):
        s = s.rstrip(" .") + " 사고가 있었습니다."
    return (date_txt + s).strip()

def make_structured_script(text: str, max_points: int = 6) -> str:
    """
    3분형 교육대본 자동 구성:
    - 전처리 → 요약 → 사건/위험/조치/질문 분류
    - 사건 문장 자연어화(to_case_sentence)
    - 도입~구호 구조 조립
    """
    topic = detect_topic(text)
    core = [soften(s) for s in ai_extract_summary(text, max_points)]
    if not core:
        return "본문이 충분하지 않아 대본을 생성할 수 없습니다."

    case, risk, act, ask, misc = [], [], [], [], []
    for s in core:
        if re.search(r"\d{4}\.\d{2}\.\d{2}", s) or any(k in s for k in ["사망", "의식불명", "사고", "발생", "사례"]):
            case.append(to_case_sentence(s))
        elif any(k in s for k in ["위험", "요인", "원인", "중독증상"]):
            risk.append(s)
        elif any(k in s for k in ["조치", "예방", "착용", "설치", "점검", "환기", "비치", "교육"]) or any(v in s for v in ACTION_VERBS):
            act.append(s)
        elif "?" in s or "확인" in s:
            ask.append(s if s.endswith("?") else (s + " 맞습니까?"))
        else:
            misc.append(s)

    # 화학물질/질식 주제일 때 액션 보강: MSDS/국소배기/호흡보호구/환기
    if topic in ["화학물질 중독 예방", "밀폐공간 질식재해 예방"]:
        stock_actions = [
            "작업 전 MSDS를 확인하고 교육받습니다.",
            "국소배기장치를 가동하고 환기 경로를 확보합니다.",
            "송기마스크 또는 공기호흡기 등 적합한 호흡보호구를 착용합니다.",
            "유해가스·산소 농도를 측정하고 기록합니다.",
            "감시자를 배치하고 통신을 유지합니다."
        ]
        # 중복 없이 보강
        for a in stock_actions:
            if a not in act:
                act.append(a)

    # 액션 문장 상한
    act = act[:5]

    # 도입 문구(주제별 자연화)
    topic_intro = {
        "화학물질 중독 예방": "오늘은 유기용제·가스 등 화학물질로 인한 중독을 예방하는 방법을 정리하겠습니다.",
        "밀폐공간 질식재해 예방": "오늘은 탱크·맨홀 등 밀폐공간에서 반복되는 질식재해를 다루겠습니다.",
        "온열질환 예방": "오늘은 폭염 시 온열질환을 막기 위한 핵심을 짚어보겠습니다.",
        "감전사고 예방": "오늘은 감전사고를 예방하기 위한 기본 절차를 확인하겠습니다.",
        "지붕 작업 추락사고 예방": "오늘은 지붕 작업 중 발생하는 추락사고를 막는 방법을 이야기하겠습니다.",
        "컨베이어 끼임사고 예방": "오늘은 컨베이어에서 자주 발생하는 끼임사고를 예방하는 방법을 안내하겠습니다.",
        "안전보건 교육": "오늘은 안전보건 기본을 간단히 정리하겠습니다."
    }
    intro = topic_intro.get(topic, f"오늘은 {topic}에 대해 이야기하겠습니다.").strip()

    # 조립
    lines = []
    lines.append(f"🦺 TBM 교육대본 – {topic}\n")
    lines.append("◎ 도입")
    lines.append(intro + "\n")

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

    # 질문이 하나도 없으면 기본 질문 제공(주제별)
    if not ask:
        if topic == "화학물질 중독 예방":
            ask = [
                "MSDS 비치와 교육이 현장에서 실제로 이루어지고 있습니까?",
                "국소배기장치가 정상 가동되고 있습니까?",
                "호흡보호구(정화통/송기)가 작업에 적합하고 관리가 되고 있습니까?"
            ]
        elif topic == "밀폐공간 질식재해 예방":
            ask = [
                "작업 전 산소/유해가스 농도 측정을 했습니까?",
                "감시자 배치와 통신 수단이 준비되어 있습니까?",
                "환기장치와 출입통제가 작동 중입니까?"
            ]

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
# 5) DOCX 내보내기 (XML 금지문자 필터)
# ==========================================================
_XML_FORBIDDEN = r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
def _xml_safe(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    return rxx.sub(_XML_FORBIDDEN, "", s)

def to_docx_bytes(script: str) -> bytes:
    doc = Document()
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

# ==========================================================
# 6) Streamlit UI (그대로 유지)
# ==========================================================
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 기능(완전 무료)**  
- TextRank + MMR 기반 요약(전처리 강화)  
- 사건문장 자연어화(날짜/사망자 표기 정리)  
- 화학물질/질식 등 주제별 액션 보강

**초기화**  
- 우상단 **🧹 초기화** 버튼
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
            extracted = read_pdf_text(uploaded.read())
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
    gen_mode = st.selectbox("🧠 생성 모드", ["TBM 기본(현행)", "자연스러운 교육대본(무료)"])
    max_points = st.slider("요약 강도(핵심문장 개수)", 3, 10, 6)

    if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
        text_for_gen = (st.session_state.get("edited_text") or "").strip()
        if not text_for_gen:
            st.warning("텍스트가 비어 있습니다. PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
        else:
            # 템플릿 자동/수동(표시용)
            if tmpl_choice == "자동 선택":
                detected = "가이드형" if ("가이드" in text_for_gen or "안내" in text_for_gen) else "사고사례형"
            else:
                detected = tmpl_choice

            with st.spinner("대본 생성 중..."):
                if gen_mode == "자연스러운 교육대본(무료)":
                    script = make_structured_script(text_for_gen, max_points=max_points)
                    subtitle = f"{detected} · 자연스러운 교육대본(무료)"
                else:
                    # TBM 기본(현행)도 전처리된 문장으로 요약해서 품질 개선
                    sents = ai_extract_summary(text_for_gen, max_points if use_ai else 6)
                    sents = [soften(s) for s in sents]
                    script = "\n".join([f"- {s}" for s in sents]) if sents else "텍스트에서 핵심 문장을 찾지 못했습니다."
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

st.caption("완전 무료. 전처리 강화(TextRank+MMR·노이즈 제거) + 사건문장 자연어화 + 주제별 액션 보강. UI 변경 없음. OCR 미지원.")
