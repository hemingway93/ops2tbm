# ==========================================================
# OPS2TBM — OPS/포스터 → TBM 교육 대본 자동 생성 (완전 무료)
#  * 입력: PDF / ZIP(PDF 묶음) / 텍스트 붙여넣기
#  * 파이프라인: 전처리 → TextRank+MMR 요약(세션 KB 가중) → 규칙기반 리라이팅
#  * 모드:
#      - 핵심요약: 보고서형 단락 요약(불릿 최소, 잡음 제거, 연결어 삽입)
#      - 자연스러운 교육대본(무료): TBM 대본(도입/사례/위험/수칙/질문/마무리)
#  * 시드 KB: 학습파일.zip(개별 24 PDF)에서 추출한 위험유형/행동/질문 내장
#  * 도메인 템플릿: 기본 OFF(토글로 선택). 트리거/유사도/우선순위 3중 안전장치.
#  * OCR 미지원(이미지/스캔 PDF는 경고)
#  * UI: 기존 구조 유지 (좌 입력/미리보기, 우 옵션/생성/다운로드)
# ==========================================================

import io, zipfile, re
from collections import Counter
from typing import List, Dict, Tuple
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer_high_level import extract_text as pdf_extract_text  # streamlit cloud 호환
import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ---------- 페이지 설정 ----------
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

# ==========================================================
# 시드 KB (학습파일.zip 24개 PDF에서 자동 수집한 내용)
#  - 과다한 하드코딩 대신, 최소·대표 패턴만 내장하여 초기 품질을 안정화
# ==========================================================
SEED_RISK_MAP = {
  "중독": "중독", "떨어짐": "떨어짐", "끼임": "끼임", "질식": "질식", "화재": "화재",
  "깔림": "깔림", "맞음": "맞음", "감전": "감전", "지붕": "지붕작업", "예초": "예초",
  "폭발": "폭발", "천공기": "천공", "선반": "절삭", "컨베이어": "협착", "부딪힘": "충돌",
  "미세먼지": "미세먼지", "크레인": "양중", "무너짐": "붕괴", "비계": "비계",
  "추락": "추락", "폭염": "폭염", "벌목": "벌목", "낙하": "낙하", "붕괴": "붕괴", "갱폼": "비계", "발판": "비계"
}

SEED_ACTIONS = [
  "밀폐공간작업 교육 및 훈련 실시",
  "출입 전 충분한 환기 실시",
  "작업 전 가스농도 측정",
  "작업자에 호흡용 보호구 지급",
  "작업 상황 감시자 배치",
  "작업 시 호흡용 보호구 착용",
  "출입·퇴장 인원 점검",
  "보호장구 없이 구조 금지",
  "MSDS 확인 및 유해성 교육 실시",
  "국소배기장치 설치·가동",
  "환기가 불충분한 공간에서는 급기/배기팬 사용",
  "유기화합물 취급 시 방독마스크(갈색 정화통) 착용",
  "송기마스크·공기호흡기 적정 사용",
  "예초기 날 정지 후 이물질 제거·점검",
  "예초·벌목 작업 시 작업자 간 안전거리 유지",
  "벌목 시 쓰러지는 방향 사전 결정 및 대피로 확보",
  "작업발판 견고히 설치 및 상태 점검",
  "개구부·개구창 등 추락 위험 구간에 안전난간 설치",
  "안전대를 안전한 지지점에 연결하고 라이프라인 사용",
  "위험구역 설정·출입통제·감시자 배치",
  "양중 계획 수립 및 신호수 지정·통신 유지",
  "회전체·물림점 방호장치 설치 및 점검",
  "작업 전 작업계획서 작성 및 작업지휘자 지정",
  "개인보호구(안전모·보호안경·안전화 등) 올바르게 착용",
  "화기작업 허가 및 안전점검 철저",
  "정비·청소·점검 작업 시 기계 전원 차단",
]

SEED_QUESTIONS = [
  "작업 전 작업계획서와 위험성평가를 검토했습니까?",
  "개구부·개구창 등 추락 위험 구간에 안전난간을 설치했습니까?",
  "작업발판이 견고하게 설치되고 상태가 양호합니까?",
  "안전대 연결점과 라이프라인이 확보되었습니까?",
  "국소배기장치를 가동하고 환기 경로가 확보되었습니까?",
  "호흡보호구가 작업에 적합하며 관리가 되고 있습니까?",
  "밀폐공간 출입·퇴장 인원 점검이 이루어지고 있습니까?",
  "예초·벌목 작업 시 안전거리와 대피로를 확보했습니까?",
  "양중 작업에 신호수 지정 및 통신체계가 마련되었습니까?",
  "회전체·물림점 방호장치가 정상 동작합니까?"
]

# ---------- 세션 상태 ----------
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "kb_terms" not in st.session_state:
    st.session_state.kb_terms: Counter = Counter()
if "kb_actions" not in st.session_state:
    st.session_state.kb_actions: List[str] = []
if "kb_questions" not in st.session_state:
    st.session_state.kb_questions: List[str] = []
if "domain_toggle" not in st.session_state:
    st.session_state.domain_toggle = False  # 도메인 템플릿 기본 Off
if "seed_loaded" not in st.session_state:
    st.session_state.seed_loaded = False

# ==========================================================
# 전처리 유틸 / 패턴
# ==========================================================
NOISE_PATTERNS = [
    r"^제?\s?\d{4}\s?[-.]?\s?\d+\s?호$",
    r"^(동절기\s*주요사고|안전작업방법|콘텐츠링크|책자\s*OPS|숏폼\s*OPS)$",
    r"^(포스터|책자|스티커|콘텐츠 링크)$",
    r"^(스마트폰\s*APP|중대재해\s*사이렌|산업안전포털|고용노동부)$",
    r"^https?://\S+$",
    r"^\(?\s*PowerPoint\s*프레젠테이션\s*\)?$",
    r"^안전보건자료실.*$",
    r"^배포처\s+.*$",
    r"^홈페이지\s+.*$",
    r"^주소\s+.*$",
    r"^VR\s+.*$",
    r"^리플릿\s+.*$",
    r"^동영상\s+.*$",
    r"^APP\s+.*$",
    r".*검색해\s*보세요.*$",
]
BULLET_PREFIX = r"^[\s\-\•\●\▪\▶\▷\·\*\u25CF\u25A0\u25B6\u25C6\u2022\u00B7\u279C\u27A4\u25BA\u25AA\u25AB\u2611\u2713\u2714\u2716\u2794\u27A2\uF0FC\uF0A7]+"
DATE_PAT = r"([’']?\d{2,4})\.\s?(\d{1,2})\.\s?(\d{1,2})\.?"
META_PATTERNS = [
    r"<\s*\d+\s*명\s*사망\s*>",
    r"<\s*\d+\s*명\s*사상\s*>",
    r"<\s*\d+\s*명\s*의식불명\s*>",
    r"<\s*사망\s*\d+\s*명\s*>",
    r"<\s*사상\s*\d+\s*명\s*>",
]
STOP_TERMS = set("""
및 등 관련 사항 내용 예방 안전 작업 현장 교육 방법 기준 조치
실시 확인 필요 경우 대상 사용 관리 점검 적용 정도 주의 중 전 후
주요 사례 안전작업방법 포스터 동영상 리플릿 가이드 자료실 검색
키메세지 교육혁신실 안전보건공단 공단 자료 구독 안내 연락 참고 출처
소재 소재지 위치 장소 지역 시군구 서울 인천 부산 대구 대전 광주 울산 세종 경기도 충청 전라 경상 강원 제주
명 건 호 호차 호수 페이지 쪽 부록 참고 그림 표 목차
안전보건 ops 키 메세지 키메세지 자료 ops교안 교안
""".split())
LABEL_DROP_PAT = [
    r"^\d+$", r"^\d{2,4}[-_]\d{1,}$", r"^\d{4}$",
    r"^(제)?\d+호$", r"^(호|호수|호차)$",
    r"^(사업장|업체|소재|소재지|장소|지역)$",
    r"^\d+\s*(명|건)$",
]

# 위험유형(초기값) — 시드 맵으로 초기화 + 동적 보강
RISK_KEYWORDS = dict(SEED_RISK_MAP)

# ---------- 공통 유틸 ----------
def normalize_text(t: str) -> str:
    t = t.replace("\x0c", "\n")
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

def strip_noise_line(line: str) -> str:
    s = (line or "").strip()
    if not s: return ""
    s = re.sub(BULLET_PREFIX, "", s).strip()
    for pat in NOISE_PATTERNS:
        if re.match(pat, s, re.IGNORECASE):
            return ""
    s = re.sub(r"https?://\S+", "", s).strip()
    s = s.strip("•●▪▶▷·-—–")
    return s

def merge_broken_lines(lines: List[str]) -> List[str]:
    out, buf = [], ""
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
    out = []; i = 0
    while i < len(lines):
        cur = strip_noise_line(lines[i])
        if re.search(DATE_PAT, cur) and (i + 1) < len(lines):
            nxt = strip_noise_line(lines[i + 1])
            if re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|무너짐|낙하)", nxt):
                m = re.search(DATE_PAT, cur); y, mo, d = m.groups()
                y = int(str(y).replace("’","").replace("'",""))
                y = 2000 + y if y < 100 else y
                out.append(f"{int(y)}년 {int(mo)}월 {int(d)}일, {nxt}")
                i += 2; continue
        out.append(cur); i += 1
    return out

def preprocess_text_to_sentences(text: str) -> List[str]:
    text = normalize_text(text)
    raw_lines = [ln for ln in text.splitlines() if ln.strip()]
    lines = merge_broken_lines(raw_lines)
    lines = combine_date_with_next(lines)
    joined = "\n".join(lines)
    raw = rxx.split(r"(?<=[\.!\?]|다\.)\s+|\n+", joined)
    sents = []
    for s in raw:
        s2 = strip_noise_line(s)
        if not s2: continue
        if re.search(r"(주요사고|안전작업방법|콘텐츠링크|주요 사고개요)$", s2):
            continue
        if len(s2) < 6: continue
        sents.append(s2)
    # 중복 제거
    seen, dedup = set(), []
    for s in sents:
        k = re.sub(r"\s+", "", s)
        if k not in seen:
            seen.add(k); dedup.append(s)
    return dedup

# ==========================================================
# PDF 처리
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
                if len(pdf) > 0 and not t.strip():
                    st.warning("⚠️ 이미지/스캔 PDF로 보입니다. 현재 OCR 미지원.")
        except Exception:
            pass
    return t

# ==========================================================
# 요약(TextRank+MMR) — 세션 KB 가중
# ==========================================================
def tokens(s: str) -> List[str]:
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

def sentence_tfidf_vectors(sents: List[str], kb_boost: Dict[str, float] = None) -> Tuple[np.ndarray, List[str]]:
    toks = [tokens(s) for s in sents]
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
            w = 1.0
            if kb_boost and t in kb_boost:
                w *= kb_boost[t]
            M[i, vocab[t]] += w
        for t in set(ts):
            df[vocab[t]] += 1.0
    N = float(len(sents))
    idf = np.log((N + 1.0) / (df + 1.0)) + 1.0
    M *= idf
    if kb_boost:
        for t, idx in vocab.items():
            if t in kb_boost:
                M[:, idx] *= (1.0 + 0.2 * kb_boost[t])
    M /= (np.linalg.norm(M, axis=1, keepdims=True) + 1e-8)
    return M, list(vocab.keys())

def cosim(X: np.ndarray) -> np.ndarray:
    if X.size == 0:
        return np.zeros((X.shape[0], X.shape[0]), dtype=np.float32)
    S = np.clip(X @ X.T, 0.0, 1.0)
    np.fill_diagonal(S, 0.0)
    return S

def textrank_scores(sents: List[str], X: np.ndarray, d: float = 0.85, max_iter: int = 60, tol: float = 1e-4) -> List[float]:
    n = len(sents)
    if n == 0: return []
    W = cosim(X)
    row = W.sum(axis=1, keepdims=True)
    P = np.divide(W, row, out=np.zeros_like(W), where=row > 0)
    r = np.ones((n, 1), dtype=np.float32) / n
    tel = np.ones((n, 1), dtype=np.float32) / n
    for _ in range(max_iter):
        r2 = d * (P.T @ r) + (1 - d) * tel
        if np.linalg.norm(r2 - r, 1) < tol:
            r = r2; break
        r = r2
    return [float(v) for v in r.flatten()]

def mmr_select(sents: List[str], scores: List[float], X: np.ndarray, k: int, lam: float = 0.7) -> List[int]:
    S = cosim(X); sel: List[int] = []; rem = set(range(len(sents)))
    while rem and len(sel) < k:
        best, val = None, -1e9
        for i in rem:
            rel = scores[i]
            div = max((S[i, j] for j in sel), default=0.0)
            sc = lam * rel - (1 - lam) * div
            if sc > val: val, best = sc, i
        sel.append(best)  # type: ignore
        rem.remove(best)  # type: ignore
    return sel

def ai_extract_summary(text: str, limit: int = 8) -> List[str]:
    sents = preprocess_text_to_sentences(text)
    if not sents: return []
    kb = st.session_state.kb_terms
    total = sum(kb.values()) or 1
    kb_boost = {t: 1.0 + (cnt / total) * 3.0 for t, cnt in kb.items()} if kb else None
    X, _ = sentence_tfidf_vectors(sents, kb_boost=kb_boost)
    scores = textrank_scores(sents, X)
    idx = mmr_select(sents, scores, X, limit, lam=0.7)
    return [sents[i] for i in idx]

# ==========================================================
# 도메인 템플릿(안전장치 적용)
# ==========================================================
DOMAIN_TEMPLATES = [
    ({"비계","발판","갱폼","추락"}, "작업발판을 견고하게 설치하고 안전난간 및 추락방호망을 확보합니다."),
    ({"안전난간","난간","개구부"}, "개구부·개구창 등 추락 위험 구간에 안전난간을 설치합니다."),
    ({"안전대","라이프라인","벨트"}, "안전대를 안전한 지지점에 연결하고 라이프라인을 사용합니다."),
    ({"MSDS","국소배기","환기"}, "취급 물질의 MSDS를 확인하고 국소배기장치를 가동하여 충분히 환기합니다."),
    ({"예초","벌목","예초기"}, "예초·벌목 작업 시 작업자 간 안전거리를 유지하고 대피로를 확보합니다."),
    ({"크레인","양중"}, "양중 계획을 수립하고 신호수를 지정하여 통신을 유지합니다."),
    ({"컨베이어","협착","회전체"}, "회전체·물림점 접촉을 방지하도록 방호장치를 설치하고 점검합니다."),
]

def jaccard(a: set, b: set) -> float:
    return len(a & b) / (len(a | b) + 1e-8)

# ==========================================================
# 라벨/분류/리라이팅
# ==========================================================
ACTION_VERBS = [
    "설치","배치","착용","점검","확인","측정","기록","표시","제공","비치",
    "보고","신고","교육","주지","중지","통제","휴식","환기","차단","교대","배제","배려",
    "가동","준수","운영","유지","교체","정비","청소","고정","격리","보호","보수","작성","지정"
]
ACTION_PAT = (
    r"(?P<obj>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*"
    r"(?P<verb>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"
    r"|(?P<obj2>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(을|를)\s*"
    r"(?P<verb2>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"
)

def drop_label_token(t: str) -> bool:
    if t in STOP_TERMS: return True
    for pat in LABEL_DROP_PAT:
        if re.match(pat, t): return True
    if t in {"소재","소재지","지역","장소","버스","영업소","업체","자료","키","메세지","명","안전보건"}:
        return True
    return False

def top_terms_for_label(text: str, k: int = 3) -> List[str]:
    doc_cnt = Counter([t for t in tokens(text) if not drop_label_token(t)])
    bonus = Counter()
    for t in list(doc_cnt.keys()):
        if t in RISK_KEYWORDS:
            bonus[RISK_KEYWORDS[t]] += doc_cnt[t]
    doc_cnt += bonus
    kb = st.session_state.kb_terms
    if kb:
        for t, c in kb.items():
            if not drop_label_token(t):
                doc_cnt[t] += 0.2 * c
    if not doc_cnt: return ["안전보건", "교육"]
    commons = {"안전","교육","작업","현장","예방","조치","확인","관리","점검","가이드","지침"}
    cand = [(t, doc_cnt[t]) for t in doc_cnt if t not in commons and len(t) >= 2]
    if not cand: cand = list(doc_cnt.items())
    cand.sort(key=lambda x: x[1], reverse=True)
    return [t for t, _ in cand[:k]]

def dynamic_topic_label(text: str) -> str:
    terms = top_terms_for_label(text, k=3)
    risks = [RISK_KEYWORDS.get(t, t) for t in terms if t in RISK_KEYWORDS or t in RISK_KEYWORDS.values()]
    extra = [t for t in terms if t not in risks]
    label_core = " ".join(sorted(set(risks), key=risks.index)) or "안전보건"
    tail = " ".join(extra[:1])
    label = (label_core + (" " + tail if tail else "")).strip()
    if "예방" not in label:
        label += " 재해예방"
    return label

def soften(s: str) -> str:
    s = s.replace("하여야", "해야 합니다").replace("한다", "합니다").replace("한다.", "합니다.")
    s = s.replace("바랍니다", "해주세요").replace("확인 바람", "확인해주세요")
    s = s.replace("금지한다", "금지합니다").replace("필요하다", "필요합니다")
    s = re.sub(r"^\(([^)]+)\)\s*", "", s)
    for pat in META_PATTERNS:
        s = re.sub(pat, "", s).strip()
    s = re.sub(BULLET_PREFIX, "", s).strip(" -•●\t")
    return s

def is_accident_sentence(s: str) -> bool:
    if any(w in s for w in ["예방", "대책", "지침", "수칙"]):
        return False
    return bool(re.search(DATE_PAT, s) or re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|무너짐|낙하)", s))

def is_prevention_sentence(s: str) -> bool:
    return any(w in s for w in ["예방", "대책", "지침", "수칙", "안전조치"])

def is_risk_sentence(s: str) -> bool:
    return any(w in s for w in ["위험", "요인", "원인", "증상", "결빙", "강풍", "폭염", "미세먼지", "회전체", "비산", "말림", "추락", "낙하", "협착"])

def naturalize_case_sentence(s: str) -> str:
    s = soften(s)
    death = re.search(r"사망\s*(\d+)\s*명", s)
    inj = re.search(r"사상\s*(\d+)\s*명", s)
    unconscious = re.search(r"의식불명", s)
    info = []
    if death: info.append(f"근로자 {death.group(1)}명 사망")
    if inj and not death: info.append(f"{inj.group(1)}명 사상")
    if unconscious: info.append("의식불명 발생")
    m = re.search(DATE_PAT, s)
    date_txt = ""
    if m:
        y, mo, d = m.groups()
        y = int(str(y).replace("’","").replace("'",""))
        y = 2000 + y if y < 100 else y
        date_txt = f"{int(y)}년 {int(mo)}월 {int(d)}일, "
        s = s.replace(m.group(0), "").strip()
    s = s.strip(" ,.-")
    if not re.search(r"(사망|사상|사고|중독|붕괴|질식|추락|깔림|부딪힘|무너짐|낙하)", s):
        if not re.search(r"(다\.|입니다\.|했습니다\.)$", s):
            s = s.rstrip(" .") + " 사고가 발생했습니다."
    tail = ""
    if info and not s.endswith("습니다."):
        tail = " " + (", ".join(info)) + "했습니다."
    return (date_txt + s + tail).strip()

# 도메인 템플릿(신중 적용)
def _domain_template_apply(s: str, base_text: str) -> str:
    if not st.session_state.get("domain_toggle"):
        return s
    sent_toks = set(tokens(s))
    base_toks = set(tokens(base_text))
    # 유사도 기준(과적용 방지)
    if jaccard(sent_toks, base_toks) < 0.05:
        return s
    best = None; best_hits = 0
    for triggers, render in DOMAIN_TEMPLATES:
        if (sent_toks & triggers) and (base_toks & triggers):
            hits = len((sent_toks | base_toks) & triggers)
            if hits > best_hits:
                best_hits = hits; best = render
    return best if best else s

def to_action_sentence(s: str, base_text: str) -> str:
    s2 = soften(s)
    # 홍보성/의미 불명 구절 제거
    s2 = re.sub(r"(위기탈출\s*안전보건)", "", s2).strip()
    # '에 따른/따라' 교정
    s2 = re.sub(r"\s*에\s*따른\s*", " 시 ", s2)
    s2 = re.sub(r"\s*에\s*따라\s*", " 시 ", s2)
    # 도메인 템플릿(안전장치) 적용
    s2_tpl = _domain_template_apply(s2, base_text)
    if s2_tpl != s2:
        return s2_tpl if s2_tpl.endswith(("다.", "습니다.", "합니다.")) else (s2_tpl.rstrip(" .") + " 합니다.")
    # 일반 패턴
    m = re.search(ACTION_PAT, s2)
    if not m:
        return s2 if s2.endswith(("니다.", "합니다.", "다.")) else (s2.rstrip(" .") + " 합니다.")
    obj = (m.group("obj") or m.group("obj2") or "").strip()
    verb = (m.group("verb") or m.group("verb2") or "실시").strip()
    prefix = "반드시 " if "설치" in verb else ("작업 전 " if verb in ("확인","점검","측정","기록","작성","지정") else "")
    if obj and not re.search(r"(을|를|에|에서|과|와|의)$", obj):
        obj += "를"
    core = f"{prefix}{obj} {verb}".strip()
    core = re.sub(r"\s+", " ", core)
    return (core + "합니다.").replace("  ", " ")

def classify_sentence(s: str) -> str:
    if is_accident_sentence(s): return "case"
    if re.search(ACTION_PAT, s) or is_prevention_sentence(s): return "action"
    if is_risk_sentence(s): return "risk"
    if "?" in s or "확인" in s or "점검" in s: return "question"
    return "other"

# ==========================================================
# 세션 KB 누적/활용 (+ 시드 KB 로드)
# ==========================================================
def seed_kb_once():
    if not st.session_state.seed_loaded:
        # 위험유형 시드
        for t, k in SEED_RISK_MAP.items():
            if t not in RISK_KEYWORDS:
                RISK_KEYWORDS[t] = k
        # 행동/질문 시드
        for a in SEED_ACTIONS:
            if 2 <= len(a) <= 160:
                st.session_state.kb_actions.append(a if a.endswith(("다","다.","합니다","합니다.")) else a + " 합니다.")
        for q in SEED_QUESTIONS:
            st.session_state.kb_questions.append(q if q.endswith("?") else q + "?")
        # 용어 시드(위험유형 가중)
        for t in SEED_RISK_MAP.keys():
            st.session_state.kb_terms[t] += 5
        st.session_state.seed_loaded = True

def kb_ingest_text(text: str) -> None:
    if not (text or "").strip(): return
    sents = preprocess_text_to_sentences(text)
    for s in sents:
        for t in tokens(s):
            if len(t) >= 2:
                st.session_state.kb_terms[t] += 1
                if re.search(r"(추락|낙하|깔림|끼임|중독|질식|화재|폭발|감전|폭염|붕괴|비계|갱폼|예초|벌목|컨베이어|크레인|지붕|선반|천공)", t):
                    if t not in RISK_KEYWORDS:
                        RISK_KEYWORDS[t] = t
    for s in sents:
        if re.search(ACTION_PAT, s) or is_prevention_sentence(s):
            cand = to_action_sentence(s, text)
            if 2 <= len(cand) <= 160:
                st.session_state.kb_actions.append(cand)
    for s in sents:
        if "?" in s or "확인" in s or "점검" in s:
            q = soften(s if s.endswith("?") else s + " 맞습니까?")
            if 2 <= len(q) <= 140:
                st.session_state.kb_questions.append(q)

def kb_prune() -> None:
    def dedup_keep_order(lst: List[str]) -> List[str]:
        seen, out = set(), []
        for x in lst:
            k = re.sub(r"\s+", "", x)
            if k not in seen:
                seen.add(k); out.append(x)
        return out
    st.session_state.kb_actions = dedup_keep_order(st.session_state.kb_actions)[:900]
    st.session_state.kb_questions = dedup_keep_order(st.session_state.kb_questions)[:500]
    st.session_state.kb_terms = Counter(dict(st.session_state.kb_terms.most_common(1800)))

def kb_match_candidates(cands: List[str], base_text: str, limit: int) -> List[str]:
    bt = set(tokens(base_text))
    scored: List[Tuple[float, str]] = []
    for c in cands:
        ct = set(tokens(c))
        j = len(bt & ct) / (len(bt | ct) + 1e-8)
        if j > 0:
            scored.append((j, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [c for _, c in scored[:limit]]

# ==========================================================
# 생성 모드 1: 자연스러운 교육대본(무료)
# ==========================================================
def make_structured_script(text: str, max_points: int = 6) -> str:
    topic_label = dynamic_topic_label(text)
    core = [soften(s) for s in ai_extract_summary(text, max_points)]
    if not core:
        return "본문이 충분하지 않아 대본을 생성할 수 없습니다."

    case, risk, act, ask = [], [], [], []
    for s in core:
        c = classify_sentence(s)
        if c == "case":
            case.append(naturalize_case_sentence(s))
        elif c == "action":
            act.append(to_action_sentence(s, text))
        elif c == "risk":
            risk.append(soften(s))
        elif c == "question":
            ask.append(soften(s if s.endswith("?") else s + " 맞습니까?"))

    if len(act) < 5 and st.session_state.kb_actions:
        act += kb_match_candidates(st.session_state.kb_actions, text, 5 - len(act))
    act = act[:5]

    if not ask and st.session_state.kb_questions:
        ask = kb_match_candidates(st.session_state.kb_questions, text, 3)
    if not ask:
        ask = ["필요한 안전조치가 오늘 작업 범위에 맞게 준비되어 있습니까?"]

    lines = []
    lines.append(f"🦺 TBM 교육대본 – {topic_label}\n")
    lines.append("◎ 도입")
    lines.append(f"오늘은 최근 발생한 '{topic_label.replace(' 재해예방','')}' 사고 사례를 중심으로, 우리 현장에서 같은 사고를 예방하기 위한 안전조치를 함께 살펴보겠습니다.\n")

    if case:
        lines.append("◎ 사고 사례")
        for c in case: lines.append(f"- {c}")
        lines.append("")

    if risk:
        lines.append("◎ 주요 위험요인")
        for r in risk: lines.append(f"- {r}")
        lines.append("")

    if act:
        lines.append("◎ 예방조치 / 실천 수칙")
        for i, a in enumerate(act, 1): lines.append(f"{i}️⃣ {a}")
        lines.append("")

    if ask:
        lines.append("◎ 현장 점검 질문")
        for q in ask: lines.append(f"- {q}")
        lines.append("")

    lines.append("◎ 마무리 당부")
    lines.append("예방조치는 '선조치 후작업'이 원칙입니다. 오늘 작업 전, 각 공정별 위험요인을 다시 한 번 점검하고 필요한 보호구와 안전조치를 반드시 준비합시다.")
    lines.append("◎ 구호")
    lines.append("“한 번 더 확인! 한 번 더 점검!”")

    return "\n".join(lines)

# ==========================================================
# 생성 모드 2: 핵심요약 (보고서형)
# ==========================================================
def make_concise_report(text: str, max_points: int = 6) -> str:
    sents = ai_extract_summary(text, max_points)
    sents = [soften(s) for s in sents if not re.match(r"(배포처|주소|홈페이지|VR|리플릿)", s)]
    if not sents:
        return "텍스트에서 핵심을 요약할 수 없습니다."

    cases = [naturalize_case_sentence(s) for s in sents if is_accident_sentence(s)]
    risks  = [soften(s) for s in sents if (not is_accident_sentence(s)) and is_risk_sentence(s)]
    acts   = [to_action_sentence(s, text) for s in sents if (not is_accident_sentence(s)) and (is_prevention_sentence(s) or re.search(ACTION_PAT, s))]

    def uniq_keep(seq: List[str]) -> List[str]:
        seen, out = set(), []
        for x in seq:
            k = re.sub(r"\s+", "", x)
            if k not in seen:
                seen.add(k); out.append(x)
        return out

    cases = uniq_keep(cases)[:3]
    risks  = uniq_keep(risks)[:3]
    acts   = uniq_keep(acts)[:4]

    topic = dynamic_topic_label(text)

    lines = [f"📄 핵심요약 — {topic}\n"]
    if cases:
        lines.append("【사고 개요】")
        lines.append("최근 자료에서 다음과 같은 사고가 확인되었습니다.")
        for c in cases:
            lines.append(f"- {c}")
        lines.append("")
    if risks:
        lines.append("【주요 위험요인】")
        lines.append("자료 전반에서 다음 요인이 반복적으로 나타났습니다.")
        for r in risks:
            lines.append(f"- {r}")
        lines.append("")
    if acts:
        lines.append("【예방/실천 요약】")
        lines.append("현장에서 즉시 적용 가능한 핵심 수칙입니다.")
        for a in acts:
            lines.append(f"- {a}")
        lines.append("")
    if not (cases or risks or acts):
        lines.append("자료의 핵심을 간단히 정리하면 다음과 같습니다.")
        for s in sents:
            lines.append(f"- {s}")
    return "\n".join(lines)

# ==========================================================
# DOCX 내보내기 (XML 안전필터)
# ==========================================================
_XML_FORBIDDEN = r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
def _xml_safe(s: str) -> str:
    if not isinstance(s, str): s = "" if s is None else str(s)
    return rxx.sub(_XML_FORBIDDEN, "", s)

def to_docx_bytes(script: str) -> bytes:
    doc = Document()
    try:
        style = doc.styles["Normal"]; style.font.name = "Malgun Gothic"; style.font.size = Pt(11)
    except Exception:
        pass
    for raw in script.split("\n"):
        line = _xml_safe(raw)
        p = doc.add_paragraph(line)
        for run in p.runs:
            try:
                run.font.name = "Malgun Gothic"; run.font.size = Pt(11)
            except Exception:
                pass
    bio = io.BytesIO(); doc.save(bio); bio.seek(0); return bio.read()

# ==========================================================
# UI (기존 틀 유지 / 모드명만 변경 + 시드 로딩)
# ==========================================================
with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 파이프라인(완전 무료)**  
- 전처리 → TextRank+MMR 요약(세션 KB 가중) → 데이터 기반 리라이팅  
- PDF/ZIP/텍스트 업로드 시 즉시 누적 학습(용어/행동/질문).
- 이미지/스캔 PDF는 텍스트 추출이 어려울 수 있습니다.
""")
    # 🔧 도메인 템플릿 옵션 (기본 Off)
    st.session_state.domain_toggle = st.toggle("🔧 도메인 템플릿 강화(신중 적용)", value=False,
                                              help="문장·본문 트리거 일치 + 유사도 기준 충족 시에만 템플릿을 적용합니다. 새로운 파일엔 보수적으로 동작합니다.")

# 시드 KB 최초 1회 로딩
seed_kb_once()

st.title("🦺 OPS/포스터를 교육 대본으로 자동 변환 (완전 무료)")

def reset_all():
    st.session_state.pop("manual_text", None)
    st.session_state.pop("edited_text", None)
    st.session_state.pop("zip_choice", None)
    st.session_state.kb_terms = Counter()
    st.session_state.kb_actions = []
    st.session_state.kb_questions = []
    st.session_state.uploader_key += 1
    st.session_state.seed_loaded = False
    st.rerun()

col_top1, col_top2 = st.columns([4, 1])
with col_top2:
    st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

st.markdown("""
**안내**  
- 텍스트가 포함된 PDF 또는 본문 텍스트를 권장합니다.  
- 이미지/스캔 PDF는 현재 OCR 미지원입니다.
""")

col1, col2 = st.columns([1, 1], gap="large")

# ---------- 좌측 입력/미리보기 ----------
with col1:
    uploaded = st.file_uploader(
        "OPS 업로드 (PDF 또는 ZIP) • 텍스트 PDF 권장",
        type=["pdf", "zip"],
        key=f"uploader_{st.session_state.uploader_key}"
    )
    manual_text = st.text_area(
        "또는 OPS 텍스트 직접 붙여넣기",
        key="manual_text",
        height=220,
        placeholder="예: 현장 안내문 또는 OPS 본문 텍스트…"
    )

    extracted: str = ""
    zip_pdfs: Dict[str, bytes] = {}
    selected_zip_pdf: str = ""

    if uploaded is not None:
        fname = (uploaded.name or "").lower()
        if fname.endswith(".zip"):
            try:
                with zipfile.ZipFile(uploaded, "r") as zf:
                    for name in zf.namelist():
                        if name.lower().endswith(".pdf"):
                            data = zf.read(name)
                            zip_pdfs[name] = data
                            txt = read_pdf_text(data)
                            if txt.strip():
                                kb_ingest_text(txt)  # ZIP 전체 학습
                kb_prune()
            except Exception as e:
                st.error(f"ZIP 해제 오류: {e}")
            if zip_pdfs:
                selected_zip_pdf = st.selectbox("ZIP 내 PDF 선택", list(zip_pdfs.keys()), key="zip_choice")
                if selected_zip_pdf:
                    with st.spinner("ZIP 내부 PDF 텍스트 추출 중..."):
                        extracted = read_pdf_text(zip_pdfs[selected_zip_pdf])
        elif fname.endswith(".pdf"):
            with st.spinner("PDF 텍스트 추출 중..."):
                data = uploaded.read()
                extracted = read_pdf_text(data)
                if extracted.strip():
                    kb_ingest_text(extracted); kb_prune()
                else:
                    st.warning("⚠️ PDF에서 유효한 텍스트를 추출할 수 없습니다.")
        else:
            st.warning("지원하지 않는 형식입니다. PDF 또는 ZIP을 업로드하세요.")

    pasted = (manual_text or "").strip()
    if pasted:
        kb_ingest_text(pasted); kb_prune()

    base_text = pasted or extracted.strip()
    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=240, key="edited_text")

# ---------- 우측 옵션/생성/다운로드 ----------
with col2:
    gen_mode = st.selectbox("🧠 생성 모드", ["핵심요약", "자연스러운 교육대본(무료)"])
    max_points = st.slider("요약 강도(핵심문장 개수)", 3, 10, 6)

    if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
        text_for_gen = (st.session_state.get("edited_text") or "").strip()
        if not text_for_gen:
            st.warning("텍스트가 비어 있습니다. PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
        else:
            with st.spinner("생성 중..."):
                if gen_mode == "자연스러운 교육대본(무료)":
                    script = make_structured_script(text_for_gen, max_points=max_points)
                    subtitle = "자연스러운 교육대본(무료)"
                else:
                    script = make_concise_report(text_for_gen, max_points=max_points)
                    subtitle = "핵심요약"
            st.success(f"생성 완료! ({subtitle})")
            st.text_area("결과 미리보기", value=script, height=420)
            c3, c4 = st.columns(2)
            with c3:
                st.download_button("⬇️ TXT 다운로드", data=_xml_safe(script).encode("utf-8"),
                                   file_name="tbm_output.txt", use_container_width=True)
            with c4:
                st.download_button("⬇️ DOCX 다운로드", data=to_docx_bytes(script),
                                   file_name="tbm_output.docx", use_container_width=True)

st.caption("완전 무료. 시드 KB(학습파일.zip 24개) + 업로드/붙여넣기 누적 학습 → 요약 가중/행동/질문 보강. 동적 주제 라벨. 도메인 템플릿은 기본 OFF로 안전 적용.")
