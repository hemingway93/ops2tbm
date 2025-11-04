# ==========================================================
# OPS2TBM — OPS/포스터 → TBM 교육 대본 자동 변환 (완전 무료)
# v2025-11-04:
#  - [①] 사고 문장 종결 보강: "사망/질식/추락..."으로 끝나도 자연스러운 종결("했습니다.") 추가
#  - [②] 질문/행동 후보 매칭 정밀화: 유사도 임계값↑ + 현재 텍스트 도메인 키워드 기반 필터
#  - [③][④] 의미 없는 짧은 문장(서술어만) 제거: 10글자 미만/동사 단독형 필터
#  - UI/레이아웃 변경 없음
# ==========================================================

import io
import zipfile
import re
from collections import Counter
from typing import List, Dict, Tuple

import numpy as np
import regex as rxx
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text
import pypdfium2 as pdfium

st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

# -------------------- 시드 KB --------------------
SEED_RISK_MAP = {
    "중독": "중독", "떨어짐": "떨어짐", "끼임": "끼임", "질식": "질식", "화재": "화재", "깔림": "깔림",
    "맞음": "맞음", "감전": "감전", "지붕": "지붕작업", "예초": "예초", "폭발": "폭발", "천공기": "천공",
    "선반": "절삭", "컨베이어": "협착", "부딪힘": "충돌", "미세먼지": "미세먼지", "크레인": "양중",
    "무너짐": "붕괴", "비계": "비계", "추락": "추락", "폭염": "폭염", "벌목": "벌목",
    "낙하": "낙하", "붕괴": "붕괴", "갱폼": "비계", "발판": "비계"
}
SEED_ACTIONS = [
    "밀폐공간작업 교육 및 훈련 실시", "출입 전 충분한 환기 실시", "작업 전 가스농도 측정 및 기록",
    "작업 상황 감시자 배치", "출입·퇴장 인원 점검", "보호장구 없이 구조 금지",
    "MSDS 확인 및 유해성 교육 실시", "국소배기장치 설치·가동", "환기가 불충분한 공간에서는 급기/배기팬 사용",
    "유기화합물 취급 시 방독마스크(갈색 정화통) 착용", "송기마스크·공기호흡기 적정 사용",
    "예초기 정지 후 이물질 제거·점검", "예초·벌목 작업 안전거리 유지 및 대피로 확보",
    "작업발판 견고히 설치 및 상태 점검", "개구부·개구창 추락 위험 구간 안전난간 설치",
    "안전대 지지점 연결 및 라이프라인 사용", "위험구역 설정·출입통제·감시자 배치",
    "양중 계획 수립 및 신호수 지정·통신 유지", "회전체·물림점 방호장치 설치 및 점검",
    "작업 전 작업계획서 작성 및 작업지휘자 지정", "개인보호구 착용(안전모·보호안경·안전화 등)",
    "화기작업 허가 및 안전점검", "정비·청소·점검 시 기계 전원 차단"
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

# -------------------- 세션 상태 --------------------
def _init_once():
    ss = st.session_state
    ss.setdefault("uploader_key", 0)
    ss.setdefault("kb_terms", Counter())
    ss.setdefault("kb_actions", [])
    ss.setdefault("kb_questions", [])
    ss.setdefault("domain_toggle", False)
    ss.setdefault("seed_loaded", False)
    ss.setdefault("last_file_diag", {})
    ss.setdefault("last_extracted_cache", "")
_init_once()

# -------------------- 유틸: 한글 조사/띄어쓰기 --------------------
def _has_final_consonant(k: str) -> bool:
    if not k:
        return False
    ch = k[-1]
    base = ord('가')
    code = ord(ch) - base
    if code < 0 or code > 11171:
        return False
    jong = code % 28
    return jong != 0

def add_obj_particle(noun: str) -> str:
    noun = noun.strip()
    if not noun:
        return noun
    particle = "을" if _has_final_consonant(noun[-1]) else "를"
    return f"{noun}{particle}"

def tidy_korean_spaces(s: str) -> str:
    # 간단한 띄어쓰기/문장부호 보정
    s = re.sub(r"\s+", " ", s)
    s = s.replace("전충분한", "전 충분한").replace("전충분히", "전 충분히")
    s = re.sub(r"\s([,.])", r"\1", s)
    return s.strip()

# -------------------- 전처리/패턴 --------------------
NOISE_PATTERNS = [
    r"^제?\s?\d{4}\s?[-.]?\s?\d+\s?호$",
    r"^(동절기\s*주요사고|안전작업방법|콘텐츠\s*링크|책자\s*OPS|숏폼\s*OPS)$",
    r"^(포스터|책자|스티커|콘텐츠 링크)$",
    r"^(스마트폰\s*APP|중대재해\s*사이렌|산업안전포털|고용노동부)$",
    r"^https?://\S+$",
    r"^\(?\s*PowerPoint\s*프레젠테이션\s*\)?$",
    r"^안전보건자료실.*$",
    r"^배포처\s+.*$",
    r"^홈페이지\s+.*$",
    r"^VR\s+.*$",
    r"^리플릿\s+.*$",
    r"^동영상\s+.*$",
    r"^APP\s+.*$",
    r".*검색해\s*보세요.*$",
]
BULLET_PREFIX = r"^[\s\-\•\●\▪\▶\▷\·\*\u25CF\u25A0\u25B6\u25C6\u2022\u00B7\u279C\u27A4\u25BA\u25AA\u25AB\u2611\u2713\u2714\u2716\u2794\u27A2]+"
DATE_PAT = r"([’']?\d{2,4})\.\s?(\d{1,2})\.\s?(\d{1,2})\.?"
META_PATTERNS = [
    r"<\s*\d+\s*명\s*사망\s*>", r"<\s*\d+\s*명\s*사상\s*>", r"<\s*\d+\s*명\s*의식불명\s*>",
    r"<\s*사망\s*\d+\s*명\s*>", r"<\s*사상\s*\d+\s*명\s*>"
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
    r"^\d+$", r"^\d{2,4}[-_]\d{1,}$", r"^\d{4}$", r"^(제)?\d+호$", r"^(호|호수|호차)$",
    r"^(사업장|업체|소재|소재지|장소|지역)$", r"^\d+\s*(명|건)$"
]
RISK_KEYWORDS = dict(SEED_RISK_MAP)

def tokens(s: str) -> List[str]:
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

def normalize_text(t: str) -> str:
    t = t.replace("\x0c", "\n")
    t = re.sub(r"[ \t]+\n", "\n", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

def strip_noise_line(line: str) -> str:
    s = (line or "").strip()
    if not s:
        return ""
    s = re.sub(BULLET_PREFIX, "", s).strip()
    for pat in NOISE_PATTERNS:
        if re.match(pat, s, re.IGNORECASE):
            return ""
    s = re.sub(r"https?://\S+", "", s).strip()
    s = s.strip("•●▪▶▷·-—–")
    return s

def merge_broken_lines(lines: List[str]) -> List[str]:
    """불릿/개행 파편을 문장 단위로 병합"""
    out, buf = [], ""
    for raw in lines:
        s = strip_noise_line(raw)
        if not s:
            continue
        s = s.lstrip("↳").strip()
        if buf and (buf.endswith((":", "·", "•", "▪", "▶", "▷")) or re.match(r"^\(.*\)$", buf)):
            buf += " " + s
        else:
            if buf:
                out.append(buf)
            buf = s
    if buf:
        out.append(buf)
    return out

def combine_date_with_next(lines: List[str]) -> List[str]:
    """날짜 줄과 사고 내용 줄을 합쳐 'YYYY년 M월 D일, 내용'으로 만듦"""
    out = []
    i = 0
    while i < len(lines):
        cur = strip_noise_line(lines[i])
        if re.search(DATE_PAT, cur) and (i + 1) < len(lines):
            nxt = strip_noise_line(lines[i + 1])
            if re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|무너짐|낙하)", nxt):
                m = re.search(DATE_PAT, cur)
                y, mo, d = m.groups()
                y = int(str(y).replace("’", "").replace("'", ""))
                y = 2000 + y if y < 100 else y
                out.append(f"{int(y)}년 {int(mo)}월 {int(d)}일, {nxt}")
                i += 2
                continue
        out.append(cur)
        i += 1
    return out

# -------- 사고 블록 병합 --------
CASE_JOIN_TRIG = ("쓰러지자", "구조하던 중", "차례로", "이어", "이후", "동시에", "결국", "그 과정에서", "외부에 있던", "현장에 있던")
CASE_KEYWORDS = ("사망", "사상", "중독", "추락", "붕괴", "낙하", "질식", "끼임", "깔림", "부딪힘", "감전", "폭발")

def stitch_case_blocks(sents: List[str]) -> List[str]:
    """연속된 사고 서술을 신호/키워드로 병합해 자연스러운 한 줄로 만든다."""
    if not sents:
        return sents
    out = []
    i = 0
    while i < len(sents):
        cur = sents[i].strip()
        merged = cur
        j = i + 1
        merged_any = False
        while j < len(sents):
            nxt = sents[j].strip()
            cond_keyword = (any(k in cur for k in CASE_KEYWORDS) and any(k in nxt for k in CASE_KEYWORDS))
            cond_trigger = (any(t in nxt for t in CASE_JOIN_TRIG) or any(t in cur for t in CASE_JOIN_TRIG))
            if cond_keyword or cond_trigger:
                sep = ", " if not merged.endswith(("다.", "습니다.", "했다.", ".")) else " "
                merged = tidy_korean_spaces(merged.rstrip(" .") + sep + nxt.lstrip(" ,"))
                cur = merged
                j += 1
                merged_any = True
            else:
                break
        out.append(merged)
        i = j if merged_any else i + 1
    # 중복 제거
    seen, dedup = set(), []
    for s in out:
        k = re.sub(r"\s+", "", s)
        if k not in seen:
            seen.add(k)
            dedup.append(s)
    return dedup

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
        if not s2:
            continue
        if re.search(r"(주요사고|안전작업방법|콘텐츠링크|주요 사고개요)$", s2):
            continue
        if len(s2) < 6:
            continue
        sents.append(s2)
    sents = stitch_case_blocks(sents)
    return sents

# -------------------- PDF 처리/진단 --------------------
def read_pdf_text_from_bytes(b: bytes, fname: str = "") -> str:
    """텍스트형 PDF 우선 추출, 실패 시 스캔 경고만 표시(OCR 미지원)."""
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
    st.session_state["last_file_diag"] = {
        "name": fname, "size_bytes": len(b), "extracted_chars": len(t),
        "note": "empty_or_scanned" if (len(t.strip()) < 10) else "ok"
    }
    return t

# -------------------- 요약/가중 --------------------
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
    if n == 0:
        return []
    W = cosim(X)
    row = W.sum(axis=1, keepdims=True)
    P = np.divide(W, row, out=np.zeros_like(W), where=row > 0)
    r = np.ones((n, 1), dtype=np.float32) / n
    tel = np.ones((n, 1), dtype=np.float32) / n
    for _ in range(max_iter):
        r2 = d * (P.T @ r) + (1 - d) * tel
        if np.linalg.norm(r2 - r, 1) < tol:
            r = r2
            break
        r = r2
    return [float(v) for v in r.flatten()]

def mmr_select(sents: List[str], scores: List[float], X: np.ndarray, k: int, lam: float = 0.7) -> List[int]:
    S = cosim(X)
    sel: List[int] = []
    rem = set(range(len(sents)))
    while rem and len(sel) < k:
        best, val = None, -1e9
        for i in rem:
            rel = scores[i]
            div = max((S[i, j] for j in sel), default=0.0)
            sc = lam * rel - (1 - lam) * div
            if sc > val:
                val, best = sc, i
        sel.append(best)
        rem.remove(best)
    return sel

def ai_extract_summary(text: str, limit: int = 8) -> List[str]:
    """TextRank+MMR 요약(세션 KB 가중치 반영)."""
    sents = preprocess_text_to_sentences(text)
    if not sents:
        return []
    kb = st.session_state["kb_terms"]
    total = sum(kb.values()) or 1
    kb_boost = {t: 1.0 + (cnt / total) * 3.0 for t, cnt in kb.items()} if kb else None
    X, _ = sentence_tfidf_vectors(sents, kb_boost=kb_boost)
    scores = textrank_scores(sents, X)
    idx = mmr_select(sents, scores, X, limit, lam=0.7)
    return [sents[i] for i in idx]

# -------------------- 템플릿/분류/NLG --------------------
DOMAIN_TEMPLATES = [
    ({"비계", "발판", "갱폼", "추락"}, "작업발판을 견고하게 설치하고 안전난간 및 추락방호망을 확보합니다."),
    ({"안전난간", "난간", "개구부"}, "개구부·개구창 등 추락 위험 구간에 안전난간을 설치합니다."),
    ({"MSDS", "국소배기", "환기"}, "취급 물질의 MSDS를 확인하고 국소배기장치를 가동하여 충분히 환기합니다."),
    ({"예초", "벌목", "예초기"}, "예초·벌목 작업 시 작업자 간 안전거리를 유지하고 대피로를 확보합니다."),
    ({"크레인", "양중"}, "양중 계획을 수립하고 신호수를 지정하여 통신을 유지합니다."),
    ({"컨베이어", "협착", "회전체"}, "회전체·물림점 접촉을 방지하도록 방호장치를 설치하고 점검합니다."),
]

def jaccard(a: set, b: set) -> float:
    return len(a & b) / (len(a | b) + 1e-8)

ACTION_VERBS = [
    "설치", "배치", "착용", "점검", "확인", "측정", "기록", "표시", "제공", "비치", "보고", "신고",
    "교육", "주지", "중지", "통제", "휴식", "환기", "차단", "교대", "배제", "배려", "가동", "준수",
    "운영", "유지", "교체", "정비", "청소", "고정", "격리", "보호", "보수", "작성", "지정"
]
ACTION_PAT = (
    r"(?P<obj>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(?P<verb>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"
    r"|(?P<obj2>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(을|를)\s*(?P<verb2>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"
)

# ---------- (③④) 의미 있는 문장 여부 판정 ----------
def is_meaningful_sentence(s: str) -> bool:
    """10자 미만 또는 동사 단독형(의미 빈약)은 제외"""
    raw = re.sub(r"\s+", "", s)
    if len(raw) < 10:
        return False
    # '합니다/한다'만 있는 단독형 배제(조사 없는 케이스 등)
    if re.fullmatch(r"[가-힣\s]+(합니다|한다)\.", s.strip()):
        return False
    return True

def soften(s: str) -> str:
    s = s.replace("하여야", "해야 합니다").replace("한다", "합니다").replace("한다.", "합니다.")
    s = s.replace("바랍니다", "해주세요").replace("확인 바람", "확인해주세요")
    s = s.replace("금지한다", "금지합니다").replace("필요하다", "필요합니다")
    s = re.sub(r"^\(([^)]+)\)\s*", "", s)
    for pat in META_PATTERNS:
        s = re.sub(pat, "", s).strip()
    s = re.sub(BULLET_PREFIX, "", s).strip(" -•●\t")
    return tidy_korean_spaces(s)

def is_accident_sentence(s: str) -> bool:
    if any(w in s for w in ["예방", "대책", "지침", "수칙"]):
        return False
    return bool(re.search(DATE_PAT, s) or re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|무너짐|낙하)", s))

def is_prevention_sentence(s: str) -> bool:
    return any(w in s for w in ["예방", "대책", "지침", "수칙", "안전조치"])

def is_risk_sentence(s: str) -> bool:
    return any(w in s for w in ["위험", "요인", "원인", "증상", "결빙", "강풍", "폭염", "미세먼지", "회전체", "비산", "말림", "추락", "낙하", "협착"])

def _domain_template_apply(s: str, base_text: str) -> str:
    if not st.session_state.get("domain_toggle"):
        return s
    sent_toks = set(tokens(s))
    base_toks = set(tokens(base_text))
    if jaccard(sent_toks, base_toks) < 0.05:
        return s
    best = None
    best_hits = 0
    for triggers, render in DOMAIN_TEMPLATES:
        if (sent_toks & triggers) and (base_toks & triggers):
            hits = len((sent_toks | base_toks) & triggers)
            if hits > best_hits:
                best_hits = hits
                best = render
    return best if best else s

def to_action_sentence(s: str, base_text: str) -> str:
    """행동/수칙 문장을 자연스럽게 리라이팅 + 조사 보정"""
    s2 = soften(s)
    s2 = re.sub(r"(위기탈출\s*안전보건)", "", s2).strip()
    s2 = re.sub(r"\s*에\s*따른\s*", " 시 ", s2)
    s2 = re.sub(r"\s*에\s*따라\s*", " 시 ", s2)

    # 도메인 템플릿(선택적) 적용
    s2_tpl = _domain_template_apply(s2, base_text)
    if s2_tpl != s2:
        txt = s2_tpl
        if not txt.endswith(("다.", "습니다.", "합니다.")):
            txt = txt.rstrip(" .") + " 합니다."
        return tidy_korean_spaces(txt)

    m = re.search(ACTION_PAT, s2)
    if not m:
        # 패턴 감지가 안 되면 종결만 정리
        txt = s2 if s2.endswith(("니다.", "합니다.", "다.")) else (s2.rstrip(" .") + " 합니다.")
        return tidy_korean_spaces(txt)

    obj = (m.group("obj") or m.group("obj2") or "").strip()
    verb = (m.group("verb") or m.group("verb2") or "실시").strip()

    # 목적어 조사 보정
    if obj and not re.search(r"(을|를|에|에서|과|와|의)$", obj):
        obj = add_obj_particle(obj)

    # 접두어(반드시/작업 전)
    prefix = "반드시 " if "설치" in verb else ("작업 전 " if verb in ("확인", "점검", "측정", "기록", "작성", "지정") else "")
    core = tidy_korean_spaces(f"{prefix}{obj} {verb}")
    return core.rstrip(" .") + " 합니다."

def classify_sentence(s: str) -> str:
    if is_accident_sentence(s):
        return "case"
    if re.search(ACTION_PAT, s) or is_prevention_sentence(s):
        return "action"
    if is_risk_sentence(s):
        return "risk"
    if "?" in s or "확인" in s or "점검" in s:
        return "question"
    return "other"

# -------------------- KB 누적/시드 --------------------
def seed_kb_once():
    if not st.session_state["seed_loaded"]:
        for t, k in SEED_RISK_MAP.items():
            if t not in RISK_KEYWORDS:
                RISK_KEYWORDS[t] = k
        for a in SEED_ACTIONS:
            if 2 <= len(a) <= 160:
                st.session_state["kb_actions"].append(a if a.endswith(("다", "다.", "합니다", "합니다.")) else a + " 합니다.")
        for q in SEED_QUESTIONS:
            st.session_state["kb_questions"].append(q if q.endswith("?") else q + "?")
        for t in SEED_RISK_MAP.keys():
            st.session_state["kb_terms"][t] += 5
        st.session_state["seed_loaded"] = True

def kb_ingest_text(text: str) -> None:
    if not (text or "").strip():
        return
    sents = preprocess_text_to_sentences(text)
    for s in sents:
        for t in tokens(s):
            if len(t) >= 2:
                st.session_state["kb_terms"][t] += 1
                if re.search(r"(추락|낙하|깔림|끼임|중독|질식|화재|폭발|감전|폭염|붕괴|비계|갱폼|예초|벌목|컨베이어|크레인|지붕|선반|천공)", t):
                    if t not in RISK_KEYWORDS:
                        RISK_KEYWORDS[t] = t
    for s in sents:
        if re.search(ACTION_PAT, s) or is_prevention_sentence(s):
            cand = to_action_sentence(s, text)
            if 2 <= len(cand) <= 160:
                st.session_state["kb_actions"].append(cand)
    for s in sents:
        if "?" in s or "확인" in s or "점검" in s:
            q = soften(s if s.endswith("?") else s + " 맞습니까?")
            if 2 <= len(q) <= 140:
                st.session_state["kb_questions"].append(q)

def kb_prune() -> None:
    def dedup_keep_order(lst: List[str]) -> List[str]:
        seen, out = set(), []
        for x in lst:
            k = re.sub(r"\s+", "", x)
            if k not in seen:
                seen.add(k)
                out.append(x)
        return out
    st.session_state["kb_actions"] = dedup_keep_order(st.session_state["kb_actions"])[:900]
    st.session_state["kb_questions"] = dedup_keep_order(st.session_state["kb_questions"])[:500]
    st.session_state["kb_terms"] = Counter(dict(st.session_state["kb_terms"].most_common(1800)))

# ---------- (②) 도메인 키워드 기반 + 유사도 임계 보강 ----------
def kb_match_candidates(cands: List[str], base_text: str, limit: int, min_sim: float = 0.15) -> List[str]:
    """현재 텍스트에 없는 위험 키워드가 들어간 후보는 배제 + 유사도 임계치 적용"""
    bt = set(tokens(base_text))
    # 현재 본문에 등장한 위험 키워드 집합(원형/매핑 모두 고려)
    present_risks = {t for t in bt if (t in RISK_KEYWORDS or t in RISK_KEYWORDS.values())}
    scored: List[Tuple[float, str]] = []
    for c in cands:
        ct = set(tokens(c))
        # 도메인 키워드 필터: 후보에 있는 위험 키워드가 현재 텍스트에 없으면 제외
        cand_risks = {RISK_KEYWORDS.get(t, t) for t in ct if (t in RISK_KEYWORDS or t in RISK_KEYWORDS.values())}
        if cand_risks and not (cand_risks & present_risks):
            continue
        j = len(bt & ct) / (len(bt | ct) + 1e-8)
        if j >= min_sim:
            scored.append((j, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [c for _, c in scored[:limit]]

# -------------------- 생성기 --------------------
def naturalize_case_sentence(s: str) -> str:
    """사고 서술을 자연어 종결로 보정(① 종결 강화)."""
    s = soften(s)
    # 사상/사망 등 수치
    death = re.search(r"사망\s*(\d+)\s*명", s)
    inj = re.search(r"사상\s*(\d+)\s*명", s)
    unconscious = re.search(r"의식불명", s)
    info = []
    if death:
        info.append(f"근로자 {death.group(1)}명 사망")
    if inj and not death:
        info.append(f"{inj.group(1)}명 사상")
    if unconscious:
        info.append("의식불명 발생")
    # 날짜
    m = re.search(DATE_PAT, s)
    date_txt = ""
    if m:
        y, mo, d = m.groups()
        y = int(str(y).replace("’", "").replace("'", ""))
        y = 2000 + y if y < 100 else y
        date_txt = f"{int(y)}년 {int(mo)}월 {int(d)}일, "
        s = s.replace(m.group(0), "").strip()
    s = s.strip(" ,.-")

    # 종결 처리: 사고 키워드로 끝나는 경우 "했습니다." 부여
    if not re.search(r"(다\.|입니다\.|했습니다\.)$", s):
        if re.search(r"(사망|사상|중독|추락|낙하|붕괴|질식|끼임|깔림|부딪힘|감전|폭발)\s*$", s):
            s = s.rstrip(" .") + "했습니다."
        elif re.search(r"(사건|사고)\s*$", s):
            s = s.rstrip(" .") + "가 발생했습니다."
        else:
            s = s.rstrip(" .") + " 사고가 발생했습니다."

    # 추가 정보 꼬리(이미 '했습니다.'로 끝난 경우는 중복 회피)
    if info and not s.endswith("했습니다."):
        s = tidy_korean_spaces(s.rstrip(" .") + " " + (", ".join(info)) + "했습니다.")
    return tidy_korean_spaces((date_txt + s).strip())

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

    # (②) 도메인/유사도 필터로 KB 보강
    if len(act) < 5 and st.session_state["kb_actions"]:
        act += kb_match_candidates(st.session_state["kb_actions"], text, 5 - len(act), min_sim=0.15)
    if not ask and st.session_state["kb_questions"]:
        ask = kb_match_candidates(st.session_state["kb_questions"], text, 3, min_sim=0.15)

    # (③④) 의미 없는 짧은 문장 제거
    act = [a for a in act if is_meaningful_sentence(a)][:5]
    ask = [q if q.endswith("?") else q + "?" for q in ask if is_meaningful_sentence(q)][:3]

    # 본문 구성
    lines = []
    lines.append(f"🦺 TBM 교육대본 – {topic_label}\n")
    lines.append("◎ 도입")
    lines.append(f"오늘은 최근 발생한 '{topic_label.replace(' 재해예방','')}' 사고 사례를 중심으로, 우리 현장에서 같은 사고를 예방하기 위한 안전조치를 함께 살펴보겠습니다.\n")
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
    lines.append("예방조치는 '선조치 후작업'이 원칙입니다. 오늘 작업 전, 각 공정별 위험요인을 다시 한 번 점검하고 필요한 보호구와 안전조치를 반드시 준비합시다.")
    lines.append("◎ 구호")
    lines.append("“한 번 더 확인! 한 번 더 점검!”")
    return "\n".join(lines)

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
                seen.add(k)
                out.append(x)
        return out

    # (③④) 의미 없는 문장 제거
    cases = uniq_keep([c for c in cases if is_meaningful_sentence(c)])[:3]
    risks = uniq_keep([r for r in risks if is_meaningful_sentence(r)])[:3]
    acts  = uniq_keep([a for a in acts if is_meaningful_sentence(a)])[:4]

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

# -------------------- 라벨/토픽 --------------------
def drop_label_token(t: str) -> bool:
    if t in STOP_TERMS:
        return True
    for pat in LABEL_DROP_PAT:
        if re.match(pat, t):
            return True
    if t in {"소재", "소재지", "지역", "장소", "버스", "영업소", "업체", "자료", "키", "메세지", "명", "안전보건"}:
        return True
    return False

def top_terms_for_label(text: str, k: int = 3) -> List[str]:
    doc_cnt = Counter([t for t in tokens(text) if not drop_label_token(t)])
    bonus = Counter()
    for t in list(doc_cnt.keys()):
        if t in RISK_KEYWORDS:
            bonus[RISK_KEYWORDS[t]] += doc_cnt[t]
    doc_cnt += bonus
    kb = st.session_state["kb_terms"]
    if kb:
        for t, c in kb.items():
            if not drop_label_token(t):
                doc_cnt[t] += 0.2 * c
    if not doc_cnt:
        return ["안전보건", "교육"]
    commons = {"안전", "교육", "작업", "현장", "예방", "조치", "확인", "관리", "점검", "가이드", "지침"}
    cand = [(t, doc_cnt[t]) for t in doc_cnt if t not in commons and len(t) >= 2]
    if not cand:
        cand = list(doc_cnt.items())
    cand.sort(key=lambda x: x[1], reverse=True)
    return [t for t, _ in cand[:k]]

def dynamic_topic_label(text: str) -> str:
    terms = top_terms_for_label(text, k=3)
    risks = [RISK_KEYWORDS.get(t, t) for t in terms if t in RISK_KEYWORDS or t in RISK_KEYWORDS.values()]
    extra = [t for t in terms if t not in risks]
    label_core = " ".join(sorted(set(risks), key=risks.index)) or "안전보건"
    tail = " ".join(extra[:1])
    label = (label_core + (" " + tail if tail else "")).strip()
    if "재해예방" not in label:
        label += " 재해예방"
    return label

# -------------------- DOCX --------------------
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

# -------------------- UI --------------------
with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 파이프라인(완전 무료, 오픈소스만 사용)**  
1) 전처리(노이즈 제거/줄 병합/날짜-사고 결합)  
2) **사고 블록 병합**(연결어·키워드 기반으로 연속 서술을 한 문장으로)  
3) TextRank + MMR 요약 (**세션 KB 가중치** 반영)  
4) 규칙형 NLG: 조사/띄어쓰기·종결 보정, 행동/질문 템플릿 보강  
5) 결과 포맷: **자연스러운 교육대본(무료)** / **핵심요약**
""")
    st.session_state["domain_toggle"] = st.toggle(
        "🔧 도메인 템플릿 강화(신중 적용)",
        value=False,
        help="문장·본문 트리거 일치 + 유사도 기준 충족 시에만 템플릿을 적용합니다."
    )

seed_kb_once()
st.title("🦺 OPS/포스터를 교육 대본으로 자동 변환 (완전 무료)")

def reset_all():
    # 전체 상태 초기화(파일 입력 키 포함)
    st.session_state.pop("manual_text", None)
    st.session_state.pop("edited_text", None)
    st.session_state.pop("zip_choice", None)
    st.session_state["kb_terms"] = Counter()
    st.session_state["kb_actions"] = []
    st.session_state["kb_questions"] = []
    st.session_state["uploader_key"] += 1
    st.session_state["seed_loaded"] = False
    st.session_state["last_file_diag"] = {}
    st.session_state["last_extracted_cache"] = ""
    st.rerun()

col_top1, col_top2 = st.columns([4, 1])
with col_top2:
    st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

st.markdown("**안내**  \n- 텍스트가 포함된 PDF 또는 본문 텍스트를 권장합니다.  \n- 이미지/스캔 PDF는 현재 OCR 미지원입니다.")

col1, col2 = st.columns([1, 1], gap="large")

# ---------- 좌측 입력/미리보기 ----------
with col1:
    uploaded = st.file_uploader(
        "OPS 업로드 (PDF 또는 ZIP) • 텍스트 PDF 권장",
        type=["pdf", "zip"],
        key=f"uploader_{st.session_state['uploader_key']}"
    )
    manual_text = st.text_area(
        "또는 OPS 텍스트 직접 붙여넣기",
        key="manual_text",
        height=220,
        placeholder="예: 현장 안내문 또는 OPS 본문 텍스트…"
    )

    extracted: str = ""
    zip_pdfs: Dict[str, bytes] = {}

    if uploaded is not None:
        fname = (uploaded.name or "").lower()
        try:
            raw_bytes = uploaded.getvalue()
        except Exception:
            raw_bytes = uploaded.read()

        if fname.endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
                    for name in zf.namelist():
                        if name.lower().endswith(".pdf"):
                            data = zf.read(name)
                            zip_pdfs[name] = data
                if zip_pdfs:
                    # ZIP 전체 학습
                    for nm, data in zip_pdfs.items():
                        txt_all = read_pdf_text_from_bytes(data, fname=f"{fname}::{nm}")
                        if txt_all.strip():
                            kb_ingest_text(txt_all)
                    kb_prune()
                    # 자동 선택 + 에디터 채움
                    first_name = sorted(zip_pdfs.keys())[0]
                    extracted = read_pdf_text_from_bytes(zip_pdfs[first_name], fname=first_name)
                    if extracted.strip():
                        st.session_state["edited_text"] = extracted
                        st.session_state["last_extracted_cache"] = extracted
                    st.success(f"ZIP 감지: {len(zip_pdfs)}개 PDF, 첫 문서 자동 선택 → {first_name}")
                else:
                    st.error("ZIP 내에 PDF가 없습니다.")
            except Exception as e:
                st.error(f"ZIP 해제 오류: {e}")

            if zip_pdfs:
                chosen = st.selectbox("ZIP 내 PDF 선택", sorted(zip_pdfs.keys()), key="zip_choice")
                if chosen and zip_pdfs.get(chosen):
                    extracted2 = read_pdf_text_from_bytes(zip_pdfs[chosen], fname=chosen)
                    if extracted2.strip():
                        st.session_state["edited_text"] = extracted2
                        st.session_state["last_extracted_cache"] = extracted2

        elif fname.endswith(".pdf"):
            extracted = read_pdf_text_from_bytes(raw_bytes, fname=fname)
            if extracted.strip():
                kb_ingest_text(extracted)
                kb_prune()
                st.session_state["edited_text"] = extracted
                st.session_state["last_extracted_cache"] = extracted
            else:
                st.warning("⚠️ PDF에서 유효한 텍스트를 추출할 수 없습니다.")

        else:
            st.warning("지원하지 않는 형식입니다. PDF 또는 ZIP을 업로드하세요.")

    pasted = (manual_text or "").strip()
    if pasted:
        kb_ingest_text(pasted)
        kb_prune()
        st.session_state["edited_text"] = pasted
        st.session_state["last_extracted_cache"] = pasted

    # 미리보기
    base_text = st.session_state.get("edited_text", "")
    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=240, key="edited_text")

    with st.expander("🧪 파일 읽기 진단(Log-lite)", expanded=False):
        diag = st.session_state.get("last_file_diag", {})
        if diag:
            st.write({
                "파일명": diag.get("name"),
                "크기(bytes)": diag.get("size_bytes"),
                "추출된 문자수": diag.get("extracted_chars"),
                "메모": diag.get("note"),
            })
        st.caption(f"현재 텍스트 박스 길이: {len(st.session_state.get('edited_text',''))} chars")

# ---------- 우측 옵션/생성/다운로드 ----------
with col2:
    gen_mode = st.selectbox("🧠 생성 모드", ["핵심요약", "자연스러운 교육대본(무료)"])
    max_points = st.slider("요약 강도(핵심문장 개수)", 3, 10, 6)

    if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
        text_for_gen = (st.session_state.get("edited_text") or "").strip()
        if not text_for_gen:
            text_for_gen = (st.session_state.get("last_extracted_cache") or "").strip()
            if text_for_gen:
                st.info("빈 입력을 최근 추출 텍스트로 자동 대체했습니다.")

        if not text_for_gen:
            st.warning("PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
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
                st.download_button(
                    "⬇️ TXT 다운로드",
                    data=_xml_safe(script).encode("utf-8"),
                    file_name="tbm_output.txt",
                    use_container_width=True
                )
            with c4:
                st.download_button(
                    "⬇️ DOCX 다운로드",
                    data=to_docx_bytes(script),
                    file_name="tbm_output.docx",
                    use_container_width=True
                )

st.caption("완전 무료. 시드 KB + 업로드 누적 학습 → 요약 가중/행동/질문 보강. 사고 블록 병합·종결 보정으로 사례 흐름 자연화.")
