# ==========================================================
# OPS2TBM — OPS/포스터 → TBM 교육 대본 자동 변환 (완전 무료)
# v2025-11-05-c (tech-notes)
#
# [프로젝트/기술 스택 주석 — 제출용]
# - 구현 형태: Web App (Streamlit)
# - 사용 언어: Python 3
# - 오픈소스 라이브러리(모두 무료, 표준 라이선스):
#   * streamlit .......... 웹 UI/상태 관리 (순수 파이썬, 서버리스 클라우드 배포 호환)
#   * pdfminer.six ........ 텍스트 기반 PDF의 본문 텍스트 추출
#   * pypdfium2 ........... PDF 이미지/스캔 여부 판단(간단 진단) 및 처리를 위한 백엔드(이번 버전은 OCR 미사용)
#   * python-docx .......... 생성된 대본을 DOCX로 내보내기
#   * regex(=regex 패키지) .. 한국어/유니코드 친화 정규식(파이썬 re 보강)
#   * numpy ............... TF-IDF/코사인 유사도/텍스트랭크 벡터 계산
#
# - AI/알고리즘(유료 API 전혀 없음 / 완전 무료·로컬):
#   * 텍스트 전처리: 잡음 제거, 헤더 감지, 줄 병합, 날짜-사고 결합
#   * 문장 분할 + 연결: OPS 문서 특유의 불릿/단답을 자연문으로 재구성
#   * 요약: TextRank + MMR(다양성 조절) — 고전 그래프 기반 요약, LLM 불필요
#   * 의미 가중치: 세션 KB(사용자가 올린 PDF/텍스트에서 동적 용어 통계)로 TF-IDF에 가중치
#   * 섹션 추출: 
#     - (1) 헤더 기반 파서(사고사례/예방수칙 등 제목을 직접 인식)
#     - (2) 헤더 없이도 불릿 클러스터를 사례형/예방형으로 자동 분류 (행동동사 패턴)
#     - (3) Fallback: 키워드/날짜 기반 사고문장과 행동문을 자동 수집
#   * 규칙형 NLG: 조사/종결/띄어쓰기 보정, “~ 합니다.” 일관 문체화, 
#                짧은 불완전 문장 병합(예: '예방 실시' → '예방 조치를 실시합니다.')
#
# - 데이터/학습(경량):
#   * 세션 KB(임시): 업로드된 PDF/텍스트에서 위험어/행동문을 누적 수집(세션 범위에 한함)
#   * 시드 KB(고정): 안전·산업 재해 관련 기초 리스크 키워드/질문/예방문장 초기값
#   * 외부 LLM/벡터DB/서버 의존 없음 (오프라인/무료 실행 가능)
#
# - 아키텍처 포인트:
#   * UI는 바꾸지 않음(요청 정책 준수). 내부 파이프라인만 개선/주석 추가.
#   * 이미지 스캔 PDF는 OCR 미지원(명확히 표시). 텍스트 PDF/복붙 텍스트에 최적화.
#   * 결과물은 TXT/DOCX로 즉시 다운로드 가능.
#
# - 보안/운영 유의점:
#   * 모든 처리는 세션 내 메모리에서만 동작(개인 토큰·클라우드 API 불필요)
#   * 문서 업로드 데이터는 서버 메모리/세션 범위에만 존재(세션 종료 시 소멸)
#   * 공단 클라우드/LMS 연동은 추후 확장 단계에서 REST/SSO 등의 방법으로 가능(본 버전은 독립형)
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

# ---------- [PDF 텍스트 추출 계층 — 오픈소스 조합 설명] ----------
# 1) pdfminer.six 추출 우선: 텍스트 기반 PDF에서 본문 텍스트를 구조적으로 추출.
# 2) 실패/부족 시 pypdfium2로 간단 로드하여 "이미지/스캔" 가능성만 감지(이번 버전 OCR 미사용).
#    - 추후 무료 OCR(예: tesseract + pyocr/ocrmypdf) 조합으로 확장 가능.
pdf_extract_text = None
try:
    from pdfminer.high_level import extract_text as _extract_text
    pdf_extract_text = _extract_text
except Exception:
    try:
        from pdfminer_high_level import extract_text as _extract_text_compat  # type: ignore
        pdf_extract_text = _extract_text_compat
    except Exception:
        pdf_extract_text = None

import pypdfium2 as pdfium

# ---------- [Streamlit UI 설정 — 변경 금지 지침 준수] ----------
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

# -------------------- 시드 KB(정적) --------------------
# - 리스크 키워드/행동 수칙/점검 질문의 "초기값"을 주입(문서 도메인에 맞춘 가벼운 prior).
SEED_RISK_MAP = {
    "중독":"중독","떨어짐":"떨어짐","끼임":"끼임","질식":"질식","화재":"화재","깔림":"깔림",
    "맞음":"맞음","감전":"감전","지붕":"지붕작업","예초":"예초","폭발":"폭발","천공기":"천공",
    "선반":"절삭","컨베이어":"협착","부딪힘":"충돌","미세먼지":"미세먼지","크레인":"양중",
    "무너짐":"붕괴","비계":"비계","추락":"추락","폭염":"폭염","벌목":"벌목","낙하":"낙하","붕괴":"붕괴",
    "갱폼":"비계","발판":"비계","화학물질":"화학물질","밀폐공간":"밀폐공간"
}
SEED_ACTIONS = [
    "밀폐공간작업 교육 및 훈련 실시","출입 전 충분한 환기 실시","작업 전 가스농도 측정 및 기록",
    "작업 상황 감시자 배치","출입·퇴장 인원 점검","보호장구 없이 구조 금지",
    "MSDS 확인 및 유해성 교육 실시","국소배기장치 설치·가동","환기가 불충분한 공간에서는 급기/배기팬 사용",
    "유기화합물 취급 시 방독마스크(갈색 정화통) 착용","송기마스크·공기호흡기 적정 사용",
    "예초기 정지 후 이물질 제거·점검","예초·벌목 작업 시 안전거리 유지 및 대피로 확보",
    "작업발판 견고히 설치 및 상태 점검","개구부·개구창 추락 위험 구간 안전난간 설치",
    "안전대 지지점 연결 및 라이프라인 사용","위험구역 설정·출입통제·감시자 배치",
    "양중 계획 수립 및 신호수 지정·통신 유지","회전체·물림점 방호장치 설치 및 점검",
    "작업 전 작업계획서 작성 및 작업지휘자 지정","개인보호구 착용(안전모·보호안경·안전화 등)",
    "화기작업 허가 및 안전점검","정비·청소·점검 시 기계 전원 차단",
    "밀폐공간 작업 시 산소·유해가스 농도 측정","위험물질 취급 시 MSDS 비치·게시 및 교육",
    "위험물질 취급 시 불침투성 보호복·방독마스크 착용","환기 실시 및 감시인 배치"
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

# ---------- [세션 상태 초기화 — 저장소/KB/캐시] ----------
def _init_once():
    ss = st.session_state
    ss.setdefault("uploader_key", 0)
    ss.setdefault("kb_terms", Counter())     # 동적 용어(토큰) 출현 빈도 — TF-IDF 가중치에 반영
    ss.setdefault("kb_actions", [])          # 동적으로 수집된 행동형 수칙 문장
    ss.setdefault("kb_questions", [])        # 동적으로 수집된 점검형 질문
    ss.setdefault("domain_toggle", False)    # 템플릿 강화 토글(보수적 적용)
    ss.setdefault("seed_loaded", False)      # 시드 KB 1회 주입 여부
    ss.setdefault("last_file_diag", {})      # 파일 진단(크기/추출문자수/메모)
    ss.setdefault("last_extracted_cache", "")# 최근 추출 텍스트 캐시
_init_once()

# -------------------- 한국어 조사/띄어쓰기 보정 --------------------
# - 규칙형 NLG의 핵심: 명사 + 목적격 조사 자동 부착, 띄어쓰기/종결統一
def _has_final_consonant(k: str) -> bool:
    if not k: return False
    ch = k[-1]
    base = ord('가'); code = ord(ch) - base
    if code < 0 or code > 11171: return False
    jong = code % 28
    return jong != 0

def add_obj_particle(noun: str) -> str:
    noun = noun.strip()
    if not noun: return noun
    particle = "을" if _has_final_consonant(noun[-1]) else "를"
    return f"{noun}{particle}"

def tidy_korean_spaces(s: str) -> str:
    s = re.sub(r"\s+", " ", s)
    s = s.replace("전충분한","전 충분한").replace("전충분히","전 충분히")
    s = re.sub(r"\s([,.])", r"\1", s)
    return s.strip()

# -------------------- 전처리 파이프라인 --------------------
# - OPS 문서 특성(불릿, 줄바꿈 분절, 머리글 파편화)을 자연문으로 묶는 단계
NOISE_PATTERNS = [
    r"^제?\s?\d{4}\s?[-.]?\s?\d+\s?호$",
    r"^(동절기\s*주요사고|안전작업방법|콘텐츠\s*링크|책자\s*OPS|숏폼\s*OPS)$",
    r"^(포스터|책자|스티커|콘텐츠 링크)$",
    r"^(스마트폰\s*APP|중대재해\s*사이렌|산업안전포털|고용노동부)$",
    r"^https?://\S+$", r"^\(?\s*PowerPoint\s*프레젠테이션\s*\)?$",
    r"^안전보건자료실.*$", r"^배포처\s+.*$", r"^홈페이지\s+.*$",
    r"^VR\s+.*$", r"^리플릿\s+.*$", r"^동영상\s+.*$", r"^APP\s+.*$",
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

RISK_KEYWORDS = dict(SEED_RISK_MAP)  # 세션 수집으로 계속 보강

def tokens(s: str) -> List[str]:
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

def normalize_text(t: str) -> str:
    t = t.replace("\x0c","\n")
    t = re.sub(r"[ \t]+\n","\n", t)
    t = re.sub(r"\n{3,}","\n\n", t)
    return t.strip()

def strip_noise_line(line: str) -> str:
    s = (line or "").strip()
    if not s: return ""
    s = re.sub(BULLET_PREFIX,"", s).strip()
    for pat in NOISE_PATTERNS:
        if re.match(pat, s, re.IGNORECASE):
            return ""
    s = re.sub(r"https?://\S+","", s).strip()
    s = s.strip("•●▪▶▷·-—–")
    return s

def _looks_like_heading(s: str) -> bool:
    return bool(re.search(r"(방법|수칙|대책|안전조치|예방|작업방법|사고사례|주요\s*사고사례)\s*[:：]?$", s))

def merge_broken_lines(lines: List[str]) -> List[str]:
    # 불완전 헤더/짧은 문장들을 옆 줄과 합쳐 자연문으로 만든다.
    out, buf = [], ""
    for raw in lines:
        s = strip_noise_line(raw)
        if not s: continue
        if _looks_like_heading(s) or s.endswith((":", "：", "-", "·")):
            if buf: out.append(buf)
            buf = s; continue
        if buf:
            if buf.endswith((":", "：", "-", "·")):
                buf = tidy_korean_spaces(buf.rstrip(" :：-·") + " " + s); continue
            if (len(buf) < 20 and not re.search(r"[.?!다]$", buf)) or (len(s) < 20 and not re.search(r"[.?!다]$", s)):
                buf = tidy_korean_spaces(buf + " " + s); continue
            if not re.search(r"[.?!다]$", buf):
                buf = tidy_korean_spaces(buf + " " + s); continue
            out.append(buf); buf = s
        else:
            buf = s
    if buf: out.append(buf)
    return out

def combine_date_with_next(lines: List[str]) -> List[str]:
    # 날짜 한 줄 + 다음 줄 사고 설명 → 한 줄로 합쳐 가독성/의미 유지
    out = []; i = 0
    while i < len(lines):
        cur = strip_noise_line(lines[i])
        if re.search(DATE_PAT, cur) and (i+1) < len(lines):
            nxt = strip_noise_line(lines[i+1])
            if re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|감전|폭발)", nxt):
                m = re.search(DATE_PAT, cur)
                y, mo, d = m.groups()
                y = int(str(y).replace("’","").replace("'","")); y = 2000 + y if y < 100 else y
                out.append(f"{int(y)}년 {int(mo)}월 {int(d)}일, {nxt}")
                i += 2; continue
        out.append(cur); i += 1
    return out

# 연결어/사고키워드 기반으로 연속 서술을 하나의 사례문으로 봉합
CASE_JOIN_TRIG = ("쓰러지자","구조하던 중","차례로","이어","이후","동시에","결국","그 과정에서","외부에 있던","현장에 있던")
CASE_KEYWORDS = ("사망","사상","중독","추락","붕괴","낙하","질식","끼임","깔림","부딪힘","감전","폭발","사고")

def stitch_case_blocks(sents: List[str]) -> List[str]:
    if not sents: return sents
    out = []; i = 0
    while i < len(sents):
        cur = sents[i].strip(); merged = cur; j = i + 1; merged_any = False
        while j < len(sents):
            nxt = sents[j].strip()
            cond_keyword = (any(k in cur for k in CASE_KEYWORDS) and any(k in nxt for k in CASE_KEYWORDS))
            cond_trigger = (any(t in nxt for t in CASE_JOIN_TRIG) or any(t in cur for t in CASE_JOIN_TRIG))
            if cond_keyword or cond_trigger:
                sep = ", " if not merged.endswith(("다.","습니다.","했다.",".")) else " "
                merged = tidy_korean_spaces(merged.rstrip(" .") + sep + nxt.lstrip(" ,"))
                cur = merged; j += 1; merged_any = True
            else:
                break
        out.append(merged); i = j if merged_any else i + 1
    seen, dedup = set(), []
    for s in out:
        k = re.sub(r"\s+","", s)
        if k not in seen:
            seen.add(k); dedup.append(s)
    return dedup

def preprocess_text_to_sentences(text: str) -> List[str]:
    # [전처리 핵심 단계] 노이즈 제거 → 줄 병합 → 날짜-사고 결합 → 문장화 → 사례문 봉합
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
        if re.search(r"(주요사고|안전작업방법|콘텐츠링크|주요 사고개요)$", s2): continue
        if len(re.sub(r"\s+","", s2)) < 4:
            continue
        sents.append(s2)
    sents = stitch_case_blocks(sents)
    return sents

# -------------------- (1) 헤더 기반 섹션 파서 --------------------
SECTION_HEADERS_CASE = [r"주요\s*사고사례", r"사고사례", r"사고\s*사례"]
SECTION_HEADERS_PREV = [
    r"안전\s*작업방법", r"밀폐공간\s*작업\s*시", r"밀폐공간작업\s*시",
    r"위험물질\s*취급\s*시", r"예방\s*수칙", r"실천\s*수칙", r"예방\s*조치",
    r"안전\s*수칙", r"작업\s*수칙"
]
def _compile_headers(headers: List[str]) -> List[re.Pattern]:
    return [re.compile(h, re.IGNORECASE) for h in headers]
HDR_CASE = _compile_headers(SECTION_HEADERS_CASE)
HDR_PREV = _compile_headers(SECTION_HEADERS_PREV)

def split_keep_lines(text: str) -> List[str]:
    t = normalize_text(text)
    lines = [ln.rstrip() for ln in t.splitlines()]
    return lines

def _is_header(line: str, hdrs: List[re.Pattern]) -> bool:
    s = strip_noise_line(line)
    return any(h.search(s) for h in hdrs)

def _is_bullet(line: str) -> bool:
    return bool(re.match(BULLET_PREFIX, line.strip()) or re.match(r"^\s*[\-·•▶▷\*]\s+", line.strip()))

def extract_section_bullets(text: str, which: str = "case") -> List[str]:
    # [섹션 파서] "사고사례/안전작업방법" 같은 제목 다음의 불릿을 모은다.
    lines = split_keep_lines(text)
    hdrs = HDR_CASE if which == "case" else HDR_PREV
    items: List[str] = []
    capture = False
    for i, raw in enumerate(lines):
        s = raw.strip()
        if not s:
            if capture: break
            continue
        if _is_header(raw, hdrs):
            capture = True
            continue
        if capture:
            if _is_header(raw, HDR_CASE + HDR_PREV):
                break
            clean = strip_noise_line(raw)
            if not clean:
                continue
            items.append(clean)
    merged = merge_broken_lines(items)
    return [x for x in merged if len(re.sub(r"\s+","", x)) >= 2]

# -------------------- (2) 헤더 없는 문서: 불릿 클러스터 + 자동 분류 --------------------
# - 행동 동사 패턴으로 예방형 분류, 사고 키워드/날짜로 사례형 분류
ACTION_VERBS = [
    "설치","배치","착용","점검","확인","측정","기록","표시","제공","비치","보고","신고",
    "교육","주지","중지","통제","휴식","환기","차단","교대","배제","배려","가동","준수",
    "운영","유지","교체","정비","청소","고정","격리","보호","보수","작성","지정"
]
ACTION_PAT = (
    r"(?P<obj>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(?P<verb>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"
    r"|(?P<obj2>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(을|를)\s*(?P<verb2>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"
)

def cluster_bullets(text: str, win: int = 1) -> List[List[str]]:
    lines = split_keep_lines(text)
    clusters: List[List[str]] = []
    cur: List[str] = []
    for ln in lines:
        if _is_bullet(ln):
            cur.append(strip_noise_line(ln))
        else:
            if cur:
                clusters.append(merge_broken_lines(cur))
                cur = []
    if cur:
        clusters.append(merge_broken_lines(cur))
    cleaned = []
    for c in clusters:
        c2 = [x for x in c if x and len(re.sub(r"\s+","", x)) >= 2]
        if c2:
            cleaned.append(c2)
    return cleaned

def looks_case(s: str) -> bool:
    return bool(re.search(r"(사망|사상|사고|중독|추락|붕괴|낙하|질식|끼임|깔림|부딪힘|감전|폭발)", s))

def looks_action(s: str) -> bool:
    return bool(re.search(ACTION_PAT, s) or re.search(r"(예방|수칙|지침|안전조치|작업방법)", s))

def classify_cluster(cluster: List[str]) -> str:
    case_hits = sum(1 for x in cluster if looks_case(x))
    act_hits  = sum(1 for x in cluster if looks_action(x))
    if case_hits > act_hits and case_hits >= 1:
        return "case"
    if act_hits >= max(1, case_hits):
        return "action"
    return "other"

def extract_clusters_by_type(text: str, kind: str) -> List[str]:
    clusters = cluster_bullets(text)
    out: List[str] = []
    for c in clusters:
        typ = classify_cluster(c)
        if typ == kind:
            out += c
    return out

# -------------------- PDF 읽기/진단 --------------------
def read_pdf_text_from_bytes(b: bytes, fname: str = "") -> str:
    # (1) pdfminer.six 텍스트 추출 → 실패 시 빈 문자열
    t = ""
    try:
        if pdf_extract_text is not None:
            with io.BytesIO(b) as bio:
                t = pdf_extract_text(bio) or ""
        else:
            t = ""
    except Exception:
        t = ""
    t = normalize_text(t)
    # (2) 텍스트가 거의 없을 때 pypdfium2로 간단 진단(이미지·스캔 추정)
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

# -------------------- 요약/임베딩 유사도 유틸 --------------------
def tokens_for_vec(s: str) -> List[str]:
    return tokens(s)

def sentence_tfidf_vectors(sents: List[str], kb_boost: Dict[str, float] = None) -> Tuple[np.ndarray, List[str]]:
    # [TF-IDF] 세션 KB에 있는 토큰은 가중치↑ → 문서/업로드 데이터에 민감한 요약
    toks = [tokens_for_vec(s) for s in sents]
    vocab: Dict[str,int] = {}
    for ts in toks:
        for t in ts:
            if t not in vocab: vocab[t] = len(vocab)
    if not vocab:
        return np.zeros((len(sents),0), dtype=np.float32), []
    M = np.zeros((len(sents), len(vocab)), dtype=np.float32)
    df = np.zeros((len(vocab),), dtype=np.float32)
    for i, ts in enumerate(toks):
        for t in ts:
            w = 1.0
            if kb_boost and t in kb_boost: w *= kb_boost[t]
            M[i, vocab[t]] += w
        for t in set(ts):
            df[vocab[t]] += 1.0
    N = float(len(sents))
    idf = np.log((N+1.0)/(df+1.0)) + 1.0
    M *= idf
    if kb_boost:
        for t, idx in vocab.items():
            if t in kb_boost: M[:, idx] *= (1.0 + 0.2*kb_boost[t])
    M /= (np.linalg.norm(M, axis=1, keepdims=True) + 1e-8)
    return M, list(vocab.keys())

def cosim(X: np.ndarray) -> np.ndarray:
    # 코사인 유사도 행렬(대각=0)
    if X.size == 0: return np.zeros((X.shape[0], X.shape[0]), dtype=np.float32)
    S = np.clip(X @ X.T, 0.0, 1.0); np.fill_diagonal(S, 0.0)
    return S

def textrank_scores(sents: List[str], X: np.ndarray, d: float=0.85, max_iter: int=60, tol: float=1e-4) -> List[float]:
    # [TextRank] 그래프 기반 전통 요약 — 링크 분석처럼 중요 문장 점수 계산
    n = len(sents)
    if n == 0: return []
    W = cosim(X); row = W.sum(axis=1, keepdims=True)
    P = np.divide(W, row, out=np.zeros_like(W), where=row>0)
    r = np.ones((n,1), dtype=np.float32)/n; tel = np.ones((n,1), dtype=np.float32)/n
    for _ in range(max_iter):
        r2 = d*(P.T @ r) + (1-d)*tel
        if np.linalg.norm(r2-r,1) < tol: r = r2; break
        r = r2
    return [float(v) for v in r.flatten()]

def mmr_select(sents: List[str], scores: List[float], X: np.ndarray, k: int, lam: float=0.7) -> List[int]:
    # [MMR] 다양성 제어 — 중복 줄이고 핵심만 뽑음
    S = cosim(X); sel: List[int] = []; rem = set(range(len(sents)))
    while rem and len(sel) < k:
        best, val = None, -1e9
        for i in rem:
            rel = scores[i]; div = max((S[i,j] for j in sel), default=0.0)
            sc = lam*rel - (1-lam)*div
            if sc > val: val, best = sc, i
        sel.append(best); rem.remove(best)
    return sel

def ai_extract_summary(text: str, limit: int=8) -> List[str]:
    # [무료 요약 AI] TextRank + MMR + 세션KB 가중치
    sents = preprocess_text_to_sentences(text)
    if not sents: return []
    kb = st.session_state["kb_terms"]; total = sum(kb.values()) or 1
    kb_boost = {t: 1.0 + (cnt/total)*3.0 for t, cnt in kb.items()} if kb else None
    X, _ = sentence_tfidf_vectors(sents, kb_boost=kb_boost)
    scores = textrank_scores(sents, X)
    idx = mmr_select(sents, scores, X, limit, lam=0.7)
    return [sents[i] for i in idx]

# -------------------- 도메인 템플릿/자연화 --------------------
def jaccard(a: set, b: set) -> float:
    return len(a & b) / (len(a | b) + 1e-8)

# - 특정 도메인(비계/양중/밀폐공간 등)에서 흔한 표현을 보수적으로 보정
DOMAIN_TEMPLATES = [
    ({"비계","발판","갱폼","추락"}, "작업발판을 견고하게 설치하고 안전난간 및 추락방호망을 확보합니다."),
    ({"안전난간","난간","개구부"}, "개구부·개구창 등 추락 위험 구간에 안전난간을 설치합니다."),
    ({"MSDS","국소배기","환기"}, "취급 물질의 MSDS를 확인하고 국소배기장치를 가동하여 충분히 환기합니다."),
    ({"예초","벌목","예초기"}, "예초·벌목 작업 시 작업자 간 안전거리를 유지하고 대피로를 확보합니다."),
    ({"크레인","양중"}, "양중 계획을 수립하고 신호수를 지정하여 통신을 유지합니다."),
    ({"컨베이어","협착","회전체"}, "회전체·물림점 접촉을 방지하도록 방호장치를 설치하고 점검합니다."),
]

def _domain_template_apply(s: str, base_text: str) -> str:
    if not st.session_state.get("domain_toggle"): return s
    sent_toks = set(tokens(s)); base_toks = set(tokens(base_text))
    if jaccard(sent_toks, base_toks) < 0.05: return s
    best = None; best_hits = 0
    for triggers, render in DOMAIN_TEMPLATES:
        if (sent_toks & triggers) and (base_toks & triggers):
            hits = len((sent_toks | base_toks) & triggers)
            if hits > best_hits: best_hits = hits; best = render
    return best if best else s

def soften(s: str) -> str:
    # 문체/종결統一 + 메타 토큰 제거
    s = s.replace("하여야","해야 합니다").replace("한다","합니다").replace("한다.","합니다.")
    s = s.replace("바랍니다","해주세요").replace("확인 바람","확인해주세요")
    s = s.replace("금지한다","금지합니다").replace("필요하다","필요합니다")
    s = re.sub(r"^\(([^)]+)\)\s*","", s)
    for pat in META_PATTERNS:
        s = re.sub(pat,"", s).strip()
    s = re.sub(BULLET_PREFIX,"", s).strip(" -•●\t")
    return tidy_korean_spaces(s)

def is_meaningful_sentence(s: str) -> bool:
    raw = re.sub(r"\s+","", s)
    if len(raw) < 4:
        return False
    if re.fullmatch(r"[가-힣\s]*합니다\.", s.strip()):
        return False
    return True

def is_accident_sentence(s: str) -> bool:
    if any(w in s for w in ["예방","대책","지침","수칙"]):
        return False
    return bool(re.search(DATE_PAT, s) or re.search(r"(사망|사상|사고|중독|추락|붕괴|낙하|질식|끼임|깔림|부딪힘|감전|폭발)", s))

def is_prevention_sentence(s: str) -> bool:
    return any(w in s for w in ["예방","대책","지침","수칙","안전조치","작업방법"]) or bool(re.search(ACTION_PAT, s))

def is_risk_sentence(s: str) -> bool:
    return any(w in s for w in ["위험","요인","원인","증상","결빙","강풍","폭염","미세먼지","회전체","비산","말림","추락","낙하","협착"])

def to_action_sentence(s: str, base_text: str) -> str:
    # [규칙형 NLG] 행동 동사/목적어 탐지 → 자연스러운 수칙 한 문장으로 보정
    s2 = soften(s)
    s2 = re.sub(r"(위기탈출\s*안전보건)", "", s2).strip()
    s2 = re.sub(r"\s*에\s*따른\s*", " 시 ", s2)
    s2 = re.sub(r"\s*에\s*따라\s*", " 시 ", s2)
    s2_tpl = _domain_template_apply(s2, base_text)
    if s2_tpl != s2:
        txt = s2_tpl
        if not txt.endswith(("다.","합니다.","습니다.")):
            txt = txt.rstrip(" .") + " 합니다."
        return tidy_korean_spaces(txt)
    m = re.search(ACTION_PAT, s2)
    if not m:
        # 동사 미탐지 시: 명사구 + 추정 동사로 보강
        nounish = re.sub(r"(의|에|에서|을|를|와|과|및)$","", s2).strip()
        if nounish and len(nounish) >= 4:
            guess_verb = "설치" if any(k in nounish for k in ["난간","방호망","발판","방호장치","장비","장치","표지"]) else "확인"
            obj = add_obj_particle(nounish)
            return tidy_korean_spaces(f"{obj} {guess_verb} 합니다.")
        txt = s2 if s2.endswith(("니다.","합니다.","다.")) else (s2.rstrip(" .") + " 합니다.")
        return tidy_korean_spaces(txt)
    obj = (m.group("obj") or m.group("obj2") or "").strip()
    verb = (m.group("verb") or m.group("verb2") or "실시").strip()
    if obj and not re.search(r"(을|를|에|에서|과|와|의)$", obj):
        obj = add_obj_particle(obj)
    prefix = "반드시 " if "설치" in verb else ("작업 전 " if verb in ("확인","점검","측정","기록","작성","지정") else "")
    core = tidy_korean_spaces(f"{prefix}{obj} {verb}")
    if re.fullmatch(r"(반드시 |작업 전 )?\s*(을|를)\s*(실시|관리|운영)\s*$", core):
        if obj.strip():
            core = tidy_korean_spaces(f"{prefix}{obj} 실시")
        else:
            core = "작업 전 안전조치 확인"
    return core.rstrip(" .") + " 합니다."

def repair_action_fragments(lines: List[str]) -> List[str]:
    # 짧은/불완전 수칙이 연속될 때 다음 줄과 합쳐 완전한 문장으로 회복
    out = []
    i = 0
    while i < len(lines):
        cur = soften(lines[i])
        cur_no_sp = re.sub(r"\s+","", cur)
        has_verb = bool(re.search(ACTION_PAT, cur)) or any(v in cur for v in ["합니다","한다","실시","설치","착용","점검","확인","배치"])
        if (len(cur_no_sp) < 20) and (not has_verb):
            merged = cur
            j = i + 1
            while j < len(lines):
                nxt = soften(lines[j])
                merged = tidy_korean_spaces(merged + " " + nxt)
                if re.search(ACTION_PAT, merged) or any(v in merged for v in ["합니다","한다","실시","설치","착용","점검","확인","배치"]):
                    break
                j += 1
            out.append(merged); i = j + 1
        else:
            out.append(cur); i += 1
    return out

# -------------------- KB(세션 동적 “경량 학습”) --------------------
def seed_kb_once():
    # 시드 KB는 1회만 주입 (세션 기준)
    if not st.session_state["seed_loaded"]:
        for t, k in SEED_RISK_MAP.items():
            if t not in RISK_KEYWORDS: RISK_KEYWORDS[t] = k
        for a in SEED_ACTIONS:
            if 2 <= len(a) <= 160:
                st.session_state["kb_actions"].append(a if a.endswith(("다","다.","합니다","합니다.")) else a + " 합니다.")
        for q in SEED_QUESTIONS:
            st.session_state["kb_questions"].append(q if q.endswith("?") else q + "?")
        for t in SEED_RISK_MAP.keys():
            st.session_state["kb_terms"][t] += 5
        st.session_state["seed_loaded"] = True

def kb_ingest_text(text: str) -> None:
    # 업로드 텍스트에서 위험어/행동문/점검질문을 추출하여 세션 KB에 축적
    if not (text or "").strip(): return
    sents = preprocess_text_to_sentences(text)
    for s in sents:
        for t in tokens(s):
            if len(t) >= 2:
                st.session_state["kb_terms"][t] += 1
                if re.search(r"(추락|낙하|깔림|끼임|중독|질식|화재|폭발|감전|폭염|붕괴|비계|갱폼|예초|벌목|컨베이어|크레인|지붕|선반|천공|화학물질|밀폐공간)", t):
                    if t not in RISK_KEYWORDS: RISK_KEYWORDS[t] = t
    action_candidates = [s for s in sents if (re.search(ACTION_PAT, s) or is_prevention_sentence(s))]
    action_candidates = repair_action_fragments(action_candidates)
    for s in action_candidates:
        cand = to_action_sentence(s, text)
        if 2 <= len(cand) <= 180:
            st.session_state["kb_actions"].append(cand)
    for s in sents:
        if "?" in s or "확인" in s or "점검" in s:
            q = soften(s if s.endswith("?") else s + " 맞습니까?")
            if 2 <= len(q) <= 160:
                st.session_state["kb_questions"].append(q)

def kb_prune() -> None:
    # 세션 KB 폭주 방지 — 중복 제거 및 상한 적용
    def dedup_keep_order(lst: List[str]) -> List[str]:
        seen, out = set(), []
        for x in lst:
            k = re.sub(r"\s+","", x)
            if k not in seen:
                seen.add(k); out.append(x)
        return out
    st.session_state["kb_actions"] = dedup_keep_order(st.session_state["kb_actions"])[:2000]
    st.session_state["kb_questions"] = dedup_keep_order(st.session_state["kb_questions"])[:800]
    st.session_state["kb_terms"] = Counter(dict(st.session_state["kb_terms"].most_common(4000)))

def kb_match_candidates(cands: List[str], base_text: str, limit: int, min_sim: float = 0.12) -> List[str]:
    # 문서 토큰과의 자카드 유사도로 세션 KB에서 적합한 수칙/질문을 선택
    bt = set(tokens(base_text))
    present_risks = {t for t in bt if (t in RISK_KEYWORDS or t in RISK_KEYWORDS.values())}
    scored: List[Tuple[float,str]] = []
    for c in cands:
        ct = set(tokens(c))
        cand_risks = {RISK_KEYWORDS.get(t, t) for t in ct if (t in RISK_KEYWORDS or t in RISK_KEYWORDS.values())}
        if cand_risks and not (cand_risks & present_risks):
            continue
        j = len(bt & ct) / (len(bt | ct) + 1e-8)
        if j >= min_sim:
            scored.append((j, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [c for _, c in scored[:limit]]

# -------------------- 사례/예방 자연화 보조 --------------------
def naturalize_case_sentence(s: str) -> str:
    s = soften(s)
    death = re.search(r"사망\s*(\d+)\s*명", s)
    inj = re.search(r"사상\s*(\d+)\s*명", s)
    unconscious = re.search(r"의식불명", s)
    info = []
    if death: info.append(f"근로자 {death.group(1)}명 사망")
    if inj and not death: info.append(f"{inj.group(1)}명 사상")
    if unconscious: info.append("의식불명 발생")
    m = re.search(DATE_PAT, s); date_txt=""
    if m:
        y, mo, d = m.groups()
        y = int(str(y).replace("’","").replace("'","")); y = 2000 + y if y < 100 else y
        date_txt = f"{int(y)}년 {int(mo)}월 {int(d)}일, "
        s = s.replace(m.group(0), "").strip()
    s = s.strip(" ,.-")
    if not re.search(r"(다\.|입니다\.|했습니다\.)$", s):
        if re.search(r"(사망|사상|중독|추락|낙하|붕괴|질식|끼임|깔림|부딪힘|감전|폭발)\s*$", s):
            s = s.rstrip(" .") + "했습니다."
        elif re.search(r"(사건|사고)\s*$", s):
            s = s.rstrip(" .") + "가 발생했습니다."
        else:
            s = s.rstrip(" .") + " 사고가 발생했습니다."
    if info and not s.endswith("했습니다."):
        s = tidy_korean_spaces(s.rstrip(" .") + " " + (", ".join(info)) + "했습니다.")
    return tidy_korean_spaces((date_txt + s).strip())

# -------------------- Fallback 추출기(헤더 없을 때) --------------------
def fallback_extract_cases(text: str, sents: List[str]) -> List[str]:
    from_cluster = extract_clusters_by_type(text, "case")
    from_sents = [x for x in sents if is_accident_sentence(x)]
    pool = from_cluster + from_sents
    pool = stitch_case_blocks(pool)
    seen, out = set(), []
    for x in pool:
        k = re.sub(r"\s+","", x)
        if k not in seen:
            seen.add(k); out.append(x)
    return out

def fallback_extract_preventions(text: str, sents: List[str]) -> List[str]:
    from_cluster = extract_clusters_by_type(text, "action")
    from_sents = [x for x in sents if is_prevention_sentence(x)]
    pool = from_cluster + from_sents
    pool = repair_action_fragments(pool)
    norm = [to_action_sentence(x, text) for x in pool if is_meaningful_sentence(x)]
    seen, out = set(), []
    for x in norm:
        k = re.sub(r"\s+","", x)
        if k not in seen:
            seen.add(k); out.append(x)
    return out

# -------------------- 라벨링 --------------------
def drop_label_token(t: str) -> bool:
    if t in STOP_TERMS: return True
    for pat in LABEL_DROP_PAT:
        if re.match(pat, t): return True
    if t in {"소재","소재지","지역","장소","버스","영업소","업체","자료","키","메세지","명","안전보건"}:
        return True
    return False

def top_terms_for_label(text: str, k: int=3) -> List[str]:
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
    if not doc_cnt: return ["안전보건","교육"]
    commons = {"안전","교육","작업","현장","예방","조치","확인","관리","점검","가이드","지침"}
    cand = [(t, doc_cnt[t]) for t in doc_cnt if t not in commons and len(t) >= 2]
    if not cand: cand = list(doc_cnt.items())
    cand.sort(key=lambda x: x[1], reverse=True)
    return [t for t,_ in cand[:k]]

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

# -------------------- 요약/생성(LLM-FREE) --------------------
def ai_extract_summary(text: str, limit: int=8) -> List[str]:
    sents = preprocess_text_to_sentences(text)
    if not sents: return []
    kb = st.session_state["kb_terms"]; total = sum(kb.values()) or 1
    kb_boost = {t: 1.0 + (cnt/total)*3.0 for t, cnt in kb.items()} if kb else None
    X, _ = sentence_tfidf_vectors(sents, kb_boost=kb_boost)
    scores = textrank_scores(sents, X)
    idx = mmr_select(sents, scores, X, limit, lam=0.7)
    return [sents[i] for i in idx]

def make_structured_script(text: str, max_points: int=6) -> str:
    topic_label = dynamic_topic_label(text)

    core = [soften(s) for s in ai_extract_summary(text, max_points)] if max_points > 0 else []
    core_actions = [s for s in core if (re.search(ACTION_PAT, s) or is_prevention_sentence(s))]
    core_actions = repair_action_fragments(core_actions)

    # 1) 헤더 기반 섹션
    case_block_raw = extract_section_bullets(text, which="case")
    prev_block_raw = extract_section_bullets(text, which="prev")

    # 2) 헤더 없을 때 클러스터/문장 기반 추출
    sents_all = preprocess_text_to_sentences(text)
    if not case_block_raw:
        case_block_raw = fallback_extract_cases(text, sents_all)
    if not prev_block_raw:
        prev_block_raw = fallback_extract_preventions(text, sents_all)

    cases_block = [naturalize_case_sentence(s) for s in case_block_raw if is_meaningful_sentence(s)]
    prev_block_raw = repair_action_fragments(prev_block_raw)
    prev_block = [to_action_sentence(s, text) for s in prev_block_raw if is_meaningful_sentence(s)]

    # 3) 요약 보조
    case_aux, risk_aux, ask_aux = [], [], []
    for s in core:
        if is_accident_sentence(s): case_aux.append(naturalize_case_sentence(s))
        elif is_risk_sentence(s):   risk_aux.append(soften(s))
        elif ("?" in s or "확인" in s or "점검" in s):
            ask_aux.append(soften(s if s.endswith("?") else s + " 맞습니까?"))

    # 4) 행동 보조
    act_aux = [to_action_sentence(s, text) for s in core_actions if is_meaningful_sentence(s)]

    # 5) KB 보강(예방 부족 시 무제한 확장)
    acts = prev_block + act_aux
    if len(acts) < 3 and st.session_state["kb_actions"]:
        acts += kb_match_candidates(st.session_state["kb_actions"], text, 8, min_sim=0.10)

    def uniq_keep(seq: List[str]) -> List[str]:
        seen, out = set(), []
        for x in seq:
            k = re.sub(r"\s+","", x)
            if k not in seen:
                seen.add(k); out.append(x)
        return out

    cases = uniq_keep(cases_block + case_aux)   # 사례: 제한 없음
    risks  = uniq_keep(risk_aux)
    asks   = uniq_keep(ask_aux or kb_match_candidates(st.session_state["kb_questions"], text, 4, min_sim=0.10))
    acts   = uniq_keep(acts)                    # 예방: 전부 출력

    lines = []
    lines.append(f"🦺 TBM 교육대본 – {topic_label}\n")
    lines.append("◎ 도입")
    lines.append(f"오늘은 최근 발생한 '{topic_label.replace(' 재해예방','')}' 사고 사례를 중심으로, 우리 현장에서 같은 사고를 예방하기 위한 안전조치를 함께 살펴보겠습니다.\n")

    if cases:
        lines.append("◎ 사고 사례")
        for c in cases:
            lines.append(f"- {c}")
        lines.append("")

    if risks:
        lines.append("◎ 주요 위험요인")
        for r in risks:
            lines.append(f"- {r}")
        lines.append("")

    if acts:
        lines.append("◎ 예방조치 / 실천 수칙")
        for i, a in enumerate(acts, 1):
            lines.append(f"{i}️⃣ {a}")
        lines.append("")

    if asks:
        lines.append("◎ 현장 점검 질문")
        for q in asks:
            lines.append(f"- {q}")
        lines.append("")

    lines.append("◎ 마무리 당부")
    lines.append("예방조치는 '선조치 후작업'이 원칙입니다. 오늘 작업 전, 각 공정별 위험요인을 다시 한 번 점검하고 필요한 보호구와 안전조치를 반드시 준비합시다.")
    lines.append("◎ 구호")
    lines.append("“한 번 더 확인! 한 번 더 점검!”")
    return "\n".join(lines)

def make_concise_report(text: str, max_points: int=6) -> str:
    # [핵심요약] 실행 경로 — 기존 UI 유지
    sents = ai_extract_summary(text, max_points)
    sents = [soften(s) for s in sents if not re.match(r"(배포처|주소|홈페이지|VR|리플릿)", s)]
    cases_blk = [naturalize_case_sentence(s) for s in extract_section_bullets(text, "case")] or \
                [naturalize_case_sentence(s) for s in fallback_extract_cases(text, preprocess_text_to_sentences(text))]
    prev_blk  = [to_action_sentence(s, text) for s in repair_action_fragments(
                    extract_section_bullets(text, "prev") or fallback_extract_preventions(text, preprocess_text_to_sentences(text))
                 )]

    act_src = [s for s in sents if (not is_accident_sentence(s)) and (is_prevention_sentence(s) or re.search(ACTION_PAT, s))]
    act_src = repair_action_fragments(act_src)
    cases = [naturalize_case_sentence(s) for s in sents if is_accident_sentence(s)]
    risks  = [soften(s) for s in sents if (not is_accident_sentence(s)) and is_risk_sentence(s)]
    acts   = [to_action_sentence(s, text) for s in act_src]

    def uniq_keep(seq: List[str]) -> List[str]:
        seen, out = set(), []
        for x in seq:
            k = re.sub(r"\s+","", x)
            if k not in seen:
                seen.add(k); out.append(x)
        return out

    cases = uniq_keep(cases_blk + cases)[:6]
    risks  = uniq_keep(risks)[:6]
    acts   = uniq_keep(prev_blk + acts)[:12]

    topic = dynamic_topic_label(text)
    lines = [f"📄 핵심요약 — {topic}\n"]
    if cases:
        lines.append("【사고 개요】"); lines.append("자료에서 확인된 주요 사고는 다음과 같습니다.")
        for c in cases: lines.append(f"- {c}")
        lines.append("")
    if risks:
        lines.append("【주요 위험요인】"); lines.append("자료 전반에서 다음 요인이 반복적으로 나타났습니다.")
        for r in risks: lines.append(f"- {r}")
        lines.append("")
    if acts:
        lines.append("【예방/실천 요약】"); lines.append("현장에서 즉시 적용 가능한 핵심 수칙입니다.")
        for a in acts: lines.append(f"- {a}")
        lines.append("")
    if not (cases or risks or acts):
        lines.append("자료의 핵심을 간단히 정리하면 다음과 같습니다.")
        for s in sents: lines.append(f"- {s}")
    return "\n".join(lines)

# -------------------- DOCX 내보내기 --------------------
_XML_FORBIDDEN = r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
def _xml_safe(s: str) -> str:
    if not isinstance(s, str): s = "" if s is None else str(s)
    return rxx.sub(_XML_FORBIDDEN, "", s)

def to_docx_bytes(script: str) -> bytes:
    # python-docx: 말굽(맑은 고딕) 기본 폰트 적용, 줄바꿈/한글 안전 문자만 삽입
    doc = Document()
    try:
        style = doc.styles["Normal"]; style.font.name = "Malgun Gothic"; style.font.size = Pt(11)
    except Exception: pass
    for raw in script.split("\n"):
        line = _xml_safe(raw)
        p = doc.add_paragraph(line)
        for run in p.runs:
            try:
                run.font.name = "Malgun Gothic"; run.font.size = Pt(11)
            except Exception: pass
    bio = io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio.read()

# -------------------- UI(기존 구성 유지) --------------------
with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 파이프라인(완전 무료, 오픈소스만 사용)**  
1) 전처리(노이즈 제거/줄 병합/날짜-사고 결합)  
2) **사례 블록 병합**(연결어·키워드로 연속 서술을 한 문장으로)  
3) **헤더 없을 때도** 불릿 클러스터 자동 분류(사례형/예방형)  
4) TextRank + MMR 요약 (**세션 KB 가중치** 반영)  
5) 규칙형 NLG: 조사/띄어쓰기·종결 보정, **예방조치 줄결합 및 문맥 보정**  
6) 결과 포맷: **자연스러운 교육대본(무료)** / **핵심요약**  
*NEW: 헤더가 없어도 사례·예방을 자동 수집합니다(섹션 파서 + 클러스터 + Fallback).*
""")
    st.session_state["domain_toggle"] = st.toggle(
        "🔧 도메인 템플릿 강화(신중 적용)",
        value=False,
        help="문장·본문 트리거 일치 + 유사도 기준 충족 시에만 템플릿을 적용합니다."
    )

seed_kb_once()
st.title("🦺 OPS/포스터를 교육 대본으로 자동 변환 (완전 무료)")

def reset_all():
    # 전체 초기화(세션 KB/캐시/업로더 키) — UI는 그대로
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

col_top1, col_top2 = st.columns([4,1])
with col_top2:
    st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

st.markdown("**안내**  \n- 텍스트가 포함된 PDF 또는 본문 텍스트를 권장합니다.  \n- 이미지/스캔 PDF는 현재 OCR 미지원입니다.")

col1, col2 = st.columns([1,1], gap="large")

with col1:
    uploaded = st.file_uploader(
        "OPS 업로드 (PDF 또는 ZIP) • 텍스트 PDF 권장",
        type=["pdf","zip"],
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
                            data = zf.read(name); zip_pdfs[name] = data
                if zip_pdfs:
                    for nm, data in zip_pdfs.items():
                        txt_all = read_pdf_text_from_bytes(data, fname=f"{fname}::{nm}")
                        if txt_all.strip():
                            kb_ingest_text(txt_all)
                    kb_prune()
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
                kb_ingest_text(extracted); kb_prune()
                st.session_state["edited_text"] = extracted
                st.session_state["last_extracted_cache"] = extracted
            else:
                st.warning("⚠️ PDF에서 유효한 텍스트를 추출할 수 없습니다.")
        else:
            st.warning("지원하지 않는 형식입니다. PDF 또는 ZIP을 업로드하세요.")

    pasted = (manual_text or "").strip()
    if pasted:
        kb_ingest_text(pasted); kb_prune()
        st.session_state["edited_text"] = pasted
        st.session_state["last_extracted_cache"] = pasted

    base_text = st.session_state.get("edited_text","")
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

with col2:
    gen_mode = st.selectbox("🧠 생성 모드", ["핵심요약","자연스러운 교육대본(무료)"])
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

st.caption("완전 무료(LLM 미사용). 헤더 유무 상관없이 사례/예방을 자동 탐지합니다(섹션 파서+클러스터+Fallback).")
