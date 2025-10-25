# =========================
# OPS2TBM (AI TextRank + 템플릿 자동선택 + 역할별 파서)
# - 텍스트 PDF / 텍스트 입력 지원
# - 이미지/스캔 PDF는 미지원(안내)
# - About(시연 멘트) 사이드바 유지
# - 템플릿: 자동/사고사례형/가이드형 (수동 선택 가능)
# =========================

import io
import re
from typing import List, Tuple, Dict

import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text
import regex as rxx
import networkx as nx

# ----------------------------
# 전처리 & 문장 처리
# ----------------------------
def clean_text(text: str) -> str:
    if not text:
        return ""
    # 비가시문자/공백 정리
    text = text.replace("\ufeff", " ")
    text = re.sub(r"[ \t]+", " ", text)
    # 표 캡션/출처/페이지 꼬리표 제거(너무 공격적이면 줄이거나 주석)
    text = re.sub(r"(출처|자료|작성|페이지)\s*[:：].*", "", text, flags=re.IGNORECASE)
    # 중복 줄바꿈 축약
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()

def split_sentences_ko(text: str) -> List[str]:
    # 문장 경계 + 줄바꿈 기준
    sents = rxx.split(r"(?<=[\.!?…]|다\.|다!|다\?)\s+|\n", text)
    sents = [s.strip(" -•●▪▶▷\t") for s in sents if len(s.strip()) > 3]
    return sents

def simple_tokenize_ko(s: str) -> List[str]:
    s = rxx.sub(r"[^0-9A-Za-z가-힣]", " ", s)
    return [t for t in s.split() if len(t) >= 2]

STOPWORDS = set(["그리고","그러나","하지만","또는","또한","등","및","이후","이전","사용","경우","관련","대한","위해","우리","등의","등을","해당"])

def jaccard_sim(a: List[str], b: List[str]) -> float:
    sa, sb = set([t for t in a if t not in STOPWORDS]), set([t for t in b if t not in STOPWORDS])
    if not sa or not sb:
        return 0.0
    inter = len(sa & sb)
    union = len(sa | sb)
    return inter / union if union else 0.0

# ----------------------------
# 키워드 & 상수
# ----------------------------
KW_OVERVIEW = ["개요","사례","사고","배경","요약","현황"]
KW_CAUSE    = ["원인","이유","문제점","부적정","미비","위험요인"]
KW_RULES    = ["예방","대책","수칙","점검","조치","확인","준수","관리","설치","착용","배치","통제"]

KW_GUIDE_CORE = ["가이드","안내","보호","건강","대응","절차","지침","매뉴얼","예방교육","상담","지원"]
KW_ROLES      = ["사업주","근로자","노동자","고객","제3자","관리자","감시자","책임자","담당자"]
KW_FLOW       = ["대응절차","절차","신고","조치","보고","상담","치료","휴식","업무중단","전환"]

SAFETY_CONSTANTS = [
    "선조치 후작업(안전설비·난간·라이프라인 설치 후 작업)",
    "감시자 배치 및 위험구역 출입 통제",
    "개인보호구 착용(안전모·안전벨트·안전화 등) 철저",
    "작업계획서·위험성평가 사전 검토 및 TBM 공유",
    "추락·협착 등 고위험 작업 시 작업중지 기준 숙지",
]

ROOF_EXTRAS = [
    "투광판(썬라이트) 위 절대 밟지 않기(취약부 표시)",
    "지붕 작업 시 안전발판·난간·라이프라인·추락방지망 설치",
    "기상(강풍·우천) 불량 시 작업 중지",
]

# ----------------------------
# 불릿/역할/절차 파서
# ----------------------------
BULLET_PAT = r"^[\s]*([\-•●▪▶▷]|\d+\)|\(\d+\)|\d+\.)\s*(.+)$"

ROLE_HEADERS = [
    ("사업주", ["사업주","고용주","경영","관리감독자"]),
    ("근로자", ["근로자","노동자","작업자","종사자"]),
    ("고객",   ["고객","이용자","제3자"]),
]

FLOW_HEADERS = ["대응절차","절차","신고 절차","대응","처리 절차"]

def extract_bullets(block: str) -> List[str]:
    out = []
    for line in block.splitlines():
        m = re.match(BULLET_PAT, line.strip())
        if m:
            out.append(m.group(2).strip())
    # 불릿이 하나도 없으면 문장 단위로라도 분리
    if not out:
        out = [l.strip(" -•●▪▶▷") for l in block.splitlines() if len(l.strip()) > 2]
    # 너무 긴 줄 컷
    res = []
    for s in out:
        s = re.sub(r"\s{2,}", " ", s)
        if len(s) > 140:
            s = s[:137] + "…"
        res.append(s)
    return res

def extract_role_sections(text: str) -> Dict[str, List[str]]:
    # 문서에서 '사업주', '근로자', '고객'으로 시작하거나 콜론/은/는 으로 이어지는 블럭을 추출
    role_map = {k: [] for k,_ in ROLE_HEADERS}
    lines = text.splitlines()
    cur_role = None
    buf = []

    def flush():
        nonlocal buf, cur_role
        if cur_role and buf:
            bullets = extract_bullets("\n".join(buf))
            role_map[cur_role].extend(bullets)
        buf = []

    for line in lines:
        norm = line.strip()
        if not norm:
            continue
        # 역할 헤더 감지
        found = None
        for role, keys in ROLE_HEADERS:
            if any(norm.startswith(k) for k in keys) or any((k+"은" in norm or k+"는" in norm or k+":" in norm) for k in keys):
                found = role
                break
        if found:
            flush()
            cur_role = found
            # 헤더 라인에 붙은 내용도 같이 보관
            rest = norm
            # 콜론 이후를 떼어 내용으로 포함
            if ":" in rest:
                rest = rest.split(":",1)[1].strip()
                if rest:
                    buf.append(rest)
            continue
        # 일반 라인
        if cur_role:
            buf.append(norm)
    flush()
    # 공백 제거
    for k in list(role_map.keys()):
        role_map[k] = [s for s in role_map[k] if s]
    return role_map

def extract_flow_section(text: str) -> List[str]:
    # 대응절차/절차 등 키워드 인근 블럭 추출
    lines = text.splitlines()
    buf = []
    capture = False
    got = []
    for ln in lines:
        t = ln.strip()
        if any(h in t for h in FLOW_HEADERS):
            capture = True
            continue
        if capture:
            if t == "" or re.match(r"^\s*[-=]+\s*$", t):
                continue
            # 섹션이 끝나는 신호(다른 큰 헤더)
            if any(key in t for _,keys in ROLE_HEADERS for key in keys):
                break
            got.append(t)
    if not got:
        return []
    return extract_bullets("\n".join(got))

# ----------------------------
# 규칙 기반 & AI(TextRank) 추출
# ----------------------------
def pick_sentences_rule(sents: List[str], keywords: List[str], limit: int) -> List[str]:
    scored = []
    for s in sents:
        score = sum(1 for k in keywords if k in s)
        if score > 0:
            scored.append((score, len(s), s))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [s for _, _, s in scored[:limit]]

def textrank_scores(sents: List[str]) -> List[float]:
    if not sents:
        return []
    tokens = [simple_tokenize_ko(s) for s in sents]
    g = nx.Graph()
    g.add_nodes_from(range(len(sents)))
    for i in range(len(sents)):
        for j in range(i+1, len(sents)):
            w = jaccard_sim(tokens[i], tokens[j])
            if w > 0:
                g.add_edge(i, j, weight=w)
    if g.number_of_edges() == 0:
        return [1.0] * len(sents)
    pr = nx.pagerank(g, weight="weight")
    return [pr.get(i, 0.0) for i in range(len(sents))]

def pick_sentences_tr(sents: List[str], kw: List[str], limit: int, scores: List[float]) -> List[str]:
    ranked = []
    for idx, s in enumerate(sents):
        kscore = 1 + sum(1 for k in kw if k in s) * 0.3
        ranked.append((scores[idx] * kscore, len(s), s))
    ranked.sort(key=lambda x: (-x[0], x[1]))
    return [s for _, _, s in ranked[:limit]]

# ----------------------------
# 템플릿 자동 선택
# ----------------------------
def detect_template(text: str) -> str:
    # 역할/가이드 신호가 많으면 'guide', 아니면 'accident'
    t = text
    role_hits = sum(t.count(k) for _, keys in ROLE_HEADERS for k in keys)
    guide_hits = sum(t.count(k) for k in (KW_GUIDE_CORE + KW_FLOW))
    accident_hits = sum(t.count(k) for k in (KW_OVERVIEW + KW_CAUSE + KW_RULES))
    if role_hits + guide_hits > accident_hits * 0.8:
        return "guide"
    return "accident"

# ----------------------------
# TBM 생성(사고사례형 / 가이드형)
# ----------------------------
def make_tbm_script_accident(raw_text: str, use_ai: bool) -> Tuple[str, dict]:
    text = clean_text(raw_text)
    sents = split_sentences_ko(text)
    if use_ai:
        scores = textrank_scores(sents)
        overview = pick_sentences_tr(sents, KW_OVERVIEW, 3, scores)
        causes   = pick_sentences_tr(sents, KW_CAUSE,   4, scores)
        rules    = pick_sentences_tr(sents, KW_RULES,   6, scores)
    else:
        overview = pick_sentences_rule(sents, KW_OVERVIEW, 3)
        causes   = pick_sentences_rule(sents, KW_CAUSE,   4)
        rules    = pick_sentences_rule(sents, KW_RULES,   6)
    if len(rules) < 4:
        rules = rules + SAFETY_CONSTANTS[: (4 - len(rules)) + 1]
    rules = rules[:6]
    # 타이틀 추정
    title = "OPS 기반 안전 TBM"
    for cand in sents[:5]:
        if any(k in cand for k in ["지붕","추락","질식","화재","협착","감전","질환","유해","중독"]):
            title = cand.strip(" .")
            if len(title) > 22:
                title = title[:20] + "…"
            break
    closing = [
        "‘잠깐이면 돼’는 가장 위험한 말입니다. 애매하면 멈추고 점검합시다.",
        "오늘 작업 전, 취약부·안전설비·PPE·감시자 여부를 다시 확인합시다.",
    ]
    chant = "한 번 더 점검! 모두가 안전!"
    # 스크립트 합성
    lines = []
    lines.append(f"🦺 TBM 대본 – 「{title}」\n")
    lines.append("◎ 인사 및 도입\n- OPS 자료를 바탕으로 재해 위험요인을 짚고, 우리 현장에서 바로 적용할 수 있는 안전수칙을 공유합니다.\n")
    lines.append("◎ 1. 사고 개요")
    for s in (overview or ["(OPS에서 개요를 찾지 못했습니다. 일반 개요로 대체합니다.)"]):
        lines.append(f"- {s}")
    lines.append("\n◎ 2. 사고 원인")
    for s in (causes or ["작업계획 부재","보호구 미착용","감시자 부재"]):
        lines.append(f"- {s}")
    lines.append("\n◎ 3. 주요 안전수칙(우리 현장 적용)")
    for s in rules:
        lines.append(f"- {s}")
    if any(k in text for k in ["지붕","썬라이트","투광판","추락"]):
        for s in ROOF_EXTRAS:
            lines.append(f"- {s}")
    lines.append("\n◎ 4. 마무리 당부")
    for s in closing:
        lines.append(f"- {s}")
    lines.append("\n◎ 마무리 구호")
    lines.append(f"- {chant}")
    script = "\n".join(lines).strip()
    parts = {"title": title,"overview": overview,"causes": causes,"rules": rules,"closing": closing,"chant": chant}
    return script, parts

def make_tbm_script_guide(raw_text: str, use_ai: bool) -> Tuple[str, dict]:
    text = clean_text(raw_text)
    sents = split_sentences_ko(text)
    # 핵심 메시지(요약)
    if use_ai:
        scores = textrank_scores(sents)
        core = pick_sentences_tr(sents, KW_GUIDE_CORE, 3, scores)
    else:
        core = pick_sentences_rule(sents, KW_GUIDE_CORE, 3)
    # 역할별 파싱
    roles = extract_role_sections(text)  # {"사업주":[...], "근로자":[...], "고객":[...]}
    # 대응절차
    flow = extract_flow_section(text)
    # 보강: 역할/절차가 비어있으면 규칙/상수로 메우기
    if sum(len(v) for v in roles.values()) == 0:
        # 키워드 기반으로라도 규칙 추출
        rules_guess = pick_sentences_rule(sents, KW_RULES + KW_GUIDE_CORE, 6)
        if not rules_guess:
            rules_guess = SAFETY_CONSTANTS[:4]
        # 역할 공통 섹션으로 합치기
        roles["사업주"] = [s for s in rules_guess[:2]]
        roles["근로자"] = [s for s in rules_guess[2:4]]
        roles["고객"]   = [s for s in rules_guess[4:6] or ["상호 존중과 배려 실천"]]
    if not flow:
        flow = ["상황 인지 즉시 보고 및 기록", "업무 일시중단·휴식 부여", "필요 시 상담·치료 지원 연계", "재발방지 대책 수립 및 공유"]
    # 타이틀
    title = "OPS 기반 안전 TBM(가이드)"
    for cand in sents[:5]:
        if any(k in cand for k in ["감정노동","건강보호","대응지침","고객응대","폭언","폭행","정신건강"]):
            title = cand.strip(" .")
            if len(title) > 22:
                title = title[:20] + "…"
            break
    closing = [
        "현장의 규정·절차를 따르고, 애매하면 즉시 보고·중지합니다.",
        "서로를 존중하는 말과 태도가 안전보건의 출발점입니다.",
    ]
    chant = "존중과 배려, 안전의 기본!"
    # 스크립트 합성(가이드형)
    lines = []
    lines.append(f"🦺 TBM 대본 – 「{title}」\n")
    lines.append("◎ 인사 및 도입\n- OPS 가이드의 핵심을 현장에서 바로 적용할 수 있도록 요약해 공유합니다.\n")
    lines.append("◎ 핵심 메시지")
    for s in (core or ["현장의 안전·건강보호를 위해 역할별 조치와 대응절차를 준수합니다."]):
        lines.append(f"- {s}")
    # 역할별
    for role in ["사업주","근로자","고객"]:
        if roles.get(role):
            lines.append(f"\n◎ {role} 수칙")
            for s in roles[role][:6]:
                lines.append(f"- {s}")
    # 대응절차
    lines.append("\n◎ 대응절차")
    for s in flow[:8]:
        lines.append(f"- {s}")
    # 마무리
    lines.append("\n◎ 마무리 당부")
    for s in closing:
        lines.append(f"- {s}")
    lines.append("\n◎ 마무리 구호")
    lines.append(f"- {chant}")
    script = "\n".join(lines).strip()
    parts = {"title": title,"core": core,"roles": roles,"flow": flow,"closing": closing,"chant": chant}
    return script, parts

# ----------------------------
# 내보내기
# ----------------------------
def to_txt_bytes(text: str) -> bytes:
    return text.encode("utf-8")

def to_docx_bytes(text: str) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Malgun Gothic'
    style.font.size = Pt(11)
    for line in text.split("\n"):
        doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ----------------------------
# PDF 텍스트 추출
# ----------------------------
def extract_text_from_pdf(file_bytes: bytes) -> str:
    try:
        return pdf_extract_text(io.BytesIO(file_bytes)) or ""
    except Exception:
        return ""

# ----------------------------
# 샘플 텍스트 (사고/가이드 2종)
# ----------------------------
SAMPLE_ACCIDENT = """2020년 2월, 지붕 썬라이트 위에서 작업 중 추락 재해가 발생. FRP 투광판의 노후로 파손 위험이 높음.
작업계획 미흡, 추락방지설비 미설치, 감시자 부재가 주요 원인.
예방을 위해 안전발판/난간/라이프라인 설치, 취약부 표시 및 출입통제, 감시자 배치 필요."""

SAMPLE_GUIDE = """감정노동 근로자 건강보호 안내. 고객의 폭언·폭행 등으로 인한 건강장해 예방과 대응절차를 제시.
사업주는 고객응대업무 지침 마련과 예방교육, 상담·치료 지원을 해야 함.
근로자는 건강장해 발생 우려 시 조치를 요구할 수 있음.
고객은 반말·욕설·무리한 요구를 자제하고 존중해야 함.
폭언 발생 시 대응절차: 중지 요청 → 책임자 보고 → 기록/증거 확보 → 휴식/상담·치료 지원 → 재발방지 대책 수립."""

# ----------------------------
# UI
# ----------------------------
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

# Sidebar: About / 시연 멘트
with st.sidebar:
    st.header("ℹ️ 소개 / 시연 멘트")
    st.markdown("""
**문제**  
OPS 문서를 현장 TBM 대본으로 바로 쓰기 어렵습니다.

**해결**  
문서 내용을 자동 분석하여,  
- **사고사례형**(개요/원인/수칙) 또는  
- **가이드형**(핵심/역할별 수칙/대응절차)  
TBM으로 변환합니다.

**시연 흐름**  
1) 파일 업로드(텍스트 PDF) 또는 텍스트 붙여넣기  
2) **모드 선택: 기본/AI(TextRank)** + **템플릿(자동/수동)**  
3) 대본 생성 → 섹션별 확인  
4) `.docx` 다운로드

**현재 버전**  
- OCR 미포함(클라우드 안정화) → 스캔본은 텍스트 붙여넣기 사용  
- 🔹 AI: TextRank 요약(그래프 기반 문장 랭킹)
""")

st.title("🦺 OPS2TBM — OPS/포스터를 TBM 대본으로 자동 변환")

st.markdown("""
**사용법**  
1) **텍스트가 포함된 PDF 업로드** 또는 **OPS 텍스트 붙여넣기**  
2) **모드(기본/AI)** 와 **템플릿(자동/사고사례형/가이드형)** 선택  
3) **TBM 대본 생성** → **.txt / .docx** 다운로드

> ⚠️ 이미지/스캔 PDF는 현재 OCR 미지원입니다. 그 경우 텍스트를 붙여넣어 주세요.
""")

col1, col2 = st.columns([1, 1])

with col1:
    uploaded = st.file_uploader("OPS 파일 업로드 (PDF만 지원)", type=["pdf"])
    manual_text = st.text_area("또는 OPS 텍스트 직접 붙여넣기", height=180, placeholder="OPS 본문을 붙여넣으세요...")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("사고 샘플", use_container_width=True):
            manual_text = SAMPLE_ACCIDENT
    with c2:
        if st.button("가이드 샘플", use_container_width=True):
            manual_text = SAMPLE_GUIDE
    with c3:
        st.write("")

    extracted = ""
    if uploaded:
        with st.spinner("PDF 텍스트 추출 중... (텍스트 PDF만 지원)"):
            data = uploaded.read()
            extracted = extract_text_from_pdf(data)
            if not extracted.strip():
                st.warning("이 PDF는 **이미지/스캔**일 가능성이 큽니다. 우측의 텍스트 입력으로 진행해 주세요.")

    base_text = manual_text.strip() if manual_text.strip() else extracted.strip()

    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=240)

with col2:
    use_ai = st.toggle("🔹 AI 요약 모드(TextRank) 켜기", value=True)
    template_mode = st.selectbox("🧩 템플릿 선택", ["자동 선택","사고사례형","가이드형"])

    if st.button("🛠️ TBM 대본 생성", type="primary", use_container_width=True):
        if not edited_text.strip():
            st.warning("텍스트가 비어 있습니다. PDF 업로드(텍스트 PDF) 또는 텍스트 입력 후 시도하세요.")
        else:
            # 템플릿 결정
            if template_mode == "자동 선택":
                detected = detect_template(edited_text)
            elif template_mode == "사고사례형":
                detected = "accident"
            else:
                detected = "guide"

            if detected == "accident":
                script, parts = make_tbm_script_accident(edited_text, use_ai=use_ai)
                subtitle = "사고사례형 템플릿 적용"
            else:
                script, parts = make_tbm_script_guide(edited_text, use_ai=use_ai)
                subtitle = "가이드형 템플릿 적용"

            st.success(f"대본 생성 완료! ({subtitle})")
            st.text_area("TBM 대본", value=script, height=420)

            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇️ .txt 다운로드", data=script.encode("utf-8"),
                                   file_name="tbm_script.txt", use_container_width=True)
            with c2:
                docx_bytes = to_docx_bytes(script)
                st.download_button("⬇️ .docx 다운로드", data=docx_bytes,
                                   file_name="tbm_script.docx", use_container_width=True)

st.caption("현재: 규칙 + TextRank 기반(경량 AI). 템플릿 자동/수동. 다음 단계: OCR 재도입·LLM 미세다듬기.")
