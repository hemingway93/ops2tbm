# ==========================================================
# OPS2TBM — 안전 교육 및 점검 일지 반영 (TBM 및 일일안전교육 일지 확장)
#  - 기존 UI/레이아웃 유지, 문서에 추가 필드 포함
# ==========================================================

import io, zipfile, re
from collections import Counter
from typing import List, Dict
import streamlit as st
from docx import Document
from docx.shared import Pt
from pdfminer.high_level import extract_text as pdf_extract_text  # ← 안정 경로
import pypdfium2 as pdfium
import numpy as np
import regex as rxx

# ----------------------------
# 세션 상태
# ----------------------------
# 세션 상태 초기화 (업로드한 파일의 키값, 용어/행동/질문 용어 집합을 저장)
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "kb_terms" not in st.session_state:
    st.session_state.kb_terms: Counter = Counter()
if "kb_actions" not in st.session_state:
    st.session_state.kb_actions: List[str] = []
if "kb_questions" not in st.session_state:
    st.session_state.kb_questions: List[str] = []

# ==========================================================
# 전처리 및 텍스트 정리
# ==========================================================
# 잡음 패턴(불필요한 헤더, 링크 등을 제거하기 위한 정규식)
NOISE_PATTERNS = [
    r"^제?\s?\d{4}\s?[-.]?\s?\d+\s?호$",
    r"^(동절기\s*주요사고|안전작업방법|콘텐츠링크|책자\s*OPS|숏폼\s*OPS)$",
    r"^(포스터|책자|스티커|콘텐츠 링크)$",
    r"^(스마트폰\s*APP|중대재해\s*사이렌|산업안전포털|고용노동부)$",
    r"^https?://\S+$",
    r"^\(?\s*PowerPoint\s*프레젠테이션\s*\)?$",
    r"^안전보건자료실.*$",
]
# 불릿포인트 문자 패턴
BULLET_PREFIX = r"^[\s\-\•\●\▪\▶\▷\·\*\u25CF\u25A0\u25B6\u25C6\u2022\u00B7\u279C\u27A4\u25BA\u25AA\u25AB\u2611\u2713\u2714\u2716\u2794\u27A2]+"
# 날짜 패턴
DATE_PAT = r"(\d{4})\.\s?(\d{1,2})\.\s?(\d{1,2})\.?"
# 사고사례에서 나타날 수 있는 패턴
META_PATTERNS = [r"<\s*\d+\s*명\s*사망\s*>", r"<\s*\d+\s*명\s*사상\s*>", r"<\s*\d+\s*명\s*의식불명\s*>"]
# 불용어(구분된 용어로부터 제거할 단어들)
STOP_TERMS = set([
    "및","등","관련","사항","내용","예방","안전","작업","현장","교육","방법","기준","조치",
    "실시","확인","필요","경우","대상","사용","관리","점검","적용","정도","주의","중","전","후",
    "주요","사례","안전작업방법","포스터","동영상","리플릿","가이드","자료실","검색",
])

def normalize_text(t: str) -> str:
    """텍스트에서 불필요한 공백과 줄 바꿈을 정리"""
    t = t.replace("\x0c","\n")
    t = re.sub(r"[ \t]+\n","\n",t)
    t = re.sub(r"\n{3,}","\n\n",t)
    return t.strip()

def strip_noise_line(line: str) -> str:
    """라인에서 불필요한 문자를 제거하고 필요한 문장만 반환"""
    s = (line or "").strip()
    if not s: return ""
    s = re.sub(BULLET_PREFIX,"",s).strip()
    for pat in NOISE_PATTERNS:
        if re.match(pat,s,re.IGNORECASE):
            return ""
    s = re.sub(r"https?://\S+","",s).strip()
    s = s.strip("•●▪▶▷·-—–")
    return s

def merge_broken_lines(lines: List[str]) -> List[str]:
    """줄바꿈이 잘못된 문장들을 병합"""
    out=[]; buf=""
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
    """날짜와 다음 문장을 합쳐서 사건을 묶어주는 함수"""
    out=[]; i=0
    while i<len(lines):
        cur=strip_noise_line(lines[i])
        if re.search(DATE_PAT,cur) and (i+1)<len(lines):
            nxt=strip_noise_line(lines[i+1])
            if re.search(r"(사망|사상|사고|중독|화재|붕괴|질식|추락|깔림|부딪힘|무너짐)",nxt):
                m=re.search(DATE_PAT,cur); y,mo,d=m.groups()
                out.append(f"{int(y)}년 {int(mo)}월 {int(d)}일, {nxt}")
                i+=2; continue
        out.append(cur); i+=1
    return out

# ==========================================================
# PDF 텍스트 추출 (텍스트형/이미지형 판단)
# ==========================================================
def read_pdf_text(b: bytes) -> str:
    """PDF 파일에서 텍스트 추출"""
    try:
        with io.BytesIO(b) as bio:
            t = pdf_extract_text(bio) or ""
    except Exception:
        t = ""
    t = normalize_text(t)
    if len(t.strip())<10:
        try:
            with io.BytesIO(b) as bio:
                pdf = pdfium.PdfDocument(bio)
                if len(pdf)>0 and not t.strip():
                    st.warning("⚠️ 이미지/스캔 PDF로 보입니다. 현재 OCR 미지원.")
        except Exception:
            pass
    return t

# ==========================================================
# 요약(TextRank+MMR) — 세션KB 용어 가중치 반영
# ==========================================================
def tokens(s: str) -> List[str]:
    """간단 토큰화(한글/영문/숫자 2자 이상)"""
    return rxx.findall(r"[가-힣a-z0-9]{2,}", s.lower())

def sentence_tfidf_vectors(sents: List[str], kb_boost: Dict[str,float]=None):
    """문장 TF-IDF + KB 용어 가중치"""
    toks=[tokens(s) for s in sents]
    vocab={}
    for ts in toks:
        for t in ts:
            if t not in vocab: vocab[t]=len(vocab)
    if not vocab: return np.zeros((len(sents),0),dtype=np.float32), []
    M=np.zeros((len(sents),len(vocab)),dtype=np.float32)
    df=np.zeros((len(vocab),),dtype=np.float32)
    for i,ts in enumerate(toks):
        for t in ts:
            w = 1.0
            if kb_boost and t in kb_boost:
                w *= kb_boost[t]
            M[i,vocab[t]] += w
        for t in set(ts):
            df[vocab[t]] += 1.0
    N=float(len(sents))
    idf=np.log((N+1.0)/(df+1.0))+1.0
    M*=idf
    if kb_boost:
        for t,idx in vocab.items():
            if t in kb_boost:
                M[:,idx] *= (1.0 + 0.2*kb_boost[t])  # 살짝 추가 부스트
    M/= (np.linalg.norm(M,axis=1,keepdims=True)+1e-8)
    return M, list(vocab.keys())

def cosim(X: np.ndarray)->np.ndarray:
    """문장 간 코사인 유사도 계산"""
    if X.size==0: return np.zeros((X.shape[0],X.shape[0]),dtype=np.float32)
    S=np.clip(X@X.T,0.0,1.0); np.fill_diagonal(S,0.0); return S

def textrank_scores(sents: List[str], X: np.ndarray, d:float=0.85, max_iter:int=60, tol:float=1e-4)->List[float]:
    """TextRank 점수 계산(유사도 행렬 기반)"""
    n=len(sents)
    if n==0: return []
    W=cosim(X)
    row=W.sum(axis=1,keepdims=True); P=np.divide(W,row,out=np.zeros_like(W),where=row>0)
    r=np.ones((n,1),dtype=np.float32)/n; tel=np.ones((n,1),dtype=np.float32)/n
    for _ in range(max_iter):
        r2=d*(P.T@r)+(1-d)*tel
        if np.linalg.norm(r2-r,1)<tol: r=r2; break
        r=r2
    return [float(v) for v in r.flatten()]

def mmr_select(sents, scores, X, k:int, lam:float=0.7)->List[int]:
    """MMR로 정보성/비중복성 균형"""
    S=cosim(X); sel=[]; rem=set(range(len(sents)))
    while rem and len(sel)<k:
        best, val=None,-1e9
        for i in rem:
            rel=scores[i]; div=max((S[i,j] for j in sel), default=0.0)
            sc=lam*rel-(1-lam)*div
            if sc>val: val, best=sc, i
        sel.append(best); rem.remove(best)
    return sel

def ai_extract_summary(text:str, limit:int=8)->List[str]:
    """전처리 문장 → KB가중 요약 문장 top-k"""
    sents=preprocess_text_to_sentences(text)
    if not sents: return []
    kb = st.session_state.kb_terms
    total = sum(kb.values()) or 1
    kb_boost = {t: 1.0 + (cnt/total)*3.0 for t,cnt in kb.items()} if kb else None
    X,_=sentence_tfidf_vectors(sents, kb_boost=kb_boost)
    scores=textrank_scores(sents, X)
    idx=mmr_select(sents,scores,X,limit,lam=0.7)
    return [sents[i] for i in idx]

# ==========================================================
# 동적 주제 라벨 — 문서+KB 상위 핵심어 조합
# ==========================================================
def top_terms_for_label(text: str, k:int=3) -> List[str]:
    doc_cnt = Counter([t for t in tokens(text) if t not in STOP_TERMS])
    kb = st.session_state.kb_terms
    if kb:
        for t, c in kb.items():
            if t in STOP_TERMS: continue
            doc_cnt[t] += 0.2 * c  # KB는 약하게 합산
    if not doc_cnt: return ["안전보건","교육"]
    commons = {"안전","교육","작업","현장","예방","조치","확인","관리","점검"}
    cand = [(t,doc_cnt[t]) for t in doc_cnt if t not in commons and len(t)>=2]
    if not cand: cand = list(doc_cnt.items())
    cand.sort(key=lambda x:x[1], reverse=True)
    return [t for t,_ in cand[:k]]

def dynamic_topic_label(text: str) -> str:
    return " · ".join(top_terms_for_label(text, k=3))

# ==========================================================
# 리라이팅 — 사건/행동/질문
# ==========================================================
ACTION_VERBS = [
    "설치","배치","착용","점검","확인","측정","기록","표시","제공","비치",
    "보고","신고","교육","주지","중지","통제","휴식","환기","차단","교대","배제","배려",
    "가동","준수","운영","유지","교체","정비","청소","고정","격리","보호","보수"
]

ACTION_PAT = r"(?P<obj>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(?P<verb>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b|(?P<obj2>[가-힣a-zA-Z0-9·\(\)\[\]\/\-\s]{2,}?)\s*(을|를)\s*(?P<verb2>" + "|".join(ACTION_VERBS) + r"|실시|운영|관리)\b"

def soften(s:str)->str:
    """명령조/서술조를 완곡한 존댓말로"""
    s = s.replace("하여야","해야 합니다").replace("한다","합니다").replace("한다.","합니다.")
    s = s.replace("바랍니다","해주세요").replace("확인 바람","확인해주세요")
    s = s.replace("금지한다","금지합니다").replace("필요하다","필요합니다")
    s = re.sub(r"^\(([^)]+)\)\s*","",s)
    for pat in META_PATTERNS:
        s = re.sub(pat,"",s).strip()
    s = re.sub(BULLET_PREFIX,"",s).strip(" -•●\t")
    return s

def naturalize_case_sentence(s:str)->str:
    """사건문 형식(YYYY.MM.DD + 내용) → 자연어"""
    s = soften(s)
    m=re.search(DATE_PAT,s)
    date_txt=""
    if m:
        y,mo,d=m.groups()
        date_txt=f"{int(y)}년 {int(mo)}월 {int(d)}일, "
        s = s.replace(m.group(0),"").strip()
    s = s.strip(" ,.-")
    if not re.search(r"(다\.|입니다\.|했습니다\.|발생했습니다\.)$",s):
        s = s.rstrip(" .") + " 사고가 발생했습니다."
    return (date_txt+s).strip()

def to_action_sentence(s:str)->str:
    """'명사구 + 동사'를 존댓말 지시형으로 변환"""
    s2 = soften(s)
    m = re.search(ACTION_PAT, s2)
    if not m:
        return s2 if s2.endswith(("니다.","합니다.","다.")) else (s2.rstrip(" .") + " 합니다.")
    obj = (m.group("obj") or m.group("obj2") or "").strip()
    verb = (m.group("verb") or m.group("verb2") or "실시").strip()
    if obj and not re.search(r"(을|를|에|에서|과|와|의)$", obj):
        obj += "를"
    return f"{obj} {verb}합니다." if obj else (s2 if s2.endswith(("니다.","합니다.","다.")) else (s2.rstrip(" .")+" 합니다."))

# ==========================================================
# KB 구축/활용 — 업로드/붙여넣기 시 자동 누적
# ==========================================================
def kb_ingest_text(text: str):
    """문서에서 용어/행동/질문을 추출하여 세션 KB에 누적"""
    sents = preprocess_text_to_sentences(text)
    for s in sents:
        for t in tokens(s):
            if len(t)>=2:
                st.session_state.kb_terms[t]+=1
    for s in sents:
        if re.search(ACTION_PAT, s):
            cand = to_action_sentence(s)
            if len(cand)<=100:   # 너무 긴 문장 제외
                st.session_state.kb_actions.append(cand)
    for s in sents:
        if "?" in s or "확인" in s or "점검" in s:
            q = soften(s if s.endswith("?") else s + " 맞습니까?")
            if len(q)<=100:
                st.session_state.kb_questions.append(q)

def kb_prune():
    """중복 제거 및 상한 적용(세션 메모리 과도 사용 방지)"""
    def dedup_keep_order(lst: List[str]) -> List[str]:
        seen=set(); out=[]
        for x in lst:
            k=re.sub(r"\s+","",x)
            if k not in seen:
                seen.add(k); out.append(x)
        return out
    st.session_state.kb_actions = dedup_keep_order(st.session_state.kb_actions)[:500]
    st.session_state.kb_questions = dedup_keep_order(st.session_state.kb_questions)[:300]
    st.session_state.kb_terms = Counter(dict(st.session_state.kb_terms.most_common(1000)))

def kb_match_candidates(cands: List[str], base_text: str, limit:int) -> List[str]:
    """현재 문서 토큰 교집합 기반으로 KB 후보 선별(도메인 일치도↑)"""
    bt = set(tokens(base_text))
    scored=[]
    for c in cands:
        ct=set(tokens(c))
        j = len(bt & ct) / (len(bt | ct)+1e-8)
        if j>0:
            scored.append((j, c))
    scored.sort(key=lambda x:x[0], reverse=True)
    return [c for _,c in scored[:limit]]

# ==========================================================
# 대본 생성 — 도입/사고/위험/행동/질문/마무리
# ==========================================================
def make_structured_script(text:str, max_points:int=6)->str:
    topic_label = dynamic_topic_label(text)  # 동적 주제 라벨
    core = [soften(s) for s in ai_extract_summary(text, max_points)]
    if not core:
        return "본문이 충분하지 않아 대본을 생성할 수 없습니다."

    case, risk, act, ask = [], [], [], []
    for s in core:
        if re.search(DATE_PAT, s) or re.search(r"(사망|사상|사례|사고|의식불명|중독|붕괴|질식|추락|깔림|부딪힘|무너짐)", s):
            case.append(naturalize_case_sentence(s))
        elif any(k in s for k in ["위험","요인","원인","증상","결빙","강풍","화재","중독","질식","미세먼지","회전체","비산","말림"]):
            risk.append(soften(s))
        elif re.search(ACTION_PAT, s) or any(v in s for v in ACTION_VERBS):
            act.append(to_action_sentence(s))
        elif "?" in s or "확인" in s or "점검" in s:
            ask.append(soften(s if s.endswith("?") else s + " 맞습니까?"))

    # 부족분은 KB에서 동일 도메인 후보로 보강
    if len(act) < 5 and st.session_state.kb_actions:
        act += kb_match_candidates(st.session_state.kb_actions, text, 5 - len(act))
    act = act[:5]

    if not ask and st.session_state.kb_questions:
        ask = kb_match_candidates(st.session_state.kb_questions, text, 3)
    if not ask:
        ask = ["필요한 안전조치가 오늘 작업 범위에 맞게 준비되어 있습니까?"]

    lines=[]
    lines.append(f"🦺 TBM 교육대본 – {topic_label}\n")
    lines.append("◎ 도입"); lines.append(f"오늘은 {topic_label}의 핵심을 짚어보겠습니다.\n")
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
        for i,a in enumerate(act,1): lines.append(f"{i}️⃣ {a}")
        lines.append("")
    if ask:
        lines.append("◎ 현장 점검 질문")
        for q in ask: lines.append(f"- {q}")
        lines.append("")
    lines.append("◎ 마무리 당부")
    lines.append("안전은 한순간의 관심에서 시작됩니다. 오늘 작업 전 서로 한 번 더 확인합시다.")
    lines.append("◎ 구호"); lines.append("“한 번 더 확인! 한 번 더 점검!”")
    return "\n".join(lines)

# ==========================================================
# DOCX (특수문자 필터)
# ==========================================================
_XML_FORBIDDEN = r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]"
def _xml_safe(s:str)->str:
    if not isinstance(s,str): s = "" if s is None else str(s)
    return rxx.sub(_XML_FORBIDDEN,"",s)

def to_docx_bytes(script:str)->bytes:
    doc=Document()
    try:
        style=doc.styles["Normal"]; style.font.name="Malgun Gothic"; style.font.size=Pt(11)
    except Exception: pass
    for raw in script.split("\n"):
        line=_xml_safe(raw)
        p=doc.add_paragraph(line)
        for run in p.runs:
            try: run.font.name="Malgun Gothic"; run.font.size=Pt(11)
            except Exception: pass
    bio=io.BytesIO(); doc.save(bio); bio.seek(0); return bio.read()

# ==========================================================
# UI (그대로 유지)
# ==========================================================
st.set_page_config(page_title="OPS2TBM", page_icon="🦺", layout="wide")

with st.sidebar:
    st.header("ℹ️ 소개 / 사용법")
    st.markdown("""
**AI 파이프라인(완전 무료)**  
- 전처리 → TextRank+MMR 요약(세션 KB 가중) → 데이터 기반 리라이팅  
- PDF/ZIP/텍스트 모두 올리면 즉시 누적 학습됩니다(용어/행동/질문).
""")

st.title("🦺 OPS/포스터를 교육 대본으로 자동 변환 (완전 무료)")

def reset_all():
    """초기화(세션 KB/입력/업로더 키)"""
    st.session_state.pop("manual_text", None)
    st.session_state.pop("edited_text", None)
    st.session_state.pop("zip_choice", None)
    st.session_state.kb_terms = Counter()
    st.session_state.kb_actions = []
    st.session_state.kb_questions = []
    st.session_state.uploader_key += 1
    st.rerun()

col_top1, col_top2 = st.columns([4,1])
with col_top2: st.button("🧹 초기화", on_click=reset_all, use_container_width=True)

st.markdown("""
**안내**  
- 텍스트가 포함된 PDF 또는 본문 텍스트를 권장합니다.  
- 이미지/스캔 PDF는 현재 OCR 미지원입니다.
""")

col1, col2 = st.columns([1,1], gap="large")

with col1:
    uploaded = st.file_uploader("OPS 업로드 (PDF 또는 ZIP) • 텍스트 PDF 권장",
                                type=["pdf","zip"], key=f"uploader_{st.session_state.uploader_key}")
    manual_text = st.text_area("또는 OPS 텍스트 직접 붙여넣기", key="manual_text",
                               height=220, placeholder="예: 현장 안내문 또는 OPS 본문 텍스트…")

    extracted=""
    zip_pdfs: Dict[str,bytes] = {}
    selected_zip_pdf=None

    # ZIP 업로드 시: 내부 모든 PDF를 학습(KB) + 선택해서 미리보기 가능
    if uploaded and uploaded.name.lower().endswith(".zip"):
        try:
            with zipfile.ZipFile(uploaded,"r") as zf:
                for name in zf.namelist():
                    if name.lower().endswith(".pdf"):
                        data = zf.read(name)
                        zip_pdfs[name]=data
                        txt = read_pdf_text(data)
                        if txt.strip():
                            kb_ingest_text(txt)      # ZIP 전체 학습
            kb_prune()
        except Exception:
            st.error("ZIP 해제 오류. 파일을 확인해 주세요.")
        if zip_pdfs:
            selected_zip_pdf = st.selectbox("ZIP 내 PDF 선택", list(zip_pdfs.keys()), key="zip_choice")

    # 단일 PDF 업로드 시: 즉시 텍스트 추출 + KB 학습
    if uploaded and uploaded.name.lower().endswith(".pdf"):
        with st.spinner("PDF 텍스트 추출 중..."):
            data = uploaded.read()
            extracted = read_pdf_text(data)
            if extracted.strip():
                kb_ingest_text(extracted)   # 🔹 단일 PDF도 즉시 학습
                kb_prune()
    elif selected_zip_pdf:
        with st.spinner("ZIP 내부 PDF 텍스트 추출 중..."):
            extracted = read_pdf_text(zip_pdfs[selected_zip_pdf])

    # 붙여넣기 텍스트도 가볍게 학습 반영
    pasted = (st.session_state.get("manual_text") or "").strip()
    if pasted:
        kb_ingest_text(pasted); kb_prune()

    base_text = pasted or extracted.strip()
    st.markdown("**추출/입력 텍스트 미리보기**")
    edited_text = st.text_area("텍스트", value=base_text, height=240, key="edited_text")

with col2:
    use_ai = st.toggle("🔹 AI 요약(TextRank+MMR) 사용", value=True)
    tmpl_choice = st.selectbox("🧩 템플릿", ["자동 선택","사고사례형","가이드형"])  # 표시만 유지
    gen_mode = st.selectbox("🧠 생성 모드", ["TBM 기본(현행)","자연스러운 교육대본(무료)"])
    max_points = st.slider("요약 강도(핵심문장 개수)", 3, 10, 6)

    if st.button("🛠️ 대본 생성", type="primary", use_container_width=True):
        text_for_gen = (st.session_state.get("edited_text") or "").strip()
        if not text_for_gen:
            st.warning("텍스트가 비어 있습니다. PDF/ZIP 업로드 또는 텍스트 입력 후 시도하세요.")
        else:
            with st.spinner("대본 생성 중..."):
                if gen_mode == "자연스러운 교육대본(무료)":
                    script = make_structured_script(text_for_gen, max_points=max_points)
                    subtitle = "자연스러운 교육대본(무료)"
                else:
                    sents = ai_extract_summary(text_for_gen, max_points if use_ai else 6)
                    sents = [soften(s) for s in sents]
                    script = "\n".join([f"- {s}" for s in sents]) if sents else "텍스트에서 핵심 문장을 찾지 못했습니다."
                    subtitle = "TBM 기본(현행)"

            st.success(f"대본 생성 완료! ({subtitle})")
            st.text_area("대본 미리보기", value=script, height=420)
            c3,c4 = st.columns(2)
            with c3:
                st.download_button("⬇️ TXT 다운로드", data=_xml_safe(script).encode("utf-8"),
                                   file_name="tbm_script.txt", use_container_width=True)
            with c4:
                st.download_button("⬇️ DOCX 다운로드", data=to_docx_bytes(script),
                                   file_name="tbm_script.docx", use_container_width=True)

st.caption("완전 무료. 업로드/붙여넣기만 해도 누적 학습 → 요약 가중/행동/질문 보강. 동적 주제 라벨. UI 변경 없음.")
