# -*- coding: utf-8 -*-
"""
한독의약박물관 관람 후기 감성분석 대시보드 (Streamlit)
------------------------------------------------------------------------
- 크롤러(네이버/다음/구글맵/카카오맵)가 만든 후기 엑셀/CSV 를 업로드하면
  자동으로 채널을 인식하고, 감성분석 + 인터랙티브 차트 + 다운로드를 제공.
- xlwings/openpyxl/엑셀이 저장한 .xlsx 를 모두 안전하게 읽습니다.
- 필요 패키지: streamlit pandas altair openpyxl xlrd  (requirements.txt 참고)
"""

import io
import re
import zipfile
import datetime
import pandas as pd
import altair as alt
import streamlit as st

st.set_page_config(
    page_title="한독의약박물관 후기 감성분석",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------- 한국어 감성 사전 -----------------------
POSITIVE_WORDS = {
    # 기본 감정
    "좋": 2, "최고": 3, "훌륭": 3, "멋": 2, "예쁘": 2, "이쁘": 2, "아름": 2,
    "만족": 3, "추천": 3, "감동": 3, "인상": 2, "유익": 3, "알차": 3, "알찬": 3,
    "재밌": 3, "재미있": 3, "재미": 2, "즐겁": 3, "즐거": 2, "행복": 3, "친절": 3,
    "깔끔": 2, "깨끗": 2, "쾌적": 2, "편하": 2, "편리": 2, "넓": 1, "다양": 2,
    "볼거리": 2, "교육적": 3, "가성비": 2, "무료": 1, "가볼만": 3, "볼만": 2,
    "강추": 3, "짱": 2, "굿": 2, "좋아요": 2, "최애": 3, "감사": 2, "완벽": 3,
    "신기": 2, "흥미": 2, "정성": 2, "따뜻": 2, "힐링": 3, "인생": 2, "대박": 3,
    "잘": 1, "좋았": 2, "좋은": 2, "좋네": 2, "좋고": 2, "멋진": 2, "예쁜": 2,
    # 박물관/전시/체험 맥락 (확장)
    "체험": 2, "배우": 2, "배웠": 2, "배울": 2, "관람": 1, "구경": 1, "구성": 1,
    "설명": 1, "친절하게": 3, "자세": 1, "유익하": 3, "알차게": 3,
    "또가": 3, "또오": 3, "또방문": 3, "재방문": 3, "다시가": 2, "다시오": 2,
    "추천해": 3, "추천합니다": 3, "추천드": 3, "가보세요": 2, "가볼": 2,
    "아이가좋아": 3, "아이들이좋아": 3, "아이와": 1, "가족": 1, "나들이": 1,
    "볼거리가많": 3, "체험이많": 3, "프로그램": 1, "잘되어": 2, "잘꾸": 2,
    "잘정리": 2, "정리가잘": 2, "잘보존": 2, "인상적": 3, "의미있": 3,
    "뜻깊": 3, "기대이상": 3, "생각보다좋": 3, "만족스": 3, "즐길": 2,
    "즐겼": 2, "체험할": 2, "체험형": 2, "좋았습니다": 3, "좋았어요": 3,
    "좋더": 2, "괜찮": 2, "쏠쏠": 2, "유용": 3,
    # 추천/권유 표현 ("꼭 ~하세요", "들어야 한다" 등 = 추천 = 긍정)
    "꼭": 2, "들어야": 3, "들어보": 2, "들으면": 2, "들으시": 2,
    "들어보세요": 3, "들어보시": 3, "해보세요": 2, "하세요": 1, "해보시": 2,
    "가보시": 2, "보세요": 1, "추천드려": 3, "추천이": 2, "강력추천": 3,
    "필수": 2, "놓치지": 2, "꼭 들": 3, "해설": 1, "설명이 좋": 3,
    "설명을 잘": 3, "설명이 자세": 3, "해설사": 1, "도슨트": 1,
}
NEGATIVE_WORDS = {
    # 기본 감정
    "별로": -3, "실망": -3, "최악": -3, "불친절": -3, "지루": -3, "재미없": -3,
    "아쉽": -2, "아쉬": -2, "부족": -2, "협소": -2, "좁": -1, "낡": -2, "오래되": -1,
    "비싸": -2, "불편": -2, "더럽": -3, "지저분": -2, "불쾌": -3, "짜증": -3,
    "화나": -3, "후회": -3, "별거": -2, "볼거없": -3, "볼게없": -3,
    "허접": -3, "엉망": -3, "무성의": -3, "썰렁": -2, "노잼": -3, "낭비": -3,
    "안좋": -2, "나쁘": -3, "최악이": -3, "실망스": -3, "별루": -3,
    # 박물관/방문 맥락 (확장)
    "웨이팅": -2, "오래기다": -3, "주차힘": -2, "주차가불편": -3,
    "설명부족": -3, "볼게없": -3, "볼것이없": -3, "볼것도없": -3,
    "규모가작": -2, "생각보다작": -2, "기대이하": -3, "돈아깝": -3,
    "시간낭비": -3, "재미없": -3, "재미가없": -3, "심심": -2, "밋밋": -2,
    "관리안": -2, "관리가안": -3, "정비가안": -2, "찾기힘": -2, "찾기어렵": -2,
    "교통불편": -2,
}
NEGATORS = ["안", "못", "없", "않", "말고", "빼고"]
LABELS = ["긍정", "중립", "부정"]
COLOR = {"긍정": "#4CAF50", "중립": "#9E9E9E", "부정": "#F44336"}

# ----------------------- 보고서용: 불용어 & 개선 과제 테마 -----------------------
STOPWORDS = set("""
그리고 그런데 그래서 하지만 그러나 또한 정말 진짜 너무 매우 조금 많이 약간 완전
있는 있어요 있었어요 있고 없는 없어요 합니다 했어요 하고 해서 하는 되는 되어 된다
같아요 같은 이런 저런 그런 여기 저기 거기 우리 저희 제가 내가 그냥 아주 계속 다시
한독 의약 박물관 한독의약박물관 후기 방문 관람 곳 것 수 때 등 및
에서 에게 의 로 으로 처럼 보다 이라고 라고 하러 위해 대한
음성 서울 오늘 어제 이번 지난 정도 위치 근처 다녀 다녀왔 왔어요 갔어요 봤어요 였어요
이곳 저곳 하나 모두 함께 통해 관련 대해 경우 시간 요금 가격 입장
""".split())

SENT_SPLIT = re.compile(r"[.!?。\n]+")
WORD_RE = re.compile(r"[가-힣]{2,}|[A-Za-z]{3,}")

IMPROVEMENT_THEMES = {
    "주차/교통": (
        ["주차", "주차장", "차댈", "교통", "찾기힘", "찾기어렵", "접근성",
         "길찾", "멀어", "버스", "대중교통"],
        "주차 공간 확대·안내 강화, 대체 주차장/대중교통 이용 안내, 진입로 안내 개선"),
    "대기/혼잡": (
        ["웨이팅", "기다", "붐비", "혼잡", "사람많", "대기"],
        "체험 프로그램 사전 예약제 도입, 시간대별 인원 분산, 대기 현황 안내"),
    "전시 규모/볼거리": (
        ["규모", "볼거리", "볼거", "볼게", "단조", "심심", "밋밋", "금방", "빨리끝"],
        "체험형·참여형 콘텐츠 확충, 순환 기획전 강화, 포토존·인터랙션 요소 추가"),
    "해설/설명": (
        ["설명", "해설", "안내문", "정보부족", "이해", "어려", "설명서"],
        "정기 해설(도슨트) 시간 확대·안내, 설명 패널 보강, 오디오가이드 도입"),
    "시설/청결": (
        ["낡", "오래되", "더럽", "지저분", "화장실", "청소", "관리", "정비",
         "냄새", "에어컨", "난방", "의자", "벤치"],
        "노후 시설 보수·정비, 청결 관리 주기 강화, 냉난방·휴게 공간 개선"),
    "직원/응대": (
        ["불친절", "직원", "응대", "태도", "무성의", "안내원", "퉁명"],
        "현장 응대 매뉴얼 교육, 친절 서비스 모니터링, 안내 인력 배치 점검"),
    "요금/가치": (
        ["비싸", "돈아깝", "가격", "요금", "입장료", "가성비", "값"],
        "요금 대비 체험 가치 강화, 할인·패키지 안내, 무료 프로그램 홍보 확대"),
    "편의시설": (
        ["매점", "카페", "먹을", "식당", "수유", "유모차", "휠체어",
         "엘리베이터", "물품보관", "사물함"],
        "편의시설(카페·휴게·수유실 등) 확충 검토, 유모차·휠체어 접근성 개선"),
}

CHANNEL_HINTS = {
    "네이버": "네이버블로그", "naver": "네이버블로그",
    "다음": "다음블로그", "daum": "다음블로그",
    "구글": "구글맵", "google": "구글맵",
    "카카오": "카카오맵", "kakao": "카카오맵",
    "인스타": "인스타그램", "instagram": "인스타그램", "insta": "인스타그램",
    "통합": "통합블로그",
}


# ----------------------- 파일 읽기 (핵심) -----------------------
def _read_csv_any(file_bytes):
    """여러 인코딩으로 CSV 읽기 시도"""
    last = None
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
        except Exception as e:
            last = e
    raise last


def _is_zip(file_bytes):
    """실제로 zip(=xlsx 컨테이너)으로 열리는지 확인"""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            names = z.namelist()
        return any(n.startswith("xl/") for n in names) or "[Content_Types].xml" in names
    except Exception:
        return False


def read_any(file_bytes, filename):
    """
    업로드 파일을 안전하게 읽기 (xlwings/openpyxl/엑셀 저장본 모두 대응).
    핵심: xlsx 는 결국 zip 컨테이너이므로,
         '.xlsx 확장자' 또는 'zip으로 열림'이면 무조건 openpyxl 로 읽고
         절대 xlrd 로 넘어가지 않는다. (OLE2 오류 방지)
    """
    name = (filename or "").lower()

    # 1) 확장자가 xlsx/xlsm → 무조건 openpyxl (xlrd 안 씀)
    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")

    # 2) 확장자가 애매해도 내용이 zip이면 xlsx 로 간주
    if _is_zip(file_bytes):
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")

    # 3) 진짜 구형 xls (OLE2 시그니처) → xlrd
    if file_bytes[:4] == b"\xD0\xCF\x11\xE0" or name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")

    # 4) 그 외 → CSV (여러 인코딩)
    try:
        return _read_csv_any(file_bytes)
    except Exception:
        pass

    # 5) 최후 폴백
    return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")


@st.cache_data(show_spinner=False)
def read_file(file_bytes, filename):
    return read_any(file_bytes, filename)


# ----------------------- 감성 판정 -----------------------
def _count_negated(text, word):
    flips = 0
    for m in re.finditer(re.escape(word), text):
        pre = text[max(0, m.start() - 4):m.start()]
        if any(neg in pre for neg in NEGATORS):
            flips += 1
    return flips


def score_text(text):
    if not isinstance(text, str) or not text.strip():
        return 0
    t = text.replace(" ", "")
    score = 0
    for word, w in POSITIVE_WORDS.items():
        key = word.replace(" ", "")
        cnt = t.count(key)
        if cnt:
            flipped = _count_negated(text, key)
            score += w * (cnt - flipped) - w * flipped
    for word, w in NEGATIVE_WORDS.items():
        key = word.replace(" ", "")
        score += w * t.count(key)
    return score


def label_from_star(star):
    # 빈값/NaN/None/'None' 문자열은 '별점 없음'으로 처리 (중립 오판 방지)
    if star is None:
        return None
    try:
        s = float(star)
    except (TypeError, ValueError):
        return None
    if s != s:          # NaN 체크
        return None
    if s >= 4:
        return "긍정"
    if s <= 2:
        return "부정"
    return "중립"       # 별점 3점만 진짜 중립


def label_from_score(score):
    # 중립 구간 축소: 0보다 크면 긍정, 작으면 부정, 딱 0(감정어 없음)만 중립
    if score > 0:
        return "긍정"
    if score < 0:
        return "부정"
    return "중립"


def guess_channel(filename):
    low = filename.lower()
    for key, ch in CHANNEL_HINTS.items():
        if key in low:
            return ch
    return "기타"


def _get(row, *keys):
    for k in keys:
        if k in row and pd.notna(row[k]) and str(row[k]).strip():
            return str(row[k])
    return ""


def _source_to_channel(src):
    """통합 파일의 '출처' 값(네이버/다음)을 표준 채널명으로"""
    s = str(src)
    if "네이버" in s:
        return "네이버블로그"
    if "다음" in s:
        return "다음블로그"
    return None


def standardize(df, channel):
    is_merged = (channel == "통합블로그")
    has_src = "출처" in df.columns          # 통합 파일에 있는 출처 컬럼
    recs = []
    for _, r in df.iterrows():
        text = _get(r, "리뷰내용", "본문", "요약")
        title = _get(r, "제목")

        # 통합 파일이면 행의 '출처'로 실제 채널을 지정(네이버/다음 분리)
        ch = channel
        if is_merged and has_src:
            mapped = _source_to_channel(r.get("출처"))
            if mapped:
                ch = mapped

        recs.append({
            "채널": ch,
            "작성일": _get(r, "작성일"),
            "작성자": _get(r, "작성자", "블로그명", "블로그ID"),
            "별점": r["별점"] if "별점" in r and pd.notna(r["별점"]) else None,
            "제목": title,
            "내용": text,
            "_full": (title + " " + text).strip(),
            "URL": _get(r, "URL"),
        })
    return pd.DataFrame(recs)


def analyze(df):
    labels, bases, scores = [], [], []
    for _, r in df.iterrows():
        star_label = label_from_star(r["별점"])
        sc = score_text(r["_full"])
        scores.append(sc)
        if star_label is not None:
            labels.append(star_label)
            bases.append("별점")
        else:
            labels.append(label_from_score(sc))
            bases.append("본문분석")
    df = df.copy()
    df["감성"] = labels
    df["판정근거"] = bases
    df["감성점수"] = scores
    return df


# ----------------------- 보고서 생성 함수 -----------------------
def _tokenize(text):
    if not isinstance(text, str):
        return []
    out = []
    for w in WORD_RE.findall(text):
        if w in STOPWORDS:
            continue
        if re.search(r"(했어요|하고|해서|하는|하니|이라|이고|였어요|어요|았어요|"
                     r"네요|으며|으면|지만|는데|하며|하면|되어|되는|된다|스러|"
                     r"스럽|같아)$", w):
            continue
        w2 = re.sub(r"(은|는|이|가|을|를|에|의|도|로|으로|와|과|만|에서|에게)$", "", w)
        if len(w2) < 2 or w2 in STOPWORDS:
            continue
        out.append(w2)
    return out


def top_keywords(texts, n=15):
    from collections import Counter
    c = Counter()
    for t in texts:
        c.update(set(_tokenize(t)))
    return c.most_common(n)


def rep_sentences(pairs, n=5):
    cands, seen = [], set()
    for score, text in pairs:
        if not isinstance(text, str):
            continue
        for sent in SENT_SPLIT.split(text):
            s = sent.strip()
            if not (12 <= len(s) <= 90):
                continue
            if s[:20] in seen:
                continue
            seen.add(s[:20])
            cands.append((abs(score), s))
    cands.sort(key=lambda x: x[0], reverse=True)
    return [s for _, s in cands[:n]]


def analyze_improvements(neg_df):
    from collections import defaultdict
    theme_hits = defaultdict(list)
    for _, r in neg_df.iterrows():
        text = str(r.get("_full", "")) or (str(r.get("제목", "")) + str(r.get("내용", "")))
        ns = text.replace(" ", "")
        for theme, (keys, _a) in IMPROVEMENT_THEMES.items():
            if any(k.replace(" ", "") in ns for k in keys):
                theme_hits[theme].append(text)
    results = []
    for theme, texts in theme_hits.items():
        keys, advice = IMPROVEMENT_THEMES[theme]
        keys_ns = [k.replace(" ", "") for k in keys]

        def rel(t):
            n = t.replace(" ", "")
            return min((n.find(k) for k in keys_ns if k in n), default=9999)
        samples, seen = [], set()
        for t in sorted(texts, key=lambda x: (rel(x), len(x))):
            sn = t.strip()
            if 8 <= len(sn) <= 80 and sn[:15] not in seen:
                seen.add(sn[:15])
                samples.append(sn)
            if len(samples) >= 2:
                break
        if not samples and texts:
            samples = [texts[0][:80]]
        results.append((theme, len(texts), samples, advice))
    results.sort(key=lambda x: x[1], reverse=True)
    return results


# ----------------------- 주제 분석 (방문객이 주로 하는 이야기) -----------------------
TOPICS = {
    "약사 체험 프로그램": ["약사", "체험", "약만들", "약재체험", "프로그램",
                    "만들어", "조제", "체험형", "체험학습"],
    "해설·설명": ["해설", "도슨트", "설명", "안내문", "들으니", "이해", "가이드", "오디오"],
    "전시·유물 관람": ["전시", "유물", "약재", "의료기구", "역사", "신기",
                 "흥미", "볼거리", "소장품", "컬렉션", "작품"],
    "아이·가족 나들이": ["아이", "가족", "아이와", "아이랑", "데리고", "초등",
                  "유아", "어린이", "아기", "엄마", "아빠", "부모"],
    "건물·공간·사진": ["건물", "정원", "산책", "사진", "포토", "예쁘", "경치",
                 "조경", "분위기", "인테리어", "예쁜", "이쁜", "고풍"],
    "가성비·무료": ["무료", "가성비", "알차", "저렴", "공짜", "값싸"],
    "친절·응대": ["친절", "응대", "직원", "선생님", "안내원", "따뜻"],
    "주차·접근성": ["주차", "교통", "접근", "찾기", "길찾", "내비", "대중교통"],
    "규모·볼거리 아쉬움": ["규모작", "작고", "볼거리부족", "금방", "볼게없",
                   "단조", "심심", "협소"],
    "대기·혼잡": ["웨이팅", "대기", "기다", "붐비", "혼잡", "사람많"],
}


def extract_phrases(texts, top_n=12):
    """자주 나오는 2~3어절 핵심 문구 (데이터 적어도 안전)"""
    from sklearn.feature_extraction.text import CountVectorizer
    docs = [" ".join(_tokenize(t)) for t in texts]
    docs = [d for d in docs if d.strip()]
    if len(docs) < 2:
        return []
    min_df = 3 if len(docs) >= 40 else (2 if len(docs) >= 8 else 1)
    for md in (min_df, 2, 1):
        try:
            vec = CountVectorizer(ngram_range=(2, 3), min_df=md)
            X = vec.fit_transform(docs)
            if X.shape[1] == 0:
                continue
            freq = X.sum(axis=0).A1
            terms = vec.get_feature_names_out()
            order = freq.argsort()[::-1]
            out = []
            for i in order:
                if freq[i] < max(2, md):
                    continue
                out.append((terms[i], int(freq[i])))
                if len(out) >= top_n:
                    break
            if out:
                return out
        except (ValueError, Exception):
            continue
    return []


def analyze_topics(df):
    """주제 사전으로 후기 분류 → 주제별 건수·비율·감성·대표후기·핵심문구"""
    from collections import Counter, defaultdict
    total = len(df)
    if total == 0:
        return []
    topic_rows = defaultdict(list)
    for _, r in df.iterrows():
        text = str(r.get("_full", "") or "")
        ns = text.replace(" ", "")
        senti = str(r.get("감성", ""))
        for topic, keys in TOPICS.items():
            if any(k.replace(" ", "") in ns for k in keys):
                topic_rows[topic].append((senti, text))
    results = []
    for topic, rows in topic_rows.items():
        texts = [t for _, t in rows]
        sentis = Counter(s for s, _ in rows if s)
        rep = ""
        for t in sorted(texts, key=len):
            if 10 <= len(t) <= 70:
                rep = t
                break
        if not rep and texts:
            rep = texts[0][:70]
        phrases = extract_phrases(texts, top_n=3)
        results.append({
            "주제": topic, "건수": len(rows),
            "비율": round(len(rows) / total * 100, 1),
            "감성분포": dict(sentis), "대표후기": rep,
            "핵심문구": [p for p, _ in phrases],
        })
    results.sort(key=lambda x: x["건수"], reverse=True)
    return results


def senti_str(dist):
    order = ["긍정", "중립", "부정"]
    parts = [f"{k}{dist[k]}" for k in order if dist.get(k)]
    return " ".join(parts) if parts else "-"


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="감성분석결과")
    return buf.getvalue()


# ============================== UI ==============================
st.title("🏛️ 한독의약박물관 관람 후기 감성분석")
st.caption("네이버·다음 블로그, 구글맵·카카오맵 리뷰를 한 곳에서 분석합니다.")

with st.sidebar:
    st.header("📂 데이터 업로드")
    st.markdown("PC 크롤러가 만든 **후기 엑셀(.xlsx)** 또는 **.csv** 를 여러 개 올려주세요.")
    uploaded = st.file_uploader(
        "후기 파일 선택",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
    st.divider()
    st.subheader("ℹ️ 채널 자동 인식")
    st.caption("파일명에 '네이버/다음/구글/카카오'가 있으면 자동 분류됩니다.")
    st.divider()
    st.caption("분석 전용 대시보드 · 후기 수집은 PC 크롤러에서 진행")

if not uploaded:
    st.info(
        "👈 왼쪽 사이드바에서 후기 파일을 업로드하면 분석이 시작됩니다.\n\n"
        "예시: `네이버 블로그_2026-08-06.xlsx`, `구글맵 리뷰_2026-08-06.xlsx`"
    )
    with st.expander("💡 이 대시보드는 무엇을 하나요?"):
        st.markdown("""
        - 업로드한 후기를 **긍정 / 중립 / 부정**으로 자동 분류합니다.
        - **별점이 있는 채널**(구글맵·카카오맵)은 별점을 우선 사용하고,
          별점이 없으면 **본문 단어**로 판정합니다.
        - 채널별·감성별 **인터랙티브 차트**와 **필터 표**를 제공합니다.
        - 결과를 **엑셀/CSV로 다운로드**할 수 있습니다.
        """)
    st.stop()

# --- 업로드 파일 분류 ---
parsed = []          # (파일명, 채널, DataFrame)
for f in uploaded:
    try:
        raw = read_file(f.getvalue(), f.name)
        channel = guess_channel(f.name)
        parsed.append((f.name, channel, raw))
    except Exception as e:
        st.error(f"'{f.name}' 읽기 실패: {e}")

# --- 중복 방지: 통합 블로그가 있으면 개별 네이버/다음은 자동 제외 ---
has_merged = any(ch == "통합블로그" for _, ch, _ in parsed)
skipped = []
kept = []
for name, ch, raw in parsed:
    if has_merged and ch in ("네이버블로그", "다음블로그"):
        skipped.append((name, ch))          # 통합에 이미 포함 → 중복 제외
    else:
        kept.append((name, ch, raw))

if skipped:
    names = ", ".join(f"'{n}'" for n, _ in skipped)
    st.warning(
        f"⚠️ **통합 블로그** 파일이 있어, 개별 블로그 파일 {names} 은(는) "
        "중복 방지를 위해 자동 제외했어요. (통합 파일에 이미 포함되어 있어요)"
    )

# --- 표준화 + 병합 ---
frames, load_info = [], []
for name, ch, raw in kept:
    std = standardize(raw, ch)
    frames.append(std)
    # 통합 파일은 내부적으로 네이버/다음으로 나뉘므로 표시를 명확히
    shown = ch
    if ch == "통합블로그":
        shown = "통합블로그(→네이버/다음 분리)"
    load_info.append((name, shown, len(std)))

if not frames:
    st.warning("읽을 수 있는 데이터가 없습니다. 파일 형식을 확인해주세요.")
    st.stop()

data = analyze(pd.concat(frames, ignore_index=True))

with st.expander(f"📥 불러온 파일 {len(load_info)}개 보기", expanded=True):
    st.table(pd.DataFrame(load_info, columns=["파일명", "인식된 채널", "건수"]))
    st.caption("※ 통합 블로그 파일은 내부의 '출처'에 따라 네이버/다음으로 자동 분류됩니다.")

# --- KPI ---
total = len(data)
counts = data["감성"].value_counts().reindex(LABELS).fillna(0).astype(int)
pos_rate = counts["긍정"] / total * 100 if total else 0
rated = pd.to_numeric(data["별점"], errors="coerce").dropna()
avg_star = rated.mean() if len(rated) else None

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("전체 후기", f"{total:,}건")
c2.metric("😊 긍정", f"{counts['긍정']:,}", f"{pos_rate:.1f}%")
c3.metric("😐 중립", f"{counts['중립']:,}")
c4.metric("😞 부정", f"{counts['부정']:,}")
c5.metric("⭐ 평균 별점", f"{avg_star:.2f}" if avg_star is not None else "-")

st.divider()

# --- 차트 ---
col_left, col_right = st.columns([1, 1.3])
with col_left:
    st.subheader("전체 감성 분포")
    pie_df = pd.DataFrame({"감성": counts.index, "건수": counts.values})
    donut = (alt.Chart(pie_df).mark_arc(innerRadius=70).encode(
        theta=alt.Theta("건수:Q"),
        color=alt.Color("감성:N",
                        scale=alt.Scale(domain=LABELS, range=[COLOR[l] for l in LABELS]),
                        legend=alt.Legend(title="감성")),
        tooltip=["감성", "건수"]).properties(height=320))
    st.altair_chart(donut, use_container_width=True)

with col_right:
    st.subheader("채널별 감성 분포")
    ch_df = data.groupby(["채널", "감성"]).size().reset_index(name="건수")
    bar = (alt.Chart(ch_df).mark_bar().encode(
        y=alt.Y("채널:N", title=None, sort="-x"),
        x=alt.X("건수:Q", title="건수"),
        color=alt.Color("감성:N",
                        scale=alt.Scale(domain=LABELS, range=[COLOR[l] for l in LABELS]),
                        legend=alt.Legend(title="감성")),
        order=alt.Order("감성:N"),
        tooltip=["채널", "감성", "건수"]).properties(height=320))
    st.altair_chart(bar, use_container_width=True)

date_series = pd.to_datetime(data["작성일"], errors="coerce")
if date_series.notna().any():
    st.subheader("월별 후기 추이")
    tmp = data.copy()
    tmp["_월"] = date_series.dt.to_period("M").astype(str)
    tmp = tmp[tmp["_월"] != "NaT"]
    trend = tmp.groupby(["_월", "감성"]).size().reset_index(name="건수")
    line = (alt.Chart(trend).mark_bar().encode(
        x=alt.X("_월:N", title="월"),
        y=alt.Y("건수:Q", title="건수"),
        color=alt.Color("감성:N",
                        scale=alt.Scale(domain=LABELS, range=[COLOR[l] for l in LABELS])),
        tooltip=["_월", "감성", "건수"]).properties(height=300))
    st.altair_chart(line, use_container_width=True)

st.divider()

# --- 상세 표 ---
st.subheader("📋 후기 상세 보기")
f1, f2, f3 = st.columns([1, 1, 2])
sel_channels = f1.multiselect("채널", sorted(data["채널"].unique()),
                              default=sorted(data["채널"].unique()))
sel_labels = f2.multiselect("감성", LABELS, default=LABELS)
keyword = f3.text_input("본문 검색어", placeholder="예: 친절, 주차, 아이 …")

view = data[data["채널"].isin(sel_channels) & data["감성"].isin(sel_labels)]
if keyword.strip():
    view = view[view["_full"].str.contains(keyword.strip(), case=False, na=False)]

show_cols = ["채널", "작성일", "작성자", "별점", "감성", "판정근거", "감성점수",
             "제목", "내용", "URL"]
st.caption(f"필터 결과: {len(view):,}건")
st.dataframe(
    view[show_cols], use_container_width=True, hide_index=True, height=380,
    column_config={
        "URL": st.column_config.LinkColumn("URL", width="small"),
        "내용": st.column_config.TextColumn("내용", width="large"),
    },
)

# ============================== 분석 보고서 ==============================
st.divider()
st.subheader("📑 분석 보고서 (보고 자료용)")

# 텍스트 컬럼 준비
data["_full"] = (data["제목"].fillna("").astype(str) + " "
                 + data["내용"].fillna("").astype(str)).str.strip()

_total = len(data)
_counts = data["감성"].value_counts().reindex(LABELS).fillna(0).astype(int)
_pos_rate = round(_counts["긍정"] / _total * 100, 1) if _total else 0
_neg = data[data["감성"] == "부정"]

rep_tab1, rep_tab5, rep_tab2, rep_tab3, rep_tab4 = st.tabs(
    ["① 플랫폼별", "② ★주제 분석", "③ 감성별 키워드·문장",
     "④ 공통 언급", "⑤ ★개선 과제"])

# --- ① 플랫폼별 ---
with rep_tab1:
    piv = (data.groupby(["채널", "감성"]).size().unstack(fill_value=0)
           .reindex(columns=LABELS, fill_value=0))
    piv["합계"] = piv.sum(axis=1)
    piv["긍정률(%)"] = (piv["긍정"] / piv["합계"] * 100).round(1)
    star = pd.to_numeric(data["별점"], errors="coerce")
    avg_star = data.assign(_s=star).groupby("채널")["_s"].mean().round(2)
    piv["평균별점"] = avg_star
    piv = piv.sort_values("긍정률(%)", ascending=False)
    tone = ("매우 긍정적" if _pos_rate >= 70 else "대체로 긍정적"
            if _pos_rate >= 50 else "긍·부정 혼재" if _pos_rate >= 30 else "개선 필요")
    st.info(f"**한 줄 총평** — 전체 {_total}건 중 긍정 {_counts['긍정']}건"
            f"({_pos_rate}%)으로, 방문객 반응은 **'{tone}'**입니다.")
    st.dataframe(piv, use_container_width=True)

# --- ② ★주제 분석 (방문객이 주로 하는 이야기) ---
with rep_tab5:
    st.markdown("#### 🎯 방문객이 주로 이야기하는 내용")
    topics = analyze_topics(data)
    if not topics:
        st.warning("주제로 분류된 후기가 없습니다.")
    else:
        top = topics[0]
        st.success(f"**핵심 발견** — 후기의 **{top['비율']}%**가 "
                   f"**'{top['주제']}'**을(를) 언급하며 가장 큰 화제입니다.")
        # 주제별 비율 막대차트
        tdf = pd.DataFrame([{"주제": t["주제"], "건수": t["건수"],
                             "비율(%)": t["비율"]} for t in topics])
        chart = (alt.Chart(tdf).mark_bar(color="#5B8FF9").encode(
            x=alt.X("건수:Q", title="언급 후기 수"),
            y=alt.Y("주제:N", sort="-x", title=None),
            tooltip=["주제", "건수", "비율(%)"]).properties(height=340))
        st.altair_chart(chart, use_container_width=True)

        st.markdown("##### 주제별 상세")
        for t in topics:
            sd = senti_str(t["감성분포"])
            with st.expander(f"▶ {t['주제']}  —  {t['건수']}건 ({t['비율']}%)  [{sd}]",
                             expanded=(t is topics[0])):
                if t["핵심문구"]:
                    st.markdown("**자주 나온 표현:** "
                                + ", ".join(f"`{p}`" for p in t["핵심문구"]))
                if t["대표후기"]:
                    st.markdown(f"**대표 후기:** \"{t['대표후기']}\"")
        # 주제 분석 CSV 다운로드
        tdf_full = pd.DataFrame([{
            "주제": t["주제"], "건수": t["건수"], "비율(%)": t["비율"],
            "감성분포": senti_str(t["감성분포"]),
            "핵심문구": " / ".join(t["핵심문구"]), "대표후기": t["대표후기"],
        } for t in topics])
        st.download_button(
            "⬇️ 주제 분석 표(CSV) 다운로드",
            data=tdf_full.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"주제분석_{datetime.date.today()}.csv",
            mime="text/csv")

# --- ③ 감성별 키워드 & 대표 문장 ---
with rep_tab2:
    for lab in LABELS:
        sub = data[data["감성"] == lab]
        if len(sub) == 0:
            continue
        emoji = {"긍정": "😊", "중립": "😐", "부정": "😞"}[lab]
        st.markdown(f"#### {emoji} {lab} ({len(sub)}건)")
        kws = top_keywords(sub["_full"].tolist(), n=15)
        if kws:
            st.markdown("**핵심 키워드:** "
                        + ", ".join(f"`{w}`({c})" for w, c in kws))
        pairs = list(zip(sub.get("감성점수", [0] * len(sub)), sub["_full"].tolist()))
        sents = rep_sentences(pairs, n=5)
        if sents:
            st.markdown("**대표 문장:**")
            for s in sents:
                st.markdown(f"> {s}")
        st.divider()

# --- ③ 공통 언급 ---
with rep_tab3:
    common = top_keywords(data["_full"].tolist(), n=15)
    st.markdown("#### 🔑 후기가 공통적으로 자주 언급한 키워드")
    if common:
        kw_df = pd.DataFrame(common, columns=["키워드", "언급 문서 수"])
        cc1, cc2 = st.columns([1, 1])
        with cc1:
            st.dataframe(kw_df, use_container_width=True, hide_index=True, height=360)
        with cc2:
            chart = (alt.Chart(kw_df).mark_bar(color="#5B8FF9").encode(
                x=alt.X("언급 문서 수:Q"),
                y=alt.Y("키워드:N", sort="-x"),
                tooltip=["키워드", "언급 문서 수"]).properties(height=360))
            st.altair_chart(chart, use_container_width=True)

# --- ④ 개선 과제 ---
with rep_tab4:
    st.markdown("#### 🛠️ 개선 과제 (부정 후기 주제별 정리)")
    if len(_neg) == 0:
        st.success("부정 후기가 없어 도출된 개선 과제가 없습니다. 👍")
    else:
        improvements = analyze_improvements(_neg)
        if not improvements:
            st.warning("부정 후기가 있으나 사전 정의된 주제에 매칭되지 않았습니다. "
                       "아래 부정 대표 문장을 직접 검토해 주세요.")
            pairs = list(zip(_neg.get("감성점수", [0] * len(_neg)),
                             _neg["_full"].tolist()))
            for s in rep_sentences(pairs, n=5):
                st.markdown(f"> {s}")
        else:
            top_theme, top_cnt = improvements[0][0], improvements[0][1]
            st.error(f"**최우선 개선 과제: '{top_theme}'** "
                     f"(부정 {top_cnt}건, 가장 많이 지적됨)")
            imp_rows = []
            for i, (theme, cnt, samples, advice) in enumerate(improvements, 1):
                with st.expander(f"{i}. {theme}  —  부정 {cnt}건", expanded=(i == 1)):
                    for s in samples:
                        st.markdown(f"- 후기: \"{s}\"")
                    st.success(f"✅ 개선 방향: {advice}")
                imp_rows.append({"우선순위": i, "개선주제": theme, "부정건수": cnt,
                                 "대표후기": " / ".join(samples), "개선방향": advice})
            # 개선 과제 CSV 다운로드
            imp_df = pd.DataFrame(imp_rows)
            st.download_button(
                "⬇️ 개선 과제 표(CSV) 다운로드",
                data=imp_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"개선과제_{datetime.date.today()}.csv",
                mime="text/csv",
            )

# --- 다운로드 ---
st.divider()
st.subheader("💾 결과 다운로드")
today = datetime.date.today().strftime("%Y-%m-%d")
d1, d2 = st.columns(2)
d1.download_button(
    "⬇️ 엑셀(.xlsx) 다운로드",
    data=to_excel_bytes(view[show_cols]),
    file_name=f"감성분석_결과_{today}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
d2.download_button(
    "⬇️ CSV 다운로드",
    data=view[show_cols].to_csv(index=False).encode("utf-8-sig"),
    file_name=f"감성분석_결과_{today}.csv",
    mime="text/csv",
    use_container_width=True,
)
