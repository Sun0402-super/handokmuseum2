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

# ===================== 은은한 인디고 테마 (Style C) =====================
#  슬레이트(#1E293B~#334155) + 인디고(#4F46E5) 포인트
THEME = {
    "navy": "#1E293B", "blue": "#4F46E5", "sky": "#6366F1",
    "bg": "#F6F7FB", "card": "#FFFFFF", "text": "#1E293B",
    "pos": "#4F46E5", "neu": "#94A3B8", "neg": "#F43F5E",
    "slate2": "#334155", "chip": "#EEF0FE", "chiptext": "#4F46E5",
}
st.markdown(f"""
<style>
    /* 전체 배경 & 폰트 */
    .stApp {{ background: {THEME['bg']}; }}
    html, body, [class*="css"] {{
        font-family: 'Pretendard','Malgun Gothic','Segoe UI',sans-serif;
        color: {THEME['text']};
    }}
    /* Streamlit 기본 상단 헤더 숨기기 (우리 배너가 그 자리를 대신함) */
    header[data-testid="stHeader"] {{ display: none; }}

    /* 본문 여백: 상단은 고정 배너 높이만큼 확보 (배너를 키워서 값도 증가) */
    .block-container {{ padding-top: 132px; padding-bottom: 2rem;
                        max-width: 1300px; }}

    /* ---- 헤더 배너: 화면 최상단 '전체 폭'(사이드바까지) 고정 ---- */
    .hero {{
        position: fixed; top: 0; left: 0; right: 0; z-index: 2147483647;
        height: 108px; display: flex; align-items: center;
        background: linear-gradient(135deg, {THEME['navy']} 0%,
                    {THEME['slate2']} 100%);
        padding: 0 40px;
        box-shadow: 0 2px 14px rgba(0,0,0,0.20);
        color: #fff;
        opacity: 1 !important;               /* 리런 중 투명해짐 방지 */
    }}
    .hero-row {{ display:flex; align-items:center; gap:18px; }}
    .hero-logo {{
        width:58px; height:58px; border-radius:16px; flex:none;
        background: rgba(255,255,255,0.08);
        border: 1px solid rgba(255,255,255,0.12);
        display:flex; align-items:center; justify-content:center;
        font-size:30px;
    }}
    .hero h1 {{ margin: 0; font-size: 27px; font-weight: 800;
                letter-spacing: -0.3px; color:#F8FAFC; line-height:1.25; }}
    .hero p  {{ margin: 4px 0 0; font-size: 14.5px; color:#CBD5E1; }}

    /* ★ Streamlit 리런 시 요소가 잠깐 투명해지는 현상 방지
       (fixed 로 뺀 배너는 부모의 opacity 변화까지는 못 벗어나므로
        배너를 감싼 조상 요소들의 opacity/transition 을 강제 고정) */
    div:has(> .hero),
    div:has(.hero) {{
        opacity: 1 !important;
        transition: none !important;
        animation: none !important;
    }}
    [data-testid="stMarkdownContainer"]:has(.hero),
    [data-testid="element-container"]:has(.hero),
    .element-container:has(.hero) {{
        opacity: 1 !important;
        transition: none !important;
    }}
    .hero, .hero * {{ opacity: 1 !important; }}

    /* 사이드바 내용도 배너 높이만큼 아래로 */
    section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {{
        padding-top: 124px;
    }}

    /* ---- KPI 카드 ---- */
    .kpi {{
        background: {THEME['card']}; border-radius: 16px; padding: 18px 20px;
        box-shadow: 0 4px 16px rgba(30,111,217,0.10);
        border: 1px solid #E8F0FB; height: 100%;
        transition: transform .15s ease, box-shadow .15s ease;
    }}
    .kpi:hover {{ transform: translateY(-3px);
                  box-shadow: 0 10px 24px rgba(30,111,217,0.18); }}
    .kpi .label {{ font-size: 13px; color:#5B7089; font-weight:600; }}
    .kpi .value {{ font-size: 30px; font-weight: 800; margin-top: 4px;
                   line-height:1.1; }}
    .kpi .sub   {{ font-size: 12.5px; color:#8AA0B6; margin-top: 3px; }}
    .kpi .icon  {{ float:right; font-size: 26px; opacity:.85; }}

    /* ---- 섹션 제목 ---- */
    h2, h3 {{ color: {THEME['navy']}; font-weight: 700; }}
    [data-testid="stHeading"] {{ color: {THEME['navy']}; }}

    /* ---- 탭 디자인 ---- */
    .stTabs [data-baseweb="tab-list"] {{ gap: 6px; }}
    .stTabs [data-baseweb="tab"] {{
        background: #EAF2FD; border-radius: 10px 10px 0 0;
        padding: 9px 16px; font-weight: 600; color: {THEME['navy']};
    }}
    .stTabs [aria-selected="true"] {{
        background: {THEME['blue']}; color: #fff !important;
    }}

    /* 다운로드/버튼 */
    .stDownloadButton button, .stButton button {{
        background: {THEME['blue']}; color:#fff; border:none;
        border-radius: 10px; font-weight:600;
    }}
    .stDownloadButton button:hover, .stButton button:hover {{
        background: {THEME['navy']};
    }}
    /* 사이드바 */
    section[data-testid="stSidebar"] {{ background:#F0F1FA; }}

    /* ---- 업로드 카드 (메인 영역) ---- */
    .upload-card {{
        background: {THEME['card']}; border: 1px solid #E7EAF3;
        border-radius: 16px; padding: 8px 20px 4px;
        box-shadow: 0 4px 16px rgba(79,70,229,0.06); margin-bottom: 18px;
    }}
    .upload-card .uc-title {{ font-size: 15px; font-weight: 700;
        color: {THEME['navy']}; margin: 10px 0 2px; }}
    .upload-card .uc-desc {{ font-size: 12.5px; color:#7A8698; margin-bottom: 8px; }}

    /* ---- 사이드바 메뉴(라디오 → 메뉴 버튼처럼) ---- */
    section[data-testid="stSidebar"] .stRadio > label {{ display:none; }}
    section[data-testid="stSidebar"] .stRadio [role="radiogroup"] {{
        gap: 6px; }}
    section[data-testid="stSidebar"] .stRadio [role="radiogroup"] label {{
        background:#FFFFFF; border:1px solid #E4E7F2; border-radius:12px;
        padding:11px 14px; font-weight:600; font-size:14.5px;
        color:{THEME['navy']}; cursor:pointer; transition:all .12s ease;
        margin:0;
    }}
    section[data-testid="stSidebar"] .stRadio [role="radiogroup"] label:hover {{
        border-color:{THEME['blue']}; box-shadow:0 2px 8px rgba(79,70,229,.10); }}
    /* 선택된 메뉴 */
    section[data-testid="stSidebar"] .stRadio [role="radiogroup"]
        label:has(input:checked) {{
        background:{THEME['blue']}; color:#fff; border-color:{THEME['blue']};
        box-shadow:0 4px 12px rgba(79,70,229,.28); }}
    /* 라디오 동그라미 숨기기 */
    section[data-testid="stSidebar"] .stRadio [role="radiogroup"]
        label > div:first-child {{ display:none; }}
</style>
""", unsafe_allow_html=True)

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
COLOR = {"긍정": "#4F46E5", "중립": "#94A3B8", "부정": "#F43F5E"}

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


# ----------------------- 실제 방문 장소 판정 -----------------------
# 장소 필터는 수집 채널(네이버/구글/카카오)이 아니라 실제 방문 지점 기준입니다.
# UI에는 항상 "전체 / 음성 / 서울"만 표시합니다.
SEOUL_URL_HINTS = [
    "place.map.kakao.com/1584706006",  # 카카오맵: 한독의약박물관 서울 후기
]
EUMSEONG_URL_HINTS = [
    # 음성 카카오맵 장소 ID를 확인하면 여기에 추가하세요.
]

SEOUL_STRONG = [
    "한독의약박물관 서울", "한독 박물관 서울", "한독뮤지엄 서울", "서울 한독", "서울관",
    "한독퓨처콤플렉스", "퓨처콤플렉스", "future complex", "강남", "강남구",
    "논현", "논현동", "서울특별시", "서울시", "서울", "신논현", "언주역", "역삼"
]
SEOUL_CONTEXT = [
    "도심", "빌딩", "사옥", "복합문화공간", "백신", "제약사옥", "라운지", "전시관 서울"
]
EUMSEONG_STRONG = [
    "한독의약박물관 음성", "한독 박물관 음성", "음성 한독", "음성관", "충북", "충청북도",
    "음성군", "대소면", "대풍산단", "대풍산업단지", "한독약품 음성", "한독 음성", "음성"
]
EUMSEONG_CONTEXT = [
    "생명갤러리", "약초원", "의약유물", "야외", "넓은", "공장", "견학", "어린이 체험",
    "가족 체험", "충청", "충북 음성"
]


def _loc_score(text, strong, context):
    t = str(text).lower()
    score = 0
    for k in strong:
        if k.lower() in t:
            score += 3
    for k in context:
        if k.lower() in t:
            score += 1
    return score


def _normalize_known_place(raw_place):
    p = str(raw_place or "")
    if not p.strip():
        return None
    if any(k in p for k in ["서울", "강남", "논현", "퓨처콤플렉스"]):
        return "서울"
    if any(k in p for k in ["음성", "충북", "충청북도", "대소", "생명갤러리", "약초원"]):
        return "음성"
    return None


def classify_visit_place(title="", text="", url="", raw_place="", channel="", filename=""):
    """
    후기 업로드 사이트가 아니라 실제 방문 지점을 음성/서울로 판정합니다.
    반환값: (장소, 판정근거, 신뢰도)
    """
    full = " ".join(map(str, [title, text, url, raw_place, channel, filename]))
    low = full.lower()

    # 1) URL 확정 힌트
    for hint in SEOUL_URL_HINTS:
        if hint.lower() in low:
            return "서울", "서울 URL 힌트", "높음"
    for hint in EUMSEONG_URL_HINTS:
        if hint.lower() in low:
            return "음성", "음성 URL 힌트", "높음"

    # 2) 원본 장소 컬럼 우선
    known = _normalize_known_place(raw_place)
    if known:
        return known, "원본 장소 컬럼", "높음"

    # 3) 키워드 점수 판정
    seoul_score = _loc_score(full, SEOUL_STRONG, SEOUL_CONTEXT)
    eum_score = _loc_score(full, EUMSEONG_STRONG, EUMSEONG_CONTEXT)

    if seoul_score >= eum_score + 2 and seoul_score > 0:
        return "서울", f"서울 키워드 우세({seoul_score}:{eum_score})", "높음" if seoul_score >= 5 else "중간"
    if eum_score >= seoul_score + 2 and eum_score > 0:
        return "음성", f"음성 키워드 우세({eum_score}:{seoul_score})", "높음" if eum_score >= 5 else "중간"

    # 4) 직접 언급 보정
    if any(k in full for k in ["서울", "강남", "논현", "퓨처콤플렉스"]):
        return "서울", "서울 직접 언급", "중간"
    if any(k in full for k in ["음성", "충북", "충청북도", "대소"]):
        return "음성", "음성 직접 언급", "중간"

    # 5) 일반 명칭만 있는 경우: 본관인 음성으로 보수 분류하되 신뢰도 낮음 표시
    if "한독의약박물관" in full or "한독 박물관" in full:
        return "음성", "일반 명칭만 있어 본관 기준 분류", "낮음"

    # UI 필터를 음성/서울만 유지하기 위한 기본값
    return "음성", "장소 단서 부족, 기본값 음성", "낮음"


def standardize(df, channel, filename=""):
    is_merged = (channel == "통합블로그")
    has_src = "출처" in df.columns          # 통합 파일에 있는 출처 컬럼
    recs = []
    for _, r in df.iterrows():
        text = _get(r, "리뷰내용", "본문", "요약", "내용", "리뷰", "후기")
        title = _get(r, "제목", "타이틀")
        url = _get(r, "URL", "url", "링크", "주소")

        # 통합 파일이면 행의 '출처'로 실제 채널을 지정(네이버/다음 분리)
        ch = channel
        if is_merged and has_src:
            mapped = _source_to_channel(r.get("출처"))
            if mapped:
                ch = mapped

        # 원본 장소는 보존하되, 분석용 장소는 실제 방문 지점(음성/서울)으로 재분류
        raw_place = _get(r, "장소", "방문장소", "지점", "place", "location")
        visit_place, place_reason, place_conf = classify_visit_place(
            title=title, text=text, url=url, raw_place=raw_place, channel=ch, filename=filename
        )

        recs.append({
            "채널": ch,                    # 어디에서 수집했는가
            "장소": visit_place,           # 실제 어디를 방문했는가: 음성/서울만 사용
            "원본장소": raw_place,
            "장소판정근거": place_reason,
            "장소판정신뢰도": place_conf,
            "작성일": _get(r, "작성일"),
            "작성자": _get(r, "작성자", "블로그명", "블로그ID"),
            "별점": r["별점"] if "별점" in r and pd.notna(r["별점"]) else None,
            "제목": title,
            "내용": text,
            "_full": (title + " " + text).strip(),
            "URL": url,
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
# ---- 헤더 배너 (칩은 사이드바 메뉴로 이동) ----
st.markdown("""
<div class="hero">
    <div class="hero-row">
        <div class="hero-logo">🏛️</div>
        <div>
            <h1>한독의약박물관 후기 분석</h1>
            <p>네이버·다음·구글맵·카카오맵 리뷰를 실제 방문 장소(음성·서울) 기준으로 감성·주제·개선과제로 자동 분석합니다.</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ---- 업로드 (배너 바로 아래, 메인 영역) ----
st.markdown("""
<div class="upload-card">
    <div class="uc-title">📂 데이터 업로드</div>
    <div class="uc-desc">PC 크롤러가 만든 후기 <b>엑셀(.xlsx)</b> 또는 <b>.csv</b> 를
        여러 개 한 번에 올려주세요. 파일명으로 수집 채널이 분류되고, 후기 내용으로 실제 방문 장소가 자동 판정됩니다.</div>
</div>
""", unsafe_allow_html=True)
uploaded = st.file_uploader(
    "후기 파일 선택",
    type=["xlsx", "xls", "csv"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not uploaded:
    st.info(
        "⬆️ 위에서 후기 파일을 업로드하면 분석이 시작됩니다.\n\n"
        "예시: `네이버 블로그_2026-08-06.xlsx`, `구글맵 리뷰_2026-08-06.xlsx`"
    )
    with st.expander("💡 이 대시보드는 무엇을 하나요?"):
        st.markdown("""
        - 업로드한 후기를 **긍정 / 중립 / 부정**으로 자동 분류합니다.
        - **별점이 있는 채널**(구글맵·카카오맵)은 별점을 우선 사용하고,
          별점이 없으면 **본문 단어**로 판정합니다.
        - **주제 분석**으로 "후기가 주로 무슨 이야기를 하는지" 비율까지 보여줍니다.
        - 장소별·감성별 **인터랙티브 차트**와 **개선 과제**를 제공합니다.
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
    std = standardize(raw, ch, name)
    frames.append(std)
    # 통합 파일은 내부적으로 네이버/다음으로 나뉘므로 표시를 명확히
    shown = ch
    if ch == "통합블로그":
        shown = "통합블로그(→네이버/다음 분리)"
    load_info.append((name, shown, len(std)))

if not frames:
    st.warning("읽을 수 있는 데이터가 없습니다. 파일 형식을 확인해주세요.")
    st.stop()

data_all = analyze(pd.concat(frames, ignore_index=True))

# 텍스트 컬럼 준비(주제/키워드 분석용)
data_all["_full"] = (data_all["제목"].fillna("").astype(str) + " "
                     + data_all["내용"].fillna("").astype(str)).str.strip()

# 장소 필터는 실제 방문 장소 기준으로 고정: 전체 / 음성 / 서울
place_options = ["전체", "음성", "서울"]

show_cols = ["장소", "채널", "작성일", "작성자", "별점", "감성", "판정근거", "감성점수",
             "장소판정신뢰도", "장소판정근거", "제목", "내용", "URL"]
today = datetime.date.today().strftime("%Y-%m-%d")


def kpi_card(label, value, sub="", icon="", color=THEME["navy"]):
    return f"""
    <div class="kpi">
        <span class="icon">{icon}</span>
        <div class="label">{label}</div>
        <div class="value" style="color:{color}">{value}</div>
        <div class="sub">{sub}</div>
    </div>"""


# ============================== 사이드바 메뉴 ==============================
with st.sidebar:
    menu = st.radio(
        "menu",
        ["📊 감성 분석", "🎯 주제 분석", "🛠️ 개선 과제", "📥 보고서 다운로드"],
        label_visibility="collapsed",
    )
    st.divider()

    # ---- 실제 방문 장소 필터: 전체 / 음성 / 서울만 표시 ----
    st.markdown("**📍 장소 필터**")
    if st.session_state.get("place_filter", "전체") not in place_options:
        st.session_state["place_filter"] = "전체"
    fcol1, fcol2 = st.columns([2.4, 1])
    with fcol2:
        if st.button("↺ 초기화", use_container_width=True,
                     help="장소 필터를 '전체'로 되돌립니다.",
                     disabled=st.session_state.get("place_filter", "전체") == "전체"):
            st.session_state["place_filter"] = "전체"
            st.rerun()
    selected_place = st.radio(
        "실제 방문 장소",
        place_options,
        label_visibility="collapsed",
        key="place_filter",
        help="업로드 사이트가 아니라 실제 방문 장소 기준입니다. 출처 채널은 보조 정보로만 사용됩니다.",
    )
    if selected_place != "전체":
        st.caption(f"현재 '{selected_place}' 방문 후기만 보고 있어요.")
    st.divider()

    with st.expander(f"📥 불러온 파일 {len(load_info)}개", expanded=False):
        st.table(pd.DataFrame(load_info, columns=["파일명", "채널", "건수"]))
    if skipped:
        st.caption("⚠️ 통합 파일이 있어 개별 블로그는 중복 제외됨")
    st.divider()
    st.caption("분석 전용 대시보드 · 후기 수집은 PC 크롤러에서 진행")

# ---- 필터 적용: 선택된 장소로 data 좁히기 (전체 메뉴가 이 data 를 사용) ----
data = (data_all if selected_place == "전체"
        else data_all[data_all["장소"] == selected_place]).reset_index(drop=True)

# ---- 공통 집계 (필터 반영된 data 기준) ----
total = len(data)
counts = data["감성"].value_counts().reindex(LABELS).fillna(0).astype(int)
pos_rate = counts["긍정"] / total * 100 if total else 0
rated = pd.to_numeric(data["별점"], errors="coerce").dropna()
avg_star = rated.mean() if len(rated) else None
_neg = data[data["감성"] == "부정"]

if selected_place != "전체" and total == 0:
    st.warning(f"'{selected_place}'에 해당하는 후기가 없습니다. "
               "사이드바에서 다른 장소를 선택해 주세요.")
    st.stop()


# ============================== ① 감성 분석 ==============================
# 장소 필터가 적용된 상태면, 어느 메뉴를 보든 상단에 안내 배지 표시
if selected_place != "전체":
    st.markdown(f"""
    <div style="display:inline-block;background:{THEME['chip']};
                color:{THEME['blue']};padding:6px 14px;border-radius:20px;
                font-size:12.5px;font-weight:700;margin-bottom:12px;">
        📍 장소 필터 적용 중: {selected_place} ({total}건)
    </div>
    """, unsafe_allow_html=True)

if menu == "📊 감성 분석":
    _avg = f"{avg_star:.2f}" if avg_star is not None else "-"
    kpis = [
        kpi_card("전체 후기", f"{total:,}", "건", "📝", THEME["navy"]),
        kpi_card("긍정", f"{counts['긍정']:,}", f"{pos_rate:.1f}%", "😊", THEME["pos"]),
        kpi_card("중립", f"{counts['중립']:,}", "건", "😐", THEME["neu"]),
        kpi_card("부정", f"{counts['부정']:,}",
                 f"{counts['부정']/total*100:.1f}%" if total else "", "😞", THEME["neg"]),
        kpi_card("평균 별점", _avg, "5점 만점", "⭐", THEME["blue"]),
    ]
    for col, html in zip(st.columns(5), kpis):
        col.markdown(html, unsafe_allow_html=True)
    st.write("")

    # 플랫폼별 요약 + 한 줄 총평
    piv = (data.groupby(["채널", "감성"]).size().unstack(fill_value=0)
           .reindex(columns=LABELS, fill_value=0))
    piv["합계"] = piv.sum(axis=1)
    piv["긍정률(%)"] = (piv["긍정"] / piv["합계"] * 100).round(1)
    _star = pd.to_numeric(data["별점"], errors="coerce")
    piv["평균별점"] = data.assign(_s=_star).groupby("채널")["_s"].mean().round(2)
    piv = piv.sort_values("긍정률(%)", ascending=False)
    tone = ("매우 긍정적" if pos_rate >= 70 else "대체로 긍정적"
            if pos_rate >= 50 else "긍·부정 혼재" if pos_rate >= 30 else "개선 필요")
    st.info(f"**한 줄 총평** — 전체 {total}건 중 긍정 {counts['긍정']}건"
            f"({pos_rate:.1f}%)으로, 방문객 반응은 **'{tone}'**입니다.")

    st.divider()
    col_left, col_right = st.columns([1, 1.3])
    with col_left:
        st.subheader("🎯 전체 감성 분포")
        pie_df = pd.DataFrame({"감성": counts.index, "건수": counts.values})
        donut = (alt.Chart(pie_df).mark_arc(innerRadius=72, stroke="#fff",
                                            strokeWidth=2).encode(
            theta=alt.Theta("건수:Q"),
            color=alt.Color("감성:N",
                            scale=alt.Scale(domain=LABELS, range=[COLOR[l] for l in LABELS]),
                            legend=alt.Legend(title="감성", orient="bottom")),
            tooltip=["감성", "건수"]).properties(height=320))
        st.altair_chart(donut, use_container_width=True)
    with col_right:
        st.subheader("📊 채널별 감성 분포")
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
        st.subheader("📈 월별 후기 추이")
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
    st.subheader("📋 플랫폼별 상세 & 후기 검색")
    st.dataframe(piv, use_container_width=True)

    f1, f2, f3 = st.columns([1, 1, 2])
    sel_channels = f1.multiselect("채널", sorted(data["채널"].unique()),
                                  default=sorted(data["채널"].unique()))
    sel_labels = f2.multiselect("감성", LABELS, default=LABELS)
    keyword = f3.text_input("본문 검색어", placeholder="예: 친절, 주차, 아이 …")
    view = data[data["채널"].isin(sel_channels) & data["감성"].isin(sel_labels)]
    if keyword.strip():
        view = view[view["_full"].str.contains(keyword.strip(), case=False, na=False)]
    st.caption(f"필터 결과: {len(view):,}건")
    st.dataframe(
        view[show_cols], use_container_width=True, hide_index=True, height=380,
        column_config={
            "URL": st.column_config.LinkColumn("URL", width="small"),
            "내용": st.column_config.TextColumn("내용", width="large"),
        },
    )

    # 감성별 키워드 & 대표 문장
    st.divider()
    st.subheader("🔑 감성별 핵심 키워드 & 대표 문장")
    for lab in LABELS:
        sub = data[data["감성"] == lab]
        if len(sub) == 0:
            continue
        emoji = {"긍정": "😊", "중립": "😐", "부정": "😞"}[lab]
        with st.expander(f"{emoji} {lab} ({len(sub)}건)", expanded=(lab == "긍정")):
            kws = top_keywords(sub["_full"].tolist(), n=15)
            if kws:
                st.markdown("**핵심 키워드:** "
                            + ", ".join(f"`{w}`({c})" for w, c in kws))
            pairs = list(zip(sub.get("감성점수", [0] * len(sub)), sub["_full"].tolist()))
            for s in rep_sentences(pairs, n=5):
                st.markdown(f"> {s}")


# ============================== ② 주제 분석 ==============================
elif menu == "🎯 주제 분석":
    st.subheader("🎯 방문객이 주로 이야기하는 내용")
    topics = analyze_topics(data)
    if not topics:
        st.warning("주제로 분류된 후기가 없습니다.")
    else:
        top = topics[0]
        st.success(f"**핵심 발견** — 후기의 **{top['비율']}%**가 "
                   f"**'{top['주제']}'**을(를) 언급하며 가장 큰 화제입니다.")
        tdf = pd.DataFrame([{"주제": t["주제"], "건수": t["건수"],
                             "비율(%)": t["비율"]} for t in topics])
        chart = (alt.Chart(tdf).mark_bar(color="#4F46E5", cornerRadiusEnd=4).encode(
            x=alt.X("건수:Q", title="언급 후기 수"),
            y=alt.Y("주제:N", sort="-x", title=None),
            tooltip=["주제", "건수", "비율(%)"]).properties(height=360))
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

        # 공통 언급 키워드
        st.divider()
        st.subheader("🔑 전체 공통 언급 키워드")
        common = top_keywords(data["_full"].tolist(), n=15)
        if common:
            kw_df = pd.DataFrame(common, columns=["키워드", "언급 문서 수"])
            cc1, cc2 = st.columns([1, 1])
            with cc1:
                st.dataframe(kw_df, use_container_width=True, hide_index=True, height=360)
            with cc2:
                kc = (alt.Chart(kw_df).mark_bar(color="#4F46E5",
                                                cornerRadiusEnd=4).encode(
                    x=alt.X("언급 문서 수:Q"),
                    y=alt.Y("키워드:N", sort="-x"),
                    tooltip=["키워드", "언급 문서 수"]).properties(height=360))
                st.altair_chart(kc, use_container_width=True)


# ============================== ③ 개선 과제 ==============================
elif menu == "🛠️ 개선 과제":
    st.subheader("🛠️ 개선 과제 (부정 후기 장소별 · 주제별 정리)")

    if len(_neg) == 0:
        st.success("부정 후기가 없어 도출된 개선 과제가 없습니다. 👍")
    else:
        # '장소' 컬럼에 값이 2개 이상이면(예: 음성/서울) 장소별로 탭을 나눠 표시
        has_place = "장소" in _neg.columns and _neg["장소"].nunique() > 1

        def render_improvements(sub_neg, key_prefix=""):
            if len(sub_neg) == 0:
                st.caption("이 장소는 부정 후기가 없습니다. 👍")
                return
            imps = analyze_improvements(sub_neg)
            if not imps:
                st.warning("사전 정의된 주제에 매칭되지 않았습니다. "
                           "아래 부정 대표 문장을 직접 검토해 주세요.")
                pairs = list(zip(sub_neg.get("감성점수", [0] * len(sub_neg)),
                                 sub_neg["_full"].tolist()))
                for s in rep_sentences(pairs, n=5):
                    st.markdown(f"> {s}")
                return
            top_theme, top_cnt = imps[0][0], imps[0][1]
            st.error(f"**최우선 개선 과제: '{top_theme}'** "
                     f"(부정 {top_cnt}건, 가장 많이 지적됨)")
            for i, (theme, cnt, samples, advice) in enumerate(imps, 1):
                label = (f"[{key_prefix}] {i}. {theme}  —  부정 {cnt}건" if key_prefix
                        else f"{i}. {theme}  —  부정 {cnt}건")
                with st.expander(label, expanded=(i == 1)):
                    for s in samples:
                        st.markdown(f"- 후기: \"{s}\"")
                    st.success(f"✅ 개선 방향: {advice}")

        if has_place:
            # 사이드바에서 특정 장소를 아직 안 골랐고(전체 보기),
            # 장소가 여러 개 섞여 있을 때만 탭으로 나눠 보여줌
            places = sorted(_neg["장소"].unique())
            place_tabs = st.tabs([f"📍 {p}" for p in places])
            for tab, place in zip(place_tabs, places):
                with tab:
                    sub = _neg[_neg["장소"] == place]
                    st.caption(f"부정 후기 {len(sub)}건")
                    render_improvements(sub, key_prefix=place)
        else:
            # ★ 탭 자동 선택: 사이드바 장소 필터로 이미 한 장소만 남았거나
            #   원래 장소 구분이 없는 데이터라 탭 없이 바로 내용을 보여줌
            if selected_place != "전체":
                st.info(f"📍 사이드바 필터 적용 중: **{selected_place}**")
            render_improvements(_neg)


# ============================== ④ 보고서 다운로드 ==============================
elif menu == "📥 보고서 다운로드":
    st.subheader("📥 보고서 다운로드")
    st.caption("분석 결과를 파일로 내려받아 보고 자료에 바로 활용하세요. "
               "(모두 UTF-8, 엑셀에서 한글 정상)")

    # 1) 전체 감성분석 결과
    st.markdown("##### 📊 전체 감성분석 결과")
    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ 엑셀(.xlsx)",
        data=to_excel_bytes(data[show_cols]),
        file_name=f"감성분석_결과_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)
    d2.download_button(
        "⬇️ CSV",
        data=data[show_cols].to_csv(index=False).encode("utf-8-sig"),
        file_name=f"감성분석_결과_{today}.csv",
        mime="text/csv", use_container_width=True)

    # 2) 주제 분석
    st.markdown("##### 🎯 주제 분석 표")
    _topics = analyze_topics(data)
    if _topics:
        tdf_full = pd.DataFrame([{
            "주제": t["주제"], "건수": t["건수"], "비율(%)": t["비율"],
            "감성분포": senti_str(t["감성분포"]),
            "핵심문구": " / ".join(t["핵심문구"]), "대표후기": t["대표후기"],
        } for t in _topics])
        st.download_button(
            "⬇️ 주제 분석 CSV",
            data=tdf_full.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"주제분석_{today}.csv", mime="text/csv")
    else:
        st.caption("주제로 분류된 후기가 없습니다.")

    # 3) 개선 과제 (장소별로 나뉘어 있으면 '장소' 컬럼 포함)
    st.markdown("##### 🛠️ 개선 과제 표")
    if len(_neg) > 0:
        has_place_dl = "장소" in _neg.columns and _neg["장소"].nunique() > 1
        groups = (list(_neg.groupby("장소")) if has_place_dl else [("전체", _neg)])

        imp_rows = []
        for place_name, place_df in groups:
            _imp = analyze_improvements(place_df)
            for i, (th, cnt, sm, adv) in enumerate(_imp, 1):
                imp_rows.append({
                    "장소": place_name, "우선순위": i, "개선주제": th,
                    "부정건수": cnt, "대표후기": " / ".join(sm), "개선방향": adv,
                })

        if imp_rows:
            imp_df = pd.DataFrame(imp_rows)
            st.download_button(
                "⬇️ 개선 과제 CSV",
                data=imp_df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"개선과제_{today}.csv", mime="text/csv")
        else:
            st.caption("사전 정의 주제에 매칭된 개선 과제가 없습니다.")
    else:
        st.caption("부정 후기가 없어 개선 과제가 없습니다. 👍")
