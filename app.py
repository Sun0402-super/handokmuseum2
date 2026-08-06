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


def standardize(df, channel):
    recs = []
    for _, r in df.iterrows():
        text = _get(r, "리뷰내용", "본문", "요약")
        title = _get(r, "제목")
        recs.append({
            "채널": channel,
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

# --- 업로드 읽기 ---
frames, load_info = [], []
for f in uploaded:
    try:
        raw = read_file(f.getvalue(), f.name)
        channel = guess_channel(f.name)
        std = standardize(raw, channel)
        frames.append(std)
        load_info.append((f.name, channel, len(std)))
    except Exception as e:
        st.error(f"'{f.name}' 읽기 실패: {e}")

if not frames:
    st.warning("읽을 수 있는 데이터가 없습니다. 파일 형식을 확인해주세요.")
    st.stop()

data = analyze(pd.concat(frames, ignore_index=True))

with st.expander(f"📥 불러온 파일 {len(load_info)}개 보기", expanded=False):
    st.table(pd.DataFrame(load_info, columns=["파일명", "인식된 채널", "건수"]))

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
