# -*- coding: utf-8 -*-
"""한독의약박물관 온라인 후기 감성 분석 Streamlit 대시보드."""
from __future__ import annotations

import io
import re
from datetime import date
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

st.set_page_config(page_title="한독의약박물관 후기 감성 분석", page_icon="💊", layout="wide")

POSITIVE = [
    "좋다", "좋았", "좋은", "유익", "재미있", "재밌", "친절", "알차", "추천", "만족",
    "깨끗", "아이가 좋아", "설명이 좋", "체험이 좋", "의미 있", "의미있", "다시 방문",
    "최고", "훌륭", "흥미", "신기", "즐겁", "볼거리", "교육적", "가볼만"
]
NEGATIVE = [
    "아쉽", "불편", "어렵", "별로", "부족", "실망", "복잡", "주차 불편", "안내 부족",
    "찾기 어렵", "기대 이하", "시간이 짧", "불친절", "지루", "재미없", "협소", "좁",
    "기다", "웨이팅", "볼 게 없", "볼게 없"
]
TOPICS = {
    "교육·체험": ["체험", "교육", "약사", "연구원", "간호사", "응급구조", "만들기", "프로그램"],
    "전시·유물": ["전시", "유물", "의약", "약장", "의료기구", "역사", "보물"],
    "해설·안내": ["해설", "도슨트", "설명", "안내", "가이드"],
    "아이·가족": ["아이", "어린이", "초등", "가족", "부모", "방학"],
    "시설·환경": ["깨끗", "시설", "로비", "정원", "카페", "주차", "건물"],
    "예약·운영": ["예약", "운영시간", "휴관", "무료", "입장", "신청"],
}
OUTPUT_COLUMNS = [
    "source", "museum_location", "platform", "title", "content_summary", "full_text", "url",
    "author", "published_date", "collected_date", "rating", "sentiment", "sentiment_score",
    "positive_score", "negative_score", "keyword", "is_duplicate", "duplicate_group",
    "priority_source", "primary_topic", "notes"
]
ALIASES = {
    "source": ["source", "장소"], "platform": ["platform", "채널", "유형", "출처"],
    "title": ["title", "제목"], "content_summary": ["content_summary", "요약"],
    "full_text": ["full_text", "리뷰내용", "본문", "내용"], "url": ["url", "URL"],
    "author": ["author", "작성자", "블로그명", "블로그ID"],
    "published_date": ["published_date", "작성일"], "rating": ["rating", "별점"],
}

def value(row, names):
    for name in names:
        if name in row.index and pd.notna(row[name]) and str(row[name]).strip():
            return str(row[name]).strip()
    return ""

def read_file(uploaded):
    if uploaded.name.lower().endswith(".xlsx"):
        return pd.read_excel(uploaded, engine="openpyxl")
    raw = uploaded.getvalue()
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f"CSV 인코딩을 읽을 수 없습니다: {last_error}")

def detect_platform(filename, url, supplied=""):
    text = f"{filename} {url} {supplied}".lower()
    if "blog.naver.com" in text or "네이버" in text:
        return "네이버 블로그"
    if "tistory" in text or "daum" in text or "다음" in text:
        return "다음 블로그"
    if "google" in text or "구글" in text:
        return "구글맵"
    if "kakao" in text or "카카오" in text:
        return "카카오맵"
    if "instagram" in text or "인스타" in text:
        return "인스타그램"
    return supplied or "기타"

def detect_location(title, text, url, source):
    combined = re.sub(r"\s+", "", f"{title} {text} {url} {source}").lower()
    seoul_terms = ["서울관", "한독의약박물관서울", "마곡", "마곡나루", "퓨처콤플렉스", "강서구"]
    eumseong_terms = ["음성", "대소면", "대소읍", "대풍산단", "팩토리투어", "한독공장"]
    if any(term in combined for term in seoul_terms):
        return "서울관"
    if any(term in combined for term in eumseong_terms):
        return "음성관"
    return "미분류"

def sentiment(text, rating):
    compact = re.sub(r"\s+", "", str(text))
    pos = sum(compact.count(w.replace(" ", "")) for w in POSITIVE)
    neg = sum(compact.count(w.replace(" ", "")) for w in NEGATIVE)
    score = pos - neg
    rating_note = ""
    try:
        star = float(rating)
        if star >= 4:
            score += 2
        elif star <= 2:
            score -= 2
        rating_note = f"별점 {star:g}점 반영"
    except (TypeError, ValueError):
        pass
    label = "긍정" if score > 0 else ("부정" if score < 0 else "중립")
    return label, score, pos, neg, rating_note or "텍스트 키워드 분석"

def topic(text):
    compact = re.sub(r"\s+", "", str(text))
    scored = {name: sum(compact.count(w.replace(" ", "")) for w in words) for name, words in TOPICS.items()}
    best = max(scored, key=scored.get)
    return best if scored[best] > 0 else "기타"

def normalize(text):
    return re.sub(r"[^0-9a-z가-힣]", "", str(text).lower())

def standardize(df, filename):
    records = []
    for _, row in df.iterrows():
        rec = {key: value(row, names) for key, names in ALIASES.items()}
        rec["platform"] = detect_platform(filename, rec["url"], rec["platform"])
        rec["source"] = rec["source"] or rec["platform"]
        rec["full_text"] = rec["full_text"] or rec["content_summary"]
        combined = f"{rec['title']} {rec['content_summary']} {rec['full_text']}"
        rec["museum_location"] = detect_location(rec["title"], combined, rec["url"], rec["source"])
        label, score, pos, neg, note = sentiment(combined, rec["rating"])
        rec.update({
            "collected_date": date.today().isoformat(), "sentiment": label, "sentiment_score": score,
            "positive_score": pos, "negative_score": neg, "keyword": "한독의약박물관",
            "is_duplicate": False, "duplicate_group": "", "priority_source": rec["platform"],
            "primary_topic": topic(combined), "notes": note,
        })
        records.append(rec)
    return pd.DataFrame(records)

def deduplicate(df, threshold):
    priority = {"네이버 블로그": 0, "다음 블로그": 1, "구글맵": 2, "카카오맵": 3, "인스타그램": 4}
    work = df.copy()
    work["_priority"] = work["platform"].map(priority).fillna(9)
    work = work.sort_values("_priority", kind="stable").reset_index(drop=True)
    representatives = []
    groups = []
    for idx, row in work.iterrows():
        url_key = normalize(row["url"])
        text_key = normalize(row["title"] or row["full_text"][:140])
        match = None
        for rep_idx, rep_url, rep_text, gid in representatives:
            same_url = bool(url_key and rep_url and url_key == rep_url)
            similar = bool(text_key and rep_text and SequenceMatcher(None, text_key, rep_text).ratio() >= threshold)
            if same_url or similar:
                match = gid
                break
        if match:
            work.at[idx, "is_duplicate"] = True
            work.at[idx, "duplicate_group"] = match
            work.at[idx, "priority_source"] = "중복 제외"
        else:
            gid = f"D{len(representatives)+1:04d}"
            work.at[idx, "duplicate_group"] = gid
            representatives.append((idx, url_key, text_key, gid))
        groups.append(work.at[idx, "duplicate_group"])
    return work.drop(columns="_priority")

def excel_bytes(data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data[OUTPUT_COLUMNS].to_excel(writer, index=False, sheet_name="후기분석")
        pd.crosstab(data["museum_location"], data["sentiment"], margins=True).to_excel(writer, sheet_name="관별감성요약")
        pd.crosstab(data["platform"], data["sentiment"], margins=True).to_excel(writer, sheet_name="플랫폼요약")
        data.groupby(["museum_location", "primary_topic"]).size().rename("건수").reset_index().to_excel(writer, index=False, sheet_name="주제분석")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column in ws.columns:
                width = min(max(len(str(cell.value or "")) for cell in column) + 2, 55)
                ws.column_dimensions[column[0].column_letter].width = width
    return output.getvalue()

st.title("💊 한독의약박물관 온라인 후기 감성 분석기")
st.caption("음성관과 서울관 후기를 구분하고, 플랫폼별·관별 감성 및 주제를 그래프로 확인합니다.")

with st.sidebar:
    st.header("분석 설정")
    threshold = st.slider("중복 판정 유사도", 0.70, 1.00, 0.90, 0.01)
    remove_duplicates = st.checkbox("중복 후기를 분석에서 제외", value=True)
    st.markdown("**관 구분 기준**")
    st.caption("서울관: 서울관·마곡·퓨처콤플렉스·강서구\n\n음성관: 음성·대소·대풍산단·팩토리투어")

files = st.file_uploader("후기 CSV 또는 Excel 업로드", type=["csv", "xlsx"], accept_multiple_files=True)
if not files:
    st.info("네이버 블로그, 다음 블로그, 구글맵, 카카오맵 결과 파일을 여러 개 동시에 올릴 수 있습니다.")
    st.stop()

frames, errors = [], []
for uploaded in files:
    try:
        raw = read_file(uploaded)
        if raw.empty:
            st.warning(f"{uploaded.name}: 데이터가 없어 건너뜁니다.")
        else:
            frames.append(standardize(raw, uploaded.name))
    except Exception as exc:
        errors.append(f"{uploaded.name}: {exc}")
if errors:
    st.error("읽지 못한 파일:\n\n" + "\n".join(f"- {e}" for e in errors))
if not frames:
    st.stop()

all_data = deduplicate(pd.concat(frames, ignore_index=True), threshold)
analysis_data = all_data[~all_data["is_duplicate"]].copy() if remove_duplicates else all_data.copy()

st.subheader("필터")
f1, f2, f3 = st.columns(3)
location_options = ["전체"] + sorted(analysis_data["museum_location"].dropna().unique().tolist())
platform_options = ["전체"] + sorted(analysis_data["platform"].dropna().unique().tolist())
sentiment_options = ["전체", "긍정", "중립", "부정"]
selected_location = f1.selectbox("관 선택", location_options)
selected_platform = f2.selectbox("플랫폼 선택", platform_options)
selected_sentiment = f3.selectbox("감성 선택", sentiment_options)

filtered = analysis_data.copy()
if selected_location != "전체": filtered = filtered[filtered["museum_location"] == selected_location]
if selected_platform != "전체": filtered = filtered[filtered["platform"] == selected_platform]
if selected_sentiment != "전체": filtered = filtered[filtered["sentiment"] == selected_sentiment]

counts = filtered["sentiment"].value_counts()
metrics = st.columns(5)
metrics[0].metric("후기", len(filtered))
metrics[1].metric("긍정", int(counts.get("긍정", 0)))
metrics[2].metric("중립", int(counts.get("중립", 0)))
metrics[3].metric("부정", int(counts.get("부정", 0)))
metrics[4].metric("중복 제외", int(all_data["is_duplicate"].sum()))

left, right = st.columns(2)
with left:
    st.subheader("관별 감성 분포")
    loc_chart = pd.crosstab(filtered["museum_location"], filtered["sentiment"]).reindex(columns=["긍정", "중립", "부정"], fill_value=0)
    st.bar_chart(loc_chart, color=["#2E8B57", "#8A8A8A", "#D9534F"])
with right:
    st.subheader("플랫폼별 감성 분포")
    platform_chart = pd.crosstab(filtered["platform"], filtered["sentiment"]).reindex(columns=["긍정", "중립", "부정"], fill_value=0)
    st.bar_chart(platform_chart, color=["#2E8B57", "#8A8A8A", "#D9534F"])

left2, right2 = st.columns(2)
with left2:
    st.subheader("감성 비율")
    sentiment_chart = filtered["sentiment"].value_counts().reindex(["긍정", "중립", "부정"], fill_value=0)
    st.bar_chart(sentiment_chart, horizontal=True, color="#4C78A8")
with right2:
    st.subheader("주요 후기 주제")
    topic_chart = filtered["primary_topic"].value_counts()
    st.bar_chart(topic_chart, horizontal=True, color="#F2A65A")

st.subheader("음성관·서울관 비교표")
comparison = pd.crosstab(filtered["museum_location"], filtered["sentiment"], margins=True).reindex(columns=["긍정", "중립", "부정", "All"], fill_value=0)
comparison = comparison.rename(columns={"All": "합계"})
st.dataframe(comparison, use_container_width=True)

st.subheader("후기 상세")
display_columns = ["museum_location", "platform", "published_date", "title", "content_summary", "rating", "sentiment", "sentiment_score", "primary_topic", "url"]
st.dataframe(filtered[display_columns], use_container_width=True, hide_index=True)

st.download_button(
    "전체 분석 결과 Excel 다운로드", excel_bytes(analysis_data),
    file_name=f"한독의약박물관_후기감성분석_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary"
)
st.caption("감성 분석과 관 구분은 키워드 기반 참고 결과입니다. 공식 평가 자료에는 원문을 사람이 최종 검토해 주세요. 미분류 후기는 상세표에서 확인 후 원문 기준으로 검토하세요.")
