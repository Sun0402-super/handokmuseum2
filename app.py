# -*- coding: utf-8 -*-
"""한독의약박물관 온라인 관람 후기 분석 Streamlit 앱."""
from __future__ import annotations

import io
import re
import math
from datetime import date
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

st.set_page_config(page_title="한독의약박물관 후기 분석", page_icon="🏛️", layout="wide")

DEFAULT_KEYWORDS = [
    "한독의약박물관", "한독 의약 박물관", "음성 한독의약박물관", "한독의약박물관 후기",
    "Handok Museum of Medicine", "Handok Museum", "Handok Museum of Medicine review",
]

POSITIVE = [
    "좋다", "좋았", "좋은", "유익", "재미있", "재밌", "친절", "알차", "추천", "만족",
    "깨끗", "아이가 좋아", "설명이 좋", "체험이 좋", "의미 있", "의미있", "다시 방문",
    "최고", "훌륭", "멋", "예쁘", "흥미", "신기", "교육적", "가볼만", "기대 이상",
]
NEGATIVE = [
    "아쉽", "불편", "어렵", "별로", "부족", "실망", "복잡", "주차 불편", "안내 부족",
    "찾기 어렵", "기대 이하", "시간이 짧", "불친절", "지루", "재미없", "협소", "좁",
    "비싸", "더럽", "웨이팅", "오래 기다", "볼 게 없", "규모가 작", "돈 아깝",
]
TOPICS = {
    "약사 체험 프로그램": ["약사", "체험", "약만들", "조제", "프로그램", "체험학습"],
    "해설·설명": ["해설", "도슨트", "설명", "안내문", "가이드", "오디오"],
    "전시·유물 관람": ["전시", "유물", "약재", "의료기구", "역사", "볼거리", "소장품"],
    "아이·가족 나들이": ["아이", "가족", "아이와", "초등", "유아", "어린이", "부모"],
    "건물·공간·사진": ["건물", "정원", "산책", "사진", "포토", "경치", "조경", "분위기"],
    "친절·응대": ["친절", "응대", "직원", "선생님", "안내원"],
    "주차·접근성": ["주차", "교통", "접근", "찾기", "길찾", "대중교통"],
    "규모·볼거리 아쉬움": ["규모", "작고", "볼거리부족", "금방", "볼게없", "협소"],
    "대기·혼잡": ["웨이팅", "대기", "기다", "붐비", "혼잡", "사람많"],
}
OUTPUT_COLUMNS = [
    "source", "platform", "title", "content_summary", "full_text", "url", "author",
    "published_date", "collected_date", "rating", "sentiment", "sentiment_score", "keyword",
    "is_duplicate", "duplicate_group", "priority_source", "notes",
]

ALIASES = {
    "title": ["title", "제목"], "content_summary": ["content_summary", "요약"],
    "full_text": ["full_text", "내용", "리뷰내용", "본문"], "url": ["url", "URL"],
    "author": ["author", "작성자", "블로그명", "블로그ID"],
    "published_date": ["published_date", "작성일"], "rating": ["rating", "별점"],
    "platform": ["platform", "채널", "출처"], "source": ["source", "장소"],
}

def first_value(row, names):
    for name in names:
        if name in row.index and pd.notna(row[name]) and str(row[name]).strip():
            return row[name]
    return ""

def infer_platform(filename: str) -> str:
    name = filename.lower()
    for key, value in [("네이버", "네이버블로그"), ("다음", "다음블로그"),
                       ("구글", "구글맵"), ("카카오", "카카오맵"), ("인스타", "인스타그램")]:
        if key in name:
            return value
    return "기타"

def read_uploaded(uploaded) -> pd.DataFrame:
    name = uploaded.name.lower()
    if name.endswith(".xlsx"):
        return pd.read_excel(uploaded, engine="openpyxl")
    raw = uploaded.getvalue()
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("파일 인코딩을 읽지 못했습니다.")

def standardize(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    platform_default = infer_platform(filename)
    records = []
    for _, row in df.iterrows():
        rec = {key: first_value(row, names) for key, names in ALIASES.items()}
        rec["platform"] = rec["platform"] or platform_default
        rec["source"] = rec["source"] or rec["platform"]
        rec["title"] = str(rec["title"] or "").strip()
        rec["content_summary"] = str(rec["content_summary"] or "").strip()
        rec["full_text"] = str(rec["full_text"] or rec["content_summary"] or "").strip()
        rec["collected_date"] = date.today().isoformat()
        rec["keyword"] = "한독의약박물관"
        rec["notes"] = ""
        records.append(rec)
    return pd.DataFrame(records)

def valid_rating(value):
    try:
        number = float(value)
        return None if math.isnan(number) else number
    except (TypeError, ValueError):
        return None

def analyze_sentiment(text: str, rating=None):
    compact = str(text or "").replace(" ", "")
    pos = sum(compact.count(w.replace(" ", "")) for w in POSITIVE)
    neg = sum(compact.count(w.replace(" ", "")) for w in NEGATIVE)
    score = pos - neg
    star = valid_rating(rating)
    if star is not None:
        label = "긍정" if star >= 4 else ("부정" if star <= 2 else "중립")
        return label, score, "별점 참고"
    label = "긍정" if score > 0 else ("부정" if score < 0 else "중립")
    return label, score, "텍스트 키워드"

def normalize_text(value: str) -> str:
    return re.sub(r"[^0-9a-z가-힣]", "", str(value).lower())

def deduplicate(df: pd.DataFrame, threshold: float) -> pd.DataFrame:
    work = df.copy().reset_index(drop=True)
    priority = {"네이버블로그": 0, "네이버": 0, "다음블로그": 1, "다음": 1}
    work["_priority"] = work["platform"].map(priority).fillna(2)
    work = work.sort_values("_priority", kind="stable").reset_index(drop=True)
    kept_norms, groups = [], []
    duplicate_flags, group_ids = [], []
    for _, row in work.iterrows():
        basis = row["title"] or row["full_text"][:100] or row["url"]
        norm = normalize_text(basis)
        matched = None
        for i, previous in enumerate(kept_norms):
            if norm and previous and SequenceMatcher(None, norm, previous).ratio() >= threshold:
                matched = groups[i]
                break
        if matched is None:
            group = f"D{len(groups)+1:04d}"
            kept_norms.append(norm)
            groups.append(group)
            duplicate_flags.append(False)
            group_ids.append(group)
        else:
            duplicate_flags.append(True)
            group_ids.append(matched)
    work["is_duplicate"] = duplicate_flags
    work["duplicate_group"] = group_ids
    work["priority_source"] = work["platform"].where(~work["is_duplicate"], "중복 제외")
    return work.drop(columns="_priority")

def create_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    summary = (df.groupby(["platform", "sentiment"]).size().unstack(fill_value=0)
               .reindex(columns=["긍정", "중립", "부정"], fill_value=0))
    summary["합계"] = summary.sum(axis=1)
    topic_rows = []
    active = df[~df["is_duplicate"]]
    for topic, words in TOPICS.items():
        mask = active["full_text"].fillna("").map(lambda x: any(w.replace(" ", "") in str(x).replace(" ", "") for w in words))
        topic_rows.append({"주제": topic, "건수": int(mask.sum()), "비율(%)": round(mask.mean()*100, 1) if len(mask) else 0})
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df[OUTPUT_COLUMNS].to_excel(writer, index=False, sheet_name="후기분석")
        summary.to_excel(writer, sheet_name="플랫폼요약")
        pd.DataFrame(topic_rows).sort_values("건수", ascending=False).to_excel(writer, index=False, sheet_name="주제분석")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for col in sheet.columns:
                width = min(max(len(str(c.value or "")) for c in col) + 2, 60)
                sheet.column_dimensions[col[0].column_letter].width = width
    return output.getvalue()

st.title("🏛️ 한독의약박물관 온라인 관람 후기 분석")
st.caption("기존 수집 프로그램의 CSV 또는 Excel 결과를 업로드하면 통합, 중복 제거, 감성 및 주제 분석 후 Excel로 내려받을 수 있습니다.")

with st.sidebar:
    st.header("분석 설정")
    selected_keyword = st.selectbox("기본 검색 키워드", DEFAULT_KEYWORDS)
    threshold = st.slider("중복 판정 유사도", 0.70, 1.00, 0.90, 0.01)
    remove_duplicates = st.checkbox("중복 행을 최종 결과에서 제외", value=True)
    st.info("네이버와 다음 블로그가 중복이면 네이버를 먼저 유지합니다.")

files = st.file_uploader(
    "수집 결과 파일을 올려주세요", type=["csv", "xlsx"], accept_multiple_files=True,
    help="네이버 블로그, 다음 블로그, 구글맵, 카카오맵, 인스타그램 결과를 함께 올릴 수 있습니다.",
)

if not files:
    st.info("먼저 CSV 또는 XLSX 파일을 1개 이상 올려주세요. 서버에서는 로그인이나 차단 우회 방식의 자동 수집을 실행하지 않습니다.")
    st.stop()

frames, errors = [], []
for uploaded in files:
    try:
        frames.append(standardize(read_uploaded(uploaded), uploaded.name))
    except Exception as exc:
        errors.append(f"{uploaded.name}: {exc}")
if errors:
    st.warning("일부 파일을 읽지 못했습니다.\n\n" + "\n".join(f"- {e}" for e in errors))
if not frames:
    st.error("분석 가능한 파일이 없습니다.")
    st.stop()

data = pd.concat(frames, ignore_index=True)
data["keyword"] = selected_keyword
sentiments = data.apply(lambda r: analyze_sentiment(f"{r['title']} {r['full_text']}", r["rating"]), axis=1)
data[["sentiment", "sentiment_score", "notes"]] = pd.DataFrame(sentiments.tolist(), index=data.index)
data = deduplicate(data, threshold)
final = data[~data["is_duplicate"]].copy() if remove_duplicates else data.copy()
for col in OUTPUT_COLUMNS:
    if col not in final.columns:
        final[col] = ""

counts = final["sentiment"].value_counts()
c1, c2, c3, c4 = st.columns(4)
c1.metric("분석 후기", len(final))
c2.metric("긍정", int(counts.get("긍정", 0)))
c3.metric("중립", int(counts.get("중립", 0)))
c4.metric("부정", int(counts.get("부정", 0)))

st.subheader("플랫폼별 감성 분포")
chart = final.groupby(["platform", "sentiment"]).size().unstack(fill_value=0).reindex(columns=["긍정", "중립", "부정"], fill_value=0)
st.bar_chart(chart)

st.subheader("분석 결과")
st.dataframe(final[OUTPUT_COLUMNS], use_container_width=True, hide_index=True)

excel = create_excel(final)
st.download_button(
    "분석 결과 Excel 다운로드", data=excel,
    file_name=f"한독의약박물관_후기분석_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
)
st.caption("감성 분석은 키워드와 별점을 토대로 한 참고용 결과입니다. 공식 평가 자료에는 원문을 사람이 최종 검토해 주세요.")
