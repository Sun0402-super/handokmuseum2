# -*- coding: utf-8 -*-
"""한독의약박물관 온라인 후기 감성 분석 Streamlit 대시보드.

주요 기능
- CSV/XLSX 여러 파일 통합
- 음성관/서울관 자동 분류
- 네이버/다음 중복 시 네이버 우선 유지
- 부정 예상 뒤의 긍정 결론 등 반전 문맥 반영
- 관별/플랫폼별 감성 그래프와 Excel 다운로드
"""
from __future__ import annotations

import io
import re
from datetime import date
from difflib import SequenceMatcher

import pandas as pd
import streamlit as st

st.set_page_config(page_title="한독의약박물관 후기 감성 분석", page_icon="💊", layout="wide")

# 일반 감성 표현
POSITIVE = [
    "좋다", "좋았다", "좋은", "유익", "재미있", "재밌", "친절", "알차", "추천", "만족",
    "깨끗", "아이가 좋아", "아이들이 좋아", "설명이 좋", "체험이 좋", "의미 있", "의미있",
    "다시 방문", "최고", "훌륭", "흥미", "신기", "즐겁", "볼거리", "교육적", "가볼만",
    "다양", "풍부", "도움", "잘 되어", "잘되어", "인상적", "기대 이상", "생각보다 좋"
]
NEGATIVE = [
    "아쉽", "불편", "어렵", "별로", "부족", "실망", "복잡", "주차 불편", "안내 부족",
    "찾기 어렵", "기대 이하", "시간이 짧", "불친절", "지루", "재미없", "협소", "좁",
    "기다", "웨이팅", "볼 게 없", "볼게 없", "관리 안", "추천하지 않"
]

# 강한 최종 평가에는 일반 단어보다 높은 가중치를 부여
STRONG_POSITIVE = [
    "너무 만족", "매우 만족", "정말 만족", "적극 추천", "꼭 추천", "다시 방문하고 싶",
    "재방문하고 싶", "아이들이 정말 좋아", "아이들이 즐거워", "너무 재밌", "정말 재미있",
    "기대 이상", "생각보다 훨씬 좋", "전혀 지루하지 않", "지루하지 않", "지루하지 않았"
]
STRONG_NEGATIVE = [
    "매우 실망", "너무 실망", "다시는 방문하지 않", "추천하지 않", "돈이 아깝",
    "시간 낭비", "전혀 재미없", "너무 지루", "매우 불편"
]

# 부정 예상이나 걱정이 뒤의 긍정 경험으로 뒤집히는 대표 패턴
POSITIVE_REVERSAL_PATTERNS = [
    r"지루(?:할|해할|해하지 않을|하지 않을|하지는 않을|하지 않을까)?[^.!?\n]{0,45}(?:그런데|하지만|했는데|였는데|아니|전혀|오히려|생각보다)[^.!?\n]{0,60}(?:재미|즐겁|흥미|만족|좋|다양|알차)",
    r"지루[^.!?\n]{0,40}(?:전혀\s*)?(?:아니었|않았|않다|않고|하지 않았|하지 않)",
    r"(?:걱정|우려)[^.!?\n]{0,45}(?:했는데|했지만|과 달리|와 달리|오히려|전혀)[^.!?\n]{0,60}(?:재미|즐겁|흥미|만족|좋|다양|알차)",
    r"(?:기대하지 않았|기대 없이|별 기대 없|큰 기대 없)[^.!?\n]{0,45}(?:그런데|했는데|였는데|오히려|생각보다)[^.!?\n]{0,60}(?:좋|만족|재미|즐겁|흥미|다양|알차|볼거리)",
    r"어려울[^.!?\n]{0,35}(?:줄 알았|것 같았|까 걱정)[^.!?\n]{0,45}(?:그런데|했는데|였는데|오히려|생각보다)[^.!?\n]{0,60}(?:쉽|이해|재미|도움|좋)",
    r"(?:아쉽|부족|별로|실망|불편)[^.!?\n]{0,35}(?:줄 알았|것 같았|까 걱정)[^.!?\n]{0,45}(?:아니|않았|오히려|생각보다)[^.!?\n]{0,55}(?:좋|만족|괜찮|편리|충분)",
]

# 부정 표현 자체가 부정되면 긍정 보정
NEGATED_NEGATIVE_PATTERNS = [
    r"지루(?:하지|하지는|하지도|하지가)\s*(?:않|않았|않고|않아서|않다)",
    r"전혀\s*지루하지\s*(?:않|않았)",
    r"재미없(?:지|지는|지도)\s*(?:않|않았)",
    r"불편(?:하지|하지는|하지도)\s*(?:않|않았)",
    r"어렵(?:지|지는|지도)\s*(?:않|않았)",
    r"아쉽(?:지|지는|지도)\s*(?:않|않았)",
    r"부족(?:하지|하지는|하지도)\s*(?:않|않았)",
    r"실망(?:하지|하지는|하지도)\s*(?:않|않았)",
    r"별로일\s*줄\s*알았는데[^.!?\n]{0,50}(?:좋|만족|괜찮|재미)",
]

# 긍정 표현이 부정되는 경우
NEGATED_POSITIVE_PATTERNS = [
    r"좋(?:지|지는|지도)\s*(?:않|않았)",
    r"재미있(?:지|지는|지도)\s*(?:않|않았)",
    r"만족(?:하지|하지는|하지도)\s*(?:않|않았)",
    r"추천(?:하지|하지는|하지도)\s*(?:않|않았)",
    r"친절(?:하지|하지는|하지도)\s*(?:않|않았)",
]

# 역접 뒤의 결론에 더 높은 가중치를 주기 위한 접속 표현
CONTRAST_MARKERS = ["하지만", "그런데", "그렇지만", "했는데", "였는데", "반면", "오히려", "생각보다", "전혀"]

# 실제 방문 경험에 대한 명확한 불만만 부정 판정의 핵심 근거로 사용한다.
# 예약 매진의 아쉬움, 아이가 힘을 써야 했다는 설명, "지루할까 걱정" 같은 예상은 제외한다.
EXPLICIT_COMPLAINT_PATTERNS = [
    r"(?:전시|체험|관람|시설|주차|안내|직원|설명|프로그램)(?:이|가|은|는|도)?[^.!?\n]{0,25}(?:별로였|실망스러|불편했|부족했|지루했|재미없었|아쉬웠|협소했|불친절했)",
    r"(?:다시는|재방문은)[^.!?\n]{0,25}(?:안|않|못)",
    r"(?:추천하지|비추천|돈이아깝|시간낭비|볼게없|볼 게 없)",
    r"(?:주차|길찾기|접근성)[^.!?\n]{0,20}(?:너무|매우)?[^.!?\n]{0,15}(?:불편|어렵)",
]

# 글 전체의 최종 추천 및 만족 표현. 발견되면 결론에 높은 가중치를 준다.
FINAL_POSITIVE_PATTERNS = [
    r"(?:강력|적극|꼭)?\s*추천(?:합니다|해요|드려요|하고 싶|할 만)",
    r"(?:정말|너무|매우)?\s*만족(?:스러웠|했|합니다|해요)",
    r"(?:유익함과 재미|유익하고 즐거|재미와 유익)",
    r"(?:최고의|뜻깊은|값진|행복한|즐거운)\s*(?:시간|체험|관람|수업|기억)",
    r"(?:다시|또)\s*(?:방문|가고 싶|오고 싶)",
    r"아이(?:들|가)?[^.!?\n]{0,30}(?:좋아했|즐거워했|흥미로워했|집중했)",
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
    "is_relevant", "relevance_reason", "target_mentions",
    "author", "published_date", "collected_date", "rating", "sentiment", "sentiment_score",
    "positive_score", "negative_score", "positive_evidence", "negative_evidence",
    "reversal_evidence", "keyword", "is_duplicate", "duplicate_group", "priority_source",
    "primary_topic", "notes"
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

def classify_relevance(title, text, url, source, platform=""):
    """한독의약박물관 자체 후기인지 먼저 판정한다.

    검색 결과에 '한독의약박물관'이 비교 대상으로 한 번 언급되거나,
    '음성'이라는 지역명만 포함된 다른 관광지가 섞이는 문제를 차단한다.
    """
    title_norm = re.sub(r"\s+", "", str(title)).lower()
    body_norm = re.sub(r"\s+", "", str(text)).lower()
    source_norm = re.sub(r"\s+", "", str(source)).lower()
    url_norm = str(url).lower()
    whole = f"{title_norm} {body_norm} {source_norm} {url_norm}"

    # 지도 리뷰는 파일/플랫폼 자체가 장소를 특정한다. 리뷰 본문에 박물관명이 없더라도 정상이다.
    # 따라서 한독의약박물관 장소 페이지에서 수집된 지도 리뷰는 직접 언급 여부로 제외하지 않는다.
    source_platform = f"{source_norm} {str(platform).lower()}"
    is_map_review = any(term in source_platform for term in ["구글맵", "구글", "google", "카카오맵", "카카오", "kakao"])
    if is_map_review:
        return True, "지도 장소 리뷰: 본문 직접 언급 불필요", 0

    target_terms = [
        "한독의약박물관", "한독박물관", "handokmuseum", "handokmuseumofmedicine",
        "대풍산단로78", "한독퓨처콤플렉스", "한독공장"
    ]
    target_mentions = sum(whole.count(term) for term in target_terms)
    title_has_target = any(term in title_norm for term in target_terms)
    source_has_target = any(term in source_norm for term in target_terms)

    # 이번 데이터에서 실제로 확인된 대표 오수집 대상과 일반적인 타 박물관 표지
    unrelated_terms = [
        "미래엔교과서박물관", "교과서박물관", "봉학골산림욕장", "품바재생예술촌",
        "품바예술촌", "반기문평화기념관", "철박물관"
    ]
    unrelated_hits = [term for term in unrelated_terms if term in title_norm or term in body_norm]

    # 제목이나 장소가 한독의약박물관이면 가장 신뢰도가 높다.
    if title_has_target or source_has_target:
        return True, "제목/장소에 한독의약박물관 명시", target_mentions

    # 다른 명소가 제목에 있고 한독은 비교 문구로만 소수 언급된 경우 제외한다.
    title_unrelated = [term for term in unrelated_terms if term in title_norm]
    if title_unrelated:
        return False, f"다른 장소 후기: {', '.join(title_unrelated)}", target_mentions

    # 본문에 정확한 명칭이 반복되면 관련 후기 가능성이 높다.
    if target_mentions >= 2:
        return True, "본문에 한독의약박물관 반복 언급", target_mentions

    # 한 번만 나온 글은 비교·추천 목록일 가능성이 있어 분석에서 제외한다.
    if target_mentions == 1:
        return False, "한독의약박물관 1회 단순 언급/비교", target_mentions

    # '음성' 지역명만 있는 다른 관광지는 관련 후기가 아니다.
    if unrelated_hits:
        return False, f"다른 관광지 내용: {', '.join(unrelated_hits)}", target_mentions
    return False, "한독의약박물관 직접 언급 없음", target_mentions

def detect_location(title, text, url, source):
    combined = re.sub(r"\s+", "", f"{title} {text} {url} {source}").lower()
    seoul_terms = ["서울관", "한독의약박물관서울", "마곡", "마곡나루", "퓨처콤플렉스", "강서구"]
    eumseong_terms = ["음성", "대소면", "대소읍", "대풍산단", "팩토리투어", "한독공장"]
    if any(term in combined for term in seoul_terms):
        return "서울관"
    if any(term in combined for term in eumseong_terms):
        return "음성관"
    return "미분류"

def unique_join(items):
    """표시용 근거에서 중복을 제거한다."""
    return ", ".join(dict.fromkeys(item for item in items if item))

def sentence_split(text):
    """후기 문장을 단순 분리한다."""
    return [part.strip() for part in re.split(r"(?<=[.!?。])\s+|[\r\n]+", str(text)) if part.strip()]

def keyword_hits(text, keywords):
    compact = re.sub(r"\s+", "", text)
    return [word for word in keywords if word.replace(" ", "") in compact]

def find_pattern_evidence(text, patterns):
    found = []
    for pattern in patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            found.append(match.group(0).strip())
    return found

def analyze_sentence(sentence):
    """문장 단위 감성 점수와 판정 근거를 계산한다."""
    compact = re.sub(r"\s+", "", sentence)
    positive_hits = keyword_hits(sentence, POSITIVE)
    negative_hits = keyword_hits(sentence, NEGATIVE)
    strong_positive_hits = keyword_hits(sentence, STRONG_POSITIVE)
    strong_negative_hits = keyword_hits(sentence, STRONG_NEGATIVE)
    negated_negative = find_pattern_evidence(sentence, NEGATED_NEGATIVE_PATTERNS)
    negated_positive = find_pattern_evidence(sentence, NEGATED_POSITIVE_PATTERNS)
    reversals = find_pattern_evidence(sentence, POSITIVE_REVERSAL_PATTERNS)
    explicit_complaints = find_pattern_evidence(sentence, EXPLICIT_COMPLAINT_PATTERNS)
    final_positives = find_pattern_evidence(sentence, FINAL_POSITIVE_PATTERNS)

    positive_score = len(positive_hits) + 2 * len(strong_positive_hits) + 3 * len(final_positives)
    # 단순 부정 단어는 약한 신호로만 사용하고, 명시적 불만 문장에 높은 가중치를 둔다.
    negative_score = len(negative_hits) + 2 * len(strong_negative_hits) + 3 * len(explicit_complaints)

    # "지루하지 않았다"처럼 부정 단어가 부정된 경우 부정 점수를 제거하고 긍정 점수를 가산
    for evidence in negated_negative:
        related = sum(1 for word in NEGATIVE if word.replace(" ", "") in re.sub(r"\s+", "", evidence))
        negative_score = max(0, negative_score - max(1, related))
        positive_score += 2

    # "좋지 않았다"처럼 긍정 단어가 부정된 경우 반대로 보정
    for evidence in negated_positive:
        related = sum(1 for word in POSITIVE if word.replace(" ", "") in re.sub(r"\s+", "", evidence))
        positive_score = max(0, positive_score - max(1, related))
        negative_score += 2

    # 부정 예상 뒤 실제 경험이 긍정인 경우 강하게 긍정 보정
    if reversals:
        positive_score += 3 * len(reversals)
        # 반전 패턴 안의 부정어는 실제 평가가 아니라 예상/우려이므로 감점하지 않는다.
        for evidence in reversals:
            related = sum(1 for word in NEGATIVE if word.replace(" ", "") in re.sub(r"\s+", "", evidence))
            negative_score = max(0, negative_score - max(1, related))

    # 역접 표현 뒤쪽의 결론을 한 번 더 평가해 최종 경험에 가중치를 준다.
    for marker in CONTRAST_MARKERS:
        if marker in sentence:
            tail = sentence.rsplit(marker, 1)[-1]
            tail_pos = keyword_hits(tail, POSITIVE)
            tail_neg = keyword_hits(tail, NEGATIVE)
            positive_score += len(tail_pos)
            negative_score += len(tail_neg)

    return {
        "positive_score": positive_score,
        "negative_score": negative_score,
        "positive_evidence": positive_hits + strong_positive_hits,
        "negative_evidence": negative_hits + strong_negative_hits,
        "reversal_evidence": negated_negative + negated_positive + reversals,
        "explicit_complaint_evidence": explicit_complaints,
        "final_positive_evidence": final_positives,
    }

def sentiment(text, rating):
    """문장별 결과를 합산하고 별점을 참고해 최종 감성을 정한다."""
    sentences = sentence_split(text) or [str(text)]
    total_pos = 0
    total_neg = 0
    positive_evidence = []
    negative_evidence = []
    reversal_evidence = []
    explicit_complaint_evidence = []
    final_positive_evidence = []

    for sentence in sentences:
        result = analyze_sentence(sentence)
        total_pos += result["positive_score"]
        total_neg += result["negative_score"]
        positive_evidence.extend(result["positive_evidence"])
        negative_evidence.extend(result["negative_evidence"])
        reversal_evidence.extend(result["reversal_evidence"])
        explicit_complaint_evidence.extend(result["explicit_complaint_evidence"])
        final_positive_evidence.extend(result["final_positive_evidence"])

    notes = ["문장 단위 키워드 분석"]
    try:
        star = float(str(rating).replace("점", "").strip())
        if star >= 4:
            total_pos += 2
        elif star <= 2:
            total_neg += 2
        notes.append(f"별점 {star:g}점 반영")
    except (TypeError, ValueError):
        pass

    score = total_pos - total_neg
    # 오분류 방지 원칙: 부정은 실제 방문 경험에 대한 명시적 불만이 있을 때만 판정한다.
    # 명시적 불만이 없고 추천·만족·즐거움 근거가 있으면 긍정으로 처리한다.
    if explicit_complaint_evidence and score < 0:
        label = "부정"
    elif total_pos > 0 or final_positive_evidence or reversal_evidence:
        label = "긍정"
    else:
        label = "중립"
    if reversal_evidence:
        notes.append("반전·부정어 문맥 반영")
    if final_positive_evidence:
        notes.append("최종 추천·만족 표현 반영")
    if negative_evidence and not explicit_complaint_evidence:
        notes.append("부정 단어는 있으나 명시적 불만 없음")

    return (
        label, score, total_pos, total_neg,
        unique_join(positive_evidence + final_positive_evidence), unique_join(negative_evidence),
        unique_join(reversal_evidence), " / ".join(notes)
    )

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
        relevant, relevance_reason, target_mentions = classify_relevance(
            rec["title"], combined, rec["url"], rec["source"], rec["platform"]
        )
        rec["museum_location"] = detect_location(rec["title"], combined, rec["url"], rec["source"]) if relevant else "분석 제외"
        if relevant:
            label, score, pos, neg, pos_ev, neg_ev, rev_ev, note = sentiment(combined, rec["rating"])
        else:
            label, score, pos, neg, pos_ev, neg_ev, rev_ev, note = "분석 제외", 0, 0, 0, "", "", "", "관련성 필터 제외"
        rec.update({
            "is_relevant": relevant, "relevance_reason": relevance_reason,
            "target_mentions": target_mentions,
            "collected_date": date.today().isoformat(), "sentiment": label, "sentiment_score": score,
            "positive_score": pos, "negative_score": neg, "positive_evidence": pos_ev,
            "negative_evidence": neg_ev, "reversal_evidence": rev_ev,
            "keyword": "한독의약박물관", "is_duplicate": False, "duplicate_group": "",
            "priority_source": rec["platform"], "primary_topic": topic(combined), "notes": note,
        })
        records.append(rec)
    return pd.DataFrame(records)

def deduplicate(df, threshold):
    priority = {"네이버 블로그": 0, "다음 블로그": 1, "구글맵": 2, "카카오맵": 3, "인스타그램": 4}
    work = df.copy()
    work["_priority"] = work["platform"].map(priority).fillna(9)
    work = work.sort_values("_priority", kind="stable").reset_index(drop=True)
    representatives = []
    for idx, row in work.iterrows():
        url_key = normalize(row["url"])
        text_key = normalize(row["title"] or row["full_text"][:140])
        match = None
        for _, rep_url, rep_text, gid in representatives:
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
    return work.drop(columns="_priority")

def excel_bytes(data):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        relevant_data = data[data["is_relevant"]].copy()
        excluded_data = data[~data["is_relevant"]].copy()
        relevant_data[OUTPUT_COLUMNS].to_excel(writer, index=False, sheet_name="후기분석")
        excluded_data[OUTPUT_COLUMNS].to_excel(writer, index=False, sheet_name="관련없음_제외목록")
        pd.crosstab(relevant_data["museum_location"], relevant_data["sentiment"], margins=True).to_excel(writer, sheet_name="관별감성요약")
        pd.crosstab(relevant_data["platform"], relevant_data["sentiment"], margins=True).to_excel(writer, sheet_name="플랫폼요약")
        relevant_data.groupby(["museum_location", "primary_topic"]).size().rename("건수").reset_index().to_excel(writer, index=False, sheet_name="주제분석")
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column in ws.columns:
                width = min(max(len(str(cell.value or "")) for cell in column) + 2, 55)
                ws.column_dimensions[column[0].column_letter].width = width
    return output.getvalue()

st.title("💊 한독의약박물관 온라인 후기 감성 분석기")
st.caption("한독의약박물관 자체 후기만 먼저 선별한 뒤, 음성관·서울관 구분과 반전 문맥을 반영해 감성을 분석합니다.")

with st.sidebar:
    st.header("분석 설정")
    threshold = st.slider("중복 판정 유사도", 0.70, 1.00, 0.90, 0.01)
    remove_duplicates = st.checkbox("중복 후기를 분석에서 제외", value=True)
    st.markdown("**관 구분 기준**")
    st.caption("서울관: 서울관·마곡·퓨처콤플렉스·강서구\n\n음성관: 음성·대소·대풍산단·팩토리투어")
    st.info("부정은 실제 방문 경험에 대한 명시적 불만이 있을 때만 판정합니다. 지도 리뷰는 장소 페이지 자체가 대상을 특정하므로 본문에 박물관명이 없어도 분석에 포함합니다.")

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

raw_data = pd.concat(frames, ignore_index=True)
excluded_data = raw_data[~raw_data["is_relevant"]].copy()
relevant_data = raw_data[raw_data["is_relevant"]].copy()
if relevant_data.empty:
    st.error("한독의약박물관 자체 후기로 판정된 글이 없습니다. 아래 제외 목록을 확인해 주세요.")
    st.dataframe(excluded_data[["title", "url", "relevance_reason"]], use_container_width=True, hide_index=True)
    st.stop()
all_data = deduplicate(relevant_data, threshold)
analysis_data = all_data[~all_data["is_duplicate"]].copy() if remove_duplicates else all_data.copy()

with st.expander(f"관련 없는 검색 결과 제외 목록 ({len(excluded_data)}건)"):
    st.caption("다른 박물관·관광지 후기, 비교 문구 1회 언급, 직접 언급이 없는 글은 감성 분석 전에 제외됩니다.")
    st.dataframe(excluded_data[["title", "platform", "relevance_reason", "url"]], use_container_width=True, hide_index=True)

st.subheader("필터")
f1, f2, f3 = st.columns(3)
location_options = ["전체"] + sorted(analysis_data["museum_location"].dropna().unique().tolist())
platform_options = ["전체"] + sorted(analysis_data["platform"].dropna().unique().tolist())
sentiment_options = ["전체", "긍정", "중립", "부정"]
selected_location = f1.selectbox("관 선택", location_options)
selected_platform = f2.selectbox("플랫폼 선택", platform_options)
selected_sentiment = f3.selectbox("감성 선택", sentiment_options)

filtered = analysis_data.copy()
if selected_location != "전체":
    filtered = filtered[filtered["museum_location"] == selected_location]
if selected_platform != "전체":
    filtered = filtered[filtered["platform"] == selected_platform]
if selected_sentiment != "전체":
    filtered = filtered[filtered["sentiment"] == selected_sentiment]

counts = filtered["sentiment"].value_counts()
metrics = st.columns(6)
metrics[0].metric("후기", len(filtered))
metrics[1].metric("긍정", int(counts.get("긍정", 0)))
metrics[2].metric("중립", int(counts.get("중립", 0)))
metrics[3].metric("부정", int(counts.get("부정", 0)))
metrics[4].metric("관련 없음 제외", len(excluded_data))
metrics[5].metric("중복 제외", int(all_data["is_duplicate"].sum()))

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

st.subheader("후기 상세 및 판정 근거")
display_columns = [
    "museum_location", "platform", "published_date", "title", "content_summary", "rating",
    "sentiment", "sentiment_score", "positive_evidence", "negative_evidence",
    "reversal_evidence", "primary_topic", "url"
]
st.dataframe(filtered[display_columns], use_container_width=True, hide_index=True)

st.download_button(
    "전체 분석 결과 Excel 다운로드", excel_bytes(pd.concat([analysis_data, excluded_data], ignore_index=True)),
    file_name=f"한독의약박물관_후기감성분석_{date.today().isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary"
)
st.caption("관련성·감성·관 구분은 규칙 기반 참고 결과입니다. 공식 평가 자료에는 판정 근거와 원문을 사람이 최종 검토해 주세요.")
