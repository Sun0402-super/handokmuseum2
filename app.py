# -*- coding: utf-8 -*-
import csv, io, re
from datetime import date
import streamlit as st

st.set_page_config(page_title="한독의약박물관 감성 분석기", page_icon="💊", layout="wide")
st.title("💊 한독의약박물관 후기 감성 분석기")
st.caption("CSV 후기 파일을 올리면 중복을 정리하고 긍정·중립·부정을 분석합니다.")

POS = ["좋다","좋았","좋은","유익","재미있","재밌","친절","알차","추천","만족","깨끗","아이가좋아","설명이좋","체험이좋","의미있","다시방문","최고","훌륭","흥미","신기"]
NEG = ["아쉽","불편","어렵","별로","부족","실망","복잡","주차불편","안내부족","찾기어렵","기대이하","시간이짧","불친절","지루","재미없","협소","좁"]

def decode(raw):
    for enc in ("utf-8-sig","utf-8","cp949","euc-kr"):
        try: return raw.decode(enc)
        except UnicodeDecodeError: pass
    return raw.decode("utf-8", errors="replace")

def val(r, *keys):
    for k in keys:
        v=r.get(k,"")
        if v is not None and str(v).strip(): return str(v).strip()
    return ""

def platform(name, url=""):
    s=(name+" "+url).lower()
    if "blog.naver.com" in s or "네이버" in s: return "네이버블로그"
    if "다음" in s or "daum" in s or "tistory" in s: return "다음블로그"
    if "구글" in s or "google" in s: return "구글맵"
    if "카카오" in s or "kakao" in s: return "카카오맵"
    if "인스타" in s or "instagram" in s: return "인스타그램"
    return "기타"

def sentiment(text, rating=""):
    compact=re.sub(r"\s+","",text)
    p=sum(compact.count(w.replace(" ","")) for w in POS)
    n=sum(compact.count(w.replace(" ","")) for w in NEG)
    try:
        star=float(rating)
        if star >= 4: label="긍정"
        elif star <= 2: label="부정"
        else: label="중립"
        basis="별점+텍스트"
    except Exception:
        label="긍정" if p>n else ("부정" if n>p else "중립")
        basis="텍스트"
    return label,p-n,basis

def norm(s): return re.sub(r"[^0-9a-z가-힣]","",s.lower())

files=st.file_uploader("CSV 파일 업로드", type=["csv"], accept_multiple_files=True)
if not files:
    st.info("네이버·다음·구글맵·카카오맵 CSV를 여러 개 동시에 올릴 수 있습니다.")
    st.stop()

rows=[]
for f in files:
    try:
        text=decode(f.getvalue())
        data=list(csv.DictReader(io.StringIO(text)))
        if not data:
            st.warning(f"{f.name}: 데이터가 없어 건너뜁니다.")
            continue
        for r in data:
            url=val(r,"URL","url")
            title=val(r,"제목","title")
            body=val(r,"리뷰내용","본문","내용","요약","full_text","content_summary")
            plat=platform(f.name,url)
            rating=val(r,"별점","rating")
            lab,score,basis=sentiment(title+" "+body,rating)
            rows.append({
              "source":val(r,"장소","source") or plat,"platform":plat,"title":title,
              "content_summary":val(r,"요약","content_summary"),"full_text":body,"url":url,
              "author":val(r,"작성자","블로그명","블로그ID","author"),
              "published_date":val(r,"작성일","published_date"),"collected_date":date.today().isoformat(),
              "rating":rating,"sentiment":lab,"sentiment_score":score,"keyword":"한독의약박물관",
              "is_duplicate":False,"duplicate_group":"","priority_source":plat,"notes":basis})
    except Exception as e:
        st.error(f"{f.name} 읽기 오류: {e}")

# URL 우선 중복 제거, 다음 검색에 잡힌 네이버 URL은 네이버로 보정
seen={}; out=[]; dup_no=0
priority={"네이버블로그":0,"다음블로그":1,"구글맵":2,"카카오맵":3,"인스타그램":4,"기타":5}
rows.sort(key=lambda r: priority.get(r["platform"],9))
for r in rows:
    key=r["url"].strip() or norm(r["title"])
    if key and key in seen:
        dup_no+=1; r["is_duplicate"]=True; r["duplicate_group"]=seen[key]
        continue
    gid=f"D{len(seen)+1:04d}"
    if key: seen[key]=gid
    r["duplicate_group"]=gid
    out.append(r)

st.success(f"분석 완료: {len(out)}건 (중복 제외 {dup_no}건)")
counts={k:sum(1 for r in out if r["sentiment"]==k) for k in ("긍정","중립","부정")}
c1,c2,c3,c4=st.columns(4)
c1.metric("전체",len(out)); c2.metric("긍정",counts["긍정"]); c3.metric("중립",counts["중립"]); c4.metric("부정",counts["부정"])
st.subheader("분석 결과")
st.dataframe(out, use_container_width=True)

cols=["source","platform","title","content_summary","full_text","url","author","published_date","collected_date","rating","sentiment","sentiment_score","keyword","is_duplicate","duplicate_group","priority_source","notes"]
buf=io.StringIO(); w=csv.DictWriter(buf,fieldnames=cols); w.writeheader(); w.writerows(out)
st.download_button("분석 결과 CSV 다운로드", buf.getvalue().encode("utf-8-sig"), f"한독의약박물관_감성분석_{date.today().isoformat()}.csv", "text/csv", type="primary")
st.caption("감성 분석은 참고용입니다. 공식 평가 자료로 사용할 때는 원문을 반드시 사람이 검토해 주세요.")
