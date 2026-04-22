
import io
import re
from pathlib import Path
from typing import Optional, List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="굿윌 매출 대시보드", page_icon="📊", layout="wide")

MONTH_PATTERNS = [
    re.compile(r"(20\d{2})[-_]?([01]?\d)"),
    re.compile(r"(20\d{2})년\s*([01]?\d)월"),
]

DAILY_NUM_COLS = ["전표건수","객단가","공급가액","실매출액","현금","현금영수증","카드","포인트","현금카드외","제휴포인트","상품권결제"]
PRODUCT_NUM_COLS = ["판매수량","공급가액","실매출액","기본단가","이익금액","현금","현금영수증","카드","상품권결제","현금카드외"]

MAJOR_DONORS = ["CJ제일제당","편의점","모던하우스","오뚜기","신세계푸드"]
DONOR_PATTERNS = {
    "CJ제일제당": [r"CJ제일제당", r"\bCJ\b"],
    "편의점": [r"GS25", r"\bCU\b", r"세븐일레븐", r"편의점"],
    "모던하우스": [r"모던하우스", r"기증파트너 》 모던", r"\b모던\b"],
    "오뚜기": [r"오뚜기"],
    "신세계푸드": [r"신세계푸드"],
}
CATEGORY_GROUPS = {
    "의류": ["의류"],
    "잡화": ["잡화"],
    "생활": ["생활용품"],
    "식품": ["식품"],
    "건강/미용": ["건강/미용"],
    "문화": ["문화용품"],
    "원가상품": ["원가상품"],
    "기타": []
}

def clean_number(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("%", "")
    if s == "":
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0

def pct(a, b):
    if b in [0, None] or pd.isna(b):
        return 0.0
    return float(a) / float(b)

def extract_month_from_filename(filename: str) -> Optional[str]:
    name = Path(filename).stem
    for pattern in MONTH_PATTERNS:
        m = pattern.search(name)
        if m:
            y = int(m.group(1))
            mm = int(m.group(2))
            if 1 <= mm <= 12:
                return f"{y:04d}-{mm:02d}"
    return None

def month_label(month_str: str) -> str:
    y, m = month_str.split("-")
    return f"{y}년 {int(m)}월"

def month_sheet_label(month_str: str) -> str:
    return f"{int(month_str.split('-')[1])}월"

def parse_daily_sales(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0)
    raw = raw.rename(columns=lambda x: str(x).strip())
    store_col = raw.columns[0]
    current_store = None
    rows = []
    for _, row in raw.iterrows():
        marker = row.get(store_col)
        if isinstance(marker, str) and "매장:" in marker:
            current_store = re.sub(r".*매장:\s*", "", marker).split("[")[0].strip().replace("밀알", "")
            continue
        sale_date = row.get("영업일자")
        if pd.isna(sale_date):
            continue
        try:
            sale_date = pd.to_datetime(sale_date)
        except Exception:
            continue
        out = {"지점명": current_store if current_store else "미확인", "영업일자": sale_date, "기준월": sale_date.strftime("%Y-%m")}
        for col in DAILY_NUM_COLS:
            out[col] = clean_number(row.get(col))
        rows.append(out)
    df = pd.DataFrame(rows)
    if not df.empty:
        df["일"] = df["영업일자"].dt.day
        df["월"] = df["영업일자"].dt.month
        df["연도"] = df["영업일자"].dt.year
        df["객단가_계산"] = df.apply(lambda r: pct(r["실매출액"], r["전표건수"]) if r["전표건수"] else 0, axis=1)
    return df

def parse_product_sales(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0)
    raw = raw.rename(columns=lambda x: str(x).strip())
    month_str = extract_month_from_filename(uploaded_file.name)
    if not month_str:
        raise ValueError(f"상품별 파일명에서 월을 읽을 수 없습니다: {uploaded_file.name}")
    category_hint = None
    rows = []
    for _, row in raw.iterrows():
        left = row.get("매장")
        if isinstance(left, str) and "상품분류1:" in left:
            category_hint = re.sub(r".*상품분류1:\s*", "", left).split("[")[0].strip()
            continue
        store = row.get("Unnamed: 1")
        product = row.get("상품")
        if pd.isna(store) or pd.isna(product):
            continue
        category = str(row.get("상품분류")).strip() if not pd.isna(row.get("상품분류")) else (category_hint or "미분류")
        out = {
            "기준월": month_str,
            "지점명": str(store).strip().replace("밀알",""),
            "상품명": str(product).strip(),
            "상품분류": category,
            "바코드": "" if pd.isna(row.get("바코드")) else str(row.get("바코드")).strip(),
        }
        for col in PRODUCT_NUM_COLS:
            out[col] = clean_number(row.get(col))
        rows.append(out)
    df = pd.DataFrame(rows)
    if not df.empty:
        df["대분류"] = df["상품분류"].astype(str).str.split("》").str[0].str.strip()
        df["기증처"] = df.apply(infer_donor, axis=1)
    return df

def infer_donor(row):
    text = f"{row.get('상품명','')} | {row.get('상품분류','')}"
    for donor, patterns in DONOR_PATTERNS.items():
        for p in patterns:
            if re.search(p, text, re.IGNORECASE):
                return donor
    if "》" in str(row.get("상품분류","")):
        parts = [x.strip() for x in str(row.get("상품분류","")).split("》")]
        if len(parts) >= 2:
            second = parts[1]
            blacklist = {"여성의류","남성의류","아동의류","하의","상의","가방","신발","잡화","문화용품","도서","생활용품","주방용품","가전","건강/미용","기업","식품","매입","제빵"}
            if second not in blacklist and len(second) >= 2:
                return second
    m = re.match(r"([A-Za-z가-힣0-9]+)\)", str(row.get("상품명","")))
    if m:
        return m.group(1)
    return "기타"

@st.cache_data(show_spinner=False)
def load_all_data(daily_files, product_files):
    daily = pd.concat([parse_daily_sales(f) for f in daily_files], ignore_index=True) if daily_files else pd.DataFrame()
    product = pd.concat([parse_product_sales(f) for f in product_files], ignore_index=True) if product_files else pd.DataFrame()
    return daily, product

def build_month_store_summary(daily):
    if daily.empty:
        return pd.DataFrame()
    out = daily.groupby(["기준월","지점명"], as_index=False).agg({
        "전표건수":"sum","공급가액":"sum","실매출액":"sum",
        "현금":"sum","현금영수증":"sum","카드":"sum","포인트":"sum","현금카드외":"sum","제휴포인트":"sum","상품권결제":"sum"
    })
    out["객단가"] = out.apply(lambda r: pct(r["실매출액"], r["전표건수"]) if r["전표건수"] else 0, axis=1)
    out = out.sort_values(["지점명","기준월"])
    out["전월실매출액"] = out.groupby("지점명")["실매출액"].shift(1)
    out["전월대비증감률"] = out.apply(lambda r: pct(r["실매출액"]-r["전월실매출액"], r["전월실매출액"]) if r["전월실매출액"] else 0, axis=1)
    return out

def normalize_category_group(x):
    x = str(x)
    for group, keywords in CATEGORY_GROUPS.items():
        if any(k in x for k in keywords):
            return group
    return "기타"

def build_classification_report(product_month):
    if product_month.empty:
        return pd.DataFrame()
    df = product_month.copy()
    df["분류그룹"] = df["대분류"].apply(normalize_category_group)
    grp = df.groupby(["지점명","분류그룹"], as_index=False).agg({"판매수량":"sum","실매출액":"sum"})
    total = grp.groupby("지점명", as_index=False)["실매출액"].sum().rename(columns={"실매출액":"지점합계"})
    grp = grp.merge(total, on="지점명", how="left")
    grp["점유율"] = grp.apply(lambda r: pct(r["실매출액"], r["지점합계"]), axis=1)
    order = ["의류","잡화","생활","식품","건강/미용","문화","원가상품","기타"]
    grp["정렬"] = grp["분류그룹"].apply(lambda x: order.index(x) if x in order else 999)
    return grp.sort_values(["지점명","정렬","분류그룹"]).drop(columns=["정렬","지점합계"])

def build_donor_report(product_month):
    if product_month.empty:
        return pd.DataFrame()
    grp = product_month.groupby(["지점명","기증처"], as_index=False).agg({"판매수량":"sum","실매출액":"sum"})
    grp["피스단가"] = grp.apply(lambda r: pct(r["실매출액"], r["판매수량"]) if r["판매수량"] else 0, axis=1)
    base = grp[grp["기증처"].isin(MAJOR_DONORS)].copy()
    top_extra = (
        grp[~grp["기증처"].isin(MAJOR_DONORS)]
        .groupby("기증처", as_index=False)["실매출액"].sum()
        .sort_values("실매출액", ascending=False)
        .head(5)["기증처"].tolist()
    )
    extra = grp[grp["기증처"].isin(top_extra)].copy()
    out = pd.concat([base, extra], ignore_index=True)
    totals = out.groupby("기증처", as_index=False)["실매출액"].sum().sort_values("실매출액", ascending=False)
    order = totals["기증처"].tolist()
    out["기증처"] = pd.Categorical(out["기증처"], categories=order, ordered=True)
    return out.sort_values(["지점명","기증처"])

def make_month_analysis_workbook_bytes(daily_df, product_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        months = sorted(product_df["기준월"].dropna().unique().tolist()) if not product_df.empty else sorted(daily_df["기준월"].dropna().unique().tolist())
        for month in months:
            product_month = product_df[product_df["기준월"] == month].copy() if not product_df.empty else pd.DataFrame()
            class_df = build_classification_report(product_month)
            donor_df = build_donor_report(product_month)
            sheet_name = month_sheet_label(month)
            row_cursor = 0
            if not class_df.empty:
                pivot_rows = []
                stores = sorted(class_df["지점명"].unique().tolist())
                categories = ["의류","잡화","생활","식품","건강/미용","문화","원가상품","기타"]
                for cat in categories:
                    row = {"구분": cat}
                    for store in stores:
                        sub = class_df[(class_df["지점명"] == store) & (class_df["분류그룹"] == cat)]
                        qty = float(sub["판매수량"].sum()) if not sub.empty else 0
                        sales = float(sub["실매출액"].sum()) if not sub.empty else 0
                        total_sales = float(class_df[class_df["지점명"] == store]["실매출액"].sum()) if store in class_df["지점명"].values else 0
                        share = pct(sales, total_sales)
                        row[f"{store}_판매수량"] = qty
                        row[f"{store}_실매출액"] = sales
                        row[f"{store}_점유율"] = share
                    pivot_rows.append(row)
                pivot_df = pd.DataFrame(pivot_rows)
                pd.DataFrame([[f"기간: {month_label(month)}", "", "", ""]]).to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=row_cursor)
                row_cursor += 2
                pd.DataFrame([["*분류별"]]).to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=row_cursor)
                row_cursor += 1
                pivot_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row_cursor)
                row_cursor += len(pivot_df) + 3
            if not donor_df.empty:
                pd.DataFrame([["*주요 기증처별"]]).to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=row_cursor)
                row_cursor += 1
                donor_pivot = donor_df.pivot_table(index="기증처", columns="지점명", values=["판매수량","실매출액","피스단가"], aggfunc="sum", fill_value=0)
                donor_pivot.columns = [f"{store}_{metric}" for metric, store in donor_pivot.columns]
                donor_pivot = donor_pivot.reset_index().rename(columns={"기증처":"구분"})
                donor_pivot.to_excel(writer, sheet_name=sheet_name, index=False, startrow=row_cursor)
        if not months:
            pd.DataFrame({"안내":["상품별 파일을 업로드하면 월매출분석자료가 생성됩니다."]}).to_excel(writer, sheet_name="안내", index=False)
    out.seek(0)
    return out.getvalue()

def make_goodwill_sales_workbook_bytes(daily_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        month_store = build_month_store_summary(daily_df)
        if month_store.empty:
            pd.DataFrame({"안내":["매출현황 파일을 업로드하면 굿윌2026매출자료가 생성됩니다."]}).to_excel(writer, sheet_name="안내", index=False)
            out.seek(0)
            return out.getvalue()
        all_months = sorted(daily_df["기준월"].dropna().unique().tolist())
        year = int(all_months[0].split("-")[0])
        stores = sorted(daily_df["지점명"].dropna().unique().tolist())
        summary_rows = []
        month_nums = [int(m.split("-")[1]) for m in all_months]
        for store in stores:
            row = {"지점명": store}
            total = 0
            for m in all_months:
                val = float(month_store[(month_store["지점명"] == store) & (month_store["기준월"] == m)]["실매출액"].sum())
                row[f"{int(m.split('-')[1])}월"] = val
                total += val
            row["합계"] = total
            summary_rows.append(row)
        summary_df = pd.DataFrame(summary_rows)
        summary_df.to_excel(writer, sheet_name="월합계", index=False, startrow=1)
        for m in all_months:
            daily_month = daily_df[daily_df["기준월"] == m].copy().sort_values(["지점명","일"])
            rows = []
            for store in sorted(daily_month["지점명"].unique().tolist()):
                sub = daily_month[daily_month["지점명"] == store]
                sales = {"구분":"매출","지점명":store}
                cnt = {"구분":"영수건수","지점명":store}
                avg = {"구분":"객단가","지점명":store}
                for day in range(1,32):
                    ds = sub[sub["일"] == day]
                    sales[str(day)] = float(ds["실매출액"].sum()) if not ds.empty else 0
                    cnt[str(day)] = float(ds["전표건수"].sum()) if not ds.empty else 0
                    avg[str(day)] = pct(sales[str(day)], cnt[str(day)]) if cnt[str(day)] else 0
                sales["월합계"] = sum(v for k,v in sales.items() if k.isdigit())
                cnt["월합계"] = sum(v for k,v in cnt.items() if k.isdigit())
                avg["월합계"] = pct(sales["월합계"], cnt["월합계"]) if cnt["월합계"] else 0
                rows.extend([sales, cnt, avg])
            month_sheet = pd.DataFrame(rows)
            title = pd.DataFrame([[f"{year}년 {int(m.split('-')[1])}월 매출"]])
            title.to_excel(writer, sheet_name=month_sheet_label(m), index=False, header=False, startrow=0)
            month_sheet.to_excel(writer, sheet_name=month_sheet_label(m), index=False, startrow=2)
    out.seek(0)
    return out.getvalue()

def kpi_block(filtered_month):
    total_sales = filtered_month["실매출액"].sum()
    total_cnt = filtered_month["전표건수"].sum()
    avg = pct(total_sales, total_cnt)
    supply = filtered_month["공급가액"].sum()
    mom_base = filtered_month["전월실매출액"].fillna(0).sum() if "전월실매출액" in filtered_month.columns else 0
    mom = pct(total_sales-mom_base, mom_base) if mom_base else 0
    cols = st.columns(5)
    metrics = [
        ("실매출액", f"{total_sales:,.0f}원"),
        ("전표건수", f"{total_cnt:,.0f}건"),
        ("객단가", f"{avg:,.0f}원"),
        ("공급가액", f"{supply:,.0f}원"),
        ("전월대비", f"{mom*100:,.1f}%"),
    ]
    for c, (label, val) in zip(cols, metrics):
        c.metric(label, val)

st.title("굿윌 매출 대시보드")
st.caption("대시보드 조회 + 월매출분석자료 생성 + 굿윌2026매출자료 생성")

with st.sidebar:
    st.header("파일 업로드")
    daily_files = st.file_uploader("매출현황 파일", type=["xlsx","xls"], accept_multiple_files=True)
    product_files = st.file_uploader("상품별 매출현황 파일", type=["xlsx","xls"], accept_multiple_files=True, help="파일명에 2026-04 또는 202604 형식의 월 표시 필요")

if not daily_files:
    st.info("먼저 매출현황 파일을 업로드해 주세요.")
    st.stop()

daily_df, product_df = load_all_data(daily_files, product_files)
if daily_df.empty:
    st.warning("매출현황 파일에서 읽을 수 있는 데이터가 없습니다.")
    st.stop()

month_store = build_month_store_summary(daily_df)
all_months = sorted(month_store["기준월"].dropna().unique().tolist())
selected_month = st.sidebar.selectbox("기준월", all_months, index=len(all_months)-1)
all_stores = sorted(month_store["지점명"].dropna().unique().tolist())
selected_stores = st.sidebar.multiselect("지점", all_stores, default=all_stores)

filtered_month = month_store[(month_store["기준월"] == selected_month) & (month_store["지점명"].isin(selected_stores))].copy()
daily_month = daily_df[(daily_df["기준월"] == selected_month) & (daily_df["지점명"].isin(selected_stores))].copy()
product_month = product_df[(product_df["기준월"] == selected_month) & (product_df["지점명"].isin(selected_stores))].copy() if not product_df.empty else pd.DataFrame()

kpi_block(filtered_month)

tab1, tab2, tab3, tab4 = st.tabs(["임원용 대시보드", "월매출 분석자료", "굿윌2026매출자료", "엑셀 생성"])

with tab1:
    left, right = st.columns([1.4, 1])
    with left:
        st.subheader(f"{month_label(selected_month)} 월별 추이")
        trend = month_store[month_store["지점명"].isin(selected_stores)].groupby("기준월", as_index=False)[["실매출액","전표건수","공급가액"]].sum()
        st.line_chart(trend.set_index("기준월")["실매출액"])
        st.subheader("일별 매출")
        daily_trend = daily_month.groupby("영업일자", as_index=False)["실매출액"].sum().sort_values("영업일자")
        st.area_chart(daily_trend.set_index("영업일자")["실매출액"])
    with right:
        st.subheader("지점별 실매출액")
        comp = filtered_month[["지점명","실매출액"]].sort_values("실매출액", ascending=False).set_index("지점명")
        st.bar_chart(comp)
        st.subheader("지점별 요약")
        view = filtered_month.copy()
        view["전월대비증감률"] = (view["전월대비증감률"] * 100).round(1)
        st.dataframe(view[["지점명","실매출액","전표건수","객단가","전월대비증감률"]], use_container_width=True, hide_index=True)
    if not product_month.empty:
        a, b = st.columns(2)
        with a:
            class_df = build_classification_report(product_month)
            class_chart = class_df.groupby("분류그룹", as_index=False)["실매출액"].sum().sort_values("실매출액", ascending=False).set_index("분류그룹")
            st.subheader("분류별 매출")
            st.bar_chart(class_chart)
        with b:
            donor_df = build_donor_report(product_month)
            donor_chart = donor_df.groupby("기증처", as_index=False)["실매출액"].sum().sort_values("실매출액", ascending=False).head(10).set_index("기증처")
            st.subheader("주요 기증처별 매출")
            st.bar_chart(donor_chart)

with tab2:
    st.subheader("월매출 분석자료 미리보기")
    if product_month.empty:
        st.info("상품별 파일을 업로드하면 분류별/주요 기증처별 분석표가 생성됩니다.")
    else:
        class_df = build_classification_report(product_month)
        donor_df = build_donor_report(product_month)
        st.markdown("#### 지점별 분류별 판매수량 / 실매출액 / 점유율")
        st.dataframe(class_df, use_container_width=True, hide_index=True)
        st.markdown("#### 주요 기증처별 판매수량 / 실매출액 / 피스단가")
        st.dataframe(donor_df, use_container_width=True, hide_index=True)

with tab3:
    st.subheader("굿윌2026매출자료 미리보기")
    summary_preview = month_store[month_store["지점명"].isin(selected_stores)].pivot_table(index="지점명", columns="기준월", values="실매출액", aggfunc="sum", fill_value=0).reset_index()
    st.markdown("#### 월합계")
    st.dataframe(summary_preview, use_container_width=True, hide_index=True)
    st.markdown("#### 선택월 일별 매출 / 영수건수 / 객단가")
    rows = []
    for store in selected_stores:
        sub = daily_month[daily_month["지점명"] == store]
        if sub.empty:
            continue
        sales = {"지점명": store, "구분": "매출"}
        cnt = {"지점명": store, "구분": "영수건수"}
        avg = {"지점명": store, "구분": "객단가"}
        for d in range(1, 32):
            day_sub = sub[sub["일"] == d]
            sales[str(d)] = float(day_sub["실매출액"].sum()) if not day_sub.empty else 0
            cnt[str(d)] = float(day_sub["전표건수"].sum()) if not day_sub.empty else 0
            avg[str(d)] = pct(sales[str(d)], cnt[str(d)]) if cnt[str(d)] else 0
        rows.extend([sales, cnt, avg])
    preview = pd.DataFrame(rows)
    st.dataframe(preview, use_container_width=True, hide_index=True)

with tab4:
    st.subheader("엑셀 생성")
    c1, c2 = st.columns(2)
    with c1:
        month_analysis_bytes = make_month_analysis_workbook_bytes(daily_df, product_df)
        st.download_button(
            "월매출분석자료.xlsx 다운로드",
            data=month_analysis_bytes,
            file_name="월매출분석자료_자동생성.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with c2:
        goodwill_bytes = make_goodwill_sales_workbook_bytes(daily_df)
        st.download_button(
            "굿윌2026매출자료.xlsx 다운로드",
            data=goodwill_bytes,
            file_name="굿윌2026매출자료_자동생성.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.divider()
st.markdown("""
**업로드 규칙**
- 매출현황 파일: 일자 기준 데이터
- 상품별 매출현황 파일: 파일명에 월 포함  
  예) `상품별 매출현황-2026-04.xlsx`, `상품별 매출현황-202604.xlsx`

**현재 자동 생성 항목**
- 월매출분석자료: 지점별 분류별 판매수량 / 실매출액 / 점유율, 주요 기증처별 판매수량 / 실매출액 / 피스단가
- 굿윌2026매출자료: 월합계, 월별 일자별 매출 / 영수건수 / 객단가
""")
