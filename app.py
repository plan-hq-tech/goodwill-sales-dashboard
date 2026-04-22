
import io
import os
import re
from pathlib import Path
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="굿윌 월매출 대시보드",
    page_icon="📊",
    layout="wide",
)

# ---------------------------
# 유틸
# ---------------------------

MONTH_PATTERNS = [
    re.compile(r"(20\d{2})[-_]?([01]\d)"),
    re.compile(r"(20\d{2})년\s*([01]?\d)월"),
]

NUMERIC_COLS_DAILY = [
    "전표건수", "객단가", "공급가액", "실매출액", "현금", "현금영수증", "카드",
    "포인트", "현금카드외", "제휴포인트", "상품권결제"
]

NUMERIC_COLS_PRODUCT = [
    "판매수량", "공급가액", "실매출액", "기본단가", "이익금액",
    "현금", "현금영수증", "카드", "상품권결제", "현금카드외"
]

def extract_month_from_filename(filename: str) -> Optional[str]:
    """파일명에서 YYYY-MM 추출"""
    name = Path(filename).stem
    for pattern in MONTH_PATTERNS:
        m = pattern.search(name)
        if m:
            year = int(m.group(1))
            month = int(m.group(2))
            if 1 <= month <= 12:
                return f"{year:04d}-{month:02d}"
    return None

def clean_number(x):
    if pd.isna(x):
        return 0
    if isinstance(x, (int, float, np.number)):
        return float(x)
    x = str(x).strip().replace(",", "").replace("%", "")
    if x == "":
        return 0
    try:
        return float(x)
    except Exception:
        return 0

def safe_pct(numer, denom):
    if denom in [0, None] or pd.isna(denom):
        return 0.0
    return float(numer) / float(denom)

def month_to_label(month_str: str) -> str:
    y, m = month_str.split("-")
    return f"{y}년 {int(m)}월"

# ---------------------------
# 원본 파서
# ---------------------------

def parse_daily_sales(uploaded_file) -> pd.DataFrame:
    """
    매출현황 파일 파서
    기대 구조:
    - 첫 행 또는 중간에 '매장: XXX점 [3]' 같은 텍스트가 존재
    - 실데이터 행에는 영업일자, 전표건수, 실매출액 등이 있음
    """
    raw = pd.read_excel(uploaded_file, sheet_name=0)
    if raw.empty:
        return pd.DataFrame()

    raw = raw.rename(columns=lambda x: str(x).strip())
    store_col = raw.columns[0]

    current_store = None
    rows = []

    for _, row in raw.iterrows():
        store_candidate = row.get(store_col)
        if isinstance(store_candidate, str) and "매장:" in store_candidate:
            current_store = re.sub(r".*매장:\s*", "", store_candidate).split("[")[0].strip()
            continue

        sale_date = row.get("영업일자")
        if pd.isna(sale_date):
            continue

        try:
            sale_date = pd.to_datetime(sale_date)
        except Exception:
            continue

        out = {
            "지점명": current_store if current_store else "미확인",
            "영업일자": sale_date,
            "기준월": sale_date.strftime("%Y-%m"),
        }

        for col in NUMERIC_COLS_DAILY:
            out[col] = clean_number(row.get(col))

        rows.append(out)

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["연도"] = df["영업일자"].dt.year
    df["월"] = df["영업일자"].dt.month
    return df

def parse_product_sales(uploaded_file) -> pd.DataFrame:
    """
    상품별 매출현황 파일 파서
    - 파일명에서 YYYY-MM 추출
    - 매장명은 Unnamed: 1 컬럼, 상품분류 정보는 '상품분류1:' 헤더 문구가 간헐적으로 등장
    """
    raw = pd.read_excel(uploaded_file, sheet_name=0)
    if raw.empty:
        return pd.DataFrame()

    raw = raw.rename(columns=lambda x: str(x).strip())
    month_str = extract_month_from_filename(uploaded_file.name)
    if not month_str:
        raise ValueError(
            f"파일명에서 월을 읽을 수 없습니다: {uploaded_file.name}\n"
            "예: 상품별 매출현황-2026-04.xlsx 또는 상품별 매출현황-202604.xlsx"
        )

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

        out = {
            "기준월": month_str,
            "지점명": str(store).strip(),
            "상품명": str(product).strip(),
            "바코드": str(row.get("바코드")).strip() if not pd.isna(row.get("바코드")) else "",
            "재고단위": str(row.get("재고 단위")).strip() if not pd.isna(row.get("재고 단위")) else "",
            "상품분류": str(row.get("상품분류")).strip() if not pd.isna(row.get("상품분류")) else (category_hint or "미분류"),
        }

        for col in NUMERIC_COLS_PRODUCT:
            out[col] = clean_number(row.get(col))

        rows.append(out)

    return pd.DataFrame(rows)

# ---------------------------
# 통합 집계
# ---------------------------

@st.cache_data(show_spinner=False)
def load_all_data(daily_files, product_files):
    daily_frames = []
    product_frames = []

    for f in daily_files:
        df = parse_daily_sales(f)
        if not df.empty:
            daily_frames.append(df)

    for f in product_files:
        df = parse_product_sales(f)
        if not df.empty:
            product_frames.append(df)

    daily = pd.concat(daily_frames, ignore_index=True) if daily_frames else pd.DataFrame()
    product = pd.concat(product_frames, ignore_index=True) if product_frames else pd.DataFrame()

    if not daily.empty:
        daily["객단가_계산"] = daily.apply(
            lambda r: safe_pct(r["실매출액"], r["전표건수"]) if r["전표건수"] else 0, axis=1
        )

    return daily, product

def build_month_summary(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame()

    grouped = (
        daily.groupby(["기준월", "지점명"], as_index=False)
        .agg({
            "전표건수": "sum",
            "공급가액": "sum",
            "실매출액": "sum",
            "현금": "sum",
            "현금영수증": "sum",
            "카드": "sum",
            "포인트": "sum",
            "현금카드외": "sum",
            "제휴포인트": "sum",
            "상품권결제": "sum",
        })
    )
    grouped["객단가"] = grouped.apply(
        lambda r: safe_pct(r["실매출액"], r["전표건수"]) if r["전표건수"] else 0, axis=1
    )

    grouped = grouped.sort_values(["기준월", "실매출액"], ascending=[True, False])
    grouped["월내매출순위"] = grouped.groupby("기준월")["실매출액"].rank(method="dense", ascending=False)
    grouped["전월실매출액"] = grouped.groupby("지점명")["실매출액"].shift(1)
    grouped["전월대비증감률"] = grouped.apply(
        lambda r: safe_pct(r["실매출액"] - r["전월실매출액"], r["전월실매출액"]) if r["전월실매출액"] else 0,
        axis=1
    )
    return grouped

def build_product_summary(product: pd.DataFrame) -> pd.DataFrame:
    if product.empty:
        return pd.DataFrame()

    grouped = (
        product.groupby(["기준월", "지점명", "상품분류"], as_index=False)
        .agg({
            "판매수량": "sum",
            "실매출액": "sum",
            "공급가액": "sum",
            "이익금액": "sum",
        })
    )
    grouped["이익률"] = grouped.apply(
        lambda r: safe_pct(r["이익금액"], r["실매출액"]) if r["실매출액"] else 0,
        axis=1
    )
    return grouped

def init_store_master(all_stores: List[str]) -> pd.DataFrame:
    return pd.DataFrame({
        "지점명": sorted(set(all_stores)),
        "사용여부": True,
        "표시순서": list(range(1, len(set(all_stores)) + 1))
    })

def apply_store_master(stores_df: pd.DataFrame, store_master: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    enabled = (
        store_master.loc[store_master["사용여부"] == True, ["지점명", "표시순서"]]
        .drop_duplicates()
    )
    if not stores_df.empty:
        stores_df = stores_df.merge(enabled, on="지점명", how="inner")
        stores_df = stores_df.sort_values(["표시순서", "기준월"])
    return stores_df, enabled

# ---------------------------
# UI
# ---------------------------

st.title("굿윌 월매출 대시보드")
st.caption("매출현황 파일 + 상품별 매출현황 파일을 업로드하면 월별 분석과 지점별 대시보드를 자동 생성합니다.")

with st.sidebar:
    st.header("파일 업로드")
    daily_files = st.file_uploader(
        "매출현황 파일 업로드",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="예: 매출현황-20260422.xlsx",
    )
    product_files = st.file_uploader(
        "상품별 매출현황 파일 업로드",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="예: 상품별 매출현황-2026-04.xlsx 또는 상품별 매출현황-202604.xlsx",
    )

if not daily_files:
    st.info("먼저 매출현황 파일을 업로드해 주세요.")
    st.stop()

try:
    daily_df, product_df = load_all_data(daily_files, product_files)
except Exception as e:
    st.error(f"파일 로딩 중 오류가 발생했습니다: {e}")
    st.stop()

if daily_df.empty:
    st.warning("매출현황 파일에서 읽을 수 있는 데이터가 없습니다.")
    st.stop()

month_summary = build_month_summary(daily_df)
product_summary = build_product_summary(product_df)

all_stores = list(daily_df["지점명"].dropna().unique())
if not product_df.empty:
    all_stores = sorted(set(all_stores) | set(product_df["지점명"].dropna().unique()))

st.sidebar.header("지점 설정")
if "store_master" not in st.session_state:
    st.session_state.store_master = init_store_master(all_stores)
else:
    current = set(st.session_state.store_master["지점명"])
    new = set(all_stores) - current
    if new:
        add_df = pd.DataFrame({"지점명": sorted(new), "사용여부": True, "표시순서": range(len(current)+1, len(current)+len(new)+1)})
        st.session_state.store_master = pd.concat([st.session_state.store_master, add_df], ignore_index=True)

edited_master = st.sidebar.data_editor(
    st.session_state.store_master,
    num_rows="dynamic",
    use_container_width=True,
    key="store_editor",
)
st.session_state.store_master = edited_master.copy()

month_summary, enabled_stores = apply_store_master(month_summary, st.session_state.store_master)
if not product_summary.empty:
    product_summary, _ = apply_store_master(product_summary, st.session_state.store_master)

available_months = sorted(month_summary["기준월"].dropna().unique())
selected_month = st.sidebar.selectbox("기준월 선택", available_months, index=len(available_months)-1 if available_months else 0)
enabled_store_names = enabled_stores.sort_values("표시순서")["지점명"].tolist()
selected_stores = st.sidebar.multiselect("지점 선택", enabled_store_names, default=enabled_store_names)

filtered_month = month_summary[
    (month_summary["기준월"] == selected_month) &
    (month_summary["지점명"].isin(selected_stores))
].copy()

daily_month = daily_df[
    (daily_df["기준월"] == selected_month) &
    (daily_df["지점명"].isin(selected_stores))
].copy()

product_month = pd.DataFrame()
if not product_df.empty:
    product_month = product_df[
        (product_df["기준월"] == selected_month) &
        (product_df["지점명"].isin(selected_stores))
    ].copy()

if filtered_month.empty:
    st.warning("선택한 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# KPI
total_sales = filtered_month["실매출액"].sum()
total_tickets = filtered_month["전표건수"].sum()
avg_ticket = safe_pct(total_sales, total_tickets)
total_supply = filtered_month["공급가액"].sum()
mom = safe_pct(
    filtered_month["실매출액"].sum() - filtered_month["전월실매출액"].fillna(0).sum(),
    filtered_month["전월실매출액"].fillna(0).sum()
)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("실매출액", f"{total_sales:,.0f}원")
c2.metric("전표건수", f"{total_tickets:,.0f}건")
c3.metric("객단가", f"{avg_ticket:,.0f}원")
c4.metric("공급가액", f"{total_supply:,.0f}원")
c5.metric("전월대비", f"{mom*100:,.1f}%")

# 월별 추이
st.subheader(f"{month_to_label(selected_month)} 요약")
trend = (
    month_summary[month_summary["지점명"].isin(selected_stores)]
    .groupby("기준월", as_index=False)[["실매출액", "전표건수", "공급가액"]]
    .sum()
    .sort_values("기준월")
)

tab1, tab2, tab3, tab4 = st.tabs(["대시보드", "지점 분석", "상품 분석", "데이터 다운로드"])

with tab1:
    col_a, col_b = st.columns([1.2, 1])

    with col_a:
        st.markdown("#### 월별 매출 추이")
        st.line_chart(trend.set_index("기준월")["실매출액"])

        st.markdown("#### 일별 매출 추이")
        daily_trend = daily_month.groupby("영업일자", as_index=False)["실매출액"].sum().sort_values("영업일자")
        st.line_chart(daily_trend.set_index("영업일자")["실매출액"])

    with col_b:
        st.markdown("#### 지점별 매출 비교")
        compare_sales = filtered_month[["지점명", "실매출액"]].sort_values("실매출액", ascending=False).set_index("지점명")
        st.bar_chart(compare_sales)

        st.markdown("#### 결제수단 구성")
        payment_cols = ["현금", "현금영수증", "카드", "포인트", "현금카드외", "제휴포인트", "상품권결제"]
        payment_sum = daily_month[payment_cols].sum().reset_index()
        payment_sum.columns = ["결제수단", "금액"]
        payment_sum = payment_sum[payment_sum["금액"] > 0]
        if not payment_sum.empty:
            st.dataframe(payment_sum, use_container_width=True, hide_index=True)

with tab2:
    st.markdown("#### 지점별 월매출 분석")
    view = filtered_month.copy()
    view["전월대비증감률"] = (view["전월대비증감률"] * 100).round(1)
    view["월내매출순위"] = view["월내매출순위"].astype(int)
    st.dataframe(
        view[["지점명", "실매출액", "공급가액", "전표건수", "객단가", "월내매출순위", "전월대비증감률"]],
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("#### 지점별 일자 상세")
    daily_store = (
        daily_month.groupby(["지점명", "영업일자"], as_index=False)[["실매출액", "전표건수", "공급가액"]]
        .sum()
        .sort_values(["지점명", "영업일자"])
    )
    st.dataframe(daily_store, use_container_width=True, hide_index=True)

with tab3:
    if product_month.empty:
        st.info("상품별 파일이 없거나, 선택한 월에 해당하는 상품 데이터가 없습니다.")
    else:
        st.markdown("#### 상품분류별 실매출액")
        cat = (
            product_month.groupby("상품분류", as_index=False)[["실매출액", "판매수량", "이익금액"]]
            .sum()
            .sort_values("실매출액", ascending=False)
        )
        st.bar_chart(cat.set_index("상품분류")["실매출액"])

        st.markdown("#### TOP 20 상품")
        top20 = (
            product_month.groupby(["상품명", "상품분류"], as_index=False)[["판매수량", "실매출액", "공급가액", "이익금액"]]
            .sum()
            .sort_values("실매출액", ascending=False)
            .head(20)
        )
        st.dataframe(top20, use_container_width=True, hide_index=True)

        st.markdown("#### 지점별 상품분류 요약")
        cat_store = (
            product_month.groupby(["지점명", "상품분류"], as_index=False)[["판매수량", "실매출액", "이익금액"]]
            .sum()
            .sort_values(["지점명", "실매출액"], ascending=[True, False])
        )
        st.dataframe(cat_store, use_container_width=True, hide_index=True)

with tab4:
    st.markdown("#### 정제 데이터 다운로드")
    out_buffer = io.BytesIO()
    with pd.ExcelWriter(out_buffer, engine="openpyxl") as writer:
        daily_df.to_excel(writer, index=False, sheet_name="RAW_매출현황_정제")
        month_summary.to_excel(writer, index=False, sheet_name="월별지점집계")
        st.session_state.store_master.to_excel(writer, index=False, sheet_name="지점마스터")
        if not product_df.empty:
            product_df.to_excel(writer, index=False, sheet_name="RAW_상품별_정제")
            product_summary.to_excel(writer, index=False, sheet_name="월별상품집계")

    st.download_button(
        label="정제 데이터 엑셀 다운로드",
        data=out_buffer.getvalue(),
        file_name=f"굿윌_대시보드_정제데이터_{selected_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.divider()
st.markdown(
    """
    **파일명 규칙 권장**
    - 매출현황 파일: `매출현황-20260422.xlsx`
    - 상품별 파일: `상품별 매출현황-2026-04.xlsx` 또는 `상품별 매출현황-202604.xlsx`

    **지점 추가/삭제 방식**
    - 왼쪽 `지점 설정` 표에서 `사용여부` 체크 해제 → 대시보드 제외
    - 새 지점을 추가하면 표에 자동 추가됨
    - `표시순서` 변경으로 정렬 가능
    """
)
