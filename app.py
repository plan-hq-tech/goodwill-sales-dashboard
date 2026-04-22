
import io
import re
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="굿윌 매출 리포트", page_icon="📈", layout="wide")

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
CATEGORY_ORDER = ["의류","잡화","생활","식품","건강/미용","문화","원가상품","기타"]

# -----------------------------
# Utility
# -----------------------------
def clean_number(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.number)):
        return float(x)
    s = str(x).strip().replace(",", "").replace("%", "")
    s = s.replace("합계:", "").replace("평균:", "").replace("카운트:", "")
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

def fmt_won(x): return f"{x:,.0f}원"
def fmt_pct(x): return f"{x*100:,.1f}%"

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

# -----------------------------
# Parsing
# -----------------------------
def parse_daily_sales(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0)
    raw = raw.rename(columns=lambda x: str(x).strip())
    store_col = raw.columns[0]
    current_store = None
    rows = []

    for _, row in raw.iterrows():
        marker = row.get(store_col)
        if isinstance(marker, str) and "매장:" in marker:
            current_store = re.sub(r".*매장:\s*", "", marker).split("[")[0].strip().replace("밀알","")
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
    return "기타"

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
        }
        for col in PRODUCT_NUM_COLS:
            out[col] = clean_number(row.get(col))
        rows.append(out)

    df = pd.DataFrame(rows)
    if not df.empty:
        df["대분류"] = df["상품분류"].astype(str).str.split("》").str[0].str.strip()
        df["기증처"] = df.apply(infer_donor, axis=1)
    return df

@st.cache_data(show_spinner=False)
def load_all_data(daily_files, product_files):
    daily = pd.concat([parse_daily_sales(f) for f in daily_files], ignore_index=True) if daily_files else pd.DataFrame()
    product = pd.concat([parse_product_sales(f) for f in product_files], ignore_index=True) if product_files else pd.DataFrame()
    return daily, product

# -----------------------------
# Aggregations
# -----------------------------
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
    if "의류" in x: return "의류"
    if "잡화" in x: return "잡화"
    if "생활용품" in x or "생활" in x: return "생활"
    if "식품" in x: return "식품"
    if "건강/미용" in x or "건강" in x or "미용" in x: return "건강/미용"
    if "문화" in x or "도서" in x: return "문화"
    if "원가상품" in x or "매입" in x: return "원가상품"
    return "기타"

def build_classification_report(product_month):
    if product_month.empty:
        return pd.DataFrame()
    df = product_month.copy()
    df["분류그룹"] = df["대분류"].apply(normalize_category_group)
    grp = df.groupby(["지점명","분류그룹"], as_index=False).agg({"판매수량":"sum","실매출액":"sum"})
    totals = grp.groupby("지점명", as_index=False)["실매출액"].sum().rename(columns={"실매출액":"지점합계"})
    grp = grp.merge(totals, on="지점명", how="left")
    grp["점유율"] = grp.apply(lambda r: pct(r["실매출액"], r["지점합계"]), axis=1)
    grp["정렬"] = grp["분류그룹"].apply(lambda x: CATEGORY_ORDER.index(x) if x in CATEGORY_ORDER else 999)
    return grp.sort_values(["지점명","정렬"]).drop(columns=["정렬","지점합계"])

def build_donor_report(product_month):
    if product_month.empty:
        return pd.DataFrame()
    grp = product_month.groupby(["지점명","기증처"], as_index=False).agg({"판매수량":"sum","실매출액":"sum"})
    grp["피스단가"] = grp.apply(lambda r: pct(r["실매출액"], r["판매수량"]) if r["판매수량"] else 0, axis=1)

    base = grp[grp["기증처"].isin(MAJOR_DONORS)].copy()
    extra = (
        grp[~grp["기증처"].isin(MAJOR_DONORS)]
        .groupby("기증처", as_index=False)["실매출액"].sum()
        .sort_values("실매출액", ascending=False)
        .head(5)
    )
    extra_names = extra["기증처"].tolist()
    extra_df = grp[grp["기증처"].isin(extra_names)].copy()

    out = pd.concat([base, extra_df], ignore_index=True)
    total_order = out.groupby("기증처", as_index=False)["실매출액"].sum().sort_values("실매출액", ascending=False)["기증처"].tolist()
    out["기증처"] = pd.Categorical(out["기증처"], categories=total_order, ordered=True)
    return out.sort_values(["기증처","지점명"])

# -----------------------------
# Excel styling helpers
# -----------------------------
TITLE_FILL = PatternFill("solid", fgColor="1F1F1F")
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
HEADER_FILL = PatternFill("solid", fgColor="EDEDED")
SUBHEADER_FILL = PatternFill("solid", fgColor="F7F7F7")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
WHITE_FONT = Font(color="FFFFFF", bold=True, size=14)
BOLD_FONT = Font(bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")

def auto_fit(ws, min_width=10, max_width=22):
    for col_cells in ws.columns:
        length = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            val = "" if cell.value is None else str(cell.value)
            length = max(length, len(val))
        ws.column_dimensions[col_letter].width = max(min(length + 2, max_width), min_width)

def style_title(ws, row, end_col, title):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    c = ws.cell(row=row, column=1, value=title)
    c.fill = TITLE_FILL
    c.font = WHITE_FONT
    c.alignment = CENTER
    ws.row_dimensions[row].height = 24

def style_section(ws, row, end_col, title):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_col)
    c = ws.cell(row=row, column=1, value=title)
    c.fill = SECTION_FILL
    c.font = Font(bold=True, size=11)
    c.alignment = LEFT

def apply_table_style(ws, start_row, end_row, start_col=1, header_rows=1, pct_cols=None, won_cols=None, int_cols=None):
    pct_cols = pct_cols or []
    won_cols = won_cols or []
    int_cols = int_cols or []
    for r in range(start_row, end_row + 1):
        for c in range(start_col, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            if r < start_row + header_rows:
                cell.fill = HEADER_FILL
                cell.font = BOLD_FONT
                cell.alignment = CENTER
            else:
                cell.alignment = CENTER if c > 1 else LEFT
            if c in pct_cols and r >= start_row + header_rows:
                cell.number_format = '0.0%'
            elif c in won_cols and r >= start_row + header_rows:
                cell.number_format = '#,##0"원"'
            elif c in int_cols and r >= start_row + header_rows:
                cell.number_format = '#,##0'
    for c in range(start_col, ws.max_column + 1):
        ws.cell(start_row, c).fill = HEADER_FILL
        ws.cell(start_row, c).font = BOLD_FONT

def add_total_row(df, key_col, numeric_cols, label="합계"):
    if df.empty:
        return df
    row = {key_col: label}
    for col in df.columns:
        if col == key_col:
            continue
        row[col] = df[col].sum() if col in numeric_cols else ""
    return pd.concat([df, pd.DataFrame([row])], ignore_index=True)

def build_report_matrix(class_df, donor_df):
    class_matrix = pd.DataFrame()
    donor_matrix = pd.DataFrame()

    if not class_df.empty:
        stores = sorted(class_df["지점명"].unique().tolist())
        rows = []
        for cat in CATEGORY_ORDER:
            row = {"구분": cat}
            for store in stores:
                sub = class_df[(class_df["지점명"] == store) & (class_df["분류그룹"] == cat)]
                qty = float(sub["판매수량"].sum()) if not sub.empty else 0
                sales = float(sub["실매출액"].sum()) if not sub.empty else 0
                total_sales = float(class_df[class_df["지점명"] == store]["실매출액"].sum()) if store in class_df["지점명"].values else 0
                share = pct(sales, total_sales)
                row[f"{store}_수량"] = qty
                row[f"{store}_금액"] = sales
                row[f"{store}_점유율"] = share
            rows.append(row)
        class_matrix = pd.DataFrame(rows)

    if not donor_df.empty:
        donors = donor_df["기증처"].astype(str).drop_duplicates().tolist()
        stores = sorted(donor_df["지점명"].unique().tolist())
        rows = []
        for donor in donors:
            row = {"구분": donor}
            for store in stores:
                sub = donor_df[(donor_df["지점명"] == store) & (donor_df["기증처"].astype(str) == donor)]
                qty = float(sub["판매수량"].sum()) if not sub.empty else 0
                sales = float(sub["실매출액"].sum()) if not sub.empty else 0
                piece = pct(sales, qty) if qty else 0
                row[f"{store}_수량"] = qty
                row[f"{store}_금액"] = sales
                row[f"{store}_피스단가"] = piece
            rows.append(row)
        donor_matrix = pd.DataFrame(rows)

    return class_matrix, donor_matrix

# -----------------------------
# Excel generators (styled)
# -----------------------------
def make_designed_month_analysis(product_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        if product_df.empty:
            pd.DataFrame({"안내":["상품별 파일 업로드 후 생성 가능합니다."]}).to_excel(writer, sheet_name="안내", index=False)
        else:
            for month in sorted(product_df["기준월"].unique().tolist()):
                ws_name = f"{int(month.split('-')[1])}월"
                pm = product_df[product_df["기준월"] == month].copy()
                class_df = build_classification_report(pm)
                donor_df = build_donor_report(pm)
                class_matrix, donor_matrix = build_report_matrix(class_df, donor_df)

                class_matrix = add_total_row(
                    class_matrix, "구분",
                    [c for c in class_matrix.columns if c.endswith("_수량") or c.endswith("_금액")]
                ) if not class_matrix.empty else class_matrix

                donor_matrix = add_total_row(
                    donor_matrix, "구분",
                    [c for c in donor_matrix.columns if c.endswith("_수량") or c.endswith("_금액")]
                ) if not donor_matrix.empty else donor_matrix

                row = 1
                if class_matrix.empty and donor_matrix.empty:
                    pd.DataFrame({"안내":["해당 월의 상품 데이터가 없습니다."]}).to_excel(writer, sheet_name=ws_name, index=False)
                    continue

                # write class
                if not class_matrix.empty:
                    class_matrix.to_excel(writer, sheet_name=ws_name, index=False, startrow=row+2)
                    ws = writer.book[ws_name]
                    style_title(ws, row, len(class_matrix.columns), f"{month_label(month)} 월매출 분석자료")
                    style_section(ws, row+2, len(class_matrix.columns), "분류별 매출 현황")
                    start = row+3
                    end = start + len(class_matrix)
                    pct_cols = [i+1 for i, col in enumerate(class_matrix.columns) if col.endswith("_점유율")]
                    won_cols = [i+1 for i, col in enumerate(class_matrix.columns) if col.endswith("_금액")]
                    int_cols = [i+1 for i, col in enumerate(class_matrix.columns) if col.endswith("_수량")]
                    apply_table_style(ws, start, end, pct_cols=pct_cols, won_cols=won_cols, int_cols=int_cols)
                    # total row highlight
                    for c in range(1, ws.max_column+1):
                        ws.cell(end, c).fill = TOTAL_FILL
                        ws.cell(end, c).font = BOLD_FONT
                    row = end + 3

                # write donor
                if not donor_matrix.empty:
                    donor_matrix.to_excel(writer, sheet_name=ws_name, index=False, startrow=row+1)
                    ws = writer.book[ws_name]
                    style_section(ws, row, len(donor_matrix.columns), "주요 기증처별 매출 현황")
                    start = row+1
                    end = start + len(donor_matrix)
                    won_cols = [i+1 for i, col in enumerate(donor_matrix.columns) if col.endswith("_금액") or col.endswith("_피스단가")]
                    int_cols = [i+1 for i, col in enumerate(donor_matrix.columns) if col.endswith("_수량")]
                    apply_table_style(ws, start, end, won_cols=won_cols, int_cols=int_cols)
                    for c in range(1, ws.max_column+1):
                        ws.cell(end, c).fill = TOTAL_FILL
                        ws.cell(end, c).font = BOLD_FONT

                    auto_fit(ws)
                    ws.freeze_panes = "B4"
    out.seek(0)
    return out.getvalue()

def make_designed_goodwill_sales(daily_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        if daily_df.empty:
            pd.DataFrame({"안내":["매출현황 파일 업로드 후 생성 가능합니다."]}).to_excel(writer, sheet_name="안내", index=False)
        else:
            month_store = build_month_store_summary(daily_df)
            summary = month_store.pivot_table(index="지점명", columns="기준월", values="실매출액", aggfunc="sum", fill_value=0).reset_index()
            summary["합계"] = summary.drop(columns=["지점명"]).sum(axis=1)

            summary.to_excel(writer, sheet_name="월합계", index=False, startrow=2)
            ws = writer.book["월합계"]
            style_title(ws, 1, len(summary.columns), "굿윌 매출자료 월합계")
            apply_table_style(
                ws, 3, 3 + len(summary),
                won_cols=list(range(2, len(summary.columns)+1))
            )
            total_row_idx = 3 + len(summary)
            for c in range(1, ws.max_column+1):
                ws.cell(total_row_idx, c).fill = TOTAL_FILL
                ws.cell(total_row_idx, c).font = BOLD_FONT
            auto_fit(ws)
            ws.freeze_panes = "B4"

            for month in sorted(daily_df["기준월"].unique().tolist()):
                ws_name = f"{int(month.split('-')[1])}월"
                dm = daily_df[daily_df["기준월"] == month].copy()
                detail = dm.groupby(["지점명","영업일자"], as_index=False).agg({"실매출액":"sum","전표건수":"sum","공급가액":"sum"})
                detail["객단가"] = detail.apply(lambda r: pct(r["실매출액"], r["전표건수"]) if r["전표건수"] else 0, axis=1)
                detail["영업일자"] = detail["영업일자"].dt.strftime("%Y-%m-%d")
                detail = detail[["지점명","영업일자","실매출액","전표건수","객단가","공급가액"]]
                detail.to_excel(writer, sheet_name=ws_name, index=False, startrow=2)

                ws = writer.book[ws_name]
                style_title(ws, 1, len(detail.columns), f"{month_label(month)} 운영자료")
                apply_table_style(
                    ws, 3, 3 + len(detail),
                    won_cols=[3,5,6],
                    int_cols=[4]
                )
                auto_fit(ws)
                ws.freeze_panes = "A4"
    out.seek(0)
    return out.getvalue()

# -----------------------------
# UI
# -----------------------------
st.title("굿윌 매출 리포트")
st.caption("대시보드 조회 + 보고서형 엑셀 다운로드")

with st.sidebar:
    st.header("파일 업로드")
    daily_files = st.file_uploader("매출현황 파일", type=["xlsx","xls"], accept_multiple_files=True)
    product_files = st.file_uploader("상품별 매출현황 파일", type=["xlsx","xls"], accept_multiple_files=True, help="파일명에 2026-04 또는 202604 형식의 월 표기 필요")

if not daily_files:
    st.info("먼저 매출현황 파일을 업로드해 주세요.")
    st.stop()

daily_df, product_df = load_all_data(daily_files, product_files)
if daily_df.empty:
    st.warning("매출현황 파일을 읽을 수 없습니다.")
    st.stop()

month_store = build_month_store_summary(daily_df)
months = sorted(month_store["기준월"].unique().tolist())
selected_month = st.sidebar.selectbox("기준월", months, index=len(months)-1)
stores = sorted(month_store["지점명"].unique().tolist())
selected_stores = st.sidebar.multiselect("지점 선택", stores, default=stores)

fm = month_store[(month_store["기준월"] == selected_month) & (month_store["지점명"].isin(selected_stores))].copy()
dm = daily_df[(daily_df["기준월"] == selected_month) & (daily_df["지점명"].isin(selected_stores))].copy()
pm = product_df[(product_df["기준월"] == selected_month) & (product_df["지점명"].isin(selected_stores))].copy() if not product_df.empty else pd.DataFrame()

# simple dashboard preview
total_sales = fm["실매출액"].sum()
total_cnt = fm["전표건수"].sum()
avg_ticket = pct(total_sales, total_cnt)
prev_sales = fm["전월실매출액"].fillna(0).sum() if "전월실매출액" in fm.columns else 0
mom = pct(total_sales - prev_sales, prev_sales) if prev_sales else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("실매출액", fmt_won(total_sales))
c2.metric("전표건수", f"{total_cnt:,.0f}건")
c3.metric("객단가", fmt_won(avg_ticket))
c4.metric("전월대비", fmt_pct(mom))

tab1, tab2, tab3 = st.tabs(["미리보기", "월매출 분석자료", "엑셀 다운로드"])

with tab1:
    st.subheader("지점별 실매출 비교")
    st.bar_chart(fm[["지점명","실매출액"]].sort_values("실매출액", ascending=False).set_index("지점명"))
    st.subheader("선택월 요약")
    st.dataframe(fm[["지점명","실매출액","전표건수","객단가","전월대비증감률"]], use_container_width=True, hide_index=True)

with tab2:
    if pm.empty:
        st.info("상품별 파일을 업로드하면 분류별/기증처별 분석이 생성됩니다.")
    else:
        class_df = build_classification_report(pm)
        donor_df = build_donor_report(pm)
        st.markdown("#### 분류별 현황")
        st.dataframe(class_df, use_container_width=True, hide_index=True)
        st.markdown("#### 주요 기증처별 현황")
        st.dataframe(donor_df, use_container_width=True, hide_index=True)

with tab3:
    st.subheader("보고서형 엑셀 다운로드")
    month_bytes = make_designed_month_analysis(product_df)
    sales_bytes = make_designed_goodwill_sales(daily_df)

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "디자인 적용 월매출분석자료.xlsx",
            data=month_bytes,
            file_name="디자인적용_월매출분석자료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("제목/섹션/헤더색/테두리/합계강조/숫자포맷 적용")
    with d2:
        st.download_button(
            "디자인 적용 굿윌매출자료.xlsx",
            data=sales_bytes,
            file_name="디자인적용_굿윌매출자료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("월합계/월별운영자료에 동일한 디자인 적용")
