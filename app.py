import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="BC KPCS – FULL 01→07 (Chuẩn VBA)", layout="wide")
st.title("📊 BÁO CÁO KPCS – ĐẦY ĐỦ 7 BẢNG (CHUẨN VBA)")

# ======================================================
# HELPERS
# ======================================================
def find_column(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None

def must_have(mapping):
    for name, col in mapping.items():
        if col is None:
            st.error(f"❌ Thiếu cột bắt buộc: {name}")
            st.stop()

@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)
    for c in df.columns:
        if "ngày" in c.lower():
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
    return df

# ======================================================
# CORE LOGIC
# ======================================================
def valid_ton(df, BH, KP, TD, start, end):
    # Tồn cuối kỳ có xét Theo dõi riêng
    return (
        (df[BH] <= end)
        & (df[KP].isna() | (df[KP] > end))
        & (
            df[TD].isna()
            | ((df[TD] >= start) & (df[TD] <= end))
        )
    )

def cnt_df(dfx, group_cols):
    if dfx.empty:
        return pd.Series(dtype=int)
    return dfx.groupby(group_cols).size()

# ======================================================
# BẢNG 01 – CHUẨN ẢNH EXCEL
# ======================================================
def bang_01(df, group_col, BH, KP, HAN, dates):
    y0 = dates["year_start_date"]
    s  = dates["report_start_date"]
    e  = dates["report_end_date"]

    for c in [BH, KP, HAN]:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    def cnt(dfx):
        if dfx.empty:
            return pd.Series(dtype=int)
        return dfx.groupby(group_col).size()

    # NĂM
    ton_dau_nam = df[(df[BH] < y0) & (df[KP].isna() | (df[KP] >= y0))]
    phat_sinh_nam = df[(df[BH] >= y0) & (df[BH] <= e)]
    khac_phuc_nam = df[(df[KP].notna()) & (df[KP] >= y0) & (df[KP] <= e)]

    # QUÝ
    ton_dau_quy = df[(df[BH] < s) & (df[KP].isna() | (df[KP] >= s))]
    phat_sinh_quy = df[(df[BH] >= s) & (df[BH] <= e)]
    khac_phuc_quy = df[(df[KP].notna()) & (df[KP] >= s) & (df[KP] <= e)]

    out = pd.DataFrame({
        "Tồn đầu năm": cnt(ton_dau_nam),
        "Phát sinh": cnt(phat_sinh_nam),
        "Khắc phục": cnt(khac_phuc_nam),
        "Tồn đầu quý": cnt(ton_dau_quy),
        "Phát sinh quý": cnt(phat_sinh_quy),
        "Khắc phục quý": cnt(khac_phuc_quy),
    }).fillna(0).astype(int)

    # F + G − H
    out["Tồn cuối quý"] = out["Tồn đầu quý"] + out["Phát sinh quý"] - out["Khắc phục quý"]

    denom = out["Tồn đầu năm"] + out["Phát sinh"]
    out["Tỷ lệ chưa khắc phục đến cuối Quý"] = np.where(denom > 0, out["Tồn cuối quý"] / denom, 0)

    # QUÁ HẠN (chỉ trên tồn cuối quý)
    ton_cuoi_df = df[(df[BH] <= e) & (df[KP].isna() | (df[KP] > e))].copy()
    ton_cuoi_df[HAN] = pd.to_datetime(ton_cuoi_df[HAN], errors="coerce")

    qua_han = ton_cuoi_df[ton_cuoi_df[HAN] < e]
    qua_han_1n = ton_cuoi_df[ton_cuoi_df[HAN] < (e - pd.DateOffset(days=365))]

    out["Quá hạn khắc phục"] = cnt(qua_han)
    out["Trong đó quá hạn trên 1 năm"] = cnt(qua_han_1n)

    return out.fillna(0)

def add_total_row(df):
    total = df.sum(numeric_only=True)
    total.name = "TỔNG"
    return pd.concat([df, total.to_frame().T])

# ======================================================
# BẢNG 02 → 07
# ======================================================
def bang_02(b01):
    return b01.loc[(b01 != 0).any(axis=1)]

def bang_03(df, DONVI, BH, KP, TD, dates, top_n=10):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)]
    return ton.groupby(DONVI).size().sort_values(ascending=False).head(top_n).to_frame("Tồn cuối quý")

def bang_04(df, BH, KP, TD, HAN, dates):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)].copy()
    ton[HAN] = pd.to_datetime(ton[HAN], errors="coerce")
    ton = ton[ton[HAN].notna()]
    ton["Số ngày quá hạn"] = (e - ton[HAN]).dt.days
    bins = [-1, 90, 180, 270, 365, 10**9]
    labels = ["<3 tháng", "3–6", "6–9", "9–12", ">1 năm"]
    ton["Nhóm"] = pd.cut(ton["Số ngày quá hạn"], bins=bins, labels=labels)
    return ton.groupby("Nhóm").size().to_frame("Số lượng")

def bang_05(df, DONVI, BH, KP, TD, HAN, dates, top_n=10):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)].copy()
    ton[HAN] = pd.to_datetime(ton[HAN], errors="coerce")
    ton = ton[ton[HAN] < e]
    return ton.groupby(DONVI).size().sort_values(ascending=False).head(top_n).to_frame("Quá hạn")

def bang_06(df, KHOI, KV, BH, KP, TD, dates):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)]
    return ton.groupby([KHOI, KV]).size().to_frame("Tồn")

def bang_07(df, KHOI, KV, DONVI, BH, KP, TD, dates):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)]
    return ton.groupby([KHOI, KV, DONVI]).size().to_frame("Tồn")

# ======================================================
# UI
# ======================================================
with st.sidebar:
    st.header("⚙️ THAM SỐ")
    start = st.date_input("Từ ngày", datetime(datetime.now().year, 1, 1))
    end   = st.date_input("Đến ngày", datetime.now())
    file  = st.file_uploader("📂 File Excel KPCS", type=["xlsx", "xls"])

if file:
    df = load_excel(file)

    BH = find_column(df, ["Ngày, tháng, năm ban hành (mm/dd/yyyy)", "Ngày ban hành"])
    KP = find_column(df, ["NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)", "Ngày hoàn tất"])
    TD = find_column(df, ["NGÀY CHUYỂN THEO DÕI RIÊNG (mm/dd/yyyy)"])
    HAN = find_column(df, ["Thời hạn hoàn thành (mm/dd/yyyy)", "Hạn KPCS"])
    DONVI = find_column(df, ["Đơn vị thực hiện KPCS trong quý", "Đơn vị"])
    KHOI = find_column(df, ["SUM (THEO Khối, KV, ĐVKD, Hội sở, Ban Dự Án QLTS)", "Khối"])
    KV = find_column(df, ["Khối, Khu vực, AMC", "Khu vực"])

    must_have({
        "Ngày ban hành": BH,
        "Ngày hoàn tất": KP,
        "Theo dõi riêng": TD,
        "Hạn": HAN,
        "Đơn vị": DONVI,
        "Khối/KV": KV
    })

    # Phân nhóm BẢNG 01 đúng ảnh Excel
    NHOM_COL = "NHÓM_BANG_01"
    df[NHOM_COL] = np.where(df[KV].str.contains("Hội sở", na=False), "Hội sở", "ĐVKD, AMC")

    dates = {
        "year_start_date": pd.to_datetime(f"{end.year}-01-01"),
        "report_start_date": pd.to_datetime(start),
        "report_end_date": pd.to_datetime(end),
    }

    # TÍNH BẢNG
    b01_raw = bang_01(df, NHOM_COL, BH, KP, HAN, dates)
    b01 = add_total_row(b01_raw)
    b02 = bang_02(b01)
    b03 = bang_03(df, DONVI, BH, KP, TD, dates)
    b04 = bang_04(df, BH, KP, TD, HAN, dates)
    b05 = bang_05(df, DONVI, BH, KP, TD, HAN, dates)
    b06 = bang_06(df, KHOI, KV, BH, KP, TD, dates) if KHOI and KV else None
    b07 = bang_07(df, KHOI, KV, DONVI, BH, KP, TD, dates) if KHOI and KV else None

    # HIỂN THỊ
    tables = {
        "BẢNG 01": b01,
        "BẢNG 02": b02,
        "BẢNG 03": b03,
        "BẢNG 04": b04,
        "BẢNG 05": b05,
        "BẢNG 06": b06,
        "BẢNG 07": b07,
    }
    for name, tb in tables.items():
        if tb is not None:
            st.subheader(name)
            if name == "BẢNG 01":
                st.dataframe(tb.style.format({"Tỷ lệ chưa khắc phục đến cuối Quý": "{:.2%}"}), use_container_width=True)
            else:
                st.dataframe(tb, use_container_width=True)

    # EXPORT
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        b01.to_excel(writer, sheet_name="BANG_01")
        b02.to_excel(writer, sheet_name="BANG_02")
        b03.to_excel(writer, sheet_name="BANG_03")
        b04.to_excel(writer, sheet_name="BANG_04")
        b05.to_excel(writer, sheet_name="BANG_05")
        if b06 is not None: b06.to_excel(writer, sheet_name="BANG_06")
        if b07 is not None: b07.to_excel(writer, sheet_name="BANG_07")

    st.download_button(
        "📥 Tải Excel ĐẦY ĐỦ 7 BẢNG",
        data=output.getvalue(),
        file_name="BC_KPCS_7_BANG_PYTHON.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("⬅️ Upload file Excel để bắt đầu")
