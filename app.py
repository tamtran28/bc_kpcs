import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="BC KPCS – ĐỦ 7 BẢNG (Chuẩn VBA)", layout="wide")
st.title("📊 BÁO CÁO KPCS – ĐẦY ĐỦ BẢNG 01 → 07 (CHUẨN VBA)")

# ======================================================
# HELPER
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
# CORE LOGIC – VBA
# ======================================================
def valid_ton(df, BH, KP, TD, start, end):
    """
    Tồn cuối kỳ theo VBA (có xét Theo dõi riêng)
    """
    return (
        (df[BH] <= end)
        &
        (df[KP].isna() | (df[KP] > end))
        &
        (
            df[TD].isna()
            | ((df[TD] >= start) & (df[TD] <= end))
        )
    )


def cnt_df(dfx, group_cols):
    if dfx.empty:
        return pd.Series(dtype=int)
    return dfx.groupby(group_cols).size()


# ======================================================
# BẢNG 01 – TỔNG HỢP (CHUẨN ẢNH EXCEL)
# ======================================================
def bang_01(df, group_cols, BH, KP, TD, HAN, dates):
    y0 = dates["year_start_date"]
    s  = dates["report_start_date"]
    e  = dates["report_end_date"]

    for c in [BH, KP, TD, HAN]:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    # ===== NĂM =====
    ton_dau_nam = df[(df[BH] < y0) & (df[KP].isna() | (df[KP] >= y0))]
    phat_sinh_nam = df[(df[BH] >= y0) & (df[BH] <= e)]
    khac_phuc_nam = df[(df[KP].notna()) & (df[KP] >= y0) & (df[KP] <= e)]

    # ===== QUÝ =====
    ton_dau_quy = df[(df[BH] < s) & (df[KP].isna() | (df[KP] >= s))]
    phat_sinh_quy = df[(df[BH] >= s) & (df[BH] <= e)]
    khac_phuc_quy = df[(df[KP].notna()) & (df[KP] >= s) & (df[KP] <= e)]

    out = pd.DataFrame({
        "Tồn đầu năm": cnt_df(ton_dau_nam, group_cols),
        "Phát sinh năm": cnt_df(phat_sinh_nam, group_cols),
        "Khắc phục năm": cnt_df(khac_phuc_nam, group_cols),
        "Tồn đầu quý": cnt_df(ton_dau_quy, group_cols),
        "Phát sinh quý": cnt_df(phat_sinh_quy, group_cols),
        "Khắc phục quý": cnt_df(khac_phuc_quy, group_cols),
    }).fillna(0).astype(int)

    # ===== TỒN CUỐI QUÝ (CÔNG THỨC VBA) =====
    out["Tồn cuối quý"] = (
        out["Tồn đầu quý"]
        + out["Phát sinh quý"]
        - out["Khắc phục quý"]
    )

    # ===== TỶ LỆ =====
    denom = out["Tồn đầu năm"] + out["Phát sinh năm"]
    out["Tỷ lệ chưa KP đến cuối quý"] = np.where(
        denom > 0, out["Tồn cuối quý"] / denom, 0
    )

    # ===== QUÁ HẠN (CHỈ TRÊN TỒN CUỐI QUÝ) =====
    ton_cuoi_df = df[
        (df[BH] <= e) &
        (df[KP].isna() | (df[KP] > e))
    ].copy()

    ton_cuoi_df[HAN] = pd.to_datetime(ton_cuoi_df[HAN], errors="coerce")

    qua_han = ton_cuoi_df[ton_cuoi_df[HAN] < e]
    qua_han_1n = ton_cuoi_df[
        ton_cuoi_df[HAN] < (e - pd.DateOffset(days=365))
    ]

    out["Quá hạn khắc phục"] = cnt_df(qua_han, group_cols)
    out["Trong đó quá hạn trên 1 năm"] = cnt_df(qua_han_1n, group_cols)

    return out.fillna(0)


# ======================================================
# BẢNG 02 – LOẠI DÒNG = 0
# ======================================================
def bang_02(b01):
    return b01.loc[(b01 != 0).any(axis=1)]


# ======================================================
# BẢNG 03 – TOP ĐƠN VỊ TỒN CUỐI QUÝ
# ======================================================
def bang_03(df, DONVI, BH, KP, TD, dates, top_n=10):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)]
    return (
        ton.groupby(DONVI)
        .size()
        .sort_values(ascending=False)
        .head(top_n)
        .to_frame("Tồn cuối quý")
    )


# ======================================================
# BẢNG 04 – PHÂN NHÓM QUÁ HẠN
# ======================================================
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


# ======================================================
# BẢNG 05 – TOP ĐƠN VỊ QUÁ HẠN
# ======================================================
def bang_05(df, DONVI, BH, KP, TD, HAN, dates, top_n=10):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)].copy()
    ton[HAN] = pd.to_datetime(ton[HAN], errors="coerce")
    ton = ton[ton[HAN] < e]

    return (
        ton.groupby(DONVI)
        .size()
        .sort_values(ascending=False)
        .head(top_n)
        .to_frame("Quá hạn")
    )


# ======================================================
# BẢNG 06 – THEO KHỐI / KHU VỰC
# ======================================================
def bang_06(df, KHOI, KV, BH, KP, TD, dates):
    s, e = dates["report_start_date"], dates["report_end_date"]
    ton = df[valid_ton(df, BH, KP, TD, s, e)]
    return ton.groupby([KHOI, KV]).size().to_frame("Tồn")


# ======================================================
# BẢNG 07 – CHI TIẾT ĐƠN VỊ
# ======================================================
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
    end = st.date_input("Đến ngày", datetime.now())
    file = st.file_uploader("📂 File Excel KPCS", type=["xlsx", "xls"])

if file:
    df = load_excel(file)

    BH = find_column(df, ["Ngày, tháng, năm ban hành (mm/dd/yyyy)", "Ngày ban hành"])
    KP = find_column(df, ["NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)", "Ngày hoàn tất"])
    TD = find_column(df, ["NGÀY CHUYỂN THEO DÕI RIÊNG (mm/dd/yyyy)"])
    HAN = find_column(df, ["Thời hạn hoàn thành (mm/dd/yyyy)", "Hạn KPCS"])
    DONVI = find_column(df, ["Đơn vị thực hiện KPCS trong quý"])
    KHOI = find_column(df, ["SUM (THEO Khối, KV, ĐVKD, Hội sở, Ban Dự Án QLTS)"])
    KV = find_column(df, ["Khối, Khu vực, AMC"])

    must_have({
        "Ngày ban hành": BH,
        "Ngày hoàn tất": KP,
        "Theo dõi riêng": TD,
        "Hạn": HAN,
        "Đơn vị": DONVI,
    })

    df["NHÓM"] = "TỔNG"

    dates = {
        "year_start_date": pd.to_datetime(f"{end.year}-01-01"),
        "report_start_date": pd.to_datetime(start),
        "report_end_date": pd.to_datetime(end),
    }

    b01 = bang_01(df, ["NHÓM"], BH, KP, TD, HAN, dates)
    b02 = bang_02(b01)
    b03 = bang_03(df, DONVI, BH, KP, TD, dates)
    b04 = bang_04(df, BH, KP, TD, HAN, dates)
    b05 = bang_05(df, DONVI, BH, KP, TD, HAN, dates)
    b06 = bang_06(df, KHOI, KV, BH, KP, TD, dates) if KHOI and KV else None
    b07 = bang_07(df, KHOI, KV, DONVI, BH, KP, TD, dates) if KHOI and KV else None

    for name, table in {
        "BẢNG 01": b01,
        "BẢNG 02": b02,
        "BẢNG 03": b03,
        "BẢNG 04": b04,
        "BẢNG 05": b05,
        "BẢNG 06": b06,
        "BẢNG 07": b07,
    }.items():
        if table is not None:
            st.subheader(name)
            st.dataframe(table, use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        b01.to_excel(writer, sheet_name="BANG_01")
        b02.to_excel(writer, sheet_name="BANG_02")
        b03.to_excel(writer, sheet_name="BANG_03")
        b04.to_excel(writer, sheet_name="BANG_04")
        b05.to_excel(writer, sheet_name="BANG_05")
        if b06 is not None:
            b06.to_excel(writer, sheet_name="BANG_06")
        if b07 is not None:
            b07.to_excel(writer, sheet_name="BANG_07")

    st.download_button(
        "📥 Tải Excel ĐẦY ĐỦ 7 BẢNG",
        data=output.getvalue(),
        file_name="BC_KPCS_7_BANG_PYTHON.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("⬅️ Upload file Excel để bắt đầu")
