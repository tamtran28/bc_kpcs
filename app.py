import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="BC KPCS – FULL 01→07 (VBA)", layout="wide")
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
    for k, v in mapping.items():
        if v is None:
            st.error(f"❌ Thiếu cột bắt buộc: {k}")
            st.stop()

@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)
    for c in df.columns:
        if "ngày" in c.lower():
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
    return df

# ======================================================
# CORE ENGINE – DÙNG CHUNG (VBA)
# ======================================================
def calc_metrics(df, group_col, BH, KP, HAN, dates):
    y0 = dates["year_start_date"]
    s  = dates["report_start_date"]
    e  = dates["report_end_date"]

    for c in [BH, KP, HAN]:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    def cnt(dfx):
        if dfx.empty:
            return pd.Series(dtype=int)
        return dfx.groupby(group_col).size()

    ton_dau_nam = df[(df[BH] < y0) & (df[KP].isna() | (df[KP] >= y0))]
    phat_sinh_nam = df[(df[BH] >= y0) & (df[BH] <= e)]
    khac_phuc_nam = df[(df[KP].notna()) & (df[KP] >= y0) & (df[KP] <= e)]

    ton_dau_quy = df[(df[BH] < s) & (df[KP].isna() | (df[KP] >= s))]
    phat_sinh_quy = df[(df[BH] >= s) & (df[BH] <= e)]
    khac_phuc_quy = df[(df[KP].notna()) & (df[KP] >= s) & (df[KP] <= e)]

    out = pd.DataFrame({
        "Tồn đầu năm": cnt(ton_dau_nam),
        "Phát sinh năm": cnt(phat_sinh_nam),
        "Khắc phục năm": cnt(khac_phuc_nam),
        "Tồn đầu quý": cnt(ton_dau_quy),
        "Phát sinh quý": cnt(phat_sinh_quy),
        "Khắc phục quý": cnt(khac_phuc_quy),
    }).fillna(0).astype(int)

    out["Tồn cuối quý"] = (
        out["Tồn đầu quý"]
        + out["Phát sinh quý"]
        - out["Khắc phục quý"]
    )

    denom = out["Tồn đầu năm"] + out["Phát sinh năm"]
    out["Tỷ lệ chưa KP đến cuối Quý"] = np.where(
        denom > 0, out["Tồn cuối quý"] / denom, 0
    )

    ton_cuoi = df[(df[BH] <= e) & (df[KP].isna() | (df[KP] > e))].copy()
    ton_cuoi[HAN] = pd.to_datetime(ton_cuoi[HAN], errors="coerce")

    out["Quá hạn khắc phục"] = cnt(ton_cuoi[ton_cuoi[HAN] < e])
    out["Trong đó quá hạn trên 1 năm"] = cnt(
        ton_cuoi[ton_cuoi[HAN] < (e - pd.DateOffset(days=365))]
    )

    return out.fillna(0)

def add_total(df, name="TỔNG CỘNG"):
    t = df.sum(numeric_only=True)
    t.name = name
    return pd.concat([df, t.to_frame().T])

# ======================================================
# BẢNG 01 – TOÀN HÀNG
# ======================================================
def bang_01(df, KV, BH, KP, HAN, dates):
    df["_NHOM01"] = np.where(
        df[KV].str.contains("Hội sở", na=False),
        "Hội sở",
        "ĐVKD, AMC"
    )
    out = calc_metrics(df, "_NHOM01", BH, KP, HAN, dates)
    return add_total(out, "TỔNG")

# ======================================================
# BẢNG 02 – ĐƠN VỊ HỘI SỞ
# ======================================================
def bang_02(df, KHOI, BH, KP, HAN, dates):
    hs = df[df["KV"].str.contains("Hội sở", na=False)]
    out = calc_metrics(hs, KHOI, BH, KP, HAN, dates)
    return add_total(out)

# ======================================================
# BẢNG 03 – TOP ĐƠN VỊ TỒN CUỐI QUÝ
# ======================================================
def bang_03(df, DONVI, BH, KP, HAN, dates, n=10):
    out = calc_metrics(df, DONVI, BH, KP, HAN, dates)
    return out.sort_values("Tồn cuối quý", ascending=False).head(n)

# ======================================================
# BẢNG 04 – DVKD THEO 5 KV + AMC
# ======================================================
def bang_04(df, KV, BH, KP, HAN, dates):
    out = calc_metrics(df, KV, BH, KP, HAN, dates)
    return add_total(out)

# ======================================================
# BẢNG 05 – TOP 10 DVKD QUÁ HẠN
# ======================================================
def bang_05(df, DONVI, BH, KP, HAN, dates):
    out = calc_metrics(df, DONVI, BH, KP, HAN, dates)
    out = out.sort_values("Quá hạn khắc phục", ascending=False).head(10)
    return add_total(out)

# ======================================================
# BẢNG 06 – CHI TIẾT PHÒNG/BAN HỘI SỞ
# ======================================================
def bang_06(df, KHOI, DONVI, BH, KP, HAN, dates):
    hs = df[df["KV"].str.contains("Hội sở", na=False)]
    tables = []
    for khoi, g in hs.groupby(KHOI):
        tong = calc_metrics(g, DONVI, BH, KP, HAN, dates).sum().to_frame().T
        tong.index = [f"Cộng {khoi}"]
        ct = calc_metrics(g, DONVI, BH, KP, HAN, dates)
        ct.index = ["   " + i for i in ct.index]
        tables += [tong, ct]
    return pd.concat(tables)

# ======================================================
# BẢNG 07 – CHI TIẾT DVKD
# ======================================================
def bang_07(df, KV, DONVI, BH, KP, HAN, dates):
    tables = []
    for kv, g in df.groupby(KV):
        tong = calc_metrics(g, DONVI, BH, KP, HAN, dates).sum().to_frame().T
        tong.index = [f"Cộng {kv}"]
        ct = calc_metrics(g, DONVI, BH, KP, HAN, dates)
        ct.index = ["   " + i for i in ct.index]
        tables += [tong, ct]
    return pd.concat(tables)

# ======================================================
# UI
# ======================================================
with st.sidebar:
    start = st.date_input("Từ ngày", datetime(datetime.now().year, 1, 1))
    end   = st.date_input("Đến ngày", datetime.now())
    file  = st.file_uploader("📂 File Excel KPCS", type=["xlsx", "xls"])

if file:
    df = load_excel(file)

    BH = find_column(df, ["Ngày, tháng, năm ban hành (mm/dd/yyyy)"])
    KP = find_column(df, ["NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)"])
    HAN = find_column(df, ["Thời hạn hoàn thành (mm/dd/yyyy)"])
    DONVI = find_column(df, ["Đơn vị thực hiện KPCS trong quý"])
    KHOI = find_column(df, ["SUM (THEO Khối, KV, ĐVKD, Hội sở, Ban Dự Án QLTS)"])
    KV = find_column(df, ["ĐVKD, AMC, Hội sở (Nhập ĐVKD hoặc Hội sở hoặc AMC)"])

    df["KV"] = df[KV]

    must_have({
        "Ngày ban hành": BH,
        "Ngày hoàn tất": KP,
        "Hạn": HAN,
        "Đơn vị": DONVI,
        "Khối/KV": KV
    })

    dates = {
        "year_start_date": pd.to_datetime(f"{end.year}-01-01"),
        "report_start_date": pd.to_datetime(start),
        "report_end_date": pd.to_datetime(end),
    }

    b01 = bang_01(df, KV, BH, KP, HAN, dates)
    b02 = bang_02(df, KHOI, BH, KP, HAN, dates)
    b03 = bang_03(df, DONVI, BH, KP, HAN, dates)
    b04 = bang_04(df, KV, BH, KP, HAN, dates)
    b05 = bang_05(df, DONVI, BH, KP, HAN, dates)
    b06 = bang_06(df, KHOI, DONVI, BH, KP, HAN, dates)
    b07 = bang_07(df, KV, DONVI, BH, KP, HAN, dates)

    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        b01.to_excel(w, "TK_KPCS_BANG_01")
        b02.to_excel(w, "TK_KPCS_BANG_02")
        b03.to_excel(w, "TK_KPCS_BANG_03")
        b04.to_excel(w, "TK_KPCS_BANG_04")
        b05.to_excel(w, "TK_KPCS_BANG_05")
        b06.to_excel(w, "TK_KPCS_BANG_06")
        b07.to_excel(w, "TK_KPCS_BANG_07")

    st.download_button(
        "📥 Tải Excel FULL 7 BẢNG",
        out.getvalue(),
        "BC_KPCS_FULL.xlsx"
    )
