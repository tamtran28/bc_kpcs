import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="BC KPCS – Python = VBA", layout="wide")
st.title("📊 BÁO CÁO KPCS – LOGIC CHUẨN VBA")

# ======================================================
# HELPER
# ======================================================
def find_column(df, names):
    for n in names:
        if n in df.columns:
            return n
    return None


def must_have(df, mapping):
    for name, col in mapping.items():
        if col is None:
            st.error(f"❌ Thiếu cột bắt buộc: {name}")
            st.stop()


# ======================================================
# LOAD DATA
# ======================================================
@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)

    for c in df.columns:
        if "ngày" in c.lower():
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    return df


# ======================================================
# CORE LOGIC – VBA 1–1
# ======================================================
def valid_ton(df, BH, KP, TD, start, end):
    return (
        (df[BH] <= end)
        & (
            df[KP].isna()
            | (df[KP] > end)
        )
        & (
            df[TD].isna()
            | (
                (df[TD] >= start)
                & (df[TD] <= end)
            )
        )
    )


def calculate_summary_metrics_vba(df, group_cols, BH, KP, TD, HAN, dates):
    y0 = dates["year_start_date"]
    s = dates["report_start_date"]
    e = dates["report_end_date"]

    def cnt(mask):
        if group_cols:
            return df.loc[mask].groupby(group_cols).size()
        return pd.Series({"ALL": mask.sum()})

    ton_dau_nam = cnt(
        (df[BH] < y0)
        & (
            df[KP].isna()
            | (df[KP] >= y0)
        )
    )

    phat_sinh_nam = cnt(
        (df[BH] >= y0) & (df[BH] <= e)
    )

    khac_phuc_nam = cnt(
        (df[KP] >= y0) & (df[KP] <= e)
    )

    ton_dau_ky = cnt(
        (df[BH] < s)
        & (
            df[KP].isna()
            | (df[KP] >= s)
        )
    )

    phat_sinh_ky = cnt(
        (df[BH] >= s) & (df[BH] <= e)
    )

    khac_phuc_ky = cnt(
        (df[KP] >= s) & (df[KP] <= e)
    )

    ton_ck = ton_dau_ky + phat_sinh_ky - khac_phuc_ky

    ton_df = df[valid_ton(df, BH, KP, TD, s, e)]

    qua_han = cnt(ton_df[ton_df[HAN] < e].index)
    qua_han_1n = cnt(ton_df[ton_df[HAN] < (e - pd.DateOffset(years=1))].index)

    out = pd.DataFrame({
        "Tồn đầu năm": ton_dau_nam,
        "Phát sinh năm": phat_sinh_nam,
        "Khắc phục năm": khac_phuc_nam,
        "Tồn đầu kỳ": ton_dau_ky,
        "Phát sinh kỳ": phat_sinh_ky,
        "Khắc phục kỳ": khac_phuc_ky,
        "Tồn cuối kỳ": ton_ck,
        "Quá hạn": qua_han,
        "Quá hạn >1 năm": qua_han_1n
    }).fillna(0).astype(int)

    denom = out["Tồn đầu năm"] + out["Phát sinh năm"]
    out["Tỷ lệ chưa KP"] = np.where(denom > 0, out["Tồn cuối kỳ"] / denom, 0)

    return out.reset_index().rename(columns={"index": "Đơn vị"})


# ======================================================
# UI
# ======================================================
with st.sidebar:
    st.header("⚙️ TÙY CHỌN")
    start = st.date_input("Từ ngày", datetime(datetime.now().year, 1, 1))
    end = st.date_input("Đến ngày", datetime.now())
    file = st.file_uploader("📂 File Excel", type=["xlsx", "xls"])

if file:
    df = load_excel(file)

    BH = find_column(df, ["Ngày, tháng, năm ban hành (mm/dd/yyyy)", "Ngày ban hành"])
    KP = find_column(df, ["NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)", "Ngày hoàn tất"])
    TD = find_column(df, ["NGÀY CHUYỂN THEO DÕI RIÊNG (mm/dd/yyyy)"])
    HAN = find_column(df, ["Thời hạn hoàn thành (mm/dd/yyyy)"])
    DONVI = find_column(df, ["Đơn vị thực hiện KPCS trong quý", "Đơn vị"])

    must_have(df, {
        "Ngày ban hành": BH,
        "Ngày hoàn tất": KP,
        "Ngày chuyển TD riêng": TD,
        "Thời hạn": HAN,
        "Đơn vị": DONVI
    })

    dates = {
        "year_start_date": pd.to_datetime(f"{end.year}-01-01"),
        "report_start_date": pd.to_datetime(start),
        "report_end_date": pd.to_datetime(end),
    }

    df["NHÓM"] = "TỔNG"

    st.subheader("📋 XEM DỮ LIỆU")
    st.dataframe(df.head())

    st.subheader("📊 BẢNG 01 – TỔNG HỢP")
    bang01 = calculate_summary_metrics_vba(
        df,
        ["NHÓM"],
        BH, KP, TD, HAN,
        dates
    )
    st.dataframe(bang01)

    if st.button("📥 TẢI EXCEL"):
        bio = BytesIO()
        with pd.ExcelWriter(bio, engine="xlsxwriter") as w:
            bang01.to_excel(w, sheet_name="BANG_01", index=False)
        st.download_button(
            "⬇️ Download",
            bio.getvalue(),
            "BC_KPCS_PYTHON.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬅️ Upload file Excel để bắt đầu")
