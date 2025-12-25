import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime

# =====================================================
# CONFIG
# =====================================================
st.set_page_config(page_title="KPCS – Chuẩn VBA", layout="wide")
st.title("📊 HỆ THỐNG BÁO CÁO KPCS (CHUẨN 1–1 VBA)")

# =====================================================
# LOAD DATA
# =====================================================
@st.cache_data
def load_data(file):
    df = pd.read_excel(file)

    date_cols = [
        'Ngày, tháng, năm ban hành (mm/dd/yyyy)',
        'NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)',
        'Thời hạn hoàn thành (mm/dd/yyyy)',
        'Ngày theo dõi riêng'
    ]
    for c in date_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    return df


# =====================================================
# CORE FILTER – CHUẨN VBA
# =====================================================
def filter_ton(df, ban_hanh, kp, theo_doi, start, end):
    return (
        (df[ban_hanh] <= end) &
        (
            df[kp].isna() |
            (df[kp] > end)
        ) &
        (
            df[theo_doi].isna() |
            (
                (df[theo_doi] >= start) &
                (df[theo_doi] <= end)
            )
        )
    )


# =====================================================
# CORE SUMMARY – BẢNG 01 (CHUẨN VBA)
# =====================================================
def calculate_summary_metrics_vba(
    df,
    groupby_cols,
    year_start_date,
    report_start_date,
    report_end_date,
    col_bh='Ngày, tháng, năm ban hành (mm/dd/yyyy)',
    col_kp='NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)',
    col_td='Ngày theo dõi riêng',
    col_han='Thời hạn hoàn thành (mm/dd/yyyy)'
):

    def valid_td(dfx, start, end):
        return (
            dfx[col_td].isna() |
            (
                (dfx[col_td] >= start) &
                (dfx[col_td] <= end)
            )
        )

    def agg(dfx):
        if dfx.empty:
            return pd.Series(dtype=int)
        return dfx.groupby(groupby_cols).size()

    ton_dau_nam = df[
        (df[col_bh] < year_start_date) &
        (
            df[col_kp].isna() |
            (df[col_kp] >= year_start_date)
        ) &
        valid_td(df, year_start_date, report_end_date)
    ]

    phat_sinh_nam = df[
        (df[col_bh] >= year_start_date) &
        (df[col_bh] <= report_end_date)
    ]

    khac_phuc_nam = df[
        (df[col_kp].notna()) &
        (df[col_kp] >= year_start_date) &
        (df[col_kp] <= report_end_date)
    ]

    ton_dau_ky = df[
        (df[col_bh] < report_start_date) &
        (
            df[col_kp].isna() |
            (df[col_kp] >= report_start_date)
        ) &
        valid_td(df, report_start_date, report_end_date)
    ]

    phat_sinh_ky = df[
        (df[col_bh] >= report_start_date) &
        (df[col_bh] <= report_end_date)
    ]

    khac_phuc_ky = df[
        (df[col_kp].notna()) &
        (df[col_kp] >= report_start_date) &
        (df[col_kp] <= report_end_date)
    ]

    ton_cuoi_ky = df[
        filter_ton(df, col_bh, col_kp, col_td, report_start_date, report_end_date)
    ]

    qua_han = ton_cuoi_ky[ton_cuoi_ky[col_han] < report_end_date]
    qua_han_1_nam = ton_cuoi_ky[
        ton_cuoi_ky[col_han] < (report_end_date - pd.DateOffset(days=365))
    ]

    summary = pd.DataFrame({
        "Tồn đầu năm": agg(ton_dau_nam),
        "Phát sinh năm": agg(phat_sinh_nam),
        "Khắc phục năm": agg(khac_phuc_nam),
        "Tồn đầu kỳ": agg(ton_dau_ky),
        "Phát sinh kỳ": agg(phat_sinh_ky),
        "Khắc phục kỳ": agg(khac_phuc_ky),
        "Tồn cuối kỳ": agg(ton_cuoi_ky),
        "Quá hạn": agg(qua_han),
        "Quá hạn >1 năm": agg(qua_han_1_nam),
    }).fillna(0).astype(int)

    return summary


# =====================================================
# BẢNG 01 → 07
# =====================================================
def bang_01(df, dates):
    kq = calculate_summary_metrics_vba(df, ['NHÓM'], **dates)
    total = pd.DataFrame(kq.sum()).T
    total.index = ['TỔNG']
    return pd.concat([kq, total])


def bang_02(df, dates):
    return bang_01(df, dates).loc[lambda x: (x != 0).any(axis=1)]


def bang_03(df, dates, top_n=10):
    ton = df[filter_ton(df, BH, KP, TD, dates['report_start_date'], dates['report_end_date'])]
    return ton.groupby('Đơn vị').size().sort_values(ascending=False).head(top_n).to_frame("Tồn cuối kỳ")


def bang_04(df, dates):
    ton = df[filter_ton(df, BH, KP, TD, dates['report_start_date'], dates['report_end_date'])]
    ton = ton.copy()
    ton["Số ngày quá hạn"] = (dates['report_end_date'] - ton[HAN]).dt.days
    bins = [-1, 90, 180, 270, 365, 10**9]
    labels = ["<3 tháng", "3–6", "6–9", "9–12", ">1 năm"]
    ton["Nhóm"] = pd.cut(ton["Số ngày quá hạn"], bins=bins, labels=labels)
    return ton.groupby("Nhóm").size().to_frame("Số lượng")


def bang_05(df, dates, top_n=10):
    ton = df[
        filter_ton(df, BH, KP, TD, dates['report_start_date'], dates['report_end_date']) &
        (df[HAN] < dates['report_end_date'])
    ]
    return ton.groupby("Đơn vị").size().sort_values(ascending=False).head(top_n).to_frame("Quá hạn")


def bang_06(df, dates):
    ton = df[filter_ton(df, BH, KP, TD, dates['report_start_date'], dates['report_end_date'])]
    return ton.groupby(["Khối", "Khu vực"]).size().to_frame("Tồn")


def bang_07(df, dates):
    ton = df[filter_ton(df, BH, KP, TD, dates['report_start_date'], dates['report_end_date'])]
    return ton.groupby(["Khối", "Khu vực", "Đơn vị"]).size().to_frame("Tồn")


# =====================================================
# UI
# =====================================================
with st.sidebar:
    st.header("⚙️ Tham số")
    start_date = st.date_input("Từ ngày", datetime(datetime.now().year, 1, 1))
    end_date = st.date_input("Đến ngày", datetime.now())
    file = st.file_uploader("📂 File Excel KPCS", type=["xls", "xlsx"])

if file:
    df = load_data(file)

    BH = 'Ngày, tháng, năm ban hành (mm/dd/yyyy)'
    KP = 'NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)'
    TD = 'Ngày theo dõi riêng'
    HAN = 'Thời hạn hoàn thành (mm/dd/yyyy)'

    df["NHÓM"] = np.where(
        df["ĐVKD, AMC, Hội sở (Nhập ĐVKD hoặc Hội sở hoặc AMC)"] == "Hội sở",
        "Hội sở",
        "ĐVKD"
    )

    dates = {
        "year_start_date": pd.to_datetime(f"{end_date.year}-01-01"),
        "report_start_date": pd.to_datetime(start_date),
        "report_end_date": pd.to_datetime(end_date),
    }

    tables = {
        "BẢNG 01": bang_01(df, dates),
        "BẢNG 02": bang_02(df, dates),
        "BẢNG 03": bang_03(df, dates),
        "BẢNG 04": bang_04(df, dates),
        "BẢNG 05": bang_05(df, dates),
        "BẢNG 06": bang_06(df, dates),
        "BẢNG 07": bang_07(df, dates),
    }

    for name, table in tables.items():
        st.subheader(name)
        st.dataframe(table)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name.replace(" ", "_"))

    st.download_button(
        "📥 Tải Excel đầy đủ 01–07",
        data=output.getvalue(),
        file_name="KPCS_FULL_PYTHON_CHUAN_VBA.xlsx"
    )
else:
    st.info("⬅️ Vui lòng tải file Excel KPCS")
