import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# ======================================================
# CONFIG
# ======================================================
st.set_page_config(page_title="BC KPCS – Python = VBA", layout="wide")
st.title("📊 BÁO CÁO KPCS – LOGIC CHUẨN VBA (FIXED)")

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


# ======================================================
# LOAD DATA
# ======================================================
@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)

    # Ép toàn bộ cột có chữ "ngày" về datetime (an toàn)
    for c in df.columns:
        if "ngày" in c.lower():
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)

    return df


# ======================================================
# CORE LOGIC – CHUẨN VBA
# ======================================================
def valid_ton(df, BH, KP, TD, start, end):
    """
    TỒN tại cuối kỳ (giống VBA):
    - Ban hành <= end
    - Chưa KP hoặc KP sau end
    - Chưa chuyển TD riêng hoặc TD riêng nằm trong kỳ
    """
    return (
        (df[BH] <= end)
        &
        (
            df[KP].isna() |
            (df[KP] > end)
        )
        &
        (
            df[TD].isna() |
            (
                (df[TD] >= start) &
                (df[TD] <= end)
            )
        )
    )


def calculate_summary_metrics_vba(df, group_cols, BH, KP, TD, HAN, dates):
    y0 = dates["year_start_date"]
    s = dates["report_start_date"]
    e = dates["report_end_date"]

    # ÉP KIỂU CHẮC CHẮN (FIX LỖI OBJECT < TIMESTAMP)
    for col in [BH, KP, TD, HAN]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    def cnt(mask):
        dfx = df.loc[mask]
        if group_cols:
            return dfx.groupby(group_cols).size()
        return pd.Series({"TỔNG": len(dfx)})

    # ===== TỒN / PHÁT SINH / KHẮC PHỤC =====
    ton_dau_nam = cnt(
        (df[BH] < y0) &
        (
            df[KP].isna() |
            (df[KP] >= y0)
        )
    )

    phat_sinh_nam = cnt(
        (df[BH] >= y0) & (df[BH] <= e)
    )

    khac_phuc_nam = cnt(
        (df[KP].notna()) &
        (df[KP] >= y0) &
        (df[KP] <= e)
    )

    ton_dau_ky = cnt(
        (df[BH] < s) &
        (
            df[KP].isna() |
            (df[KP] >= s)
        )
    )

    phat_sinh_ky = cnt(
        (df[BH] >= s) & (df[BH] <= e)
    )

    khac_phuc_ky = cnt(
        (df[KP].notna()) &
        (df[KP] >= s) &
        (df[KP] <= e)
    )

    # ===== TỒN CUỐI KỲ (TÍNH TRỰC TIẾP THEO VBA) =====
    ton_cuoi_ky = ton_dau_ky + phat_sinh_ky - khac_phuc_ky

    # ===== QUÁ HẠN (CHỈ TÍNH TRÊN HỒ SƠ CÒN TỒN) =====
    ton_df = df.loc[valid_ton(df, BH, KP, TD, s, e)].copy()

    # ÉP KIỂU HẠN RIÊNG (QUAN TRỌNG)
    ton_df[HAN] = pd.to_datetime(ton_df[HAN], errors="coerce")

    qua_han = cnt(
        ton_df[HAN].notna() &
        (ton_df[HAN] < e)
    )

    qua_han_1n = cnt(
        ton_df[HAN].notna() &
        (ton_df[HAN] < (e - pd.DateOffset(years=1)))
    )

    # ===== GHÉP KẾT QUẢ =====
    out = pd.DataFrame({
        "Tồn đầu năm": ton_dau_nam,
        "Phát sinh năm": phat_sinh_nam,
        "Khắc phục năm": khac_phuc_nam,
        "Tồn đầu kỳ": ton_dau_ky,
        "Phát sinh kỳ": phat_sinh_ky,
        "Khắc phục kỳ": khac_phuc_ky,
        "Tồn cuối kỳ": ton_cuoi_ky,
        "Quá hạn": qua_han,
        "Quá hạn >1 năm": qua_han_1n
    }).fillna(0).astype(int)

    denom = out["Tồn đầu năm"] + out["Phát sinh năm"]
    out["Tỷ lệ chưa KP"] = np.where(denom > 0, out["Tồn cuối kỳ"] / denom, 0)

    return out.reset_index().rename(columns={"index": "Nhóm"})


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

    # ===== MAP CỘT (ĐÚNG FILE BẠN) =====
    BH = find_column(df, [
        "Ngày, tháng, năm ban hành (mm/dd/yyyy)",
        "Ngày ban hành"
    ])

    KP = find_column(df, [
        "NGÀY HOÀN TẤT KPCS (mm/dd/yyyy)",
        "Ngày hoàn tất"
    ])

    TD = find_column(df, [
        "NGÀY CHUYỂN THEO DÕI RIÊNG (mm/dd/yyyy)"
    ])

    HAN = find_column(df, [
        "Thời hạn hoàn thành (mm/dd/yyyy)",
        "Hạn KPCS"
    ])

    must_have({
        "Ngày ban hành": BH,
        "Ngày hoàn tất": KP,
        "Ngày chuyển TD riêng": TD,
        "Hạn KPCS": HAN
    })

    # ===== NHÓM (GIỐNG VBA – CÓ THỂ MỞ RỘNG) =====
    df["NHÓM"] = "TỔNG"

    dates = {
        "year_start_date": pd.to_datetime(f"{end.year}-01-01"),
        "report_start_date": pd.to_datetime(start),
        "report_end_date": pd.to_datetime(end),
    }

    st.subheader("📊 BẢNG 01 – TỔNG HỢP (CHUẨN VBA)")
    bang01 = calculate_summary_metrics_vba(
        df,
        ["NHÓM"],
        BH, KP, TD, HAN,
        dates
    )
    st.dataframe(bang01, use_container_width=True)

    # ===== EXPORT =====
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        bang01.to_excel(writer, sheet_name="BANG_01", index=False)

    st.download_button(
        "📥 Tải Excel BẢNG 01",
        data=output.getvalue(),
        file_name="BC_KPCS_BANG_01_PYTHON.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("⬅️ Vui lòng upload file Excel KPCS")
