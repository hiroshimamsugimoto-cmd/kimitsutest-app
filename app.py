import streamlit as st
import openpyxl
from datetime import datetime
from io import BytesIO
import base64

# === 設定 ===
TEMPLATE = "気密試験記録.xlsx"  # 同じフォルダにテンプレートExcelを置く

st.title("🧾 気密試験記録 入力フォーム")

# --- 入力項目 ---
st.subheader("試験情報入力")
system_name = st.text_input("系統名")
test_pressure = st.text_input("試験圧力 (MPa)")
test_range = st.text_input("試験範囲")
test_medium = st.text_input("試験媒体")
test_time = st.text_input("放置時間 (h)")
gauge_no = st.text_input("使用圧力計機器No.")
test_location = st.text_input("測定場所")

# --- 日時 ---
st.subheader("開始日時")
col1, col2, col3 = st.columns(3)
with col1:
    start_date = st.date_input("日付", datetime.now().date())
with col2:
    start_hour = st.number_input("時", min_value=0, max_value=23, value=9)
with col3:
    start_min = st.number_input("分", min_value=0, max_value=59, value=0)

st.subheader("終了日時")
col4, col5, col6 = st.columns(3)
with col4:
    end_date = st.date_input("日付 ", datetime.now().date())
with col5:
    end_hour = st.number_input("時 ", min_value=0, max_value=23, value=10)
with col6:
    end_min = st.number_input("分 ", min_value=0, max_value=59, value=0)

# --- 測定値入力 ---
st.subheader("測定値入力")

col5, col6 = st.columns(2)
with col5:
    P1 = st.text_input("開始圧力 (MPa)", placeholder="例：0.0799")
with col6:
    T1 = st.text_input("開始温度 (℃)", placeholder="例：27.2")

col7, col8 = st.columns(2)
with col7:
    P2p = st.text_input("終了圧力 (MPa)", placeholder="例：0.0815")
with col8:
    T2 = st.text_input("終了温度 (℃)", placeholder="例：29.8")

tester = st.text_input("試験実施者")

# --- 数値変換 ---
def safe_float(v):
    try:
        return float(v.strip()) if v else None
    except:
        return None

P1 = safe_float(P1)
T1 = safe_float(T1)
P2p = safe_float(P2p)
T2 = safe_float(T2)

# --- 判定処理 ---
if st.button("判定・保存"):

    if None in (P1, T1, P2p, T2):
        st.error("⚠ 数値入力が不足しています。")
    else:
        # ボイル・シャルル補正
        P2_corr = P2p * (T1 + 273.15) / (T2 + 273.15)
        delta_P = P2_corr - P1
        tolerance = P1 * 0.01  # ±1%

        if abs(delta_P) <= tolerance:
            result = "合格"
            result_color = "green"
        else:
            result = "不合格"
            result_color = "red"

        # 結果表示
        st.markdown("## 📊 計算結果（ボイル・シャルルの法則に基づく補正）")
        st.write(f"- 補正後終了圧力 P2_corr: **{P2_corr:.4f} MPa**")
        st.write(f"- 圧力変化量 ΔP: **{delta_P:.4f} MPa**")
        st.write(f"- 判定範囲: ±**{tolerance:.4f} MPa**")
        st.markdown(f"### <span style='color:{result_color};'>判定結果: {result}</span>", unsafe_allow_html=True)

        # --- Excel 出力 ---
        wb = openpyxl.load_workbook(TEMPLATE)
        ws = wb.active

        ws["C6"].value = system_name
        ws["C7"].value = test_pressure
        ws["C8"].value = test_range
        ws["C9"].value = test_medium
        ws["C10"].value = test_time
        ws["C11"].value = gauge_no
        ws["C12"].value = test_location

        ws["C14"].value = str(start_date)
        ws["E14"].value = f"{start_hour}:{start_min:02d}"
        ws["C15"].value = str(end_date)
        ws["E15"].value = f"{end_hour}:{end_min:02d}"

        ws["C17"].value = P1
        ws["C18"].value = T1
        ws["E17"].value = P2p
        ws["E18"].value = T2

        ws["F17"].value = P2_corr
        ws["G17"].value = delta_P
        ws["H17"].value = f"±{tolerance:.4f}"
        ws["I17"].value = result
        ws["C20"].value = tester

        # Excel保存
        output = BytesIO()
        wb.save(output)
        excel_data = output.getvalue()
        b64 = base64.b64encode(excel_data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64}" download="気密試験記録.xlsx">📥 Excelをダウンロード</a>'
        st.markdown(href, unsafe_allow_html=True)
