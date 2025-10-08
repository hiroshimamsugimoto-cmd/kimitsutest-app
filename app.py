import streamlit as st
import openpyxl
from datetime import datetime
import base64
from io import BytesIO

# === 設定 ===
TEMPLATE = "気密試験記録.xlsx"   # 同じフォルダにテンプレートExcelを置く

st.title("📑 気密試験記録 入力フォーム")

# --- 入力項目 ---
系統名 = st.text_input("系統名")
試験圧力 = st.text_input("試験圧力 (MPa)")
試験範囲 = st.text_input("試験範囲")
試験媒体 = st.text_input("試験媒体")
放置時間 = st.text_input("放置時間 (h)")
使用機器No = st.text_input("使用圧力計機器No.")
測定場所 = st.text_input("測定場所")

# --- 開始日時 ---
st.subheader("開始日時")
col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    開始日 = st.date_input("日付", key="start_date")
with col2:
    開始時 = st.text_input("時", value="", placeholder="00", key="start_hour")
with col3:
    開始分 = st.text_input("分", value="", placeholder="00", key="start_minute")

# --- 終了日時 ---
st.subheader("終了日時")
col4, col5, col6 = st.columns([2, 1, 1])
with col4:
    終了日 = st.date_input("日付", key="end_date")
with col5:
    終了時 = st.text_input("時", value="", placeholder="00", key="end_hour")
with col6:
    終了分 = st.text_input("分", value="", placeholder="00", key="end_minute")

# --- 入力検証とdatetime生成 ---
try:
    開始時 = int(開始時) if 開始時.strip() != "" else 0
    開始分 = int(開始分) if 開始分.strip() != "" else 0
    終了時 = int(終了時) if 終了時.strip() != "" else 0
    終了分 = int(終了分) if 終了分.strip() != "" else 0

    開始時刻 = datetime.strptime(f"{開始時:02d}:{開始分:02d}", "%H:%M").time()
    終了時刻 = datetime.strptime(f"{終了時:02d}:{終了分:02d}", "%H:%M").time()

    開始日時 = datetime.combine(開始日, 開始時刻)
    終了日時 = datetime.combine(終了日, 終了時刻)
except ValueError:
    st.error("⚠ 時間の入力は 0〜23 時・0〜59 分の数値で入力してください。")

# --- 測定値入力 ---
st.subheader("測定値入力")

col5, col6 = st.columns([2, 2])
with col5:
    P1 = st.number_input("開始圧力 (MPa)", value=None, format="%.4f", step=None)
with col6:
    T1 = st.number_input("開始温度 (℃)", value=None, format="%.1f", step=None)

col7, col8 = st.columns([2, 2])
with col7:
    P2p = st.number_input("終了圧力 (MPa)", value=None, format="%.4f", step=None)
with col8:
    T2 = st.number_input("終了温度 (℃)", value=None, format="%.1f", step=None)


試験実施者 = st.text_input("試験実施者")

# --- 保存ボタン ---
if st.button("判定・保存"):
    if None in (P1, P2p, T1, T2):
        st.warning("⚠ 圧力・温度のすべてを入力してください。")
    else:
        wb = openpyxl.load_workbook(TEMPLATE)
        ws = wb["気密試験記録"]

        # Excelに書き込み
        ws["D3"] = 系統名
        ws["D4"] = 試験圧力
        ws["M4"] = 試験範囲
        ws["D5"] = 試験媒体
        ws["M5"] = 放置時間
        ws["D6"] = 使用機器No
        ws["M6"] = 測定場所
        ws["D8"] = 開始日時.strftime("%Y/%m/%d %H:%M")
        ws["M8"] = 終了日時.strftime("%Y/%m/%d %H:%M")

        ws["A10"] = f"{P1:.4f} "
        ws["C10"] = f"{T1:.1f} "
        ws["E10"] = f"{P2p:.4f}"
        ws["G10"] = f"{T2:.1f} "
        ws["E11"] = 試験実施者

        # --- ボイル・シャルルの法則で補正 ---
        try:
            T1_K = T1 + 273.15
            T2_K = T2 + 273.15
            P2_corr = P2p * (T1_K / T2_K)
            deltaP = P2_corr - P1

            # ✅ 判定範囲を固定 ±0.001MPa に変更
            判定範囲 = 0.001
            合否 = "合格" if abs(deltaP) <= 判定範囲 else "不合格"

            # Excelへ結果を書き込み
            ws["J10"] = f"{P2_corr:.4f} MPa"
            ws["M10"] = f"{deltaP:.4f} MPa"
            ws["O10"] = f"±{判定範囲:.4f} MPa"
            ws["M11"] = 合否

            # --- Streamlit画面出力 ---
            st.markdown("### 🧮 計算結果（ボイル・シャルルの法則に基づく補正）")
            st.write(f"- 補正後終了圧力 P2_corr: **{P2_corr:.4f} MPa**")
            st.write(f"- 圧力変化量 ⊿P: **{deltaP:.4f} MPa**")
            st.write(f"- 判定範囲: **±{判定範囲:.4f} MPa**")

            if 合否 == "合格":
                st.markdown(f"<h4 style='color:green;'>✅ 判定結果: {合否}</h4>", unsafe_allow_html=True)
            else:
                st.markdown(f"<h4 style='color:red;'>❌ 判定結果: {合否}</h4>", unsafe_allow_html=True)

        except Exception as e:
            st.error(f"⚠ 計算エラー: {e}")

        # --- ダウンロード処理 ---
        output = BytesIO()
        wb.save(output)
        excel_data = output.getvalue()
        filename = f"気密試験記録_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        b64 = base64.b64encode(excel_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 Excelをダウンロード</a>'
        st.markdown(href, unsafe_allow_html=True)
