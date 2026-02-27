import streamlit as st
import pandas as pd

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide") # 画面を広く使う設定
st.title("🌟 AI自動シフト作成アプリ (データ読み込みテスト中)")
st.write("作成した新しいエクセルファイルをアップロードしてください。")

# エクセルをアップロードする入り口
uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        st.info("AIがエクセルの中身を解析しています...")

        # 1. 「スタッフ設定」シートを読み込む
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        st.success("✅ 「スタッフ設定」シートの読み込みに成功しました！")
        st.dataframe(df_staff) # 画面に表を表示

        # 2. 「希望休・先月履歴」シートを読み込む
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        st.success("✅ 「希望休・先月履歴」シートの読み込みに成功しました！")
        st.dataframe(df_history)

        # 3. 「必要人数設定」シートを読み込む
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        st.success("✅ 「必要人数設定」シートの読み込みに成功しました！")
        st.dataframe(df_req)

        st.balloons() # 成功の風船を飛ばす！
        st.info("素晴らしいです！AIがあなたの作った3つのシートを完璧に認識しました。次はここに『シフト作成の頭脳（OR-Tools）』を組み込んでいきます！")

    except ValueError:
        # シート名が間違っていた場合のエラーメッセージ
        st.error("⚠️ エラー：シート名が見つかりません。エクセルの下のタブの名前が「スタッフ設定」「希望休・先月履歴」「必要人数設定」と一字一句同じになっているか（空白などが入っていないか）確認してください。")
    except Exception as e:
        st.error(f"⚠️ 予期せぬエラーが発生しました: {e}")
