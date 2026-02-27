import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="centered")
st.title("🌟 AI自動シフト作成アプリ")
st.write("スタッフの名前が書かれたエクセルをアップロードしてください。")

# --- 1. エクセルをアップロードする画面 ---
uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    # エクセルを読み込む
    df = pd.read_excel(uploaded_file)
    st.success("エクセルを読み込みました！")
    st.dataframe(df) # 画面に表を表示

    # エクセル内に「スタッフ名」という列があるか確認
    if "スタッフ名" not in df.columns:
        st.error("エラー：エクセルの1行目に「スタッフ名」という見出しを作ってください。")
    else:
        # --- 2. ここからOR-Toolsの計算（本格版） ---
        if st.button("シフトを自動作成する！"):
            with st.spinner('AI店長がパズルを解いています...'):
                
                # スタッフのリストと日数（今回は仮で30日）を取得
                staff_names = df["スタッフ名"].tolist()
                num_staff = len(staff_names)
                num_days = 30
                
                model = cp_model.CpModel()
                shifts = {}
                
                # マス目を作る
                for e in range(num_staff):
                    for d in range(num_days):
                        shifts[(e, d)] = model.NewBoolVar(f'shift_{e}_{d}')
                
                # ルール1：毎日、必ず「2人」が出勤する（本格的！）
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d)] for e in range(num_staff)) == 2)
                
                # パズルを解かせる
                solver = cp_model.CpSolver()
                status = solver.Solve(model)
                
                # --- 3. 結果をエクセルにしてダウンロードさせる ---
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨")
                    
                    # 結果を新しい表（データフレーム）にまとめる
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e]}
                        for d in range(num_days):
                            # 出勤(1)なら〇、休み(0)なら×にする
                            row[f"{d+1}日"] = "〇" if solver.Value(shifts[(e, d)]) == 1 else "休"
                        result_data.append(row)
                    
                    result_df = pd.DataFrame(result_data)
                    st.dataframe(result_df) # 完成した表を画面に出す
                    
                    # エクセルファイルに変換する魔法
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成したエクセルをダウンロード",
                        data=processed_data,
                        file_name="完成版_自動シフト.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("条件が厳しすぎてシフトが組めませんでした。ルールを見直してください。")
