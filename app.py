import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ2：夜勤セット実装テスト)")
st.write("「D → E → 公」の絶対ルールをAIが守れるかテストします！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        # エクセルの読み込み
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        
        # AIに「スタッフの名前」と「人数」を覚えさせる
        staff_names = df_staff["スタッフ名"].dropna().tolist()
        num_staff = len(staff_names)
        
        # AIに「日付（カレンダーの列）」を覚えさせる
        # （※"スタッフ名"や空白の列名を除外して、日付の列だけを抽出します）
        date_columns = [col for col in df_history.columns if col != "スタッフ名" and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        st.success(f"✅ {num_staff}名のスタッフと、{num_days}日分のカレンダーを認識しました！")
        
        if st.button("シフトを自動作成する！（夜勤セット発動🌙）"):
            with st.spinner('AI店長が夜勤セットのパズルを解いています...'):
                
                model = cp_model.CpModel()
                
                # 今回使う勤務の種類
                shift_types = ['A', 'D', 'E', '公']
                
                # ① シフトのマス目を作る（裏側の準備）
                shifts = {}
                for e in range(num_staff):
                    for d in range(num_days):
                        for s in shift_types:
                            shifts[(e, d, s)] = model.NewBoolVar('')
                            
                # ② ルール1: 毎日必ずどれか1つの勤務に就く
                for e in range(num_staff):
                    for d in range(num_days):
                        model.AddExactlyOne(shifts[(e, d, s)] for s in shift_types)
                        
                # ③ ルール2: 夜勤セットの絶対ルール（D -> E -> 公）
                for e in range(num_staff):
                    for d in range(num_days - 2): # 最終日付近は枠外にはみ出ないように処理
                        # もし今日が「D」なら、明日は必ず「E」にしなさい
                        model.AddImplication(shifts[(e, d, 'D')], shifts[(e, d+1, 'E')])
                        # もし明日が「E」なら、明後日は必ず「公」にしなさい
                        model.AddImplication(shifts[(e, d+1, 'E')], shifts[(e, d+2, '公')])
                        
                # 枠の終端処理（月末の最後の2日間にDを入れると翌月にはみ出るので、一旦禁止にする）
                for e in range(num_staff):
                    model.Add(shifts[(e, num_days-1, 'D')] == 0)
                    model.Add(shifts[(e, num_days-2, 'D')] == 0)

                # ④ ルール3: テスト用に、全員に最低2回の夜勤(D)をやらせる
                for e in range(num_staff):
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) >= 2)
                
                # パズルを解かせる！
                solver = cp_model.CpSolver()
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨")
                    
                    # カレンダー形式の表（エクセルと同じ形）にまとめる
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e]}
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    row[date_columns[d]] = s # 該当する日付の列に記号を入れる
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)
                    st.dataframe(result_df) # 画面に表示
                    
                    # エクセル出力の準備
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成したシフト（エクセル）をダウンロード",
                        data=processed_data,
                        file_name="完成版_夜勤セットテスト.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("条件が厳しすぎてシフトが組めませんでした。")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
