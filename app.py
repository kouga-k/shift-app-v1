import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ3.1：条件緩和テスト)")
st.write("「夜勤セット」＋「夜勤の必要人数」＋「リーダー配置」を計算します！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        
        staff_names = df_staff["スタッフ名"].tolist()
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        num_staff = len(staff_names)
        
        date_columns = [col for col in df_history.columns if col != "スタッフ名" and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        night_req_row = df_req[df_req.iloc[:, 0] == "夜勤人数"]
        if not night_req_row.empty:
            # エクセルのB列(インデックス1)以降から、日数分だけ数字を取得する。
            night_req_values = night_req_row.iloc[0, 1:].dropna().tolist()
            # もし数字が足りなければ、最後の数字（または2）で埋める
            last_val = night_req_values[-1] if night_req_values else 2
            night_req_list = night_req_values + [last_val] * (num_days - len(night_req_values))
        else:
            night_req_list = [2] * num_days
            
        st.success(f"✅ {num_staff}名のスタッフデータを読み込みました。計算を開始します...")
        
        if st.button("シフトを自動作成する！（フェーズ3.1🔥）"):
            with st.spinner('AI店長が複雑なパズルを解いています...'):
                
                model = cp_model.CpModel()
                shift_types = ['A', 'D', 'E', '公']
                
                shifts = {}
                for e in range(num_staff):
                    for d in range(num_days):
                        for s in shift_types:
                            shifts[(e, d, s)] = model.NewBoolVar(f'shift_{e}_{d}_{s}')
                            
                # ルール1: 毎日必ずどれか1つの勤務
                for e in range(num_staff):
                    for d in range(num_days):
                        model.AddExactlyOne(shifts[(e, d, s)] for s in shift_types)
                        
                # ルール2: 夜勤セット（D -> E -> 公） ※月末のはみ出しも許容する（翌月のことは一旦気にしない）
                for e in range(num_staff):
                    for d in range(num_days):
                        if d + 1 < num_days:
                            model.AddImplication(shifts[(e, d, 'D')], shifts[(e, d+1, 'E')])
                        if d + 2 < num_days:
                            model.AddImplication(shifts[(e, d+1, 'E')], shifts[(e, d+2, '公')])

                # ルール3: 毎日の「夜勤(D)」の必要人数を守る
                for d in range(num_days):
                    target_night = int(night_req_list[d])
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == target_night)

                # ルール4: リーダー1名、またはサブ2名以上を「日勤(A)」に配置
                for d in range(num_days):
                    leadership_score = 0
                    for e in range(num_staff):
                        role = str(staff_roles[e])
                        if "リーダー" in role:
                            leadership_score += 2 * shifts[(e, d, 'A')]
                        elif "サブ" in role:
                            leadership_score += 1 * shifts[(e, d, 'A')]
                    # その日の日勤の合計ポイントが2以上であること！
                    model.Add(leadership_score >= 2)

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 15.0 # タイマーを少し長めに
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨ リーダー/サブの配置も完璧です！")
                    
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e], "役割": staff_roles[e]}
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    row[date_columns[d]] = s
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)
                    st.dataframe(result_df)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成したシフトをダウンロード",
                        data=processed_data,
                        file_name="完成版_フェーズ3.xlsx"
                    )
                else:
                    st.error("❌ 条件が厳しすぎてシフトが組めませんでした。スタッフ人数を増やすか、夜勤の必要人数を減らしたエクセルで再度試してください。")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
