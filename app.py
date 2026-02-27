import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ4：希望休と公休回数)")
st.write("「夜勤ロック」＋「日勤・夜勤の人数」＋「希望休の取得」＋「公休回数」を計算します！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        
        staff_names = df_staff["スタッフ名"].tolist()
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        
        # 公休回数を取得（空白ならとりあえず8回にする）
        if "公休回数" in df_staff.columns:
            staff_off_days = df_staff["公休回数"].fillna(8).tolist()
        else:
            staff_off_days = [8] * len(staff_names)
            
        num_staff = len(staff_names)
        
        date_columns = [col for col in df_history.columns if col != "スタッフ名" and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        # 夜勤人数の取得
        night_req_row = df_req[df_req.iloc[:, 0] == "夜勤人数"]
        night_req_list = night_req_row.iloc[0, 1:].dropna().tolist() + [2]*num_days if not night_req_row.empty else [2]*num_days
        
        # 日勤人数の取得（新機能！）
        day_req_row = df_req[df_req.iloc[:, 0] == "日勤人数"]
        day_req_list = day_req_row.iloc[0, 1:].dropna().tolist() + [3]*num_days if not day_req_row.empty else [3]*num_days
            
        st.success(f"✅ {num_staff}名のデータ、希望休、必要人数を読み込みました！")
        
        if st.button("シフトを自動作成する！（フェーズ4🔥）"):
            with st.spinner('AI店長がみんなの希望休と人数パズルを解いています...（最大20秒）'):
                
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
                        
                # ルール2: 夜勤セットの【完全ロック】
                for e in range(num_staff):
                    model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0:
                            model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days:
                            model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

                # ルール3: 毎日の「夜勤(D)」の必要人数
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == int(night_req_list[d]))

                # 🌟 ルール4: 毎日の「日勤(A)」の必要人数（指定人数"以上"配置する）
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'A')] for e in range(num_staff)) >= int(day_req_list[d]))

                # ルール5: リーダー配置（日勤にリーダー1名orサブ2名）
                for d in range(num_days):
                    leadership_score = sum(
                        (2 if "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * shifts[(e, d, 'A')]
                        for e in range(num_staff)
                    )
                    model.Add(leadership_score >= 2)

                # 🌟 ルール6: 希望休の絶対反映
                for e in range(num_staff):
                    for d in range(num_days):
                        # エクセルの該当マスの文字を取得
                        cell_value = str(df_history.iloc[e, d+1]).strip()
                        if cell_value == "公":
                            # もしエクセルに「公」と書いてあったら、絶対に「公休」にする！
                            model.Add(shifts[(e, d, '公')] == 1)

                # 🌟 ルール7: 月間の「公休回数」ノルマを達成する
                for e in range(num_staff):
                    target_off = int(staff_off_days[e])
                    # 1ヶ月の「公」の合計が、エクセルの公休回数とピッタリ一致すること！
                    model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == target_off)

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 20.0 
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨ 希望休も公休回数も完璧に守られています！")
                    
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
                        file_name="完成版_フェーズ4.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 条件が厳しすぎて組めませんでした。（原因例：公休希望が多すぎる、人数が足りない、など）")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
