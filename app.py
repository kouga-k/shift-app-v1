import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ5：夜勤上限＆夜勤不可)")
st.write("各スタッフの「夜勤上限（0なら夜勤不可）」を守ってシフトを組みます！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        
        staff_names = df_staff["スタッフ名"].dropna().tolist()
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        staff_off_days = df_staff["公休回数"].fillna(8).tolist() if "公休回数" in df_staff.columns else [8]*len(staff_names)
        
        # 🌟 新機能：夜勤上限の取得（空欄の場合は上限なしとして仮に10回とする）
        if "夜勤上限" in df_staff.columns:
            staff_night_limits = df_staff["夜勤上限"].fillna(10).tolist()
        else:
            staff_night_limits = [10] * len(staff_names)

        num_staff = len(staff_names)
        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        night_req_row = df_req[df_req.iloc[:, 0] == "夜勤人数"]
        day_req_row = df_req[df_req.iloc[:, 0] == "日勤人数"]
        
        night_req_list = []
        day_req_list = []
        for col in date_columns:
            if not night_req_row.empty and col in night_req_row.columns:
                val = night_req_row[col].values[0]
                night_req_list.append(int(val) if pd.notna(val) else 2)
            else:
                night_req_list.append(2)
            if not day_req_row.empty and col in day_req_row.columns:
                val = day_req_row[col].values[0]
                day_req_list.append(int(val) if pd.notna(val) else 3)
            else:
                day_req_list.append(3)
            
        st.success(f"✅ データの読み込み完了！各スタッフの夜勤上限を考慮して計算します...")
        
        if st.button("シフトを自動作成する！（フェーズ5🔥）"):
            with st.spinner('AI店長がみんなの希望休と夜勤上限パズルを解いています...（最大30秒）'):
                
                model = cp_model.CpModel()
                shift_types = ['A', 'D', 'E', '公']
                
                shifts = {}
                for e in range(num_staff):
                    for d in range(num_days):
                        for s in shift_types:
                            shifts[(e, d, s)] = model.NewBoolVar(f'shift_{e}_{d}_{s}')
                            
                # ルール1: 毎日必ずどれか1つ
                for e in range(num_staff):
                    for d in range(num_days):
                        model.AddExactlyOne(shifts[(e, d, s)] for s in shift_types)
                        
                # ルール2: 夜勤セットの完全ロック
                for e in range(num_staff):
                    model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0:
                            model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days:
                            model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

                # ルール3: 毎日の「夜勤(D)」の必要人数
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])

                # ルール4: 毎日の「日勤(A)」の必要人数
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'A')] for e in range(num_staff)) >= day_req_list[d])

                # ルール5: リーダー配置（日勤にリーダー1名orサブ2名）
                for d in range(num_days):
                    leadership_score = sum(
                        (2 if "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * shifts[(e, d, 'A')]
                        for e in range(num_staff)
                    )
                    model.Add(leadership_score >= 2)

                # ルール6: 希望休の完全ピンポイント検索
                for e, staff_name in enumerate(staff_names):
                    for d, date_col in enumerate(date_columns):
                        if date_col in df_history.columns:
                            target_row = df_history[df_history["スタッフ名"] == staff_name]
                            if not target_row.empty:
                                cell_value = str(target_row[date_col].values[0]).strip()
                                if cell_value == "公":
                                    model.Add(shifts[(e, d, '公')] == 1)

                # ルール7: 公休回数のノルマ
                for e in range(num_staff):
                    target_off = int(staff_off_days[e])
                    model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == target_off)

                # 🌟 新ルール8: スタッフごとの「夜勤(D)」の上限回数（0なら夜勤不可）
                for e in range(num_staff):
                    max_night = int(staff_night_limits[e])
                    # 1ヶ月の夜勤(D)の合計が、エクセルの上限の数字以下であること！
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= max_night)

                solver = cp_model.CpSolver()
                # 複雑な条件になったので、考える時間を少し（30秒）長くします
                solver.parameters.max_time_in_seconds = 30.0 
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨ 各スタッフの夜勤上限も完璧に守られています！")
                    
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e], "役割": staff_roles[e]}
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    row[date_columns[d]] = s
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)

                    # 右側の集計（個人の回数）
                    result_df['日勤(A)回数'] = (result_df[date_columns] == 'A').sum(axis=1)
                    result_df['夜勤(D)回数'] = (result_df[date_columns] == 'D').sum(axis=1)
                    result_df['公休回数'] = (result_df[date_columns] == '公').sum(axis=1)
                    
                    # 🌟 上限の確認用に、エクセルに書いた「夜勤上限」の数字も右端に表示する
                    result_df['夜勤上限(設定値)'] = staff_night_limits

                    # 下側の集計（毎日の人数）
                    summary_A = {"スタッフ名": "【日勤(A) 合計】", "役割": ""}
                    summary_D = {"スタッフ名": "【夜勤(D) 合計】", "役割": ""}
                    summary_Off = {"スタッフ名": "【公休 合計】", "役割": ""}
                    
                    for col in ['日勤(A)回数', '夜勤(D)回数', '公休回数', '夜勤上限(設定値)']:
                        summary_A[col] = ""
                        summary_D[col] = ""
                        summary_Off[col] = ""

                    for col in date_columns:
                        summary_A[col] = (result_df[col] == 'A').sum()
                        summary_D[col] = (result_df[col] == 'D').sum()
                        summary_Off[col] = (result_df[col] == '公').sum()

                    summary_df = pd.DataFrame([summary_A, summary_D, summary_Off])
                    final_df = pd.concat([result_df, summary_df], ignore_index=True)

                    st.dataframe(final_df)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成したシフトをダウンロード",
                        data=processed_data,
                        file_name="完成版_夜勤上限対応.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 条件が厳しすぎて組めませんでした。（原因例：夜勤の上限を厳しくしすぎて、毎日の夜勤人数を確保できない、など）")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
