import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ4.2：見やすい集計機能付き)")
st.write("「A」「D」「公」の回数や人数を自動集計し、目がチカチカしないようにしました！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        
        staff_names = df_staff["スタッフ名"].dropna().tolist()
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        staff_off_days = df_staff["公休回数"].fillna(8).tolist() if "公休回数" in df_staff.columns else [8]*len(staff_names)
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
            
        st.success(f"✅ データの読み込み完了！シフトを計算します...")
        
        if st.button("シフトを自動作成する！（フェーズ4.2🔥）"):
            with st.spinner('AI店長がみんなの希望休と人数パズルを解いています...（最大20秒）'):
                
                model = cp_model.CpModel()
                shift_types = ['A', 'D', 'E', '公']
                
                shifts = {}
                for e in range(num_staff):
                    for d in range(num_days):
                        for s in shift_types:
                            shifts[(e, d, s)] = model.NewBoolVar(f'shift_{e}_{d}_{s}')
                            
                for e in range(num_staff):
                    for d in range(num_days):
                        model.AddExactlyOne(shifts[(e, d, s)] for s in shift_types)
                        
                for e in range(num_staff):
                    model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0:
                            model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days:
                            model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                    model.Add(sum(shifts[(e, d, 'A')] for e in range(num_staff)) >= day_req_list[d])

                    leadership_score = sum(
                        (2 if "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * shifts[(e, d, 'A')]
                        for e in range(num_staff)
                    )
                    model.Add(leadership_score >= 2)

                for e, staff_name in enumerate(staff_names):
                    for d, date_col in enumerate(date_columns):
                        if date_col in df_history.columns:
                            target_row = df_history[df_history["スタッフ名"] == staff_name]
                            if not target_row.empty:
                                cell_value = str(target_row[date_col].values[0]).strip()
                                if cell_value == "公":
                                    model.Add(shifts[(e, d, '公')] == 1)

                for e in range(num_staff):
                    target_off = int(staff_off_days[e])
                    model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == target_off)

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 20.0 
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨ 集計欄をご確認ください。")
                    
                    # 1. 基本のシフト表を作成
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e], "役割": staff_roles[e]}
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    row[date_columns[d]] = s
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)

                    # 🌟 2. 横の集計（個人の回数）を追加
                    result_df['日勤(A)回数'] = (result_df[date_columns] == 'A').sum(axis=1)
                    result_df['夜勤(D)回数'] = (result_df[date_columns] == 'D').sum(axis=1)
                    result_df['公休回数'] = (result_df[date_columns] == '公').sum(axis=1)

                    # 🌟 3. 下の集計（毎日の人数）を追加
                    summary_A = {"スタッフ名": "【日勤(A) 合計】", "役割": ""}
                    summary_D = {"スタッフ名": "【夜勤(D) 合計】", "役割": ""}
                    summary_Off = {"スタッフ名": "【公休 合計】", "役割": ""}
                    
                    # 右端の集計列は空欄にする
                    for col in ['日勤(A)回数', '夜勤(D)回数', '公休回数']:
                        summary_A[col] = ""
                        summary_D[col] = ""
                        summary_Off[col] = ""

                    # 日ごとにA, D, 公の数を数える
                    for col in date_columns:
                        summary_A[col] = (result_df[col] == 'A').sum()
                        summary_D[col] = (result_df[col] == 'D').sum()
                        summary_Off[col] = (result_df[col] == '公').sum()

                    # 表を合体させる
                    summary_df = pd.DataFrame([summary_A, summary_D, summary_Off])
                    final_df = pd.concat([result_df, summary_df], ignore_index=True)

                    # 画面に表示
                    st.dataframe(final_df)
                    
                    # エクセル出力
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成したシフト（集計付き）をダウンロード",
                        data=processed_data,
                        file_name="完成版_集計付きシフト.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 条件が厳しすぎて組めませんでした。（希望休が重なりすぎて人数が足りないなど）")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
