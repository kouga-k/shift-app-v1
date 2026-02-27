import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ4.1：ズレ防止・完全一致版)")
st.write("希望休や人数の「ズレ」を修正し、正確にシフトを組みます！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        
        # --- 1. データの安全な読み込み（ズレ防止） ---
        staff_names = df_staff["スタッフ名"].dropna().tolist()
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        staff_off_days = df_staff["公休回数"].fillna(8).tolist() if "公休回数" in df_staff.columns else [8]*len(staff_names)
        num_staff = len(staff_names)
        
        # 「必要人数設定」シートのB列以降（日付）をカレンダーの基準にする
        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        # 毎日の必要人数を「日付」と一致させて取得する
        night_req_row = df_req[df_req.iloc[:, 0] == "夜勤人数"]
        day_req_row = df_req[df_req.iloc[:, 0] == "日勤人数"]
        
        night_req_list = []
        day_req_list = []
        for col in date_columns:
            # 夜勤人数
            if not night_req_row.empty and col in night_req_row.columns:
                val = night_req_row[col].values[0]
                night_req_list.append(int(val) if pd.notna(val) else 2)
            else:
                night_req_list.append(2)
            # 日勤人数
            if not day_req_row.empty and col in day_req_row.columns:
                val = day_req_row[col].values[0]
                day_req_list.append(int(val) if pd.notna(val) else 3)
            else:
                day_req_list.append(3)
            
        st.success(f"✅ {num_staff}名のスタッフと、{num_days}日分のカレンダーを正確に認識しました！")
        
        if st.button("シフトを自動作成する！（フェーズ4.1🔥）"):
            with st.spinner('AI店長がみんなの希望休と人数パズルを解いています...'):
                
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

                # ルール4: 毎日の「日勤(A)」の必要人数（指定人数"以上"）
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'A')] for e in range(num_staff)) >= day_req_list[d])

                # ルール5: リーダー配置（日勤にリーダー1名orサブ2名）
                for d in range(num_days):
                    leadership_score = sum(
                        (2 if "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * shifts[(e, d, 'A')]
                        for e in range(num_staff)
                    )
                    model.Add(leadership_score >= 2)

                # 🌟 ルール6: 希望休の「完全ピンポイント検索（VLOOKUP方式）」
                for e, staff_name in enumerate(staff_names):
                    for d, date_col in enumerate(date_columns):
                        # 希望休シートにこの日付（例: 1, 2, 3...）の列があるか確認
                        if date_col in df_history.columns:
                            # スタッフ名を検索して行を特定
                            target_row = df_history[df_history["スタッフ名"] == staff_name]
                            if not target_row.empty:
                                cell_value = str(target_row[date_col].values[0]).strip()
                                if cell_value == "公":
                                    # 見つけたら絶対に休みにする
                                    model.Add(shifts[(e, d, '公')] == 1)

                # ルール7: 公休回数のノルマ
                for e in range(num_staff):
                    target_off = int(staff_off_days[e])
                    model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == target_off)

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 20.0 
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨ 希望休も人数もズレなく反映されています！")
                    
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
                        file_name="完成版_ズレ修正版.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 条件が厳しすぎて組めませんでした。（希望休が重なりすぎて人数が足りないなど）")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
