import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
from openpyxl.styles import PatternFill

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ9：カレンダー自動色付け版)")
st.write("土日祝の全自動色付けと、パート(P〇)の対応、AI相談機能を搭載した完成形です！")

# --- 妥協案のセッション管理 ---
if 'allow_4_days_work' not in st.session_state:
    st.session_state.allow_4_days_work = False
if 'allow_night_before_3_days' not in st.session_state:
    st.session_state.allow_night_before_3_days = False
if 'allow_sub_only' not in st.session_state:
    st.session_state.allow_sub_only = False

# 🗓️ 作成する年月の選択（祝日判定用）
st.write("---")
st.write("🗓️ **作成するシフトの「年」と「月」を選んでください（祝日の自動判定に使います）**")
today = datetime.date.today()
col_y, col_m = st.columns(2)
with col_y:
    target_year = st.selectbox("作成年", [today.year, today.year + 1], index=0)
with col_m:
    # 現在の月の「次の月」をデフォルトにする
    next_month = today.month + 1 if today.month < 12 else 1
    target_month = st.selectbox("作成月", list(range(1, 13)), index=next_month - 1)
st.write("---")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・前月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="日別設定")
        
        staff_names = df_staff["スタッフ名"].dropna().tolist()
        num_staff = len(staff_names)
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        staff_off_days = df_staff["公休数"].fillna(8).tolist()
        staff_night_ok = df_staff["夜勤可否"].fillna("〇").tolist()
        
        if "パート" in df_staff.columns:
            staff_part_shifts = df_staff["パート"].fillna("").astype(str).tolist()
        else:
            staff_part_shifts = [""] * num_staff
        
        staff_night_limits = []
        for i in range(num_staff):
            if staff_night_ok[i] == "×":
                staff_night_limits.append(0)
            else:
                val = df_staff["夜勤上限"].iloc[i]
                staff_night_limits.append(int(val) if pd.notna(val) else 10)
        
        staff_overtime_ok = df_staff["残業可否"].fillna("〇").tolist()
        staff_sun_d = []
        staff_sun_e = []
        for i in range(num_staff):
            if staff_night_ok[i] == "×":
                staff_sun_d.append("×")
                staff_sun_e.append("×")
            else:
                staff_sun_d.append(df_staff["日曜Dカウント"].fillna("〇").iloc[i])
                staff_sun_e.append(df_staff["日曜Eカウント"].fillna("〇").iloc[i])

        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        weekdays = df_req.iloc[0, 1:num_days+1].tolist()
        
        def get_req_row(label, default_val):
            row = df_req[df_req.iloc[:, 0] == label]
            if not row.empty:
                return [int(x) if pd.notna(x) else default_val for x in row.iloc[0, 1:num_days+1]]
            return [default_val] * num_days

        def get_str_row(label, default_val):
            row = df_req[df_req.iloc[:, 0] == label]
            if not row.empty:
                return [str(x).strip() if pd.notna(x) else default_val for x in row.iloc[0, 1:num_days+1]]
            return [default_val] * num_days

        day_req_list = get_req_row("日勤人数", 3)
        absolute_req_list = get_str_row("絶対確保", "")
        night_req_list = get_req_row("夜勤人数", 2)

        st.success(f"✅ データの読み込み完了！祝日判定用のカレンダーを準備しました！")
        
        # 💬 AIからのご相談エリア
        with st.expander("📩 AI店長への特別許可（シフトが組めない時だけ開いてください）", expanded=True):
            st.warning("👩‍💼 **AI店長からのご相談:**\n\n『申し訳ありません、現在の人数と希望休ではどうしてもシフトが組めません💦 もしよろしければ、今回だけ以下のルールのどれかを特別に許可していただけませんか？』")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.session_state.allow_4_days_work = st.checkbox("🙏 誰かが「最大4連勤」になることを許可する", value=st.session_state.allow_4_days_work)
            with col2:
                st.session_state.allow_night_before_3_days = st.checkbox("🙏 誰かの夜勤直前が「3日勤」になることを許可する", value=st.session_state.allow_night_before_3_days)
            with col3:
                st.session_state.allow_sub_only = st.checkbox("🙏 リーダー不在時、「サブ1名＋他」の配置を許可する", value=st.session_state.allow_sub_only)

        if st.button("シフトを自動作成する！（自動色付け版🔥）"):
            with st.spinner('AI店長がパズルを解き、カレンダーに色を塗っています...（最大45秒）'):
                
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
                    if staff_night_ok[e] == "×":
                        for d in range(num_days):
                            model.Add(shifts[(e, d, 'D')] == 0)
                            model.Add(shifts[(e, d, 'E')] == 0)

                for e, staff_name in enumerate(staff_names):
                    target_row = df_history[df_history.iloc[:, 0] == staff_name]
                    if not target_row.empty:
                        last_month_last_day = str(target_row.iloc[0, 5]).strip()
                        if last_month_last_day == "D":
                            model.Add(shifts[(e, 0, 'E')] == 1)
                            if num_days > 1:
                                model.Add(shifts[(e, 1, '公')] == 1)
                        elif last_month_last_day == "E":
                            model.Add(shifts[(e, 0, '公')] == 1)

                for e in range(num_staff):
                    if staff_night_ok[e] != "×":
                        target_row = df_history[df_history.iloc[:, 0] == staff_names[e]]
                        if not target_row.empty:
                            if str(target_row.iloc[0, 5]).strip() != "D":
                                model.Add(shifts[(e, 0, 'E')] == 0)
                        for d in range(num_days):
                            if d > 0:
                                model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                            if d + 1 < num_days:
                                model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                    
                    actual_day_staff = sum(
                        shifts[(e, d, 'A')] for e in range(num_staff) if "新人" not in str(staff_roles[e])
                    )
                    
                    if absolute_req_list[d] == "〇":
                        model.Add(actual_day_staff >= day_req_list[d])
                    else:
                        model.Add(actual_day_staff >= day_req_list[d] - 1)

                for d in range(num_days):
                    leadership_score = sum(
                        (2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * shifts[(e, d, 'A')]
                        for e in range(num_staff)
                    )
                    if st.session_state.allow_sub_only:
                        model.Add(leadership_score >= 1)
                    else:
                        model.Add(leadership_score >= 2)

                for e, staff_name in enumerate(staff_names):
                    target_row = df_history[df_history.iloc[:, 0] == staff_name]
                    if not target_row.empty:
                        for d in range(num_days):
                            col_idx = 6 + d
                            if col_idx < len(df_history.columns):
                                cell_value = str(target_row.iloc[0, col_idx]).strip()
                                if cell_value == "公":
                                    model.Add(shifts[(e, d, '公')] == 1)

                for e in range(num_staff):
                    model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                    if staff_night_ok[e] != "×":
                        model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

                for e in range(num_staff):
                    for d in range(num_days - 3):
                        model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <= 3)
                        
                        if st.session_state.allow_4_days_work:
                            if d < num_days - 4:
                                model.Add(shifts[(e, d, 'A')] + shifts[(e, d+1, 'A')] + shifts[(e, d+2, 'A')] + shifts[(e, d+3, 'A')] + shifts[(e, d+4, 'A')] <= 4)
                        else:
                            model.Add(shifts[(e, d, 'A')] + shifts[(e, d+1, 'A')] + shifts[(e, d+2, 'A')] + shifts[(e, d+3, 'A')] <= 3)

                        if st.session_state.allow_night_before_3_days == False:
                            model.AddImplication(shifts[(e, d+3, 'D')], shifts[(e, d, 'A')] + shifts[(e, d+1, 'A')] + shifts[(e, d+2, 'A')] <= 2)

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 45.0
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨")
                    
                    # 🌟 曜日と祝日の判定（列名を 1(月) のように結合）
                    new_date_columns = []
                    for d_val, w_val in zip(date_columns, weekdays):
                        try:
                            # 祝日判定
                            dt = datetime.date(target_year, target_month, int(d_val))
                            if jpholiday.is_holiday(dt):
                                new_date_columns.append(f"{d_val}({w_val}・祝)")
                            else:
                                new_date_columns.append(f"{d_val}({w_val})")
                        except ValueError:
                            # 31日が存在しない月などの安全対策
                            new_date_columns.append(f"{d_val}({w_val})")
                    
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e], "役割": staff_roles[e], "パート": staff_part_shifts[e]}
                        role_str = str(staff_roles[e])
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    # P〇の処理
                                    if s == 'A' and str(staff_part_shifts[e]).strip() not in ["", "nan"]:
                                        row[new_date_columns[d]] = str(staff_part_shifts[e]).strip()
                                    else:
                                        row[new_date_columns[d]] = s
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)

                    # 集計処理
                    result_df['日勤(A・P)回数'] = result_df[new_date_columns].apply(lambda x: x.str.contains('A|P|Ｐ', na=False)).sum(axis=1)
                    result_df['夜勤(D)回数'] = (result_df[new_date_columns] == 'D').sum(axis=1)
                    result_df['公休回数'] = (result_df[new_date_columns] == '公').sum(axis=1)
                    
                    sunday_d_counts = []
                    sunday_e_counts = []
                    for e in range(num_staff):
                        d_count = 0
                        e_count = 0
                        for d in range(num_days):
                            if str(weekdays[d]).strip() == "日":
                                col_name = new_date_columns[d]
                                if result_df.loc[e, col_name] == 'D' and staff_sun_d[e] == "〇":
                                    d_count += 1
                                if result_df.loc[e, col_name] == 'E' and staff_sun_e[e] == "〇":
                                    e_count += 1
                        sunday_d_counts.append(d_count)
                        sunday_e_counts.append(e_count)
                        
                    result_df['日曜D回数(〇のみ)'] = sunday_d_counts
                    result_df['日曜E回数(〇のみ)'] = sunday_e_counts

                    summary_A = {"スタッフ名": "【日勤(A・P) 合計】", "役割": "", "パート": ""}
                    summary_D = {"スタッフ名": "【夜勤(D) 合計】", "役割": "", "パート": ""}
                    summary_Off = {"スタッフ名": "【公休 合計】", "役割": "", "パート": ""}
                    
                    for col in ['日勤(A・P)回数', '夜勤(D)回数', '公休回数', '日曜D回数(〇のみ)', '日曜E回数(〇のみ)']:
                        summary_A[col] = ""
                        summary_D[col] = ""
                        summary_Off[col] = ""

                    for d, col in enumerate(new_date_columns):
                        a_count = 0
                        for e in range(num_staff):
                            val = str(result_df.loc[e, col])
                            if (val == 'A' or "P" in val or "Ｐ" in val) and "新人" not in str(staff_roles[e]):
                                a_count += 1
                        summary_A[col] = a_count
                        summary_D[col] = (result_df[col] == 'D').sum()
                        summary_Off[col] = (result_df[col] == '公').sum()

                    summary_df = pd.DataFrame([summary_A, summary_D, summary_Off])
                    final_df = pd.concat([result_df, summary_df], ignore_index=True)

                    # 🌟 画面上での色付け設定（Streamlit）
                    def highlight_cols(s):
                        if s.name is None:
                            return [''] * len(s)
                        col_name = str(s.name)
                        if "祝" in col_name or "(日)" in col_name:
                            return ['background-color: #FFE6E6'] * len(s) # 薄い赤
                        elif "(土)" in col_name:
                            return ['background-color: #E6F2FF'] * len(s) # 薄い青
                        else:
                            return [''] * len(s)

                    st.dataframe(final_df.style.apply(highlight_cols, axis=0))
                    
                    # 🌟 エクセル出力での色付け設定（openpyxl）
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='完成シフト')
                        worksheet = writer.sheets['完成シフト']
                        
                        fill_sat = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid") # 薄い青
                        fill_sun_hol = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # 薄い赤
                        
                        # 列名を見て、土日祝ならその列全体に色を塗る
                        for col_idx, col_name in enumerate(final_df.columns, 1):
                            col_str = str(col_name)
                            if "祝" in col_str or "(日)" in col_str:
                                for row in range(1, len(final_df) + 2):
                                    worksheet.cell(row=row, column=col_idx).fill = fill_sun_hol
                            elif "(土)" in col_str:
                                for row in range(1, len(final_df) + 2):
                                    worksheet.cell(row=row, column=col_idx).fill = fill_sat
                                    
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成した究極のシフトをダウンロード（色付き）",
                        data=processed_data,
                        file_name="完成版_実務完全版_色付き.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 【AI店長より】\n申し訳ありません、どうしても今の条件ではシフトが破綻してしまいます😭\n上の「📩AI店長への特別許可」を開いて、どれか1つでも許可のチェックを入れてから再度お試しください！")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: 部品（jpholiday）が入っていないか、エクセルの形式が間違っています。({e})")
