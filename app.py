import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
from openpyxl.styles import PatternFill

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ10.1：バグ修正＆厳格モード)")
st.write("AIの勝手な妥協を禁止し、イレギュラーが発生した箇所を「色」で警告します！")

# --- 妥協案のセッション管理 ---
if 'allow_day_minus_1' not in st.session_state:
    st.session_state.allow_day_minus_1 = False
if 'allow_4_days_work' not in st.session_state:
    st.session_state.allow_4_days_work = False
if 'allow_night_before_3_days' not in st.session_state:
    st.session_state.allow_night_before_3_days = False
if 'allow_sub_only' not in st.session_state:
    st.session_state.allow_sub_only = False

st.write("---")
st.write("🗓️ **作成するシフトの「年」と「月」を選んでください**")
today = datetime.date.today()
col_y, col_m = st.columns(2)
with col_y:
    target_year = st.selectbox("作成年", [today.year, today.year + 1], index=0)
with col_m:
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

        st.success(f"✅ データの読み込み完了！AIの勝手な妥協を禁止し、厳格モードで準備しました。")
        
        # 💬 AIからのご相談エリア
        with st.expander("📩 AI店長への特別許可（※エラーで組めない時だけチェックを入れてください）", expanded=True):
            st.warning("👩‍💼 **AI店長からのご相談:**\n\n『基本はルールを100%死守して時間をかけて考えます！でも、どうしても無理な場合だけ、以下の妥協を許可してください💦』")
            col1, col2 = st.columns(2)
            with col1:
                st.session_state.allow_day_minus_1 = st.checkbox("🙏 日勤人数の「マイナス1」を許可する（※絶対確保の日は除く）", value=st.session_state.allow_day_minus_1)
                st.session_state.allow_sub_only = st.checkbox("🙏 リーダー不在時、「サブ1名＋他」の配置を許可する", value=st.session_state.allow_sub_only)
            with col2:
                st.session_state.allow_4_days_work = st.checkbox("🙏 誰かが「最大4連勤」になることを許可する（※該当箇所は黄色で警告します）", value=st.session_state.allow_4_days_work)
                st.session_state.allow_night_before_3_days = st.checkbox("🙏 誰かの夜勤直前が「3日勤」になることを許可する（※該当箇所はオレンジ色で警告します）", value=st.session_state.allow_night_before_3_days)

        if st.button("時間をかけて完璧なシフトを考える（最大3分🔥）"):
            with st.spinner('AI店長がたっぷり時間を使って、完璧なパズルを計算しています...（最大3分かかります。お待ちください🙏）'):
                
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

                # 前月履歴の引き継ぎ
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

                # 夜勤セットのロック
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

                # 人数の確保
                for d in range(num_days):
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                    
                    actual_day_staff = sum(
                        shifts[(e, d, 'A')] for e in range(num_staff) if "新人" not in str(staff_roles[e])
                    )
                    
                    if absolute_req_list[d] == "〇":
                        model.Add(actual_day_staff >= day_req_list[d])
                    else:
                        if st.session_state.allow_day_minus_1:
                            model.Add(actual_day_staff >= day_req_list[d] - 1)
                        else:
                            model.Add(actual_day_staff >= day_req_list[d])

                # 役割配置
                for d in range(num_days):
                    leadership_score = sum(
                        (2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * shifts[(e, d, 'A')]
                        for e in range(num_staff)
                    )
                    if st.session_state.allow_sub_only:
                        model.Add(leadership_score >= 1)
                    else:
                        model.Add(leadership_score >= 2)

                # 希望休
                for e, staff_name in enumerate(staff_names):
                    target_row = df_history[df_history.iloc[:, 0] == staff_name]
                    if not target_row.empty:
                        for d in range(num_days):
                            col_idx = 6 + d
                            if col_idx < len(df_history.columns):
                                cell_value = str(target_row.iloc[0, col_idx]).strip()
                                if cell_value == "公":
                                    model.Add(shifts[(e, d, '公')] == 1)

                # 回数ノルマ
                for e in range(num_staff):
                    model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                    if staff_night_ok[e] != "×":
                        model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

                # 🌟 バグ修正箇所：連勤・連休制限（文法エラーを修正しました）
                for e in range(num_staff):
                    for d in range(num_days - 3):
                        model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <= 3)
                        
                        if st.session_state.allow_4_days_work:
                            if d < num_days - 4:
                                model.Add(shifts[(e, d, 'A')] + shifts[(e, d+1, 'A')] + shifts[(e, d+2, 'A')] + shifts[(e, d+3, 'A')] + shifts[(e, d+4, 'A')] <= 4)
                        else:
                            model.Add(shifts[(e, d, 'A')] + shifts[(e, d+1, 'A')] + shifts[(e, d+2, 'A')] + shifts[(e, d+3, 'A')] <= 3)

                        if st.session_state.allow_night_before_3_days == False:
                            # 【修正】OR-Toolsの正しい文法（OnlyEnforceIf）を使用
                            model.Add(shifts[(e, d, 'A')] + shifts[(e, d+1, 'A')] + shifts[(e, d+2, 'A')] <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

                solver = cp_model.CpSolver()
                # 最大3分（180秒）考える
                solver.parameters.max_time_in_seconds = 180.0
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨")
                    
                    new_date_columns = []
                    for d_val, w_val in zip(date_columns, weekdays):
                        try:
                            dt = datetime.date(target_year, target_month, int(d_val))
                            if jpholiday.is_holiday(dt):
                                new_date_columns.append(f"{d_val}({w_val}・祝)")
                            else:
                                new_date_columns.append(f"{d_val}({w_val})")
                        except ValueError:
                            new_date_columns.append(f"{d_val}({w_val})")
                    
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e], "役割": staff_roles[e], "パート": staff_part_shifts[e]}
                        role_str = str(staff_roles[e])
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    if s == 'A' and str(staff_part_shifts[e]).strip() not in ["", "nan"]:
                                        row[new_date_columns[d]] = str(staff_part_shifts[e]).strip()
                                    else:
                                        row[new_date_columns[d]] = s
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)

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

                    # 🌟 警告ハイライト関数
                    def highlight_warnings(df):
                        styles = pd.DataFrame('', index=df.index, columns=df.columns)
                        
                        # 1. 毎日の集計行のチェック（赤色）
                        for d, col_name in enumerate(new_date_columns):
                            actual_a = df.loc[len(staff_names), col_name]
                            target_a = day_req_list[d]
                            if actual_a != "" and actual_a < target_a:
                                styles.loc[len(staff_names), col_name] = 'background-color: #FFCCCC; color: red; font-weight: bold;'

                        # 2. 該当スタッフの4連勤、夜勤前3連勤のチェック
                        for e in range(num_staff):
                            for d in range(num_days):
                                def is_work(day_idx):
                                    if day_idx >= num_days: return False
                                    v = str(df.loc[e, new_date_columns[day_idx]])
                                    return v == 'A' or 'P' in v or 'Ｐ' in v or v == 'D' or v == 'E'

                                # 4連勤の検出（黄色）
                                if is_work(d) and is_work(d+1) and is_work(d+2) and is_work(d+3):
                                    styles.loc[e, new_date_columns[d]] = 'background-color: #FFFF99;'
                                    styles.loc[e, new_date_columns[d+1]] = 'background-color: #FFFF99;'
                                    styles.loc[e, new_date_columns[d+2]] = 'background-color: #FFFF99;'
                                    styles.loc[e, new_date_columns[d+3]] = 'background-color: #FFFF99;'

                                # 夜勤前3連勤（オレンジ）
                                if d + 3 < num_days:
                                    v1 = str(df.loc[e, new_date_columns[d]])
                                    v2 = str(df.loc[e, new_date_columns[d+1]])
                                    v3 = str(df.loc[e, new_date_columns[d+2]])
                                    v4 = str(df.loc[e, new_date_columns[d+3]])
                                    if (v1=='A' or 'P' in v1) and (v2=='A' or 'P' in v2) and (v3=='A' or 'P' in v3) and v4=='D':
                                        styles.loc[e, new_date_columns[d]] = 'background-color: #FFD580;'
                                        styles.loc[e, new_date_columns[d+1]] = 'background-color: #FFD580;'
                                        styles.loc[e, new_date_columns[d+2]] = 'background-color: #FFD580;'
                                        styles.loc[e, new_date_columns[d+3]] = 'background-color: #FFD580;'
                        return styles

                    st.dataframe(final_df.style.apply(highlight_warnings, axis=None))
                    
                    # エクセル出力での色付け設定
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_df.to_excel(writer, index=False, sheet_name='完成シフト')
                        worksheet = writer.sheets['完成シフト']
                        
                        fill_red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                        fill_yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                        fill_orange = PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid")
                        
                        for d, col_name in enumerate(new_date_columns):
                            actual_a = final_df.loc[len(staff_names), col_name]
                            if actual_a != "" and actual_a < day_req_list[d]:
                                worksheet.cell(row=len(staff_names)+2, column=d+4).fill = fill_red

                        for e in range(num_staff):
                            for d in range(num_days):
                                def is_work(day_idx):
                                    if day_idx >= num_days: return False
                                    v = str(final_df.loc[e, new_date_columns[day_idx]])
                                    return v == 'A' or 'P' in v or 'Ｐ' in v or v == 'D' or v == 'E'

                                if is_work(d) and is_work(d+1) and is_work(d+2) and is_work(d+3):
                                    worksheet.cell(row=e+2, column=d+4).fill = fill_yellow
                                    worksheet.cell(row=e+2, column=d+5).fill = fill_yellow
                                    worksheet.cell(row=e+2, column=d+6).fill = fill_yellow
                                    worksheet.cell(row=e+2, column=d+7).fill = fill_yellow

                                if d + 3 < num_days:
                                    v1 = str(final_df.loc[e, new_date_columns[d]])
                                    v2 = str(final_df.loc[e, new_date_columns[d+1]])
                                    v3 = str(final_df.loc[e, new_date_columns[d+2]])
                                    v4 = str(final_df.loc[e, new_date_columns[d+3]])
                                    if (v1=='A' or 'P' in v1) and (v2=='A' or 'P' in v2) and (v3=='A' or 'P' in v3) and v4=='D':
                                        worksheet.cell(row=e+2, column=d+4).fill = fill_orange
                                        worksheet.cell(row=e+2, column=d+5).fill = fill_orange
                                        worksheet.cell(row=e+2, column=d+6).fill = fill_orange
                                        worksheet.cell(row=e+2, column=d+7).fill = fill_orange
                                        
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 完成したシフトをダウンロード（警告色付き）",
                        data=processed_data,
                        file_name="完成版_厳格モード_警告付き.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 【AI店長より】\n3分間フルに考えましたが、どうしてもルール違反なしではシフトが組めません😭\n上の「📩AI店長への特別許可」を開いて、妥協を許可してから再度ボタンを押してください！")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白の行があります。({e})")
