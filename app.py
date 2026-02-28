import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
from openpyxl.styles import PatternFill
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ16：完全安定・バグ修正版)")
st.write("エラーを修正し、妥協優先度（1,2,3...）と残業の割合公平化を安全に実行します！")

if 'allow_day_minus_1' not in st.session_state:
    st.session_state.allow_day_minus_1 = False
if 'allow_4_days_work' not in st.session_state:
    st.session_state.allow_4_days_work = False
if 'allow_night_before_3_days' not in st.session_state:
    st.session_state.allow_night_before_3_days = False
if 'allow_sub_only' not in st.session_state:
    st.session_state.allow_sub_only = False
if 'allow_consecutive_overtime' not in st.session_state:
    st.session_state.allow_consecutive_overtime = False

st.write("---")
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
        
        # 🌟 安全なデータ取得（リストの長さエラーを絶対に防ぐ）
        staff_names = df_staff["スタッフ名"].dropna().tolist()
        num_staff = len(staff_names)
        
        def get_staff_col(col_name, default_val, is_int=False):
            res = []
            for i in range(num_staff):
                if col_name in df_staff.columns and pd.notna(df_staff[col_name].iloc[i]):
                    val = df_staff[col_name].iloc[i]
                    res.append(int(val) if is_int else str(val).strip())
                else:
                    res.append(default_val)
            return res

        staff_roles = get_staff_col("役割", "一般")
        staff_off_days = get_staff_col("公休数", 8, is_int=True)
        staff_night_ok = get_staff_col("夜勤可否", "〇")
        staff_overtime_ok = get_staff_col("残業可否", "〇")
        staff_part_shifts = get_staff_col("パート", "")
        
        staff_night_limits = []
        raw_limits = get_staff_col("夜勤上限", 10, is_int=True)
        for i in range(num_staff):
            staff_night_limits.append(0 if staff_night_ok[i] == "×" else raw_limits[i])

        raw_sun_d = get_staff_col("日曜Dカウント", "〇")
        raw_sun_e = get_staff_col("日曜Eカウント", "〇")
        staff_sun_d = ["×" if staff_night_ok[i] == "×" else raw_sun_d[i] for i in range(num_staff)]
        staff_sun_e = ["×" if staff_night_ok[i] == "×" else raw_sun_e[i] for i in range(num_staff)]

        # 妥協優先度の取得
        staff_comp_lvl = []
        for i in range(num_staff):
            val = ""
            if "妥協優先度" in df_staff.columns and pd.notna(df_staff["妥協優先度"].iloc[i]):
                val = str(df_staff["妥協優先度"].iloc[i]).strip()
            elif "連勤妥協OK" in df_staff.columns and pd.notna(df_staff["連勤妥協OK"].iloc[i]):
                val = str(df_staff["連勤妥協OK"].iloc[i]).strip()
            
            if val in ["〇", "1", "1.0"]: staff_comp_lvl.append(1)
            elif val in ["2", "2.0"]: staff_comp_lvl.append(2)
            elif val in ["3", "3.0"]: staff_comp_lvl.append(3)
            else: staff_comp_lvl.append(0)

        # カレンダーの取得
        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        def get_req_col(label, default_val, is_int=True):
            row = df_req[df_req.iloc[:, 0] == label]
            res = []
            for d in range(num_days):
                if not row.empty and (d + 1) < len(df_req.columns):
                    val = row.iloc[0, d + 1]
                    if pd.notna(val):
                        res.append(int(val) if is_int else str(val).strip())
                        continue
                res.append(default_val)
            return res

        day_req_list = get_req_col("日勤人数", 3)
        night_req_list = get_req_col("夜勤人数", 2)
        overtime_req_list = get_req_col("残業人数", 0)
        absolute_req_list = get_req_col("絶対確保", "", is_int=False)

        weekdays = []
        for d in range(num_days):
            if (d + 1) < len(df_req.columns):
                val = df_req.iloc[0, d + 1]
                weekdays.append(str(val).strip() if pd.notna(val) else "")
            else:
                weekdays.append("")

        st.success("✅ データの読み込みに成功しました！（エラー修正済みです）")
        
        with st.expander("⚙️ 【高度な設定】条件緩和ルールの優先順位（※エラーで作成できない場合のみ設定）", expanded=True):
            st.info("シフトが組めない場合、AIは以下の「優先順位 1」の項目から順番に条件を緩和（妥協）して再計算します。")
            options = ["許可しない（絶対死守）", "優先順位 1（最初に妥協）", "優先順位 2", "優先順位 3（最終手段）"]
            col1, col2 = st.columns(2)
            with col1:
                st.write("**■ 人数と役割の緩和**")
                opt_minus_1 = st.selectbox("日勤人数の「マイナス1」許容", options, index=0)
                opt_sub_only = st.selectbox("役割配置「サブ1名のみ」の許容", options, index=0)
            with col2:
                st.write("**■ 連続勤務の緩和（※エクセルの妥協優先度に沿って適用）**")
                opt_4_days = st.selectbox("対象者の「最大4連勤」許容", options, index=0)
                opt_night_3 = st.selectbox("対象者の「夜勤前3日勤」許容", options, index=0)
                opt_ot_consec = st.selectbox("やむを得ない「残業(A残)2日連続」の許容", options, index=0)

        def get_penalty_weight(opt_str):
            if "許可しない" in opt_str: return -1
            elif "優先順位 1" in opt_str: return 100
            elif "優先順位 2" in opt_str: return 1000
            elif "優先順位 3" in opt_str: return 10000
            return -1

        def solve_shift(random_seed):
            model = cp_model.CpModel()
            shift_types = ['A', 'A残', 'D', 'E', '公']
            
            shifts = {}
            for e in range(num_staff):
                for d in range(num_days):
                    for s in shift_types:
                        shifts[(e, d, s)] = model.NewBoolVar(f'shift_{e}_{d}_{s}')
                        
            model.AddHint(shifts[(0, 0, 'A')], random.choice([0, 1]))

            for e in range(num_staff):
                for d in range(num_days):
                    model.AddExactlyOne(shifts[(e, d, s)] for s in shift_types)
                    
            for e in range(num_staff):
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 0)
                        model.Add(shifts[(e, d, 'E')] == 0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'A残')] == 0)

            # 前月履歴の読み込み（安全処理付き）
            for e, staff_name in enumerate(staff_names):
                target_row = df_history[df_history.iloc[:, 0] == staff_name]
                if not target_row.empty:
                    last_month_last_day = str(target_row.iloc[0, 5]).strip() if target_row.shape[1] > 5 else ""
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
                        l_day = str(target_row.iloc[0, 5]).strip() if target_row.shape[1] > 5 else ""
                        if l_day != "D":
                            model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0:
                            model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days:
                            model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

            for e in range(num_staff):
                for d in range(num_days - 6):
                    model.Add(shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(e, d+6, 'D')] <= 2)

            penalties = []
            
            # 人数確保
            w_minus_1 = get_penalty_weight(opt_minus_1)
            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                actual_day_staff = sum((shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                
                if absolute_req_list[d] == "〇" or w_minus_1 == -1:
                    model.Add(actual_day_staff >= day_req_list[d])
                else:
                    model.Add(actual_day_staff >= day_req_list[d] - 1)
                    minus_var = model.NewBoolVar('')
                    model.Add(actual_day_staff == day_req_list[d] - 1).OnlyEnforceIf(minus_var)
                    penalties.append(minus_var * w_minus_1)

            # 役割配置
            w_sub_only = get_penalty_weight(opt_sub_only)
            for d in range(num_days):
                leadership_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if w_sub_only == -1:
                    model.Add(leadership_score >= 2)
                else:
                    model.Add(leadership_score >= 1)
                    sub_var = model.NewBoolVar('')
                    model.Add(leadership_score == 1).OnlyEnforceIf(sub_var)
                    penalties.append(sub_var * w_sub_only)

            # 希望休とノルマ
            for e, staff_name in enumerate(staff_names):
                target_row = df_history[df_history.iloc[:, 0] == staff_name]
                if not target_row.empty:
                    for d in range(num_days):
                        col_idx = 6 + d
                        if col_idx < target_row.shape[1]:
                            cell_value = str(target_row.iloc[0, col_idx]).strip()
                            if cell_value == "公":
                                model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×":
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

            # 連勤・連休制限（優先順位付き）
            w_4_days = get_penalty_weight(opt_4_days)
            w_night_3 = get_penalty_weight(opt_night_3)
            
            for e in range(num_staff):
                target_weight = staff_comp_lvl[e]
                
                for d in range(num_days - 3):
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <= 3)
                    
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                        
                    # 4連勤チェック
                    if w_4_days != -1 and target_weight > 0:
                        if d < num_days - 4:
                            model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf(p_var)
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * (w_4_days * target_weight))
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3)

                    # 夜勤前3日勤チェック
                    if w_night_3 != -1 and target_weight > 0:
                        np_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Not())
                        
                        final_p = model.NewIntVar(0, w_night_3 * target_weight, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, d+3, 'D')]])
                        penalties.append(final_p)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            # 残業連続制限
            w_ot_consec = get_penalty_weight(opt_ot_consec)
            for e in range(num_staff):
                for d in range(num_days - 1):
                    if w_ot_consec == -1:
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(ot_var * w_ot_consec)

            # 残業割合の公平化
            total_ot_req = sum(overtime_req_list)
            total_day_req = sum(day_req_list) 
            if total_ot_req > 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        actual_days_worked = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        actual_ot = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        
                        ideal_ot_scaled = actual_days_worked * total_ot_req
                        actual_ot_scaled = actual_ot * total_day_req
                        
                        diff = model.NewIntVar(-10000, 10000, f'diff_{e}')
                        abs_diff = model.NewIntVar(0, 10000, f'abs_diff_{e}')
                        
                        model.Add(diff == actual_ot_scaled - ideal_ot_scaled)
                        model.AddAbsEquality(abs_diff, diff)
                        penalties.append(abs_diff)
            
            if penalties:
                model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0
            solver.parameters.random_seed = random_seed
            status = solver.Solve(model)
            
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                return solver, shifts
            else:
                return None, None


        if st.button("設定に基づき、シフトを【3パターン】作成する！（最大3分🔥）"):
            with st.spinner('AIが優先順位と割合を計算し、3パターンのシフトを考えています...（最大3分かかります）'):
                
                results = []
                for seed in [1, 42, 99]:
                    solver, shifts = solve_shift(seed)
                    if solver:
                        results.append((solver, shifts))

                if not results:
                    st.error("❌ 【AI店長より】申し訳ありません、どうしてもシフトが組めません😭 緩和条件の「優先順位」を選択してください！")
                else:
                    st.success(f"✨シフトが完成しました！ {len(results)}パターンのご提案があります！✨")
                    
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

                    tab_names = [f"提案パターン {i+1}" for i in range(len(results))]
                    tabs = st.tabs(tab_names)
                    
                    for i, (solver, shifts) in enumerate(results):
                        with tabs[i]:
                            shift_types = ['A', 'A残', 'D', 'E', '公']
                            result_data = []
                            for e in range(num_staff):
                                row = {"スタッフ名": staff_names[e], "役割": staff_roles[e], "パート": staff_part_shifts[e]}
                                for d in range(num_days):
                                    for s in shift_types:
                                        if solver.Value(shifts[(e, d, s)]) == 1:
                                            if (s == 'A' or s == 'A残') and str(staff_part_shifts[e]).strip() not in ["", "nan"]:
                                                row[new_date_columns[d]] = str(staff_part_shifts[e]).strip()
                                            else:
                                                row[new_date_columns[d]] = s
                                result_data.append(row)
                                
                            result_df = pd.DataFrame(result_data)

                            result_df['日勤(A・P)回数'] = result_df[new_date_columns].apply(lambda x: x.str.contains('A|P|Ｐ', na=False) & ~x.str.contains('残', na=False)).sum(axis=1)
                            result_df['残業(A残)回数'] = (result_df[new_date_columns] == 'A残').sum(axis=1)
                            
                            def calc_ratio(row):
                                if row['日勤(A・P)回数'] > 0:
                                    return f"{(row['残業(A残)回数'] / row['日勤(A・P)回数']) * 100:.1f}%"
                                return "0.0%"
                            
                            result_df['残業割合'] = result_df.apply(calc_ratio, axis=1)

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
                            summary_A_zan = {"スタッフ名": "【残業(A残) 合計】", "役割": "", "パート": ""}
                            summary_D = {"スタッフ名": "【夜勤(D) 合計】", "役割": "", "パート": ""}
                            summary_Off = {"スタッフ名": "【公休 合計】", "役割": "", "パート": ""}
                            
                            for col in ['日勤(A・P)回数', '残業(A残)回数', '残業割合', '夜勤(D)回数', '公休回数', '日曜D回数(〇のみ)', '日曜E回数(〇のみ)']:
                                summary_A[col] = ""
                                summary_A_zan[col] = ""
                                summary_D[col] = ""
                                summary_Off[col] = ""

                            for d, col in enumerate(new_date_columns):
                                a_count = 0
                                for e in range(num_staff):
                                    val = str(result_df.loc[e, col])
                                    if (val == 'A' or val == 'A残' or "P" in val or "Ｐ" in val) and "新人" not in str(staff_roles[e]):
                                        a_count += 1
                                summary_A[col] = a_count
                                summary_A_zan[col] = (result_df[col] == 'A残').sum()
                                summary_D[col] = (result_df[col] == 'D').sum()
                                summary_Off[col] = (result_df[col] == '公').sum()

                            summary_df = pd.DataFrame([summary_A, summary_A_zan, summary_D, summary_Off])
                            final_df = pd.concat([result_df, summary_df], ignore_index=True)

                            def highlight_warnings(df):
                                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                                
                                for d, col_name in enumerate(new_date_columns):
                                    actual_a = df.loc[len(staff_names), col_name]
                                    target_a = day_req_list[d]
                                    if actual_a != "":
                                        if actual_a < target_a:
                                            styles.loc[len(staff_names), col_name] = 'background-color: #FFCCCC; color: red; font-weight: bold;'
                                        elif actual_a > target_a:
                                            styles.loc[len(staff_names), col_name] = 'background-color: #CCFFFF; color: blue; font-weight: bold;'

                                for e in range(num_staff):
                                    for d in range(num_days):
                                        def is_work(day_idx):
                                            if day_idx >= num_days: return False
                                            v = str(df.loc[e, new_date_columns[day_idx]])
                                            return v == 'A' or v == 'A残' or 'P' in v or 'Ｐ' in v or v == 'D' or v == 'E'

                                        if is_work(d) and is_work(d+1) and is_work(d+2) and is_work(d+3):
                                            styles.loc[e, new_date_columns[d]] = 'background-color: #FFFF99;'
                                            styles.loc[e, new_date_columns[d+1]] = 'background-color: #FFFF99;'
                                            styles.loc[e, new_date_columns[d+2]] = 'background-color: #FFFF99;'
                                            styles.loc[e, new_date_columns[d+3]] = 'background-color: #FFFF99;'

                                        if d + 3 < num_days:
                                            v1 = str(df.loc[e, new_date_columns[d]])
                                            v2 = str(df.loc[e, new_date_columns[d+1]])
                                            v3 = str(df.loc[e, new_date_columns[d+2]])
                                            v4 = str(df.loc[e, new_date_columns[d+3]])
                                            
                                            v1_is_a = (v1=='A' or v1=='A残' or 'P' in v1 or 'Ｐ' in v1)
                                            v2_is_a = (v2=='A' or v2=='A残' or 'P' in v2 or 'Ｐ' in v2)
                                            v3_is_a = (v3=='A' or v3=='A残' or 'P' in v3 or 'Ｐ' in v3)
                                            
                                            if v1_is_a and v2_is_a and v3_is_a and v4=='D':
                                                styles.loc[e, new_date_columns[d]] = 'background-color: #FFD580;'
                                                styles.loc[e, new_date_columns[d+1]] = 'background-color: #FFD580;'
                                                styles.loc[e, new_date_columns[d+2]] = 'background-color: #FFD580;'
                                                styles.loc[e, new_date_columns[d+3]] = 'background-color: #FFD580;'
                                return styles

                            st.dataframe(final_df.style.apply(highlight_warnings, axis=None))
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                final_df.to_excel(writer, index=False, sheet_name='完成シフト')
                            processed_data = output.getvalue()
                            
                            st.download_button(
                                label=f"📥 【パターン {i+1}】 をエクセルでダウンロード（色なし）",
                                data=processed_data,
                                file_name=f"完成版_パターン{i+1}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_btn_{i}"
                            )
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白の行があります。({e})")
