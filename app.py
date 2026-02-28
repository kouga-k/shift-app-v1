import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ18：偏り防止＆厳格人数管理)")
st.write("「緩和」は本当に組めない時の最終手段とし、夜勤・残業が月内で偏らないように調整します。")

if 'allow_day_minus_1' not in st.session_state: st.session_state.allow_day_minus_1 = False
if 'allow_4_days_work' not in st.session_state: st.session_state.allow_4_days_work = False
if 'allow_night_before_3_days' not in st.session_state: st.session_state.allow_night_before_3_days = False
if 'allow_sub_only' not in st.session_state: st.session_state.allow_sub_only = False
if 'allow_consecutive_overtime' not in st.session_state: st.session_state.allow_consecutive_overtime = False

st.write("---")
today = datetime.date.today()
col_y, col_m = st.columns(2)
with col_y: target_year = st.selectbox("作成年", [today.year, today.year + 1], index=0)
with col_m: target_month = st.selectbox("作成月", list(range(1, 13)), index=(today.month % 12))
st.write("---")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・前月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="日別設定")
        
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

        st.success("✅ データの読み込み完了！")
        
        with st.expander("⚙️ 【高度な設定】緩和ルールの優先順位（※どうしても組めない時だけ設定）", expanded=True):
            st.info("※「緩和」は本当にどうしても組めない時の【最終手段】としてのみAIが使用します。勝手な乱用はしません。")
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
            types = ['A', 'A残', 'D', 'E', '公']
            shifts = {(e, d, s): model.NewBoolVar('') for e in range(num_staff) for d in range(num_days) for s in types}
            model.AddHint(shifts[(0, 0, 'A')], random.choice([0, 1]))

            for e in range(num_staff):
                for d in range(num_days):
                    model.AddExactlyOne(shifts[(e, d, s)] for s in types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 0); model.Add(shifts[(e, d, 'E')] == 0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'A残')] == 0)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    last_month_last_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_month_last_day == "D":
                        model.Add(shifts[(e, 0, 'E')] == 1)
                        if num_days > 1:
                            model.Add(shifts[(e, 1, '公')] == 1)
                    elif last_month_last_day == "E":
                        model.Add(shifts[(e, 0, '公')] == 1)

            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
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
            
            # 🌟 人数確保の厳格化（日曜ルールと勝手な+1の制限）
            w_minus_1 = get_penalty_weight(opt_minus_1)
            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum((shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                req = day_req_list[d]
                is_sun = ('日' in weekdays[d])
                is_abs = (absolute_req_list[d] == "〇")

                if is_sun:
                    # 日曜日は「+1(過剰)」を絶対に許さない
                    model.Add(act_day <= req)
                    if is_abs or w_minus_1 == -1:
                        model.Add(act_day == req) # 緩和不可ならピッタリ
                    else:
                        model.Add(act_day >= req - 1)
                        minus_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(minus_var)
                        model.Add(act_day != req - 1).OnlyEnforceIf(minus_var.Not())
                        penalties.append(minus_var * w_minus_1 * 100) # ペナルティを100倍にして最終手段化
                else:
                    # 平日は「+1(過剰)」までは許容
                    model.Add(act_day <= req + 1)
                    if is_abs or w_minus_1 == -1:
                        model.Add(act_day >= req) # 緩和不可なら絶対に不足させない
                    else:
                        model.Add(act_day >= req - 1)
                        minus_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(minus_var)
                        model.Add(act_day != req - 1).OnlyEnforceIf(minus_var.Not())
                        penalties.append(minus_var * w_minus_1 * 100)

            w_sub_only = get_penalty_weight(opt_sub_only)
            for d in range(num_days):
                leadership_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if w_sub_only == -1:
                    model.Add(leadership_score >= 2)
                else:
                    model.Add(leadership_score >= 1)
                    sub_var = model.NewBoolVar('')
                    model.Add(leadership_score == 1).OnlyEnforceIf(sub_var)
                    penalties.append(sub_var * w_sub_only * 100)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    for d in range(num_days):
                        col_idx = 6 + d
                        if col_idx < tr.shape[1]:
                            cell_value = str(tr.iloc[0, col_idx]).strip()
                            if cell_value == "公":
                                model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×":
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

            w_4_days = get_penalty_weight(opt_4_days)
            w_night_3 = get_penalty_weight(opt_night_3)
            
            for e in range(num_staff):
                target_weight = staff_comp_lvl[e]
                for d in range(num_days - 3):
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <= 3)
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                        
                    if w_4_days != -1 and target_weight > 0:
                        if d < num_days - 4:
                            model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf(p_var)
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * w_4_days * target_weight * 100)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3)

                    if w_night_3 != -1 and target_weight > 0:
                        np_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Not())
                        final_p = model.NewIntVar(0, w_night_3 * target_weight * 100, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, d+3, 'D')]])
                        penalties.append(final_p)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            w_ot_consec = get_penalty_weight(opt_ot_consec)
            for e in range(num_staff):
                for d in range(num_days - 1):
                    if w_ot_consec == -1:
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(ot_var * w_ot_consec * 100)

            # 🌟 NEW: 月内での配置バランス（前後半の偏り防止）
            mid_day = num_days // 2
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    d_first = sum(shifts[(e, d, 'D')] for d in range(mid_day))
                    d_second = sum(shifts[(e, d, 'D')] for d in range(mid_day, num_days))
                    diff_d = model.NewIntVar(-100, 100, '')
                    abs_diff_d = model.NewIntVar(0, 100, '')
                    model.Add(diff_d == d_first - d_second)
                    model.AddAbsEquality(abs_diff_d, diff_d)
                    penalties.append(abs_diff_d * 50) # 偏りにペナルティ
                
                if staff_overtime_ok[e] != "×":
                    ot_first = sum(shifts[(e, d, 'A残')] for d in range(mid_day))
                    ot_second = sum(shifts[(e, d, 'A残')] for d in range(mid_day, num_days))
                    diff_ot = model.NewIntVar(-100, 100, '')
                    abs_diff_ot = model.NewIntVar(0, 100, '')
                    model.Add(diff_ot == ot_first - ot_second)
                    model.AddAbsEquality(abs_diff_ot, diff_ot)
                    penalties.append(abs_diff_ot * 50)

            # 夜勤回数と残業割合の公平化
            total_night_req = sum(night_req_list)
            night_staff_count = sum(1 for ok in staff_night_ok if ok != "×")
            if total_night_req > 0 and night_staff_count > 0:
                for e in range(num_staff):
                    if staff_night_ok[e] != "×":
                        act_n = sum(shifts[(e, d, 'D')] for d in range(num_days))
                        diff_n = model.NewIntVar(-10000, 10000, '')
                        abs_diff_n = model.NewIntVar(0, 10000, '')
                        model.Add(diff_n == (act_n * night_staff_count) - total_night_req)
                        model.AddAbsEquality(abs_diff_n, diff_n)
                        penalties.append(abs_diff_n)

            total_ot_req = sum(overtime_req_list); total_day_req = sum(day_req_list) 
            if total_ot_req > 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        act_d = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        act_o = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        diff = model.NewIntVar(-10000, 10000, '')
                        abs_diff = model.NewIntVar(0, 10000, '')
                        model.Add(diff == (act_o * total_day_req) - (act_d * total_ot_req))
                        model.AddAbsEquality(abs_diff, diff)
                        penalties.append(abs_diff)
            
            if penalties: model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0
            solver.parameters.random_seed = random_seed
            return (solver, shifts) if solver.Solve(model) in [cp_model.OPTIMAL, cp_model.FEASIBLE] else (None, None)


        if st.button("設定に基づき、シフトを【3パターン】作成する！"):
            with st.spinner('AIが優先順位とバランスを計算し、3パターンのシフトを考えています...（最大3分）'):
                results = [res for seed in [1, 42, 99] if (res := solve_shift(seed))[0]]
                if not results: st.error("❌ 条件が厳しすぎます。設定画面で緩和する条件の「優先順位」を選択してください！")
                else:
                    st.success(f"✨完成！ {len(results)}パターン提案します！✨")
                    cols = []
                    for d_val, w_val in zip(date_columns, weekdays):
                        try:
                            dt = datetime.date(target_year, target_month, int(d_val))
                            if jpholiday.is_holiday(dt): cols.append(f"{d_val}({w_val}・祝)")
                            else: cols.append(f"{d_val}({w_val})")
                        except ValueError:
                            cols.append(f"{d_val}({w_val})")

                    tabs = st.tabs([f"パターン {i+1}" for i in range(len(results))])
                    
                    for i, (solver, shifts) in enumerate(results):
                        with tabs[i]:
                            data = []
                            for e in range(num_staff):
                                row = {"スタッフ名": staff_names[e]}
                                for d in range(num_days):
                                    for s in ['A', 'A残', 'D', 'E', '公']:
                                        if solver.Value(shifts[(e, d, s)]):
                                            if (s == 'A' or s == 'A残') and str(staff_part_shifts[e]).strip() not in ["", "nan"]:
                                                row[cols[d]] = str(staff_part_shifts[e]).strip()
                                            else:
                                                row[cols[d]] = s
                                data.append(row)
                                
                            df_res = pd.DataFrame(data)

                            # 集計欄
                            sum_A = {"スタッフ名": "【日勤(A・P) 合計】"}
                            for c in cols: sum_A[c] = ""

                            for d, c in enumerate(cols):
                                a_count = 0
                                for e in range(num_staff):
                                    val = str(df_res.loc[e, c])
                                    if (val == 'A' or val == 'A残' or "P" in val or "Ｐ" in val) and "新人" not in str(staff_roles[e]):
                                        a_count += 1
                                sum_A[c] = a_count

                            df_fin = pd.concat([df_res, pd.DataFrame([sum_A])], ignore_index=True)

                            # 色塗り関数
                            def highlight_warnings(df):
                                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                                for d, col_name in enumerate(cols):
                                    actual_a = df.loc[len(staff_names), col_name]
                                    target_a = day_req_list[d]
                                    if actual_a != "":
                                        if actual_a < target_a:
                                            styles.loc[len(staff_names), col_name] = 'background-color: #FFCCCC; color: red; font-weight: bold;'
                                        elif actual_a > target_a:
                                            styles.loc[len(staff_names), col_name] = 'background-color: #CCFFFF; color: blue; font-weight: bold;'

                                for e in range(num_staff):
                                    for d in range(num_days):
                                        def is_day_work(day_idx):
                                            if day_idx >= num_days: return False
                                            v = str(df.loc[e, cols[day_idx]])
                                            return v == 'A' or v == 'A残' or 'P' in v or 'Ｐ' in v

                                        if is_day_work(d) and is_day_work(d+1) and is_day_work(d+2) and is_day_work(d+3):
                                            styles.loc[e, cols[d]] = 'background-color: #FFFF99;'
                                            styles.loc[e, cols[d+1]] = 'background-color: #FFFF99;'
                                            styles.loc[e, cols[d+2]] = 'background-color: #FFFF99;'
                                            styles.loc[e, cols[d+3]] = 'background-color: #FFFF99;'

                                        if d + 3 < num_days:
                                            if is_day_work(d) and is_day_work(d+1) and is_day_work(d+2) and str(df.loc[e, cols[d+3]]) == 'D':
                                                styles.loc[e, cols[d]] = 'background-color: #FFD580;'
                                                styles.loc[e, cols[d+1]] = 'background-color: #FFD580;'
                                                styles.loc[e, cols[d+2]] = 'background-color: #FFD580;'
                                                styles.loc[e, cols[d+3]] = 'background-color: #FFD580;'
                                return styles

                            st.dataframe(df_fin.style.apply(highlight_warnings, axis=None))
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df_fin.to_excel(writer, index=False, sheet_name='完成シフト')
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
