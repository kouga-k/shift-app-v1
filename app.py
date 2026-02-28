import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
from openpyxl.styles import PatternFill
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ")
st.write("複雑な条件と公平性を考慮し、最適なシフトを3パターン提案します。")

# セッション管理
for key in ['allow_day_minus_1', 'allow_4_days_work', 'allow_night_before_3_days', 'allow_sub_only', 'allow_consecutive_overtime']:
    if key not in st.session_state: st.session_state[key] = False

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
        staff_roles = df_staff["役割"].fillna("一般").tolist()
        staff_off_days = df_staff["公休数"].fillna(8).tolist()
        staff_night_ok = df_staff["夜勤可否"].fillna("〇").tolist()
        staff_overtime_ok = df_staff["残業可否"].fillna("〇").tolist()
        staff_part_shifts = df_staff["パート"].fillna("").astype(str).tolist() if "パート" in df_staff.columns else [""] * num_staff
        
        staff_comp_lvl = []
        comp_col = df_staff.get("妥協優先度", df_staff.get("連勤妥協OK", pd.Series([""] * num_staff)))
        for val in comp_col:
            v = str(val).strip()
            if v in ["〇", "1", "1.0"]: staff_comp_lvl.append(1)
            elif v in ["2", "2.0"]: staff_comp_lvl.append(2)
            elif v in ["3", "3.0"]: staff_comp_lvl.append(3)
            else: staff_comp_lvl.append(0)
        
        staff_night_limits = [0 if ok == "×" else int(v) if pd.notna(v) else 10 for ok, v in zip(staff_night_ok, df_staff.get("夜勤上限", pd.Series([10]*num_staff)))]
        staff_sun_d = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, df_staff.get("日曜Dカウント", pd.Series(["〇"]*num_staff)).fillna("〇"))]
        staff_sun_e = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, df_staff.get("日曜Eカウント", pd.Series(["〇"]*num_staff)).fillna("〇"))]

        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        weekdays = df_req.iloc[0, 1:num_days+1].tolist()
        
        def get_row(label, d_val, is_int=True):
            r = df_req[df_req.iloc[:, 0] == label]
            if not r.empty: return [int(x) if pd.notna(x) else d_val for x in r.iloc[0, 1:num_days+1]] if is_int else [str(x).strip() if pd.notna(x) else d_val for x in r.iloc[0, 1:num_days+1]]
            return [d_val] * num_days

        day_req_list = get_row("日勤人数", 3)
        night_req_list = get_row("夜勤人数", 2)
        overtime_req_list = get_row("残業人数", 0)
        absolute_req_list = get_row("絶対確保", "", is_int=False)

        st.success("✅ データ読み込み完了！")
        
        with st.expander("⚙️ 【高度な設定】条件緩和ルールの優先順位（※エラーで作成できない場合のみ設定）", expanded=True):
            st.info("シフトが組めない場合、AIは以下の「優先順位 1」の項目から順番に条件を緩和（妥協）して再計算します。")
            options = ["許可しない（絶対死守）", "優先順位 1（最初に妥協）", "優先順位 2", "優先順位 3（最終手段）"]
            col1, col2 = st.columns(2)
            with col1:
                st.write("**■ 人数と役割の緩和**")
                opt_minus_1 = st.selectbox("日勤人数の「マイナス1」許容", options, index=0)
                opt_sub_only = st.selectbox("役割配置「サブ1名のみ」の許容", options, index=0)
            with col2:
                st.write("**■ 連続勤務の緩和（※エクセルの妥協ターゲットにのみ適用）**")
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
                    for d in range(num_days): model.Add(shifts[(e, d, 'A残')] == 0)
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×":
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))
            
            for e, name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == name]
                if not tr.empty:
                    last_d = str(tr.iloc[0, 5]).strip()
                    if last_d == "D":
                        model.Add(shifts[(e, 0, 'E')] == 1)
                        if num_days > 1: model.Add(shifts[(e, 1, '公')] == 1)
                    elif last_d == "E": model.Add(shifts[(e, 0, '公')] == 1)
                    for d in range(num_days):
                        cv = str(tr.iloc[0, 6+d]).strip() if 6+d < len(df_history.columns) else ""
                        if cv == "公": model.Add(shifts[(e, d, '公')] == 1)
                
                if staff_night_ok[e] != "×":
                    if not tr.empty and str(tr.iloc[0, 5]).strip() != "D": model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0: model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d+1 < num_days: model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

                for d in range(num_days - 6): model.Add(shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(e, d+6, 'D')] <= 2)

            penalties = []
            
            w_minus_1 = get_penalty_weight(opt_minus_1)
            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                if absolute_req_list[d] == "〇" or w_minus_1 == -1:
                    model.Add(act_day >= day_req_list[d])
                else:
                    model.Add(act_day >= day_req_list[d] - 1)
                    minus_var = model.NewBoolVar('')
                    model.Add(act_day == day_req_list[d] - 1).OnlyEnforceIf(minus_var)
                    penalties.append(minus_var * w_minus_1)

            w_sub_only = get_penalty_weight(opt_sub_only)
            for d in range(num_days):
                l_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if w_sub_only == -1:
                    model.Add(l_score >= 2)
                else:
                    model.Add(l_score >= 1)
                    sub_var = model.NewBoolVar('')
                    model.Add(l_score == 1).OnlyEnforceIf(sub_var)
                    penalties.append(sub_var * w_sub_only)

            w_4_days = get_penalty_weight(opt_4_days)
            w_night_3 = get_penalty_weight(opt_night_3)
            
            for e in range(num_staff):
                target_weight = staff_comp_lvl[e]
                for d in range(num_days - 3):
                    model.Add(sum(shifts[(e, d+i, '公')] for i in range(4)) <= 3)
                    work = lambda x: shifts[(e, x, 'A')] + shifts[(e, x, 'A残')]
                    
                    if w_4_days != -1 and target_weight > 0:
                        if d < num_days - 4: model.Add(sum(work(d+i) for i in range(5)) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(sum(work(d+i) for i in range(4)) == 4).OnlyEnforceIf(p_var)
                        model.Add(sum(work(d+i) for i in range(4)) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * (w_4_days * target_weight))
                    else:
                        model.Add(sum(work(d+i) for i in range(4)) <= 3)

                    if w_night_3 != -1 and target_weight > 0:
                        np_var = model.NewBoolVar('')
                        model.Add(sum(work(d+i) for i in range(3)) == 3).OnlyEnforceIf(np_var)
                        model.Add(sum(work(d+i) for i in range(3)) <= 2).OnlyEnforceIf(np_var.Not())
                        final_p = model.NewIntVar(0, w_night_3 * target_weight, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, d+3, 'D')]])
                        penalties.append(final_p)
                    else:
                        model.Add(sum(work(d+i) for i in range(3)) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            w_ot_consec = get_penalty_weight(opt_ot_consec)
            for e in range(num_staff):
                for d in range(num_days - 1):
                    if w_ot_consec == -1:
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(ot_var * w_ot_consec)

            tot_ot = sum(overtime_req_list); tot_day = sum(day_req_list)
            if tot_ot > 0 and tot_day > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        act_d = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        act_o = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        diff = model.NewIntVar(-10000, 10000, ''); abs_diff = model.NewIntVar(0, 10000, '')
                        model.Add(diff == (act_o * tot_day) - (act_d * tot_ot))
                        model.AddAbsEquality(abs_diff, diff)
                        penalties.append(abs_diff)
            
            if penalties: model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0
            solver.parameters.random_seed = random_seed
            return (solver, shifts) if solver.Solve(model) in [cp_model.OPTIMAL, cp_model.FEASIBLE] else (None, None)

        if st.button("設定に基づき、シフトを【3パターン】作成する！"):
            with st.spinner('AIが優先順位と割合を計算し、3パターンのシフトを考えています...（最大3分）'):
                results = [res for seed in [1, 42, 99] if (res := solve_shift(seed))[0]]
                if not results: st.error("❌ 条件が厳しすぎます。設定画面で緩和する条件の「優先順位」を選択してください！")
                else:
                    st.success(f"✨完成！ {len(results)}パターン提案します！✨")
                    
                    # 🌟 修正ポイント：存在しない日付を無視する安全装置
                    cols = []
                    for d_val, w_val in zip(date_columns, weekdays):
                        try:
                            dt = datetime.date(target_year, target_month, int(d_val))
                            if jpholiday.is_holiday(dt):
                                cols.append(f"{d_val}({w_val}・祝)")
                            else:
                                cols.append(f"{d_val}({w_val})")
                        except ValueError:
                            # 2月30日などのエラーを回避
                            cols.append(f"{d_val}({w_val})")
                            
                    tabs = st.tabs([f"パターン {i+1}" for i in range(len(results))])
                    
                    for i, (solver, shifts) in enumerate(results):
                        with tabs[i]:
                            data = []
                            for e in range(num_staff):
                                row = {"スタッフ名": staff_names[e], "役割": staff_roles[e], "パート": staff_part_shifts[e]}
                                for d in range(num_days):
                                    for s in ['A', 'A残', 'D', 'E', '公']:
                                        if solver.Value(shifts[(e, d, s)]):
                                            row[cols[d]] = str(staff_part_shifts[e]).strip() if s in ['A','A残'] and str(staff_part_shifts[e]).strip() else s
                                data.append(row)
                            df_res = pd.DataFrame(data)
                            
                            df_res['日勤(A・P)回数'] = df_res[cols].apply(lambda x: x.str.contains('A|P|Ｐ', na=False) & ~x.str.contains('残', na=False)).sum(axis=1)
                            df_res['残業(A残)回数'] = (df_res[cols] == 'A残').sum(axis=1)
                            df_res['残業割合'] = df_res.apply(lambda r: f"{(r['残業(A残)回数']/r['日勤(A・P)回数'])*100:.1f}%" if r['日勤(A・P)回数']>0 else "0.0%", axis=1)
                            df_res['夜勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                            df_res['公休回数'] = (df_res[cols] == '公').sum(axis=1)
                            df_res['日曜D回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'D') if staff_sun_d[e] == "〇" else 0 for e in range(num_staff)]
                            df_res['日曜E回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'E') if staff_sun_e[e] == "〇" else 0 for e in range(num_staff)]

                            sum_A, sum_Az, sum_D, sum_O = {k: "" for k in df_res.columns}, {k: "" for k in df_res.columns}, {k: "" for k in df_res.columns}, {k: "" for k in df_res.columns}
                            sum_A.update({"スタッフ名": "【日勤(A・P) 合計】"}); sum_Az.update({"スタッフ名": "【残業(A残) 合計】"})
                            sum_D.update({"スタッフ名": "【夜勤(D) 合計】"}); sum_O.update({"スタッフ名": "【公休 合計】"})
                            
                            for d, c in enumerate(cols):
                                sum_A[c] = sum(1 for e in range(num_staff) if str(df_res.loc[e, c]) in ['A', 'A残'] or 'P' in str(df_res.loc[e, c]) and "新人" not in str(staff_roles[e]))
                                sum_Az[c] = (df_res[c] == 'A残').sum(); sum_D[c] = (df_res[c] == 'D').sum(); sum_O[c] = (df_res[c] == '公').sum()

                            df_fin = pd.concat([df_res, pd.DataFrame([sum_A, sum_Az, sum_D, sum_O])], ignore_index=True)

                            def hl(df):
                                s = pd.DataFrame('', index=df.index, columns=df.columns)
                                for d, c in enumerate(cols):
                                    v = df.loc[len(staff_names), c]
                                    if v != "" and v < day_req_list[d]: s.loc[len(staff_names), c] = 'background-color: #FFCCCC; color: red;'
                                    elif v != "" and v > day_req_list[d]: s.loc[len(staff_names), c] = 'background-color: #CCFFFF; color: blue;'
                                for e in range(num_staff):
                                    for d in range(num_days):
                                        w = lambda x: x < num_days and str(df.loc[e, cols[x]]) in ['A', 'A残', 'D', 'E'] or 'P' in str(df.loc[e, cols[x]])
                                        if w(d) and w(d+1) and w(d+2) and w(d+3):
                                            for i in range(4): s.loc[e, cols[d+i]] = 'background-color: #FFFF99;'
                                        if d+3 < num_days:
                                            v_a = lambda x: str(df.loc[e, cols[x]]) in ['A', 'A残'] or 'P' in str(df.loc[e, cols[x]])
                                            if v_a(d) and v_a(d+1) and v_a(d+2) and str(df.loc[e, cols[d+3]]) == 'D':
                                                for i in range(4): s.loc[e, cols[d+i]] = 'background-color: #FFD580;'
                                return s

                            st.dataframe(df_fin.style.apply(hl, axis=None))
                            out = io.BytesIO()
                            with pd.ExcelWriter(out, engine='openpyxl') as w: df_fin.to_excel(w, index=False, sheet_name='完成シフト')
                            st.download_button(f"📥 【パターン {i+1}】 をダウンロード（色なし）", out.getvalue(), f"完成版_パターン{i+1}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_{i}")

    except Exception as e:
        st.error(f"⚠️ エラー: エクセル形式または項目名を確認してください。({e})")
