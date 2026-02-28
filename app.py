import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
import random
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🤝 AIシフト作成 Co-Pilot")
st.write("現場のルールと思いやりを考慮して、AIと一緒にシフトを作るシステムです。")

# 🌟 NEW: スマホユーザーのための便利リンクエリア
with st.expander("📱 【事前準備】入力用エクセルをお持ちでない方へ", expanded=True):
    st.markdown("""
    **方法1：クラウドで入力する（おすすめ）**
    以下のリンクからGoogleスプレッドシートを開き、スマホで直接入力してください。
    入力後、`ファイル ＞ ダウンロード ＞ Microsoft Excel (.xlsx)` を選んで保存してください。
    👉 [入力用スプレッドシートを開く（※ここにあなたのスプレッドシートのURLを貼れます）](#)
    """)
    
    st.markdown("**方法2：空のテンプレートをダウンロードして使う**")
    # 空のデータフレームを作ってエクセル化し、ダウンロードさせる
    output_tmpl = io.BytesIO()
    with pd.ExcelWriter(output_tmpl, engine='openpyxl') as writer:
        pd.DataFrame(columns=["スタッフ名", "役割", "公休数", "夜勤可否", "夜勤上限", "残業可否", "日曜Dカウント", "日曜Eカウント", "パート", "妥協優先度", "定時確保数"]).to_excel(writer, index=False, sheet_name='スタッフ設定')
        pd.DataFrame(columns=["スタッフ名", "前月27", "前月28", "前月29", "前月30", "前月31", "1", "2", "3", "4", "5"]).to_excel(writer, index=False, sheet_name='希望休・前月履歴')
        pd.DataFrame({"項目": ["曜日", "日勤人数", "夜勤人数", "残業人数", "絶対確保"], "1": ["月", 3, 2, 0, ""], "2": ["火", 3, 2, 0, ""]}).to_excel(writer, index=False, sheet_name='日別設定')
    st.download_button(label="📥 空のテンプレート（.xlsx）をダウンロード", data=output_tmpl.getvalue(), file_name="入力用テンプレート.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if 'needs_compromise' not in st.session_state:
    st.session_state.needs_compromise = False

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
                else: res.append(default_val)
            return res

        staff_roles = get_staff_col("役割", "一般")
        staff_off_days = get_staff_col("公休数", 9, is_int=True)
        staff_night_ok = get_staff_col("夜勤可否", "〇")
        staff_overtime_ok = get_staff_col("残業可否", "〇")
        staff_part_shifts = get_staff_col("パート", "")
        staff_night_limits = [0 if ok == "×" else int(v) if pd.notna(v) else 10 for ok, v in zip(staff_night_ok, get_staff_col("夜勤上限", 10, is_int=True))]
        staff_min_normal_a = get_staff_col("定時確保数", 2, is_int=True)

        staff_comp_lvl = []
        for i in range(num_staff):
            val = ""
            if "妥協優先度" in df_staff.columns and pd.notna(df_staff["妥協優先度"].iloc[i]): val = str(df_staff["妥協優先度"].iloc[i]).strip()
            elif "連勤妥協OK" in df_staff.columns and pd.notna(df_staff["連勤妥協OK"].iloc[i]): val = str(df_staff["連勤妥協OK"].iloc[i]).strip()
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
        weekdays = [str(df_req.iloc[0, d+1]).strip() if (d+1) < len(df_req.columns) and pd.notna(df_req.iloc[0, d+1]) else "" for d in range(num_days)]

        st.success("✅ データの読み込み完了！まずは妥協なしの「理想のシフト」を作れるかテストします。")

        def solve_shift(random_seed, allow_minus_1=False, allow_4_days=False, allow_night_3=False, allow_sub_only=False, allow_ot_consec=False, allow_night_consec_3=False):
            model = cp_model.CpModel()
            types = ['A', 'A残', 'D', 'E', '公']
            shifts = {(e, d, s): model.NewBoolVar('') for e in range(num_staff) for d in range(num_days) for s in types}
            
            random.seed(random_seed)
            for e in range(num_staff):
                for d in range(num_days): model.AddHint(shifts[(e, d, 'A')], random.choice([0, 1]))
                for d in range(num_days): model.AddExactlyOne(shifts[(e, d, s)] for s in types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 0); model.Add(shifts[(e, d, 'E')] == 0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days): model.Add(shifts[(e, d, 'A残')] == 0)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    last_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_day == "D":
                        model.Add(shifts[(e, 0, 'E')] == 1)
                        if num_days > 1: model.Add(shifts[(e, 1, '公')] == 1)
                    elif last_day == "E": model.Add(shifts[(e, 0, '公')] == 1)

            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                        if l_day != "D": model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0: model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days: model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

            penalties = []
            
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    for d in range(num_days - 3): model.Add(shifts[(e, d, 'E')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, 'D')] <= 3)
                    for d in range(num_days - 4): model.Add(shifts[(e, d, 'E')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] + shifts[(e, d+4, 'D')] <= 4)
                    
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty and tr.shape[1] > 5:
                        l_5 = [str(tr.iloc[0, i]).strip() for i in range(1, 6)]
                        if l_5[4] == "E":
                            if num_days > 2: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, '公')] + shifts[(e, 2, 'D')] <= 2)
                            if num_days > 3: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, '公')] + shifts[(e, 2, '公')] + shifts[(e, 3, 'D')] <= 3)
                        if l_5[3] == "E" and l_5[4] == "公":
                            if num_days > 1: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, 'D')] <= 1)
                            if num_days > 2: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, '公')] + shifts[(e, 2, 'D')] <= 2)

            for e, staff_name in enumerate(staff_names):
                if staff_night_ok[e] != "×":
                    past_D = [0] * 5
                    tr = df_history[df_history.iloc[:, 0] == staff_name]
                    if not tr.empty:
                        for i in range(5):
                            if (i+1) < tr.shape[1] and str(tr.iloc[0, i+1]).strip() == "D": past_D[i] = 1
                    
                    all_D = past_D + [shifts[(e, d, 'D')] for d in range(num_days)]
                    for i in range(len(all_D) - 6):
                        window = all_D[i : i+7]
                        if not allow_night_consec_3:
                            if any(isinstance(x, cp_model.IntVar) for x in window): model.Add(sum(window) <= 2)
                        else:
                            if any(isinstance(x, cp_model.IntVar) for x in window):
                                n3_var = model.NewBoolVar('')
                                model.Add(sum(window) >= 3).OnlyEnforceIf(n3_var)
                                model.Add(sum(window) <= 2).OnlyEnforceIf(n3_var.Not())
                                penalties.append(n3_var * 5000)

            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum((shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                req = day_req_list[d]
                is_sun = ('日' in weekdays[d])
                is_abs = (absolute_req_list[d] == "〇")

                if is_abs:
                    model.Add(act_day >= req)
                    over_var = model.NewIntVar(0, 100, ''); diff = model.NewIntVar(-100, 100, '')
                    model.Add(diff == act_day - req); model.AddMaxEquality(over_var, [0, diff])
                    penalties.append(over_var * 1) 
                elif is_sun:
                    model.Add(act_day <= req)
                    if not allow_minus_1: model.Add(act_day == req)
                    else:
                        model.Add(act_day >= req - 1); m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var); model.Add(act_day == req).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * 1000)
                else:
                    if not allow_minus_1: model.Add(act_day >= req)
                    else:
                        model.Add(act_day >= req - 1); m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var); model.Add(act_day != req - 1).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * 1000)
                    over_var = model.NewIntVar(0, 100, ''); diff = model.NewIntVar(-100, 100, '')
                    model.Add(diff == act_day - req); model.AddMaxEquality(over_var, [0, diff])
                    penalties.append(over_var * 100)

                l_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if not allow_sub_only: model.Add(l_score >= 2)
                else:
                    model.Add(l_score >= 1); sub_var = model.NewBoolVar('')
                    model.Add(l_score == 1).OnlyEnforceIf(sub_var); penalties.append(sub_var * 1000)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    for d in range(num_days):
                        col_idx = 6 + d
                        if col_idx < tr.shape[1] and str(tr.iloc[0, col_idx]).strip() == "公": model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×": model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

            limit_groups = {}
            for e in range(num_staff):
                if staff_night_ok[e] != "×" and int(staff_night_limits[e]) > 0:
                    limit_groups.setdefault(int(staff_night_limits[e]), []).append(e)
            for limit, members in limit_groups.items():
                if len(members) >= 2:
                    actual_nights = [sum(shifts[(m, d, 'D')] for d in range(num_days)) for m in members]
                    max_n = model.NewIntVar(0, limit, ''); min_n = model.NewIntVar(0, limit, '')
                    model.AddMaxEquality(max_n, actual_nights); model.AddMinEquality(min_n, actual_nights)
                    model.Add(max_n - min_n <= 1)

            for e in range(num_staff):
                for d in range(num_days - 3): model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <= 3)
                for d in range(num_days - 2):
                    is_3_off = model.NewBoolVar('')
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] == 3).OnlyEnforceIf(is_3_off)
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] <= 2).OnlyEnforceIf(is_3_off.Not())
                    penalties.append(is_3_off * 500) 

                is_2_offs = []
                for d in range(num_days - 1):
                    is_2_off = model.NewBoolVar('')
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] == 2).OnlyEnforceIf(is_2_off)
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] <= 1).OnlyEnforceIf(is_2_off.Not())
                    is_2_offs.append(is_2_off)
                has_any_2_off = model.NewBoolVar('')
                model.Add(sum(is_2_offs) >= 1).OnlyEnforceIf(has_any_2_off); model.Add(sum(is_2_offs) == 0).OnlyEnforceIf(has_any_2_off.Not())
                penalties.append(has_any_2_off.Not() * 300) 

            for e in range(num_staff):
                target_lvl = staff_comp_lvl[e]
                w_base = 10 ** target_lvl if target_lvl > 0 else 0
                for d in range(num_days - 3):
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                    if allow_4_days and target_lvl > 0:
                        if d < num_days - 4: model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf(p_var)
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * w_base)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3)

                    if allow_night_3 and target_lvl > 0:
                        np_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Not())
                        final_p = model.NewIntVar(0, w_base, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, d+3, 'D')]])
                        penalties.append(final_p)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            for e in range(num_staff):
                for d in range(num_days - 1):
                    if not allow_ot_consec: model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(ot_var * 500)

            for e in range(num_staff):
                if staff_overtime_ok[e] != "×":
                    total_day_work = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                    b_has_work = model.NewBoolVar('')
                    model.Add(total_day_work > 0).OnlyEnforceIf(b_has_work); model.Add(total_day_work == 0).OnlyEnforceIf(b_has_work.Not())
                    min_a = int(staff_min_normal_a[e])
                    total_a_normal = sum(shifts[(e, d, 'A')] for d in range(num_days))
                    model.Add(total_a_normal >= min_a).OnlyEnforceIf(b_has_work)

            ot_burden_scores = []
            for e in range(num_staff):
                if staff_overtime_ok[e] != "×":
                    total_work_score = sum(shifts[(e, d, 'A')] + (shifts[(e, d, 'A残')] * 2) for d in range(num_days)) 
                    ot_burden_scores.append(total_work_score)
            
            if ot_burden_scores:
                max_b = model.NewIntVar(0, 100, ''); min_b = model.NewIntVar(0, 100, '')
                model.AddMaxEquality(max_b, ot_burden_scores); model.AddMinEquality(min_b, ot_burden_scores)
                penalties.append((max_b - min_b) * 50)

            for e in range(num_staff):
                ot_bias = random.randint(-2, 2); night_bias = random.randint(-2, 2); off_bias = random.randint(-2, 2)
                if staff_overtime_ok[e] != "×": penalties.append(sum(shifts[(e, d, 'A残')] for d in range(num_days)) * ot_bias)
                if staff_night_ok[e] != "×": penalties.append(sum(shifts[(e, d, 'D')] for d in range(num_days)) * night_bias)
                penalties.append(sum(shifts[(e, d, '公')] for d in range(num_days)) * off_bias)
                for d in range(num_days): penalties.append(shifts[(e, d, 'A')] * random.randint(-1, 1))
            
            if penalties: model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 45.0 
            solver.parameters.random_seed = random_seed
            status = solver.Solve(model)
            
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE: return solver, shifts
            else: return None, None

        if not st.session_state.needs_compromise:
            if st.button("▶️ 【STEP 1】まずは妥協なしで理想のシフトを計算する（3パターン）"):
                with st.spinner('AIが「妥協なし」の完璧なシフトを3パターン模索中...'):
                    results = []
                    for seed in [1, 42, 99]:
                        solver, shifts = solve_shift(seed, False, False, False, False, False, False)
                        if solver: results.append((solver, shifts))
                        
                    if results: st.success(f"🎉 なんと！妥協なしで完璧なシフトが {len(results)} パターン組めました！")
                    else:
                        st.session_state.needs_compromise = True
                        st.rerun()
        else:
            st.error("⚠️ 【AI店長からのご報告】\n申し訳ありません。現在の人数と希望休では、すべてのルールを完璧に守ってシフトを組むことは不可能でした...")
            st.warning("💡 以下のいずれかの「妥協案」を許可して、再計算を指示してください。")
            
            with st.container():
                st.markdown("### 📝 妥協の提案リスト")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**■ 人数と役割について**")
                    allow_minus_1 = st.checkbox("日勤人数の「マイナス1」を許可する（絶対確保日以外）")
                    allow_sub_only = st.checkbox("役割配置を「サブ1名＋他」まで下げることを許可する")
                with col2:
                    st.markdown("**■ 対象スタッフへの連勤お願い**")
                    allow_4_days = st.checkbox("対象者への「最大4連勤」のお願いを許可する")
                    allow_night_3 = st.checkbox("対象者への「夜勤前3日連続日勤」のお願いを許可する")
                
                st.markdown("**■ その他の例外ルール**")
                col3, col4 = st.columns(2)
                with col3:
                    allow_night_consec_3 = st.checkbox("やむを得ない「月またぎ含む、夜勤セット3連続」を許可する")
                with col4:
                    allow_ot_consec = st.checkbox("やむを得ない「残業(A残)の2日連続」を許可する")

            if st.button("🔄 【STEP 3】チェックした妥協案を許可して、3パターンのシフトを作る！"):
                with st.spinner('許可された妥協案をもとに、AIが再計算しています...'):
                    results = []
                    for seed in [1, 42, 99]:
                        solver, shifts = solve_shift(seed, allow_minus_1, allow_4_days, allow_night_3, allow_sub_only, allow_ot_consec, allow_night_consec_3)
                        if solver: results.append((solver, shifts))

                    if not results: st.error("😭 まだ条件が厳しすぎます！もう少しだけ他の妥協案も許可してもらえませんか？")
                    else:
                        st.success(f"✨ ありがとうございます！許可いただいた条件内で、{len(results)}パターンのシフトが完成しました！")
                        st.session_state.needs_compromise = False

        if 'results' in locals() and results:
            cols = []
            for d_val, w_val in zip(date_columns, weekdays):
                try:
                    dt = datetime.date(target_year, target_month, int(d_val))
                    if jpholiday.is_holiday(dt): cols.append(f"{d_val}({w_val}・祝)")
                    else: cols.append(f"{d_val}({w_val})")
                except ValueError: cols.append(f"{d_val}({w_val})")

            tabs = st.tabs([f"提案パターン {i+1}" for i in range(len(results))])
            for i, (solver, shifts) in enumerate(results):
                with tabs[i]:
                    data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e]}
                        for d in range(num_days):
                            for s in ['A', 'A残', 'D', 'E', '公']:
                                if solver.Value(shifts[(e, d, s)]):
                                    if (s == 'A' or s == 'A残') and str(staff_part_shifts[e]).strip() not in ["", "nan"]: row[cols[d]] = str(staff_part_shifts[e]).strip()
                                    else: row[cols[d]] = s
                        data.append(row)
                        
                    df_res = pd.DataFrame(data)

                    df_res['日勤(A/P)回数'] = df_res[cols].apply(lambda x: x.str.contains('A|P|Ｐ', na=False) & ~x.str.contains('残', na=False)).sum(axis=1)
                    df_res['残業(A残)回数'] = (df_res[cols] == 'A残').sum(axis=1)
                    df_res['残業割合(%)'] = df_res.apply(lambda r: f"{(r['残業(A残)回数']/r['日勤(A/P)回数'])*100:.1f}%" if r['日勤(A/P)回数']>0 else "0.0%", axis=1)
                    df_res['夜勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                    df_res['公休回数'] = (df_res[cols] == '公').sum(axis=1)
                    df_res['日曜D回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'D') if staff_sun_d[e] == "〇" else 0 for e in range(num_staff)]
                    df_res['日曜E回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'E') if staff_sun_e[e] == "〇" else 0 for e in range(num_staff)]

                    sum_A = {"スタッフ名": "【日勤(A/P) 合計人数】"}
                    sum_Az = {"スタッフ名": "【残業(A残) 合計人数】"}
                    sum_D = {"スタッフ名": "【夜勤(D) 合計人数】"}
                    sum_O = {"スタッフ名": "【公休 合計人数】"}
                    
                    for c in ['日勤(A/P)回数', '残業(A残)回数', '残業割合(%)', '夜勤(D)回数', '公休回数', '日曜D回数', '日曜E回数']:
                        sum_A[c] = ""; sum_Az[c] = ""; sum_D[c] = ""; sum_O[c] = ""

                    for d, c in enumerate(cols):
                        sum_A[c] = sum(1 for e in range(num_staff) if str(df_res.loc[e, c]) in ['A', 'A残'] or 'P' in str(df_res.loc[e, c]) and "新人" not in str(staff_roles[e]))
                        sum_Az[c] = (df_res[c] == 'A残').sum()
                        sum_D[c] = (df_res[c] == 'D').sum()
                        sum_O[c] = (df_res[c] == '公').sum()

                    df_fin = pd.concat([df_res, pd.DataFrame([sum_A, sum_Az, sum_D, sum_O])], ignore_index=True)

                    def highlight_warnings(df):
                        styles = pd.DataFrame('', index=df.index, columns=df.columns)
                        for d, col_name in enumerate(cols):
                            actual_a = df.loc[len(staff_names), col_name]
                            target_a = day_req_list[d]
                            if actual_a != "":
                                if actual_a < target_a: styles.loc[len(staff_names), col_name] = 'background-color: #FFCCCC; color: red; font-weight: bold;'
                                elif actual_a > target_a: styles.loc[len(staff_names), col_name] = 'background-color: #CCFFFF; color: blue; font-weight: bold;'
                        
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
                                        
                                if d + 8 < num_days:
                                    if str(df.loc[e, cols[d]]) == 'D' and str(df.loc[e, cols[d+3]]) == 'D' and str(df.loc[e, cols[d+6]]) == 'D':
                                        for k in range(9): styles.loc[e, cols[d+k]] = 'background-color: #E6E6FA;'
                        return styles

                    st.dataframe(df_fin.style.apply(highlight_warnings, axis=None))
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_fin.to_excel(writer, index=False, sheet_name='完成シフト')
                        worksheet = writer.sheets['完成シフト']
                        
                        font_meiryo = Font(name='Meiryo')
                        border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                        align_center = Alignment(horizontal='center', vertical='center')
                        align_left = Alignment(horizontal='left', vertical='center')
                        
                        fill_sat = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
                        fill_sun = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
                        fill_short = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                        fill_over = PatternFill(start_color="CCFFFF", end_color="CCFFFF", fill_type="solid")
                        fill_4days = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                        fill_n3 = PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid")
                        fill_n3_consec = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")

                        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                            for cell in row:
                                cell.font = font_meiryo
                                cell.border = border_thin
                                cell.alignment = align_left if cell.column == 1 else align_center

                        for c_idx, col_name in enumerate(cols):
                            if "土" in col_name:
                                for r_idx in range(1, len(df_fin) + 2): worksheet.cell(row=r_idx, column=c_idx+2).fill = fill_sat
                            elif "日" in col_name or "祝" in col_name:
                                for r_idx in range(1, len(df_fin) + 2): worksheet.cell(row=r_idx, column=c_idx+2).fill = fill_sun

                        row_a_idx = len(staff_names) + 2
                        for d, col_name in enumerate(cols):
                            actual_a = df_fin.loc[len(staff_names), col_name]
                            if actual_a != "":
                                if actual_a < day_req_list[d]: worksheet.cell(row=row_a_idx, column=d+2).fill = fill_short
                                elif actual_a > day_req_list[d]: worksheet.cell(row=row_a_idx, column=d+2).fill = fill_over

                        for e in range(num_staff):
                            for d in range(num_days):
                                def is_d_work(day_idx):
                                    if day_idx >= num_days: return False
                                    v = str(df_fin.loc[e, cols[day_idx]])
                                    return v == 'A' or v == 'A残' or 'P' in v or 'Ｐ' in v

                                if is_d_work(d) and is_d_work(d+1) and is_d_work(d+2) and is_d_work(d+3):
                                    for k in range(4): worksheet.cell(row=e+2, column=d+k+2).fill = fill_4days

                                if d + 3 < num_days:
                                    if is_d_work(d) and is_d_work(d+1) and is_d_work(d+2) and str(df_fin.loc[e, cols[d+3]]) == 'D':
                                        for k in range(4): worksheet.cell(row=e+2, column=d+k+2).fill = fill_n3
                                        
                                if d + 8 < num_days:
                                    if str(df_fin.loc[e, cols[d]]) == 'D' and str(df_fin.loc[e, cols[d+3]]) == 'D' and str(df_fin.loc[e, cols[d+6]]) == 'D':
                                        for k in range(9): worksheet.cell(row=e+2, column=d+k+2).fill = fill_n3_consec

                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label=f"📥 【パターン {i+1}】 をエクセルでダウンロード（レイアウト完成版）",
                        data=processed_data,
                        file_name=f"完成版_対話型シフト_{i+1}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_btn_{i}"
                    )
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白の行があります。({e})")
