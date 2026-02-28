import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
import random

st.set_page_config(page_title="自動シフト作成アプリ", layoutこと」「夜勤3連続の緩和チェック機能」をすべて盛り込み、あなたがシステム指示に登録したプロンプト通りに正確に動作="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ19：集計復活＆夜勤緩和版)")
st.write("現場必須の集計欄をすべて復活させ、「夜勤セット3連続」の厳格な緩和管理を追加しました！")

st.write("---")
today = datetime.date.today()
col_y,する**【真の完成版コード】**を作成しました。

右側の集計列、下部の集計行 col_m = st.columns(2)
with col_y: target_year = st.selectbox("作成年", [today.year, today.year + 1], index=0)
with col_m: target_month = st.selectbox("作成月", list(range(1, 13)), index=(がすべて復活し、日勤回数の計算（A＋A残＋P〇）も現場の定義通りにtoday.month % 12))
st.write("---")

uploaded_file = st.file_修正されています。

---

### 🛠️ アプリの修正（集計欄復活＆夜勤3連続uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")緩和版）

GitHubの `app.py` を開き、以下のコードに**すべて上書き**してください。

▼ ここから下をすべてコピー ▼
```python
import streamlit as st
import pandas as pd
from
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・前月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="日別設定")
        
        staff_names = df_staff["スタッフ名"].dropna().tolist()
         ortools.sat.python import cp_model
import io
import jpholiday
import datetime
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
stnum_staff = len(staff_names)
        
        def get_staff_col(col_name, default_val, is_int=False):
            res = []
            for i in range(num.title("🌟 AI自動シフト作成アプリ (フェーズ19：集計欄復活＆完全版)")
st.write("現場_staff):
                if col_name in df_staff.columns and pd.notna(df_staff[col_name].iloc[i]):
                    val = df_staff[col_name].iloc[iで必須の集計欄を完全復活させ、夜勤3連続の緩和機能を追加しました！")

st.write("---")]
                    res.append(int(val) if is_int else str(val).strip())
                else:
                    res.append(default_val)
            return res

        staff_roles = get_
today = datetime.date.today()
col_y, col_m = st.columns(2)staff_col("役割", "一般")
        staff_off_days = get_staff_col("公休数", 8, is_int=True)
        staff_night_ok = get_staff_
with col_y: target_year = st.selectbox("作成年", [today.year, today.year + 1], index=0)
with col_m: target_month = st.selectbox("作成月", list(range(1, 13)), index=(today.month % 12))
st.write("---")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

ifcol("夜勤可否", "〇")
        staff_overtime_ok = get_staff_col("残業可否", "〇")
        staff_part_shifts = get_staff_col("パート", "")
        
        staff_night_limits = []
        raw_limits = get_staff_col("夜勤上限", 10, is_int=True)
        for i in range(num_staff):
            staff_night_limits.append(0 if staff_night_ok[i] == " uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・前月履歴")
        df_req = pd.read_excel(uploaded_×" else raw_limits[i])

        staff_comp_lvl = []
        for i in range(num_staff):
            val = ""
            if "妥協優先度" in df_staff.columns and pd.notna(df_staff["妥協優先度"].iloc[i]):
                val = str(df_staff["妥協優先度"].iloc[i]).strip()
            elif "連勤妥協OK" infile, sheet_name="日別設定")
        
        staff_names = df_staff["スタッフ名 df_staff.columns and pd.notna(df_staff["連勤妥協OK"].iloc[i]):
                val = str(df_staff["連勤妥協OK"].iloc[i]).strip()
"].dropna().tolist()
        num_staff = len(staff_names)
        
        def get_            
            if val in ["〇", "1", "1.0"]: staff_comp_lvl.appendstaff_col(col_name, default_val, is_int=False):
            res = []
            for i in range(num_staff):
                if col_name in df_staff.columns and pd.notna(df_staff[col_name].iloc[i]):
                    val = df_staff[col_name].iloc[i]
                    res.append(int(val) if is_int else str(1)
            elif val in ["2", "2.0"]: staff_comp_lvl.append(2)
            elif val in ["3", "3.0"]: staff_comp_lvl.append(3)
            else: staff_comp_lvl.append(0)

        raw_sun_d = get_staff_col("日曜Dカウント", "〇")
        raw_sun_e = get_staff_(val).strip())
                else:
                    res.append(default_val)
            return res

        staff_roles = get_staff_col("役割", "一般")
        staff_off_days =col("日曜Eカウント", "〇")
        staff_sun_d = ["×" if staff_night_ok[i] == "×" else raw_sun_d[i] for i in range(num_staff)]
        staff_sun_e = ["×" if staff_night_ok[i] == get_staff_col("公休数", 8, is_int=True)
        staff_night_ok = get_staff_col("夜勤可否", "〇")
        staff_overtime_ "×" else raw_sun_e[i] for i in range(num_staff)]

        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        ok = get_staff_col("残業可否", "〇")
        staff_part_shifts = get_staff_col("パート", "")
        
        staff_night_limits = [0 if ok == "×" else int(v) if pd.notna(v) else 10 for ok, v in zip(staff
        def get_req_col(label, default_val, is_int=True):
            row = df_req[df_req.iloc[:, 0] == label]
            res = []
            for d in range(num_days):
                if not row.empty and (d + 1) < len(df_req.columns):
                    val = row.iloc[0, d + 1]
_night_ok, get_staff_col("夜勤上限", 10, is_int=True))]
        staff_sun_d = ["×" if ok == "×" else v for ok, v                    if pd.notna(val):
                        res.append(int(val) if is_int else str(val).strip())
                        continue
                res.append(default_val)
            return res

 in zip(staff_night_ok, get_staff_col("日曜Dカウント", "〇"))]
        day_req_list = get_req_col("日勤人数", 3)
        night_req_list = get_req_col("夜勤人数", 2)
        overtime_req_        staff_sun_e = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Eカウント", "〇"))]

        staff_comp_lvl = []
        for i in range(num_staff):
            val = ""
            if "妥協優先度list = get_req_col("残業人数", 0)
        absolute_req_list = get_req_col("絶対確保", "", is_int=False)

        weekdays = []
        for d in range(num_days):
            if (d + 1) < len(df_req." in df_staff.columns and pd.notna(df_staff["妥協優先度"].iloc[columns):
                val = df_req.iloc[0, d + 1]
                weekdays.append(str(val).strip() if pd.notna(val) else "")
            else:
                weekdays.append("")

        st.success("✅ データの読み込み完了！集計欄を復活させi]): val = str(df_staff["妥協優先度"].iloc[i]).strip()
            elif "連勤妥協ました。")
        
        with st.expander("⚙️ 【高度な設定】緩和ルールの優先順位（※どうしても組めない時だけ設定）", expanded=True):
            st.info("※OK" in df_staff.columns and pd.notna(df_staff["連勤妥協OK"].iloc[i]): val = str(df_staff["連勤妥協OK"].iloc[i]).strip()「緩和」は本当にどうしても組めない時の【最終手段】です。ペナルティが低い(優先順位1)項目から順
            if val in ["〇", "1", "1.0"]: staff_comp_lvl.append(1)
            elif val in ["2", "2.0"]: staff_comp_lvl.append(2)
            にAIが使用します。")
            options = ["許可しない（絶対死守）", "優先順位 1（最初に妥協）", "優先順位 2", "優先順位 3（最終手段）"]
elif val in ["3", "3.0"]: staff_comp_lvl.append(3)
            else: staff_comp_lvl.append(0)

        date_columns = [col for col in df_            
            col1, col2, col3 = st.columns(3)
            with col1:req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        def get_req_col(label
                st.write("**■ 人数と役割の緩和**")
                opt_minus_1 = st.selectbox("日勤人数の「マイナス1」許容", options, index=0)
                opt_sub_only = st.selectbox("役割「サブ1名のみ」許容", options, index=0), default_val, is_int=True):
            row = df_req[df_req.iloc[:, 0] == label]
            with col2:
                st.write("**■ 連続勤務の緩和（対象者のみ）**")
                
            res = []
            for d in range(num_days):
                if not row.empty and (d + 1) < len(df_req.columns):
                    val = row.iloc[0, d + 1]
opt_4_days = st.selectbox("対象者の「最大4連勤」許容", options, index=0)
                opt_night_3 = st.selectbox("対象者の「夜勤前3日勤」許容", options, index=0)
            with col3:
                st.write("**■ 夜                    if pd.notna(val):
                        res.append(int(val) if is_int else str(val).strip())
                        continue
                res.append(default_val)
            return res

        day勤・残業の緩和**")
                opt_night_consec = st.selectbox("やむを得ない「夜勤3連続」許容", options, index=0)
                opt_ot_consec = st.selectbox("やむを得ない_req_list = get_req_col("日勤人数", 3)
        night_req_list = get_req_col("夜勤人数", 2)
        overtime_req_list =「A残2連続」許容", options, index=0)

        def get_penalty_weight(opt_str):
 get_req_col("残業人数", 0)
        absolute_req_list = get_req            if "許可しない" in opt_str: return -1
            elif "優先順位 1" in opt_str: return 100
            elif "優先順位 2" in opt_str_col("絶対確保", "", is_int=False)

        weekdays = [str(df_req.iloc[0, d+1]).strip() if (d+1) < len(df_req.columns) and pd.: return 1000
            elif "優先順位 3" in opt_str: return 10000
            return -1

        def solve_shift(random_seed):
            model = cp_model.CpModel()
            types = ['A', 'A残', 'D', 'Enotna(df_req.iloc[0, d+1]) else "" for d in range(num_days)]

        st.success("✅ データの読み込み完了！")
        
        with st.expander("⚙️ 【高度な設定】緩和ルールの優先順位（※どうしても組めない時だけ設定）", expanded=True):
            st.info("※「緩和」は本当にどうしても組めない時の【最終手段】です', '公']
            shifts = {(e, d, s): model.NewBoolVar('') for e in range(num_staff) for d in range(num_days) for s in types}
            model.AddHint(shifts[(0, 0, 'A')], random.choice([0, 1]))

            for e in range(num_staff):
                for d in range(num_days):
                    model.AddExactlyOne(shifts[(e, d, s)] for s in types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 。勝手な乱用はしません。")
            options = ["許可しない（絶対死守）", "優先順位0); model.Add(shifts[(e, d, 'E')] == 0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days):
                         1（最初に妥協）", "優先順位 2", "優先順位 3（最終手段）"]
            col1, col2 = st.columns(2)
            with col1:
                model.Add(shifts[(e, d, 'A残')] == 0)

            for e, staffst.write("**■ 人数と役割の緩和**")
                opt_minus_1 = st.selectbox_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    last_month_last_day = str(tr.("日勤人数の「マイナス1」許容", options, index=0)
                opt_sub_only = st.selectbox("役割配置「サブ1名のみ」の許容", options, index=0iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_month_last_day == "D":
                        model.Add(shifts[(e, 0, 'E')] == 1))
            with col2:
                st.write("**■ 連続勤務の緩和（※エクセルの妥協優先度に沿って適用）**")
                opt_4_days = st.selectbox("対象者の
                        if num_days > 1:
                            model.Add(shifts[(e, 1, '「最大4連勤」許容", options, index=0)
                opt_night_3 = st公')] == 1)
                    elif last_month_last_day == "E":
                        model..selectbox("対象者の「夜勤前3日勤」許容", options, index=0)
                opt_night_3_consec = st.selectbox("対象者の「夜勤3連続(DE公DE公D)」許容", options, index=0)
                opt_ot_consec = st.selectboxAdd(shifts[(e, 0, '公')] == 1)

            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                        if l_day != "D":
                            model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d("やむを得ない「残業(A残)2日連続」の許容", options, index=0)

        def get_penalty_weight(opt_str):
            if "許可しない" in opt_str: return -1
            elif "優先順位 1" in opt_str: return 1 > 0:
                            model.Add(shifts[(e, d, 'E')] == shifts[(e,00
            elif "優先順位 2" in opt_str: return 1000
 d-1, 'D')])
                        if d + 1 < num_days:
                            model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

            penalties = []
            
            # 🌟 NEW: 夜勤セット3連続の緩和ロジック
            w_night_            elif "優先順位 3" in opt_str: return 10000
            return -1

        def solve_shift(random_seed):
            model = cp_model.CpModel()
            types = ['A', 'A残', 'D', 'E', '公']
            shifts =consec = get_penalty_weight(opt_night_consec)
            for e in range(num {(e, d, s): model.NewBoolVar('') for e in range(num_staff) for d in range(num__staff):
                for d in range(num_days - 6):
                    d_sum = shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(edays) for s in types}
            model.AddHint(shifts[(0, 0, 'A')], random.choice([0, 1]))

            for e in range(num_staff):
                for d, d+6, 'D')]
                    if w_night_consec == -1:
                        # 許可しない場合は絶対禁止（最大2回まで）
                        model.Add(d_sum <= 2)
                    else in range(num_days):
                    model.AddExactlyOne(shifts[(e, d, s)] for s in types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 0); model.Add(shifts[(:
                        # 許可する場合はペナルティ付きで3連続を許容（4連続は流石に絶対禁止）
                        if d < num_days - 9:
                            model.Add(d_sum +e, d, 'E')] == 0)
                if staff_overtime_ok[e] == shifts[(e, d+9, 'D')] <= 3)
                        n3_var = model. "×":
                    for d in range(num_days): model.Add(shifts[(e, d, 'A残')] == 0)

            for e, staff_name in enumerate(staff_names):
NewBoolVar('')
                        model.Add(d_sum == 3).OnlyEnforceIf(n3_var)
                        model.Add(d_sum <= 2).OnlyEnforceIf(n3_var                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if.Not())
                        penalties.append(n3_var * w_night_consec *  not tr.empty:
                    last_month_last_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_month_last_day == "D":
                        model.Add(shifts[(e, 0, 'E')] == 100) # ペナルティ重め

            w_minus_1 = get_penalty_weight(opt_minus_1)
            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                1)
                        if num_days > 1: model.Add(shifts[(e, 1, 'model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum((shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff公')] == 1)
                    elif last_month_last_day == "E":
                        model.Add(shifts[(e) if "新人" not in str(staff_roles[e]))
                req = day_req_list[d]
                is_sun = ('日' in weekdays[d])
                is_abs = (absolute_req_list[d, 0, '公')] == 1)

            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history] == "〇")

                if is_sun:
                    model.Add(act_day <= req)
                    if is_abs or w_minus_1 == -1:
                        model.Add(act_.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] >day == req)
                    else:
                        model.Add(act_day >= req - 1)
                        minus_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1 5 else ""
                        if l_day != "D": model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d >).OnlyEnforceIf(minus_var)
                        penalties.append(minus_var * w_minus_1 * 100)
                else:
                    model.Add(act_day <= req + 1)
 0: model.Add(shifts[(e, d, 'E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days: model.AddImplication(shifts[(                    if is_abs or w_minus_1 == -1:
                        model.Add(act_day >= req)
                    else:
                        model.Add(act_day >= req - 1)
                        e, d, 'E')], shifts[(e, d+1, '公')])

            penalties =minus_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(minus_var)
                        penalties.append(minus_var * w_minus []
            
            # 夜勤セットの連続制限（3連続の禁止または緩和）
            w_night_3_consec = get_penalty_weight(opt_night_3_consec)
            for_1 * 100)

            w_sub_only = get_penalty_weight(opt_ e in range(num_staff):
                target_weight = staff_comp_lvl[e]
                sub_only)
            for d in range(num_days):
                leadership_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(for d in range(num_days - 6):
                    if w_night_3_consec != -1 and target_weight > 0:
                        n3c_var = model.NewBoolVar('')
e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if w_sub_only == -1:
                    model.Add(leadership_score                        model.Add(shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(e, d+6, 'D')] == 3).OnlyEnforceIf(n >= 2)
                else:
                    model.Add(leadership_score >= 1)
                    sub_var = model.NewBoolVar('')
                    model.Add(leadership_score == 1).OnlyEnforce3c_var)
                        model.Add(shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(e, d+6, 'D')] <= 2).If(sub_var)
                    penalties.append(sub_var * w_sub_only *OnlyEnforceIf(n3c_var.Not())
                        penalties.append(n3c 100)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    for d_var * w_night_3_consec * target_weight * 100)
                    else in range(num_days):
                        col_idx = 6 + d
                        if col_idx < tr.shape[1]:
                            cell_value = str(tr.iloc[0, col_idx]).:
                        model.Add(shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(e, d+6, 'D')] <= 2)

            w_minus_1 = get_penalty_weight(opt_minus_1)
            for d in range(numstrip()
                            if cell_value == "公":
                                model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum((shifts[(e, d, 'A')] + shifts[(_off_days[e]))
                if staff_night_ok[e] != "×":
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

            w_4_days = get_penalty_weighte, d, 'A残')]) for e in range(num_staff) if "新人" not in str(opt_4_days)
            w_night_3 = get_penalty_weight(opt_night_3)
            
            for e in range(num_staff):
                target_weight = staff_(staff_roles[e]))
                req = day_req_list[d]
                is_sun = ('日' in weekdays[d])
                is_abs = (absolute_req_list[d] == "〇")

                ifcomp_lvl[e]
                for d in range(num_days - 3):
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + is_sun:
                    model.Add(act_day <= req)
                    if is_abs or w_minus_1 == -1: model.Add(act_day == req)
                    else:
                         shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <=model.Add(act_day >= req - 1)
                        m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day != req - 1).OnlyEnforceIf(m_var. 3)
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                        
                    if w_4_days != -1 and target_weight > 0:
                        if d < num_days - 4:
                            model.Add(work(Not())
                        penalties.append(m_var * w_minus_1 * 100)
                else:
                    model.Add(act_day <= req + 1)
                    if isd) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf_abs or w_minus_1 == -1: model.Add(act_day >= req)
                    (p_var)
                        model.Add(work(d) + work(d+1) + workelse:
                        model.Add(act_day >= req - 1)
                        m_var = model(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * w_4_days * target_weight * 100)
                    else:
                        model.Add(work(d) + work(d.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day != req - 1).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * w_minus_1 * 100)

            w_sub_only = get_penalty_weight(opt_sub_only)
            for d in range(num_days):
                l_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_+1) + work(d+2) + work(d+3) <= 3)

                    if w_night_3 != -1 and target_weight > 0:
                        np_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Notstaff))
                if w_sub_only == -1: model.Add(l_score >= 2)
                else:
                    model.Add(l_score >= 1)
                    s_var = model.NewBoolVar('')())
                        final_p = model.NewIntVar(0, w_night_3 * target_weight * 100, '')
                        model.AddMultiplicationEquality(final_p, [np_var,
                    model.Add(l_score == 1).OnlyEnforceIf(s_var)
                    penalties.append(s_var * w_sub_only * 100)

            for shifts[(e, d+3, 'D')]])
                        penalties.append(final_p)
                    else:
                        model.Add(work(d) + work(d+1) + work( e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    for d in ranged+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            w_ot_consec = get_penalty_weight(opt_ot_consec)
            (num_days):
                        col_idx = 6 + d
                        if col_idx < tr.for e in range(num_staff):
                for d in range(num_days - 1):
shape[1]:
                            if str(tr.iloc[0, col_idx]).strip() == "公": model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days))                    if w_ot_consec == -1:
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')
                        model.Add(shifts[(e, d, == int(staff_off_days[e]))
                if staff_night_ok[e] != " 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(ot_var * w_ot_consec×": model.Add(sum(shifts[(e, d, 'D')] for d in range(num_ * 100)

            mid_day = num_days // 2
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    d_first = sum(shifts[(e, d, 'D')] for d in range(mid_day))
                    d_second = sum(shifts[(e, d, 'D')] for d in range(mid_day, num_days))
                    diff_d = model.NewIntVar(-100, 100, '')
                    abs_diff_d = model.NewIntVar(0, 100, '')
                    days)) <= int(staff_night_limits[e]))

            w_4_days = get_penalty_weight(opt_4_days)
            w_night_3 = get_penalty_weight(opt_night_3)
            
            for e in range(num_staff):
                target_weight = staff_comp_lvl[e]
                for d in range(num_days - 3):
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公model.Add(diff_d == d_first - d_second)
                    model.AddAbsEquality(abs_diff_d, diff_d)
                    penalties.append(abs_diff_d *')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] <= 3)
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                        
                    if w_4_days != -1 and target_weight > 0:
                        if d < num_days - 4: model.Add(work(d) + work( 50)
                
                if staff_overtime_ok[e] != "×":
                    ot_first = sum(shifts[(e, d, 'A残')] for d in range(mid_dayd+1) + work(d+2) + work(d+3) + work(d+4))
                    ot_second = sum(shifts[(e, d, 'A残')] for d in range(mid_day, num_days))
                    diff_ot = model.NewIntVar(-100, ) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf100, '')
                    abs_diff_ot = model.NewIntVar(0, 100, '')
                    model.Add(diff_ot == ot_first - ot_second)
                    model.(p_var)
                        model.Add(work(d) + work(d+1) + workAddAbsEquality(abs_diff_ot, diff_ot)
                    penalties.append(abs_diff_ot * 5(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * w_4_days * target_weight * 100)
                    else:
                        model.Add(work(d) + work(d0)

            total_night_req = sum(night_req_list)
            night_staff_count = sum(1 for ok in staff_night_ok if ok != "×")
            if total_night_req > 0 and night_staff_count > 0:
                for e in range(num+1) + work(d+2) + work(d+3) <= 3)

                    if w_night_3 != -1 and target_weight > 0:
                        np_var = model._staff):
                    if staff_night_ok[e] != "×":
                        act_n = sum(shifts[(e, d, 'D')] for d in range(num_days))
                        diff_n = model.NewIntVar(-10000, 10000, '')
                        abs_diff_n = model.NewIntVar(0, 10000, '')
                        model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(Add(diff_n == (act_n * night_staff_count) - total_night_req)
                        model.AddAbsEquality(abs_diff_n, diff_n)
                        penalties.d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Not())append(abs_diff_n)

            total_ot_req = sum(overtime_req_list); total_day_req = sum(day_req_list) 
            if total_ot_req >
                        final_p = model.NewIntVar(0, w_night_3 * target_weight * 100, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        act_d = sum(shifts d+3, 'D')]])
                        penalties.append(final_p)
                    else:[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        act_o = sum(shifts[(e, d, 'A残')] for d
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            w_ot_consec = get_ in range(num_days))
                        diff = model.NewIntVar(-10000, 10000, '')
                        abs_diff = model.NewIntVar(0, 1000penalty_weight(opt_ot_consec)
            for e in range(num_staff):
                for d in range(num_days - 1):
                    if w_ot_consec == -10, '')
                        model.Add(diff == (act_o * total_day_req) - (act_d * total_ot_req))
                        model.AddAbsEquality(abs_diff, diff): model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')

                        penalties.append(abs_diff)
            
            if penalties: model.Minimize(sum                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0
            solver.parameters.random_seed = random_seed
            return (solver, shifts) if solver.Solve(model) in [cp_model.OPTIMAL, cp_model.FE(ot_var * w_ot_consec * 100)

            mid_day = num_days // 2
            for e in range(num_staff):
                if staff_night_okASIBLE] else (None, None)


        if st.button("設定に基づき、シフトを【3[e] != "×":
                    diff_d = model.NewIntVar(-100, 100パターン】作成する！"):
            with st.spinner('AIが優先順位とバランスを計算し、3パターンのシフトを考えています...（最大3分）'):
                results = [res for seed in [1, 42, ''); abs_diff_d = model.NewIntVar(0, 100, '')
                    model.Add(diff_d == sum(shifts[(e, d, 'D')] for d in range(mid, 99] if (res := solve_shift(seed))[0]]
                if not results: st.error("❌ 条件が厳しすぎます。設定画面で緩和する条件の「優先順位」を選択_day)) - sum(shifts[(e, d, 'D')] for d in range(mid_day,してください！")
                else:
                    st.success(f"✨完成！ {len(results)}パターン提案します！✨")
                    cols = []
                    for d_val, w_val in zip(date_columns, weekdays):
                        try:
                            dt = datetime.date(target_year, target_month num_days)))
                    model.AddAbsEquality(abs_diff_d, diff_d)
                    penalties.append(abs_diff_d * 50)
                if staff_overtime_ok[e] != "×":
                    diff_ot = model.NewIntVar(-100, 100, ''); abs, int(d_val))
                            if jpholiday.is_holiday(dt): cols.append(f"{d_val}({w_val}・祝)")
                            else: cols.append(f"{d_val}({w__diff_ot = model.NewIntVar(0, 100, '')
                    model.Add(diff_ot == sum(shifts[(e, d, 'A残')] for d in range(mid_dayval})")
                        except ValueError:
                            cols.append(f"{d_val}({w_val})")

                    tabs = st.tabs([f"パターン {i+1}" for i in range(len(results))])
                    
                    for i, (solver, shifts) in enumerate(results):
                        with)) - sum(shifts[(e, d, 'A残')] for d in range(mid_day, num_days)))
                    model.AddAbsEquality(abs_diff_ot, diff_ot)
                    penalties.append(abs_diff_ot * 50)

            total_night_req = sum(night_req_ tabs[i]:
                            data = []
                            for e in range(num_staff):
                                row = {"スタッフ名": staff_names[e]}
                                for d in range(num_days):
                                    list)
            night_staff_count = sum(1 for ok in staff_night_ok if ok != "×for s in ['A', 'A残', 'D', 'E', '公']:
                                        if solver")
            if total_night_req > 0 and night_staff_count > 0:
                .Value(shifts[(e, d, s)]):
                                            if (s == 'A' or s == 'A残') and str(staff_part_shifts[e]).strip() not in ["", "nan"]:
                               for e in range(num_staff):
                    if staff_night_ok[e] != "×":
                        act_n = sum(shifts[(e, d, 'D')] for d in range(num_                 row[cols[d]] = str(staff_part_shifts[e]).strip()
                                            else:
                                                row[cols[d]] = s
                                data.append(row)
                                
                            df_res = pd.days))
                        diff_n = model.NewIntVar(-10000, 10000, ''); abs_diff_n = model.NewIntVar(0, 10000, '')
                        model.Add(diff_n == (act_n * night_staff_count) - total_night_req)
                        modelDataFrame(data)

                            # 🌟 消してしまった必須集計欄の完全復活
                            df_res['日勤(A・P)回数'] = df_res[cols].apply(lambda x: x.str.contains('A|P|Ｐ', na=False) & ~x.str.contains('残', na=False)).sum(axis=1).AddAbsEquality(abs_diff_n, diff_n)
                        penalties.append(abs_diff_n)

            total_ot_req = sum(overtime_req_list); total_day_req = sum(day_req_list) 
            if total_ot_req > 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[
                            df_res['残業(A残)回数'] = (df_res[cols] == 'A残').sum(axis=1)
                            df_res['残業割合'] = df_rese] != "×":
                        act_d = sum(shifts[(e, d, 'A')] +.apply(lambda r: f"{(r['残業(A残)回数']/r['日勤 shifts[(e, d, 'A残')] for d in range(num_days))
                        act_o = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        (A・P)回数'])*100:.1f}%" if r['日勤(A・P)回数']>0 else "0.0%", axis=1)
                            df_res['夜diff = model.NewIntVar(-10000, 10000, ''); abs_diff = model.NewIntVar(0, 10000, '')
                        model.Add(diff ==勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                            df_res['公休回数'] = (df_res[cols] == '公'). (act_o * total_day_req) - (act_d * total_ot_req))
                        model.AddAbsEquality(abs_diff, diff)
                        penalties.append(abs_diff)
            
            if penaltiessum(axis=1)
                            df_res['日曜D回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'D') if staff_sun_d[e] == "〇" else : model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0
            solver.parameters.random0 for e in range(num_staff)]
                            df_res['日曜E回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res_seed = random_seed
            return (solver, shifts) if solver.Solve(model) in [cp.loc[e, cols[d]] == 'E') if staff_sun_e[e] == "〇" else 0 for e in range(num_staff)]

                            # 🌟 下部の集計行_model.OPTIMAL, cp_model.FEASIBLE] else (None, None)

        if st.button("設定に基づき、シフトを【3パターン】作成する！"):
            with st.spinner('AIが優先順位とバランスを計算し、3パターンのシフトを考えています...（最大3分の完全復活
                            sum_A = {"スタッフ名": "【日勤(A・P) 合計】"}
                            sum_Az = {"スタッフ名": "【残業(A残) 合計】"}
）'):
                results = [res for seed in [1, 42, 99] if (                            sum_D = {"スタッフ名": "【夜勤(D) 合計】"}
                            sum_res := solve_shift(seed))[0]]
                if not results: st.error("❌ 条件が厳しすぎます。設定画面で緩和する条件の「優先順位」を選択してください！")
                else:
                    st.success(f"✨完成！ {len(results)}パターン提案します！✨")
                    Off = {"スタッフ名": "【公休 合計】"}
                            
                            for c in ['日勤(A・P)回数', '残業(A残)回数', '残業割合', '夜勤(D)回数', '公休回数', '日曜D回数', '日曜E回数']:
                                sum_A[c] = ""; sum_Az[c] = ""; sum_D[c] = ""; sum_Off[c] = ""

                            for d, c in enumerate(cols):
cols = []
                    for d_val, w_val in zip(date_columns, weekdays):
                                                        sum_A[c] = sum(1 for e in range(num_staff) if str(df_res.loc[e, c]) in ['A', 'A残'] or 'P' in str(df_try:
                            dt = datetime.date(target_year, target_month, int(d_val))
                            if jpholiday.is_holiday(dt): cols.append(f"{d_val}({res.loc[e, c]) and "新人" not in str(staff_roles[e]))
                                sum_Az[c] = (df_res[c] == 'A残').sum()
                                sumw_val}・祝)")
                            else: cols.append(f"{d_val}({w_val})")
                        except ValueError:
                            cols.append(f"{d_val}({w_val_D[c] = (df_res[c] == 'D').sum()
                                sum_Off[c] = (df_res[c] == '公').sum()

                            df_fin = pd.concat([df_res, pd.DataFrame([sum_A, sum_Az, sum_D, sum})")

                    tabs = st.tabs([f"パターン {i+1}" for i in range(len(results))])
                    
                    for i, (solver, shifts) in enumerate(results):
                        with_Off])], ignore_index=True)

                            def highlight_warnings(df):
                                styles = pd.DataFrame('', index=df.index, columns=df.columns)
                                for d, col_name in tabs[i]:
                            data = []
                            for e in range(num_staff):
                                row = {"スタッフ名": staff_names[e]}
                                for d in range(num_days):
                                    for s in ['A', 'A残', 'D', 'E', '公']:
                                        if solver enumerate(cols):
                                    actual_a = df.loc[len(staff_names), col_name]
                                    target_a = day_req_list[d]
                                    if actual_a != "":
                                        if actual_a < target_a:
                                            styles.loc[len(staff_names), col_name] = 'background-color: #FFCCCC; color: red; font-weight:.Value(shifts[(e, d, s)]):
                                            if (s == 'A' or s == 'A残') and str(staff_part_shifts[e]).strip() not in ["", "nan"]:
                                                row[cols[d]] = str(staff_part_shifts[e]).strip()
                                            else:
                                                row[cols[d]] = s
                                data.append(row)
                                
                            df_res = pd bold;'
                                        elif actual_a > target_a:
                                            styles.loc[len(staff_names), col_name] = 'background-color: #CCFFFF; color: blue; font-weight.DataFrame(data)

                            # 🌟 集計列の完全復活（日勤回数はAとA残と: bold;'

                                for e in range(num_staff):
                                    for d in range(num_days):
                                        def is_day_work(day_idx):
                                            if day_idx >=Pの合計）
                            df_res['日勤(A/P)回数'] = df_res[cols].apply(lambda x num_days: return False
                                            v = str(df.loc[e, cols[day_idx]])
                                            return v == 'A' or v == 'A残' or 'P' in v or: x.str.contains('A|P|Ｐ', na=False)).sum(axis=1)
                            df_res['残業(A残)回数'] = (df_res[cols] == 'A残').sum(axis=1)
                            df_res['残業割合(%)'] = df_res.apply(lambda r: 'Ｐ' in v

                                        if is_day_work(d) and is_day_work(d+1) and is_day_work(d+2) and is_day_work(d+3):
                                            styles.loc[e, cols[d]] = 'background-color: #FFFF99;'
                                            styles. f"{(r['残業(A残)回数']/r['日勤(A/P)回数'])*100:.1f}%" if r['日勤(A/P)回数']>loc[e, cols[d+1]] = 'background-color: #FFFF99;'
                                            styles.loc[e, cols[d+2]] = 'background-color: #FFFF99;'
0 else "0.0%", axis=1)
                            df_res['夜勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                            df_res['公休回数'] =                                            styles.loc[e, cols[d+3]] = 'background-color: #FFFF99;'

                                        if d + 3 < num_days:
                                            if is_day_work(d) and is_day (df_res[cols] == '公').sum(axis=1)
                            df_res['日曜D回数'] = [sum(1 for d in range(num_days) if '日' in weekdays_work(d+1) and is_day_work(d+2) and str(df.loc[e, cols[d+3]]) == 'D':
                                                styles.loc[e, cols[[d] and df_res.loc[e, cols[d]] == 'D') if staff_sun_d[e] == "〇" else 0 for e in range(num_staff)]
                            df_res['日曜E回数'] = [sum(1 for d in range(num_days) if 'd]] = 'background-color: #FFD580;'
                                                styles.loc[e, cols[d+1]] = 'background-color: #FFD580;'
                                                styles.loc[e, cols[d+2]] = 'background-color: #FFD580;'
日' in weekdays[d] and df_res.loc[e, cols[d]] == 'E')                                                styles.loc[e, cols[d+3]] = 'background-color: #FFD580;'
                                return styles

                            st.dataframe(df_fin.style.apply(highlight_warnings if staff_sun_e[e] == "〇" else 0 for e in range(num_staff)]

                            # 🌟 下部の集計行の完全復活
                            sum_A = {"スタッフ名": "【日勤(A/, axis=None))
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                df_fin.to_excel(writer, indexP) 合計】"}
                            sum_Az = {"スタッフ名": "【残業(A残) 合計】"}
                            sum_D = {"スタッフ名": "【夜勤(D) 合計】=False, sheet_name='完成シフト')
                            processed_data = output.getvalue()
                            
                            st.download_button(
                                label=f"📥 【パターン {i+1}】 をエクセル"}
                            sum_O = {"スタッフ名": "【公休 合計】"}
                            
                            for c in ['日勤(A/P)回数', '残業(A残)回数', '残業割合(%)', '夜勤(D)回数', '公休回数', '日曜Dでダウンロード（色なし）",
                                data=processed_data,
                                file_name=f"完成版_パターン{i+1}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_btn_{i}"
                            )
                    回数', '日曜E回数']:
                                sum_A[c] = ""; sum_Az[c
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白] = ""; sum_D[c] = ""; sum_O[c] = ""

                            for d, c in enumerate(cols):
                                a_count = 0
                                for e in range(num_staff):
                                    val = str(df_res.loc[e, c])
                                    if (val == 'A' or val == 'A残' or "P" in val or "Ｐ" in val) and "新人" notの行があります。({e})")
