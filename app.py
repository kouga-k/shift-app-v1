import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🤝 AIシフト作成 Co-Pilot (フェーズ23：多様性MAX＆完全版)")
st.write("夜勤だけでなく、残業（A残）や公休の配置もパターンごとに劇的に変わるようにしました！")

# 状態管理
if 'needs_compromise' not in st.session_state:
    st.session_state.needs_compromise = False

st.write("---")
today = datetime.date.today()
col_y, col_m = st.columns(2)
with col_y: target_year = st.selectbox("作成年", [today.year, today.year + 1], index=0)
withれてしまいましたので、今回は出力を少しスリムにし、**絶対に途切れないよう最後まで完全出力**します！

---

### 🛠️ アプリの修正（残業・公休のランダム拡張版）

GitHubの `app.py` を開き、以下のコードに**すべて丸ごと上書き**してください。

▼ ここから下をすべてコピー ▼
```python
import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🤝 AIシフト作成 Co-Pilot col_m: target_month = st.selectbox("作成月", list(range(1, 13)), index=(today.month % 12))
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
            for i in (フェーズ24：残業・夜勤の完全ランダム化)")
st.write("夜勤だけでなく「残業(A残)」や「公休」の配置にも揺らぎを与え、全く違う3パターンを提案します！")

# 状態管理
if 'needs_compromise' not in st.session_state:
    st.session_state.needs_compromise = False

st.write("---")
today = datetime.date.today()
col_y, col_m = st.columns(2)
with col_y: target_year = st.selectbox("作成年", [today.year, today.year + 1], range(num_staff):
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
 index=0)
with col_m: target_month = st.selectbox("作成月", list(range(1, 13)), index=(today.month % 12))
st.write("---")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・前月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="日別設定")
        
        staff_names = df_staff["        
        staff_night_limits = [0 if ok == "×" else int(v) if pd.notna(v) else 10 for ok, v in zip(staff_night_ok, get_staff_col("夜勤上限", 10, is_int=True))]
        staff_sun_d = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Dカウント", "〇"))]
        staff_sun_e = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Eカウント", "〇"))]

        staff_comp_lvl = []スタッフ名"].dropna().tolist()
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
        staff_off_days = get_staff_col("公休数", 8, is_int=True)
        staff_night_ok = get_staff_col("夜勤可否", "〇")
        staff_overtime_
        for i in range(num_staff):
            val = ""
            if "妥協優先度" in df_staff.columns and pd.notna(df_staff["妥協優先度"].iloc[i]): val = str(df_staff["妥協優先度"].iloc[i]).strip()
            elif "連勤妥協OK" in df_staff.columns and pd.notna(df_staff["連勤妥協OK"].iloc[i]): val = str(df_staff["連勤妥協OK"].iloc[i]).strip()
            
            if val in ["〇", "1", "1.0"]: staff_comp_lvl.append(1)
            elif val in ["2", "2.0"]: staff_comp_lvl.append(2)
            elif val in ["3", "3.0"]: staff_comp_lvl.append(3)
            else: staff_comp_lvl.append(0)

        date_columns = [col for col in df_req.columns if col != df_req.columns[0]ok = get_staff_col("残業可否", "〇")
        staff_part_shifts = get_staff_col("パート", "")
        
        staff_night_limits = [0 if ok == "×" else int(v) if pd.notna(v) else 10 for ok, v in zip(staff_night_ok, get_staff_col("夜勤上限", 10, is_int=True))]
        staff_sun_d = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Dカウント", "〇"))]
        staff_sun_e = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Eカウント", "〇"))]

        staff_comp_lvl = []
        for i in range(num_staff):
            val = ""
            if "妥協優先度" in df_staff.columns and pd.notna(df and not str(col).startswith("Unnamed")]
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
        absolute_req_list = get_staff["妥協優先度"].iloc[i]): val = str(df_staff["妥協優先度"].iloc[i]).strip()
            elif "連勤妥協OK" in df_staff.columns and pd.notna(df_staff["連勤妥協OK"].iloc[i]): val = str(df_staff["連勤妥協OK"].iloc[i]).strip()
            
            if val in ["〇", "1", "1.0"]: staff_comp_lvl.append(1)
            elif val in ["2", "2.0"]: staff_comp_lvl.append(2)
            elif val in ["3", "3.0"]:_req_col("絶対確保", "", is_int=False)

        weekdays = [str(df_req.iloc[0, d+1]).strip() if (d+1) < len(df_req.columns) and pd.notna(df_req.iloc[0, d+1]) else "" for d in range(num_days)]

        st.success("✅ データの読み込み完了！まずは妥協なしの「理想のシフト」を作れるかテストします。")

        def solve_shift(random_seed, allow_minus_1=False, allow_4_days=False, allow_night_3=False, allow_sub_only=False, allow_ot_consec=False, allow_night_consec_3=False):
            model = cp_model.CpModel()
            types = ['A', staff_comp_lvl.append(3)
            else: staff_comp_lvl.append(0)

        date_columns = [col for col in df_req.columns if col != df_req.columns[0] and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        def get_req_col(label, default_val, is_int=True):
            row = df_req[df_req.iloc[:, 0] == label]
            res = []
            for d in range(num_days): 'A残', 'D', 'E', '公']
            shifts = {(e, d, s): model.NewBoolVar('') for e in range(num_staff) for d in range(num_days) for s in types}
            
            # ランダムシードの設定（これでパターンごとの動きを変える）
            random.seed(random_seed)

            for e in range(num_staff):
                for d in range(num_
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

days):
                    model.AddExactlyOne(shifts[(e, d, s)] for s in types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 0); model.Add(shifts[(e, d, 'E')] == 0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days): model.Add(shifts[(e, d, 'A残')] == 0)

            # 前月履歴
            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    last_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_day == "D":
                        model.Add(shifts[(e, 0, 'E')] == 1)
                        if num_days > 1: model.Add(shifts[(e, 1, '公')] == 1)
                    elif last_day == "E":
                        model.Add(shifts[(        st.success("✅ データの読み込み完了！まずは妥協なしの「理想のシフト」を作れるかテストします。")

        def solve_shift(random_seed, allow_minus_1=False, allow_4_days=False, allow_night_3=False, allow_sub_only=False, allow_ot_consec=False, allow_night_consec_3=False):
            model = cp_model.CpModel()
            types = ['A', 'A残', 'D', 'E', '公']
            shifts = {(e, d, s): model.NewBoolVar('') for e in range(num_staff) for d in range(num_days) for s in types}
            
            random.seed(randome, 0, '公')] == 1)

            # 夜勤セットの絶対ルール
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                        if l_day != "D": model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0: model.Add(shifts[(e, d, '_seed)

            for e in range(num_staff):
                for d in range(num_days):
                    model.AddExactlyOne(shifts[(e, d, s)] for s in types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days):
                        model.Add(shifts[(e, d, 'D')] == 0); model.Add(shifts[(e, d, 'E')] == 0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days): model.Add(shifts[(e, d, 'A残E')] == shifts[(e, d-1, 'D')])
                        if d + 1 < num_days: model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

            penalties = []
            
            # 夜勤ループと3連続防止
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    for d in range(num_days - 3): model.Add(shifts[(e, d, 'E')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, 'D')] <= 3)
                    for d in range(num_days - 4): model.Add(shifts[(e, d, 'E')] +')] == 0)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    last_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_day == "D":
                        model.Add(shifts[(e, 0, 'E')] == 1)
                        if num_days > 1 shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e, d+3, '公')] + shifts[(e, d+4, 'D')] <= 4)
                    
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty and tr.shape[1] > 5:
                        l_5 = [str(tr.iloc[0, i]).strip() for i in range(1, 6)]
                        if l_5[4] == "E":
                            if num_days >: model.Add(shifts[(e, 1, '公')] == 1)
                    elif last_day == "E":
                        model.Add(shifts[(e, 0, '公')] == 1)

            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                        if l_day != "D": model.Add(shifts[(e, 0, 'E')] == 0)
                    for d in range(num_days):
                        if d > 0: model.Add(shifts[(e, d, 'E')] == 2: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, '公')] + shifts[(e, 2, 'D')] <= 2)
                            if num_days > 3: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, '公')] + shifts[(e, 2, '公')] + shifts[(e, 3, 'D')] <= 3)
                        if l_5[3] == "E" and l_5[4] == "公":
                            if num_days > 1: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, 'D')] <= 1)
                            if num_days > 2: model.Add(shifts[(e, 0, '公')] + shifts[(e, 1, '公')] + shifts[(e, 2, 'D')] <= 2)

            for e, staff_name in enumerate(staff_names):
                if staff_night_ok[e] != "×":
                    past_D = [0] * 5
                    tr = df shifts[(e, d-1, 'D')])
                        if d + 1 < num_days: model.AddImplication(shifts[(e, d, 'E')], shifts[(e, d+1, '公')])

            penalties = []
            
            for e, staff_name in enumerate(staff_names):
                if staff_night_ok[e] != "×":
                    past_D = [0] * 5
                    tr = df_history[df_history.iloc[:, 0] == staff_name]
                    if not tr.empty:
                        for i in range(5):
                            if (i+1) < tr.shape[1] and str(tr.iloc[0, i+1]).strip_history[df_history.iloc[:, 0] == staff_name]
                    if not tr.empty:
                        for i in range(5):
                            if (i+1) < tr.shape[1] and str(tr.iloc[0, i+1]).strip() == "D": past_D[i] = 1
                    
                    all_D = past_D + [shifts[(e, d, 'D')] for d in range(num_days)]
                    for i in range(len(all_D) - 6):
                        window = all_D[i : i+7]
                        if not allow_night_consec_3:
                            if any(isinstance(x, cp_model.IntVar) for x in window): model.() == "D": past_D[i] = 1
                    
                    all_D = past_D + [shifts[(e, d, 'D')] for d in range(num_days)]
                    for i in range(len(all_D) - 6):
                        window = all_D[i : i+7]
                        if not allow_night_consec_3:
                            if any(isinstance(x, cp_model.IntVar) for x in window): model.Add(sum(window) <= 2)
                        else:
                            if any(isinstance(x, cp_model.IntVar) for xAdd(sum(window) <= 2)
                        else:
                            if any(isinstance(x, cp_model.IntVar) for x in window):
                                n3_var = model.NewBoolVar('')
                                model.Add(sum(window) >= 3).OnlyEnforceIf(n3_var)
                                model.Add(sum(window) <= 2).OnlyEnforceIf(n3_var.Not())
                                penalties.append(n3_var * 5000)

            # 日勤人数の誘導ロジック
            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d]) in window):
                                n3_var = model.NewBoolVar('')
                                model.Add(sum(window) >= 3).OnlyEnforceIf(n3_var)
                                model.Add(sum(window) <= 2).OnlyEnforceIf(n3_var.Not())
                                penalties.append(n3_var * 5000)

            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum((shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                req = day_req_list[d]
                is_sun = ('日' in weekdays[d])
                is_abs = (absolute_req_list[d] == "〇")

                if is_abs:
                    model.Add(act_day >= req)
                    over_var = model.NewIntVar(0, 100, '')
                    diff = model.NewIntVar(-100, 100, '')
残')] for e in range(num_staff)) == overtime_req_list[d])
                
                act_day = sum((shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                req = day_req_list[d]
                is_sun = ('日' in weekdays[d])
                is_abs = (absolute_req_list[d] == "〇")

                if is_abs:
                    model.Add(act_day >= req)
                    over_var = model.NewIntVar(0, 100, '')
                    diff = model.NewIntVar(-100, 100, '')
                    model.Add(diff == act_day - req)
                    model.AddMaxEquality(over_var, [0, diff])
                    penalties.append(over_var * 1) 
                elif is_sun:
                    model.Add(act_day <= req)
                    if not allow                    model.Add(diff == act_day - req)
                    model.AddMaxEquality(over_var, [0, diff])
                    penalties.append(over_var * 1) 
                elif is_sun:
                    model.Add(act_day <= req)
                    if not allow_minus_1: model.Add(act_day == req)
                    else:
                        model.Add(act_day >= req - 1)
                        m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day == req).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * 1000)
                else:
                    if not allow_minus_1: model.Add(act_day >= req)
                    else:
                        model.Add(act_day >= req - 1)
                        m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day != req -_minus_1: model.Add(act_day == req)
                    else:
                        model.Add(act_day >= req - 1)
                        m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day == req).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * 1000)
                else:
                    if not allow_minus_1: model.Add(act_day >= req)
                    else:
                        model.Add( 1).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * 1000)
                    
                    over_var = model.NewIntVar(0, 100, '')
                    diff = model.NewIntVar(-100, 100, '')
                    model.Add(diff == act_day - req)
                    model.AddMaxEquality(over_var, [0, diff])
                    penalties.append(over_var * 100)

                l_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0)act_day >= req - 1)
                        m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day != req - 1).OnlyEnforceIf(m_var.Not())
                        penalties.append(m_var * 1000)
                    
                    over_var = model.NewIntVar(0, 100, '')
                    diff = model.NewIntVar(-100, 100, '')
                    model.Add(diff == act_day - req)
                    model.AddMaxEquality(over_var, [0, diff])
                    penalties.append(over_var * 100)

                l_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if " * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if not allow_sub_only: model.Add(l_score >= 2)
                else:
                    model.Add(l_score >= 1)
                    sub_var = model.NewBoolVar('')
                    model.Add(l_score == 1).OnlyEnforceIf(sub_var)
                    penalties.append(sub_var * 1000)

            # 希望休・回数ノルマ
            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    for d in range(num_days):
                        col_idx = 6 + d
                        if col_idx < tr.shape[1]:
                            if str(tr.iloc[0, col_idx]).strip() == "公サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                if not allow_sub_only: model.Add(l_score >= 2)
                else:
                    model.Add(l_score >= 1)
                    sub_var = model.NewBoolVar('')
                    model.Add(l_score == 1).OnlyEnforceIf(sub_var)
                    penalties.append(sub_var * 1000)

            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                ": model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×":
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

            # 🌟 夜勤回数の厳格な公平化
            limit_groups = {}
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    limit = intif not tr.empty:
                    for d in range(num_days):
                        col_idx = 6 + d
                        if col_idx < tr.shape[1]:
                            if str(tr.iloc[0, col_idx]).strip() == "公": model.Add(shifts[(e, d, '公')] == 1)

            for e in range(num_staff):
                model.Add(sum(shifts[(e, d, '公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×":
                    model.Add(sum(shifts[(e, d, 'D')] for d in range(num_days)) <= int(staff_night_limits[e]))

            limit_groups = {}
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    limit = int(staff_night_limits[(staff_night_limits[e])
                    if limit > 0:
                        if limit not in limit_groups: limit_groups[limit] = []
                        limit_groups[limit].append(e)
            for limit, members in limit_groups.items():
                if len(members) >= 2:
                    actual_nights = [sum(shifts[(m, d, 'D')] for d in range(num_days)) for m in members]
                    max_n = model.NewIntVar(0, limit, ''); min_n = model.NewIntVar(0, limit, '')
                    model.AddMaxEquality(max_n, actual_nights); model.AddMinEquality(min_n, actual_nights)
                    model.Add(max_n - min_n <= 1)

            # 思いやりの連休コントロール
            for e in range(num_staff):
                for d in range(num_days - 3): model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] + shifts[(e])
                    if limit > 0:
                        if limit not in limit_groups: limit_groups[limit] = []
                        limit_groups[limit].append(e)
            for limit, members in limit_groups.items():
                if len(members) >= 2:
                    actual_nights = [sum(shifts[(m, d, 'D')] for d in range(num_days)) for m in members]
                    max_n = model.NewIntVar(0, limit, ''); min_n = model.NewIntVar(0, limit, '')
                    model.AddMaxEquality(max_n, actual_nights); model.e, d+3, '公')] <= 3)

                for d in range(num_days - 2):
                    is_3_off = model.NewBoolVar('')
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] == 3).OnlyEnforceIf(is_3_off)
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] + shifts[(e, d+2, '公')] <= 2).OnlyEnforceIf(is_3_off.Not())
                    penalties.append(is_3_off * 500AddMinEquality(min_n, actual_nights)
                    model.Add(max_n - min_n <= 1)

            for e in range(num_staff):
                target_lvl = staff_comp_lvl[e]
                w_base = 10 ** target_lvl if target_lvl > 0 else 0
                
                for d in range(num_days - 3):
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                        
                    if allow_4_days and target_lvl > 0:
                        if d <) 

                is_2_offs = []
                for d in range(num_days - 1):
                    is_2_off = model.NewBoolVar('')
                    model.Add(shifts[(e, d, '公')] + shifts[(e, d+1, '公')] == 2).OnlyEnforceIf(is_2_off)
                    model.Add(shifts[(e, d, '公')] + num_days - 4: model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf(p_var)
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * w_base)
                    else: shifts[(e, d+1, '公')] <= 1).OnlyEnforceIf(is_2_off.Not())
                    is_2_offs.append(is_2_off)
                
                has_any_2_off = model.NewBoolVar('')
                model.Add(sum(is_2_offs) >= 1).OnlyEnforceIf(has_any_2_off) 
                model.Add(sum(is_2_offs) == 0).OnlyEnforceIf(has_any_2_off.Not())
                penalties.append(has_any_2_off.Not() * 300) 

            for e in range(num_staff):
                target_lvl = staff_comp_lvl[e]
                w_base = 10 ** target_lvl if target_lvl > 0 else 0
                
                for d in range(num_days - 3):
                    def work(day): return shifts[(e, day, 'A')] + shifts[(e, day, '
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3)

                    if allow_night_3 and target_lvl > 0:
                        np_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Not())
                        final_p = model.NewIntVar(0, w_base, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, d+3, 'D')]])
                        penalties.append(final_p)
                    else:
                        model.A残')]
                        
                    if allow_4_days and target_lvl > 0:
                        if d < num_days - 4: model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                        p_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) == 4).OnlyEnforceIf(p_var)
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3).OnlyEnforceIf(p_var.Not())
                        penalties.append(p_var * w_base)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3)

                    if allow_night_3 and target_lvl > 0:
                        Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            for e in range(num_staff):
                for d in range(num_days - 1):
                    if not allow_ot_consec: model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.Newnp_var = model.NewBoolVar('')
                        model.Add(work(d) + work(d+1) + work(d+2) == 3).OnlyEnforceIf(np_var)
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(np_var.Not())
                        final_p = model.NewIntVar(0, w_base, '')
                        model.AddMultiplicationEquality(final_p, [np_var, shifts[(e, d+3, 'D')]])
                        penalties.append(final_pBoolVar('')
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf(ot_var)
                        penalties.append(ot_var * 500)

            mid_day = num_days // 2
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    diff_d = model.NewIntVar(-100, 100, ''); abs_diff_d = model.NewIntVar(0, 100, '')
                    model.Add(diff_d == sum(shifts[(e, d, 'D')] for d in range(mid_day)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) <= 2).OnlyEnforceIf(shifts[(e, d+3, 'D')])

            for e in range(num_staff):
                for d in range(num_days - 1):
                    if not allow_ot_consec: model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)
                    else:
                        ot_var = model.NewBoolVar('')
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] == 2).OnlyEnforceIf)) - sum(shifts[(e, d, 'D')] for d in range(mid_day, num_days)))
                    model.AddAbsEquality(abs_diff_d, diff_d)
                    penalties.append(abs_diff_d * 5)
                
                if staff_overtime_ok[e] != "×":
                    diff_ot = model.NewIntVar(-100, 100, ''); abs_diff_ot = model.NewIntVar(0, 100, '')
                    model.Add(diff_(ot_var)
                        penalties.append(ot_var * 500)

            mid_day = num_days // 2
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    diff_d = model.NewIntVar(-100, 100, ''); abs_diff_d = model.NewIntVar(0, 100, '')
                    model.Add(diff_d == sum(shifts[(e, d, 'D')] for d in range(mid_day)) - sum(shifts[(e, d, 'D')] for d in range(mid_day, num_days)))
                    model.AddAbsEquality(abs_diff_dot == sum(shifts[(e, d, 'A残')] for d in range(mid_day)) - sum(shifts[(e, d, 'A残')] for d in range(mid_day, num_days)))
                    model.AddAbsEquality(abs_diff_ot, diff_ot)
                    penalties.append(abs_diff_ot * 5)

            total_ot_req = sum(overtime_req_list); total_day_req = sum(day_req_list) 
            if total_ot_req > 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        act_d = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        act_o = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        diff = model.NewIntVar(-10000, 10000, ''); abs_diff = model.NewIntVar(0, 1, diff_d)
                    penalties.append(abs_diff_d * 5)
                
                if staff_overtime_ok[e] != "×":
                    diff_ot = model.NewIntVar(-100, 100, ''); abs_diff_ot = model.NewIntVar(0, 100, '')
                    model.Add(diff_ot == sum(shifts[(e, d, 'A残')] for d in range(mid_day)) - sum(shifts[(e, d, 'A残')] for d in range(mid_day, num_days)))
                    model.AddAbsEquality(abs_diff_ot, diff_ot)
                    penalties.append(abs_diff_ot * 5)

            total_ot_req = sum(overtime_req_list); total_day_req = sum(day_req_list) 
            if total_ot_req > 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        act_d = sum(shifts[(e,0000, '')
                        model.Add(diff == (act_o * total_day_req) - (act_d * total_ot_req))
                        model.AddAbsEquality(abs_diff, diff)
                        penalties.append(abs_diff)

            # 🌟 NEW: 夜勤・残業・日勤の「ランダムな揺らぎ（スパイス）」を全開にする！
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    # 夜勤 d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        act_o = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        diff = model.NewIntVar(-10000, 10000, ''); abs_diff = model.NewIntVar(0, 10000, '')
                        model.Add(diff == (act_o * total_day_req) - (act_d * total_ot_req))
                        model.AddAbsEquality(abs_diff, diff)
                        penaltiesの担当者にランダムな罰金(-3〜3)を与え、パターンごとに担当を変える
                    act_n = sum(shifts[(e, d, 'D')] for d in range(num_days))
                    penalties.append(act_n * random.randint(-3, 3))
                
                if staff_overtime_ok[e] != "×":
                    # 残業の担当者にもランダムな罰金(-2〜2)を与え、パターンごとに担当を変える
                    act_o = sum(shifts[(e, d,.append(abs_diff)

            # 🌟 NEW: パターンを劇的に変化させる「ランダムな揺らぎ（スパイス）」を強化！
            for e in range(num_staff):
                # 人ごとに「A残を好むか」「Dを好むか」「公休を好むか」のランダムな好みを設定（-2〜2点）
                ot_bias = random.randint(-2, 2)
                night_bias = random.randint(-2, 2)
                off_bias = random.randint(-2, 2)
                
                 'A残')] for d in range(num_days))
                    penalties.append(act_o * random.randint(-2, 2))

                for d in range(num_days):
                    # 毎日のAや公の配置自体も、細かく散らす
                    penalties.append(shifts[(e, d, 'A')] * random.randint(-1, 1))
                    penalties.append(shifts[(e, d, '公')] * random.randint(-1, 1))
            
            if penalties: model.Minimize(sum(penal# ペナルティとして足し込む（マイナス点ならAIはそのシフトを積極的に配置しようとする）
                if staff_overtime_ok[e] != "×":
                    penalties.append(sum(shifts[(e, d, 'A残')] for d in range(num_days)) * ot_bias)
                if staff_night_ok[e] != "×":
                    penalties.append(sum(shifts[(e, d, 'D')] for d in range(num_days)) * night_bias)
                ties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 30.0 
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
                        if solver: results.appendpenalties.append(sum(shifts[(e, d, '公')] for d in range(num_days)) * off_bias)
                
                # 日々の配置自体にもランダムな揺らぎ（-1〜1）を与える
                for d in range(num_days):
                    penalties.append(shifts[(e, d, 'A')] * random.randint(-1, 1))
            
            if penalties: model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 30.0 
            solver.parameters.random_seed = random_seed
            status = solver.Solve(model)
            
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE: return solver, shifts
            else: return None, None

        if not st.session_state.needs_compromise:
            if st.button("▶️ 【STEP 1】まずは妥((solver, shifts))
                        
                    if results:
                        st.success(f"🎉 なんと！妥協なしで完璧なシフトが {len(results)} パターン組めました！")
                    else:
                        st.session_state.needs_compromise = True
                        st.rerun()
        else:
            st.error("⚠️ 【AI店長からのご報告】\n申し訳ありません。現在の人数と希望休では、すべてのルールを完璧に守ってシフトを組むことは物理的に不可能でした...")
            st.warning("💡 以下のいずれかの「妥協案」を許可して、再計算を指示してください。（※妥協したくない協なしで理想のシフトを計算する（3パターン）"):
                with st.spinner('AIが「妥協なし」の完璧なシフトを3パターン模索中...'):
                    results = []
                    for seed in [1, 42, 99]:
                        solver, shifts = solve_shift(seed, False, False, False, False, False, False)
                        if solver: results.append((solver, shifts))
                        
                    if results:
                        st.success(f"🎉 なんと！妥協なしで完璧なシフトが {len(results)} パターン組めました！")
                    else:
                        st.session_state.needs_compromise = True
                        st.rerun()項目はチェックを外したままでOKです）")
            
            with st.container():
                st.markdown("### 📝 妥協の提案リスト")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**■ 人数と役割について**")
                    allow_minus_1 = st.checkbox("日勤人数の「マイナス1」を許可する（絶対確保日以外）")
                    allow_sub_only = st.checkbox("役割配置を「サブ1名＋他」まで下げることを許可する")
                with col2:
                    st.markdown("**■ 対象スタッフ（エクセルで1,2,3設定）への連勤お願い**")
                    allow_4_days = st.
        else:
            st.error("⚠️ 【AI店長からのご報告】\n申し訳ありません。現在の人数と希望休では、すべてのルールを完璧に守ってシフトを組むことは物理的に不可能でした...")
            st.warning("💡 以下のいずれかの「妥協案」を許可して、再計算を指示してください。（※妥協したくない項目はチェックを外したままでOKです）")
            
            with st.container():
                st.markdown("### 📝 妥協の提案リスト")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**■ 人数と役割について**")
                    allow_minus_1 = st.checkbox("日勤人数の「マイナス1」を許可する（絶対確保日以外）")
                    allow_sub_only = st.checkbox("役割配置をcheckbox("対象者への「最大4連勤」のお願いを許可する")
                    allow_night_3 = st.checkbox("対象者への「夜勤前3日連続日勤」のお願いを許可する")
                
                st.markdown("**■ その他の例外ルール**")
                col3, col4 = st.columns(2)
                with col3:
                    allow_night_consec_3 = st.checkbox("やむを得ない「月またぎ含む、夜勤セット3連続」を許可する")
                with col4:
                    allow_ot_consec = st.checkbox("やむを得ない「残業(A残)の2日連続」を許可する")

            if st.button("🔄 【STEP 「サブ1名＋他」まで下げることを許可する")
                with col2:
                    st.markdown("**■ 対象スタッフ（エクセルで1,2,3設定）への連勤お願い**")
                    allow_4_days = st.checkbox("対象者への「最大4連勤」のお願いを許可する")
                    allow_night_3 = st.checkbox("対象者への「夜勤前3日連続日勤」のお願いを許可する")
                
                st.markdown("**■ その他の例外ルール**")
                col3, col4 = st.columns(2)
                with col3:
                    allow_night_consec_3 = st.checkbox("やむを得ない「月またぎ含む、夜勤セット33】チェックした妥協案を許可して、3パターンのシフトを作る！"):
                with st.spinner('許可された妥協案をもとに、AIが再計算しています...'):
                    results = []
                    for seed in [1, 42, 99]:
                        solver, shifts = solve_shift(seed, allow_minus_1, allow_4_days, allow_night_3, allow_sub_only, allow_ot_consec, allow_night_consec_3)
                        if solver: results.append((solver, shifts))

                    if not results: st.error("😭 まだ条件が厳しすぎます！もう少しだけ他の妥協案も許可してもらえませんか？")
                    else:
                        st.success(f"✨ ありがとうございます！許可いただいた条件内で、{len(results)}パターンのシフトが完成しました！")
                        st.session_state.needs_compromise = False

        if 'results' in連続」を許可する")
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
                        st.success(f"✨ ありがとうございます！許可いただいた条件内で、{len( locals() and results:
            cols = []
            for d_val, w_val in zip(date_columns, weekdays):
                try:
                    dt = datetime.date(target_year, target_month, int(d_val))
                    if jpholiday.is_holiday(dt): cols.append(f"{d_val}({w_val}・祝)")
                    else: cols.append(f"{d_val}({w_val})")
                except ValueError: cols.append(f"{d_val}({w_val})")

            tabs = st.tabs([f"提案パターン {i+1}" for i in range(len(results))])
            for iresults)}パターンのシフトが完成しました！")
                        st.session_state.needs_compromise = False

        # --- 以下、画面描画処理（省略なし） ---
        if 'results' in locals() and results:
            cols = []
            for d_val, w_val in zip(date_columns, weekdays):
                try:
                    dt = datetime.date(target_year, target_month, int(d_val))
                    if jpholiday.is_holiday(dt): cols.append(f"{d_val}({w_val}・祝)")
                    else: cols.append(f"{d_val}({w_val})")
                except, (solver, shifts) in enumerate(results):
                with tabs[i]:
                    data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e]}
                        for d in range(num_days):
                            for s in ['A', 'A残', 'D', 'E', '公']:
                                if solver.Value(shifts[(e, d, s)]):
                                    if (s == 'A' or s == 'A残') and str(staff_part_shifts[e]).strip() not in ["", "nan"]: row[cols[d]] = str(staff_part_shifts[e]).strip()
                                    else: ValueError: cols.append(f"{d_val}({w_val})")

            tabs = st.tabs([f"提案パターン {i+1}" for i in range(len(results))])
            for i, (solver, shifts) in enumerate(results):
                with tabs[i]:
                    data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e]}
                        for d in range(num_days):
                            for s in ['A', 'A残', 'D', 'E', '公']:
                                if solver.Value(shifts[(e, d, s)]):
                                    if (s == 'A' or s == 'A残') and str(staff_part row[cols[d]] = s
                        data.append(row)
                        
                    df_res = pd.DataFrame(data)

                    df_res['日勤(A/P)回数'] = df_res[cols].apply(lambda x: x.str.contains('A|P|Ｐ', na=False) & ~x.str.contains('残', na=False)).sum(axis=1)
                    df_res['残業(A残)回数'] = (df_res[cols] == 'A残').sum(axis=1)
                    df_res['残業割合(%)'] = df_res_shifts[e]).strip() not in ["", "nan"]: row[cols[d]] = str(staff_part_shifts[e]).strip()
                                    else: row[cols[d]] = s
                        data.append(row)
                        
                    df_res = pd.DataFrame(data)

                    df_res['日勤(A/P)回数'] = df_res[cols].apply(lambda x: x.str.contains('A|P|Ｐ', na=False) & ~x.str.contains('残', na=False)).sum(axis=1)
                    df_res['残業(A残)回数'] = (df.apply(lambda r: f"{(r['残業(A残)回数']/r['日勤(A/P)回数'])*100:.1f}%" if r['日勤(A/P)回数']>0 else "0.0%", axis=1)
                    df_res['夜勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                    df_res['公休回数'] = (df_res[cols] == '公').sum(axis=1)
                    df_res['日曜D回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'D') if staff_sun_d[e] == "〇" else _res[cols] == 'A残').sum(axis=1)
                    df_res['残業割合(%)'] = df_res.apply(lambda r: f"{(r['残業(A残)回数']/r['日勤(A/P)回数'])*100:.1f}%" if r['日勤(A/P)回数']>0 else "0.0%", axis=1)
                    df_res['夜勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                    df_res['公休回数'] = (df_res[cols] == '公').sum(axis=1)
                    df_res['日曜D回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'D') if staff_sun_d[0 for e in range(num_staff)]
                    df_res['日曜E回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'E') if staff_sun_e[e] == "〇" else 0 for e in range(num_staff)]

                    sum_A = {"スタッフ名": "【日勤(A/P) 合計人数】"}
                    sum_Az = {"スタッフ名": "【残業(A残) 合計人数】"}
                    sum_D = {"スタッフ名": "【夜勤(D) 合計人数】"}
                    sum_O = {"スタッフ名": "【公e] == "〇" else 0 for e in range(num_staff)]
                    df_res['日曜E回数'] = [sum(1 for d in range(num_days) if '日' in weekdays[d] and df_res.loc[e, cols[d]] == 'E') if staff_sun_e[e] == "〇" else 0 for e in range(num_staff)]

                    sum_A = {"スタッフ名": "【日勤(A/P) 合計人数】"}
                    sum_Az = {"スタッフ名": "【残業(A残) 合計人数】"}
                    sum_D = {"スタッフ名": "【夜勤(D) 合計人数】"}
                    sum_O休 合計人数】"}
                    
                    for c in ['日勤(A/P)回数', '残業(A残)回数', '残業割合(%)', '夜勤(D)回数', '公休回数', '日曜D回数', '日曜E回数']:
                        sum_A[c] = ""; sum_Az[c] = ""; sum_D[c] = ""; sum_O[c] = ""

                    for d, c in enumerate(cols):
                        sum_A[c] = sum(1 for e in range(num_staff) if str(df_res.loc[e = {"スタッフ名": "【公休 合計人数】"}
                    
                    for c in ['日勤(A/P)回数', '残業(A残)回数', '残業割合(%)', '夜勤(D)回数', '公休回数', '日曜D回数', '日曜E回数']:
                        sum_A[c] = ""; sum_Az[c] = ""; sum_D[c] = ""; sum_O[c] = ""

                    for d, c in enumerate(cols):
                        sum_A[c] = sum(1 for e in range(num_staff) if str(df_res.loc[e, c]) in ['A', 'A残'] or 'P' in, c]) in ['A', 'A残'] or 'P' in str(df_res.loc[e, c]) and "新人" not in str(staff_roles[e]))
                        sum_Az[c] = (df_res[c] == 'A残').sum()
                        sum_D[c] = (df_res[c] == 'D').sum()
                        sum_O[c] = (df_res[c] == '公').sum()

                    df_fin = pd.concat([df_res, pd.DataFrame([sum_A, sum_Az, sum_D, sum_O])], ignore_index=True)

                    def highlight_warnings(df):
                        styles = pd.DataFrame('', index str(df_res.loc[e, c]) and "新人" not in str(staff_roles[e]))
                        sum_Az[c] = (df_res[c] == 'A残').sum()
                        sum_D[c] = (df_res[c] == 'D').sum()
                        sum_O[c] = (df_res[c] == '公').sum()

                    df_fin = pd.concat([df_res, pd.DataFrame([sum_A, sum_Az, sum_D, sum_O])], ignore_index=True)

                    def highlight_warnings(df):
                        styles = pd.DataFrame('', index=df.index, columns=df.columns)
                        for d, col_name in enumerate(cols):
                            actual_a = df.loc[len(staff_names),=df.index, columns=df.columns)
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
                                    v = str(df. col_name]
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

                                if is_day_work(d) and is_day_work(d+1) and is_dayloc[e, cols[day_idx]])
                                    return v == 'A' or v == 'A残' or 'P' in v or 'Ｐ' in v

                                if is_day_work(d) and is_day_work(d+1) and is_day_work(d+2) and is_day_work(d+3):
                                    styles.loc[e, cols[d]] = 'background-color: #FFFF99;'
                                    styles.loc[e, cols[d+1]] = 'background-color: #FFFF99;'
                                    styles.loc[e, cols[d+2]] = 'background-_work(d+2) and is_day_work(d+3):
                                    styles.loc[e, cols[d]] = 'background-color: #FFFF99;'
                                    styles.loc[e, cols[d+1]] = 'background-color: #FFFF99;'
                                    styles.loc[e, cols[d+2]] = 'background-color: #FFFF99;'
                                    styles.loc[e, cols[d+3]] = 'background-color: #FFFF99;'

                                if d + 3 < num_days:
                                    if is_day_work(d) and is_day_work(color: #FFFF99;'
                                    styles.loc[e, cols[d+3]] = 'background-color: #FFFF99;'

                                if d + 3 < num_days:
                                    if is_day_work(d) and is_day_work(d+1) and is_day_work(d+2) and str(df.loc[e, cols[d+3]]) == 'D':
                                        styles.loc[e, cols[d]] = 'background-color: #FFD580;'
                                        styles.loc[e, cols[d+1]] = 'background-colord+1) and is_day_work(d+2) and str(df.loc[e, cols[d+3]]) == 'D':
                                        styles.loc[e, cols[d]] = 'background-color: #FFD580;'
                                        styles.loc[e, cols[d+1]] = 'background-color: #FFD580;'
                                        styles.loc[e, cols[d+2]] = 'background-color: #FFD580;'
                                        styles.loc[e, cols[d+3]] = 'background-color: #FFD580;': #FFD580;'
                                        styles.loc[e, cols[d+2]] = 'background-color: #FFD580;'
                                        styles.loc[e, cols[d+3]] = 'background-color: #FFD580;'
                        return styles

                    st.dataframe(df_fin.style.apply(highlight_warnings, axis=None))
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_fin.to_excel(writer, index=False, sheet_name='完成シフト')

                        return styles

                    st.dataframe(df_fin.style.apply(highlight_warnings, axis=None))
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_fin.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label=f"📥 【パターン {i+1}】 をエクセルでダウンロード（色なし）",
                        data=processed_data,
                        file_name=f"完成版_対                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label=f"📥 【パターン {i+1}】 をエクセルでダウンロード（色なし）",
                        data=processed_data,
                        file_name=f"完成版_対話型シフト_{i+1}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_btn_{i}"
                    )
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白の行があります。({e})")
