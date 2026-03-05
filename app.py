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
st.write("「定時確保」や「残業の逆比例公平化」を搭載した、実務完全対応のシフト作成システムです。")

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
        
        staff_sun_d = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Dカウント", "〇"))]
        staff_sun_e = ["×" if ok == "×" else v for ok, v in zip(staff_night_ok, get_staff_col("日曜Eカウント", "〇"))]

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
            # 位置番号(d+1)ではなく列名(date_columns[d])で引くことでズレを防止
            row = df_req[df_req.iloc[:, 0] == label]
            res = []
            for d in range(num_days):
                col_name = date_columns[d]
                try:
                    if not row.empty and col_name in df_req.columns:
                        val = row.iloc[0][col_name]
                        if pd.notna(val):
                            res.append(int(val) if is_int else str(val).strip())
                            continue
                except Exception:
                    pass
                res.append(default_val)
            return res

        day_req_list = get_req_col("日勤人数", 3)
        night_req_list = get_req_col("夜勤人数", 2)
        overtime_req_list = get_req_col("残業人数", 0)
        absolute_req_list = get_req_col("絶対確保", "", is_int=False)
        # 曜日はエクセルの値を使わず、選択された年月と日付から自動計算（ズレ防止）
        _weekday_map = ["月", "火", "水", "木", "金", "土", "日"]
        weekdays = []
        for _d_val in date_columns:
            try:
                _dt = datetime.date(target_year, target_month, int(_d_val))
                weekdays.append(_weekday_map[_dt.weekday()])
            except (ValueError, TypeError):
                weekdays.append("")

        # =====================================
        # 🔍 不可能理由の診断関数（新規追加）
        # =====================================
        def diagnose_infeasibility():
            """
            シフトが組めない原因を詳細に分析して表示する。
            ① スタッフ別 稼働余力テーブル
            ② 日別 人員過不足テーブル
            ③ 前月末の夜勤チェーン影響
            ④ 希望休が集中している日
            ⑤ リーダー不足・夜勤総量・残業要員などの全体チェック
            """

            # ────────────────────────────────────────────
            # 事前集計：スタッフ別の希望休（固定公休）一覧
            # ────────────────────────────────────────────
            fixed_off = [0] * num_staff          # スタッフeの固定公休日数
            fixed_off_days_list = [[] for _ in range(num_staff)]  # スタッフeの固定公休日リスト
            fixed_off_per_day = [0] * num_days   # d日目の固定公休人数

            # 前月末シフト（前月最終日）
            prev_last_shift = {}   # {e: "D"/"E"/"A"/...}
            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    # 前月最終日（列5）
                    if tr.shape[1] > 5:
                        prev_last_shift[e] = str(tr.iloc[0, 5]).strip()
                    # 希望休（G列=index6以降のみ対象。前月列との列名衝突を防ぐ）
                    history_today_cols = list(tr.columns[6:])
                    for d in range(num_days):
                        col_name = date_columns[d]
                        if col_name in history_today_cols and str(tr.iloc[0][col_name]).strip() == "公":
                            fixed_off[e] += 1
                            fixed_off_days_list[e].append(d)
                            fixed_off_per_day[d] += 1

            # ────────────────────────────────────────────
            # 前月末夜勤チェーンで今月頭が強制される日
            # ────────────────────────────────────────────
            # forced_block[d] = [(staff_name, reason), ...]
            forced_block = [[] for _ in range(num_days)]
            for e, staff_name in enumerate(staff_names):
                last = prev_last_shift.get(e, "")
                if last == "D":
                    # 1日目→E確定、2日目→公確定
                    if num_days > 0: forced_block[0].append((staff_name, "前月末がDのためE確定"))
                    if num_days > 1: forced_block[1].append((staff_name, "前月末がDのためE翌日→公確定"))
                elif last == "E":
                    # 1日目→公確定
                    if num_days > 0: forced_block[0].append((staff_name, "前月末がEのため公確定"))

            # ────────────────────────────────────────────
            # ① スタッフ別 稼働余力テーブル
            # ────────────────────────────────────────────
            staff_table_rows = []
            staff_warnings = []
            for e in range(num_staff):
                total_days = num_days
                f_off = fixed_off[e]
                required_off = staff_off_days[e]
                remaining_off_to_assign = max(0, required_off - f_off)
                free_days = total_days - f_off  # 希望休を除いた日数
                work_days = free_days - remaining_off_to_assign  # 実働可能日数（公休引き後）

                # 前月チェーンで強制ブロックされる日数
                chain_block = 0
                if prev_last_shift.get(e, "") == "D": chain_block = 2
                elif prev_last_shift.get(e, "") == "E": chain_block = 1
                actual_work = work_days - chain_block

                night_limit = staff_night_limits[e] if staff_night_ok[e] != "×" else 0
                night_label = f"{night_limit}回上限" if staff_night_ok[e] != "×" else "夜勤不可"

                # 判定
                if required_off > total_days:
                    status = "🔴 公休数が日数超過"
                    staff_warnings.append(f"**{staff_names[e]}**：公休数({required_off}日) が月の日数({total_days}日)を超えています！")
                elif actual_work < 0:
                    status = "🔴 実働日数が足りない"
                    staff_warnings.append(f"**{staff_names[e]}**：希望休({f_off}日)+残り公休({remaining_off_to_assign}日)+前月チェーン({chain_block}日) = {f_off+remaining_off_to_assign+chain_block}日が拘束され、実働できる日が足りません。")
                elif actual_work <= 3:
                    status = "🟡 実働日数が少ない"
                    staff_warnings.append(f"**{staff_names[e]}**：実働可能日数が **{actual_work}日** しかありません（希望休{f_off}日＋残り公休{remaining_off_to_assign}日＋前月チェーン{chain_block}日）。")
                else:
                    status = "🟢 問題なし"

                staff_table_rows.append({
                    "スタッフ名": staff_names[e],
                    "役割": staff_roles[e],
                    "月日数": total_days,
                    "希望休(固定)": f_off,
                    "残り公休割当": remaining_off_to_assign,
                    "前月チェーン拘束": chain_block,
                    "実働可能日数": max(0, actual_work),
                    "夜勤": night_label,
                    "判定": status,
                })

            df_staff_diag = pd.DataFrame(staff_table_rows)

            # ────────────────────────────────────────────
            # ② 日別 人員過不足テーブル
            # ────────────────────────────────────────────
            day_table_rows = []
            day_warnings = []
            for d in range(num_days):
                day_label = f"{date_columns[d]}({weekdays[d]})"

                # その日にブロックされているスタッフを集計
                blocked_names = []
                # 希望休
                for e in range(num_staff):
                    if d in fixed_off_days_list[e]:
                        blocked_names.append(f"{staff_names[e]}(希望休)")
                # 前月チェーン強制
                for (sname, reason) in forced_block[d]:
                    if not any(sname in b for b in blocked_names):
                        blocked_names.append(f"{sname}({reason})")

                blocked_count = len(blocked_names)
                # 夜勤(D)はその日にDシフトが入るので日勤から除外
                night_consuming = night_req_list[d]  # D勤務人数
                # Eシフト：前日がDの人はその日E確定（今月1日以降）
                e_consuming = night_req_list[d - 1] if d > 0 else 0

                available_for_day = num_staff - blocked_count - night_consuming - e_consuming
                required_day = day_req_list[d]
                required_ot  = overtime_req_list[d]
                total_required = required_day + required_ot

                gap = available_for_day - total_required

                if gap < 0:
                    status = f"🔴 {abs(gap)}人不足"
                    blocked_str = "、".join(blocked_names[:5]) + ("…他" if len(blocked_names) > 5 else "")
                    day_warnings.append(
                        f"**{day_label}**：必要{total_required}人 に対し最大{max(0,available_for_day)}人しか確保できません。"
                        + (f"（ブロック：{blocked_str}）" if blocked_names else "")
                    )
                elif gap <= 1:
                    status = "🟡 ギリギリ"
                else:
                    status = "🟢 OK"

                day_table_rows.append({
                    "日付": day_label,
                    "必要日勤": required_day,
                    "必要残業": required_ot,
                    "必要夜勤(D)": night_consuming,
                    "希望休人数": fixed_off_per_day[d],
                    "前月チェーン拘束": len(forced_block[d]),
                    "利用可能人数(推定)": max(0, available_for_day),
                    "判定": status,
                })

            df_day_diag = pd.DataFrame(day_table_rows)

            # ────────────────────────────────────────────
            # ③ 希望休が集中している日を検出
            # ────────────────────────────────────────────
            congestion_days = []
            for d in range(num_days):
                if fixed_off_per_day[d] >= max(2, num_staff // 3):
                    names_off = [staff_names[e] for e in range(num_staff) if d in fixed_off_days_list[e]]
                    congestion_days.append(f"**{date_columns[d]}日({weekdays[d]})**：{len(names_off)}名が希望休（{', '.join(names_off)}）")

            # ────────────────────────────────────────────
            # ④ 夜勤総量・リーダー・残業の全体チェック
            # ────────────────────────────────────────────
            global_issues = []
            night_capable = [e for e in range(num_staff) if staff_night_ok[e] != "×"]
            total_night_capacity = sum(staff_night_limits[e] for e in night_capable)
            total_night_required = sum(night_req_list)
            if total_night_capacity < total_night_required:
                global_issues.append(f"🌙 **夜勤総量不足**：月合計夜勤 **{total_night_required}回** 必要なのに、スタッフ上限合計は **{total_night_capacity}回** です。（不足：{total_night_required-total_night_capacity}回）")
            if len(night_capable) < 2:
                global_issues.append(f"🌙 **夜勤可能スタッフが{len(night_capable)}名**：D→E→公のルールを回すには最低2名必要です。")

            leader_count = sum(1 for r in staff_roles if "主任" in str(r) or "リーダー" in str(r) or "サブ" in str(r))
            if leader_count == 0:
                global_issues.append("👤 **リーダー/サブリーダーが0名**：毎日の日勤にリーダー系が必須のためシフトが組めません。")
            elif leader_count == 1:
                global_issues.append(f"👤 **リーダー系が1名のみ**：公休・夜勤でリーダーが不在になる日が発生します。")

            ot_capable = sum(1 for e in range(num_staff) if staff_overtime_ok[e] != "×")
            total_ot_required = sum(overtime_req_list)
            if total_ot_required > 0 and ot_capable == 0:
                global_issues.append(f"⏰ **残業要員が0名**：残業が必要な日がありますが「残業可否〇」のスタッフがいません。")

            return df_staff_diag, staff_warnings, df_day_diag, day_warnings, congestion_days, global_issues

        st.success("✅ データの読み込み完了！まずは妥協なしの「理想のシフト」を作れるかテストします。")

        def solve_shift(random_seed, allow_minus_1=False, allow_4_days=False, allow_night_3=False, allow_sub_only=False, allow_ot_consec=False, allow_night_consec_3=False, allow_sun_minus_1=False):
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
                    elif last_day == "E":
                        model.Add(shifts[(e, 0, '公')] == 1)

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
                    # 日曜は設定人数固定。allow_sun_minus_1かつリーダー在勤の場合のみ-1を許可
                    model.Add(act_day <= req)
                    if not allow_sun_minus_1:
                        model.Add(act_day == req)
                    else:
                        # リーダー以上が1名以上いるか確認
                        leader_present = sum((1 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff))
                        has_leader = model.NewBoolVar('')
                        model.Add(leader_present >= 1).OnlyEnforceIf(has_leader)
                        model.Add(leader_present == 0).OnlyEnforceIf(has_leader.Not())
                        # リーダーがいる場合のみ-1を許可
                        model.Add(act_day >= req - 1).OnlyEnforceIf(has_leader)
                        model.Add(act_day >= req).OnlyEnforceIf(has_leader.Not())
                        m_var = model.NewBoolVar('')
                        model.Add(act_day == req - 1).OnlyEnforceIf(m_var)
                        model.Add(act_day != req - 1).OnlyEnforceIf(m_var.Not())
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
                    # G列(index6)以降のみ対象。前月列との列名衝突を防ぐ
                    history_today_cols = list(tr.columns[6:])
                    for d in range(num_days):
                        col_name = date_columns[d]
                        if col_name in history_today_cols and str(tr.iloc[0][col_name]).strip() == "公":
                            model.Add(shifts[(e, d, '公')] == 1)

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

        def solve_shift_fast(flags):
            """妥協案の必要性チェック専用。時間を短縮して実現可否だけ判定する。"""
            m1, f4, n3, sub, ot, nc3, sun1 = flags
            # 45秒では遅いので、内部で短い制限を使う専用ソルバーを呼ぶ
            # → solve_shiftのmax_time_in_secondsを短縮した版を直接組む
            from ortools.sat.python import cp_model as _cp
            _model = _cp.CpModel()
            _types = ['A', 'A残', 'D', 'E', '公']
            _shifts = {(e, d, s): _model.NewBoolVar('') for e in range(num_staff) for d in range(num_days) for s in _types}
            # 基本制約のみ（ペナルティなし）でFEASIBLEチェック
            for e in range(num_staff):
                for d in range(num_days): _model.AddExactlyOne(_shifts[(e, d, s)] for s in _types)
                if staff_night_ok[e] == "×":
                    for d in range(num_days): _model.Add(_shifts[(e,d,'D')]==0); _model.Add(_shifts[(e,d,'E')]==0)
                if staff_overtime_ok[e] == "×":
                    for d in range(num_days): _model.Add(_shifts[(e,d,'A残')]==0)
            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    last_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                    if last_day == "D":
                        _model.Add(_shifts[(e,0,'E')]==1)
                        if num_days > 1: _model.Add(_shifts[(e,1,'公')]==1)
                    elif last_day == "E": _model.Add(_shifts[(e,0,'公')]==1)
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        l_day = str(tr.iloc[0, 5]).strip() if tr.shape[1] > 5 else ""
                        if l_day != "D": _model.Add(_shifts[(e,0,'E')]==0)
                    for d in range(num_days):
                        if d > 0: _model.Add(_shifts[(e,d,'E')] == _shifts[(e,d-1,'D')])
                        if d+1 < num_days: _model.AddImplication(_shifts[(e,d,'E')], _shifts[(e,d+1,'公')])
            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    history_today_cols = list(tr.columns[6:])
                    for d in range(num_days):
                        col_name = date_columns[d]
                        if col_name in history_today_cols and str(tr.iloc[0][col_name]).strip() == "公":
                            _model.Add(_shifts[(e,d,'公')]==1)
            for e in range(num_staff):
                _model.Add(sum(_shifts[(e,d,'公')] for d in range(num_days)) == int(staff_off_days[e]))
                if staff_night_ok[e] != "×": _model.Add(sum(_shifts[(e,d,'D')] for d in range(num_days)) <= int(staff_night_limits[e]))
            for e in range(num_staff):
                if staff_night_ok[e] != "×":
                    past_D = [0]*5
                    tr = df_history[df_history.iloc[:, 0] == staff_names[e]]
                    if not tr.empty:
                        for i in range(5):
                            if (i+1) < tr.shape[1] and str(tr.iloc[0,i+1]).strip() == "D": past_D[i]=1
                    all_D = past_D + [_shifts[(e,d,'D')] for d in range(num_days)]
                    for i in range(len(all_D)-6):
                        window = all_D[i:i+7]
                        if not nc3 and any(isinstance(x, _cp.IntVar) for x in window): _model.Add(sum(window)<=2)
            for d in range(num_days):
                _model.Add(sum(_shifts[(e,d,'D')] for e in range(num_staff)) == night_req_list[d])
                _model.Add(sum(_shifts[(e,d,'A残')] for e in range(num_staff)) == overtime_req_list[d])
                act_day = sum((_shifts[(e,d,'A')]+_shifts[(e,d,'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e]))
                req = day_req_list[d]; is_sun = ('日' in weekdays[d]); is_abs = (absolute_req_list[d]=="〇")
                if is_abs: _model.Add(act_day >= req)
                elif is_sun:
                    _model.Add(act_day <= req)
                    if not sun1: _model.Add(act_day == req)
                    else: _model.Add(act_day >= req-1)
                else:
                    if not m1: _model.Add(act_day >= req)
                    else: _model.Add(act_day >= req-1)
                l_score = sum((2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0)*(_shifts[(e,d,'A')]+_shifts[(e,d,'A残')]) for e in range(num_staff))
                if not sub: _model.Add(l_score >= 2)
                else: _model.Add(l_score >= 1)
            for e in range(num_staff):
                target_lvl = staff_comp_lvl[e]
                for d in range(num_days-3):
                    def _work(day, _e=e): return _shifts[(_e,day,'A')]+_shifts[(_e,day,'A残')]
                    if not f4 or target_lvl == 0: _model.Add(_work(d)+_work(d+1)+_work(d+2)+_work(d+3)<=3)
                    if not n3 or target_lvl == 0: _model.Add(_work(d)+_work(d+1)+_work(d+2)<=2).OnlyEnforceIf(_shifts[(e,d+3,'D')])
            for e in range(num_staff):
                for d in range(num_days-1):
                    if not ot: _model.Add(_shifts[(e,d,'A残')]+_shifts[(e,d+1,'A残')]<=1)
            for e in range(num_staff):
                if staff_overtime_ok[e] != "×":
                    total_day_work = sum(_shifts[(e,d,'A')]+_shifts[(e,d,'A残')] for d in range(num_days))
                    b_has = _model.NewBoolVar('')
                    _model.Add(total_day_work>0).OnlyEnforceIf(b_has); _model.Add(total_day_work==0).OnlyEnforceIf(b_has.Not())
                    _model.Add(sum(_shifts[(e,d,'A')] for d in range(num_days)) >= int(staff_min_normal_a[e])).OnlyEnforceIf(b_has)
            _solver = _cp.CpSolver()
            _solver.parameters.max_time_in_seconds = 12.0
            _solver.parameters.random_seed = 1
            _status = _solver.Solve(_model)
            return _status in (_cp.OPTIMAL, _cp.FEASIBLE)

        # 妥協案の名称・引数位置の定義（順番固定）
        # flags = (allow_minus_1, allow_4_days, allow_night_3, allow_sub_only, allow_ot_consec, allow_night_consec_3, allow_sun_minus_1)
        COMPROMISE_LABELS = [
            "日勤人数の「マイナス1」（平日・祝）",
            "最大4連勤のお願い",
            "夜勤前3日連続日勤のお願い",
            "役割配置をサブ1名まで下げる",
            "残業(A残)の2日連続",
            "夜勤セット3連続（月またぎ含む）",
            "日曜の出勤人数を設定より-1（リーダー在勤条件）",
        ]
        ALL_ON = (True, True, True, True, True, True, True)

        if 'min_compromise_result' not in st.session_state:
            st.session_state.min_compromise_result = None
        if 'step3_failed' not in st.session_state:
            st.session_state.step3_failed = False

        if not st.session_state.needs_compromise:
            if st.button("▶️ 【STEP 1】まずは妥協なしで理想のシフトを計算する（5パターン）"):
                with st.spinner('AIが「妥協なし」の完璧なシフトを5パターン模索中...'):
                    results = []
                    for seed in [7, 42, 137, 512, 9999]:
                        solver, shifts = solve_shift(seed, False, False, False, False, False, False, False)
                        if solver: results.append((solver, shifts))

                    if results:
                        st.success(f"🎉 妥協なしで完璧なシフトが {len(results)} パターン組めました！")
                    else:
                        st.session_state.needs_compromise = True
                        st.session_state.min_compromise_result = None
                        st.session_state.step3_failed = False
                        st.rerun()

        else:
            st.error("⚠️ 妥協なしではシフトを組めませんでした。妥協案を選んで再計算してください。")

            # ── 診断レポート ──
            df_staff_diag, staff_warnings, df_day_diag, day_warnings, congestion_days, global_issues = diagnose_infeasibility()
            with st.expander("🔍 **【原因診断レポート】** ← クリックして確認", expanded=False):
                if global_issues:
                    st.markdown("### ⚠️ 全体チェックで検出された問題")
                    for g in global_issues: st.error(g)
                    st.markdown("---")
                st.markdown("### 👥 ① スタッフ別 稼働余力チェック")
                def color_staff_status(val):
                    if "🔴" in str(val): return "background-color:#FFD0D0;font-weight:bold;"
                    if "🟡" in str(val): return "background-color:#FFF5CC;"
                    if "🟢" in str(val): return "background-color:#D8F5D8;"
                    return ""
                st.dataframe(df_staff_diag.style.applymap(color_staff_status, subset=["判定"]), use_container_width=True)
                if staff_warnings:
                    for w in staff_warnings: st.warning(w)
                st.markdown("---")
                st.markdown("### 📅 ② 日別 人員過不足チェック")
                def color_day_status(val):
                    if "🔴" in str(val): return "background-color:#FFD0D0;font-weight:bold;"
                    if "🟡" in str(val): return "background-color:#FFF5CC;"
                    if "🟢" in str(val): return "background-color:#D8F5D8;"
                    return ""
                st.dataframe(df_day_diag.style.applymap(color_day_status, subset=["判定"]), use_container_width=True)
                if day_warnings:
                    for w in day_warnings: st.error(w)
                if congestion_days:
                    st.markdown("### 🗓️ ③ 希望休が集中している日")
                    for c in congestion_days: st.warning(c)
                st.caption("※ この診断は推定値です。")

            st.markdown("---")

            # ── STEP2：自動分析ボタン ──
            if st.button("🔎 【STEP 2】どの妥協案が本当に必要か自動で調べる（約1分）"):
                progress = st.progress(0, text="全妥協案ONで確認中...")
                can_all = solve_shift_fast(ALL_ON)
                if not can_all:
                    st.error("😭 全妥協案をONにしても組めません。希望休や人数設定を見直してください。")
                    st.session_state.min_compromise_result = None
                else:
                    necessary = [True] * 7
                    for i in range(7):
                        progress.progress((i+1)/8, text=f"「{COMPROMISE_LABELS[i]}」が不要か確認中... ({i+1}/7)")
                        flags_test = list(ALL_ON)
                        flags_test[i] = False
                        if solve_shift_fast(tuple(flags_test)):
                            necessary[i] = False
                    progress.progress(1.0, text="分析完了！")
                    st.session_state.min_compromise_result = necessary

            if st.session_state.min_compromise_result is not None:
                necessary = st.session_state.min_compromise_result
                needed   = [COMPROMISE_LABELS[i] for i in range(7) if necessary[i]]
                unneeded = [COMPROMISE_LABELS[i] for i in range(7) if not necessary[i]]
                col_n, col_u = st.columns(2)
                with col_n:
                    st.error("**🔴 必要な妥協案**")
                    for n in (needed or ["なし"]): st.markdown(f"　{n}")
                with col_u:
                    st.success("**🟢 不要な妥協案**")
                    for u in (unneeded or ["なし"]): st.markdown(f"　{u}")
                st.markdown("---")

            # ── 妥協案チェックボックス（日曜-1はここには出さない） ──
            _r = st.session_state.min_compromise_result or [False]*7
            st.markdown("### 📝 【STEP 2】妥協案を選んでください")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**■ 人数・役割**")
                allow_minus_1  = st.checkbox("絶対確保日・日曜以外の日勤人数を-1にする", value=bool(_r[0]))
                allow_sub_only = st.checkbox("役割配置を「サブ1名＋他」まで下げる", value=bool(_r[3]))
            with col2:
                st.markdown("**■ 連勤・夜勤**")
                allow_4_days   = st.checkbox("対象者への「最大4連勤」を許可する", value=bool(_r[1]))
                allow_night_3  = st.checkbox("対象者への「夜勤前3日連続日勤」を許可する", value=bool(_r[2]))
            st.markdown("**■ その他**")
            col3, col4 = st.columns(2)
            with col3:
                allow_night_consec_3 = st.checkbox("「夜勤セット3連続（月またぎ）」を許可する", value=bool(_r[5]))
            with col4:
                allow_ot_consec = st.checkbox("「残業A残の2日連続」を許可する", value=bool(_r[4]))

            if st.button("🔄 【STEP 3】選んだ妥協案で5パターン作成"):
                with st.spinner('計算中...'):
                    results = []
                    for seed in [7, 42, 137, 512, 9999]:
                        solver, shifts = solve_shift(seed, allow_minus_1, allow_4_days, allow_night_3, allow_sub_only, allow_ot_consec, allow_night_consec_3, False)
                        if solver: results.append((solver, shifts))
                    if results:
                        st.success(f"✨ {len(results)}パターン完成！")
                        st.session_state.needs_compromise = False
                        st.session_state.step3_failed = False
                    else:
                        st.session_state.step3_failed = True
                        st.rerun()

            # ── STEP3で失敗した場合のみ、日曜-1を追加オプションとして提示 ──
            if st.session_state.step3_failed:
                st.error("❌ 選んだ妥協案でもシフトが組めませんでした。")
                st.markdown("---")
                st.markdown("### 🌟 【STEP 4】日曜日の出勤人数も-1にしますか？")
                st.warning("日曜は設定人数を固定しているため、さらに厳しい場合はここで許可できます。")
                allow_sun_minus_1 = st.checkbox("日曜の出勤人数を設定より-1にする（リーダー在勤時のみ）", value=bool(_r[6]))

                if st.button("🔄 【STEP 4】日曜-1も加えて5パターン作成"):
                    with st.spinner('計算中...'):
                        results = []
                        for seed in [7, 42, 137, 512, 9999]:
                            solver, shifts = solve_shift(seed, allow_minus_1, allow_4_days, allow_night_3, allow_sub_only, allow_ot_consec, allow_night_consec_3, allow_sun_minus_1)
                            if solver: results.append((solver, shifts))
                        if results:
                            st.success(f"✨ {len(results)}パターン完成！")
                            st.session_state.needs_compromise = False
                            st.session_state.step3_failed = False
                        else:
                            st.error("😭 まだ組めません。希望休や人数設定を見直してください。")

        if 'results' in locals() and results:
            cols = []
            for d_val, w_val in zip(date_columns, weekdays):
                try:
                    dt = datetime.date(target_year, target_month, int(d_val))
                    if jpholiday.is_holiday(dt): cols.append(f"{d_val}({w_val}・祝)")
                    else: cols.append(f"{d_val}({w_val})")
                except ValueError: cols.append(f"{d_val}({w_val})")

            # 希望休セル(e, d)の集合を事前に構築（黄色塗り用）
            hope_off_set = set()
            for e, staff_name in enumerate(staff_names):
                tr = df_history[df_history.iloc[:, 0] == staff_name]
                if not tr.empty:
                    history_today_cols = list(tr.columns[6:])
                    for d in range(num_days):
                        col_name = date_columns[d]
                        if col_name in history_today_cols and str(tr.iloc[0][col_name]).strip() == "公":
                            hope_off_set.add((e, d))

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
                    df_res['夜勤(D)回数'] = (df_res[cols] == 'D').sum(axis=1)
                    df_res['公休回数'] = (df_res[cols] == '公').sum(axis=1)

                    sum_A  = {"スタッフ名": "【日勤(A/P) 合計人数】"}
                    sum_Az = {"スタッフ名": "【残業(A残) 合計人数】"}
                    sum_D  = {"スタッフ名": "【夜勤(D) 人数】"}
                    sum_E  = {"スタッフ名": "【夜勤明け(E) 人数】"}

                    for c in ['日勤(A/P)回数', '残業(A残)回数', '夜勤(D)回数', '公休回数']:
                        sum_A[c] = ""; sum_Az[c] = ""; sum_D[c] = ""; sum_E[c] = ""

                    for d, c in enumerate(cols):
                        sum_A[c]  = sum(1 for e in range(num_staff) if str(df_res.loc[e, c]) in ['A', 'A残'] or 'P' in str(df_res.loc[e, c]) and "新人" not in str(staff_roles[e]))
                        sum_Az[c] = (df_res[c] == 'A残').sum()
                        sum_D[c]  = (df_res[c] == 'D').sum()
                        sum_E[c]  = (df_res[c] == 'E').sum()

                    df_fin = pd.concat([df_res, pd.DataFrame([sum_A, sum_Az, sum_D, sum_E])], ignore_index=True)

                    def highlight_warnings(df):
                        styles = pd.DataFrame('', index=df.index, columns=df.columns)
                        
                        for d, col_name in enumerate(cols):
                            if "土" in col_name: styles.iloc[:, d+1] = 'background-color: #E6F2FF;'
                            elif "日" in col_name or "祝" in col_name: styles.iloc[:, d+1] = 'background-color: #FFE6E6;'

                        # 希望休セルを黄色に（他の色より先に塗って上書きされる前提）
                        for e in range(num_staff):
                            for d in range(num_days):
                                if (e, d) in hope_off_set:
                                    styles.loc[e, cols[d]] = 'background-color: #FFFF00; font-weight: bold;'
                        
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
                                    styles.loc[e, cols[d]] = 'background-color: #FFFF99; font-weight: bold; color: black;'
                                    styles.loc[e, cols[d+1]] = 'background-color: #FFFF99; font-weight: bold; color: black;'
                                    styles.loc[e, cols[d+2]] = 'background-color: #FFFF99; font-weight: bold; color: black;'
                                    styles.loc[e, cols[d+3]] = 'background-color: #FFFF99; font-weight: bold; color: black;'

                                if d + 3 < num_days:
                                    if is_day_work(d) and is_day_work(d+1) and is_day_work(d+2) and str(df.loc[e, cols[d+3]]) == 'D':
                                        styles.loc[e, cols[d]] = 'background-color: #FFD580; font-weight: bold; color: black;'
                                        styles.loc[e, cols[d+1]] = 'background-color: #FFD580; font-weight: bold; color: black;'
                                        styles.loc[e, cols[d+2]] = 'background-color: #FFD580; font-weight: bold; color: black;'
                                        styles.loc[e, cols[d+3]] = 'background-color: #FFD580; font-weight: bold; color: black;'
                                        
                                if d + 8 < num_days:
                                    if str(df.loc[e, cols[d]]) == 'D' and str(df.loc[e, cols[d+3]]) == 'D' and str(df.loc[e, cols[d+6]]) == 'D':
                                        for k in range(9):
                                            styles.loc[e, cols[d+k]] = 'background-color: #E6E6FA; font-weight: bold; color: black;'
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

                        fill_sat       = PatternFill(start_color="E6F2FF", end_color="E6F2FF", fill_type="solid")
                        fill_sun       = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
                        fill_short     = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                        fill_over      = PatternFill(start_color="CCFFFF", end_color="CCFFFF", fill_type="solid")
                        fill_4days     = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                        fill_n3        = PatternFill(start_color="FFD580", end_color="FFD580", fill_type="solid")
                        fill_n3_consec = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")
                        fill_hope_off  = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

                        # ── 印刷用：各スタッフ行の下に空白行を挿入（下から挿入でズレ防止）
                        for e in range(num_staff - 1, -1, -1):
                            insert_row = e + 3  # ヘッダー1行 + スタッフ行(e+2) の次
                            worksheet.insert_rows(insert_row)

                        # 挿入後の行番号マップ: スタッフe → Excelの行番号
                        # e番目スタッフは (e+1)*2 行目（ヘッダーが1行、スタッフ+空白で2行ずつ）
                        def staff_row(e): return (e + 1) * 2
                        # 集計行（日勤合計）は staff_names の後
                        sum_row_start = num_staff * 2 + 2

                        # ── 全セルにフォント・ボーダー・アライン ──
                        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                            for cell in row:
                                cell.font = font_meiryo
                                cell.border = border_thin
                                cell.alignment = align_left if cell.column == 1 else align_center

                        # ── 土曜・日祝の列塗り ──
                        for c_idx, col_name in enumerate(cols):
                            if "土" in col_name:
                                for r_idx in range(1, worksheet.max_row + 1):
                                    worksheet.cell(row=r_idx, column=c_idx+2).fill = fill_sat
                            elif "日" in col_name or "祝" in col_name:
                                for r_idx in range(1, worksheet.max_row + 1):
                                    worksheet.cell(row=r_idx, column=c_idx+2).fill = fill_sun

                        # ── 日勤合計行の色（不足=赤、超過=青）──
                        for d, col_name in enumerate(cols):
                            actual_a = df_fin.loc[len(staff_names), col_name]
                            if actual_a != "":
                                if actual_a < day_req_list[d]:   worksheet.cell(row=sum_row_start, column=d+2).fill = fill_short
                                elif actual_a > day_req_list[d]: worksheet.cell(row=sum_row_start, column=d+2).fill = fill_over

                        # ── スタッフ行の色（希望休・4連勤・夜勤前3連・夜勤3連続）──
                        for e in range(num_staff):
                            xl_row = staff_row(e)
                            for d in range(num_days):
                                def is_d_work(day_idx, _e=e):
                                    if day_idx >= num_days: return False
                                    v = str(df_fin.loc[_e, cols[day_idx]])
                                    return v == 'A' or v == 'A残' or 'P' in v or 'Ｐ' in v

                                # 希望休（黄色）
                                if (e, d) in hope_off_set:
                                    worksheet.cell(row=xl_row, column=d+2).fill = fill_hope_off
                                    continue  # 希望休は他の色より優先

                                if is_d_work(d) and is_d_work(d+1) and is_d_work(d+2) and is_d_work(d+3):
                                    for k in range(4): worksheet.cell(row=xl_row, column=d+k+2).fill = fill_4days

                                if d + 3 < num_days:
                                    if is_d_work(d) and is_d_work(d+1) and is_d_work(d+2) and str(df_fin.loc[e, cols[d+3]]) == 'D':
                                        for k in range(4): worksheet.cell(row=xl_row, column=d+k+2).fill = fill_n3

                                if d + 8 < num_days:
                                    if str(df_fin.loc[e, cols[d]]) == 'D' and str(df_fin.loc[e, cols[d+3]]) == 'D' and str(df_fin.loc[e, cols[d+6]]) == 'D':
                                        for k in range(9): worksheet.cell(row=xl_row, column=d+k+2).fill = fill_n3_consec

                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label=f"📥 【パターン {i+1}】 をエクセルでダウンロード（レイアウト完成版）",
                        data=processed_data,
                        file_name=f"完成版_レイアウト適用シフト_{i+1}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_btn_{i}"
                    )
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白の行があります。({e})")
