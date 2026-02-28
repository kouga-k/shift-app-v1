import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import jpholiday
import datetime
from openpyxl.styles import PatternFill
import random

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ13：残業割合の公平化)")
st.write("「日勤回数に対する残業の割合」が全員平等になるよう、3パターンのシフトを提案します！")

# --- 妥協案のセッション管理 ---
if 'allow_day_minus_1' not in st.session_state:
    st.session_state.allow_day_minus_1 = False
if 'allow_4_days_work' not in st.session_state:
    st.session_state.allow_4_days_work = False
if 'allow_sub_only' not in st.session_state:
    st.session_state.allow_sub_only = False
if 'allow_consecutive_overtime' not in st.session_state:
    st.session_state.allow_consecutive_overtime = False

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
        staff_overtime_ok = df_staff["残業可否"].fillna("〇").tolist()
        
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
        overtime_req_list = get_req_row("残業人数", 0)
        night_req_list = get_req_row("夜勤人数", 2)

        st.success(f"✅ データの読み込み完了！")
        
        with st.expander("📩 AI店長への特別許可（※エラーで組めない時だけチェックを入れてください）", expanded=True):
            st.warning("👩‍💼 **AI店長からのご相談:**\n\n『どうしても無理な場合だけ、以下の妥協を許可してください💦』")
            col1, col2 = st.columns(2)
            with col1:
                st.session_state.allow_day_minus_1 = st.checkbox("🙏 日勤人数の「マイナス1」を許可する", value=st.session_state.allow_day_minus_1)
                st.session_state.allow_sub_only = st.checkbox("🙏 リーダー不在時、「サブ1名＋他」を許可する", value=st.session_state.allow_sub_only)
            with col2:
                st.session_state.allow_4_days_work = st.checkbox("🙏 誰かが「最大4連勤」になることを許可する（※黄色で警告）", value=st.session_state.allow_4_days_work)
                st.session_state.allow_consecutive_overtime = st.checkbox("🙏 やむを得ない「残業(A残)の2日連続」を許可する", value=st.session_state.allow_consecutive_overtime)

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

            for e in range(num_staff):
                for d in range(num_days - 6):
                    model.Add(shifts[(e, d, 'D')] + shifts[(e, d+3, 'D')] + shifts[(e, d+6, 'D')] <= 2)

            for d in range(num_days):
                model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == night_req_list[d])
                model.Add(sum(shifts[(e, d, 'A残')] for e in range(num_staff)) == overtime_req_list[d])
                
                actual_day_staff = sum(
                    (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')]) for e in range(num_staff) if "新人" not in str(staff_roles[e])
                )
                
                if absolute_req_list[d] == "〇":
                    model.Add(actual_day_staff >= day_req_list[d])
                else:
                    if st.session_state.allow_day_minus_1:
                        model.Add(actual_day_staff >= day_req_list[d] - 1)
                    else:
                        model.Add(actual_day_staff >= day_req_list[d])

            for d in range(num_days):
                leadership_score = sum(
                    (2 if "主任" in str(staff_roles[e]) or "リーダー" in str(staff_roles[e]) else 1 if "サブ" in str(staff_roles[e]) else 0) * (shifts[(e, d, 'A')] + shifts[(e, d, 'A残')])
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
                    
                    def work(day):
                        return shifts[(e, day, 'A')] + shifts[(e, day, 'A残')]
                        
                    if st.session_state.allow_4_days_work:
                        if d < num_days - 4:
                            model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) + work(d+4) <= 4)
                    else:
                        model.Add(work(d) + work(d+1) + work(d+2) + work(d+3) <= 3)

            if not st.session_state.allow_consecutive_overtime:
                for e in range(num_staff):
                    for d in range(num_days - 1):
                        model.Add(shifts[(e, d, 'A残')] + shifts[(e, d+1, 'A残')] <= 1)

            # 🌟 究極の「残業割合」公平化ロジック
            # 月間の総残業枠と総日勤枠（概算）を計算
            total_ot_req = sum(overtime_req_list)
            total_day_req = sum(day_req_list) # 基本値での概算
            
            # 残業可能なスタッフ全員について、ペナルティ（理想からのズレ）を計算
            penalties = []
            if total_ot_req > 0 and total_day_req > 0:
                for e in range(num_staff):
                    if staff_overtime_ok[e] != "×":
                        # この人の実際の日勤合計（A + A残）
                        actual_days_worked = sum(shifts[(e, d, 'A')] + shifts[(e, d, 'A残')] for d in range(num_days))
                        # この人の実際の残業合計（A残）
                        actual_ot = sum(shifts[(e, d, 'A残')] for d in range(num_days))
                        
                        # 【掛け算のトリック】
                        # 理想の残業数 = (実際の日勤数) × (総残業枠 / 総日勤枠)
                        # つまり：実際の日勤数 × 総残業枠 ＝ 理想の残業数 × 総日勤枠
                        # これを利用して、両辺の差（ズレ）をペナルティとする
                        
                        ideal_ot_scaled = actual_days_worked * total_ot_req
                        actual_ot_scaled = actual_ot * total_day_req
                        
                        # ズレの絶対値を計算するための変数
                        diff = model.NewIntVar(-10000, 10000, f'diff_{e}')
                        abs_diff = model.NewIntVar(0, 10000, f'abs_diff_{e}')
                        
                        model.Add(diff == actual_ot_scaled - ideal_ot_scaled)
                        model.AddAbsEquality(abs_diff, diff)
                        penalties.append(abs_diff)
                
                # ペナルティの合計を最小化しろ！と命令する
                if penalties:
                    model.Minimize(sum(penalties))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 60.0 # 1パターンにつき最大60秒
            solver.parameters.random_seed = random_seed
            status = solver.Solve(model)
            
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                return solver, shifts
            else:
                return None, None


        if st.button("公平なシフトを【3パターン】作成する！（最大3分🔥）"):
            with st.spinner('AI店長が全く違う3つのシフトを同時に考えています...（最大3分かかります）'):
                
                results = []
                for seed in [1, 42, 99]:
                    solver, shifts = solve_shift(seed)
                    if solver:
                        results.append((solver, shifts))

                if not results:
                    st.error("❌ 【AI店長より】申し訳ありません、どうしてもシフトが組めません😭 妥協を許可してから再度お試しください！")
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

                            # 🌟 集計：日勤の分母（AとA残とP）と、分子（A残）をそれぞれ出す
                            result_df['日勤(A・P)回数'] = result_df[new_date_columns].apply(lambda x: x.str.contains('A|P|Ｐ', na=False)).sum(axis=1)
                            result_df['残業(A残)回数'] = (result_df[new_date_columns] == 'A残').sum(axis=1)
                            
                            # 🌟 残業の割合（％）を表示する列を追加
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
                                    if actual_a != "" and actual_a < target_a:
                                        styles.loc[len(staff_names), col_name] = 'background-color: #FFCCCC; color: red; font-weight: bold;'
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
                                return styles

                            st.dataframe(final_df.style.apply(highlight_warnings, axis=None))
                            
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                final_df.to_excel(writer, index=False, sheet_name='完成シフト')
                                worksheet = writer.sheets['完成シフト']
                                fill_red = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                                fill_yellow = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
                                
                                for d, col_name in enumerate(new_date_columns):
                                    actual_a = final_df.loc[len(staff_names), col_name]
                                    if actual_a != "" and actual_a < day_req_list[d]:
                                        worksheet.cell(row=len(staff_names)+2, column=d+4).fill = fill_red
                                for e in range(num_staff):
                                    for d in range(num_days):
                                        def is_work(day_idx):
                                            if day_idx >= num_days: return False
                                            v = str(final_df.loc[e, new_date_columns[day_idx]])
                                            return v == 'A' or v == 'A残' or 'P' in v or 'Ｐ' in v or v == 'D' or v == 'E'
                                        if is_work(d) and is_work(d+1) and is_work(d+2) and is_work(d+3):
                                            worksheet.cell(row=e+2, column=d+4).fill = fill_yellow
                                            worksheet.cell(row=e+2, column=d+5).fill = fill_yellow
                                            worksheet.cell(row=e+2, column=d+6).fill = fill_yellow
                                            worksheet.cell(row=e+2, column=d+7).fill = fill_yellow
                                                
                            processed_data = output.getvalue()
                            
                            st.download_button(
                                label=f"📥 【パターン {i+1}】 をエクセルでダウンロード",
                                data=processed_data,
                                file_name=f"完成版_残業割合公平化_パターン{i+1}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_btn_{i}"
                            )
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: エクセルの形式が間違っているか、空白の行があります。({e})")
