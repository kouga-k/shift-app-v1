import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

st.set_page_config(page_title="自動シフト作成アプリ", layout="wide")
st.title("🌟 AI自動シフト作成アプリ (フェーズ3：リーダー＆人数配置)")
st.write("「夜勤セット」＋「夜勤の必要人数」＋「リーダーorサブ2名の配置」を計算します！")

uploaded_file = st.file_uploader("エクセルファイル (.xlsx) を選択", type=["xlsx"])

if uploaded_file:
    try:
        # エクセルの読み込み
        df_staff = pd.read_excel(uploaded_file, sheet_name="スタッフ設定")
        df_history = pd.read_excel(uploaded_file, sheet_name="希望休・先月履歴")
        df_req = pd.read_excel(uploaded_file, sheet_name="必要人数設定")
        
        # スタッフ情報と役割の取得
        staff_names = df_staff["スタッフ名"].tolist()
        staff_roles = df_staff["役割"].fillna("一般").tolist() # 空白は「一般」にする
        num_staff = len(staff_names)
        
        # 日付列の取得（カレンダーの列）
        date_columns = [col for col in df_history.columns if col != "スタッフ名" and not str(col).startswith("Unnamed")]
        num_days = len(date_columns)
        
        # 夜勤の必要人数を取得（"夜勤人数"という行を探してリストにする）
        night_req_row = df_req[df_req.iloc[:, 0] == "夜勤人数"]
        if not night_req_row.empty:
            # B列目以降の数字を取得（足りない分は2人で埋める）
            night_req_list = night_req_row.iloc[0, 1:].fillna(2).tolist()[:num_days]
        else:
            night_req_list = [2] * num_days # 見つからなければ毎日2人にする
            
        st.success(f"✅ {num_staff}名のスタッフデータを読み込みました。リーダーとサブの配置を計算します！")
        
        if st.button("シフトを自動作成する！（フェーズ3発動🔥）"):
            with st.spinner('AI店長が複雑なパズルを解いています...（少し時間がかかります）'):
                
                model = cp_model.CpModel()
                shift_types = ['A', 'D', 'E', '公']
                
                # ① マス目を作る
                shifts = {}
                for e in range(num_staff):
                    for d in range(num_days):
                        for s in shift_types:
                            shifts[(e, d, s)] = model.NewBoolVar(f'shift_{e}_{d}_{s}')
                            
                # ② ルール1: 毎日必ずどれか1つの勤務
                for e in range(num_staff):
                    for d in range(num_days):
                        model.AddExactlyOne(shifts[(e, d, s)] for s in shift_types)
                        
                # ③ ルール2: 夜勤セット（D -> E -> 公）
                for e in range(num_staff):
                    for d in range(num_days - 2):
                        model.AddImplication(shifts[(e, d, 'D')], shifts[(e, d+1, 'E')])
                        model.AddImplication(shifts[(e, d+1, 'E')], shifts[(e, d+2, '公')])
                for e in range(num_staff):
                    model.Add(shifts[(e, num_days-1, 'D')] == 0)
                    model.Add(shifts[(e, num_days-2, 'D')] == 0)

                # ④ ルール3: 毎日の「夜勤(D)」の必要人数を守る
                for d in range(num_days):
                    target_night = int(night_req_list[d] if d < len(night_req_list) else 2)
                    model.Add(sum(shifts[(e, d, 'D')] for e in range(num_staff)) == target_night)

                # ⑤ ルール4: リーダー1名、またはサブ2名以上を「日勤(A)」に配置する
                for d in range(num_days):
                    leadership_score = 0
                    for e in range(num_staff):
                        role = str(staff_roles[e])
                        if "リーダー" in role:
                            # リーダーが日勤(A)なら2ポイント
                            leadership_score += 2 * shifts[(e, d, 'A')]
                        elif "サブ" in role:
                            # サブが日勤(A)なら1ポイント
                            leadership_score += 1 * shifts[(e, d, 'A')]
                    # その日の日勤の合計ポイントが2以上であること！
                    model.Add(leadership_score >= 2)

                # パズルを解かせる！
                solver = cp_model.CpSolver()
                # 複雑なパズルなので、最大10秒で諦めるようにタイマーをセット
                solver.parameters.max_time_in_seconds = 10.0
                status = solver.Solve(model)
                
                if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                    st.success("✨シフトが完成しました！✨ リーダー/サブの配置も完璧です！")
                    
                    # 結果をまとめる
                    result_data = []
                    for e in range(num_staff):
                        row = {"スタッフ名": staff_names[e], "役割": staff_roles[e]}
                        for d in range(num_days):
                            for s in shift_types:
                                if solver.Value(shifts[(e, d, s)]) == 1:
                                    row[date_columns[d]] = s
                        result_data.append(row)
                        
                    result_df = pd.DataFrame(result_data)
                    st.dataframe(result_df)
                    
                    # エクセル出力
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='完成シフト')
                    processed_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 リーダー配置済みのシフトをダウンロード",
                        data=processed_data,
                        file_name="完成版_フェーズ3.xlsx"
                    )
                else:
                    st.error("条件が厳しすぎてシフトが組めませんでした。スタッフの人数や夜勤の必要人数を見直してください。")
                    
    except Exception as e:
        st.error(f"⚠️ エラーが発生しました: {e}")
