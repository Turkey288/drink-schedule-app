import streamlit as st
import subprocess
import pandas as pd
import os
import sys
from datetime import datetime

st.set_page_config(page_title="飲料店排班系統", layout="wide")

st.title("🥤 飲料店排班系統 第一版")

st.sidebar.title("功能選單")
page = st.sidebar.radio(
    "請選擇功能",
    [
        "店長排班後台",
        "員工查看班表",
        "排班檢查",
        "人力分析",
        "LINE公告",
        "月曆班表",
        "排班調整",
    ],
)


def 找最新班表():
    output_files = [
        f for f in os.listdir(".")
        if f.endswith(".xlsx") and "自動排班結果" in f
    ]

    if os.path.exists("output"):
        output_files += [
            os.path.join("output", f)
            for f in os.listdir("output")
            if f.endswith(".xlsx") and "自動排班結果" in f
        ]

    if not output_files:
        return None

    return max(output_files, key=os.path.getmtime)


def 格式化日期(value):
    text = str(value).strip()
    text = text.replace("月", "/").replace("日", "")

    if "/" in text:
        parts = text.split("/")
        if len(parts) >= 2:
            try:
                return f"{int(parts[0])}/{int(parts[1])}"
            except Exception:
                return text

    return text


def 轉換班別(value):
    value = str(value).strip()

    if value == "":
        return "上班"
    if value == "例":
        return "例假"
    if value == "休":
        return "休息日"
    if value in ["○", "o", "O"]:
        return "可加班休"
    if value == "空":
        return "空班"

    return value


def 班別轉原始符號(value):
    對照 = {
        "上班": "",
        "例假": "例",
        "休息日": "休",
        "可加班休": "○",
        "空班": "空",
    }
    return 對照.get(value, "")


def 是否算上班(value):
    value = str(value).strip()
    return value in ["", "○", "o", "O"]


def 是否可加班休(value):
    value = str(value).strip()
    return value in ["○", "o", "O"]


def 是否技術人員(姓名):
    技術名單 = [
        "王森弘",
        "古峻燐",
        "萬郁茹",
        "彭笙祐",
        "劉裕平",
        "莊絨媗",
    ]
    return 姓名 in 技術名單


def 讀取班表檔案():
    latest_file = 找最新班表()

    if not latest_file:
        return None, None

    df = pd.read_excel(
        latest_file,
        sheet_name=0,
        header=None
    ).fillna("")

    return latest_file, df


def 取得員工資料(df):
    合法身份 = ["副店長", "組長", "正職", "PT", "PT新人"]

    return df[
        df[1].astype(str).str.strip().isin(合法身份)
    ].copy()


def 取得日期欄位(df):
    日期列 = df.iloc[2]
    日期欄位 = []

    for col in range(2, df.shape[1]):
        日期 = str(日期列[col]).strip()
        if 日期 and "/" in 日期:
            日期欄位.append(col)

    return 日期欄位


def 取得日期欄位對照(df):
    日期列 = df.iloc[2]
    對照 = {}

    for col in 取得日期欄位(df):
        日期 = 格式化日期(日期列[col])
        對照[日期] = col

    return 對照


def 建立人力分析表(df):
    員工資料 = 取得員工資料(df)
    日期欄位 = 取得日期欄位(df)
    日期列 = df.iloc[2]

    分析資料 = []

    for col in 日期欄位:
        日期 = 格式化日期(日期列[col])
        上班人數 = 0
        技術人數 = 0

        for _, row in 員工資料.iterrows():
            姓名 = str(row[0]).strip()
            班別 = str(row[col]).strip()

            if 是否算上班(班別):
                上班人數 += 1

                if 是否技術人員(姓名):
                    技術人數 += 1

        分析資料.append({
            "日期": 日期,
            "上班人數": 上班人數,
            "技術人數": 技術人數,
        })

    return pd.DataFrame(分析資料)


def 建立人力警示表(分析表):
    警示資料 = []

    for _, row in 分析表.iterrows():
        日期 = row["日期"]
        上班人數 = row["上班人數"]
        技術人數 = row["技術人數"]

        if 上班人數 < 7:
            警示資料.append({
                "日期": 日期,
                "問題": "人數不足",
                "目前人數": 上班人數,
                "標準": "至少7人",
            })

        if 技術人數 < 3:
            警示資料.append({
                "日期": 日期,
                "問題": "技術不足",
                "目前人數": 技術人數,
                "標準": "至少3人",
            })

    return pd.DataFrame(警示資料)


def 建立可加班推薦表(df, 警示表):
    員工資料 = 取得員工資料(df)
    日期欄位對照 = 取得日期欄位對照(df)

    推薦資料 = []

    if 警示表.empty:
        return pd.DataFrame(推薦資料)

    for _, warn in 警示表.iterrows():
        日期 = 格式化日期(warn["日期"])

        if 日期 not in 日期欄位對照:
            continue

        col = 日期欄位對照[日期]

        技術推薦 = []
        一般推薦 = []

        for _, emp in 員工資料.iterrows():
            姓名 = str(emp[0]).strip()
            身份 = str(emp[1]).strip()
            班別 = str(emp[col]).strip()

            if 是否可加班休(班別):
                if 是否技術人員(姓名):
                    技術推薦.append(f"{姓名}（技術）")
                else:
                    一般推薦.append(f"{姓名}（{身份}）")

        推薦名單 = 技術推薦 + 一般推薦

        推薦資料.append({
            "日期": 日期,
            "問題": warn["問題"],
            "目前人數": warn["目前人數"],
            "推薦可加班人員": "、".join(推薦名單) if 推薦名單 else "無可加班休人員",
        })

    return pd.DataFrame(推薦資料)


def 班表評分(df):
    分析表 = 建立人力分析表(df)
    員工資料 = 取得員工資料(df)
    日期欄位 = 取得日期欄位(df)

    score = 100
    問題 = []

    缺工天數 = len(分析表[分析表["上班人數"] < 7])
    if 缺工天數 > 0:
        score -= 缺工天數 * 3
        問題.append(f"人數低於7人的日期：{缺工天數} 天")

    技術不足天數 = len(分析表[分析表["技術人數"] < 3])
    if 技術不足天數 > 0:
        score -= 技術不足天數 * 5
        問題.append(f"技術人數低於3人的日期：{技術不足天數} 天")

    for _, row in 員工資料.iterrows():
        姓名 = str(row[0]).strip()
        身份 = str(row[1]).strip()
        上班天數 = 0

        for col in 日期欄位:
            班別 = str(row[col]).strip()
            if 是否算上班(班別):
                上班天數 += 1

        if 身份 in ["副店長", "組長", "正職"]:
            if 上班天數 > 26:
                score -= 2
                問題.append(f"{姓名} 上班天數偏高：{上班天數} 天")
            if 上班天數 < 18:
                score -= 2
                問題.append(f"{姓名} 上班天數偏低：{上班天數} 天")

        if 身份 in ["PT", "PT新人"]:
            if 上班天數 > 26:
                score -= 2
                問題.append(f"{姓名} PT上班天數偏高：{上班天數} 天")
            if 上班天數 < 4:
                score -= 2
                問題.append(f"{姓名} PT上班天數偏低：{上班天數} 天")

    if score < 0:
        score = 0

    return score, 問題


if page == "店長排班後台":
    st.header("店長排班後台")

    uploaded_file = st.file_uploader("上傳排班 Excel 檔", type=["xlsx"])

    if uploaded_file:
        os.makedirs("uploads", exist_ok=True)
        os.makedirs("output", exist_ok=True)

        input_path = "uploads/上傳資料.xlsx"

        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success("Excel 上傳成功")

        if st.button("🚀 產生班表"):
            with st.spinner("正在產生班表..."):
                result = subprocess.run(
                    [sys.executable, "auto_schedule.py"],
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )

            if result.returncode == 0:
                st.success("班表產生成功")

                if result.stdout:
                    with st.expander("📋 執行紀錄"):
                        st.code(result.stdout)

                latest_file, preview_df = 讀取班表檔案()

                if latest_file is not None:
                    st.subheader("📅 班表預覽")
                    st.dataframe(preview_df, use_container_width=True, height=600)

                    with open(latest_file, "rb") as f:
                        st.download_button(
                            "📥 下載班表 Excel",
                            data=f,
                            file_name=os.path.basename(latest_file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                else:
                    st.warning("有執行成功，但找不到輸出的 Excel 檔。")

            else:
                st.error("排班失敗")
                if result.stderr:
                    st.code(result.stderr)


if page == "員工查看班表":
    st.header("👤 員工查看班表")

    latest_file, df = 讀取班表檔案()

    if latest_file is None:
        st.warning("目前找不到已產生的班表，請先到店長排班後台產生班表。")
    else:
        try:
            員工資料 = 取得員工資料(df)

            if 員工資料.empty:
                st.warning("找不到員工資料，請確認班表格式是否正確。")
            else:
                員工名單 = 員工資料[0].astype(str).tolist()
                selected_name = st.selectbox("請選擇員工", 員工名單)

                row = 員工資料[員工資料[0].astype(str) == selected_name]

                日期列 = df.iloc[2]
                星期列 = df.iloc[3]
                日期欄位 = 取得日期欄位(df)

                result_rows = []

                for col in 日期欄位:
                    日期 = 格式化日期(日期列[col])
                    星期 = 星期列[col]
                    原始班別 = row.iloc[0, col]
                    班別 = 轉換班別(原始班別)

                    result_rows.append({
                        "日期": 日期,
                        "星期": 星期,
                        "班別": 班別,
                    })

                personal_df = pd.DataFrame(result_rows)

                st.subheader(f"📅 {selected_name} 的班表")

                col1, col2, col3, col4, col5, col6 = st.columns(6)
                col1.metric("例假", (personal_df["班別"] == "例假").sum())
                col2.metric("休息日", (personal_df["班別"] == "休息日").sum())
                col3.metric("可加班休", (personal_df["班別"] == "可加班休").sum())
                col4.metric("空班", (personal_df["班別"] == "空班").sum())
                col5.metric("上班", (personal_df["班別"] == "上班").sum())
                col6.metric("總天數", len(personal_df))

                st.dataframe(personal_df, use_container_width=True, height=600)

        except Exception as e:
            st.error("讀取員工班表失敗")
            st.code(str(e))


if page == "排班檢查":
    st.header("⚠️ 排班檢查")

    latest_file = 找最新班表()

    if latest_file is None:
        st.warning("目前找不到已產生的班表，請先到店長排班後台產生班表。")
    else:
        try:
            check_df = pd.read_excel(latest_file, sheet_name="違規檢查").fillna("")

            if check_df.empty:
                st.success("未發現排班問題")
            else:
                st.subheader("📋 違規檢查明細")
                st.dataframe(check_df, use_container_width=True, height=600)

        except Exception as e:
            st.warning("找不到『違規檢查』工作表，或讀取失敗。")
            st.code(str(e))


if page == "人力分析":
    st.header("📊 人力分析儀表板")

    latest_file, df = 讀取班表檔案()

    if latest_file is None:
        st.warning("目前找不到已產生的班表，請先到店長排班後台產生班表。")
    else:
        try:
            分析表 = 建立人力分析表(df)
            警示表 = 建立人力警示表(分析表)
            推薦表 = 建立可加班推薦表(df, 警示表)
            score, issues = 班表評分(df)

            st.subheader("🏆 班表健康度")

            if score >= 90:
                st.success(f"班表評分：{score} 分")
            elif score >= 75:
                st.warning(f"班表評分：{score} 分")
            else:
                st.error(f"班表評分：{score} 分")

            if issues:
                with st.expander("查看扣分原因"):
                    for item in issues:
                        st.write("•", item)

            st.divider()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("平均上班人數", round(分析表["上班人數"].mean(), 1))
            c2.metric("平均技術人數", round(分析表["技術人數"].mean(), 1))
            c3.metric("最低上班人數", int(分析表["上班人數"].min()))
            c4.metric("最低技術人數", int(分析表["技術人數"].min()))

            st.divider()

            st.subheader("👥 每日上班人數")
            st.line_chart(分析表.set_index("日期")["上班人數"])

            st.subheader("🧋 每日技術人數")
            st.line_chart(分析表.set_index("日期")["技術人數"])

            st.divider()

            st.subheader("🚨 人力不足警示")

            if 警示表.empty:
                st.success("本月無人力不足日期")
            else:
                st.error(f"共發現 {len(警示表)} 筆人力警示")
                st.dataframe(警示表, use_container_width=True)

            st.divider()

            st.subheader("📌 可加班人員推薦")

            if 推薦表.empty:
                st.success("目前無需推薦加班人員")
            else:
                st.dataframe(推薦表, use_container_width=True)

            st.divider()

            st.subheader("📋 完整人力表")
            st.dataframe(分析表, use_container_width=True, height=600)

        except Exception as e:
            st.error("分析失敗")
            st.code(str(e))


if page == "LINE公告":
    st.header("📢 LINE 群組公告產生器")

    latest_file, df = 讀取班表檔案()

    if latest_file is None:
        st.warning("目前找不到已產生的班表，請先產生班表。")
    else:
        try:
            分析表 = 建立人力分析表(df)
            警示表 = 建立人力警示表(分析表)
            推薦表 = 建立可加班推薦表(df, 警示表)

            月份文字 = "本月"

            if not 分析表.empty:
                第一日 = str(分析表.iloc[0]["日期"])
                if "/" in 第一日:
                    月份文字 = 第一日.split("/")[0] + "月"

            警示日期 = []
            if not 警示表.empty:
                警示日期 = 警示表["日期"].astype(str).unique().tolist()

            警示日期文字 = "、".join(警示日期) if 警示日期 else "無"

            公告文字 = f"""【{月份文字}班表已更新】

請各位同仁至排班系統查看個人班表。

○ 代表「可加班休」：
原則上仍屬休假，但若當日人力不足，會優先詢問是否可支援。

⚠️ 人力需注意日期：
{警示日期文字}

如有班表問題，請私訊店長協調，謝謝大家。
"""

            st.subheader("📋 公告預覽")
            st.text_area("可直接複製到 LINE 群組", 公告文字, height=260)

            st.subheader("📌 可加班推薦明細")

            if 推薦表.empty:
                st.success("目前沒有需要特別找人支援的日期。")
            else:
                st.dataframe(推薦表, use_container_width=True)

        except Exception as e:
            st.error("公告產生失敗")
            st.code(str(e))


if page == "月曆班表":
    from streamlit.components.v1 import html

    st.header("📅 員工月曆班表")

    latest_file, df = 讀取班表檔案()

    if latest_file is None:
        st.warning("目前找不到已產生的班表，請先產生班表。")
    else:
        try:
            員工資料 = 取得員工資料(df)

            員工名單 = 員工資料[0].astype(str).tolist()
            selected_name = st.selectbox("請選擇員工", 員工名單)

            row = 員工資料[員工資料[0].astype(str) == selected_name]

            日期列 = df.iloc[2]
            星期列 = df.iloc[3]
            日期欄位 = 取得日期欄位(df)

            星期順序 = ["一", "二", "三", "四", "五", "六", "日"]

            班表資料 = []

            for col in 日期欄位:
                日期 = 格式化日期(日期列[col])
                星期 = str(星期列[col]).strip()
                原始班別 = str(row.iloc[0, col]).strip()
                班別 = 轉換班別(原始班別)

                班表資料.append({
                    "日期": 日期,
                    "星期": 星期,
                    "班別": 班別,
                })

            st.subheader(f"🗓️ {selected_name} 的月曆班表")

            個人表 = pd.DataFrame(班表資料)

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("例假", (個人表["班別"] == "例假").sum())
            col2.metric("休息日", (個人表["班別"] == "休息日").sum())
            col3.metric("可加班休", (個人表["班別"] == "可加班休").sum())
            col4.metric("空班", (個人表["班別"] == "空班").sum())
            col5.metric("上班", (個人表["班別"] == "上班").sum())

            css = """
            <style>
            body{background:#0e1117;color:white;}
            .calendar{display:grid;grid-template-columns:repeat(7,1fr);gap:10px;margin-top:15px;}
            .weekday{background:#262730;text-align:center;padding:10px;border-radius:10px;font-weight:bold;color:white;}
            .day{min-height:95px;padding:10px;border-radius:12px;border:1px solid #333;background:#1e1f26;color:white;}
            .today{border:2px solid #00d4ff;box-shadow:0 0 15px #00d4ff55;}
            .date{font-weight:bold;margin-bottom:8px;}
            .work{color:#4ade80;font-weight:bold;}
            .holiday{color:#ff6b6b;font-weight:bold;}
            .rest{color:#4ecdc4;font-weight:bold;}
            .support{color:#ffd166;font-weight:bold;}
            .empty{color:#999;font-weight:bold;}
            </style>
            """

            html_content = '<div class="calendar">'

            for w in 星期順序:
                html_content += f'<div class="weekday">{w}</div>'

            第一個星期 = 班表資料[0]["星期"]

            if 第一個星期 in 星期順序:
                空白格數 = 星期順序.index(第一個星期)

                for _ in range(空白格數):
                    html_content += "<div></div>"

            今天 = f"{datetime.now().month}/{datetime.now().day}"

            for item in 班表資料:
                日期 = item["日期"]
                班別 = item["班別"]

                if 班別 == "上班":
                    class_name = "work"
                elif 班別 == "例假":
                    class_name = "holiday"
                elif 班別 == "休息日":
                    class_name = "rest"
                elif 班別 == "可加班休":
                    class_name = "support"
                elif 班別 == "空班":
                    class_name = "empty"
                else:
                    class_name = "work"

                day_class = "day today" if 日期 == 今天 else "day"

                html_content += f"""
                <div class="{day_class}">
                    <div class="date">{日期}</div>
                    <div class="{class_name}">{班別}</div>
                </div>
                """

            html_content += "</div>"

            html(css + html_content, height=900, scrolling=True)

            st.caption("紅色＝例假｜綠色＝休息日｜黃色＝可加班休｜灰色＝空班｜藍框＝今天")

        except Exception as e:
            st.error("月曆班表讀取失敗")
            st.code(str(e))


if page == "排班調整":
    st.header("🔧 排班調整中心")

    latest_file, df = 讀取班表檔案()

    if latest_file is None:
        st.warning("目前找不到已產生的班表，請先產生班表。")
    else:
        try:
            員工資料 = 取得員工資料(df)

            員工名單 = 員工資料[0].astype(str).tolist()
            selected_name = st.selectbox("選擇員工", 員工名單)

            日期欄位 = 取得日期欄位(df)
            日期列 = df.iloc[2]

            日期選單 = [
                格式化日期(日期列[col])
                for col in 日期欄位
            ]

            selected_date = st.selectbox("選擇日期", 日期選單)

            員工index = 員工資料[
                員工資料[0].astype(str) == selected_name
            ].index[0]

            日期col = None

            for col in 日期欄位:
                if 格式化日期(df.iloc[2, col]) == selected_date:
                    日期col = col
                    break

            目前班別 = 轉換班別(df.iat[員工index, 日期col])

            新班別 = st.selectbox(
                "修改班別",
                ["上班", "例假", "休息日", "可加班休", "空班"],
                index=["上班", "例假", "休息日", "可加班休", "空班"].index(目前班別)
                if 目前班別 in ["上班", "例假", "休息日", "可加班休", "空班"]
                else 0,
            )

            st.info(f"目前班別：{目前班別}")

            if st.button("💾 儲存修改"):
                df.iat[員工index, 日期col] = 班別轉原始符號(新班別)

                df.to_excel(
                    latest_file,
                    index=False,
                    header=False,
                )

                st.success(f"{selected_name} {selected_date} 已修改為 {新班別}")

        except Exception as e:
            st.error("排班調整失敗")
            st.code(str(e))
