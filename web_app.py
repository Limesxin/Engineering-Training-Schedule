import streamlit as st
import pandas as pd
import re
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 网页全局与核心参数配置
# ==========================================
st.set_page_config(page_title="工程训练排课系统", page_icon="🛠️", layout="wide")

ADMIN_PASSWORD = "888"

MASTER_FILE = '2025-2026工程训练_0308.xlsx'
SUB_FILE = '各工种场地课表_最新版03082.xlsx'

# 【全新升级核心正则】增加对第4个分组(老师姓氏)的精准提取
GLOBAL_PATTERN = re.compile(r'^([AB]?)\s*(.*?)\s*(\d+\'?-\d+\'?|考\d+)\s*(?:[（\(](.*?)[）\)])?\s*$')


# ==========================================
# 2. 核心功能一：读取与缓存数据
# ==========================================
@st.cache_data
def load_all_data():
    df_master = pd.read_excel(MASTER_FILE, sheet_name='排课表')
    df_master = df_master.fillna("")

    all_sub_sheets = pd.read_excel(SUB_FILE, sheet_name=None)
    for key in all_sub_sheets:
        all_sub_sheets[key] = all_sub_sheets[key].fillna("")

    return df_master, all_sub_sheets


# ==========================================
# 3. 美化渲染模块
# ==========================================
def display_multiline_table(df):
    df_display = df.replace(r'\n', '<br>', regex=True)
    html = df_display.to_html(escape=False, index=False)

    css = """<style>
.custom-excel-table table {
    width: 100%;
    border-collapse: collapse;
    font-size: 14px;
    font-family: sans-serif;
    margin-top: 10px;
}
.custom-excel-table th {
    background-color: #f0f2f6;
    color: #31333F;
    border: 1px solid #e6e9ef;
    padding: 10px;
    text-align: center !important;
    white-space: nowrap;
}
.custom-excel-table td {
    border: 1px solid #e6e9ef;
    padding: 8px;
    text-align: center;
    vertical-align: middle;
    line-height: 1.6;
}
</style>"""
    st.write(css + f'<div class="custom-excel-table" style="overflow-x: auto;">{html}</div>', unsafe_allow_html=True)


# ==========================================
# 4. 核心功能二：联动同步引擎
# ==========================================
def sync_sub_sheets(df_master_latest):
    week_cols = {}
    for col in df_master_latest.columns:
        match = re.search(r'第(\d+)周', col)
        if match:
            week_cols[col] = int(match.group(1))

    workshop_schedule = {}
    days = ['周一', '周二', '周三', '周四', '周五']

    def init_ws(ws_name):
        if ws_name not in workshop_schedule:
            workshop_schedule[ws_name] = {w: {d: {'上午': [], '下午': []} for d in days} for w in range(1, 22)}

    for index, row in df_master_latest.iterrows():
        class_name = str(row.get('教学班名称', '')).strip()
        day = str(row.get('星期', '')).strip()
        if day not in days or not class_name: continue

        for col_name, week_num in week_cols.items():
            cell_val = str(row[col_name])
            if cell_val and cell_val != 'nan':
                for line in cell_val.split('\n'):
                    line = line.strip()
                    if not line: continue

                    match = GLOBAL_PATTERN.match(line)
                    if match:
                        ws_name = match.group(2).strip()
                        time_suffix = match.group(3)

                        if not ws_name:
                            ws_name = '考试' if '考' in time_suffix else '未命名项目'

                        init_ws(ws_name)
                        display_text = f"{class_name} {line}"

                        is_am, is_pm = False, False
                        if '考' in time_suffix:
                            num = int(time_suffix.replace('考', ''))
                            if num <= 4: is_am = True
                            if num >= 5: is_pm = True
                        else:
                            parts = time_suffix.split('-')
                            start = int(parts[0].replace("'", ""))
                            end = int(parts[1].replace("'", ""))
                            if start <= 4: is_am = True
                            if end >= 5: is_pm = True

                        if is_am: workshop_schedule[ws_name][week_num][day]['上午'].append(display_text)
                        if is_pm: workshop_schedule[ws_name][week_num][day]['下午'].append(display_text)

    with pd.ExcelWriter(SUB_FILE, engine='openpyxl') as writer:
        all_ws = list(workshop_schedule.keys())
        if '理论' in all_ws:
            all_ws.remove('理论')
            all_ws = ['理论'] + sorted(all_ws)
        else:
            all_ws = sorted(all_ws)

        for ws in all_ws:
            ws_data = []
            for w in range(1, 22):
                row_dict = {'周次': f'第{w}周'}
                for d in days:
                    am_list = workshop_schedule[ws][w][d]['上午']
                    pm_list = workshop_schedule[ws][w][d]['下午']
                    am_str = "【上午】\n" + "\n".join(am_list) if am_list else "【上午】"
                    pm_str = "【下午】\n" + "\n".join(pm_list) if pm_list else "【下午】"
                    row_dict[d] = "" if not am_list and not pm_list else f"{am_str}\n{'-' * 18}\n{pm_str}"
                ws_data.append(row_dict)

            df_ws = pd.DataFrame(ws_data)
            df_ws.to_excel(writer, index=False, sheet_name=ws)

            worksheet = writer.sheets[ws]
            worksheet.column_dimensions['A'].width = 12
            for col_idx in range(2, 7):
                worksheet.column_dimensions[get_column_letter(col_idx)].width = 32
            fill_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1,
                                           max_col=worksheet.max_column):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                    if cell.row == 1:
                        cell.font = Font(bold=True)
                    elif cell.row % 2 == 0:
                        cell.fill = fill_gray


# ==========================================
# 5. 网页前端 UI 与交互逻辑
# ==========================================
st.title("🛠️ 工程训练排课与场地查询系统")

try:
    df_master, all_sub_sheets = load_all_data()
except FileNotFoundError as e:
    st.error(f"❌ 找不到文件，请确保总表和分表都在同一个文件夹内！")
    st.stop()

st.sidebar.header("⚙️ 导航与控制面板")
# 【新增第三个选项：自由组合】
view_mode = st.sidebar.radio("👀 请选择视图模式：", ["📚 查看大总表", "📍 查看场地分表", "🧑‍🏫 个人专属课表 (自由组合)"])
st.sidebar.markdown("---")

st.sidebar.subheader("🔒 管理员通道")
input_pwd = st.sidebar.text_input("请输入修改密码解锁编辑模式：", type="password")

is_admin = (input_pwd == ADMIN_PASSWORD)
if is_admin:
    st.sidebar.success("✅ 密码正确，已解锁在线编辑权限！")
elif input_pwd:
    st.sidebar.error("❌ 密码错误！")

# ----------------- 模式一：大总表 -----------------
if view_mode == "📚 查看大总表":
    st.subheader("📌 当前视图：【工程训练全景大总表】")

    if is_admin:
        st.info("💡 提示：编辑模式下，双击包含多行文本的单元格会弹出悬浮编辑框。修改完成后，请点击下方保存按钮。")
        edited_df = st.data_editor(
            df_master,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic"
        )

        if st.button("💾 保存修改，并同步更新所有场地课表", type="primary"):
            with st.spinner("正在保存总表并为您重新拆解分发课程，请稍候..."):
                edited_df.to_excel(MASTER_FILE, index=False, sheet_name='排课表')
                sync_sub_sheets(edited_df)
                st.cache_data.clear()
            st.success("🎉 修改已成功保存！所有场地分表已联动更新！页面即将刷新...")
            st.rerun()
    else:
        display_multiline_table(df_master)

# ----------------- 模式二：单一场地分表 -----------------
elif view_mode == "📍 查看场地分表":
    sheet_names = list(all_sub_sheets.keys())
    selected_sheet = st.sidebar.selectbox("🎯 请选择要查看的场地/工种：", sheet_names)
    st.subheader(f"📌 当前视图：【{selected_sheet}】场地课表")
    display_multiline_table(all_sub_sheets[selected_sheet])

# ----------------- 模式三：个人自由组合（全新功能） -----------------
elif view_mode == "🧑‍🏫 个人专属课表 (自由组合)":
    st.subheader("📌 当前视图：【个人专属自由组合课表】")
    st.info("💡 请在下方自由勾选您负责的工种以及您代课的姓氏标识，系统将为您生成大合体课表！")

    # 动态扫描提取所有工种和老师标识
    available_ws = set()
    available_teachers = set()

    week_cols = {col: int(re.search(r'第(\d+)周', col).group(1)) for col in df_master.columns if
                 re.search(r'第(\d+)周', col)}

    for index, row in df_master.iterrows():
        for col in week_cols.keys():
            val = str(row[col])
            if val and val != 'nan':
                for line in val.split('\n'):
                    match = GLOBAL_PATTERN.match(line.strip())
                    if match:
                        ws = match.group(2).strip()
                        time_s = match.group(3)
                        teacher = match.group(4).strip() if match.group(4) else ""

                        if not ws: ws = '考试' if '考' in time_s else '未命名项目'
                        available_ws.add(ws)
                        if teacher: available_teachers.add(teacher)

    # UI：双列多选框
    col1, col2 = st.columns(2)
    with col1:
        selected_ws = st.multiselect("🎯 1. 请选择工种 (可多选):", sorted(list(available_ws)))
    with col2:
        selected_teachers = st.multiselect("👤 2. 请选择代课标识 (可多选):", sorted(list(available_teachers)))

    if not selected_ws and not selected_teachers:
        st.warning("👈 请在上方至少选择一项工种或代课标识以生成您的专属课表。")
    else:
        # 生成合体课表逻辑
        days = ['周一', '周二', '周三', '周四', '周五']
        custom_schedule = {w: {d: {'上午': [], '下午': []} for d in days} for w in range(1, 22)}

        for index, row in df_master.iterrows():
            class_name = str(row.get('教学班名称', '')).strip()
            day = str(row.get('星期', '')).strip()
            if day not in days or not class_name: continue

            for col_name, week_num in week_cols.items():
                val = str(row[col_name])
                if val and val != 'nan':
                    for line in val.split('\n'):
                        line = line.strip()
                        if not line: continue
                        match = GLOBAL_PATTERN.match(line)
                        if match:
                            ws_name = match.group(2).strip()
                            time_suffix = match.group(3)
                            teacher_name = match.group(4).strip() if match.group(4) else ""
                            if not ws_name: ws_name = '考试' if '考' in time_suffix else '未命名项目'

                            # 【核心判断逻辑】属于选中的工种，或者包含选中的代课老师
                            if (ws_name in selected_ws) or (teacher_name in selected_teachers):
                                display_text = f"{class_name} {line}"

                                is_am, is_pm = False, False
                                if '考' in time_suffix:
                                    num = int(time_suffix.replace('考', ''))
                                    if num <= 4: is_am = True
                                    if num >= 5: is_pm = True
                                else:
                                    parts = time_suffix.split('-')
                                    start = int(parts[0].replace("'", ""))
                                    end = int(parts[1].replace("'", ""))
                                    if start <= 4: is_am = True
                                    if end >= 5: is_pm = True

                                if is_am: custom_schedule[week_num][day]['上午'].append(display_text)
                                if is_pm: custom_schedule[week_num][day]['下午'].append(display_text)

        # 组装展示用 DataFrame
        custom_data = []
        for w in range(1, 22):
            row_dict = {'周次': f'第{w}周'}
            for d in days:
                am_list = custom_schedule[w][d]['上午']
                pm_list = custom_schedule[w][d]['下午']
                am_str = "【上午】\n" + "\n".join(am_list) if am_list else "【上午】"
                pm_str = "【下午】\n" + "\n".join(pm_list) if pm_list else "【下午】"
                row_dict[d] = "" if not am_list and not pm_list else f"{am_str}\n{'-' * 18}\n{pm_str}"
            custom_data.append(row_dict)

        df_custom = pd.DataFrame(custom_data)
        st.success("✨ 合成完毕！以下是为您量身定制的专属课表：")
        display_multiline_table(df_custom)