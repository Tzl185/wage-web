import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
import zipfile
import tempfile
import io

st.set_page_config(page_title="工资表处理工具", layout="centered")

# ----------------------
# 处理合并单元格表头
# ----------------------
def get_flat_column_names(filepath, header_row=4):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    merged_dict = {}
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = (
            merged_range.min_col, merged_range.min_row,
            merged_range.max_col, merged_range.max_row
        )
        value = ws.cell(row=min_row, column=min_col).value
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                merged_dict[(row, col)] = value

    col_names = []
    for col in range(1, ws.max_column + 1):
        upper = merged_dict.get((header_row - 1, col)) or ws.cell(row=header_row - 1, column=col).value
        lower = merged_dict.get((header_row, col)) or ws.cell(row=header_row, column=col).value
        parts = []
        if upper: parts.append(str(upper).strip())
        if lower and lower != upper: parts.append(str(lower).strip())
        col_name = "-".join(parts) if parts else f"列{col}"
        col_names.append(col_name)
    return col_names

# ----------------------
# 处理 Excel 文件夹
# ----------------------
def process_file_a(folder_path, output_file="文件A_汇总结果.xlsx"):
    log = io.StringIO()
    all_data = []
    all_values = {}

    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith(('.xlsx')) and not filename.startswith('~$'):
                filepath = os.path.join(root, filename)
                try:
                    print(f"\n📄 正在处理: {filename}", file=log)
                    columns = get_flat_column_names(filepath, header_row=4)
                    df = pd.read_excel(filepath, header=3, engine="openpyxl")
                    df.columns = columns
                    budget_unit_col = columns[1]
                    wage_cols = columns[16:30]
                    df_filtered = df[[budget_unit_col] + wage_cols]
                    df_filtered[wage_cols] = df_filtered[wage_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                    df_grouped = df_filtered.groupby(budget_unit_col).sum()

                    for budget_unit, row in df_grouped.iterrows():
                        for wage_type in wage_cols:
                            value = row[wage_type]
                            wage_type_str = str(wage_type).strip()
                            if "绩效工资" in wage_type_str:
                                wage_type_str = wage_type_str.replace("绩效工资", "基础性绩效")
                            if "行政医疗" in wage_type_str:
                                wage_type_str = wage_type_str.replace("行政医疗", "职工基本医疗（行政）")
                            elif "事业医疗" in wage_type_str:
                                wage_type_str = wage_type_str.replace("事业医疗", "基本医疗（事业）")
                            elif "医疗保险" in wage_type_str:
                                wage_type_str = wage_type_str.replace("医疗保险", "基本医疗")
                            key = (str(budget_unit).strip(), wage_type_str)
                            all_values[key] = value

                    if not df_grouped.empty:
                        all_data.append(df_grouped)

                except Exception as e:
                    print(f"❌ 文件 {filename} 处理失败: {e}", file=log)

    if all_data:
        df_all = pd.concat(all_data)
        df_final = df_all.groupby(df_all.index).sum()
        output_path = os.path.join(folder_path, output_file)
        df_final.to_excel(output_path)
        print(f"\n✅ 汇总完成，已保存到: {output_path}", file=log)
        return output_path, log.getvalue()
    else:
        print("❌ 没有找到有效数据，请检查Excel格式或列名", file=log)
        return None, log.getvalue()

# ----------------------
# Streamlit 界面
# ----------------------
st.title("📊 工资表自动处理工具（带合并单元格支持）")

uploaded_zip = st.file_uploader("请上传包含多个工资表的 ZIP 文件", type=["zip"])

if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "upload.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(tmpdir)

        # 展示解压内容
        st.markdown("### 📂 解压后的文件列表：")
        for root, _, files in os.walk(tmpdir):
            for name in files:
                st.markdown(f"- `{os.path.join(root, name).replace(tmpdir, '')}`")

        # 处理 Excel 文件夹
        st.markdown("---")
        st.markdown("### ⚙️ 正在处理文件A...")

        output_path, log_text = process_file_a(tmpdir)

        st.markdown("### 📋 处理日志：")
        st.text(log_text)

        if output_path:
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 下载汇总结果（文件A）",
                    data=f,
                    file_name="文件A_汇总结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
