import streamlit as st
import pandas as pd
import zipfile
import os
import shutil
from openpyxl import load_workbook

def process_file_a(folder_path, output_file="文件A_汇总结果.xlsx"):
    all_data = []
    positive_values = {}

    for filename in os.listdir(folder_path):
        if filename.endswith(('.xls', '.xlsx')) and not filename.startswith('~$'):
            filepath = os.path.join(folder_path, filename)
            try:
                df = pd.read_excel(filepath, header=3)
                budget_unit_col = df.columns[1]
                wage_cols = df.columns[16:30]
                df_filtered = df[[budget_unit_col] + list(wage_cols)]
                df_filtered[wage_cols] = df_filtered[wage_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                df_grouped = df_filtered.groupby(budget_unit_col).sum()
                all_data.append(df_grouped)
            except Exception as e:
                print(f"{filename} 错误: {e}")

    if all_data:
        df_all = pd.concat(all_data)
        df_final = df_all.groupby(df_all.index).sum()
        output_path = os.path.join(folder_path, output_file)
        df_final.to_excel(output_path)

        for budget_unit, row in df_final.iterrows():
            for wage_type in wage_cols:
                value = row[wage_type]
                if value > 0:
                    if "绩效工资" in wage_type:
                        wage_type = wage_type.replace("绩效工资", "基础性绩效")
                    key = (str(budget_unit).strip(), str(wage_type).strip())
                    positive_values[key] = value
        return output_path, positive_values
    else:
        return None, None

def update_file_b(file_a_path, file_b_path, output_dir):
    df_a = pd.read_excel(file_a_path, index_col=0)
    wage_cols = df_a.columns

    positive_values = {}
    for budget_unit, row in df_a.iterrows():
        for wage_type in wage_cols:
            value = row[wage_type]
            if value > 0:
                if "绩效工资" in wage_type:
                    wage_type = wage_type.replace("绩效工资", "基础性绩效")
                key = (str(budget_unit).strip(), str(wage_type).strip())
                positive_values[key] = value

    wb = load_workbook(file_b_path)
    sheet = wb.active
    j_col_index = 10
    for row_idx in range(2, sheet.max_row + 1):
        unit_info = str(sheet.cell(row=row_idx, column=1).value or "")
        budget_project = str(sheet.cell(row=row_idx, column=2).value or "")
        unit_info_cleaned = unit_info.replace("-", "").replace(" ", "")
        for (budget_unit, wage_type), value in positive_values.items():
            budget_unit_cleaned = budget_unit.replace(" ", "")
            if ((budget_unit_cleaned in unit_info_cleaned or unit_info_cleaned in budget_unit_cleaned)
                    and wage_type in budget_project):
                sheet.cell(row=row_idx, column=j_col_index).value = value
                break

    output_path = os.path.join(output_dir, "updated_" + os.path.basename(file_b_path))
    wb.save(output_path)
    return output_path

# -------- Streamlit 界面部分 --------
st.title("Excel 工资处理自动化网站")
st.write("上传压缩包和模板文件，自动生成汇总文件和更新后的模板。")

zip_file = st.file_uploader("上传 Excel 压缩包（多个工资表）", type="zip")
template_file = st.file_uploader("上传模板文件（项目细化导入模板）", type=["xls", "xlsx"])

if st.button("运行程序"):
    if zip_file and template_file:
        with st.spinner("正在处理，请稍候..."):
            work_dir = "temp_upload"
            os.makedirs(work_dir, exist_ok=True)

            zip_path = os.path.join(work_dir, "uploaded.zip")
            with open(zip_path, "wb") as f:
                f.write(zip_file.read())

            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(work_dir)

            template_path = os.path.join(work_dir, "template.xlsx")
            with open(template_path, "wb") as f:
                f.write(template_file.read())

            a_path, _ = process_file_a(work_dir)
            if a_path:
                b_path = update_file_b(a_path, template_path, work_dir)

                with open(a_path, "rb") as f:
                    st.download_button("下载汇总文件A", f, file_name="文件A_汇总结果.xlsx")

                with open(b_path, "rb") as f:
                    st.download_button("下载更新后的文件B", f, file_name="updated_项目细化导入模板.xlsx")
            else:
                st.error("处理失败，请检查上传的文件格式")

            shutil.rmtree(work_dir)
    else:
        st.warning("请上传所有文件")
