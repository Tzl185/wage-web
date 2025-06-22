import streamlit as st
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import zipfile
import tempfile
import io
import xlrd  # 用于读取 .xls
from openpyxl.workbook import Workbook

st.set_page_config(page_title="工资表处理工具", layout="centered")

# -----------------------
def convert_xls_to_xlsx(xls_path):
    df = pd.read_excel(xls_path, header=None, engine='xlrd')
    new_path = xls_path + "x"
    df.to_excel(new_path, header=False, index=False, engine='openpyxl')
    return new_path

def get_merged_headers(path, header_row=4):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    merged = {}
    for rng in ws.merged_cells.ranges:
        for row in range(rng.min_row, rng.max_row + 1):
            for col in range(rng.min_col, rng.max_col + 1):
                merged[(row, col)] = ws.cell(rng.min_row, rng.min_col).value
    headers = []
    for col in range(1, ws.max_column + 1):
        top = merged.get((header_row - 1, col)) or ws.cell(header_row - 1, col).value
        bottom = merged.get((header_row, col)) or ws.cell(header_row, col).value
        parts = [str(p).strip() for p in [top, bottom] if p and p != top]
        headers.append("-".join(parts) if parts else f"列{col}")
    return headers

# -----------------------
def process_file_a(folder_path, output_file="文件A_汇总结果.xlsx", log_area=None):
    all_data = []
    all_values = {}

    for root, _, files in os.walk(folder_path):
        for fname in files:
            if fname.endswith(('.xls', '.xlsx')) and not fname.startswith('~$'):
                fpath = os.path.join(root, fname)
                try:
                    if log_area: log_area.write(f"\n📄 正在处理: {fname}\n")

                    if fpath.endswith(".xls"):
                        fpath = convert_xls_to_xlsx(fpath)
                        if log_area: log_area.write("📎 已转换为 .xlsx\n")

                    headers = get_merged_headers(fpath)
                    df = pd.read_excel(fpath, header=3, engine='openpyxl')
                    df.columns = headers

                    budget_col = headers[1]
                    wage_cols = headers[16:30]

                    df_filtered = df[[budget_col] + wage_cols]
                    df_filtered[wage_cols] = df_filtered[wage_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
                    df_grouped = df_filtered.groupby(budget_col).sum()

                    for unit, row in df_grouped.iterrows():
                        for wage_type in wage_cols:
                            value = row[wage_type]
                            name = str(wage_type).strip()
                            name = name.replace("绩效工资", "基础性绩效")
                            name = name.replace("行政医疗", "职工基本医疗（行政）")
                            name = name.replace("事业医疗", "基本医疗（事业）")
                            name = name.replace("医疗保险", "基本医疗")
                            all_values[(str(unit).strip(), name)] = value

                    if not df_grouped.empty:
                        all_data.append(df_grouped)

                    if log_area: log_area.write("✅ 处理完成\n")

                except Exception as e:
                    if log_area: log_area.write(f"❌ 错误: {e}\n")

    if all_data:
        df_all = pd.concat(all_data)
        df_final = df_all.groupby(df_all.index).sum()
        out_path = os.path.join(folder_path, output_file)
        df_final.to_excel(out_path)
        if log_area: log_area.write(f"\n✅ 汇总完成，保存至：{output_file}\n")
        return out_path
    else:
        if log_area: log_area.write("❌ 没有找到有效数据，请检查 Excel 格式或列名\n")
        return None

# -----------------------
st.title("📊 工资表自动处理工具（实时日志）")
uploaded_zip = st.file_uploader("请上传包含工资表的压缩包（.zip）", type=["zip"])

if uploaded_zip:
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "upload.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())

        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(tmpdir)

        st.markdown("### 📂 解压内容如下：")
        for root, _, files in os.walk(tmpdir):
            for f in files:
                st.markdown(f"- `{os.path.join(root, f).replace(tmpdir, '')}`")

        st.markdown("---")
        st.markdown("### 🔧 正在分析并生成文件A...")
        log_placeholder = st.empty()

        with st.spinner("⏳ 正在处理中..."):
            log_text = st.empty()
            file_a_path = process_file_a(tmpdir, log_area=log_text)

        st.markdown("### 📜 处理日志：")
        st.text(log_text)

        if file_a_path:
            with open(file_a_path, "rb") as f:
                st.download_button(
                    label="📥 下载汇总结果文件A",
                    data=f,
                    file_name="文件A_汇总结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
