import os
import pandas as pd

def convert_xls_to_xlsx_batch(folder):
    for filename in os.listdir(folder):
        if filename.endswith('.xls') and not filename.startswith('~$'):
            xls_path = os.path.join(folder, filename)
            xlsx_path = xls_path + "x"  # 生成 .xlsx 文件名

            print(f"转换文件：{filename} → {os.path.basename(xlsx_path)}")
            try:
                df = pd.read_excel(xls_path, header=None, engine='xlrd')
                df.to_excel(xlsx_path, header=False, index=False, engine='openpyxl')
            except Exception as e:
                print(f"转换失败：{filename}，原因：{e}")

if __name__ == "__main__":
    folder_path = r"你的xls文件夹路径"
    convert_xls_to_xlsx_batch(folder_path)
    print("转换完成！")
