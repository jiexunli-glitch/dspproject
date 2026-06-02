from pathlib import Path
import pandas as pd

input_folder = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/成果報告/新增第五層結果")
output_folder = input_folder / "亮點分析"
output_folder.mkdir(parents=True, exist_ok=True)

excel_files = list(input_folder.glob("*.xlsx"))

print("找到檔案數：", len(excel_files))

for file in excel_files:
    try:
        df = pd.read_excel(file)

        # 檢查欄位是否存在
        if "第四層" not in df.columns:
            print(f"跳過：{file.name}，沒有「第四層」欄位")
            print(df.columns.tolist())
            continue

        # 清理文字
        df["第四層_clean"] = (
            df["第四層"]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace("　", "", regex=False)
            .str.replace("\n", "", regex=False)
            .str.strip()
        )

        # 篩選：只要第四層包含「亮點成果」
        df_filtered = df[df["第四層_clean"].str.contains("(四)亮點成果", na=False)]

        # 不要輸出輔助欄位
        df_filtered = df_filtered.drop(columns=["第四層_clean"])

        output_name = file.stem + "_亮點成果.xlsx"
        output_path = output_folder / output_name

        df_filtered.to_excel(output_path, index=False)

        print(f"完成：{output_name}")
        print(f"保留筆數：{len(df_filtered)}")

        # 印出第四層有哪些值，方便檢查
        print("第四層出現的值：")
        print(df["(四)亮點成果"].dropna().unique()[:10])

    except Exception as e:
        print(f"錯誤：{file.name}")
        print(e)

print("全部完成！")


#%%
from pathlib import Path
import pandas as pd

input_folder = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/成果報告/新增第五層結果")
output_folder = input_folder / "亮點分析"
output_folder.mkdir(parents=True, exist_ok=True)

excel_files = list(input_folder.glob("*.xlsx"))

print("找到檔案數：", len(excel_files))

for file in excel_files:
    try:
        df = pd.read_excel(file)

        # 清理第四層文字
        fourth = (
            df["第四層"]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace("　", "", regex=False)
            .str.replace("\n", "", regex=False)
            .str.strip()
        )

        # 重點：抓「亮點」即可
        df_filtered = df[fourth.str.contains("亮點", na=False)]

        output_name = file.stem + "_亮點成果.xlsx"
        output_path = output_folder / output_name

        df_filtered.to_excel(output_path, index=False)

        print(f"完成：{output_name}")
        print(f"保留筆數：{len(df_filtered)}")

    except Exception as e:
        print(f"錯誤：{file.name}")
        print(e)

print("全部完成！")
# %%
from pathlib import Path
import pandas as pd

input_folder = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/成果報告/新增第五層結果/112到113")
output_folder = input_folder / "亮點分析"
output_folder.mkdir(parents=True, exist_ok=True)

excel_files = list(input_folder.glob("*.xlsx"))

print("找到檔案數：", len(excel_files))

for file in excel_files:
    try:
        df = pd.read_excel(file)

        # 清理第四層文字
        fourth = (
            df["第三層"]
            .astype(str)
            .str.replace(" ", "", regex=False)
            .str.replace("　", "", regex=False)
            .str.replace("\n", "", regex=False)
            .str.strip()
        )

        # 重點：抓「亮點」即可
        df_filtered = df[fourth.str.contains("亮點", na=False)]

        output_name = file.stem + "_亮點成果.xlsx"
        output_path = output_folder / output_name

        df_filtered.to_excel(output_path, index=False)

        print(f"完成：{output_name}")
        print(f"保留筆數：{len(df_filtered)}")

    except Exception as e:
        print(f"錯誤：{file.name}")
        print(e)

print("全部完成！")
# %%
