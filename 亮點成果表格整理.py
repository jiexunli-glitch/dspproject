#%%
from pathlib import Path
import pandas as pd

input_folder = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/亮點成果表格")
output_folder = input_folder / "亮點分析表格"
output_folder.mkdir(parents=True, exist_ok=True)

excel_files = list(input_folder.glob("*.xlsx"))

print("找到檔案數：", len(excel_files))
# %%
file_path = r"C:/Users/0128M/OneDrive/桌面/智庫/week4/亮點成果表格/第二期製造部門溫室氣體減量行動方案108年成果報告_階層劃分資料_新增第五層_亮點成果.xlsx"
df = pd.read_excel(file_path)
print(df.columns.tolist())
print(df.head())

# %%
#ai應用
# =========================
# 格式一
# =========================

import pandas as pd
import json
import re
import time
from pathlib import Path

from google.oauth2 import service_account
import vertexai
from vertexai.generative_models import GenerativeModel


# =========================
# 1. 基本設定
# =========================

KEY_PATH = "C:/Users/0128M\OneDrive/桌面/智庫/api/internship-495506-1c354f4e33b5.json"
credentials = service_account.Credentials.from_service_account_file(KEY_PATH)
PROJECT_ID = 'internship-495506'#"你的專案ID" 
LOCATION = "us-central1"

INPUT_FOLDER = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/亮點成果表格/格式一")
OUTPUT_FOLDER = INPUT_FOLDER / "G亮點分析結果_格式一"
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

credentials = service_account.Credentials.from_service_account_file(KEY_PATH)

vertexai.init(
    project=PROJECT_ID,
    location=LOCATION,
    credentials=credentials
)


model = GenerativeModel("gemini-2.5-flash")


# =========================
# 2. 輸出欄位
# =========================

extract_keys = [
    "既有問題",
    "製程改善策略",
    "製程改善策略的執行方法",
    "能源轉換技術",
    "能源轉換技術的執行方法",
    "循環經濟措施",
    "循環經濟措施的執行方法",
    "投資金額",
    "回收年限",
    "關鍵效能參數",
    "計畫成效",
    "減碳量",
    "對應的國際規範連結",
    "抽取依據"
]


# =========================
# 3. 工具函數
# =========================

def clean_value(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    if x.lower() == "nan":
        return ""
    return x


def build_document_id(pdf_name):
    pdf_name = clean_value(pdf_name)
    if pdf_name == "":
        return "unknown_document"

    text = re.sub(r"[^\w\u4e00-\u9fff]+", "_", pdf_name)
    return text.strip("_")


def build_prompt(source_text):
    return f"""
你是一個資訊抽取系統，不是分析系統。

你的任務是從「原文」中抽取已明確出現的資訊。

請嚴格遵守：
1. 只能抽取原文中明確出現的內容。
2. 不可以推論。
3. 不可以補充背景知識。
4. 不可以改寫成自己的說法。
5. 如果單一個column得到的抽取結果有很多項，請以圓點的條列式進行整理。
6. 如果原文沒有明確提到，請填「未能了解」。
7. 請輸出合法 JSON。
8. 不要加上 ```json。
9. 不要輸出 Markdown。
10.不要加任何解釋文字，只能輸出 JSON。

欄位定義：

既有問題：指產業在進行改善前所面臨的技術瓶頸或能源管理痛點，也就是該產業原本就存在的問題、本次行動方案要解決的目標問題。先寫問題的名稱，再說明問題的內容。
製程改善策略：指達成減碳目標的具體措施之名稱。
製程改善策略的執行方法：指在進行製程改善策略時，為了達成減碳的目標具體執行的方式。
能源轉換技術：指為了調整能源結構，將高碳燃料轉向低碳或零碳能源的技術方法之名稱。
能源轉換技術的執行方法：指在進行能源轉換技術時，為了調整能源結構而具體執行的方式。
循環經濟措施：為了能達成資源的重複利用、延長資源生命週期而實施的措施。
循環經濟措施的執行方法：指在進行循環經濟措施時，為了達成資源的重複利用、延長資源生命週期具體執行的方式。
投資金額：指企業為了實施減碳改善計畫所投入的資金總額，包含設備購置費、政府補助款及企業自籌款。請區分並標註「政府補助款」與「廠商新增投資額（自籌款）」。
回收年限：指該項投資預計多少年可以透過節省的成本回收。
計畫成效：指實施該計畫後的效果與進展，以增加的經濟與資源價值的量化數據表示。包含節能率、投資回收年限、產值增加額或資源回收率。
關鍵效能參數:內文中有提到可以用來評估該策略會用到的單位。
減碳量：指透過改善措施後，每年所減少排放的溫室氣體總量。
對應的國際規範鏈結：指該減碳成果如何協助產業對接國際減碳要求或倡議。
抽取依據：請引用原文中最能支持上述抽取結果的短句。

請輸出以下 JSON：

{{
  "既有問題": "",
  "製程改善策略": "",
  "製程改善策略的執行方法": "",
  "能源轉換技術": "",
  "能源轉換技術的執行方法": "",
  "循環經濟措施": "",
  "循環經濟措施的執行方法": "",
  "投資金額": "",
  "回收年限": "",
  "關鍵效能參數": "",
  "計畫成效": "",
  "減碳量": "",
  "對應的國際規範連結": "",
  "抽取依據": ""
}}

原文如下：

{source_text}
"""


def parse_json_response(text):
    text = text.strip()

    text = re.sub(r"^```json", "", text)
    text = re.sub(r"^```", "", text)
    text = re.sub(r"```$", "", text)
    text = text.strip()

    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        text = match.group(0)

    try:
        data = json.loads(text)

        for key in extract_keys:
            if key not in data:
                data[key] = "未能了解"
            # 強制轉成純文字
            for key in data:

                if isinstance(data[key], list):

                   data[key] = "；".join(
                       [str(x) for x in data[key]]
                   )

                elif isinstance(data[key], dict):
 
                    temp = []

                    for k, v in data[key].items():
                        temp.append(f"{k}:{v}")

                    data[key] = "；".join(temp)
                else:
                    data[key] = str(data[key])

            return data
        
    except Exception:
        data = {key: "解析失敗" for key in extract_keys}
        data["抽取依據"] = text

        return data


def extract_with_gemini(source_text):
    source_text = clean_value(source_text)

    if source_text == "":
        data = {key: "未能了解" for key in extract_keys}
        data["抽取依據"] = "第七層內文為空"
        return data

    prompt = build_prompt(source_text)

    try:
        response = model.generate_content(
            prompt,
            generation_config={
                "response_mime_type": "application/json"
            }
        )

        return parse_json_response(response.text)

    except Exception as e:
        data = {key: "API錯誤" for key in extract_keys}
        data["抽取依據"] = str(e)
        return data


# =========================
# 4. 處理整個資料夾
# =========================

all_results = []

excel_files = list(INPUT_FOLDER.glob("*.xlsx"))

print("找到 Excel 檔案數：", len(excel_files))

for file_path in excel_files:

    if file_path.name.startswith("~$"):
        continue

    print("=" * 80)
    print("正在處理：", file_path.name)

    df = pd.read_excel(file_path)

    file_results = []

    for idx, row in df.iterrows():

        pdf_name = clean_value(row.get("pdf檔名", ""))
        step = clean_value(row.get("第四層", ""))
        action_plan = clean_value(row.get("第五層", ""))
        industry = clean_value(row.get("第六層", ""))
        source_text = clean_value(row.get("第七層", ""))

        match = re.search(r"\d+", pdf_name)
        if match:
            year_id = match.group(0)
        else:
            year_id = "unknown"
        chunk_id = f"{year_id}_{idx + 1:04d}"

        print(f"正在抽取第 {idx + 1} 筆 / 共 {len(df)} 筆")

        ai_result = extract_with_gemini(source_text)

        output_row = {
            "chunk_id": chunk_id,
            "pdf檔名": pdf_name,
            "第五層": action_plan,
            "source_text": source_text
        }

        output_row.update(ai_result)

        file_results.append(output_row)
        all_results.append(output_row)

        time.sleep(1)

    file_df = pd.DataFrame(file_results)

    output_path = OUTPUT_FOLDER / f"{file_path.stem}_亮點成果抽取.xlsx"
    file_df.to_excel(output_path, index=False)

    print("已輸出：", output_path)


# =========================
# 5. 輸出總表
# =========================

all_df = pd.DataFrame(all_results)

summary_path = OUTPUT_FOLDER / "亮點成果抽取總表.xlsx"
all_df.to_excel(summary_path, index=False)

print("=" * 80)
print("全部完成！")
print("每份檔案結果存在：", OUTPUT_FOLDER)
print("總表存在：", summary_path)



# %%
# =========================
# 格式二 整理表格
# =========================

import pandas as pd
import re
from pathlib import Path

INPUT_FOLDER2 = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/亮點成果表格/格式二")
OUTPUT_FOLDER2 = INPUT_FOLDER2 / "產業欄位整理_格式二"
OUTPUT_FOLDER2.mkdir(parents=True, exist_ok=True)

# =========================
# 2. 工具函數
# =========================

def clean_value(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    if x.lower() == "nan":
        return ""
    return x


def detect_industry(text):
    """
    偵測 xxx業 / xxx產業
    """

    text = clean_value(text)

    match = re.search(r"[\u4e00-\u9fffA-Za-z0-9]+(?:產業|業)", text)

    if match:
        return match.group(0)

    return ""

def remove_industry(text):
    text = clean_value(text)

    # 刪掉 xxx業 / xxx產業
    text = re.sub(
        r"[\u4e00-\u9fffA-Za-z0-9]+(?:產業|業)",
        "",
        text
    )

    # 刪掉殘留編號，例如 (1)、（1）、1.、5.
    text = re.sub(r"[（(]\d+[）)]", "", text)
    text = re.sub(r"^\s*\d+\.\s*$", "", text)
    text = re.sub(r"^\s*\d+\.\s*", "", text)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# =========================
# 3. 處理整個資料夾
# =========================

excel_files = list(INPUT_FOLDER2.glob("*.xlsx"))

print("找到 Excel 檔案數：", len(excel_files))

for file_path in excel_files:

    if file_path.name.startswith("~$"):
        continue

    print("=" * 80)
    print("正在處理：", file_path.name)

    df = pd.read_excel(file_path)

    industries = []

    # 逐列處理
    for idx, row in df.iterrows():

        level6 = clean_value(row.get("第六層", ""))
        level7 = clean_value(row.get("第七層", ""))

        industry = ""

        # 先從第六層找
        industry = detect_industry(level6)

        # 找不到再從第七層找
        if industry == "":
            industry = detect_industry(level7)

        industries.append(industry)

        # 從原本欄位刪掉產業
        df.at[idx, "第六層"] = remove_industry(level6)
        df.at[idx, "第七層"] = remove_industry(level7)

    # 新增 industry 欄位
    df["industry"] = industries

    # 把 industry 放在第七層與第八層中間
    cols = list(df.columns)

    cols.remove("industry")

    if "第八層" in cols:
        insert_pos = cols.index("第八層")
        cols.insert(insert_pos, "industry")
    else:
        cols.append("industry")

    df = df[cols]

    # 改欄位名稱
    df = df.rename(columns={
        "第五層": "政策推動面向",
        "第六層": "技術措施類型",
        "第七層": "技術措施",
        "第八層": "內文"
    })

    # 輸出
    output_path2 = OUTPUT_FOLDER2 / f"{file_path.stem}_產業欄位整理.xlsx"

    df.to_excel(output_path2, index=False)

    print("已輸出：", output_path2)

print("=" * 80)
print("全部完成！")
print("結果存在：", OUTPUT_FOLDER2)
# %%

# %%
#ai應用
# =========================
# 格式二
# =========================

import pandas as pd
import json
import re
import time
from pathlib import Path

from google.oauth2 import service_account
import vertexai
from vertexai.generative_models import GenerativeModel


# =========================
# 1. 基本設定
# =========================

KEY_PATH = "C:/Users/0128M\OneDrive/桌面/智庫/api/internship-495506-1c354f4e33b5.json"
credentials = service_account.Credentials.from_service_account_file(KEY_PATH)
PROJECT_ID = 'internship-495506'#"你的專案ID" 
LOCATION = "us-central1"

INPUT_FOLDER3 = Path(r"C:/Users/0128M/OneDrive/桌面/智庫/week4/亮點成果表格/格式二/產業欄位整理_格式二")
OUTPUT_FOLDER3 = INPUT_FOLDER3 / "G亮點分析結果_格式二"
OUTPUT_FOLDER3.mkdir(parents=True, exist_ok=True)

credentials = service_account.Credentials.from_service_account_file(KEY_PATH)

vertexai.init(
    project=PROJECT_ID,
    location=LOCATION,
    credentials=credentials
)


model = GenerativeModel("gemini-2.5-flash")


# =========================
# 2. 輸出欄位
# =========================

extract_keys = [
    "既有問題",
    "製程改善策略",
    "製程改善策略的執行方法",
    "能源轉換技術",
    "能源轉換技術的執行方法",
    "循環經濟措施",
    "循環經濟措施的執行方法",
    "投資金額",
    "回收年限",
    "關鍵效能參數",
    "計畫成效",
    "減碳量",
    "對應的國際規範連結",
    "抽取依據"
]


# =========================
# 3. 工具函數
# =========================

def clean_value(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    if x.lower() == "nan":
        return ""
    return x

def clean_source_text(text):
    text = clean_value(text)

    # 先抓出所有「內容」欄位，只保留內容，不要名稱
    contents = re.findall(
        r"[\"']?內容[\"']?\s*[:：]\s*[\"']([^\"']+)[\"']",
        text
    )

    if contents:
        text = "；".join(contents)
    else:
        # 如果不是字典格式，才用一般清理
        text = re.sub(r"[\"']?問題名稱[\"']?\s*[:：]", "", text)
        text = re.sub(r"[\"']?問題內容[\"']?\s*[:：]", "", text)
        text = re.sub(r"[\"']?名稱[\"']?\s*[:：]\s*[^,，;；}]+[,，;；]?", "", text)
        text = re.sub(r"[\"']?內容[\"']?\s*[:：]", "", text)

    # 清掉殘留符號
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"[\"']", "", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def normalize_bullets(text):
    text = str(text)

    # 大部分模型輸出的 * 或 -
    text = re.sub(r"^\s*[\*\-]\s*", "· ", text, flags=re.MULTILINE)

    # 已經是大黑點時也轉掉
    text = text.replace("•", "·")

    return text.strip()


def build_prompt(source_text):
    return f"""
你是一個資訊抽取系統，不是分析系統。

你的任務是從「原文」中抽取已明確出現的資訊。

請嚴格遵守：
1. 只能抽取原文中明確出現的內容。
2. 不可以推論。
3. 不可以補充背景知識。
4. 不可以改寫成自己的說法。
5. 如果單一個column得到的抽取結果有很多項，請以圓點的條列式進行整理。
6. 如果原文沒有明確提到，請填「未能了解」。
7. 請輸出合法 JSON。
8. 不要加上 ```json。
9. 不要輸出 Markdown。
10.不要加任何解釋文字，只能輸出 JSON。

欄位定義：

既有問題：指產業在進行改善前所面臨的技術瓶頸或能源管理痛點，也就是該產業原本就存在的問題、本次行動方案要解決的目標問題。先寫問題的名稱，再說明問題的內容。
製程改善策略：指達成減碳目標的具體措施之名稱。
製程改善策略的執行方法：指在進行製程改善策略時，為了達成減碳的目標具體執行的方式。
能源轉換技術：指為了調整能源結構，將高碳燃料轉向低碳或零碳能源的技術方法之名稱。
能源轉換技術的執行方法：指在進行能源轉換技術時，為了調整能源結構而具體執行的方式。
循環經濟措施：為了能達成資源的重複利用、延長資源生命週期而實施的措施。
循環經濟措施的執行方法：指在進行循環經濟措施時，為了達成資源的重複利用、延長資源生命週期具體執行的方式。
投資金額：指企業為了實施減碳改善計畫所投入的資金總額，包含設備購置費、政府補助款及企業自籌款。請區分並標註「政府補助款」與「廠商新增投資額（自籌款）」。
回收年限：指該項投資預計多少年可以透過節省的成本回收。
計畫成效：指實施該計畫後的效果與進展，以增加的經濟與資源價值的量化數據表示。包含節能率、投資回收年限、產值增加額或資源回收率。
關鍵效能參數:內文中有提到可以用來評估該策略會用到的單位。
減碳量：指透過改善措施後，每年所減少排放的溫室氣體總量。
對應的國際規範鏈結：指該減碳成果如何協助產業對接國際減碳要求或倡議。
抽取依據：請引用原文中最能支持上述抽取結果的短句。

請輸出以下 JSON：

{{
  "既有問題": "",
  "製程改善策略": "",
  "製程改善策略的執行方法": "",
  "能源轉換技術": "",
  "能源轉換技術的執行方法": "",
  "循環經濟措施": "",
  "循環經濟措施的執行方法": "",
  "投資金額": "",
  "回收年限": "",
  "關鍵效能參數": "",
  "計畫成效": "",
  "減碳量": "",
  "對應的國際規範連結": "",
  "抽取依據": ""
}}

原文如下：

{source_text}
"""


def parse_json_response(text):
    text = text.strip()

    text = re.sub(r"^```json", "", text)
    text = re.sub(r"^```", "", text)
    text = re.sub(r"```$", "", text)
    text = text.strip()

    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        text = match.group(0)

    try:
        data = json.loads(text)

        for key in extract_keys:
            if key not in data:
                data[key] = "未能了解"
            # 強制轉成純文字
            for key in list(data.keys()):

                if isinstance(data[key], list):

                   data[key] = "；".join(
                       [str(x) for x in data[key]]
                   )

                elif isinstance(data[key], dict):
 
                    temp = []

                    for k, v in data[key].items():
                        temp.append(f"{k}:{v}")

                    data[key] = "；".join(temp)
                else:
                    data[key] = str(data[key])
                
                data[key] = normalize_bullets(data[key])
            clean_data = {key: data.get(key, "未能了解") for key in extract_keys}

            return clean_data
        
    except Exception:
        data = {key: "解析失敗" for key in extract_keys}
        data["抽取依據"] = text

        return data


def extract_with_gemini(source_text):
    source_text = clean_value(source_text)

    if source_text == "":
        data = {key: "未能了解" for key in extract_keys}
        data["抽取依據"] = "內文為空"
        return data

    prompt = build_prompt(source_text)

    try:
        response = model.generate_content(
            prompt,
            generation_config={
                "response_mime_type": "application/json"
            }
        )

        return parse_json_response(response.text)

    except Exception as e:
        data = {key: "API錯誤" for key in extract_keys}
        data["抽取依據"] = str(e)
        return data


# =========================
# 4. 處理整個資料夾
# =========================

all_results = []

excel_files = list(INPUT_FOLDER3.glob("*.xlsx"))

print("找到 Excel 檔案數：", len(excel_files))

for file_path in excel_files:

    if file_path.name.startswith("~$"):
        continue

    print("=" * 80)
    print("正在處理：", file_path.name)

    df = pd.read_excel(file_path)

    file_results = []

    for idx, row in df.iterrows():

        pdf_name = clean_value(row.get("pdf檔名", ""))
        policy_axis = clean_value(row.get("政策推動面向", ""))
        technology_type = clean_value(row.get("技術措施類型", ""))
        technology = clean_value(row.get("技術措施", ""))
        industry = clean_value(row.get("industry", ""))
        source_text = clean_source_text(row.get("內文", ""))

        match = re.search(r"\d+", pdf_name)
        if match:
            year_id = match.group(0)
        else:
            year_id = "unknown"
        chunk_id = f"{year_id}_{idx + 1:04d}"

        print(f"正在抽取第 {idx + 1} 筆 / 共 {len(df)} 筆")

        ai_result = extract_with_gemini(source_text)

        output_row = {
            "chunk_id": chunk_id,
            "pdf檔名": pdf_name,
            "政策推動面向": policy_axis,
            "技術措施類型": technology_type,
            "技術措施": technology,
            "industry":industry,
            "內文": source_text
        }

        output_row.update(ai_result)

        file_results.append(output_row)
        all_results.append(output_row)

        time.sleep(1)

    file_df = pd.DataFrame(file_results)

    output_path3 = OUTPUT_FOLDER3 / f"{file_path.stem}_亮點成果抽取.xlsx"
    file_df.to_excel(output_path3, index=False)

    print("已輸出：", output_path3)


# =========================
# 5. 輸出總表
# =========================

all_df = pd.DataFrame(all_results)

summary_path3 = OUTPUT_FOLDER3 / "亮點成果抽取總表.xlsx"
all_df.to_excel(summary_path3, index=False)

print("=" * 80)
print("全部完成！")
print("每份檔案結果存在：", OUTPUT_FOLDER3)
print("總表存在：", summary_path3)
# %%
