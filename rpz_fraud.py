import os
import re
import zipfile
import validators
import logging
import pandas as pd
import tabula
import win32com.client as win32
from datetime import datetime
from PyPDF2 import PdfReader


def generate_csv(pdf_files, folder_path):
    csv_files = []
    for file in pdf_files:
        file_path = os.path.join(folder_path, file)
        try:
            # Extract all tables from all pages
            dfs = tabula.read_pdf(file_path, pages='all', multiple_tables=True, lattice=True)
        except Exception as e:
            print(f"❌ 無法解析 {file}：{e}")
            continue

        if not dfs or all(df.empty for df in dfs):
            print(f"❌ {file} 中沒有有效表格，未產生 CSV。")
            continue

        # Standardize columns using first non-empty dataframe
        standard_columns = None
        all_pages = []

        for i, df in enumerate(dfs):
            if df.empty:
                print(f"⚠️ {file} 表格 {i+1} 為空，略過")
                continue

            if standard_columns is None:
                standard_columns = df.columns
            df.columns = standard_columns  # Force column alignment
            df = df.reindex(columns=standard_columns)
            all_pages.append(df)

        if all_pages:
            output_df = pd.concat(all_pages, ignore_index=True)
            output_name = os.path.splitext(file)[0] + ".csv"
            output_path = os.path.join(folder_path, output_name)
            output_df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(f"✅ 已儲存為：{output_name}（共 {len(output_df)} 筆）")
            csv_files.append(output_name)

    return csv_files

def load_iana_tlds_from_csv(csv_path="C:\\Users\\user\\Desktop\\RPZ\\RPZ_auto\\TLDs.csv"):
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"❌ 無法找到 TLD 對照表檔案：{csv_path}")
    df = pd.read_csv(csv_path)
    return set(tld.strip().lower() for tld in df["TLD"].dropna())


def safe_read_csv(file_path):
    for enc in ['utf-8-sig', 'utf-8', 'cp950', 'big5']:
        try:
            return pd.read_csv(file_path, encoding=enc)
        except Exception:
            continue
    raise Exception(f"❌ 無法讀取 CSV：{file_path}，請檢查編碼")


def process_folder_fraud(folder_path, folder_name, base_folder, outlook):
    print(f"→ 開始處理資料夾：{folder_name}")
    logging.info(f"[→] 開始處理資料夾：{folder_name}")
    today = datetime.today()
    roc_date = f"{today.year - 1911}{today.strftime('%m%d')}"
    today_str = f"{today.year}/{today.month}/{today.day}"

    serial_number = folder_name.split()[0]
    document_file = [f for f in os.listdir(folder_path) if f.lower().endswith('號.pdf')]
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf') and f != document_file[0]]
    # csv_files = generate_csv(pdf_files, folder_path)
    csv_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.csv')]

    issue_messages = []

    if not document_file or not csv_files:
        raise Exception("該資料夾缺少公文或附件")

    document_file_path = os.path.join(folder_path, document_file[0])
    reader = PdfReader(document_file_path)
    text = "\n".join(page.extract_text() for page in reader.pages if page.extract_text())

    match_docnum = re.search(r'刑詐防字第\d+號', text.replace(" ", ""))
    if not match_docnum:
        match_docnum = re.search(r'調資肆字第\d+號', text.replace(" ", ""))

    if not match_docnum:
        raise Exception("公文無法擷取文號")
    doc_number = match_docnum.group(0)


    csv_folder = os.path.join(base_folder, "自動化CSV檔")
    os.makedirs(csv_folder, exist_ok=True)
    output_name = f"RPZ 1.0 {roc_date}-{serial_number}.csv"
    output_path = os.path.join(csv_folder, output_name)

    whitelist_path = os.path.join("白名單.csv")
    if not os.path.exists(whitelist_path):
        issue_messages.append("找不到白名單.csv")
        logging.error("❌ 找不到白名單.csv")

    whitelist_df = pd.read_csv(whitelist_path)
    sensitive_keywords = set(str(k).strip().lower() for k in whitelist_df.iloc[:, 0].dropna())

    illegal_format = []
    whitelist_match = []
    valid_tlds = load_iana_tlds_from_csv()
    illegal_TLD = []

    data_rows = []
    for file in csv_files:
        file_path = os.path.join(folder_path, file)
        df = safe_read_csv(file_path)
        columns = [col.strip().lower() for col in df.columns]
        acceptable_column_names = ['domain name', 'domainname', 'Domain Name', 'DomainName', 
                           'Domain name', 'Domainname', 'Domain', 'domain']
        domain_col = [df.columns[i] for i, c in enumerate(columns) if c in acceptable_column_names][0]

        for idx, row in df.iterrows():
            csv_serial_number = str(row['編號']).strip() if '編號' in row else str(idx + 1)
            domain = str(row[domain_col]).strip().lower()
            website_type = str(row['網站性質']).strip() if '網站性質' in row else "未分類"
            email = str(row['承辦人email']).strip().lower() if '承辦人email' in row else "未提供"
            law_refer = str(row['法律依據']).strip().lower() if '法律依據' in row else "未提供"
            unit = str(row['聲請單位']).strip().lower() if '聲請單位' in row else "未提供"
            channel = str(row['申訴管道']).strip().lower() if '申訴管道' in row else "未提供"

            if domain in sensitive_keywords:
                whitelist_match.append(domain)
            if "http://" in domain or "/" in domain or "." not in domain or ".." in domain or domain.startswith("-") or domain.endswith("-") or len(domain) < 3:
                illegal_format.append(domain)

            tld = domain.split('.')[-1]
            if tld not in valid_tlds:
                illegal_TLD.append(domain)

            row_data = {
                "編號": csv_serial_number,
                "domain": domain,
                "網站性質": website_type
            }
            if len(data_rows) == 0:
                row_data.update({
                    "承辦人email": email,
                    "法律依據": law_refer,
                    "聲請單位": unit,
                    "申訴管道": channel,
                    "文號": doc_number,
                    "收文日期": today_str,
                    "類型": "行政機關命令"
                })
            data_rows.append(row_data)


    df_out = pd.DataFrame(data_rows)
    df_out.to_csv(output_path, index=False, encoding='utf-8-sig')

    zip_path = os.path.join(folder_path, f"{csv_files[0].split('.')[0]}.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for f in document_file + pdf_files:
            zipf.write(os.path.join(folder_path, f), arcname=f)

    # mail = outlook.CreateItem(0)
    # mail.To = "rpz10@daar.twnic.tw; kathysung@twnic.tw; chung@twnic.tw"
    # mail.Subject = folder_name
    # mail.Body = ""
    # mail.Attachments.Add(zip_path)
    # mail.Attachments.Add(output_path)

    # session = outlook.GetNamespace("MAPI")
    # account = session.Accounts.Item(1)

    # mail.SendUsingAccount = account 
    # mail.SaveSentMessageFolder = account.DeliveryStore.GetDefaultFolder(5)  

    # mail.Send()

