import os
import re
import zipfile
import validators
import logging
import pandas as pd
import win32com.client as win32
from datetime import datetime
from PyPDF2 import PdfReader

def is_valid_domain(domain):
    pattern = r'^(?!\-)([a-zA-Z0-9\-]{1,63}(?<!\-)\.)+[a-zA-Z]{2,}$'
    return re.match(pattern, domain) is not None

def load_iana_tlds_from_csv(csv_path="C:\\Users\\user\\Desktop\\RPZ\\RPZ_auto\\TLDs.csv"):
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"❌ 無法找到 TLD 對照表檔案：{csv_path}")
    df = pd.read_csv(csv_path)
    return set(tld.strip().lower() for tld in df["TLD"].dropna())

def process_folder_toxic(folder_path, folder_name, base_folder, outlook):
    print(f"→ 開始處理資料夾：{folder_name}")
    logging.info(f"[→] 開始處理資料夾：{folder_name}")
    today = datetime.today()
    roc_date = f"{today.year - 1911}{today.strftime('%m%d')}"
    today_str = f"{today.year}/{today.month}/{today.day}"

    serial_number = folder_name.split()[0]
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    ods_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.ods')]

    issue_messages = []

    if not pdf_files or not ods_files:
        issue_messages.append("該公文資料夾缺少PDF或ODS")
        logging.error("❌ 該公文資料夾缺少PDF或ODS")

    pdf_path = os.path.join(folder_path, pdf_files[0])
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() for page in reader.pages if page.extract_text())

    match_docnum = re.search(r'警署刑毒緝字第\d+號', text.replace(" ", ""))
    if not match_docnum:
        issue_messages.append("公文無法擷取文號")
        logging.error("❌ 公文無法擷取文號")

    doc_number = match_docnum.group(0)

    whitelist_path = os.path.join("白名單.csv")
    if not os.path.exists(whitelist_path):
        issue_messages.append("找不到白名單.csv")
        logging.error("❌ 找不到白名單.csv")

    whitelist_df = pd.read_csv(whitelist_path)
    sensitive_keywords = set(str(k).strip().lower() for k in whitelist_df.iloc[:, 0].dropna())

    quantity_match = re.search(r'等 \d+', text)
    if not quantity_match:
        issue_messages.append("公文找不到數量")
        logging.error("❌ 公文找不到數量")
        
    pdf_quantity = int(quantity_match.group(0)[1:].replace(" ", ""))

    csv_folder = os.path.join(base_folder, "自動化CSV檔")
    os.makedirs(csv_folder, exist_ok=True)
    output_name = f"RPZ 1.0 {roc_date}-{serial_number}.csv"
    output_path = os.path.join(csv_folder, output_name)

    illegal_format = []
    whitelist_match = []
    valid_tlds = load_iana_tlds_from_csv()
    illegal_TLD = []

    data_rows = []
    for file in ods_files:
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path, engine='odf')
        columns = [col.strip().lower() for col in df.columns]
        acceptable_column_names = ['domain name', 'domainname', 'Domain Name', 'DomainName', 
                           'Domain name', 'Domainname', 'Domain', 'domain']
        domain_col = [df.columns[i] for i, c in enumerate(columns) if c in acceptable_column_names][0]

        for idx, row in df.iterrows():
            domain = str(row[domain_col]).strip().lower()
            email = str(row['警政署承辦人E-mail']).strip().lower() if '警政署承辦人E-mail' in row else "fs993072@cib.npa.gov.tw"

            if len(data_rows) == 0:
                # match_dn = re.search(r"有關[\s:：]*([a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,})", text)
                match_dn = re.search(r"有關[\s:：]*([a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,}([/\w\-]*)?)", text)
                if not match_dn:
                    issue_messages.append("⚠️ 無法從 PDF 中擷取『有關』後的 domain")
                    logging.error("❌ 無法從 PDF 中擷取『有關』後的 domain")
                    # raise Exception("⚠️ 無法從 PDF 中擷取『有關』後的 domain")
                pdf_domain = match_dn.group(1).strip().lower()
                if domain != pdf_domain:
                    issue_messages.append(f"❌ ODS 第一列 domain [{domain}] 與 PDF 中擷取的 [{pdf_domain}] 不一致")
                    logging.error(f"❌ ODS 第一列 domain [{domain}] 與 PDF 中 擷取的 [{pdf_domain}] 不一致")
                    # raise Exception(f"❌ ODS 第一列 domain [{domain}] 與 PDF 中擷取的 [{pdf_domain}] 不一致")

            if domain in sensitive_keywords:
                whitelist_match.append(domain)
            if "http://" in domain or "/" in domain or "." not in domain or ".." in domain or domain.startswith("-") or domain.endswith("-") or len(domain) < 3:
                illegal_format.append(domain)


            tld = domain.split('.')[-1]
            if tld not in valid_tlds:
                illegal_TLD.append(domain)

            row_data = {
                "編號": len(data_rows) + 1,
                "domain": domain,
                "網站性質": "電子商務"
            }
            if len(data_rows) == 0:
                row_data.update({
                    "承辦人email": email,
                    "法律依據": "毒品危害防制條例",
                    "聲請單位": "內政部警政署",
                    "申訴管道": "https://www.npa.gov.tw/ch/mailbox/mailnpa/mailnpa?module=mailnpa&id=7448",
                    "文號": doc_number,
                    "收文日期": today_str,
                    "類型": "行政機關命令"
                })
            data_rows.append(row_data)

    if len(data_rows) != pdf_quantity:
        issue_messages.append("⚠️ domain 數量與 PDF 標示不符")
        logging.error("❌ domain 數量與 PDF 標示不符")
        # raise Exception("domain 數量與 PDF 標示不符")


    df_out = pd.DataFrame(data_rows)
    df_out.to_csv(output_path, index=False, encoding='utf-8-sig')

    if illegal_TLD:
        os.makedirs(os.path.join(base_folder, '疑似不正確的域名TLD'), exist_ok=True)
        illegal_df = pd.DataFrame({"疑似不正確的域名TLD": illegal_TLD})
        illegal_csv_name = f"疑似不正確的域名TLD紀錄_{roc_date}-{serial_number}.csv"
        illegal_csv_path = os.path.join(base_folder, '疑似不正確的域名TLD', illegal_csv_name)
        illegal_df.to_csv(illegal_csv_path, index=False, encoding="utf-8-sig")
        issue_messages.append(f"⚠️ 發現 {len(illegal_TLD)} 筆疑似不正確的域名TLD")

    if illegal_format:
        malformed_df = pd.DataFrame({"格式錯誤domain": illegal_format})
        malformed_csv_name = f"格式錯誤domain_{roc_date}-{serial_number}.csv"
        malformed_csv_path = os.path.join(folder_path, malformed_csv_name)
        malformed_df.to_csv(malformed_csv_path, index=False, encoding="utf-8-sig")
        issue_messages.append(f"⚠️ 發現 {len(illegal_format)} 筆格式錯誤 domain")

    if whitelist_match:
        whitelist_df = pd.DataFrame({"白名單命中domain": whitelist_match})
        whitelist_csv_name = f"白名單命中domain_{roc_date}-{serial_number}.csv"
        whitelist_csv_path = os.path.join(folder_path, whitelist_csv_name)
        whitelist_df.to_csv(whitelist_csv_path, index=False, encoding="utf-8-sig")
        issue_messages.append(f"⚠️ 發現 {len(whitelist_match)} 筆白名單命中 domain")

    # 建立壓縮檔
    zip_path = os.path.join(folder_path, f"{ods_files[0].split('.')[0]}.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for f in pdf_files + ods_files:
            zipf.write(os.path.join(folder_path, f), arcname=f)


    if not issue_messages:
        mail = outlook.CreateItem(0)
        mail.To = "rpz10@daar.twnic.tw; kathysung@twnic.tw; chung@twnic.tw"
        mail.Subject = folder_name
        mail.Body = ""
        mail.Attachments.Add(zip_path)
        mail.Attachments.Add(output_path)

        session = outlook.GetNamespace("MAPI")
        account = session.Accounts.Item(1)

        mail.SendUsingAccount = account
        mail.SaveSentMessageFolder = account.DeliveryStore.GetDefaultFolder(5)
        mail.Send()
    else:
        raise Exception("⚠️ " + "；".join(issue_messages))



