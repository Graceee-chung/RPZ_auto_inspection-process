# RPZ_auto_inspection-process
This project automates processing government RPZ (Response Policy Zone) case folders by extracting domains from files, validating formats and top-level domains, checking against whitelists, and generating standardized csv outputs. Then, archives results, zips documents, and integrates with Outlook for automatic email reporting.

## Features
- Parse official documents (PDF/ODS/CSV) for domain information.
- Validate domain formats, top-level domains (TLDs), and whitelist matches.
- Generate unified CSV outputs with metadata (case number, issue date, agency).
- Export error reports for malformed domains and suspicious TLDs.
- Compress files and organize them into result folders (finished/error).
- Send automated email reports with attachments via Outlook.

---

## File Overview

- **`rpz_main.py`**  
  Entry point of the system. Iterates through case folders, determines document type (fraud, toxic, smoke), and routes them to the correct processing function. Handles error catching, logging, and moves folders into `finished` or `error` directories.

- **`rpz_fraud.py`**  
  Handles “fraud” case folders. Extracts data from PDFs and CSVs, validates domains, checks TLDs, generates the final CSV, and packages outputs into ZIP archives. Prepares email attachments for Outlook.

- **`rpz_toxic.py`**  
  Handles “toxic” case folders. Extracts PDF and ODS data, ensures PDF content matches the ODS, validates domains and TLDs, generates output CSV, and creates error logs for mismatches or formatting issues.

- **`rpz_smoke.py`**  
  Handles “smoke” case folders. Similar to toxic handling but with domain-specific defaults (e.g., Ministry of Health). Produces CSV outputs, logs illegal formats, suspicious TLDs, and whitelist hits.

- **`run_rpz.bat`**  
  A Windows batch script for quickly running the automation system from the command line.

- **`TLDs.csv`**  
  Reference list of valid IANA top-level domains. Used to validate whether domains extracted from documents are legitimate.

---

## Requirements
- **Python 3.10+**  
- Libraries: `pandas`, `PyPDF2`, `tabula`, `validators`, `win32com.client` (via pywin32), `odfpy`  
- Microsoft Outlook (for automated mail sending)

