import os
import re
import time
import pdfplumber
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.utils import get_column_letter

# ---------- Configuration ----------
PDF_FOLDER     = r"C:\Users\310221754\OneDrive - Philips\HA Work Order"
EXCEL_PATH     = r"C:\Users\310221754\Desktop\py_wo\wo_excel\CM_WorkOrder.xlsx"
PROCESSED_LOG  = r"C:\Users\310221754\Desktop\py_wo\processed.txt"
PROCESSING_LOG = r"C:\Users\310221754\Desktop\py_wo\processing.txt"

CHECK_INTERVAL = 5 * 60   # 5 minutes
DEBUG          = True     # Set False to silence preview output

# ---------- Queried time (set once at program start in main()) ----------
QUERIED_TIME = None  # type: datetime | None


# ---------- PDF text extraction ----------
def extract_pdf_text(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


# ---------- Column G: Printed Date (yyyy/mm/dd hh:mm:ss, space optional) ----------
PRINTED_DATE_RE = r"(\d{4})/(\d{1,2})/(\d{1,2})\s*(\d{1,2}):(\d{2}):(\d{2})"

def extract_printed_date(text):
    """
    Extract Printed Date in format yyyy/mm/dd hh:mm:ss.
    Handles the PDF quirk where spaces are stripped, e.g.
        'PrintedDate:2026/02/1010:25:46'
    Only searches near the 'Printed Date' label.
    Returns a datetime object, or None if not found.
    """
    lines = text.splitlines()
    match = None

    for i, line in enumerate(lines):
        if re.search(r"Print(?:ed)?\s*Date", line, re.IGNORECASE):
            m = re.search(PRINTED_DATE_RE, line)
            if m:
                match = m
                break
            for j in range(i + 1, min(i + 4, len(lines))):
                m = re.search(PRINTED_DATE_RE, lines[j])
                if m:
                    match = m
                    break
            if match:
                break

    if not match:
        m = re.search(
            r"Print(?:ed)?\s*Date[\s\S]{0,80}?" + PRINTED_DATE_RE,
            text, re.IGNORECASE,
        )
        if m:
            match = m

    if not match:
        return None

    try:
        y, mo, d, h, mi, s = map(int, match.groups())
        return datetime(y, mo, d, h, mi, s)
    except (ValueError, TypeError):
        return None


# ---------- Column B: Location ----------
def extract_location(text):
    """
    Extract location code. Tries multiple patterns:
      1. After a 'Location' label  -> capture the code on same/next line
      2. A PMH-xxxxx-xxxxx style code anywhere
      3. A generic AAA-NN-... code near the top of the document
    """
    m = re.search(
        r"Location[^\n:：]*[:：]\s*([A-Za-z0-9][A-Za-z0-9\-/ ]*?)(?:\s{2,}|\n|$)",
        text, re.IGNORECASE,
    )
    if m:
        val = m.group(1).strip()
        if val:
            return val

    m = re.search(r"(PMH[-\s][A-Z0-9\-/]+)", text, re.IGNORECASE)
    if m:
        return m.group(1).replace(" ", "-").strip()

    m = re.search(
        r"\b([A-Z]{2,5}-[A-Z0-9]{1,6}-[A-Z0-9]{1,6}(?:-[A-Z0-9]{1,6})*)\b",
        text,
    )
    if m:
        return m.group(1).strip()

    return ""


# ---------- Column H: Description (NOT Standard Description) ----------
def extract_description(text):
    m = re.search(r"詳細內容\s*[:：]\s*(.+)", text)
    if m:
        val = m.group(1).strip()
        val = re.split(r"\s{2,}|Remarks|備註|CreationDate|登記日期", val)[0].strip()
        return val if val else None

    m = re.search(
        r"(?<!Standard\s)(?<!Standard)Description\s*[:：]\s*(.+)",
        text,
    )
    if m:
        val = m.group(1).strip()
        val = re.split(r"\s{2,}|Remarks|備註", val)[0].strip()
        return val if val else None

    return None


# ---------- Column J: Maint Contract No. (digits + dashes only) ----------
def extract_maint_contract_no(text):
    m = re.search(
        r"Maint\s*Contract\s*No\.?\s*[^\n:：]*[:：]\s*(\S+)",
        text, re.IGNORECASE,
    )
    if not m:
        return None

    val = m.group(1).strip()

    if re.fullmatch(r"[\d\-]+", val) and re.search(r"\d", val):
        return val

    print(f"    [WARN] maint_contract_no '{val}' is not numeric — writing blank")
    return None


# ---------- Field parsing ----------
def parse_fields(text):
    def grab(pattern, flags=re.IGNORECASE | re.MULTILINE):
        m = re.search(pattern, text, flags)
        return m.group(1).strip() if m else ""

    data = {
        "work_order_no": grab(r"Work\s*Order\s*No\.?\s*[^\n:：]*[:：]\s*(\S+)"),
        "location": extract_location(text),
        "ha_contact": grab(
            r"HA\s*Contact\s*Name[^\n:：]*[:：]\s*([^\n]+?)\s*(?:Division|Department|$)"
        ),
        "serial_no": grab(r"Serial\s*Number[^\n:：]*[:：]\s*(\S+)"),
        "model": grab(
            r"Model[^\n:：]*[:：]\s*([^\n]+?)\s*(?:Description|Remarks|$)"
        ),
        "asset_no": grab(r"Asset\s*Number[^\n:：]*[:：]\s*(\S+)"),
        "printed_date": extract_printed_date(text),
        "description": extract_description(text),
        "wo_type": grab(r"Work\s*Order\s*Type[^\n:：]*[:：]\s*(\S+)"),
        "maint_contract_no": extract_maint_contract_no(text),
        "maint_body_type": grab(
            r"Maint\s*Body\s*Type[^\n:：]*[:：]\s*([A-Z][A-Z/&\-]*)"
        ),
    }

    if not data["maint_body_type"]:
        data["maint_body_type"] = grab(
            r"Maint\s*Body\s*Type[^\n:：]*[:：]\s*(\S+)"
        )

    return data


# ---------- Excel writer ----------
def append_to_excel(data):
    wb = load_workbook(EXCEL_PATH)
    ws = wb.active

    next_row = ws.max_row + 1
    if next_row < 2:
        next_row = 2

    ws.cell(row=next_row, column=1,  value=data["work_order_no"])
    ws.cell(row=next_row, column=2,  value=data["location"])
    ws.cell(row=next_row, column=3,  value=data["ha_contact"])
    ws.cell(row=next_row, column=4,  value=data["serial_no"])
    ws.cell(row=next_row, column=5,  value=data["model"])
    ws.cell(row=next_row, column=6,  value=data["asset_no"])

    # Column G: real datetime + display format
    pd_cell = ws.cell(row=next_row, column=7, value=data["printed_date"])
    if isinstance(data["printed_date"], datetime):
        pd_cell.number_format = "yyyy/mm/dd hh:mm:ss"

    ws.cell(row=next_row, column=8,  value=data["description"])
    ws.cell(row=next_row, column=9,  value=data["wo_type"])
    ws.cell(row=next_row, column=10, value=data["maint_contract_no"])
    ws.cell(row=next_row, column=11, value=data["maint_body_type"])

    # --- Auto-fit column widths based on longest cell content ---
    for col_idx in range(1, 12):  # columns A through K
        letter = get_column_letter(col_idx)
        max_len = 0
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
            val = row[0]
            if val is None:
                continue
            if isinstance(val, datetime):
                text_len = len("yyyy/mm/dd hh:mm:ss")  # = 19
            else:
                text_len = len(str(val))
            if text_len > max_len:
                max_len = text_len
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 50)

    wb.save(EXCEL_PATH)
    print(f"    [OK] Row {next_row} written: WO {data['work_order_no']}")


# ---------- Read all Work Order No. already stored in Excel ----------
def load_excel_work_orders():
    if not os.path.exists(EXCEL_PATH):
        return set()
    try:
        wb = load_workbook(EXCEL_PATH, read_only=True)
        ws = wb.active
        wos = set()
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row and row[0]:
                wos.add(str(row[0]).strip())
        return wos
    except Exception as e:
        print(f"    [WARN] Could not read Excel: {e}")
        return set()


# ---------- Processed / Processing logs ----------
def load_processed():
    """
    Returns a set of filenames already processed.
    Accepts both new format ('filename.pdf | 2026/04/27 14:23:05')
    and old format ('filename.pdf') for backward compatibility.
    """
    if not os.path.exists(PROCESSED_LOG):
        return set()
    processed = set()
    with open(PROCESSED_LOG, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            # New format: "filename | timestamp"  -> take the part before '|'
            name = line.split("|", 1)[0].strip()
            if name:
                processed.add(name)
    return processed


def mark_processed(filename):
    """
    Append a line of the form:
        filename.pdf | 2026/04/27 14:23:05
    where the timestamp is the queried time of this program run.
    """
    qt_str = (
        QUERIED_TIME.strftime("%Y/%m/%d %H:%M:%S")
        if isinstance(QUERIED_TIME, datetime)
        else "(unknown)"
    )
    with open(PROCESSED_LOG, "a", encoding="utf-8") as f:
        f.write(f"{filename} | {qt_str}\n")


def write_processing_log(pending_files):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(PROCESSING_LOG, "w", encoding="utf-8") as f:
        f.write(f"# Processing queue – updated {timestamp}\n")
        f.write(f"# Total pending: {len(pending_files)}\n")
        if pending_files:
            for item in pending_files:
                f.write(item + "\n")
        else:
            f.write("# (none – all PDFs processed)\n")


# ---------- File-ready check ----------
def is_file_ready(path, wait=2):
    try:
        size1 = os.path.getsize(path)
        time.sleep(wait)
        size2 = os.path.getsize(path)
        return size1 == size2 and size1 > 0
    except OSError:
        return False


# ---------- Process one PDF ----------
def process_pdf(pdf_path, excel_wos):
    filename = os.path.basename(pdf_path)
    try:
        if not is_file_ready(pdf_path):
            print(f"    [SKIP] {filename} not ready yet, will retry next cycle.")
            return False

        text = extract_pdf_text(pdf_path)

        # if DEBUG:
        #     print("    ---------- Extracted text (preview) ----------")
        #     for line in text.splitlines():
        #         print(f"      {line}")
        #     print("    -----------------------------------------------")

        data = parse_fields(text)

        if DEBUG:
            print("    ---------- Parsed fields ----------")
            for k, v in data.items():
                print(f"      {k:20s}: {v}")
            print("    -----------------------------------")

        wo = (data.get("work_order_no") or "").strip()
        if wo and wo in excel_wos:
            print(f"    [SKIP] WO {wo} already exists in Excel — marking as processed.")
            mark_processed(filename)
            return True

        append_to_excel(data)
        mark_processed(filename)
        if wo:
            excel_wos.add(wo)
        return True
    except Exception as e:
        print(f"    [ERROR] {filename}: {e}\n")
        print("    Please check if the CM_WorkOrder excel is still opened.")
        print("    If yes, please close it and restart the program again to syn the left-behind pdf\n")
        return False


# ---------- Determine pending PDFs ----------
def compute_pending(all_pdfs, processed_set, excel_wos):
    pending = []
    for f in all_pdfs:
        if f in processed_set:
            continue

        full_path = os.path.join(PDF_FOLDER, f)
        wo_in_excel = False
        try:
            if is_file_ready(full_path, wait=0):
                text = extract_pdf_text(full_path)
                m = re.search(
                    r"Work\s*Order\s*No\.?\s*[^\n:：]*[:：]\s*(\S+)",
                    text, re.IGNORECASE)
                wo = m.group(1).strip() if m else ""
                if wo and wo in excel_wos:
                    print(f"    [AUTO] {f} (WO {wo}) already in Excel — marking processed.")
                    mark_processed(f)
                    wo_in_excel = True
        except Exception as e:
            print(f"    [WARN] Could not peek {f}: {e}")

        if not wo_in_excel:
            pending.append(f)
    return pending


# ---------- Process every unprocessed PDF in folder ----------
def process_all_pending():
    processed = load_processed()
    excel_wos = load_excel_work_orders()

    try:
        all_pdfs = sorted(
            f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")
        )
    except FileNotFoundError:
        print(f"    [WARN] PDF folder not found: {PDF_FOLDER}")
        write_processing_log([])
        return

    pending = compute_pending(all_pdfs, processed, excel_wos)
    write_processing_log(pending)

    if not pending:
        print("    No new PDFs to process.")
        return

    print(f"    Found {len(pending)} unprocessed PDF(s). Processing one by one...")
    for f in pending:
        full_path = os.path.join(PDF_FOLDER, f)
        print(f"    -> Processing: {f}")
        process_pdf(full_path, excel_wos)

    processed = load_processed()
    all_pdfs = sorted(
        f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")
    )
    remaining = [f for f in all_pdfs if f not in processed]
    write_processing_log(remaining)


# ---------- Main loop ----------
def main():
    global QUERIED_TIME
    QUERIED_TIME = datetime.now()  # set once at program start

    os.makedirs(PDF_FOLDER, exist_ok=True)
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    for p in (PROCESSED_LOG, PROCESSING_LOG):
        if not os.path.exists(p):
            open(p, "w", encoding="utf-8").close()

    print("=== Work Order PDF → Excel Watcher (5-min polling) ===")
    print(f"Watching folder : {PDF_FOLDER}")
    print(f"Writing to      : {EXCEL_PATH}")
    print(f"Processed log   : {PROCESSED_LOG}")
    print(f"Processing log  : {PROCESSING_LOG}")
    #print(f"Check interval  : every {CHECK_INTERVAL // 60} minute(s)")
    print(f"Queried time    : {QUERIED_TIME.strftime('%Y/%m/%d %H:%M:%S')}\n")

    try:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{now}] Scanning folder...")
        process_all_pending()
        print(f"    You may restart the program if new wo PDFs are added to the folder wo_pdf.\n")

    except KeyboardInterrupt:
        print("\nStopped by user.")


if __name__ == "__main__":
    main()