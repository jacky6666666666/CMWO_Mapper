import os
import re
import time
import pdfplumber
from openpyxl import load_workbook
from datetime import datetime, date, timedelta
from openpyxl.utils import get_column_letter

# ---------- Configuration ----------
# PDF_FOLDER     = r"C:\Users\310221754\OneDrive - Philips\HA Work Order"
# EXCEL_PATH     = r"C:\Users\310221754\Desktop\py_wo\wo_excel\CM_WorkOrder.xlsx"
# PROCESSED_LOG  = r"C:\Users\310221754\Desktop\py_wo\processed.txt"
# PROCESSING_LOG = r"C:\Users\310221754\Desktop\py_wo\processing.txt"


DELETED_LOG    = r"C:\Users\320163620\Desktop\py_wo\deleted.txt"   # <-- NEW
PDF_FOLDER     = r"C:\Users\320163620\OneDrive - Philips\test1"
EXCEL_PATH     = r"C:\Users\320163620\Desktop\py_wo\wo_excel\CM_WorkOrder.xlsx"
PROCESSED_LOG  = r"C:\Users\320163620\Desktop\py_wo\processed.txt"
PROCESSING_LOG = r"C:\Users\320163620\Desktop\py_wo\processing.txt"


CHECK_INTERVAL = 5 * 60   # 5 minutes
DEBUG          = True     # Set False to silence preview output
DELETE_AFTER_WRITE = True # <-- NEW  master switch for the delete feature

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
LBL_LOC   = r"Location\s*Institution\s*[/\\]\s*Block\s*[/\\]\s*Floor\s*[/\\]\s*Room"
LBL_ASSET = r"Asset\s*Location\s*Description"


def extract_location(text):
    """
    Extract the hospital location code that sits between the labels
    'Location Institution/Block/Floor/Room' and 'Asset Location Description'.
    """

    m = re.search(
        LBL_LOC + r"[\s:：]*(.*?)\s*" + LBL_ASSET,
        text, re.IGNORECASE | re.DOTALL,
    )
    if m:
        val = _clean(m.group(1))
        if _looks_like_code(val):
            return val

    m = re.search(
        LBL_ASSET + r"[\s:：]*(.*?)\s*" + LBL_LOC,
        text, re.IGNORECASE | re.DOTALL,
    )
    if m:
        val = _clean(m.group(1))
        if _looks_like_code(val):
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


def _clean(val: str) -> str:
    val = val.strip()
    val = re.sub(r"\s+", " ", val)
    val = re.sub(r"\s*-\s*", "-", val)
    val = re.sub(r"\s*/\s*", "/", val)
    return val


def _looks_like_code(val: str) -> bool:
    if not val:
        return False
    if len(val) > 60:
        return False
    return bool(re.search(r"[\d\-]", val))


# ---------- Column H: Description ----------
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


# ---------- Column J: Maint Contract No. ----------
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


def cal_oncall_engineer(date_input: datetime):
    oncallEengineer = ["Shelly", "Derick", "Jack"]
    on_call_index = 0
    root_date = datetime.strptime("2026/01/1", "%Y/%m/%d")
    stripped_date_input = date_input.date()
    time_diff = datetime.now() - datetime.strptime("2026/01/1", "%Y/%m/%d")
    num_of_days = int(time_diff.days)
    all_date_dict = {}

    for i in range(0, num_of_days):
        date = (root_date + timedelta(days=i)).date()
        temp_dict = {}
        if (i + 3) % 7 == 0 and on_call_index < len(oncallEengineer):
            on_call_index += 1
        if on_call_index >= len(oncallEengineer):
            on_call_index = 0
        On_call_Engineer = oncallEengineer[on_call_index]
        temp_dict = {date: On_call_Engineer}
        all_date_dict.update(temp_dict)
        temp_dict.clear()

    if all_date_dict.get(stripped_date_input) is not None:
        temp_on_call_engineer = all_date_dict[stripped_date_input]
        all_date_dict.clear()
        return temp_on_call_engineer
    else:
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
        "On_call_Engineer": cal_oncall_engineer(extract_printed_date(text)),
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

    pd_cell = ws.cell(row=next_row, column=7, value=data["printed_date"])
    if isinstance(data["printed_date"], datetime):
        pd_cell.number_format = "yyyy/mm/dd hh:mm:ss"

    ws.cell(row=next_row, column=8,  value=data["description"])
    ws.cell(row=next_row, column=9,  value=data["wo_type"])
    ws.cell(row=next_row, column=10, value=data["maint_contract_no"])
    ws.cell(row=next_row, column=11, value=data["maint_body_type"])
    ws.cell(row=next_row, column=12, value=data["On_call_Engineer"])

    for col_idx in range(1, 13):
        letter = get_column_letter(col_idx)
        max_len = 0
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
            val = row[0]
            if val is None:
                continue
            if isinstance(val, datetime):
                text_len = len("yyyy/mm/dd hh:mm:ss")
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
    if not os.path.exists(PROCESSED_LOG):
        return set()
    processed = set()
    with open(PROCESSED_LOG, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            name = line.split("|", 1)[0].strip()
            if name:
                processed.add(name)
    return processed


def mark_processed(filename):
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


# ---------- NEW: Delete the PDF and log the deletion ----------
def log_deletion(filename: str, status: str, reason: str = ""):
    """
    Append a deletion record to DELETED_LOG.

    Format:
        filename.pdf | 2026/04/27 14:23:05 | OK
        filename.pdf | 2026/04/27 14:23:05 | FAIL | <reason>
    """
    ts = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    line = f"{filename} | {ts} | {status}"
    if reason:
        line += f" | {reason}"
    try:
        with open(DELETED_LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"    [WARN] Could not write to delete log: {e}")


def delete_pdf(pdf_path: str) -> bool:
    """
    Delete the given PDF file and record the outcome in DELETED_LOG.
    Returns True on success, False on failure.
    """
    filename = os.path.basename(pdf_path)
    if not DELETE_AFTER_WRITE:
        return False

    try:
        if not os.path.exists(pdf_path):
            print(f"    [DEL-SKIP] {filename} no longer exists.")
            log_deletion(filename, "SKIP", "file not found")
            return False

        os.remove(pdf_path)
        print(f"    [DEL-OK] {filename} deleted after Excel write.")
        log_deletion(filename, "OK")
        return True
    except PermissionError as e:
        print(f"    [DEL-FAIL] {filename}: file is locked ({e}).")
        log_deletion(filename, "FAIL", f"PermissionError: {e}")
        return False
    except OSError as e:
        print(f"    [DEL-FAIL] {filename}: {e}")
        log_deletion(filename, "FAIL", f"OSError: {e}")
        return False


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
            # Duplicate WO still deletes the PDF so the folder stays clean
            delete_pdf(pdf_path)                                   # <-- NEW
            return True

        append_to_excel(data)
        mark_processed(filename)
        if wo:
            excel_wos.add(wo)

        # Only delete after a confirmed successful Excel write
        delete_pdf(pdf_path)                                       # <-- NEW
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
                    delete_pdf(full_path)                          # <-- NEW
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
    try:
        all_pdfs = sorted(
            f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")
        )
    except FileNotFoundError:
        all_pdfs = []
    remaining = [f for f in all_pdfs if f not in processed]
    write_processing_log(remaining)


# ---------- Main loop ----------
def main():
    global QUERIED_TIME
    QUERIED_TIME = datetime.now()

    os.makedirs(PDF_FOLDER, exist_ok=True)
    os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
    for p in (PROCESSED_LOG, PROCESSING_LOG, DELETED_LOG):          # <-- NEW: DELETED_LOG added
        if not os.path.exists(p):
            open(p, "w", encoding="utf-8").close()

    print("=== Work Order PDF → Excel Watcher (5-min polling) ===")
    print(f"Watching folder : {PDF_FOLDER}")
    print(f"Writing to      : {EXCEL_PATH}")
    print(f"Processed log   : {PROCESSED_LOG}")
    print(f"Processing log  : {PROCESSING_LOG}")
    print(f"Deleted log     : {DELETED_LOG}")                       # <-- NEW
    print(f"Delete after OK : {DELETE_AFTER_WRITE}")                # <-- NEW
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