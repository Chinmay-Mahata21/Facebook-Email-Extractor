
from openpyxl import load_workbook, Workbook
from pathlib import Path


INPUT_FILE = Path("input/facebook_urls.xlsx")
OUTPUT_FILE = Path("output/extracted_emails.xlsx")
MAX_URLS = 1                                        #It is for row limitation


def read_facebook_urls():
    """Read Facebook URLs from the input Excel file."""

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}"
        )

    workbook = load_workbook(INPUT_FILE, read_only=True)
    sheet = workbook.active

    # Find the Facebook URL column
    headers = [
        str(cell.value).strip().lower()
        if cell.value else ""
        for cell in sheet[1]
    ]

    if "facebook url" in headers:
        column_index = headers.index("facebook url") + 1
        start_row = 2
    else:
        # If there is no header, use the first column
        column_index = 1
        start_row = 1

    urls = []

    for row in range(start_row, sheet.max_row + 1):
        value = sheet.cell(row, column_index).value

        if value:
            url = str(value).strip()

            if url.startswith(("https://", "http://")):
                urls.append(url)

    workbook.close()

    # Remove duplicate URLs while preserving order
    return list(dict.fromkeys(urls))[:MAX_URLS]


def initialize_output():
    """Create output Excel file with required columns."""

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    if OUTPUT_FILE.exists():
        workbook = load_workbook(OUTPUT_FILE)
        sheet = workbook.active

        if sheet.max_row == 0:
            sheet.append(["Facebook URL", "Email", "Status"])

    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Results"

        sheet.append([
            "Facebook URL",
            "Email",
            "Status"
        ])

    workbook.save(OUTPUT_FILE)
    workbook.close()


def get_processed_urls():
    """Get URLs already processed."""

    if not OUTPUT_FILE.exists():
        return set()

    workbook = load_workbook(OUTPUT_FILE, read_only=True)
    sheet = workbook.active

    processed = set()

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0]:
            processed.add(row[0])

    workbook.close()

    return processed


def save_result(url, email, status):
    """Save one result immediately."""

    workbook = load_workbook(OUTPUT_FILE)
    sheet = workbook.active

    sheet.append([
        url,
        email if email else "",
        status
    ])

    workbook.save(OUTPUT_FILE)
    workbook.close()