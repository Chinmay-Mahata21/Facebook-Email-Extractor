
from openpyxl import load_workbook, Workbook
from pathlib import Path


INPUT_FILE = Path("input/facebook_urls.xlsx")
OUTPUT_FILE = Path("output/extracted_emails.xlsx")


def read_facebook_urls(input_file=INPUT_FILE, max_urls=None):
    """
    Read Facebook URLs from Excel.

    max_urls=None means process all URLs.
    """

    input_file = Path(input_file)

    if not input_file.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_file}"
        )

    workbook = load_workbook(input_file, read_only=True)
    sheet = workbook.active

    try:
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
            column_index = 1
            start_row = 1

        urls = []

        for row in range(start_row, sheet.max_row + 1):
            value = sheet.cell(row, column_index).value

            if value:
                url = str(value).strip()

                if url.startswith(("https://", "http://")):
                    urls.append(url)

    finally:
        workbook.close()

    # Remove duplicates while preserving order
    urls = list(dict.fromkeys(urls))

    # Apply optional URL limit
    if max_urls is not None:
        if max_urls < 1:
            raise ValueError("URL limit must be at least 1.")

        urls = urls[:max_urls]

    return urls


def initialize_output(output_file=OUTPUT_FILE):
    """Create output Excel file with required columns."""

    output_file = Path(output_file)

    output_file.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    if output_file.exists():
        workbook = load_workbook(output_file)
        sheet = workbook.active

        if sheet.max_row == 1 and sheet["A1"].value is None:
            sheet.append([
                "Facebook URL",
                "Email",
                "Status"
            ])

    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Results"

        sheet.append([
            "Facebook URL",
            "Email",
            "Status"
        ])

    workbook.save(output_file)
    workbook.close()


def get_processed_urls(output_file=OUTPUT_FILE):
    """Get URLs already processed."""

    output_file = Path(output_file)

    if not output_file.exists():
        return set()

    workbook = load_workbook(output_file, read_only=True)
    sheet = workbook.active

    processed = set()

    try:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0]:
                processed.add(row[0])

    finally:
        workbook.close()

    return processed


def save_result(url, email, status, output_file=OUTPUT_FILE):
    """Save one result immediately."""

    output_file = Path(output_file)

    workbook = load_workbook(output_file)
    sheet = workbook.active

    sheet.append([
        url,
        email if email else "",
        status
    ])

    workbook.save(output_file)
    workbook.close()
