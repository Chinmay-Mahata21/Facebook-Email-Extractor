# Facebook Business Email Extractor

A Python tool that reads Facebook Page URLs from an Excel file and looks
for publicly visible business email addresses in each page's **Intro**
section.

The first version is intentionally limited to the Intro section. It does
not navigate to the About, Contact, Photos, or other sections.

## Features

-   Reads Facebook Page URLs from an `.xlsx` file.
-   Processes up to **3 unique URLs per run**.
-   Opens pages using Playwright and Chromium.
-   Checks the page's Intro section for email addresses.
-   Saves the Facebook URL, extracted email address, and processing
    status to Excel.
-   Saves each result immediately, so completed results remain available
    if the run is interrupted.
-   Skips URLs that already appear in the output workbook.
-   Removes duplicate URLs from the input list while preserving their
    order.

## Project structure

``` text
Facebook_Email_Extractor/
│
├── main.py
├── scraper.py
├── excel_handler.py
├── requirements.txt
│
├── input/
│   └── facebook_urls.xlsx
│
└── output/
    └── extracted_emails.xlsx
```

The `output` folder and results workbook are created automatically when
the program runs.

## Requirements

-   Python 3.10 or later recommended
-   Windows, macOS, or Linux
-   Internet connection
-   Facebook pages must be accessible without bypassing Facebook's
    access restrictions

## Installation

### 1. Download the project

Open a terminal in the project folder.

### 2. Create and activate a virtual environment (recommended)

**Windows:**

``` bash
python -m venv venv
venv\Scripts\activate
```

**macOS / Linux:**

``` bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install dependencies

``` bash
pip install -r requirements.txt
```

### 4. Install Playwright Chromium

``` bash
playwright install chromium
```

## Prepare the input Excel file

Create this file:

``` text
input/facebook_urls.xlsx
```

Use the following column heading:

``` text
Facebook URL
```

Example:

  Facebook URL
  -----------------------------------
  https://www.facebook.com/example1
  https://www.facebook.com/example2
  https://www.facebook.com/example3

Add one Facebook Page URL per row.

The program looks for the `Facebook URL` heading. If that heading is
absent, it reads URLs from the first column. URLs must begin with
`https://` or `http://`.

### URL limit

The current limit is **3 unique URLs per execution**. The program
selects the first three unique valid URLs from the input list.

To change the limit, edit `MAX_URLS` in `excel_handler.py`:

``` python
MAX_URLS = 3
```

## Run the program

From the project folder, run:

``` bash
python main.py
```

The program processes the selected URLs one at a time and displays
progress and results in the terminal.

## Output

Results are saved to:

``` text
output/extracted_emails.xlsx
```

The workbook contains exactly these columns:

  Facebook URL        Email                     Status
  ------------------- ------------------------- -----------
  Facebook Page URL   Extracted email address   Found
  Facebook Page URL   *(blank)*                 Not Found
  Facebook Page URL   *(blank)*                 Error

### Status meanings

-   **Found** --- One or more email addresses were detected in the
    extracted Intro text.
-   **Not Found** --- The Intro section was identified, but no email
    address was detected in its text.
-   **Error** --- The page could not be processed or the Intro section
    could not be identified.

If multiple email addresses are detected in the Intro text, they are
stored in the Email cell separated by semicolons.

## Resuming and reprocessing

The program checks the output workbook and skips URLs that are already
recorded there. Results are appended as each URL is processed.

This also means that URLs recorded with `Error` are skipped on later
runs. To retry those URLs with the current version, back up the output
workbook and remove the corresponding rows (or remove the output
workbook to start a fresh results file).

**Note about the 3-URL limit:** The limit is applied to the first three
unique URLs before already-processed URLs are skipped. If those three
URLs are already in the output workbook, the program will not
automatically move on to later URLs in the input file.

## Important limitations

-   The tool only checks text extracted from the page's Intro section.
    It does not inspect other Facebook sections or follow external
    website links.
-   It only works with information that Facebook makes publicly
    accessible to the browser session.
-   Facebook may show login prompts, restrict access, or change its page
    layout. As a result, some pages may return `Error` even when an
    email is visible to a logged-in user.
-   Email extraction depends on the Intro section being identified and
    its text being available to Playwright.
-   The current implementation does not log in to Facebook or bypass
    login requirements, CAPTCHAs, or other access controls.
-   A `Not Found` result means no email was detected in the extracted
    Intro text; it does not prove that the business has no email
    address.

Use the tool responsibly and respect Facebook's terms, applicable
privacy rules, and website access restrictions.

## Technologies used

-   **Python** --- application logic
-   **Playwright** --- browser automation
-   **Chromium** --- browser engine
-   **OpenPyXL** --- reading and writing Excel workbooks
-   **Regular expressions (`re`)** --- detecting email addresses

## License

No license has been specified for this project yet.
