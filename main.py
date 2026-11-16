# YOU CAN LIMIT THE SCRAPING FROM excel_handler.py
import time

from excel_handler import (
    read_facebook_urls,
    initialize_output,
    get_processed_urls,
    save_result
)

from scraper import (
    create_browser,
    scrape_facebook_page
)


def main():

    print("=" * 55)
    print("       FACEBOOK BUSINESS EMAIL EXTRACTOR")
    print("=" * 55)

    # Read input URLs
    try:
        urls = read_facebook_urls()

    except Exception as error:
        print(f"Input error: {error}")
        return

    if not urls:
        print("No valid Facebook URLs found in Excel.")
        return

    # Prepare output file
    initialize_output()

    # Skip already processed URLs
    processed_urls = get_processed_urls()

    remaining_urls = [
        url for url in urls
        if url not in processed_urls
    ]

    print(f"\nTotal URLs: {len(urls)}")
    print(f"Already processed: {len(processed_urls)}")
    print(f"Remaining: {len(remaining_urls)}")

    if not remaining_urls:
        print("\nAll URLs have already been processed.")
        return

    # Start browser
    playwright = None
    browser = None
    context = None

    found_count = 0
    not_found_count = 0
    error_count = 0

    try:
        playwright, browser, context, page = create_browser()

        for index, url in enumerate(remaining_urls, start=1):

            print("\n" + "-" * 55)
            print(f"Progress: {index}/{len(remaining_urls)}")
            print(f"URL: {url}")

            email, status = scrape_facebook_page(page, url)

            # Save immediately
            save_result(url, email, status)

            if status == "Found":
                found_count += 1
                print(f"Email: {email}")
                print("Status: FOUND")

            elif status == "Not Found":
                not_found_count += 1
                print("Email: Not available")
                print("Status: NOT FOUND")

            else:
                error_count += 1
                print("Status: ERROR")

            # Be respectful of the website
            if index < len(remaining_urls):
                time.sleep(3)

    except KeyboardInterrupt:
        print("\n\nProcess interrupted by user.")

    except Exception as error:
        print(f"\nUnexpected error: {error}")

    finally:

        if context:
            context.close()

        if browser:
            browser.close()

        if playwright:
            playwright.stop()

        print("\n" + "=" * 55)
        print("PROCESS SUMMARY")
        print("=" * 55)

        print(f"Found:     {found_count}")
        print(f"Not Found: {not_found_count}")
        print(f"Errors:    {error_count}")

        print("\nResults saved in:")
        print("output/extracted_emails.xlsx")

        print("=" * 55)


if __name__ == "__main__":
    main()