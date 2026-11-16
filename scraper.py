
import re
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


EMAIL_PATTERN = re.compile(
    r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
)



def extract_intro_text(page):
    """
    Locate the Intro heading and extract text from its
    nearest containing card.
    """

    intro_text = page.locator("body").evaluate("""
        body => {

            const elements = Array.from(
                body.querySelectorAll("*")
            );

            const headings = elements.filter(element => {
                return (
                    element.innerText &&
                    element.innerText.trim() === "Intro"
                );
            });

            for (const heading of headings) {

                let current = heading.parentElement;

                for (let i = 0; i < 10 && current; i++) {

                    const text = current.innerText || "";

                    // Find the smallest suitable Intro card
                    if (
                        text.length >= 20 &&
                        text.length <= 1500
                    ) {

                        const emailPattern =
                            /[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}/i;

                        if (emailPattern.test(text)) {
                            return text;
                        }
                    }

                    current = current.parentElement;
                }

                // Fallback: return a reasonably sized Intro card
                current = heading.parentElement;

                for (let i = 0; i < 10 && current; i++) {

                    const text = current.innerText || "";

                    if (
                        text.length >= 20 &&
                        text.length <= 1000
                    ) {
                        return text;
                    }

                    current = current.parentElement;
                }
            }

            return null;
        }
    """)

    return intro_text


def extract_email(text):
    """Extract unique email addresses from text."""

    if not text:
        return []

    emails = EMAIL_PATTERN.findall(text)

    # Remove duplicates
    return list(dict.fromkeys(
        email.strip(".,;:()[]{}<>")
        for email in emails
    ))



# def debug_page(page):
#     """Save diagnostic information about the loaded Facebook page."""

#     from pathlib import Path

#     output_dir = Path("output")
#     output_dir.mkdir(exist_ok=True)

#     # Save screenshot
#     page.screenshot(
#         path=str(output_dir / "debug_page.png"),
#         full_page=True
#     )

#     # Save visible page text
#     body_text = page.locator("body").inner_text()

#     with open(
#         output_dir / "debug_page.txt",
#         "w",
#         encoding="utf-8"
#     ) as file:
#         file.write(body_text)

    # Print relevant information
    print("\n========== PAGE DEBUG ==========")
    print("URL:", page.url)
    print("Title:", page.title())
    print("\nVISIBLE PAGE TEXT:")
    print(body_text[:6000])
    print("\n================================")



def scrape_facebook_page(page, url):
    """
    Scrape one Facebook page.

    Returns:
        email, status
    """

    try:
        page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=45000
        )

        # Allow visible page content to render
        page.wait_for_timeout(4000)
        #debug_page(page)

        body_text = page.locator("body").inner_text(
            timeout=15000
        )

        lower_text = body_text.lower()

        # Detect common login/access restrictions
        restricted_phrases = [
            "log in to facebook",
            "log into facebook",
            "you must log in",
            "content isn't available",
            "this content isn't available",
            "page isn't available"
        ]

        if any(
            phrase in lower_text
            for phrase in restricted_phrases
        ):
            return "", "Error"

        # Inspect only the Intro section
        intro_text = extract_intro_text(page)
        # print("\n--- DEBUG INFORMATION ---")
        # print("Page URL:", page.url)
        # print("Page Title:", page.title())
        # print("Intro Text:", repr(intro_text))
        # print("-------------------------\n")

        if not intro_text:
            print("DEBUG: Intro section was not detected.")
            return "", "Error"

        emails = extract_email(intro_text)

        if emails:
            return "; ".join(emails), "Found"

        return "", "Not Found"

    except PlaywrightTimeoutError:
        return "", "Error"

    except Exception as error:
        print(f"Scraping error: {error}")
        return "", "Error"


def create_browser():
    """
    Launch Chromium for public page access.
    No login or access-control bypass is performed.
    """

    playwright = sync_playwright().start()

    browser = playwright.chromium.launch(
        headless=True
    )

    context = browser.new_context(
        viewport={
            "width": 1366,
            "height": 768
        },
        locale="en-US"
    )

    page = context.new_page()

    return playwright, browser, context, page