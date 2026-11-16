
import re
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# Email validation pattern
EMAIL_PATTERN = re.compile(
    r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
)


def extract_intro_data(page):
    """
    Locate the Intro section and extract its text
    and publicly displayed mailto links.
    """

    intro_data = page.locator("body").evaluate("""
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

                // Find the smallest suitable Intro card
                for (let i = 0; i < 10 && current; i++) {

                    const text = current.innerText || "";

                    const mailtoLinks = Array.from(
                        current.querySelectorAll('a[href^="mailto:"]')
                    );

                    const hasEmail = (
                        /[A-Z0-9._%+-]+@[A-Z0-9.-]+\\.[A-Z]{2,}/i
                        .test(text)
                        || mailtoLinks.length > 0
                    );

                    if (
                        text.length >= 20 &&
                        text.length <= 1500 &&
                        hasEmail
                    ) {
                        return {
                            text: text,
                            mailto: mailtoLinks.map(
                                link => link.getAttribute("href")
                            )
                        };
                    }

                    current = current.parentElement;
                }

                // Fallback when no email is present
                current = heading.parentElement;

                for (let i = 0; i < 10 && current; i++) {

                    const text = current.innerText || "";

                    if (
                        text.length >= 20 &&
                        text.length <= 1000
                    ) {
                        return {
                            text: text,
                            mailto: Array.from(
                                current.querySelectorAll(
                                    'a[href^="mailto:"]'
                                )
                            ).map(link => link.getAttribute("href"))
                        };
                    }

                    current = current.parentElement;
                }
            }

            return null;
        }
    """)

    return intro_data


def extract_email(intro_data):
    """
    Extract unique email addresses from Intro text
    and mailto links.
    """

    if not intro_data:
        return []

    text = intro_data.get("text", "")

    # Extract emails from visible text
    emails = EMAIL_PATTERN.findall(text)

    # Extract emails from mailto links
    for link in intro_data.get("mailto", []):

        if not link:
            continue

        address = link.replace("mailto:", "", 1)
        address = address.split("?")[0]

        emails.extend(EMAIL_PATTERN.findall(address))

    # Clean and remove duplicates (case-insensitive)
    unique_emails = {}

    for email in emails:

        email = email.strip(".,;:()[]{}<>")

        if EMAIL_PATTERN.fullmatch(email):
            unique_emails[email.lower()] = email

    return list(unique_emails.values())


def detect_unavailable_page(body_text):
    """Check for common Facebook unavailable-page messages."""

    unavailable_phrases = [
        "this content isn't available",
        "this page isn't available",
        "the link you followed may be broken",
        "the page may have been removed"
    ]

    lower_text = body_text.lower()

    return any(
        phrase in lower_text
        for phrase in unavailable_phrases
    )



def detect_login_required(page):
    """
    Detect Facebook login redirects and visible login popups.
    """

    try:
        current_url = page.url.lower()

        # Check whether Facebook redirected to a login page
        if "/login" in current_url or "login.php" in current_url:
            return True

        # Check for visible login input fields
        login_inputs = page.locator(
            'input[name="email"], '
            'input[name="pass"], '
            'input[type="password"], '
            'input[autocomplete="username"]'
        )

        for index in range(login_inputs.count()):
            if login_inputs.nth(index).is_visible():
                return True

    except Exception:
        pass

    return False


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

        # Allow page content to render
        page.wait_for_timeout(4000)

        body_text = page.locator("body").inner_text(
            timeout=15000
        )

        # Check whether the page is unavailable
        if detect_unavailable_page(body_text):
            return "", "Page Unavailable"

        # Inspect only the Intro section
        intro_data = extract_intro_data(page)

        # If Intro cannot be accessed, check login requirement
        if not intro_data:

            if detect_login_required(page):
                return "", "Login Required"

            return "", "Error"

        # Extract emails
        emails = extract_email(intro_data)

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
