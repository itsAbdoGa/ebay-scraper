import re
import time
from dataclasses import dataclass

from playwright.sync_api import Browser, Page, TimeoutError as PlaywrightTimeoutError
from selectolax.parser import HTMLParser

REQUEST_DELAY_SECONDS = 1.5
PAGE_TIMEOUT_MS = 60_000
PRODUCT_SELECTOR_TIMEOUT_MS = 15_000

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

VISIT_STORE_RE = re.compile(r"^Visit the (.+?) Store$")

BLOCK_MARKERS = (
    "opfcaptcha.amazon.com",
    "validateCaptcha",
    "csm-captcha-instrumentation",
    "Click the button below to continue shopping",
    "To discuss automated access to Amazon data",
    "Sorry, we just need to make sure you're not a robot",
    "Enter the characters you see below",
)
PRODUCT_MARKERS = (
    "wayfinding-breadcrumbs_feature_div",
    "productTitle",
    "po-brand",
    "bylineInfo",
)
PRODUCT_WAIT_SELECTOR = "#productTitle, #wayfinding-breadcrumbs_feature_div, tr.po-brand"
MIN_PRODUCT_PAGE_BYTES = 50_000
BODY_PREVIEW_CHARS = 500

AMAZON_DP_URL = "https://www.amazon.com/dp/{asin}"


@dataclass
class PageFetchResult:
    url: str
    final_url: str
    status_code: int
    html: str

    @property
    def content_length(self) -> int:
        return len(self.html.encode("utf-8"))


class AmazonBlockedError(Exception):
    def __init__(
        self,
        url: str,
        status_code: int,
        final_url: str,
        content_length: int,
        reason: str,
        body_preview: str,
    ) -> None:
        self.status_code = status_code
        self.final_url = final_url
        self.content_length = content_length
        self.reason = reason
        self.body_preview = body_preview
        message = (
            f"Amazon blocked the request ({reason}): "
            f"status={status_code}, final_url={final_url}, "
            f"bytes={content_length}"
        )
        super().__init__(message)


def analyze_page(url: str, result: PageFetchResult) -> None:
    html = result.html
    lowered = html.lower()
    content_length = result.content_length

    if result.status_code and result.status_code != 200:
        raise AmazonBlockedError(
            url=url,
            status_code=result.status_code,
            final_url=result.final_url,
            content_length=content_length,
            reason=f"unexpected HTTP status {result.status_code}",
            body_preview=html[:BODY_PREVIEW_CHARS],
        )

    if any(marker.lower() in lowered for marker in BLOCK_MARKERS):
        raise AmazonBlockedError(
            url=url,
            status_code=result.status_code,
            final_url=result.final_url,
            content_length=content_length,
            reason="captcha or bot-detection page",
            body_preview=html[:BODY_PREVIEW_CHARS],
        )

    if content_length < MIN_PRODUCT_PAGE_BYTES and not any(
        marker in html for marker in PRODUCT_MARKERS
    ):
        raise AmazonBlockedError(
            url=url,
            status_code=result.status_code,
            final_url=result.final_url,
            content_length=content_length,
            reason="response too small and missing product-page markers",
            body_preview=html[:BODY_PREVIEW_CHARS],
        )


def print_response_diagnostics(result: PageFetchResult) -> None:
    print(f"  status: {result.status_code}")
    print(f"  final url: {result.final_url}")
    print(f"  bytes: {result.content_length}")


def extract_main_category(tree: HTMLParser) -> str:
    breadcrumbs = tree.css_first("#wayfinding-breadcrumbs_feature_div")
    if breadcrumbs is None:
        return ""

    first_link = breadcrumbs.css_first("a")
    if first_link is None:
        return ""

    return first_link.text(strip=True)


def extract_brand(tree: HTMLParser) -> str:
    brand_row = tree.css_first("tr.po-brand")
    if brand_row is not None:
        value = brand_row.css_first("td.a-span9 span")
        if value is not None:
            brand = value.text(strip=True)
            if brand:
                return brand

    logo = tree.css_first("#brandLogoHiResByline")
    if logo is not None:
        brand = logo.attributes.get("alt", "").strip()
        if brand:
            return brand

    store_link = tree.css_first("#visitStoreDesktopUrl")
    if store_link is not None:
        match = VISIT_STORE_RE.match(store_link.text(strip=True))
        if match:
            return match.group(1).strip()

    byline = tree.css_first("#bylineInfo")
    if byline is not None:
        byline_link = byline.css_first("a")
        if byline_link is not None:
            text = byline_link.text(strip=True)
            match = VISIT_STORE_RE.match(text)
            if match:
                return match.group(1).strip()
            if text:
                return text

    return ""


def extract_product_details(html: str) -> dict[str, str]:
    tree = HTMLParser(html)
    return {
        "main_category": extract_main_category(tree),
        "brand": extract_brand(tree),
    }


def create_browser_context(browser: Browser):
    context = browser.new_context(
        locale="en-US",
        user_agent=USER_AGENT,
        viewport={"width": 1366, "height": 768},
    )
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
    )
    return context


def warm_up_session(page: Page) -> None:
    page.goto(
        "https://www.amazon.com/",
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )
    time.sleep(1.5)


def fetch_product_page(page: Page, url: str) -> PageFetchResult:
    response = page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    try:
        page.wait_for_selector(
            PRODUCT_WAIT_SELECTOR,
            timeout=PRODUCT_SELECTOR_TIMEOUT_MS,
        )
    except PlaywrightTimeoutError:
        pass

    html = page.content()
    status_code = response.status if response is not None else 0
    return PageFetchResult(
        url=url,
        final_url=page.url,
        status_code=status_code,
        html=html,
    )


def scrape_product(
    page: Page,
    url: str,
    *,
    title: str = "",
) -> dict[str, str]:
    try:
        fetch_result = fetch_product_page(page, url)
        analyze_page(url, fetch_result)
        print_response_diagnostics(fetch_result)
        details = extract_product_details(fetch_result.html)
        result = {
            "title": title,
            "url": url,
            "main_category": details["main_category"],
            "brand": details["brand"],
        }
        print(f"  category: {result['main_category'] or '(not found)'}")
        print(f"  brand: {result['brand'] or '(not found)'}")
        return result
    except AmazonBlockedError as error:
        result = {
            "title": title,
            "url": url,
            "main_category": "",
            "brand": "",
            "error": str(error),
            "status_code": error.status_code,
            "final_url": error.final_url,
            "response_bytes": error.content_length,
            "block_reason": error.reason,
        }
        print(f"  status: {error.status_code}")
        print(f"  final url: {error.final_url}")
        print(f"  bytes: {error.content_length}")
        print(f"  blocked: {error.reason}")
        print(f"  body preview: {error.body_preview!r}")
        return result
    except PlaywrightTimeoutError as error:
        result = {
            "title": title,
            "url": url,
            "main_category": "",
            "brand": "",
            "error": str(error),
        }
        print(f"  error: {error}")
        return result
