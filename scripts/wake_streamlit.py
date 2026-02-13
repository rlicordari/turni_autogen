import os
import time
from urllib.parse import urlparse, urlunparse, urlencode, parse_qsl

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def with_cache_busting(url: str) -> str:
    parts = list(urlparse(url))
    q = dict(parse_qsl(parts[4], keep_blank_values=True))
    q["keepalive"] = str(int(time.time()))
    parts[4] = urlencode(q)
    return urlunparse(parts)


def main() -> int:
    app_url = os.environ.get("STREAMLIT_APP_URL", "").strip()
    if not app_url:
        print("ERROR: STREAMLIT_APP_URL env var is empty.")
        return 2

    app_url = app_url.rstrip("/") + "/"
    url = with_cache_busting(app_url)

    chrome_path = os.environ.get("CHROME_PATH")
    chromedriver_path = os.environ.get("CHROMEDRIVER_PATH")

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--lang=en-US")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome Safari/537.36"
    )

    if chrome_path:
        options.binary_location = chrome_path

    service = Service(executable_path=chromedriver_path) if chromedriver_path else Service()

    driver = webdriver.Chrome(service=service, options=options)

    try:
        print(f"Opening: {url}")
        driver.get(url)

        wait = WebDriverWait(driver, 25)
        # Streamlit sleep button: "Yes, get this app back up!"
        # Robust selector: case-insensitive contains() via translate()
        button_xpath = (
            "//button[contains(translate(normalize-space(.), "
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
            "'get this app back up')]"
        )

        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, button_xpath)))
            print("Sleep wake button found → clicking...")
            btn.click()

            # Confirm something changed: button disappears OR app container appears
            try:
                WebDriverWait(driver, 60).until(
                    EC.any_of(
                        EC.invisibility_of_element_located((By.XPATH, button_xpath)),
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, '[data-testid="stAppViewContainer"]')
                        ),
                    )
                )
                print("Wake click done ✅ (button disappeared or app container appeared)")
            except TimeoutException:
                # Not fatal: Streamlit may keep the button in DOM while loading.
                print("WARNING: Wake click sent, but couldn't confirm within 60s. Keeping as success.")
        except TimeoutException:
            print("No sleep button detected. App likely already awake ✅")

        # Extra "real traffic": touch health endpoint too
        health_url = app_url.rstrip("/") + "/_stcore/health"
        print(f"Also touching health endpoint: {health_url}")
        driver.get(with_cache_busting(health_url))

        print("Done.")
        return 0

    except Exception as e:
        print(f"ERROR: {type(e).__name__}: {e}")
        return 1
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
