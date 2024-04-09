from selenium.webdriver import ChromeOptions
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import undetected_chromedriver

def get_driver(cookies_dir=None, is_headless=False):
    """Возвращает обычный селениум драйвер, закоменчен ещё и андетекстед, но с ним не работает headless"""
    chrome_options = ChromeOptions()
    chrome_options.add_argument('--window-size=1920,1080')
    """chrome_options.add_argument('--enable-logging')
    chrome_options.add_argument('--no-sandbox')"""
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument(
        "user-agent=User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable_automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    #mac/linux chromedriver path = ./chromedriver
    driver = undetected_chromedriver.Chrome(driver_executable_path="chromedriver.exe", chrome_options=chrome_options)
    driver.headless = True
    # driver = undetected_chromedriver.Chrome(driver_executable_path=chrome_driver_path, chrome_options=chrome_options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
            '''
    })
    driver.maximize_window()
    return driver