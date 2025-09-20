from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_driver():
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)  # 让浏览器不自动关闭
    service = Service("E:/chromedriver-win64/chromedriver-win64/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver


def do_wait(driver, xpath_str, wait_time=10):
    return WebDriverWait(driver, wait_time).until(
        EC.presence_of_element_located((By.XPATH, xpath_str))
    )


