from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time

options = Options()

driver = webdriver.Chrome(options=options)

driver.get("https://www.google.com")

print("Page title:", driver.title)

time.sleep(5)

driver.quit()