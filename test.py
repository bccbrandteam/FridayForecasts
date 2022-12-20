#!/usr/bin/python3

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--headless')
options.add_argument('--disable-dev-shm-usage')
options.add_argument("--remote-debugging-port=9222")

try:
  driver = webdriver.Chrome(chrome_options=options)
  driver.get("https://www.google.com")
  s = driver.find_element_by_name("q")
  assert s.is_displayed() is True
  print("ok")
except Exception as ex:
  print(ex)

driver.quit()
