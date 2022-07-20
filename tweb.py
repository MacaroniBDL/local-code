from http.server import executable
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
import time


service = ChromeService(executable_path = ChromeDriverManager().install())
driver = webdriver.Chrome(service = service)

driver.get("https://github.com/nytimes/covid-19-data/raw/master/rolling-averages/us-states.csv")

