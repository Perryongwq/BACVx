import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService

chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors-spki-list')
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_experimental_option('detach',True)

driver = webdriver.Chrome(chrome_options=chrome_options)
# driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()),chrome_options=chrome_options)
driver.get("http://163.50.34.40:8080/prass/jsp/monitor/ProgressDisplayViewer.jsp?dsn=orMesPRASS&Referer=http://163.50.34.40:8080/prass/jsp/shintyoku/index.jsp")
driver.set_window_position(x=0,y=0)
driver.set_window_size(int(1920*4/5),1080)

driver.switch_to.frame(driver.find_element(By .NAME, 'up'))
driver.find_element(By .NAME, 'Lotno').send_keys("22X0282300")
driver.find_element(By .NAME, 'Search').click()
driver.switch_to.default_content()
driver.switch_to.frame(driver.find_element(By .NAME, 'down'))
driver.find_element(By .XPATH,'/html/body/table[2]/tbody/tr[6]/td[3]/font').location_once_scrolled_into_view

time.sleep(100)