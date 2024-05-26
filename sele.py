from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time

from excel_to_py import main

# Config
login_time = 30  # Time for login (in seconds)
new_msg_time = 5  # TTime for a new message (in seconds)
send_msg_time = 5  # Time for sending a message (in seconds)
country_code = 91  # Set your country code
action_time = 2  # Set time for button click action
# image_path = 'image.png'        # Absolute path to you image

# Create driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))


# Encode Message Tex


# Open browser with default link
link = "https://web.whatsapp.com"
driver.get(link)
driver.maximize_window()

dic = {}
dic = main()  # main worker function with real data in dictionary format
# print(dic)
time.sleep(60)
driver.minimize_window()
wait = WebDriverWait(driver, 100)
# Loop Through Numbers List
print("sending...")
for k, i in dic.items():
    print(k, i)
    link = f"https://web.whatsapp.com/send/?phone=+91{k}&text={i}"
    driver.get(link)
    message = wait.until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span',
            )
        )
    )
    # sendmsg()
    message.click()
    print("sent to " + k)
    # content = driver.switch_to.active_element
    # action = ActionChains(driver)
    # action.send_keys(Keys.ENTER)
    # action.perform()
    # content.send_keys(i)
    # content.send_keys(Keys.ENTER)
    time.sleep(2)
    # message_box ='//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div[2]/div[1]
    # message = wait.until(EC.presence_of_element_located((By.XPATH, message_box)))
    # message.send_key(i + Keys.ENTER)

#    -----------------
time.sleep(100)
