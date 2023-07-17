import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from auth_data import user, password


print("Режимы работы:")
print("1 - Открыть estore.gz и авторизоваться на сайте")
print("2 - Ввести ссылку")
print("3 - Заполнить данные")

mode = int(input("Введите номер режима: "))

options = webdriver.ChromeOptions()

def set_driver_options(options:webdriver.ChromeOptions):
    # user-agent
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 YaBrowser/23.5.2.625 Yowser/2.5 Safari/537.36")

    # for ChromeDriver version 79.0.3945.16 or over
    options.add_argument("--disable-blink-features=AutomationControlled")

    # options.add_argument('--ignore-certificate-errors')
    # options.add_argument('--ignore-ssl-errors')
    
    # options.debugger_address = 'localhost:8989'
    
    set_driver_options(options)

caps = DesiredCapabilities().CHROME
caps['pageLoadStrategy'] = 'eager'

service = Service(desired_capabilities=caps, executable_path=r"C:\WebDriver\chromedriver\chromedriver.exe")
driver = webdriver.Chrome(service=service, options=options)

try:
    
        # Passing authentication...
    def authentication(driver:webdriver.Chrome):
        try:
            email_input = driver.find_element(By.XPATH, "//input[@name='login[username]']")
            email_input.clear()
            email_input.send_keys(user)

            password_input = driver.find_element(By.ID, "login-password")
            password_input.clear()
            password_input.send_keys(password)
            
            password_input.send_keys(Keys.ENTER)
        except Exception:
            print("Поля аутентификации не найдены или Вы уже авторизованы")
            pass



    if mode == 1:
        driver.get("https://account.gz-spb.ru/login")
        authentication(driver)

    if mode == 2:
        # url = str(input("Введите ссылку: "))
        # driver.get(url)
        driver.get("https://estore.gz-spb.ru/electronicshop/offer/create/297894?backurl=L3Byb2NlZHVyZS9mb3JtL3ZpZXcvMjk3ODk0Lz9iYWNrdXJsPUwyVnNaV04wY205dWFXTnphRzl3TDJOaGRHRnNiMmN2Y0hKdlkyVmtkWEpsTDJsdVpHVjRMejlyWlhsM2IzSmtjejBsUkRBbFFrRWxSREFsUWpBbFJEQWxRa1FsUkRFbE9EWW1iMlptWlhKZlpHRjBaVjlsYm1RdFpuSnZiVDB4T0M0d055NHlNREl6Sm1aMWJHeFRaV0Z5WTJnOU1DWT0=")
        
    if mode == 3:
        pass

    
except Exception as ex:
    print(ex)
# finally:
#     driver.close()
#     driver.quit()