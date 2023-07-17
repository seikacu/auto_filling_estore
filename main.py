import tkinter as tk

from tkinter import filedialog

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

from openpyxl.utils.exceptions import InvalidFileException

from openpyxl import load_workbook

from auth_data import user, password 


print("Режимы работы:")
print("1 - Открыть estore.gz и авторизоваться на сайте")
print("2 - Ввести ссылку")
print("3 - Заполнить форму данными из excel (консольный вариант)")
print("4 - Заполнить форму данными из excel (диалоговое окно)")

mode = int(input("Введите номер режима: "))

options = webdriver.ChromeOptions()

def set_driver_options(options:webdriver.ChromeOptions):
    # user-agent
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 YaBrowser/23.5.2.625 Yowser/2.5 Safari/537.36")

    # for ChromeDriver version 79.0.3945.16 or over
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    options.debugger_address = 'localhost:8989'
    
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

    # заполнить set_textarea на сайте
    def set_textarea(driver:webdriver.Chrome, name, arg, j):
        try:
            textarea = driver.find_element(By.XPATH, f"//textarea[@name='offer[offerRows][items][{j}][{arg}]']")
            textarea.clear()
            textarea.send_keys(name)
        except NoSuchElementException:
            print(f"Кнопка {arg} не найдена")
            pass

    # заполнить input на сайте
    def set_input(driver:webdriver.Chrome, name, arg, j):
        try:
            inputVal = driver.find_element(By.XPATH, f"//input[@name='offer[offerRows][items][{j}][{arg}]']")
            inputVal.clear()
            inputVal.send_keys(name)
        except NoSuchElementException:
            print(f"Кнопка {arg} не найдена")
            pass
       
    def get_wb():
        
        fileName = str(input("Введите имя файла excel (без .xlsx): "))
        # Load workbook
        wb = load_workbook(f"./{fileName}.xlsx")
        return wb
        
    def get_sheet(wb:load_workbook):
               
        # Get a sheet by name 
        sheet = wb['Лист1']
        return sheet
    
    # заполнение формы данными из excel (диалоговый режим)
    def process_wb(wb:load_workbook):
        
        sheet = get_sheet(wb)
        fill_form_from_sheet(sheet)
        wb.close()
    
    # открыть файл excel через диалоговое окно
    def open_xlsx_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                wb = load_workbook(file_path)
                print("Файл успешно открыт:", file_path)
                # заполнение формы данными из excel
                process_wb(wb)
            except InvalidFileException:
                print("Ошибка: Неверный формат файла или файл поврежден.")
            except Exception as e:
                print("Произошла ошибка:", e)
    
    # заполнение формы данными из excel
    def fill_form_from_sheet(sheet):
        i = 2
        j = 0
        while True:
            
            # Наименование товара, работы, услуги *
            nameProduct = sheet[f'A{i}'].value 
            if nameProduct == None:
                break
            
            # Уточняющие характеристики
            characteristics = sheet[f'B{i}'].value 
            # Срок гарантии *
            warranty = sheet[f'C{i}'].value 
            # Цена без НДС за единицу (руб.) *
            # PriceWithoutVAT = sheet[f'D{i}'].value 
            # Цена с НДС за единицу (руб.) *
            PriceWithVAT = sheet[f'E{i}'].value 
            # Ставка НДС (%): *
            # VATRate = sheet[f'F{i}'].value
                            
            # заполнить "Уточняющие характеристики" на сайте
            set_textarea(driver, characteristics, "comment", j)
            # заполнить "Уточняющие характеристики" на сайте
            set_textarea(driver, warranty, "warranty", j)
            # заполнить "Цена с НДС за единицу (руб.) *" на сайте
            set_input(driver, str(PriceWithVAT), "offerPriceNdsView", j)
            
            i += 1
            j += 1
    # режим 1
    if mode == 1:
        driver.get("https://account.gz-spb.ru/login")
        authentication(driver)

    # режим 2
    if mode == 2:
        url = str(input("Введите ссылку: "))
        driver.get(url)
        # driver.get("https://estore.gz-spb.ru/electronicshop/offer/create/297894?backurl=L3Byb2NlZHVyZS9mb3JtL3ZpZXcvMjk3ODk0Lz9iYWNrdXJsPUwyVnNaV04wY205dWFXTnphRzl3TDJOaGRHRnNiMmN2Y0hKdlkyVmtkWEpsTDJsdVpHVjRMejlyWlhsM2IzSmtjejBsUkRBbFFrRWxSREFsUWpBbFJEQWxRa1FsUkRFbE9EWW1iMlptWlhKZlpHRjBaVjlsYm1RdFpuSnZiVDB4T0M0d055NHlNREl6Sm1aMWJHeFRaV0Z5WTJnOU1DWT0=")
        
    # режим 3
    if mode == 3:
        wb = get_wb()
        sheet = get_sheet(wb)
        fill_form_from_sheet(sheet)
        wb.close()
        
    # режим 4
    if mode == 4:
        # Создание окна
        root = tk.Tk()
        root.title("Открыть XLSX файл")
        
        # Создание кнопки "Открыть файл"
        open_button = tk.Button(root, text="Открыть XLSX файл", command=open_xlsx_file)
        open_button.pack(pady=20)
        
        # Запуск основного цикла обработки событий
        root.mainloop()
        
    print("Программа завершена")
        
except Exception as ex:
    print(ex)
# finally:
#     driver.close()
#     driver.quit()