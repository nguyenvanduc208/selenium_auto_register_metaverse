import openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from time import sleep

def handle_selenium(email, password, code, name, phone):
    driver = webdriver.Chrome('./chromedriver')
    url = 'https://metahome.digital/sign-up'

    print('---- Open url: ', url)
    driver.get("https://metahome.digital/sign-up")
    wait = WebDriverWait(driver, 10)
    try:
        email_field = wait.until(EC.presence_of_element_located((By.ID, '-email')))
        email_field.send_keys(email)
        print("-----  Load email field")
        
        password_field = wait.until(EC.presence_of_element_located((By.ID, '-password')))
        password_field.send_keys(password)
        print("-----  Load password field")
        
        confirm_password_field = wait.until(EC.presence_of_element_located((By.ID, '-confirmPassword')))
        confirm_password_field.send_keys(password)
        print("-----  Load confirm password field")
        
        code_field = wait.until(EC.presence_of_element_located((By.ID, '-referralCode')))
        code_field.send_keys(code)
        print("-----  Load code field")
        
        check_box = wait.until(EC.presence_of_element_located((By.ID, 'fullAgreement')))
        check_box.click()
        print("-----  Load check box field")
        
        submit_button = wait.until(EC.presence_of_element_located((By.XPATH,
                                                                "//button[@type='submit' and contains(@class, 'mt-3') and contains(@class, 'btn_signup')]")))
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(2)
        print("-----  Load check button submit")
        submit_button.click()
        
        print("\n=============================================================================\n")
        
        sleep(2)
        print("=========== Login ===========")
        
        email_field = wait.until(EC.presence_of_element_located((By.ID, '-email')))
        email_field.send_keys(email)
        print("------- Load email field -------")
        
        password_field = wait.until(EC.presence_of_element_located((By.ID, '-password')))
        password_field.send_keys(password)
        print("------- Load password field -------")
        sleep(1)
        submit_button = wait.until(EC.presence_of_element_located((By.XPATH,
                                                                "//button[@type='submit' and contains(@class, 'mt-3')]")))
        submit_button.click()
        sleep(2)
        print("------- Get my profile: ",name, ' -------')
        driver.get("https://metahome.digital/mypage")
        
        add_profile_element = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='button' and @class='btn_add' and text()='등록하기']")))
        add_profile_element.click()
        sleep(1.5)
        
        name_field = wait.until(EC.presence_of_element_located((By.ID, '-name')))
        name_field.send_keys(name)
        print("------- Send text Name: ", name, ' -------')
        
        phone_field = wait.until(EC.presence_of_element_located((By.ID, '-phoneNumber')))
        phone_field.send_keys(phone)
        print("------- Send text Phone: ", phone, ' -------')
        submit_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='확인']")))
        submit_element.click()
        sleep(2)
        driver.close()
        return True
        

    except TimeoutException as e:
        print("------- ERROR -------")
        print(str(e))
        driver.close()
        return False
    except Exception as e:
        print("------- ERROR -------")
        print(str(e))
        driver.close()
        return False



path_excel = "./data/60tk.xlsx"
try:
    print("Load data excel: ", path_excel)
    workbook = openpyxl.load_workbook(path_excel)
    sheet = workbook.active

    error = True
    while error:
        try:
            start_id = int(input("Start id: "))
            error = False
        except Exception:
            print("Gia tri khong hop le. Nhap lai")
            error = True
            
    error = True
    while error:
        try:
            end_id = int(input("End id: "))
            error = False
        except Exception:
            print("Gia tri khong hop le. Nhap lai")
            error = True
    password_data = input("Password: ")
    
    print("======= Start auto submit =======")
    for row in sheet.iter_rows(values_only=True):
        if row[0] >= start_id and row[0] <= end_id:
            email_data = row[3]
            code_data = row[4]
            name_data = row[1]
            phone_data = '84'+str(int(row[2]))
            print(f"======= Data id {row[0]} =======")
            print("++++ Name: ", name_data)
            print("++++ Phone: ", phone_data)
            print("++++ Email: ", email_data)
            print("++++ Password: ", password_data)
            print("++++ Code: ", code_data)
            
            if handle_selenium(email_data, password_data, code_data, name_data, phone_data):
                print(f'============= Xu ly thanh cong ID: {row[0]}==============')
                print('============= Dang cho voi thoi gian 70s ==============')
                sleep(70)
            else:
                print(f'============= Chuan bi xu ly id tiep theo: {int(row[0])+1} ==============')
                continue
        else:
            continue

except Exception as e:
    print("------- ERROR -------")
    print(e)