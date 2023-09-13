from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import pythoncom
import time

from download_driver import DownloadDriver

class TrainBooking:
    def __init__(self, start_station, end_station, date, passenger_count, name_id, train_number=None, start_time=None,
                 end_time=None):
        self.start_station = start_station
        self.end_station = end_station
        self.date = date
        self.passenger_count = passenger_count
        self.driver = None
        self.name_id = name_id
        self.train_number = train_number
        self.start_time = start_time
        self.end_time = end_time

    def setup_and_book(self, book_func):
        pythoncom.CoInitialize()
        self.setup_webdriver()
        self.select_start_station()
        self.select_end_station()
        while True:
            self.accept_prompt()
            self.select_date()
            self.select_passenger_count()
            book_func()
            self.submit_form()
            if self.wait_for_confirmation():
                break

    def book_ticket_by_time_slot(self):
        self.setup_and_book(self.by_time_slot)

    def book_ticket_by_train_number(self):
        self.setup_and_book(self.by_train_number)

    def setup_webdriver(self):
        chrome_options = webdriver.ChromeOptions()
        try:
            # 嘗試設置 WebDriver
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_experimental_option("detach", True)  # 將 WebDriver 設置為 "detach" 模式
            service = Service("chromedriver.exe")
            self.driver = webdriver.Chrome(service=service, options=chrome_options)

        except Exception as e:
            print(f"捕獲到異常：{e}")
            version_str = "Current browser version is "
            print("當前的 ChromeDriver 版本與 Chrome 瀏覽器版本不兼容，正在嘗試下載適配的 ChromeDriver 版本...")
            start_index = str(e).find(version_str)
            if start_index != -1:
                start_index += len(version_str)
                end_index = str(e).find(" ", start_index)
                version = str(e)[start_index:end_index]
                print(f"當前Chrome version: {version}")
            # 如果 WebDriver 設置失敗，運行 download_driver.py
            else:
                # 如果無法從錯誤訊息中提取版本號，可以提示用戶手動輸入或使用其他方法
                version = input("請手動輸入您的 Chrome 主版本號的前三碼（例如：117）：")

            DownloadDriver.download_and_setup(version)
            # 再次嘗試設置 WebDriver
            service = Service("chromedriver.exe")
            self.driver = webdriver.Chrome(service=service, options=chrome_options)

        self.driver.get("https://www.railway.gov.tw/tra-tip-web/tip/tip001/tip123/query")
        WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.ID, "pid")))

    def accept_prompt(self):
        self.driver.find_element(By.ID, "pid").send_keys(self.name_id)

    def select_start_station(self):
        self.driver.find_element(By.ID, "startStation1").clear()
        self.driver.find_element(By.ID, "startStation1").send_keys(self.start_station)

    def select_end_station(self):
        self.driver.find_element(By.ID, "endStation1").clear()
        self.driver.find_element(By.ID, "endStation1").send_keys(self.end_station)

    def select_date(self):
        date_input = self.driver.find_element(By.ID, "rideDate1")
        date_input.clear()
        date_input.send_keys(self.date)

    def select_passenger_count(self):
        self.driver.find_element(By.ID, "normalQty1").clear()
        self.driver.find_element(By.ID, "normalQty1").send_keys(str(self.passenger_count))

    def by_train_number(self):
        self.driver.find_element(By.ID, "normalQty1").clear()
        self.driver.find_element(By.ID, "trainNoList1").send_keys(self.train_number)

    def by_time_slot(self):
        self.driver.find_element(By.XPATH, "//*[@id='queryForm']/div[1]/div[3]/div[2]/label[2]").click()
        self.driver.find_element(By.ID, "startTime1").send_keys(self.start_time)
        self.driver.find_element(By.ID, "endTime1").send_keys(self.end_time)
        self.driver.find_element(By.XPATH,
                                 "//*[@id='queryForm']/div[2]/div[4]/div/div[2]/label[1]").click()
        self.driver.find_element(By.XPATH,
                                 "//*[@id='queryForm']/div[2]/div[4]/div/div[2]/label[2]").click()
        self.driver.find_element(By.XPATH,
                                 "//*[@id='queryForm']/div[2]/div[4]/div/div[2]/label[3]").click()
        self.driver.find_element(By.XPATH,
                                 "//*[@id='queryForm']/div[2]/div[4]/div/div[2]/label[4]").click()

    def submit_form(self):
        self.driver.find_element(By.XPATH, "//*[@id='queryForm']/div[5]/input").click()

    def wait_for_confirmation(self):
        while True:
            try:
                reset_element = self.driver.find_element(By.ID, "reset")
                time.sleep(4)
                reset_element.click()
                time.sleep(1)
                return False
            except:
                pass

            try:
                self.driver.find_element(By.ID, "goBack")
                print("已找到可搭乘之車票。")
                os._exit(0)
            except:
                pass

if __name__ == "__main__":
    print("This is the train booking logic.")
