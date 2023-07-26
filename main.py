from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import SessionNotCreatedException, WebDriverException
from ttkthemes import ThemedTk
from tkinter import messagebox
import re
import os
import pythoncom
import time
import datetime
import threading
import tkinter as tk
from datetime import date

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

        except SessionNotCreatedException:
            # 如果 WebDriver 設置失敗，運行 download_driver.py
            print("當前的 ChromeDriver 版本與 Chrome 瀏覽器版本不兼容，正在嘗試下載適配的 ChromeDriver 版本...")
            DownloadDriver.download_and_setup()
            # 再次嘗試設置 WebDriver
            service = Service("chromedriver.exe")
            self.driver = webdriver.Chrome(service=service, options=chrome_options)

        except WebDriverException:
            # 如果 WebDriver 設置失敗，運行 download_driver.py
            print("當前環境沒有找到ChromeDriver，正在嘗試下載適配的 ChromeDriver 版本...")
            DownloadDriver.download_and_setup()
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


# 支援提示文字
class PlaceholderEntry(tk.Entry):
    # 呼叫父類別（tk.Entry）的構造函數
    # **kwargs用來傳送鍵值對的可變數量的參數字典
    def __init__(self, master=None, placeholder="PLACEHOLDER", color='grey', **kwargs):
        super().__init__(master, **kwargs)

        # 儲存提示文字、提示文字的顏色，以及預設的文字顏色
        self.placeholder = placeholder
        self.placeholder_color = color
        self.default_fg_color = self['fg']

        # 將焦點獲取與焦點失去事件綁定到對應的方法
        self.bind("<FocusIn>", self.foc_in)
        self.bind("<FocusOut>", self.foc_out)

        # 初始時顯示提示文字
        self.put_placeholder()

    # 插入提示文字並將文字顏色變更為提示文字的顏色。
    def put_placeholder(self):
        self.insert(0, self.placeholder)
        self['fg'] = self.placeholder_color

    # 如果目前的文字是提示文字，則清除它並將文字顏色設置為預設顏色。
    def foc_in(self, *args):
        if self['fg'] == self.placeholder_color:
            self.delete('0', 'end')
            self['fg'] = self.default_fg_color

    # 如果輸入框是空的，則插入提示文字並將文字顏色設置為提示文字的顏色。
    def foc_out(self, *args):
        if not self.get():
            self.put_placeholder()


class TrainBookingGUI:
    station_options = [
        "1000-臺北",
        "1020-板橋",
        "1080-桃園",
        "3390-員林",
        "3360-彰化",
        "3420-田中",
        "4220-臺南",
    ]

    def __init__(self):
        self.window = ThemedTk(theme="arc")

        self.window.title("Train Booking")
        self.window.geometry("400x400")  # 設置窗口大小為 400x400
        self.create_widgets()

    def create_widgets(self):
        self.clear_widgets()  # 加入這行以清除現有部件

        self.mode_label = tk.Label(self.window, text="訂票方式：")
        self.mode_label.pack()

        self.mode_var = tk.StringVar(self.window)
        self.mode_var.set("車次")
        self.mode_option_menu = tk.OptionMenu(self.window, self.mode_var, "車次", "時段")
        self.mode_option_menu.pack()

        self.submit_button = tk.Button(self.window, text="確定", command=self.on_submit)
        self.submit_button.pack()

    def on_submit(self):
        self.clear_widgets()  # 加入這行以清除現有部件
        self.mode = self.mode_var.get()
        if self.mode == "車次":
            self.create_common_widgets(self.create_train_number_widgets)
        if self.mode == "時段":
            self.create_common_widgets(self.create_time_slot_widgets)

    def create_common_widgets(self, specific_widgets_func):
        self.clear_widgets()

        self.start_station_label = tk.Label(self.window, text="起始車站：")
        self.start_station_label.pack()

        self.start_station_var = tk.StringVar(self.window)
        self.start_station_var.set("1000-臺北")
        self.start_station_option_menu = tk.OptionMenu(self.window, self.start_station_var, *self.station_options)
        self.start_station_option_menu.pack()
        self.end_station_label = tk.Label(self.window, text="終點車站：")
        self.end_station_label.pack()

        self.end_station_var = tk.StringVar(self.window)
        self.end_station_var.set("1080-桃園")
        self.end_station_option_menu = tk.OptionMenu(self.window, self.end_station_var, *self.station_options)
        self.end_station_option_menu.pack()

        self.date_label = tk.Label(self.window, text="日期：")
        self.date_label.pack()
        today = date.today().strftime("%Y%m%d")
        self.date_entry = PlaceholderEntry(self.window, placeholder=today)
        self.date_entry.pack()

        self.passenger_count_label = tk.Label(self.window, text="乘客人數：")
        self.passenger_count_label.pack()
        self.passenger_count_entry = tk.Entry(self.window)
        self.passenger_count_entry.pack()

        self.name_id_label = tk.Label(self.window, text="身分證字號：")
        self.name_id_label.pack()
        self.name_id_entry = tk.Entry(self.window)
        self.name_id_entry.pack()

        specific_widgets_func()

        self.return_button = tk.Button(self.window, text="返回", command=self.create_widgets)
        self.return_button.pack()

    def create_train_number_widgets(self):
        self.train_number_label = tk.Label(self.window, text="車次：")
        self.train_number_label.pack()
        self.train_number_entry = tk.Entry(self.window)
        self.train_number_entry.pack()

        self.submit_button = tk.Button(self.window, text="確定", command=self.on_train_number_submit)
        self.submit_button.pack()

    def on_train_number_submit(self):
        if self.validate_input():
            start_station = self.start_station_var.get()
            end_station = self.end_station_var.get()
            date = self.date_entry.get()
            passenger_count = int(self.passenger_count_entry.get())
            name_id = self.name_id_entry.get()
            train_number = self.train_number_entry.get()

            booking = TrainBooking(start_station, end_station, date, passenger_count, name_id, train_number)
            threading.Thread(target=booking.book_ticket_by_train_number).start()
        else:
            messagebox.showerror("錯誤", "有不明原因導致輸入值有誤，請檢查並重新輸入。")

    def create_time_slot_widgets(self):
        self.start_time_label = tk.Label(self.window, text="起始時間：")
        self.start_time_label.pack()
        self.start_time_spinbox = tk.Spinbox(self.window,
                                             values=[f"{i:02d}:{j:02d}" for i in range(24) for j in range(0, 60, 30)])
        self.start_time_spinbox.delete(0, "end")  # 首先刪除 Spinbox 中的所有內容
        self.start_time_spinbox.insert(0, "18:00")  # 然後插入預設的時間
        self.start_time_spinbox.pack()

        self.end_time_label = tk.Label(self.window, text="結束時間：")
        self.end_time_label.pack()
        self.end_time_spinbox = tk.Spinbox(self.window,
                                           values=[f"{i:02d}:{j:02d}" for i in range(24) for j in range(0, 60, 30)])
        self.end_time_spinbox.delete(0, "end")  # 首先刪除 Spinbox 中的所有內容
        self.end_time_spinbox.insert(0, "21:00")  # 然後插入預設的時間
        self.end_time_spinbox.pack()

        self.submit_button = tk.Button(self.window, text="確定", command=self.on_time_slot_submit)
        self.submit_button.pack()

    def on_time_slot_submit(self):
        if self.validate_input():  # 如果輸入的資料是正確的
            start_station = self.start_station_var.get()
            end_station = self.end_station_var.get()
            date = self.date_entry.get()
            passenger_count = int(self.passenger_count_entry.get())
            name_id = self.name_id_entry.get()
            start_time = self.start_time_spinbox.get()
            end_time = self.end_time_spinbox.get()

            booking = TrainBooking(start_station, end_station, date, passenger_count, name_id, None, start_time,
                                   end_time)
            threading.Thread(target=booking.book_ticket_by_time_slot).start()
        else:
            messagebox.showerror("錯誤", "有不明原因導致輸入值有誤，請檢查並重新輸入。")

    def validate_input(self):
        validation_rules = []
        if self.mode == "車次":
            validation_rules = [
                self.validate_station(),
                self.validate_date(),
                self.validate_passenger_count(),
                self.validate_name_id(),
                self.validate_train_number(),
            ]
        if self.mode == "時段":
            validation_rules = [
                self.validate_station(),
                self.validate_date(),
                self.validate_passenger_count(),
                self.validate_name_id(),
                self.validate_time(),
            ]

        for rule in validation_rules:
            if not rule:
                return False
        return True

    # 起始與終點站驗證
    def validate_station(self):
        start_station = self.start_station_var.get()
        end_station = self.end_station_var.get()

        if start_station == end_station:
            messagebox.showerror("錯誤", "輸入的起始車站不能與終點車站相同，請檢查並重新輸入。")
            return False
        return True

    # 日期驗證
    def validate_date(self):
        date = self.date_entry.get()

        if not date:
            messagebox.showerror("錯誤", "輸入的日期不能為空，請檢查並重新輸入。")
            return False

        # 將輸入的日期和當天的日期進行比較
        # 檢查日期格式是否正確
        try:
            input_date = datetime.datetime.strptime(date, "%Y%m%d").date()  # 將輸入的日期由字串形式轉換為日期形式
        except ValueError:
            messagebox.showerror("錯誤", "輸入的日期格式不正確，請按照 'YYYYMMDD' 的格式輸入，例如:20231231")
            return False

        today = datetime.date.today()  # 獲取當天的日期
        if input_date < today:
            messagebox.showerror("錯誤", "輸入的日期不能早於當天的日期，請檢查並重新輸入。")
            return False
        return True

    # 票數驗證
    def validate_passenger_count(self):
        passenger_count = self.passenger_count_entry.get()

        if not passenger_count:
            messagebox.showerror("錯誤", "輸入的乘客數量不能為空，請檢查並重新輸入。")
            return False

        # 驗證輸入的乘客數量是否為 1 到 10 的整數
        if not re.match(r"^\d+$", passenger_count) or not 1 <= int(passenger_count) <= 10:
            messagebox.showerror("錯誤", "輸入的乘客數量必須是 1 到 10 的整數，請檢查並重新輸入。")
            return False
        return True

    # 身分證驗證
    def validate_name_id(self):
        name_id = self.name_id_entry.get()

        if not name_id:
            messagebox.showerror("錯誤", "輸入的身分證字號不能為空，請檢查並重新輸入。")
            return False

        # 檢查身分證號碼的長度
        if len(name_id) != 10:
            messagebox.showerror("錯誤", "輸入的身分證號碼長度不正確，身分證號碼應該是 10 個字符長，請檢查並重新輸入。")
            return False

        # 檢查身分證號碼的第一個字符是否是大寫字母
        if not name_id[0].isalpha():
            messagebox.showerror("錯誤", "輸入的身分證號碼的第一個字符必須是大寫字母，請檢查並重新輸入。")
            return False

        # 檢查身分證號碼的最後九個字符是否都是數字
        if not name_id[1:].isdigit():
            messagebox.showerror("錯誤", "輸入的身分證號碼的最後九個字符必須都是數字，請檢查並重新輸入。")
            return False

        if not self.verifyID(name_id):
            messagebox.showerror("錯誤", "輸入的身分證號碼格式不正確，請檢查並重新輸入。")
            return False
        return True

    # 車次驗證
    def validate_train_number(self):
        train_number = self.train_number_entry.get() if hasattr(self, 'train_number_entry') else None
        # 驗證輸入的車次號碼是否為整數
        if train_number is not None and not train_number.isdigit():
            messagebox.showerror("錯誤", "輸入的車次號碼必須是整數，請檢查並重新輸入。")
            return False
        return True

    # 時間驗證
    def validate_time(self):
        start_time = self.start_time_spinbox.get() if hasattr(self, 'start_time_spinbox') else None
        end_time = self.end_time_spinbox.get() if hasattr(self, "end_time_spinbox") else None

        if start_time is not None and end_time is not None:
            if start_time >= end_time:
                messagebox.showerror("錯誤", "結束時間不能早於或等於起始時間，請檢查並重新輸入。")
                return False
        return True

    def verifyID(self, id):
        # 英文代號對應數值表（個位數乘以 9 加上十位數）
        alphaTable = {'A': 1, 'B': 10, 'C': 19, 'D': 28, 'E': 37, 'F': 46,
                      'G': 55, 'H': 64, 'I': 39, 'J': 73, 'K': 82, 'L': 2, 'M': 11,
                      'N': 20, 'O': 48, 'P': 29, 'Q': 38, 'R': 47, 'S': 56, 'T': 65,
                      'U': 74, 'V': 83, 'W': 21, 'X': 3, 'Y': 12, 'Z': 30}

        # 計算總和值
        sum = alphaTable[id[0]] + int(id[1]) * 8 + int(id[2]) * 7 + int(id[3]) * 6 + int(id[4]) * 5 + int(
            id[5]) * 4 + int(id[6]) * 3 + int(id[7]) * 2 + int(id[8]) * 1 + int(id[9])

        # 驗證餘數
        if sum % 10 == 0:
            return True  # 身分證字號驗證通過
        else:
            return False  # 身分證字號有誤

    def clear_widgets(self):
        for widget in self.window.winfo_children():
            widget.pack_forget()

    def run(self):
        self.window.mainloop()

    @staticmethod
    def start():
        gui = TrainBookingGUI()
        gui.run()


if __name__ == "__main__":
    TrainBookingGUI.start()
