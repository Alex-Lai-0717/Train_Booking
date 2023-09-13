import requests
import os
import subprocess
import platform
import zipfile
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import shutil
import time

class DownloadDriver:
    @staticmethod
    def get_chrome_version_linux():
        print("正在獲取 Chrome 版本...")
        command = "google-chrome --version"
        process = subprocess.Popen(command.split(), stdout=subprocess.PIPE)
        output, error = process.communicate()
        if error is not None:
            print(f"Error: {error}")
        else:
            version = output.decode("utf-8").strip().split()[2]
            print(f"Chrome 版本： {version}")
            version = version.rsplit('.', 1)[0]  # 忽略版本號的最後一部分
            return version

    @staticmethod
    def download_chromedriver(version, os_type):
        print(f"正在下載 Chrome 版本 {version} 的最新 ChromeDriver...")
        driver_url = f"https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{version}/{os_type}/chromedriver-{os_type}.zip"
        response = requests.get(driver_url, stream=True)
        with open("chromedriver.zip", "wb") as file:
            for chunk in response.iter_content(chunk_size=128):
                file.write(chunk)

        print("正在解壓縮 ChromeDriver...")
        with zipfile.ZipFile('chromedriver.zip', 'r') as file:
            file.extractall()
        shutil.move("chromedriver-win64/chromedriver.exe", "./chromedriver.exe")
        shutil.rmtree("chromedriver-win64")
        print("ChromeDriver 準備就緒。")

    @staticmethod
    def setup_webdriver():
        print("正在設置 WebDriver...")
        chrome_options = webdriver.ChromeOptions()
        service = Service("chromedriver.exe")
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("正在測試ChromeDriver可被使用...")
        driver.get("https://www.google.com")
        time.sleep(3)
        print("WebDriver 測試成功。")
        print("WebDriver 已設置完成。")
        return driver

    @staticmethod
    def remove_temp_files():
        print("正在移除臨時文件...")
        try:
            os.remove("chromedriver.zip")
        except OSError as e:
            print(f"Error: {e.strerror}")

    @staticmethod
    def download_and_setup_win(version):
        DownloadDriver.download_chromedriver(version, "win64")
        driver = DownloadDriver.setup_webdriver()
        DownloadDriver.remove_temp_files()
        return driver

    @staticmethod
    def download_and_setup_linux(version):
        DownloadDriver.download_chromedriver(version, "linux64")
        driver = DownloadDriver.setup_webdriver()
        DownloadDriver.remove_temp_files()
        return driver

    @staticmethod
    def download_and_setup(version):
        driver = None
        system = platform.system()
        if system == "Windows":
            driver = DownloadDriver.download_and_setup_win(version)
        elif system == "Linux":
            driver = DownloadDriver.download_and_setup_linux(version)
        return driver
