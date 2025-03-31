import os
import datetime
import time
import jpholiday
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv

load_dotenv()

# 環境変数の読み込み
log_file_path = os.getenv("LOG_FILE_PATH")
company_holidays = os.getenv("COMPANY_HOLIDAYS").split(",")
login_id = os.getenv("LOGIN_ID")
password = os.getenv("PASSWORD")
check_in_url = os.getenv("CHECK_IN_URL")

# ログファイルの作成
os.makedirs(log_file_path, exist_ok=True)

now = datetime.datetime.now()
log_filename = os.path.join(log_file_path, now.strftime("%Y%m%d_%H%M%S") + "_log.txt")

print("Good morning!\nBefore I go to work, I check if today is a holiday.")

print("First. Checking if today is Saturday or Sunday...")

#今日が土曜日か日曜日の場合は終了
if datetime.datetime.now().weekday() in [5, 6]:
    print("Today is Saturday or Sunday. Exiting...")
    exit()

print("Second. Checking if today is holidays...")

#今日が祝日の場合は終了
year = datetime.datetime.now().year
holidays = jpholiday.year_holidays(year)
if datetime.datetime.now().date() in holidays:
    print("Today is a holiday. Exiting...")
    exit()

print("Third. Checking if today is a company holiday...")

# 会社の指定休業日の場合も終了
now_date = datetime.datetime.now().strftime("%m-%d")
if now_date in company_holidays:
    print("Today is a company holiday. Exiting...")
    exit()

print("Today is a working day. Let's go to work!")

# 実行を数秒待機
# カウントダウン
print("Starting check-in process in 10 seconds...")
for i in range(10, 0, -1):
    print(f"{i}...")
    time.sleep(1)

print("\nStarting Login process...\n")

# Chromeのオプションを設定
options = Options()
options.add_argument('--start-maximized')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
# options.add_argument('--headless')  # 必要に応じてヘッドレスモードを使用

webdriver_path = ChromeDriverManager().install()
if os.path.splitext(webdriver_path)[1] != '.exe':
    webdriver_dir_path = os.path.dirname(webdriver_path)
    webdriver_path = os.path.join(webdriver_dir_path, 'chromedriver.exe')
chrome_service = Service(executable_path=webdriver_path)
chrome = webdriver.Chrome(service=chrome_service, options=options)

with open(log_filename, "w", encoding="utf-8") as log_file:
    try:
        log_file.write("Start Time: {}\n".format(time.strftime("%Y/%m/%d %H:%M:%S")))

        # open check-in URL and log in
        log_file.write("Opening check-in URL...\n")
        chrome.get(check_in_url)
        log_file.write("Successfully opened check-in URL\n")

        # ログインIDとパスワードを入力
        log_file.write("Entering login ID and password...\n")
        login_id_field = WebDriverWait(chrome, 10).until(
            EC.presence_of_element_located((By.ID, "user_login_id"))
        )
        password_field = WebDriverWait(chrome, 10).until(
            EC.presence_of_element_located((By.ID, "user_password"))
        )

        log_file.write("Entering Login ID and password...\n")
        login_id_field.send_keys(login_id)
        log_file.write("Login ID entered.\n")
        password_field.send_keys(password)
        log_file.write("Password entered.\n")
        # パスワードを入力後、Enterキーを押す
        password_field.send_keys(Keys.RETURN)
        log_file.write("Password entered and Enter key pressed.\n")

        log_file.write("Successfully login\n")

        # 3秒待機
        log_file.write("Waiting for 3 seconds...\n")
        time.sleep(3)
        log_file.write("3 seconds wait completed.\n")

        # 出勤ボタンを探す
        log_file.write("Searching for check-in button...\n")
        for attempt in range(3):
            try:
                attendance_button = WebDriverWait(chrome, 10).until(
                    EC.presence_of_element_located((By.ID, "btnIN1"))
                )
                break
            except Exception as ex:
                log_file.write(f"Attempt {attempt + 1}: Attendance button not found. Retrying...\n")
                print(f"Attempt {attempt + 1}: Attendance button not found. Retrying...")
                time.sleep(1)
        else:
            raise Exception("Attendance button not found after 3 attempts.")

        log_file.write("Attendance button found.\n")

        # 出勤ボタンをクリック
        log_file.write("Clicking Attendance button...\n")
        attendance_button.click()
        log_file.write("Attendance button clicked.\n")

        log_file.write("Job completed!\n")
        log_file.write("End Time: {}\n".format(time.strftime("%Y/%m/%d %H:%M:%S")))
        log_file.write("\nCheck in time: {}\n".format(time.strftime("%H:%M:%S")))
        log_file.write("\nI hope you have a good day today too.\n")

        print("\nCheck in time: {}\n".format(time.strftime("%H:%M:%S")))
        print("\nI hope you have a good day today too.\n")

    except Exception as ex:
        log_file.write("Error during check-in process: {}\n".format(str(ex)))
        print("Error during check-in process: {}".format(str(ex)))

    finally:
        # 終了後はブラウザを閉じずに表示させ続ける
        while True:
            time.sleep(1)