# 檔案名稱：run_register.py

import os
import platform
import sys
import time
import json
import random
import requests
from datetime import datetime
from appium import webdriver
from appium.options.android import UiAutomator2Options
from appium.webdriver.common.appiumby import AppiumBy
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from weasyprint import HTML
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError

# --- 關鍵修正：設定函式庫搜尋路徑 (for macOS) ---
if platform.system() == "Darwin":
    # 檢查並設定 Homebrew 的函式庫路徑
    homebrew_lib_path = '/opt/homebrew/lib'
    if os.path.exists(homebrew_lib_path):
        os.environ['DYLD_LIBRARY_PATH'] = homebrew_lib_path
# ------------------------------------

# =================================================================
# 設定區塊
# =================================================================
# --- Slack Bot 設定 (請務必填寫) ---
SLACK_BOT_TOKEN = "xoxb-XXX-XXX-XXX"  # 請填入您的 xoxb- 開頭的金鑰
SLACK_CHANNEL_ID = "C0XXXXX"  # 請填入您的 C 開頭的頻道 ID

# --- Appium 設定 ---
CAPABILITIES = {
    "platformName": "Android", "appium:automationName": "UiAutomator2",
    "appium:udid": "emulator-5554", "appium:appPackage": "tw.com.icash.a.icashpay.debuging",
    "appium:appActivity": "tw.net.pic.m.wallet.activity.WelcomeActivity",
    "appium:noReset": False, "appium:autoGrantPermissions": True,
}
APPIUM_SERVER_URL = 'http://localhost:4723'

# --- 元素定位符 (已整合您提供的所有新物件) ---
LOCATORS = {
    'pass_button': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/title_right_text'),
    'ad_close': (AppiumBy.XPATH, "//*[@resource-id='tw.com.icash.a.icashpay.debuging:id/label' and @text='關閉']"),
    'mine_tab': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/personal_text'),
    'login_register_button': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/text'),
    'register': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/btRegister'),
    'entermobilenumber': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/user_phones_text'),
    'enterloginame': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/user_code_text'),
    'enterpassword': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/user_pwd_text'),
    'enterpasswordagain': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/user_double_confirm_pwd_text'),
    'CheckBox1': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/cb_register_policies'),
    'CheckBox2': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/cb_op_register_policies'),
    'checkbox3': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/cb_register_policies_2'),
    'nextstep1': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/leftButton'),
    'nextstep2': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/leftButton'),
    'ERROR_POPUP_CONFIRM': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/top_text'),
    'IDbutton': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/textView112'),
    'enterrealname': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/et_user_name'),
    'birthdaybox': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/cl_register_bd'),
    'birthdayconfirm': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/btnSubmit'),
    'IDnumer': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/et_id_no'),
    'IDdate': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/textView43'),
    'IDLocation': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/tv_issue_loc'),
    'Hsinchu': (AppiumBy.XPATH, "//*[@resource-id='tw.com.icash.a.icashpay.debuging:id/text1' and @text='竹市']"),
    'nextstep3': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/leftButton'),
    'nextstep4': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/leftButton'),
    '6num1': (AppiumBy.XPATH,
              "(//android.widget.EditText[@resource-id='tw.com.icash.a.icashpay.debuging:id/editText'])[1]"),
    '6num2': (AppiumBy.XPATH,
              "(//android.widget.EditText[@resource-id='tw.com.icash.a.icashpay.debuging:id/editText'])[2]"),
    '6numconfirm': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/tvConfirm'),
    'nextstep5': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/leftButton'),
    'nexttime': (AppiumBy.ID, 'tw.com.icash.a.icashpay.debuging:id/tv_done'),
}


# =================================================================
# 功能函式區塊
# =================================================================

def generate_report(results, test_status):
    """產生報告 (HTML + PDF)，使用您指定的原始簡潔樣式"""
    html_filename = 'iCash_Pay_Register_Report.html'
    pdf_filename = 'iCash_Pay_Register_Report.pdf'
    styles = "<style>body{font-family:sans-serif;} table{border-collapse:collapse;} th,td{border:1px solid #ccc;padding:5px;}</style>"
    html_body = f"<h1>iCash Pay App 自動化註冊報告</h1><p><b>測試日期：</b>{datetime.now().strftime('%Y/%m/%d')}</p><p><b>測試結果：</b>{test_status}</p><hr><h2>測試步驟</h2><table><thead><tr><th>測試項目</th><th>測試結果</th></tr></thead><tbody>"
    for step in results:
        html_body += f"<tr><td>{step['步驟']}</td><td>{step['結果']}</td></tr>"
    html_body += "</tbody></table>"
    full_html = f"<!DOCTYPE html><html lang='zh-Hant'><head><meta charset='utf-8'><title>自動化測試報告</title>{styles}</head><body>{html_body}</body></html>"
    try:
        with open(html_filename, 'w', encoding='utf-8') as f:
            f.write(full_html)
        print(f"HTML 報告已成功產生： {html_filename}")
        HTML(string=full_html).write_pdf(pdf_filename)
        print(f"PDF 報告已成功產生： {pdf_filename}")
        return pdf_filename
    except Exception as e:
        print(f"!!! 產生報告時發生錯誤: {e}")
        return None


def upload_pdf_to_slack(pdf_path, test_status, results):
    """直接將 PDF 檔案上傳到 Slack"""
    if not SLACK_BOT_TOKEN.startswith("xoxb-"):
        print("未設定 Slack Bot Token，略過上傳。")
        return
    if not os.path.exists(pdf_path):
        print(f"找不到 PDF 檔案 {pdf_path}，略過上傳。")
        return

    client = WebClient(token=SLACK_BOT_TOKEN)

    if test_status == "Passed":
        comment = f"✅ 自動化註冊測試成功！\n報告日期：{datetime.now().strftime('%Y/%m/%d')}"
    else:
        failed_step = results[-1]['步驟'] if results else '未知步驟'
        comment = f"❌ 自動化註冊測試失敗！\n報告日期：{datetime.now().strftime('%Y/%m/%d')}\n失敗於步驟：{failed_step}"

    try:
        print("正在將 PDF 報告上傳到 Slack...")
        client.files_upload_v2(
            channel=SLACK_CHANNEL_ID,
            file=pdf_path,
            title=f"iCash_Pay_Register_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            initial_comment=comment
        )
        print("報告已成功上傳到 Slack！")
    except SlackApiError as e:
        print(f"!!! 上傳到 Slack 失敗: {e.response['error']}")


def generate_taiwan_id():
    """產生一組符合規則的隨機台灣身分證號碼"""
    letters = "ABCDEFGHJKLMNPQRSTUVXYWZIO"
    letter = random.choice(letters)
    gender = str(random.randint(1, 2))
    body = ''.join(random.choice('0123456789') for _ in range(7))
    letter_map = {'A': 10, 'B': 11, 'C': 12, 'D': 13, 'E': 14, 'F': 15, 'G': 16, 'H': 17, 'I': 34, 'J': 18, 'K': 19,
                  'L': 20, 'M': 21, 'N': 22, 'O': 35, 'P': 23, 'Q': 24, 'R': 25, 'S': 26, 'T': 27, 'U': 28, 'V': 29,
                  'W': 32, 'X': 30, 'Y': 31, 'Z': 33}
    n12 = letter_map[letter]
    n1 = n12 // 10
    n2 = n12 % 10
    nums = list(map(int, gender + body))
    checksum = n1 * 1 + n2 * 9 + nums[0] * 8 + nums[1] * 7 + nums[2] * 6 + nums[3] * 5 + nums[4] * 4 + nums[5] * 3 + \
               nums[6] * 2 + nums[7] * 1
    check_digit = (10 - (checksum % 10)) % 10
    return f"{letter}{gender}{body}{check_digit}"


def swipe_with_adb(driver, x, start_y, end_y, swipes=1):
    """(終極武器) 透過 ADB Shell 指令執行多次、精準的滑動"""
    for i in range(swipes):
        print(f"    -> 正在執行第 {i + 1}/{swipes} 次滑動...")
        command_args = ['input', 'swipe', str(x), str(start_y), str(x), str(end_y), '300']
        driver.execute_script('mobile: shell', {
            'command': command_args[0],
            'args': command_args[1:]
        })
        time.sleep(0.5)


def run_automation():
    """主執行流程 (全新註冊流程)"""
    driver = None
    results = []
    test_status = "Passed"

    def record_step(description, status="passed"):
        results.append({'步驟': description, '結果': status})
        print(f"執行： {description} [{status}]")

    try:
        options = UiAutomator2Options().load_capabilities(CAPABILITIES)
        driver = webdriver.Remote(APPIUM_SERVER_URL, options=options)
        wait = WebDriverWait(driver, 20)
        record_step("啟動 App")
        time.sleep(5)

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(LOCATORS['pass_button'])).click();
            record_step("略過引導頁")
        except:
            record_step("略過引導頁", "skipped (未出現)")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(LOCATORS['ad_close'])).click();
            record_step("關閉廣告彈窗")
        except:
            record_step("關閉廣告彈窗", "skipped (未出現)")

        time.sleep(2)

        wait.until(EC.element_to_be_clickable(LOCATORS['mine_tab'])).click();
        record_step("點擊我的頁面")
        wait.until(EC.element_to_be_clickable(LOCATORS['login_register_button'])).click();
        record_step("點擊登入/註冊")
        wait.until(EC.element_to_be_clickable(LOCATORS['register'])).click();
        record_step("點擊註冊按鈕")

        random_phone = f"09{random.randint(10000000, 99999999)}"
        random_account = f"aitest{random.randint(1000, 9999)}"
        wait.until(EC.visibility_of_element_located(LOCATORS['entermobilenumber'])).send_keys(random_phone);
        record_step(f"輸入隨機手機號: {random_phone}")
        wait.until(EC.visibility_of_element_located(LOCATORS['enterloginame'])).send_keys(random_account);
        record_step(f"輸入隨機帳號: {random_account}")

        try:
            WebDriverWait(driver, 3).until(EC.element_to_be_clickable(LOCATORS['ERROR_POPUP_CONFIRM'])).click()
            raise Exception(f"手機號或帳號 ({random_phone}/{random_account}) 已被註冊，測試中止")
        except TimeoutException:
            print("手機號與帳號可用，繼續下一步...")

        wait.until(EC.visibility_of_element_located(LOCATORS['enterpassword'])).send_keys('Aa123456');
        record_step("輸入密碼")
        wait.until(EC.visibility_of_element_located(LOCATORS['enterpasswordagain'])).send_keys('Aa123456');
        record_step("再次輸入密碼")

        wait.until(EC.element_to_be_clickable(LOCATORS['CheckBox1'])).click();
        record_step("勾選同意條款1")
        wait.until(EC.element_to_be_clickable(LOCATORS['CheckBox2'])).click();
        record_step("勾選同意條款2")
        wait.until(EC.element_to_be_clickable(LOCATORS['checkbox3'])).click();
        record_step("勾選同意條款3")

        wait.until(EC.element_to_be_clickable(LOCATORS['nextstep1'])).click();
        record_step("點擊下一步(1)")
        wait.until(EC.element_to_be_clickable(LOCATORS['nextstep2'])).click();
        record_step("點擊送出")
        wait.until(EC.element_to_be_clickable(LOCATORS['IDbutton'])).click();
        record_step("點擊國民身分證")
        wait.until(EC.visibility_of_element_located(LOCATORS['enterrealname'])).send_keys('測試一');
        record_step("輸入真實姓名")
        wait.until(EC.element_to_be_clickable(LOCATORS['birthdaybox'])).click();
        record_step("點擊生日選擇框")

        # --- 最終、最可靠的生日滑動邏輯 (ADB Shell 版) ---
        WebDriverWait(driver, 10).until(EC.presence_of_element_located(LOCATORS['birthdayconfirm']))
        record_step("日期選擇器已出現")

        record_step("密集滑動年份...")
        swipe_with_adb(driver, x=285, start_y=1300, end_y=1500, swipes=3)

        wait.until(EC.element_to_be_clickable(LOCATORS['birthdayconfirm'])).click();
        record_step("點擊生日確定")

        random_id_number = generate_taiwan_id()
        wait.until(EC.visibility_of_element_located(LOCATORS['IDnumer'])).send_keys(random_id_number);
        record_step(f"輸入隨機身分證號: {random_id_number}")

        wait.until(EC.element_to_be_clickable(LOCATORS['IDdate'])).click();
        record_step("點擊發證日期")
        wait.until(EC.element_to_be_clickable(LOCATORS['birthdayconfirm'])).click();
        record_step("點擊發證日期確定")
        wait.until(EC.element_to_be_clickable(LOCATORS['IDLocation'])).click();
        record_step("點擊發證地點")
        wait.until(EC.element_to_be_clickable(LOCATORS['Hsinchu'])).click();
        record_step("選擇新竹市")

        wait.until(EC.element_to_be_clickable(LOCATORS['nextstep3'])).click();
        record_step("點擊下一步(3)")
        wait.until(EC.element_to_be_clickable(LOCATORS['nextstep4'])).click();
        record_step("點擊確認")

        wait.until(EC.visibility_of_element_located(LOCATORS['6num1'])).send_keys('135790');
        record_step("輸入6位數密碼")
        wait.until(EC.visibility_of_element_located(LOCATORS['6num2'])).send_keys('135790');
        record_step("再次輸入6位數密碼")
        wait.until(EC.element_to_be_clickable(LOCATORS['6numconfirm'])).click();
        record_step("點擊6位數密碼確認")

        wait.until(EC.element_to_be_clickable(LOCATORS['nextstep5'])).click();
        record_step("點擊下一步(5)")
        wait.until(EC.element_to_be_clickable(LOCATORS['nexttime'])).click();
        record_step("點擊下次再說")

        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable(LOCATORS['ad_close'])).click();
            record_step("關閉最終廣告")
        except:
            record_step("關閉最終廣告", "skipped (未出現)")

        print("註冊流程成功完成！")

    except Exception as e:
        test_status = "Failed"
        error_message = f"測試失敗於: {results[-1]['步驟'] if results else '啟動 App'}"
        record_step(error_message, "failed")
        print(f"!!! 測試流程中斷！詳細錯誤: {e}")
    finally:
        if driver:
            driver.quit()
            record_step("關閉 APP")

        pdf_path = generate_report(results, test_status)

        if pdf_path:
            upload_pdf_to_slack(pdf_path, test_status, results)

        print("\n--- 所有流程已完成！ ---")


if __name__ == "__main__":
    run_automation()
