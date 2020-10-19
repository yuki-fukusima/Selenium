"""
検温サイト自動入力プログラム
"""

# 必要なライブラリのインポート
import time
from selenium import webdriver

# Chromeブラウザを起動する
driver = webdriver.Chrome("c:/driver/chromedriver.exe")

# 検温サイトを開く
driver.get("https://forms.office.com/Pages/ResponsePage.aspx?id=R3as8E6_Hk25fEKE8jiPZAP3vGlF6fZAkT_xCIYiawJUMVc2NkxIWjlXTThJMjJTRk1XMlhKVVA2TyQlQCN0PWcu")

# 測定日欄をクラス名で検索する
calender = driver.find_element_by_class_name('button-content')
# カレンダーボタンクリック
calender.click()
# カレンダーから今日日付をクラス名で検索する
today = driver.find_element_by_css_selector('.picker__day.picker__day--infocus.picker__day--today.picker__day--highlighted')
# 今日日付をクリック
today.click()

# 所属をクラス名で検索する
section = driver.find_element_by_class_name('select-placeholder')
# 所属欄クリック
section.click()

# ファイルから個人情報を読み取り
with open('data.txt', 'r', encoding='utf-8') as f:
    # 所属部署
    fSection = f.readline().rstrip('\n')
    # 社員番号
    fCode = f.readline()

    # 所属をクラス名で検索する
    sectionList = driver.find_elements_by_class_name('select-option-content')
    for item in sectionList:
        if item.text == fSection:
            mySection = item

    # プルダウンクリック
    mySection.click()

    # 社員IDをクラス名で検索する
    id = driver.find_element_by_css_selector('.office-form-question-textbox.office-form-textfield-input.form-control.border-no-radius')
    # 社員IDを入力
    id.send_keys(fCode)

# 結果をクラス名で検索する
kennonn = driver.find_element_by_css_selector('.office-form-question-choice-row.office-form-question-choice-text-row')
# 結果ラジオボタンクリック
kennonn.click()

# 送信ボタンをクラス名で検索する
submitButton = driver.find_element_by_css_selector('.__submit-button__.office-form-bottom-button.office-form-theme-button.button-control.light-background-button')
# 送信ボタンクリック
submitButton.click()

# 5秒待つ
time.sleep(5)
# ブラウザを閉じる
driver.quit()