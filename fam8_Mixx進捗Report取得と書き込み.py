import time
import traceback
import pandas as pd
import gspread
from io import StringIO
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# 【設定項目】Googleスプレッドシート関連
#↓対象のスプレッドシートのリンクからスプレッドシートキーをとる
SPREADSHEET_KEY = "1F1S7-32FEdKywdfKSvgNOo_1K16oFOVS2St_GiCyOHA"
CREDENTIAL_FILE = "mixxexperiment-f8deafcf59d7.json"
ADFRAME_LIST_SHEET = "取得広告枠リスト(Mixx)"
# 【設定項目】スプレッドシート認証情報（冒頭に追加）
SHEET_NAME = "ピクシブ"



# 【設定項目】ログイン情報
ID = 'admin'
PASSWORD = 'fhC7UPJiforgKTJ8'

def setup_driver():
    """ WebDriver をセットアップして返す """
    print("[INFO] WebDriver を起動します")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.get("https://admin.fam-8.net/report/index.php")
    print("[INFO] WebDriver の起動完了")
    return driver

def direct_fam8_login(driver):
    """ ログイン処理 """
    try:
        print("[INFO] ログイン処理を開始します")
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[1]/td/input')))
        element.clear()
        element.send_keys(ID)
        element = driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[2]/td/input')
        element.clear()
        element.send_keys(PASSWORD)
        driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[3]/td/input[2]').click()
        time.sleep(2)
        print("[INFO] ログイン完了")
    except Exception as e:
        print("[ERROR] ログイン処理でエラー発生")
        print(traceback.format_exc())
        exit(1)

def operate_browser(driver):
    """ ブラウザ操作 """
    try:
        print("[INFO] 広告枠の設定処理を開始します")
        wait = WebDriverWait(driver, 10)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidemenu"]/div[3]/a[4]/div'))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="display_modesummary_mode"]'))).click()
        time.sleep(1)
        status_element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main_area"]/form/div[1]/table[2]/tbody/tr[1]/td/select[2]')))
        status_element.send_keys("オン")
        time.sleep(1)
        wait = WebDriverWait(driver, 5)
        input_element = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main_area"]/form/div[1]/input[7]')))
        input_element.clear()
        input_element.send_keys("Mixx")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main_area"]/form/div[1]/input[10]'))).click()
        time.sleep(1)
        net_button_xpath = '//*[@id="tbl_data"]/tbody/tr[1]/th[16]/a'
        print(f"[INFO] 'ネット' ボタンのクリックを試行: {net_button_xpath}")
        net_button = wait.until(EC.element_to_be_clickable((By.XPATH, net_button_xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", net_button)
        net_button.click()
        time.sleep(2)
        print("[INFO] 操作完了")
    except Exception as e:
        print("[ERROR] ブラウザ操作でエラー発生")
        print(traceback.format_exc())
    finally:
        print("[INFO] 対象スプレッドシートに書き込む情報を取得したブラウザを閉じます。")
#追加で変更
def get_adframe_ids(sheet_key, sheet_name, start_row=3, column='A'):
    """ スプレッドシートから広告枠IDを取得（無効なデータを除外） """
    creds = Credentials.from_service_account_file(
        CREDENTIAL_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    worksheet = client.open_by_key(sheet_key).worksheet(sheet_name)
    col_number = ord(column.upper()) - ord('A') + 1
    all_values = worksheet.col_values(col_number)[start_row - 1:]

        #【必ず挿入する場所はココ👇】
    print("[超重要確認]シート名:", sheet_name)
    print("[超重要確認]使用したスプレッドシートキー:", sheet_key)
    print("[超重要確認]認証情報ファイルのパス:", CREDENTIAL_FILE)
    print("[超重要確認]取得開始行番号:", start_row)
    print("[超重要確認]取得対象列:", column)
    print("[超重要確認]スプレッドシートから実際に取得した生データ:", all_values)

    # バリデーション処理: "0" や空白セルを除外し、数値のみ取得
    adframe_ids = []
    for val in all_values:
        val = val.strip()
        if val.isdigit() and val != "0":  # 数値かつ "0" ではないもののみ追加
            adframe_ids.append(val)

    print(f"[INFO] スプレッドシートから取得した広告枠IDリスト（バリデーション済み）: {adframe_ids}")
    return adframe_ids

#追加で変更
def extract_imp_values(driver, target_adframe_ids):
    """ 表から指定の広告枠IDに対応する 'Imp' の値を取得する """
    try:
        print("[INFO] 表データを取得します")
        table_html = driver.find_element(By.ID, "tbl_data").get_attribute("outerHTML")
        df = pd.read_html(table_html, header=0)[0]

        print("[DEBUG] 取得したデータフレーム:")
        print(df.head(10))

        adframe_id_col = df.columns[0]
        df[adframe_id_col] = df[adframe_id_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

        if "Imp" in df.columns:
            imp_col_num = df.columns.get_loc("Imp") + 1
            print(f"[DEBUG] 'Imp' カラムは {imp_col_num} 列目に存在")
        else:
            print("[ERROR] 'Imp' 列が見つかりません")
            return None

        imp_values = {}
        missing_ids = []  # 見つからなかったIDを記録

        for adframe_id in target_adframe_ids:
            matched_rows = df[df[adframe_id_col] == str(adframe_id)]
            if matched_rows.empty:
                print(f"[WARNING] スプレッドシートの広告枠ID {adframe_id} はブラウザで見つかりませんでした")
                imp_values[adframe_id] = None
                missing_ids.append(adframe_id)  # 見つからなかったIDを記録
            else:
                row_index = matched_rows.index[0] + 2
                imp_value = driver.find_element(By.XPATH, f'//*[@id="tbl_data"]/tbody/tr[{row_index}]/td[{imp_col_num}]').text.strip()
                imp_values[adframe_id] = imp_value

        # 見つからなかったIDをまとめてログに出力
        if missing_ids:
            print(f"[WARNING] 以下の広告枠IDはブラウザで見つかりませんでした: {missing_ids}")

        return imp_values
    except Exception as e:
        print("[ERROR] 表データの取得中にエラー発生")
        print(traceback.format_exc())
        return None

def get_imp_data_from_browser():
    driver = None
    try:
        print("[INFO] スクリプトを開始します")
        driver = setup_driver()
        direct_fam8_login(driver)
        operate_browser(driver)

        adframe_ids = get_adframe_ids(SPREADSHEET_KEY, ADFRAME_LIST_SHEET)
        print(f"[INFO] 取得した広告枠IDリスト: {adframe_ids}")
        imp_values = extract_imp_values(driver, adframe_ids)

        driver.quit()
        return imp_values  # 辞書をreturnする
    except Exception as e:
        print("[ERROR] スクリプト実行中にエラー発生")
        print(traceback.format_exc())
    #finally:
        #print("[INFO] スクリプト完了後もブラウザを開いたままにします。")
        #exit(0)
        # time.sleep(1000)



# 【設定項目】スプレッドシート認証情報
#↓対象のスプレッドシートのリンクからスプレッドシートキーをとる
SPREADSHEET_KEY = "1F1S7-32FEdKywdfKSvgNOo_1K16oFOVS2St_GiCyOHA"

SHEET_NAME = "ピクシブ"  # 書き込み先のシート名
CREDENTIAL_FILE = "mixxexperiment-f8deafcf59d7.json"  # 認証情報ファイル

# 【書き込みデータ】（広告枠IDとそれに対応するImp値の辞書）
# ②の処理の冒頭（main内）で、関数①の結果を受け取る形にする
imp_values = get_imp_data_from_browser()


def authenticate_gspread():
    """Googleスプレッドシートに認証し、クライアントを取得する。"""
    creds = Credentials.from_service_account_file(
        CREDENTIAL_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    client = gspread.authorize(creds)
    return client

def get_target_date():
    """
    本日 - 1日 の日付を "YYYY/MM/DD" 形式で取得する。
    スプレッドシート内のフォーマットに合わせて、年付きの日付を検索する。
    """
    target_date_raw = datetime.today() - timedelta(days=1)
    return target_date_raw.strftime("%Y/%m/%d")  # "2025/03/06" のような形式

def find_target_row(worksheet):
    """
    スプレッドシート内の日付（A列）を検索し、本日 - 1日 の日付がある行を特定する。
    :param worksheet: 操作対象のスプレッドシートのワークシートオブジェクト
    :return: 見つかった日付の行番号（見つからなければ None）
    """
    target_date = get_target_date()
    date_column = worksheet.col_values(1)[2:]  # A列（3行目以降）を取得

    print(f"[DEBUG] 検索対象の日付: {target_date}")
    print(f"[DEBUG] 取得した日付一覧: {date_column}")

    for idx, cell in enumerate(date_column, start=3):  # 3行目からループ
        if cell.strip() == target_date:
            return idx
    return None  # 見つからなかった場合

def find_target_column(worksheet, adframe_id):
    """
    広告枠ID {81810} のように {} を付与して検索し、その列を特定する。
    :param worksheet: 操作対象のスプレッドシートのワークシートオブジェクト
    :param adframe_id: 検索対象の広告枠ID
    :return: 見つかった広告枠IDの列番号（見つからなければ None）
    """
    search_id = f"{{{adframe_id}}}"  # 検索時に {} を付与
    data = worksheet.get_all_values()  # スプレッドシートの全データを取得

    print(f"[DEBUG] 検索する広告枠ID: {search_id}")

    for row_idx, row in enumerate(data, start=1):  # 各行をループ
        for col_idx, cell in enumerate(row, start=1):  # 各列をループ
            if cell.strip() == search_id:
                return col_idx  # 見つかった列の番号を返す
    return None  # 見つからなかった場合

def col_number_to_letter(col_num):
    """
    数値の列番号を Excel/Googleスプレッドシートの列名に変換（例：1 → A, 10 → J）。
    :param col_num: 列番号（1から始まる）
    :return: 文字列の列名（例："J"）
    """
    letter = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def write_imp_value(worksheet, row, col, value):
    """
    指定されたセルにデータを書き込む。
    :param worksheet: 操作対象のスプレッドシートのワークシートオブジェクト
    :param row: 書き込む行番号
    :param col: 書き込む列番号
    :param value: 書き込むデータ
    """
    col_letter = col_number_to_letter(col)  # 列番号をアルファベット表記に変換
    worksheet.update_cell(row, col, value)
    print(f"[SUCCESS] 書き込み完了: {value} を「{SHEET_NAME}」の ({col_letter}{row}) に書き込みました。")

if __name__ == "__main__":
    try:
        print("[INFO] ピクシブのスプレッドシートに接続します...")
        client = authenticate_gspread()
        worksheet = client.open_by_key(SPREADSHEET_KEY).worksheet(SHEET_NAME)

        print("[INFO] 書き込む対象日付を検索中...")
        target_row = find_target_row(worksheet)
        if target_row is None:
            print("[ERROR] 書き込む対象の日付が見つかりませんでした。")
            exit(1)
        #変更
        for adframe_id, imp_value in imp_values.items():

            print(f"[INFO] 広告枠ID {adframe_id} を検索中...")
            target_col = find_target_column(worksheet, adframe_id)
            if target_col is None:
                print(f"[ERROR] 広告枠ID {adframe_id} が見つかりませんでした。")
                continue  # 他のIDの処理を続行

            col_letter = col_number_to_letter(target_col)
            print(f"[INFO] 書き込みセルを特定: 「{SHEET_NAME}」の ({col_letter}{target_row})")

            write_imp_value(worksheet, target_row, target_col, imp_value)

        print("[INFO] 全ての処理が完了しました！")

    except Exception as e:
        print("[ERROR] 処理中にエラーが発生しました:", str(e))