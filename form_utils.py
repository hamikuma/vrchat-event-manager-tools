import os
import sys
import time

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


LOG_HANDLER = None


def set_log_handler(handler):
    """GUI 側からログ受け取り関数を差し込むためのフック。

    handler が None でない場合、log_success/log_failure 呼び出し時に
    文字列メッセージを handler(message) で通知する。
    """
    global LOG_HANDLER
    LOG_HANDLER = handler

# ========================
# ログ出力関数
# ========================
def log_success(message):
    """処理成功時のログ出力（絵文字は使わず、どの環境でも出せる記号のみ使用）。"""
    text = f"[OK] {message}"
    print(text, flush=True)
    if LOG_HANDLER is not None:
        try:
            LOG_HANDLER(text)
        except Exception:
            # GUI 側のハンドラで例外が出ても処理全体が落ちないようにする
            pass


def log_failure(message="エラーが発生しました。"):
    """エラー時のログ出力（絵文字を避けて文字コード依存の問題を防ぐ）。"""
    text = f"[WARN] {message}"
    print(text, flush=True)
    if LOG_HANDLER is not None:
        try:
            LOG_HANDLER(text)
        except Exception:
            pass


def retry_func(func, *args, max_retries=3, **kwargs):
    """任意の関数を最大 ``max_retries`` 回までリトライするラッパー。

    すべてのリトライで失敗した場合は最後の例外を送出し、
    呼び出し元で明示的にエラーとして扱えるようにします。
    """
    last_exception = None

    for attempt in range(1, max_retries + 1):
        try:
            func(*args, **kwargs)
            return True
        except Exception as e:
            last_exception = e
            log_failure(f"リトライ {attempt}/{max_retries}: {e}")
            time.sleep(0.5)

    log_failure(f"最大リトライ回数({max_retries})に達しました。処理を中断します。")
    if last_exception is not None:
        raise last_exception

# ========================
# 共通操作関数
# ========================
def fill_input_by_label(driver, wait, label_text, value):
    try:
        input_elem = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//span[contains(text(), '{label_text}')]/ancestor::div[contains(@class, 'HoXoMd')]/following::input[@type='text'][1]"
        )))
        input_elem.clear()
        input_elem.send_keys(value)
        log_success(f"「{label_text}」に入力が完了しました")
    except Exception as e:
        log_failure(f"「{label_text}」の入力に失敗しました: {e}")
        raise


def fill_input_by_label_with_retry(driver, wait, label_text, value, max_retries=3):
    return retry_func(fill_input_by_label, driver, wait, label_text, value, max_retries=max_retries)

def fill_textarea_by_label(driver, wait, label_text, value):
    try:
        textarea_elem = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//span[contains(text(), '{label_text}')]/ancestor::div[contains(@class, 'HoXoMd')]/following::textarea[1]"
        )))
        textarea_elem.clear()
        textarea_elem.send_keys(value)
        log_success(f"「{label_text}」のテキストエリアに入力が完了しました")
    except Exception as e:
        log_failure(f"「{label_text}」のテキストエリア入力に失敗しました: {e}")
        raise


def fill_textarea_by_label_with_retry(driver, wait, label_text, value, max_retries=3):
    return retry_func(fill_textarea_by_label, driver, wait, label_text, value, max_retries=max_retries)

def select_option_by_label(driver, wait, label_text, option_text):
    try:
        label_elem = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//span[contains(text(), '{label_text}')]"
        )))
        container = label_elem.find_element(By.XPATH, "./ancestor::div[contains(@class, 'Qr7Oae')]")
        dropdown = container.find_element(By.XPATH, ".//div[@role='listbox']")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown)
        time.sleep(0.5)
        dropdown.click()
        time.sleep(1)
        option_elem = wait.until(EC.element_to_be_clickable((
            By.XPATH, f"//div[@role='option'][.//span[contains(text(), '{option_text}')]]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option_elem)
        option_elem.click()
        time.sleep(0.5)
        log_success(f"「{label_text}」に「{option_text}」を選択しました")
    except Exception as e:
        log_failure(f"「{label_text}」の選択に失敗しました: {e}")
        raise


def select_option_by_label_with_retry(driver, wait, label_text, option_text, max_retries=3):
    return retry_func(select_option_by_label, driver, wait, label_text, option_text, max_retries=max_retries)

def click_button_by_text(driver, wait, button_text):
    try:
        button = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//div[@role='button' and .//span[text()='{button_text}']]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
        time.sleep(0.5)
        try:
            WebDriverWait(driver, 1).until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "ThHDze"))
            )
        except:
            pass
        WebDriverWait(driver, 1).until(EC.element_to_be_clickable((
            By.XPATH, f"//div[@role='button' and .//span[text()='{button_text}']]"
        ))).click()
        log_success(f"「{button_text}」ボタンをクリックしました")
    except Exception as e:
        log_failure(f"「{button_text}」ボタンのクリックに失敗しました: {e}")
        raise


def click_button_by_text_with_retry(driver, wait, button_text, max_retries=3):
    return retry_func(click_button_by_text, driver, wait, button_text, max_retries=max_retries)

def fill_datetime_by_label(driver, wait, label_text, date_str, hour_str, minute_str):
    try:
        label_elem = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//span[contains(text(), '{label_text}')]"
        )))
        container = label_elem.find_element(By.XPATH, "./ancestor::div[contains(@class, 'Qr7Oae')]")
        date_input = container.find_element(By.XPATH, ".//input[@type='date']")
        hour_input = container.find_element(By.XPATH, ".//input[@type='text' and @aria-label='時']")
        minute_input = container.find_element(By.XPATH, ".//input[@type='text' and @aria-label='分']")

        date_input.clear()
        date_input.send_keys(date_str)
        hour_input.clear()
        hour_input.send_keys(hour_str)
        minute_input.clear()
        minute_input.send_keys(minute_str)

        log_success(f"「{label_text}」の日時入力が完了しました")
    except Exception as e:
        log_failure(f"「{label_text}」の日時入力に失敗しました: {e}")
        raise


def fill_datetime_by_label_with_retry(driver, wait, label_text, date_str, hour_str, minute_str, max_retries=3):
    return retry_func(
        fill_datetime_by_label,
        driver,
        wait,
        label_text,
        date_str,
        hour_str,
        minute_str,
        max_retries=max_retries,
    )

def check_multiple_checkboxes_by_labels(driver, wait, label_text, target_labels):
    try:
        # ラベル要素の検索修正
        label_elem = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//div[contains(@class, 'HoXoMd') and .//span[contains(text(), '{label_text}')]]"
        )))
        container = label_elem.find_element(By.XPATH, "./ancestor::div[contains(@class, 'Qr7Oae')]")

        # チェックボックス取得
        checkbox_labels = container.find_elements(By.XPATH, ".//*[@role='checkbox']")

        for checkbox in checkbox_labels:
            label = checkbox.get_attribute("aria-label") or checkbox.text.strip()
            is_checked = checkbox.get_attribute("aria-checked") == "true"

            if label in target_labels and not is_checked:
                checkbox.click()
                log_success(f"「{label_text}」の「{label}」にチェックを入れました")
            elif label in target_labels and is_checked:
                log_success(f"「{label_text}」の「{label}」は既にチェックされています")
            elif label not in target_labels and is_checked:
                checkbox.click()
                log_success(f"「{label_text}」の「{label}」のチェックを外しました")
    except Exception as e:
        log_failure(f"「{label_text}」の複数選択チェックに失敗しました: {e}")
        raise


def check_multiple_checkboxes_by_labels_with_retry(driver, wait, label_text, target_labels, max_retries=3):
    return retry_func(
        check_multiple_checkboxes_by_labels,
        driver,
        wait,
        label_text,
        target_labels,
        max_retries=max_retries,
    )

def select_radio_by_label(driver, wait, label_text, option_text):
    try:
        # ラベル要素取得
        label_elem = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//*[contains(text(), '{label_text}')]"
        )))
        container = label_elem.find_element(By.XPATH, "./ancestor::div[contains(@class, 'Qr7Oae')]")

        # aria-label で完全一致するラジオボタンを検索
        radio = container.find_element(By.XPATH, f".//*[@role='radio' and @aria-label='{option_text}']")

        # スクロールしてクリック
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", radio)
        time.sleep(0.5)
        if radio.get_attribute("aria-checked") != "true":
            radio.click()
            log_success(f"「{label_text}」で「{option_text}」を選択しました")
        else:
            log_success(f"「{label_text}」の「{option_text}」は既に選択されています")
    except Exception as e:
        log_failure(f"「{label_text}」のラジオボタン選択に失敗しました: {e}")
        raise


def select_radio_by_label_with_retry(driver, wait, label_text, option_text, max_retries=3):
    return retry_func(select_radio_by_label, driver, wait, label_text, option_text, max_retries=max_retries)

def wait_for_label(driver, label_text, timeout=10):
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, f"//span[contains(text(), '{label_text}')]"))
        )
        log_success(f"「{label_text}」が表示されました（ページ遷移完了）")
    except Exception as e:
        log_failure(f"「{label_text}」の表示待ちに失敗: {e}")
        raise


def wait_for_label_with_retry(driver, label_text, timeout=10, max_retries=3):
    return retry_func(wait_for_label, driver, label_text, timeout, max_retries=max_retries)

def wait_for_form_section_change(driver, previous_section):
    try:
        WebDriverWait(driver, 10).until(
            lambda d: d.find_element(By.XPATH, "//div[@role='list']") != previous_section
        )
    except Exception as e:
        log_failure(f"セクション切り替え待ちに失敗: {e}")
        raise


def wait_for_form_section_change_with_retry(driver, previous_section, max_retries=3):
    return retry_func(wait_for_form_section_change, driver, previous_section, max_retries=max_retries)

def get_config_path(filename="config.json"):
    # GUI から渡された設定ファイルパスがあればそれを優先して使用する
    env_path = os.environ.get("VRC_EVENT_CONFIG_PATH")
    if env_path:
        return env_path

    if getattr(sys, 'frozen', False):
        # PyInstallerでビルドされた実行ファイルの場合（.exe）
        base_path = os.path.dirname(sys.executable)
    else:
        # 通常の.pyスクリプト実行時
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, filename)