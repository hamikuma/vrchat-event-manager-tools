import json
import time
import sys, os
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from form_utils import (
    log_success, log_failure,
    fill_input_by_label_with_retry, fill_textarea_by_label_with_retry,
    select_option_by_label_with_retry, click_button_by_text_with_retry,
    fill_datetime_by_label_with_retry, check_multiple_checkboxes_by_labels_with_retry,
    wait_for_label_with_retry, wait_for_form_section_change_with_retry, select_radio_by_label_with_retry, get_config_path,
    ensure_reply_email_checkbox_on, normalize_date_for_html,
)
from cft_utils import get_cft_paths, terminate_cft_processes


MAX_LOST_BROWSER_RETRIES = 3

def validate_config(config):
    """config.json に必須のキーが揃っているかをチェックする"""
    required_keys = [
        "form_url",
        "event_name",
        "start_hour",
        "start_minute",
        "end_hour",
        "end_minute",
        "event_host",
    ]

    missing = [key for key in required_keys if key not in config]
    if missing:
        log_failure("config.json に必須項目が不足しています:")
        for key in missing:
            log_failure(f"  - {key}")
        sys.exit(1)


def load_config():
    """設定ファイルを読み込み、バリデーションした結果を返す"""
    config_path = get_config_path()
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except FileNotFoundError as e:
        log_failure(f"設定ファイルが見つかりません: {config_path} {e}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        log_failure(f"設定ファイルがJSONとして正しく読み取れません: {config_path} {e}")
        sys.exit(1)

    validate_config(config)
    return config

def _run_impl(config, retry_count: int = 0) -> None:
    # ========================
    #  プロファイル格納先の設定
    # ========================

    if getattr(sys, 'frozen', False):
        profile_dir = os.path.join(os.path.dirname(sys.executable), "profile")
    else:
        profile_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "profile")

    # ========================
    # Chrome for Testing の取得（未取得ならダウンロード）
    # ========================
    try:
        chrome_exe, driver_exe = get_cft_paths()
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing の準備に失敗しました: {e}")
        sys.exit(1)

    # ========================
    # Chrome WebDriverのオプションを設定（CfT を使用）
    # ========================
    options = webdriver.ChromeOptions()
    options.binary_location = chrome_exe
    options.add_argument('--user-data-dir=' + profile_dir)
    options.add_argument('--profile-directory=Default')
    options.add_argument("--start-maximized")
    # 一部環境での起動クラッシュを避けるための互換オプション
    options.add_argument("--disable-gpu")
    options.add_experimental_option("detach", True)
    options.add_argument("--log-level=3")  # エラーだけ表示（INFO, WARNING, ERROR → 0〜3）
    options.add_experimental_option("excludeSwitches", ["enable-logging"])  # DevToolsやConsoleログを抑制

    # ========================
    # WebDriverを起動して Googleフォームを開く（CfT と対応する chromedriver を使用）
    # ========================

    driver = None
    wait = None

    for attempt in range(1, MAX_LOST_BROWSER_RETRIES + 1):
        # --- WebDriver 起動 ---
        try:
            driver = webdriver.Chrome(service=Service(driver_exe), options=options)
            # フォームの読み込み遅延に備えて待ち時間を少し長めに確保
            wait = WebDriverWait(driver, 10)
        except WebDriverException as e:
            message = str(e)
            # プロファイルディレクトリのロックなどで起動できない典型ケースを検出
            if "user data directory is already in use" in message or "profile is in use" in message:
                log_failure(
                    "WebDriverの起動に失敗しました: プロファイルが使用中の可能性があります。\n"
                    "ブラウザ(CfT)をすべて閉じてから、もう一度実行してください。"
                )
            # Chrome が起動直後にクラッシュし、DevToolsActivePort が作成されない典型ケース
            elif "DevToolsActivePort file doesn't exist" in message or "Chrome failed to start: crashed" in message:
                log_failure(
                    "WebDriverの起動に失敗しました: Chrome for Testing が正常に起動できませんでした。\n"
                    "ウイルス対策ソフト等でブロックされていないか確認し、\n"
                    "ツールと同じフォルダ内の cft フォルダを一度削除してから再実行してください。"
                )
            else:
                log_failure(f"WebDriverの起動に失敗しました: {e}")
            sys.exit(1)

        # --- Googleフォームを開く ---
        try:
            driver.get(config["form_url"])
            log_success("Googleフォームを開きました")
            break
        except WebDriverException as e:
            message = str(e)
            # 起動直後にウィンドウが閉じられた / クラッシュした典型ケース
            if "no such window" in message or "web view not found" in message:
                log_failure(
                    f"Chrome ウィンドウが起動直後に閉じられました（{attempt}/{MAX_LOST_BROWSER_RETRIES} 回目）。\n"
                    "セキュリティソフトや OS によるブロック、またはブラウザのクラッシュが考えられます。"
                )
                try:
                    driver.quit()
                except Exception:
                    pass

                if attempt >= MAX_LOST_BROWSER_RETRIES:
                    log_failure(
                        "Chrome ウィンドウを 3 回再起動しましたが、毎回すぐに閉じられました。\n"
                        "ウイルス対策ソフトの設定や Windows のイベントログを確認し、環境側の問題がないか確認してください。"
                    )
                    sys.exit(1)

                # 少し待ってから WebDriver/Chrome を作り直す
                time.sleep(1.0)
                continue

            # それ以外のエラーは従来どおり即座に失敗として扱う
            log_failure(f"Googleフォームを開くのに失敗しました: {e}")
            sys.exit(1)


    # ========================
    # 本番処理：フォームの自動入力
    # ========================

    try:
        # メールアドレスのチェックは無条件で最初にONを試みる
        ensure_reply_email_checkbox_on(driver, wait)

        fill_input_by_label_with_retry(driver, wait, "イベント名", config["event_name"])
        select_radio_by_label_with_retry(
            driver, wait, "Android対応可否", config.get("android_support", "PC/android")
        )

        # 開始日・終了日のデフォルト
        # - 開始日: 空欄なら当日
        # - 終了日: 空欄なら開始日と同じ（開始日も空欄なら当日）
        start_raw = config.get("start_date", "")
        end_raw = config.get("end_date", "")

        start_date = normalize_date_for_html(start_raw)
        if (end_raw or "").strip():
            end_date = normalize_date_for_html(end_raw)
        else:
            end_date = start_date

        fill_datetime_by_label_with_retry(
            driver,
            wait,
            "開始日時",
            start_date,
            config["start_hour"],
            config["start_minute"],
        )
        fill_datetime_by_label_with_retry(
            driver,
            wait,
            "終了日時",
            end_date,
            config["end_hour"],
            config["end_minute"],
        )

        select_option_by_label_with_retry(driver, wait, "イベントを登録しますか", "イベントを登録する")

        previous_section = driver.find_element("xpath", "//div[@role='list']")

        # 念のためもう一度
        ensure_reply_email_checkbox_on(driver, wait)

        click_button_by_text_with_retry(driver, wait, "次へ", max_retries=1)
        wait_for_form_section_change_with_retry(driver, previous_section)
        wait_for_label_with_retry(driver, "イベント主催者")

        fill_input_by_label_with_retry(driver, wait, "イベント主催者", config["event_host"])
        fill_textarea_by_label_with_retry(
            driver,
            wait,
            "イベント内容",
            config.get("event_content", ""),
        )
        check_multiple_checkboxes_by_labels_with_retry(
            driver,
            wait,
            "イベントジャンル",
            config.get("genres", []),
        )
        fill_textarea_by_label_with_retry(
            driver,
            wait,
            "参加条件",
            config.get("participation_conditions", ""),
        )
        fill_textarea_by_label_with_retry(
            driver,
            wait,
            "参加方法",
            config.get("participation_method", ""),
        )
        fill_textarea_by_label_with_retry(
            driver,
            wait,
            "備考",
            config.get("remarks", ""),
        )

        # 海外向け告知:
        #   - チェック ON のとき: 「希望する」を選択
        #   - チェック OFF のとき: 初期値「選択」に戻す
        overseas_on = bool(config.get("overseas_announcement"))
        overseas_value = "希望する" if overseas_on else "選択"

        try:
            # まずプルダウン形式を試す
            select_option_by_label_with_retry(
                driver,
                wait,
                "海外ユーザー向け告知",
                overseas_value,
            )
        except Exception:
            # 失敗した場合はラジオボタン形式を試す
            try:
                select_radio_by_label_with_retry(
                    driver,
                    wait,
                    "海外ユーザー向け告知",
                    overseas_value,
                )
            except Exception:
                # ここで例外を握りつぶすと後続処理に進めるが、
                # 詳細なログは form_utils 側で既に出力されている
                pass

        fill_textarea_by_label_with_retry(
            driver,
            wait,
            "X告知文",
            config.get("x_announcement", ""),
        )

        log_success("自動入力完了。スクリプトを終了します（ブラウザはそのまま）。")
        sys.exit(0)
    except Exception as e:
        message = str(e)
        message_lower = message.lower()

        # ブラウザ（Chrome for Testing）との接続が切れたと判断できる代表的なパターンをまとめて扱う
        lost_browser_keywords = [
            "no such window",
            "web view not found",
            "connection aborted.",
            "max retries exceeded with url",
            "failed to establish a new connection",
            "connection refused",
            "err_connection_refused",
        ]

        if any(keyword in message_lower for keyword in lost_browser_keywords):
            log_failure(
                "フォーム入力中にブラウザとの接続が失われました。\n"
                "Chrome for Testing の対象ウィンドウが OS やブラウザ側の理由で終了し、\n"
                "自動入力用のセッションが無効になった可能性があります。\n"
                "(画面上に別の Chrome ウィンドウが残っている場合でも、\n"
                "自動入力に使用していたウィンドウとは別インスタンスの可能性があります)\n"
                f"CfT を閉じてから自動的にやり直します ({retry_count + 1}/{MAX_LOST_BROWSER_RETRIES} 回目)。"
            )

            # まず現在の WebDriver / CfT を可能な範囲で終了させる
            try:
                if driver is not None:
                    driver.quit()
            except Exception:
                pass

            try:
                terminate_cft_processes()
            except Exception:
                # CfT の終了に失敗しても再試行自体は続行する
                pass

            if retry_count + 1 >= MAX_LOST_BROWSER_RETRIES:
                log_failure("ブラウザとの接続喪失が複数回発生したため、自動再試行を終了します。")
                sys.exit(1)

            # 少し待ってから処理全体を再実行
            time.sleep(2.0)
            _run_impl(config, retry_count + 1)
            return

        # それ以外の例外は従来通り通常のエラーとして扱う
        if isinstance(e, WebDriverException):
            log_failure(f"フォーム入力中にエラーが発生しました: {e}")
        else:
            log_failure(f"フォーム入力中に予期しないエラーが発生しました: {e}")
        sys.exit(1)


def main() -> None:
    """GUI/CLI エントリーポイント。

    設定ファイルを読み込み、ドライバーがブラウザを見失った場合は
    CfT を閉じた上で一定回数まで自動リトライする。
    """

    config = load_config()
    _run_impl(config, 0)


if __name__ == "__main__":
    main()