import json
import os
import sys
import zipfile
from urllib.request import urlopen
from shutil import copyfileobj

import psutil

from form_utils import log_success, log_failure


CFT_METADATA_URL = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
CFT_PLATFORM = "win64"
DEFAULT_CFT_TIMEOUT = 300.0  # seconds


def _base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _cft_root() -> str:
    return os.path.join(_base_dir(), "cft")


def terminate_cft_processes() -> None:
    """Chrome for Testing (chrome / chromedriver) のプロセスを終了する。

    cft ディレクトリ配下の chrome.exe / chromedriver.exe のみを対象とし、
    通常インストールされた Chrome には影響を与えないようにする。
    """

    root = _cft_root()
    chrome_dir = os.path.normcase(os.path.join(root, "chrome-win64"))
    driver_dir = os.path.normcase(os.path.join(root, "chromedriver-win64"))

    try:
        for proc in psutil.process_iter(["exe", "name"]):
            exe_path = proc.info.get("exe") or ""
            if not exe_path:
                continue

            exe_norm = os.path.normcase(exe_path)

            # Chrome for Testing 本体
            if exe_norm.startswith(chrome_dir):
                try:
                    proc.terminate()
                    proc.wait(timeout=5)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                except psutil.TimeoutExpired:
                    try:
                        proc.kill()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                continue

            # ChromeDriver (念のため)
            if exe_norm.startswith(driver_dir):
                try:
                    proc.terminate()
                    proc.wait(timeout=5)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                except psutil.TimeoutExpired:
                    try:
                        proc.kill()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
    except Exception as e:  # noqa: BLE001
        # CfT の終了に失敗しても致命的ではないため、警告ログのみ出す。
        log_failure(f"Chrome for Testing のプロセス終了に失敗しました: {e}")


def _download_file(url: str, dest_path: str, timeout: float = DEFAULT_CFT_TIMEOUT) -> None:
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    log_success(f"Chrome for Testing をダウンロード中: {url}")
    try:
        # タイムアウトを明示しておかないと、ネットワーク不調時に
        # ダウンロード処理が無限に待ち続ける可能性があるため、
        # デフォルトで数分（DEFAULT_CFT_TIMEOUT）待つようにしています。
        with urlopen(url, timeout=timeout) as resp, open(dest_path, "wb") as out_f:
            copyfileobj(resp, out_f)
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing のダウンロードに失敗しました: {e}")
        raise


def _extract_zip(zip_path: str, dest_dir: str) -> None:
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(dest_dir)
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing の展開に失敗しました: {e}")
        raise


def _ensure_cft_downloaded() -> tuple[str, str]:
    root = _cft_root()
    chrome_exe = os.path.join(root, "chrome-win64", "chrome.exe")
    driver_exe = os.path.join(root, "chromedriver-win64", "chromedriver.exe")

    # 既にダウンロード済みならそのまま使う
    if os.path.exists(chrome_exe) and os.path.exists(driver_exe):
        return chrome_exe, driver_exe

    os.makedirs(root, exist_ok=True)

    log_success("Chrome for Testing が見つからないため、公式サイトから取得します。")

    # 最新の安定版バージョンと、対応するダウンロード URL を取得
    try:
        # メタデータ取得もネットワーク状況に依存するため、
        # ダウンロード自体と同様にタイムアウトを設定する。
        with urlopen(CFT_METADATA_URL, timeout=DEFAULT_CFT_TIMEOUT) as resp:
            meta = json.load(resp)
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing のメタデータ取得に失敗しました: {e}")
        raise

    try:
        stable = meta["channels"]["Stable"]
        downloads = stable["downloads"]

        def _find_url(kind: str) -> str:
            for entry in downloads.get(kind, []):
                if entry.get("platform") == CFT_PLATFORM:
                    return entry["url"]
            raise RuntimeError(f"{kind}({CFT_PLATFORM}) のダウンロード URL が見つかりませんでした")

        chrome_url = _find_url("chrome")
        driver_url = _find_url("chromedriver")
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing のダウンロード URL 解決に失敗しました: {e}")
        raise

    chrome_zip = os.path.join(root, "chrome-win64.zip")
    driver_zip = os.path.join(root, "chromedriver-win64.zip")

    _download_file(chrome_url, chrome_zip)
    _download_file(driver_url, driver_zip)

    _extract_zip(chrome_zip, root)
    _extract_zip(driver_zip, root)

    # ZIP は削除しておく（任意）
    try:
        os.remove(chrome_zip)
        os.remove(driver_zip)
    except OSError:
        pass

    if not (os.path.exists(chrome_exe) and os.path.exists(driver_exe)):
        raise RuntimeError("Chrome for Testing の実行ファイルが正しく展開されませんでした")

    log_success("Chrome for Testing のセットアップが完了しました。")
    return chrome_exe, driver_exe


def get_cft_paths() -> tuple[str, str]:
    """Chrome for Testing の chrome.exe / chromedriver.exe のパスを返す。

    未ダウンロードの場合は公式サイトから取得して展開する。
    失敗した場合は例外を送出するので、呼び出し側でログと終了処理を行う。
    """

    return _ensure_cft_downloaded()
