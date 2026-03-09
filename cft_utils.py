import json
import os
import shutil
import sys
import tempfile
import zipfile
from urllib.request import urlopen
from shutil import copyfileobj

import psutil

from form_utils import log_success, log_failure


CFT_METADATA_URL = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"
CFT_PLATFORM = "win64"
DEFAULT_CFT_TIMEOUT = 300.0  # seconds
DEFAULT_CFT_DOWNLOAD_RETRIES = 3


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


def _download_file(
    url: str,
    dest_path: str,
    timeout: float = DEFAULT_CFT_TIMEOUT,
    retries: int = DEFAULT_CFT_DOWNLOAD_RETRIES,
) -> None:
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    last_exception = None

    for attempt in range(1, retries + 1):
        log_success(f"Chrome for Testing をダウンロード中 ({attempt}/{retries}): {url}")
        try:
            # タイムアウトを明示しておかないと、ネットワーク不調時に
            # ダウンロード処理が無限に待ち続ける可能性があるため、
            # デフォルトで数分（DEFAULT_CFT_TIMEOUT）待つようにしています。
            with urlopen(url, timeout=timeout) as resp, open(dest_path, "wb") as out_f:
                copyfileobj(resp, out_f)
            return
        except Exception as e:  # noqa: BLE001
            last_exception = e
            try:
                os.remove(dest_path)
            except OSError:
                pass
            if attempt < retries:
                log_failure(f"Chrome for Testing のダウンロードに失敗したため再試行します: {e}")
            else:
                log_failure(f"Chrome for Testing のダウンロードに失敗しました: {e}")

    if last_exception is not None:
        raise last_exception


def _extract_zip(zip_path: str, dest_dir: str) -> None:
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(dest_dir)
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing の展開に失敗しました: {e}")
        raise


def _download_and_install_archive(
    url: str,
    zip_filename: str,
    extracted_dirname: str,
    expected_binary: str,
) -> None:
    root = _cft_root()
    final_dir = os.path.join(root, extracted_dirname)
    archive_path = os.path.join(root, zip_filename)
    partial_archive_path = archive_path + ".part"
    temp_extract_root = tempfile.mkdtemp(prefix=extracted_dirname + "-", dir=root)

    try:
        for leftover in (archive_path, partial_archive_path):
            try:
                os.remove(leftover)
            except OSError:
                pass

        _download_file(url, partial_archive_path)
        os.replace(partial_archive_path, archive_path)
        _extract_zip(archive_path, temp_extract_root)

        extracted_dir = os.path.join(temp_extract_root, extracted_dirname)
        expected_path = os.path.join(extracted_dir, expected_binary)
        if not os.path.exists(expected_path):
            raise RuntimeError(
                f"{extracted_dirname} の展開結果に必要なファイルが見つかりませんでした: {expected_binary}"
            )

        if os.path.exists(final_dir):
            shutil.rmtree(final_dir)
        os.replace(extracted_dir, final_dir)
    finally:
        try:
            os.remove(archive_path)
        except OSError:
            pass
        try:
            os.remove(partial_archive_path)
        except OSError:
            pass
        shutil.rmtree(temp_extract_root, ignore_errors=True)


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

    _download_and_install_archive(
        chrome_url,
        "chrome-win64.zip",
        "chrome-win64",
        "chrome.exe",
    )
    _download_and_install_archive(
        driver_url,
        "chromedriver-win64.zip",
        "chromedriver-win64",
        "chromedriver.exe",
    )

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
