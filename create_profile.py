import os
import subprocess
import sys
import json

from form_utils import (
    log_success,
    log_failure,
    get_config_path,
)
from cft_utils import get_cft_paths


def main():
    # ========================
    # 設定ファイル（config.json）の存在確認
    # ========================
    config_path = get_config_path()
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            json.load(f)
    except FileNotFoundError:
        log_failure(f"設定ファイルが見つかりません: {config_path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        log_failure(f"設定ファイルがJSONとして正しく読み取れません: {config_path} {e}")
        sys.exit(1)

    # === プロファイル格納先の設定 ===
    if getattr(sys, "frozen", False):
        profile_dir = os.path.dirname(sys.executable) + "\\profile"
    else:
        profile_dir = os.path.dirname(os.path.abspath(__file__)) + "\\profile"

    # === ディレクトリが存在しなければ作成 ===
    try:
        os.makedirs(profile_dir, exist_ok=True)
    except OSError as e:
        log_failure(f"プロファイルディレクトリの作成に失敗しました: {profile_dir} ({e})")
        sys.exit(1)

    # === Chrome for Testing の取得（未取得ならダウンロード） ===
    try:
        chrome_exe, _ = get_cft_paths()
    except Exception as e:  # noqa: BLE001
        log_failure(f"Chrome for Testing の準備に失敗しました: {e}")
        sys.exit(1)

    # === Chrome起動コマンドの準備 ===
    launch_cmd = [
        chrome_exe,
        f"--user-data-dir={profile_dir}",
        "--profile-directory=Default",
    ]

    log_success(
        f"Googleアカウントにログインいただき以下に専用プロファイルを作成します。:\n{profile_dir}"
    )

    # === Chromeを起動 ===
    try:
        subprocess.Popen(launch_cmd)
    except OSError as e:
        # 実行権限の問題や、セキュリティソフトによるブロックなどで
        # Chrome が起動できなかったケースをユーザーに分かりやすく伝える
        log_failure(f"Chrome の起動に失敗しました: {e}")
        sys.exit(1)

    # 待機させることでユーザーに説明を見せる
    log_success("Chromeが起動しました。Googleアカウントにログイン後、ブラウザを閉じてください")
    sys.exit(0)


if __name__ == "__main__":
    main()