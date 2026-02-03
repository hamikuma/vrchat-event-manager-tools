import json
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional

from PySide6.QtCore import QObject, QThread, Signal, Slot
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox

import form_utils


# -------------
# Config 管理
# -------------

REQUIRED_KEYS = [
    "form_url",
    "event_name",
    "start_hour",
    "start_minute",
    "end_hour",
    "end_minute",
    "event_host",
]


@dataclass
class ExtrasVariables:
    A: str = ""
    B: str = ""
    C: str = ""
    D: str = ""
    E: str = ""


@dataclass
class TemplateItem:
    title: str = ""
    body: str = ""
    # テンプレートごとの備考（URL などの任意テキスト）
    notes: str = ""


@dataclass
class Extras:
    variables: ExtrasVariables = field(default_factory=ExtrasVariables)
    templates: List[TemplateItem] = field(
        default_factory=lambda: [TemplateItem() for _ in range(5)]
    )


@dataclass
class AppConfig:
    form_url: str = ""
    record_the_email_address_to_reply: bool = True
    event_name: str = ""
    android_support: str = "PC/android"
    start_date: str = ""
    start_hour: str = "00"
    start_minute: str = "00"
    end_date: str = ""
    end_hour: str = "00"
    end_minute: str = "00"
    event_host: str = ""
    event_content: str = ""
    genres: List[str] = field(default_factory=list)
    participation_conditions: str = ""
    participation_method: str = ""
    remarks: str = ""
    x_announcement: str = ""
    overseas_announcement: bool = False
    extras: Extras = field(default_factory=Extras)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppConfig":
        extras_data = data.get("extras", {}) or {}
        variables_data = extras_data.get("variables", {}) or {}
        templates_data = extras_data.get("templates", []) or []

        variables = ExtrasVariables(
            A=str(variables_data.get("A", "")),
            B=str(variables_data.get("B", "")),
            C=str(variables_data.get("C", "")),
            D=str(variables_data.get("D", "")),
            E=str(variables_data.get("E", "")),
        )

        templates: List[TemplateItem] = []
        for i in range(5):
            item = templates_data[i] if i < len(templates_data) else {}
            templates.append(
                TemplateItem(
                    title=str(item.get("title", "")),
                    body=str(item.get("body", "")),
                    notes=str(item.get("notes", "")),
                )
            )

        extras = Extras(variables=variables, templates=templates)

        return cls(
            form_url=data.get("form_url", ""),
            record_the_email_address_to_reply=bool(
                data.get("record_the_email_address_to_reply", True)
            ),
            event_name=data.get("event_name", ""),
            android_support=data.get("android_support", "PC/android"),
            start_date=data.get("start_date", ""),
            start_hour=data.get("start_hour", "00"),
            start_minute=data.get("start_minute", "00"),
            end_date=data.get("end_date", ""),
            end_hour=data.get("end_hour", "00"),
            end_minute=data.get("end_minute", "00"),
            event_host=data.get("event_host", ""),
            event_content=data.get("event_content", ""),
            genres=list(data.get("genres", []) or []),
            participation_conditions=data.get("participation_conditions", ""),
            participation_method=data.get("participation_method", ""),
            remarks=data.get("remarks", ""),
            x_announcement=data.get("x_announcement", ""),
            overseas_announcement=bool(data.get("overseas_announcement", False)),
            extras=extras,
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "form_url": self.form_url,
            "record_the_email_address_to_reply": self.record_the_email_address_to_reply,
            "event_name": self.event_name,
            "android_support": self.android_support,
            "start_date": self.start_date,
            "start_hour": self.start_hour,
            "start_minute": self.start_minute,
            "end_date": self.end_date,
            "end_hour": self.end_hour,
            "end_minute": self.end_minute,
            "event_host": self.event_host,
            "event_content": self.event_content,
            "genres": list(self.genres),
            "participation_conditions": self.participation_conditions,
            "participation_method": self.participation_method,
            "remarks": self.remarks,
            "x_announcement": self.x_announcement,
            "overseas_announcement": self.overseas_announcement,
            "extras": {
                "variables": {
                    "A": self.extras.variables.A,
                    "B": self.extras.variables.B,
                    "C": self.extras.variables.C,
                    "D": self.extras.variables.D,
                    "E": self.extras.variables.E,
                },
                "templates": [
                    {"title": t.title, "body": t.body, "notes": t.notes}
                    for t in self.extras.templates
                ],
            },
        }


class ConfigManager(QObject):
    configChanged = Signal(AppConfig)
    configPathChanged = Signal(str)

    def __init__(self, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self._config_path: str = form_utils.get_config_path()
        self._config: AppConfig = AppConfig()

    @property
    def config_path(self) -> str:
        return self._config_path

    @config_path.setter
    def config_path(self, path: str) -> None:
        self._config_path = path
        self.configPathChanged.emit(path)

    @property
    def config(self) -> AppConfig:
        return self._config

    def load(self, path: Optional[str] = None) -> None:
        if path is not None:
            self.config_path = path
        try:
            with open(self._config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except FileNotFoundError:
            raise
        except json.JSONDecodeError as e:
            raise ValueError(f"JSON の形式が不正です: {e}") from e

        self._config = AppConfig.from_dict(data)
        self.configChanged.emit(self._config)

    def save(self, path: Optional[str] = None) -> None:
        if path is not None:
            self.config_path = path
        data = self._config.to_dict()
        with open(self._config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    # GUI 側から dict でもらった値で上書きする想定
    def update_from_dict(self, values: Dict[str, Any]) -> None:
        cfg = self._config
        for key in (
            "form_url",
            "record_the_email_address_to_reply",
            "event_name",
            "android_support",
            "start_date",
            "start_hour",
            "start_minute",
            "end_date",
            "end_hour",
            "end_minute",
            "event_host",
            "event_content",
            "genres",
            "participation_conditions",
            "participation_method",
            "remarks",
            "x_announcement",
            "overseas_announcement",
        ):
            if key in values:
                setattr(cfg, key, values[key])

        extras = values.get("extras", {}) or {}
        vars_values = extras.get("variables", {}) or {}
        for attr in ("A", "B", "C", "D", "E"):
            v = vars_values.get(attr)
            if v is not None:
                setattr(cfg.extras.variables, attr, str(v))

        tmpl_values = extras.get("templates", []) or []
        for i in range(min(5, len(tmpl_values))):
            item = tmpl_values[i]
            cfg.extras.templates[i].title = str(item.get("title", ""))
            cfg.extras.templates[i].body = str(item.get("body", ""))
            cfg.extras.templates[i].notes = str(item.get("notes", ""))

        self.configChanged.emit(cfg)


# -----------------
# バリデーション
# -----------------

class ValidationError(Exception):
    pass


def _normalize_date(value: str) -> Optional[str]:
    raw = (value or "").strip()
    if not raw:
        return None
    # 「月曜」〜「日曜」指定はバリデーション上は許可するが、値はそのまま保持する
    weekday_keywords = {
        "月曜",
        "月曜日",
        "火曜",
        "火曜日",
        "水曜",
        "水曜日",
        "木曜",
        "木曜日",
        "金曜",
        "金曜日",
        "土曜",
        "土曜日",
        "日曜",
        "日曜日",
    }

    if raw in weekday_keywords:
        # None を返すことで呼び出し側で元の文字列を維持する
        return None

    # それ以外は日付として解釈を試みる（保存時は YYYYMMDD に正規化）
    for fmt in ("%Y%m%d", "%Y/%m/%d", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(raw, fmt)
            return dt.strftime("%Y%m%d")
        except ValueError:
            continue
    raise ValidationError(f"日付の形式が不正です: {value}")


def validate_config_data(data: Dict[str, Any]) -> Dict[str, Any]:
    for key in REQUIRED_KEYS:
        value = data.get(key)
        if value is None or (isinstance(value, str) and not value.strip()):
            raise ValidationError(f"必須項目「{key}」が入力されていません。")

    # 日付
    for date_key in ("start_date", "end_date"):
        if date_key in data:
            try:
                normalized = _normalize_date(data.get(date_key, ""))
                if normalized is not None:
                    data[date_key] = normalized
            except ValidationError as e:
                raise ValidationError(str(e))

    # 時刻範囲チェック
    def _check_int_range(key: str, min_v: int, max_v: int) -> None:
        v = str(data.get(key, "0")).strip()
        try:
            iv = int(v)
        except ValueError:
            raise ValidationError(f"{key} は数値を入力してください。")
        if not (min_v <= iv <= max_v):
            raise ValidationError(f"{key} は {min_v}〜{max_v} の範囲で入力してください。")
        data[key] = f"{iv:02d}"

    _check_int_range("start_hour", 0, 23)
    _check_int_range("end_hour", 0, 23)
    _check_int_range("start_minute", 0, 59)
    _check_int_range("end_minute", 0, 59)

    genres = data.get("genres", []) or []
    if not isinstance(genres, list):
        raise ValidationError("genres は配列である必要があります。")
    # 空配列は許容（イベントジャンルは任意項目）

    return data


# -----------------
# 実行スレッド
# -----------------


class RunnerThread(QThread):
    logMessage = Signal(str)
    finishedWithStatus = Signal(str)  # "success" / "error"

    def __init__(self, mode: str, config_path: str, parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self._mode = mode  # "create_profile" or "autofill"
        self._config_path = config_path

    def _log_handler(self, message: str) -> None:
        self.logMessage.emit(message)

    def run(self) -> None:
        os.environ["VRC_EVENT_CONFIG_PATH"] = self._config_path
        form_utils.set_log_handler(self._log_handler)

        try:
            if self._mode == "create_profile":
                import create_profile

                try:
                    create_profile.main()
                except SystemExit as e:
                    code = e.code if isinstance(e.code, int) else 0
                    status = "success" if code == 0 else "error"
                    self.finishedWithStatus.emit(status)
                    return
            elif self._mode == "autofill":
                import autofill

                try:
                    autofill.main()
                except SystemExit as e:
                    code = e.code if isinstance(e.code, int) else 0
                    status = "success" if code == 0 else "error"
                    self.finishedWithStatus.emit(status)
                    return
            else:
                self.logMessage.emit(f"未知のモードです: {self._mode}")
                self.finishedWithStatus.emit("error")
                return
        except Exception as e:  # noqa: BLE001
            self.logMessage.emit(f"[WARN] 実行中にエラーが発生しました: {e}")
            self.finishedWithStatus.emit("error")
        finally:
            form_utils.set_log_handler(None)


# -----------------
# MainWindow 連携
# -----------------


class AppController(QObject):
    def __init__(self, window: "MainWindow", parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self.window = window
        self.config_manager = ConfigManager(self)
        self.runner_thread: Optional[RunnerThread] = None

        # 初期 config 読み込み
        self._init_config()

        # シグナル接続
        # gui_design.py 側で以下の Signal を定義している前提
        # - configFileSelectRequested
        # - configSaveRequested(dict)
        # - createProfileRequested(dict)
        # - autofillRequested(dict)
        # - templateSaveRequested(dict)
        # window 側の API も仮定して使用
        self._connect_signals()

    def _init_config(self) -> None:
        default_path = self.config_manager.config_path
        try:
            self.config_manager.load(default_path)
            self.window.set_config_path(default_path)
            self.window.set_form_values(self.config_manager.config.to_dict())
        except FileNotFoundError:
            # 初回はファイルがなくてもよい
            self.window.set_config_path(default_path)
        except ValueError as e:
            self._show_error(str(e))

    def _connect_signals(self) -> None:
        # ConfigManager -> Window
        self.config_manager.configChanged.connect(
            lambda cfg: self.window.set_form_values(cfg.to_dict())
        )
        self.config_manager.configPathChanged.connect(self.window.set_config_path)

        # Window -> Controller
        if hasattr(self.window, "configFileSelectRequested"):
            self.window.configFileSelectRequested.connect(self.on_select_config_file)
        if hasattr(self.window, "configSaveRequested"):
            self.window.configSaveRequested.connect(self.on_save_config_requested)
        if hasattr(self.window, "createProfileRequested"):
            self.window.createProfileRequested.connect(
                self.on_create_profile_requested
            )
        if hasattr(self.window, "autofillRequested"):
            self.window.autofillRequested.connect(self.on_autofill_requested)
        if hasattr(self.window, "templateSaveRequested"):
            self.window.templateSaveRequested.connect(self.on_template_save_requested)

    # ---------- UI ヘルパ ----------

    def _show_error(self, message: str) -> None:
        QMessageBox.critical(self.window, "エラー", message)

    def _show_info(self, message: str) -> None:
        QMessageBox.information(self.window, "情報", message)

    # ---------- Config 操作 ----------

    @Slot()
    def on_select_config_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self.window,
            "設定ファイルを選択",
            self.config_manager.config_path,
            "JSON ファイル (*.json)",
        )
        if not path:
            return
        try:
            self.config_manager.load(path)
            self.window.set_form_values(self.config_manager.config.to_dict())
            self._show_info("設定ファイルを読み込みました。")
        except FileNotFoundError:
            self._show_error("選択した設定ファイルが見つかりません。")
        except ValueError as e:
            self._show_error(str(e))

    @Slot(dict)
    def on_save_config_requested(self, values: Dict[str, Any]) -> None:
        try:
            validated = validate_config_data(values)
        except ValidationError as e:
            self._show_error(str(e))
            return

        self.config_manager.update_from_dict(validated)

        path, _ = QFileDialog.getSaveFileName(
            self.window,
            "設定ファイルを保存",
            self.config_manager.config_path,
            "JSON ファイル (*.json)",
        )
        if not path:
            return
        try:
            self.config_manager.save(path)
            self._show_info("設定ファイルを保存しました。")
        except OSError as e:  # noqa: BLE001
            self._show_error(f"設定ファイルの保存に失敗しました: {e}")

    # ---------- 実行処理 ----------

    def _start_runner(self, mode: str, values: Dict[str, Any]) -> None:
        if self.runner_thread is not None and self.runner_thread.isRunning():
            self._show_error("既に処理が実行中です。終了を待ってから再度実行してください。")
            return

        try:
            validated = validate_config_data(values)
        except ValidationError as e:
            self._show_error(str(e))
            return

        self.config_manager.update_from_dict(validated)
        try:
            self.config_manager.save()
        except OSError as e:  # noqa: BLE001
            self._show_error(f"設定ファイルの保存に失敗しました: {e}")
            return

        self.runner_thread = RunnerThread(mode, self.config_manager.config_path, self)
        self.runner_thread.logMessage.connect(self.window.append_log_message)
        self.runner_thread.finishedWithStatus.connect(self.on_runner_finished)

        if mode == "create_profile":
            self.window.set_status("プロファイル作成中...")
        elif mode == "autofill":
            self.window.set_status("自動入力実行中...")
        self.window.set_running(True)

        self.runner_thread.start()

    @Slot(dict)
    def on_create_profile_requested(self, values: Dict[str, Any]) -> None:
        self._start_runner("create_profile", values)

    @Slot(dict)
    def on_autofill_requested(self, values: Dict[str, Any]) -> None:
        self._start_runner("autofill", values)

    @Slot(dict)
    def on_template_save_requested(self, values: Dict[str, Any]) -> None:
        """おまけタブのテンプレートだけを現在の設定ファイルに上書き保存する。"""
        # テンプレート保存時は、他の必須項目が未入力でも動作するよう
        # validate_config_data は通さず、extras 部分をそのまま反映する
        extras = values.get("extras", {}) or {}
        minimal: Dict[str, Any] = {"extras": extras}

        self.config_manager.update_from_dict(minimal)
        try:
            self.config_manager.save()
            self._show_info("テンプレートを設定ファイルに上書き保存しました。")
        except OSError as e:  # noqa: BLE001
            self._show_error(f"テンプレートの保存に失敗しました: {e}")

    @Slot(str)
    def on_runner_finished(self, status: str) -> None:
        if status == "success":
            self.window.set_status("完了")
        else:
            self.window.set_status("エラー")
        self.window.set_running(False)


# -----------------
# エントリポイント
# -----------------


def main() -> None:
    app = QApplication(sys.argv)

    # gui_design.py から MainWindow をインポート
    try:
        from gui_design import MainWindow  # type: ignore
    except Exception as e:  # noqa: BLE001
        # exe 化した環境でも原因が分かるように標準エラーにも出力
        print(f"GUI デザインの読み込みに失敗しました: {e!r}", file=sys.stderr)
        QMessageBox.critical(None, "起動エラー", f"GUI デザインの読み込みに失敗しました: {e}")
        sys.exit(1)

    window = MainWindow()
    controller = AppController(window)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
