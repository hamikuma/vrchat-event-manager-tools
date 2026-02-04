from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any, Dict, List

from PySide6.QtCore import Qt, Signal, QTimer, QMimeData
from PySide6.QtGui import QFont, QIcon, QPixmap, QPalette, QColor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QProgressBar,
    QScrollArea,
    QTabWidget,
    QTimeEdit,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QSizePolicy,
)

import os
import sys


class PlainCopyTextEdit(QTextEdit):
    """コピー/ペーストともにプレーンテキストのみ扱う QTextEdit。

    - Ctrl+C や右クリックコピーでも、クリップボードには text/plain のみを入れる。
    - 他アプリからの貼り付け時も装飾付きの HTML を無視し、テキストだけを挿入する。
    """

    def createMimeDataFromSelection(self) -> QMimeData:  # type: ignore[override]
        cursor = self.textCursor()
        text = cursor.selectedText()
        # 行区切り用の U+2029 を通常の改行に置き換え
        text = text.replace("\u2029", "\n")
        mime = QMimeData()
        mime.setText(text)
        return mime

    def insertFromMimeData(self, source: QMimeData) -> None:  # type: ignore[override]
        """ペースト時に書式を捨ててプレーンテキストだけを挿入する。"""
        if source is None:
            return super().insertFromMimeData(source)

        text = source.text()
        if not text:
            # プレーンテキストが無い場合は通常の処理にフォールバック
            return super().insertFromMimeData(source)

        # 改行コードや U+2029 を統一
        text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\u2029", "\n")
        cursor = self.textCursor()
        cursor.insertText(text)


def _resource_base_dir() -> str:
    """リソースファイルを探す基準ディレクトリを返す。

    - 通常の .py 実行時: このファイル(gui_design.py)のあるフォルダ
    - PyInstaller の exe 実行時: 実行ファイル(.exe)のあるフォルダ

    こうすることで、配布時は exe と同じ階層に image.png を置けば
    読み込まれるようになる。
    """

    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


GENRE_CHOICES = [
    "アバター試着会",
    "改変アバター交流会",
    "その他交流会",
    "VR飲み会",
    "店舗系イベント",
    "音楽系イベント",
    "学術系イベント",
    "ロールプレイ",
    "初心者向けイベント",
    "定期イベント",
]


@dataclass
class TemplateWidgets:
    title_edit: QLineEdit
    body_edit: QTextEdit
    output_edit: QTextEdit
    notes_edit: QTextEdit
    save_button: QPushButton
    copy_output_button: QPushButton


class MainWindow(QMainWindow):
    # Controller から接続される Signal
    configFileSelectRequested = Signal()
    configSaveRequested = Signal(dict)
    createProfileRequested = Signal(dict)
    autofillRequested = Signal(dict)
    templateSaveRequested = Signal(dict)

    def __init__(self) -> None:
        super().__init__()

        # OS側のダークモード等に引きずられない、固定のライトテーマパレットを適用
        # （Qtの自動テーマ追従を避けるため、ここで明示的にスタイルとパレットを固定する）
        self._setup_palette()

        self.setWindowTitle("VRChatイベントカレンダー入力支援ツール")
        self.resize(1200, 1000)

        # 共通フォント設定（日本語・絵文字に配慮）
        base_font = QFont("Yu Gothic UI", 10)
        self.setFont(base_font)

        # 全体スタイル（パステル調の水色ベース）
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #e0f2ff;
            }
            QWidget#HeaderWidget {
                background-color: transparent;
            }
            QTabWidget::pane {
                border: 1px solid #9ac5ff;
                border-radius: 8px;
                background: #f4f9ff;
            }
            QTabBar::tab {
                background: #d7ebff;
                border: 1px solid #9ac5ff;
                border-radius: 10px 10px 0 0;
                color: #1e3a8a;
                padding: 6px 14px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #bcdfff;
            }
            QLabel {
                color: #1e293b;
            }
            QGroupBox {
                border: 1px solid #bfdbfe;
                border-radius: 8px;
                margin-top: 8px;
                background-color: #f8fbff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 6px 0 6px;
                color: #1d4ed8;
            }
            QLineEdit, QTextEdit, QComboBox {
                background: #ffffff;
                color: #111827;  /* OSのダークモードでも文字色が白にならないよう固定 */
                border: 1px solid #bfdbfe;
                border-radius: 6px;
                padding: 4px 6px;
            }
            QTextEdit {
                padding: 6px;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border-radius: 999px;  /* すべてのボタンをより丸く */
                padding: 6px 18px;
                border: none;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:disabled {
                background-color: #93c5fd;
            }
            QCheckBox {
                color: #1e293b;
                spacing: 6px;
            }
            QProgressBar {
                border: 1px solid #bfdbfe;
                border-radius: 6px;
                color: #1e293b;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #60a5fa;
                border-radius: 6px;
            }
            """
        )

        # アイコン・ヘッダー画像
        self._setup_icon()

        self._config_path_edit: QLineEdit | None = None
        self._form_widgets: Dict[str, Any] = {}
        self._genre_checkboxes: Dict[str, QCheckBox] = {}
        self._extras_var_edits: Dict[str, QLineEdit] = {}
        self._template_widgets: List[TemplateWidgets] = []

        self._log_edit: QTextEdit | None = None
        self._status_label: QLabel | None = None
        self._progress_bar: QProgressBar | None = None
        self._create_button: QPushButton | None = None
        self._autofill_button: QPushButton | None = None

        central = QWidget(self)
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(16, 8, 16, 12)
        main_layout.setSpacing(8)

        # 上部ヘッダー: 説明テキストのみ（画像はオプション）
        header_widget = self._create_header()
        main_layout.addWidget(header_widget)

        # タブ
        tabs = QTabWidget(self)
        tabs.addTab(self._create_input_tab(), "入力内容")
        tabs.addTab(self._create_execute_tab(), "実行 / ログ")
        tabs.addTab(self._create_extras_tab(), "おまけ")
        main_layout.addWidget(tabs)

    def _setup_palette(self) -> None:
        """アプリ全体のパレット/スタイルを固定し、OSのダークモードに影響されないようにする。"""

        app = QApplication.instance()
        if app is None:
            return

        # OS依存のスタイルではなく、Qt標準の Fusion スタイルに固定
        app.setStyle("Fusion")

        palette = QPalette()

        # ウィンドウ/背景色
        palette.setColor(QPalette.Window, QColor("#e0f2ff"))
        palette.setColor(QPalette.Base, QColor("#ffffff"))
        palette.setColor(QPalette.AlternateBase, QColor("#f8fbff"))

        # 文字色
        palette.setColor(QPalette.WindowText, QColor("#1e293b"))
        palette.setColor(QPalette.Text, QColor("#111827"))
        palette.setColor(QPalette.ButtonText, QColor("#ffffff"))
        palette.setColor(QPalette.ToolTipBase, QColor("#111827"))
        palette.setColor(QPalette.ToolTipText, QColor("#f9fafb"))

        # ボタン/強調色
        palette.setColor(QPalette.Button, QColor("#3b82f6"))
        palette.setColor(QPalette.Highlight, QColor("#60a5fa"))
        palette.setColor(QPalette.HighlightedText, QColor("#ffffff"))

        app.setPalette(palette)

    # -----------------
    # ヘッダー / アイコン
    # -----------------

    def _setup_icon(self) -> None:
        base_dir = _resource_base_dir()
        icon_path = os.path.join(base_dir, "favicon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

    def _create_header(self) -> QWidget:
        widget = QWidget(self)
        widget.setObjectName("HeaderWidget")
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(8, 8, 8, 4)
        layout.setSpacing(12)
        widget.setMaximumHeight(180)

        # 左側に画像（存在する場合）
        img_label = QLabel(widget)
        base_dir = _resource_base_dir()
        img_path = os.path.join(base_dir, "image.png")
        if os.path.exists(img_path):
            pix = QPixmap(img_path)
            if not pix.isNull():
                pix = pix.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                img_label.setPixmap(pix)
            img_label.setMinimumSize(160, 160)
        layout.addWidget(img_label)

        # 右側に説明テキスト
        text_container = QWidget(widget)
        text_layout = QVBoxLayout(text_container)
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_label = QLabel(
            "使い方：\n"
            "  1. 『入力内容』タブでイベント情報を入力します。\n"
            "  2. 『実行 / ログ』タブからプロファイル作成を実行し、Googleにログインします(初回のみ)。\n"
            "  3. 『実行 / ログ』タブからフォーム自動入力を実行します。",
            text_container,
        )
        # ヘッダーの使い方テキストは少し大きめのフォントにする
        header_font = QFont(text_label.font())
        ps = header_font.pointSize()
        if ps <= 0:
            header_font.setPointSize(11)
        else:
            header_font.setPointSize(ps + 1)
        text_label.setFont(header_font)
        text_label.setWordWrap(True)
        text_layout.addWidget(text_label)
        text_layout.addStretch(1)

        layout.addWidget(text_container, 1)
        # 画像とテキストの縦位置（横軸）を中央で揃える
        layout.setAlignment(img_label, Qt.AlignVCenter | Qt.AlignLeft)
        layout.setAlignment(text_container, Qt.AlignVCenter | Qt.AlignLeft)
        return widget

    # -----------------
    # 入力内容タブ
    # -----------------

    def _create_input_tab(self) -> QWidget:
        tab = QWidget(self)
        outer_layout = QVBoxLayout(tab)

        # 設定ファイルパス行
        path_layout = QHBoxLayout()
        path_label = QLabel("設定ファイルパス", tab)
        self._config_path_edit = QLineEdit(tab)
        browse_button = QPushButton("参照...", tab)
        save_button = QPushButton("入力内容保存", tab)

        browse_button.clicked.connect(self.configFileSelectRequested)
        save_button.clicked.connect(self._emit_save_requested)

        path_layout.addWidget(path_label)
        path_layout.addWidget(self._config_path_edit)
        path_layout.addWidget(browse_button)
        path_layout.addWidget(save_button)
        outer_layout.addLayout(path_layout)

        # スクロール可能なフォームエリア
        scroll = QScrollArea(tab)
        scroll.setWidgetResizable(True)
        form_container = QWidget(scroll)
        form_container.setStyleSheet("background-color: #ffffff;")
        scroll.setWidget(form_container)
        form_layout = QVBoxLayout(form_container)

        # 基本情報フォーム
        basic_group = QGroupBox("基本情報", form_container)
        basic_form = QFormLayout(basic_group)

        def add_line_edit(key: str, label_text: str) -> QLineEdit:
            edit = QLineEdit(basic_group)
            basic_form.addRow(QLabel(label_text, basic_group), edit)
            self._form_widgets[key] = edit
            # 予約変数に関係するフィールドは変更時に出力テキストを更新
            if key in {
                "event_name",
                "start_date",
                "end_date",
                "start_time",
                "end_time",
                "event_host",
            }:
                edit.textChanged.connect(self._update_template_outputs)
            return edit

        # form_url
        add_line_edit("form_url", "フォームURL(固定)")

        # イベント名（幅固定）
        event_name_edit = add_line_edit("event_name", "イベント名")
        event_name_edit.setMinimumWidth(320)
        event_name_edit.setMaximumWidth(320)

        # Android 対応可否
        android_combo = QComboBox(basic_group)
        android_combo.addItems(["PC", "PC/android", "android only"])
        android_combo.setMinimumWidth(180)
        android_combo.setMaximumWidth(180)
        basic_form.addRow(QLabel("Android対応可否", basic_group), android_combo)
        self._form_widgets["android_support"] = android_combo

        # 日時
        # 日付入力のヘルプ（ツールチップ）
        start_date_help_text = (
            "YYYYMMDD or YYYY/MM/DD or YYYY-MM-DD: 指定の日付\n"
            "空欄: 当日の日付\n"
            "月曜: 当日含む直近の月曜日の日付\n"
            "火曜: 当日含む直近の火曜日の日付\n"
            "水曜: 当日含む直近の水曜日の日付\n"
            "木曜: 当日含む直近の木曜日の日付\n"
            "金曜: 当日含む直近の金曜日の日付\n"
            "土曜: 当日含む直近の土曜日の日付\n"
            "日曜: 当日含む直近の日曜日の日付"
        )

        end_date_help_text = (
            "YYYYMMDD or YYYY/MM/DD or YYYY-MM-DD: 指定の日付\n"
            "空欄: 開始日と同じ日付 (開始日も空欄なら当日)\n"
            "月曜: 当日含む直近の月曜日の日付\n"
            "火曜: 当日含む直近の火曜日の日付\n"
            "水曜: 当日含む直近の水曜日の日付\n"
            "木曜: 当日含む直近の木曜日の日付\n"
            "金曜: 当日含む直近の金曜日の日付\n"
            "土曜: 当日含む直近の土曜日の日付\n"
            "日曜: 当日含む直近の日曜日の日付"
        )

        # 開始日: 自由入力テキストフォーム
        start_date_edit = add_line_edit("start_date", "開始日(YYYYMMDD,空欄で当日)")
        start_date_edit.setMinimumWidth(200)
        start_date_edit.setMaximumWidth(220)
        start_date_edit.setToolTip(start_date_help_text)

        # 開始時刻: 自由入力の hh:mm 形式テキスト（全角も許容）
        start_time_edit = add_line_edit("start_time", "開始時刻(hh:mm)")
        start_time_edit.setPlaceholderText("例: 23:00")
        start_time_edit.setMinimumWidth(90)
        start_time_edit.setMaximumWidth(110)

        # 終了日も同様に自由入力のテキストフォーム
        end_date_edit = add_line_edit("end_date", "終了日(YYYYMMDD,空欄で開始日と同じ)")
        end_date_edit.setMinimumWidth(200)
        end_date_edit.setMaximumWidth(220)
        end_date_edit.setToolTip(end_date_help_text)

        # 終了時刻も同様に自由入力の hh:mm 形式テキスト（全角も許容）
        end_time_edit = add_line_edit("end_time", "終了時刻(hh:mm)")
        end_time_edit.setPlaceholderText("例: 23:59")
        end_time_edit.setMinimumWidth(90)
        end_time_edit.setMaximumWidth(110)

        # イベント主催者（幅固定）
        event_host_edit = add_line_edit("event_host", "イベント主催者")
        event_host_edit.setMinimumWidth(320)
        event_host_edit.setMaximumWidth(320)

        form_layout.addWidget(basic_group)

        # 詳細設定
        detail_group = QGroupBox("詳細設定", form_container)
        detail_layout = QVBoxLayout(detail_group)

        # ジャンル
        genre_group = QGroupBox("イベントジャンル", detail_group)
        genre_layout = QGridLayout(genre_group)
        for idx, name in enumerate(GENRE_CHOICES):
            cb = QCheckBox(name, genre_group)
            row = idx // 3
            col = idx % 3
            genre_layout.addWidget(cb, row, col)
            self._genre_checkboxes[name] = cb
        detail_layout.addWidget(genre_group)

        # 複数行テキスト
        def add_text_edit(key: str, label_text: str) -> None:
            box = QGroupBox(label_text, detail_group)
            v = QVBoxLayout(box)
            edit = PlainCopyTextEdit(box)

            # 補足欄のみ高さを1行分程度にする
            if key == "remarks":
                fm = edit.fontMetrics()
                line_height = fm.lineSpacing()
                # 余白を考慮して +16px 程度を加算
                edit.setFixedHeight(line_height + 16)

            v.addWidget(edit)
            detail_layout.addWidget(box)
            self._form_widgets[key] = edit

        add_text_edit("event_content", "イベント内容")
        add_text_edit("participation_conditions", "参加条件")
        add_text_edit("participation_method", "参加方法")
        add_text_edit("remarks", "補足")
        add_text_edit("x_announcement", "X 告知文")

        # チェックボックス
        overseas_cb = QCheckBox("海外向け告知を行う", detail_group)
        self._form_widgets["overseas_announcement"] = overseas_cb
        detail_layout.addWidget(overseas_cb)

        email_cb = QCheckBox("メールアドレスを返信表示に記録 (常にON)", detail_group)
        email_cb.setChecked(True)
        email_cb.setEnabled(False)
        self._form_widgets["record_the_email_address_to_reply"] = email_cb
        detail_layout.addWidget(email_cb)

        form_layout.addWidget(detail_group)

        # 入力フォーム全体がタブの高さいっぱいに広がるように、
        # 追加のストレッチは入れずスクロールエリアをそのまま配置する
        outer_layout.addWidget(scroll)
        return tab

    # -----------------
    # 実行 / ログタブ
    # -----------------

    def _create_execute_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)

        btn_layout = QHBoxLayout()
        self._create_button = QPushButton("プロファイル作成", tab)
        self._autofill_button = QPushButton("フォーム自動入力", tab)
        self._create_button.clicked.connect(self._emit_create_requested)
        self._autofill_button.clicked.connect(self._emit_autofill_requested)
        btn_layout.addWidget(self._create_button)
        btn_layout.addWidget(self._autofill_button)
        layout.addLayout(btn_layout)

        self._log_edit = PlainCopyTextEdit(tab)
        self._log_edit.setReadOnly(True)
        layout.addWidget(self._log_edit)

        status_layout = QHBoxLayout()
        self._status_label = QLabel("待機中", tab)
        self._progress_bar = QProgressBar(tab)
        self._progress_bar.setRange(0, 0)  # インジケータ的に使用
        self._progress_bar.setVisible(False)
        status_layout.addWidget(self._status_label)
        status_layout.addWidget(self._progress_bar)
        layout.addLayout(status_layout)

        return tab

    # -----------------
    # おまけタブ
    # -----------------

    def _create_extras_tab(self) -> QWidget:
        tab = QWidget(self)
        layout = QVBoxLayout(tab)

        # タブ全体の説明（メイン機能とは独立したおまけであることを明示）
        desc_label = QLabel(
            "このタブは、イベント告知文などを作成・保持するための『おまけ』機能です。\n"
            "フォーム自動入力のメイン機能とは直接関係しない補助的な機能なので、\n"
            "必要な方だけご利用ください。",
            tab,
        )
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)

        # 変数 A〜E
        vars_group = QGroupBox("変数 A〜E", tab)
        vars_layout = QGridLayout(vars_group)
        for idx, name in enumerate(["A", "B", "C", "D", "E"]):
            label = QLabel(name, vars_group)
            edit = QLineEdit(vars_group)
            row = idx // 3
            col = (idx % 3) * 2
            vars_layout.addWidget(label, row, col)
            vars_layout.addWidget(edit, row, col + 1)
            self._extras_var_edits[name] = edit
            # 変数 A〜E が変わったらテンプレート出力を更新
            edit.textChanged.connect(self._update_template_outputs)
        layout.addWidget(vars_group)

        # テンプレートブロック ×5（スクロール）
        scroll = QScrollArea(tab)
        scroll.setWidgetResizable(True)
        container = QWidget(scroll)
        scroll.setWidget(container)
        v = QVBoxLayout(container)

        for i in range(5):
            block = QGroupBox(f"テンプレート {i+1}", container)
            b_layout = QVBoxLayout(block)

            title_edit = QLineEdit(block)
            b_layout.addWidget(QLabel("タイトル", block))
            b_layout.addWidget(title_edit)



            # 本文テンプレートと出力テキストを左右に並べるためのレイアウト
            content_row = QHBoxLayout()
            left_col = QVBoxLayout()
            right_col = QVBoxLayout()

            body_edit = PlainCopyTextEdit(block)
            body_edit.setFixedHeight(80)
            left_col.addWidget(QLabel("本文テンプレート", block))
            left_col.addWidget(body_edit)

            btn_row = QHBoxLayout()
            save_button = QPushButton("上書き保存", block)

            # ボタンは固定幅にして横いっぱいに広がらないようにする
            for btn, width in ((save_button, 110),):
                btn.setFixedWidth(width)
                sp = btn.sizePolicy()
                sp.setHorizontalPolicy(QSizePolicy.Fixed)
                btn.setSizePolicy(sp)

            btn_row.addWidget(save_button)
            btn_row.addStretch(1)
            left_col.addLayout(btn_row)

            output_edit = PlainCopyTextEdit(block)
            output_edit.setFixedHeight(80)
            output_edit.setReadOnly(True)
            right_col.addWidget(QLabel("出力テキスト", block))
            right_col.addWidget(output_edit)

            copy_output_button = QPushButton("クリップボードにコピー", block)
            # 出力用のコピー ボタンも同じ幅にしてテキストが見切れないようにする
            copy_output_button.setFixedWidth(140)
            sp_out = copy_output_button.sizePolicy()
            sp_out.setHorizontalPolicy(QSizePolicy.Fixed)
            copy_output_button.setSizePolicy(sp_out)
            # 出力テキストのコピー ボタンを右側カラムの左寄せで配置
            right_btn_row = QHBoxLayout()
            right_btn_row.addWidget(copy_output_button)

            # 出力テキスト用のコピー完了メッセージラベル（初期は非表示）
            output_status_label = QLabel("", block)
            output_status_label.setStyleSheet("color: #6b7280; font-size: 10px;")
            output_status_label.setVisible(False)
            right_btn_row.addWidget(output_status_label)

            right_btn_row.addStretch(1)
            right_col.addLayout(right_btn_row)

            # 左右カラムを 1:1 で並べる
            content_row.addLayout(left_col, 1)
            content_row.addLayout(right_col, 1)
            b_layout.addLayout(content_row)

            # タイトルと本文テンプレートの間に補足入力フォームを追加
            notes_label = QLabel("補足", block)
            b_layout.addWidget(notes_label)

            # 補足テキスト本体はラベルのすぐ下に配置
            notes_edit = PlainCopyTextEdit(block)
            notes_edit.setFixedHeight(40)
            b_layout.addWidget(notes_edit)

            # コピー ボタンは補足テキストの「下」に配置するレイアウト
            notes_row = QHBoxLayout()
            notes_copy_button = QPushButton("クリップボードにコピー", block)
            notes_copy_button.setFixedWidth(140)
            sp_notes = notes_copy_button.sizePolicy()
            sp_notes.setHorizontalPolicy(QSizePolicy.Fixed)
            notes_copy_button.setSizePolicy(sp_notes)
            notes_row.addWidget(notes_copy_button)

            # コピー完了メッセージ用の小さなラベル（初期は非表示）
            notes_status_label = QLabel("", block)
            notes_status_label.setStyleSheet("color: #6b7280; font-size: 10px;")
            notes_status_label.setVisible(False)
            notes_row.addWidget(notes_status_label)

            # 必要であれば右側に余白を入れてバランスをとる
            notes_row.addStretch(1)
            b_layout.addLayout(notes_row)

            tw = TemplateWidgets(
                title_edit=title_edit,
                body_edit=body_edit,
                output_edit=output_edit,
                notes_edit=notes_edit,
                save_button=save_button,
                copy_output_button=copy_output_button,
            )
            self._template_widgets.append(tw)

            # 補足入力欄のコピー
            notes_copy_button.clicked.connect(
                lambda _, w=notes_edit, lbl=notes_status_label: self._copy_to_clipboard(
                    w.toPlainText(), lbl
                )
            )

            # 出力テキストのコピー
            copy_output_button.clicked.connect(
                lambda _, w=output_edit, lbl=output_status_label: self._copy_to_clipboard(
                    w.toPlainText(), lbl
                )
            )

            # 上書き保存ボタン: 現在の入力内容（特にテンプレート）を
            # コントローラ経由で config.json に保存してもらう
            save_button.clicked.connect(self._emit_template_save_requested)

            # 本文テンプレートの変更で即座に出力テキストを更新
            body_edit.textChanged.connect(self._update_template_outputs)

            v.addWidget(block)

        # 予約変数の説明（スクロール内の最下段）
        info_group = QGroupBox("予約変数の説明", container)
        info_layout = QVBoxLayout(info_group)
        info_label = QLabel(
            "テンプレート本文では次の予約変数が使えます:\n"
            "  {TODAY} / {TOMORROW} : 本日 / 翌日の日付 (YYYY/MM/DD)\n"
            "  {EVENT_NAME}        : イベント名\n"
            "  {START_DATE}        : 開始日 (未入力時は当日)\n"
            "  {END_DATE}          : 終了日 (未入力時は当日)\n"
            "  {HOST}              : イベント主催者\n"
            "  {START_TIME}        : 開始時刻 (HH:MM)\n"
            "  {END_TIME}          : 終了時刻 (HH:MM)\n"
            "また {A}〜{E} は上部の変数入力欄の値で置き換えられます。",
            info_group,
        )
        info_label.setWordWrap(True)
        info_layout.addWidget(info_label)
        v.addWidget(info_group)

        v.addStretch(1)
        layout.addWidget(scroll)

        return tab

    # -----------------
    # コントローラ連携用 API
    # -----------------

    def _gather_form_values(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {}

        def get_text(key: str) -> str:
            w = self._form_widgets.get(key)
            if isinstance(w, QLineEdit):
                return w.text()
            if isinstance(w, QTextEdit):
                return w.toPlainText()
            if isinstance(w, QComboBox):
                return w.currentText()
            return ""

        data["form_url"] = get_text("form_url")
        data["event_name"] = get_text("event_name")

        # Android 対応
        android_combo = self._form_widgets.get("android_support")
        if isinstance(android_combo, QComboBox):
            data["android_support"] = android_combo.currentText()

        data["start_date"] = get_text("start_date")
        data["end_date"] = get_text("end_date")

        # "HH:MM" 形式（全角数字・全角コロンも含む）を
        # start_hour/start_minute, end_hour/end_minute に分解
        def _normalize_time_text(text: str) -> str:
            # 全角数字・全角コロンを半角に統一
            table = str.maketrans("０１２３４５６７８９：", "0123456789:")
            return text.translate(table)

        def split_time(key: str) -> tuple[str, str]:
            text = _normalize_time_text(get_text(key).strip())
            if not text:
                return "", ""
            parts = text.split(":", 1)
            if len(parts) != 2:
                return text, "00"
            h, m = parts[0].strip(), parts[1].strip()
            return h, m

        sh, sm = split_time("start_time")
        eh, em = split_time("end_time")
        data["start_hour"] = sh
        data["start_minute"] = sm
        data["end_hour"] = eh
        data["end_minute"] = em

        data["event_host"] = get_text("event_host")
        data["event_content"] = get_text("event_content")
        data["participation_conditions"] = get_text("participation_conditions")
        data["participation_method"] = get_text("participation_method")
        data["remarks"] = get_text("remarks")
        data["x_announcement"] = get_text("x_announcement")

        # チェックボックス系
        overseas_cb = self._form_widgets.get("overseas_announcement")
        if isinstance(overseas_cb, QCheckBox):
            data["overseas_announcement"] = overseas_cb.isChecked()

        email_cb = self._form_widgets.get("record_the_email_address_to_reply")
        if isinstance(email_cb, QCheckBox):
            data["record_the_email_address_to_reply"] = email_cb.isChecked()

        # ジャンル
        genres = [name for name, cb in self._genre_checkboxes.items() if cb.isChecked()]
        data["genres"] = genres

        # extras
        extras_vars: Dict[str, str] = {}
        for name, edit in self._extras_var_edits.items():
            extras_vars[name] = edit.text()

        templates: List[Dict[str, str]] = []
        for tw in self._template_widgets:
            templates.append(
                {
                    "title": tw.title_edit.text(),
                    "body": tw.body_edit.toPlainText(),
                    "notes": tw.notes_edit.toPlainText(),
                }
            )

        data["extras"] = {"variables": extras_vars, "templates": templates}
        return data

    # Controller から呼ばれるメソッド

    def set_config_path(self, path: str) -> None:
        if self._config_path_edit is not None:
            self._config_path_edit.setText(path)

    def set_form_values(self, values: Dict[str, Any]) -> None:
        def set_text(key: str, text: str) -> None:
            w = self._form_widgets.get(key)
            if isinstance(w, QLineEdit):
                w.setText(text)
            elif isinstance(w, QTextEdit):
                w.setPlainText(text)

        set_text("form_url", values.get("form_url", ""))
        set_text("event_name", values.get("event_name", ""))

        android_combo = self._form_widgets.get("android_support")
        if isinstance(android_combo, QComboBox):
            idx = android_combo.findText(values.get("android_support", "PC/android"))
            if idx >= 0:
                android_combo.setCurrentIndex(idx)

        # 日付は QDateEdit に反映（存在しない場合は当日）
        start_date_widget = self._form_widgets.get("start_date")
        if isinstance(start_date_widget, QDateEdit):
            raw = values.get("start_date", "")
            text = str(raw).strip()
            if text:
                # 既存のフォーマットを考慮しつつパース
                for fmt in ("yyyyMMdd", "yyyy/MM/dd", "yyyy-MM-dd"):
                    qd = QDate.fromString(text, fmt)
                    if qd.isValid():
                        start_date_widget.setDate(qd)
                        break
                else:
                    # 解釈できない場合は「空」扱い（minimumDate）
                    start_date_widget.setDate(start_date_widget.minimumDate())
            else:
                # 空の場合は minimumDate を使って「空」表示
                start_date_widget.setDate(start_date_widget.minimumDate())
        else:
            set_text("start_date", values.get("start_date", ""))

        end_date_widget = self._form_widgets.get("end_date")
        if isinstance(end_date_widget, QDateEdit):
            raw = values.get("end_date", "")
            text = str(raw).strip()
            if text:
                for fmt in ("yyyyMMdd", "yyyy/MM/dd", "yyyy-MM-dd"):
                    qd = QDate.fromString(text, fmt)
                    if qd.isValid():
                        end_date_widget.setDate(qd)
                        break
                else:
                    end_date_widget.setDate(end_date_widget.minimumDate())
            else:
                end_date_widget.setDate(end_date_widget.minimumDate())
        else:
            set_text("end_date", values.get("end_date", ""))

        # HH:MM 形式で時刻をまとめて表示（自由入力テキストに反映）
        start_hour = str(values.get("start_hour", "")).zfill(2) if values.get("start_hour") not in (None, "") else ""
        start_minute = str(values.get("start_minute", "")).zfill(2) if values.get("start_minute") not in (None, "") else ""
        end_hour = str(values.get("end_hour", "")).zfill(2) if values.get("end_hour") not in (None, "") else ""
        end_minute = str(values.get("end_minute", "")).zfill(2) if values.get("end_minute") not in (None, "") else ""

        start_time = f"{start_hour}:{start_minute}" if start_hour and start_minute else ""
        end_time = f"{end_hour}:{end_minute}" if end_hour and end_minute else ""

        set_text("start_time", start_time)
        set_text("end_time", end_time)

        set_text("event_host", values.get("event_host", ""))
        set_text("event_content", values.get("event_content", ""))
        set_text(
            "participation_conditions", values.get("participation_conditions", "")
        )
        set_text(
            "participation_method", values.get("participation_method", "")
        )
        set_text("remarks", values.get("remarks", ""))
        set_text("x_announcement", values.get("x_announcement", ""))

        overseas_cb = self._form_widgets.get("overseas_announcement")
        if isinstance(overseas_cb, QCheckBox):
            overseas_cb.setChecked(bool(values.get("overseas_announcement", False)))

        email_cb = self._form_widgets.get("record_the_email_address_to_reply")
        if isinstance(email_cb, QCheckBox):
            email_cb.setChecked(
                bool(values.get("record_the_email_address_to_reply", True))
            )

        # ジャンル
        genres = set(values.get("genres", []) or [])
        for name, cb in self._genre_checkboxes.items():
            cb.setChecked(name in genres)

        extras = values.get("extras", {}) or {}
        vars_values = extras.get("variables", {}) or {}
        for name, edit in self._extras_var_edits.items():
            edit.setText(str(vars_values.get(name, "")))

        templates_values = extras.get("templates", []) or []
        for i, tw in enumerate(self._template_widgets):
            if i < len(templates_values):
                item = templates_values[i]
                tw.title_edit.setText(str(item.get("title", "")))
                tw.body_edit.setPlainText(str(item.get("body", "")))
                # 補足があれば反映
                tw.notes_edit.setPlainText(str(item.get("notes", "")))

        # テンプレート出力を更新
        self._update_template_outputs()

    def _format_date_display(self, raw: str | None) -> str:
        """YYYYMMDD / YYYY/MM/DD / YYYY-MM-DD を YYYY/MM/DD に整形。

        空文字や None の場合は空文字を返す（START_DATE/END_DATE では呼び出し側で当日扱いにする）。
        """
        if not raw:
            return ""
        text = str(raw).strip()
        if not text:
            return ""
        for fmt in ("%Y%m%d", "%Y/%m/%d", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(text, fmt)
                return dt.strftime("%Y/%m/%d")
            except ValueError:
                continue
        # 解釈できない場合はそのまま返す
        return text

    def _update_template_outputs(self) -> None:
        """おまけタブの出力テキストを、変数・予約変数で展開して更新する。"""
        if not self._template_widgets:
            return

        # 現在のフォーム値をまとめて取得
        values = self._gather_form_values()

        # 予約変数
        today = datetime.today().date()
        tomorrow = today + timedelta(days=1)

        start_date_raw = values.get("start_date", "")
        end_date_raw = values.get("end_date", "")

        start_date_disp = self._format_date_display(start_date_raw)
        end_date_disp = self._format_date_display(end_date_raw)

        # 空欄の場合の扱い
        # - 開始日: 空なら当日
        # - 終了日: 空なら開始日と同じ（開始日も空なら当日）
        if not start_date_disp:
            start_date_disp = today.strftime("%Y/%m/%d")
        if not end_date_disp:
            end_date_disp = start_date_disp

        # START_DATE / END_DATE 用に年なし (MM/DD) 形式を作成
        def _to_month_day(text: str) -> str:
            if not text:
                return ""
            s = text.strip()
            for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d"):
                try:
                    dt = datetime.strptime(s, fmt)
                    return dt.strftime("%m/%d")
                except ValueError:
                    continue
            # フォーマット外の場合は、先頭4文字+区切りを除いて返す簡易処理
            if len(s) > 5 and s[4] in ("/", "-"):
                return s[5:]
            return s

        start_md = _to_month_day(start_date_disp)
        end_md = _to_month_day(end_date_disp)

        def _time_str(h_key: str, m_key: str) -> str:
            h = str(values.get(h_key, "")).strip()
            m = str(values.get(m_key, "")).strip()
            if not h or not m:
                return ""
            try:
                ih = int(h)
                im = int(m)
                return f"{ih:02d}:{im:02d}"
            except ValueError:
                return f"{h}:{m}"

        start_time = _time_str("start_hour", "start_minute")
        end_time = _time_str("end_hour", "end_minute")

        reserved: Dict[str, str] = {
            "TODAY": today.strftime("%Y/%m/%d"),
            "TOMORROW": tomorrow.strftime("%Y/%m/%d"),
            "EVENT_NAME": str(values.get("event_name", "")),
            # START_DATE / END_DATE は年なし MM/DD で展開
            "START_DATE": start_md,
            "END_DATE": end_md,
            "HOST": str(values.get("event_host", "")),
            "START_TIME": start_time,
            "END_TIME": end_time,
        }

        extras = values.get("extras", {}) or {}
        vars_values: Dict[str, str] = {
            k: str(v) for k, v in (extras.get("variables", {}) or {}).items()
        }

        # 各テンプレートごとに展開
        templates = extras.get("templates", []) or []
        for i, tw in enumerate(self._template_widgets):
            body = ""
            if i < len(templates):
                body = str(templates[i].get("body", ""))
            else:
                body = tw.body_edit.toPlainText()

            text = body
            # 1. 予約変数
            for name, val in reserved.items():
                text = text.replace("{" + name + "}", val)
            # 2. ユーザー変数 A〜E
            for name, val in vars_values.items():
                text = text.replace("{" + name + "}", val)

            tw.output_edit.setPlainText(text)

    def append_log_message(self, message: str) -> None:
        if self._log_edit is not None:
            self._log_edit.append(message)

    def set_status(self, text: str) -> None:
        if self._status_label is not None:
            self._status_label.setText(text)
        if self._progress_bar is not None:
            self._progress_bar.setVisible(text not in ("待機中", "完了", "エラー"))

    def set_running(self, running: bool) -> None:
        if self._create_button is not None:
            self._create_button.setEnabled(not running)
        if self._autofill_button is not None:
            self._autofill_button.setEnabled(not running)

    # -----------------
    # 内部ヘルパ
    # -----------------

    def _emit_save_requested(self) -> None:
        data = self._gather_form_values()
        self.configSaveRequested.emit(data)

    def _emit_create_requested(self) -> None:
        data = self._gather_form_values()
        self.createProfileRequested.emit(data)

    def _emit_autofill_requested(self) -> None:
        data = self._gather_form_values()
        self.autofillRequested.emit(data)

    def _emit_template_save_requested(self) -> None:
        """おまけタブのテンプレートを現在の設定ファイルに上書き保存する。"""
        data = self._gather_form_values()
        self.templateSaveRequested.emit(data)

    def _copy_to_clipboard(self, text: str, status_label: QLabel | None = None) -> None:
        # テキストが空でも「コピーしました」は表示する仕様とする
        clipboard = self.clipboard()
        clipboard.setText(text or "")

        # おまけタブのコピー時は、ボタン横に小さな「コピーしました」を2秒だけ表示
        if status_label is not None:
            status_label.setText("コピーしました")
            status_label.setVisible(True)

            def _hide_label() -> None:
                status_label.clear()
                status_label.setVisible(False)

            QTimer.singleShot(2000, _hide_label)

    def clipboard(self):
        # QApplication インスタンスからクリップボードを取得
        from PySide6.QtWidgets import QApplication

        return QApplication.clipboard()
