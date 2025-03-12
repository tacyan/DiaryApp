import sys
import os
import json
import datetime
import re
import calendar
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QTextEdit, QPushButton, QLabel, QCalendarWidget, QComboBox, 
                            QLineEdit, QMessageBox, QTabWidget, QGridLayout, QListWidget,
                            QListWidgetItem, QFileDialog, QColorDialog, QFontDialog, QMenu,
                            QAction, QToolBar, QStatusBar, QSplitter, QDialog, QCheckBox)
from PyQt5.QtGui import QFont, QIcon, QTextCharFormat, QColor, QTextCursor, QTextListFormat, QTextBlockFormat, QImage, QTextImageFormat, QPen
from PyQt5.QtCore import Qt, QDate, QTimer, QSize, QUrl
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent

class CustomCalendar(QCalendarWidget):
    """
    カスタムカレンダーウィジェット
    日記のある日付とお気に入りの日付を視覚的に表示する
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setGridVisible(True)
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.setHorizontalHeaderFormat(QCalendarWidget.SingleLetterDayNames)
        
        # 日記のある日付のリスト
        self.diary_dates = []
        
        # お気に入りの日付のリスト
        self.favorite_dates = []
        
    def updateCells(self):
        """
        カレンダーのセルを更新する
        """
        # 日記のある日付を取得
        self.diary_dates = []
        self.favorite_dates = []
        
        # 親ウィンドウからメタデータを取得
        if hasattr(self.parent, 'metadata') and hasattr(self.parent, 'diary_folder'):
            # 日記ファイルから日付を抽出
            for file_name in os.listdir(self.parent.diary_folder):
                if file_name.endswith('.json') and file_name != "metadata.json":
                    try:
                        # ファイル名から日付部分を抽出（yyyy-MM-dd_title-slug.json）
                        date_part = file_name.split('_')[0]
                        date = QDate.fromString(date_part, 'yyyy-MM-dd')
                        if date.isValid():
                            if date not in self.diary_dates:
                                self.diary_dates.append(date)
                            
                            # お気に入りかどうかチェック
                            file_key = file_name[:-5]  # .jsonを除去
                            if file_key in self.parent.metadata["favorites"] and date not in self.favorite_dates:
                                self.favorite_dates.append(date)
                    except:
                        continue
        
        super().updateCells()
    
    def paintCell(self, painter, rect, date):
        """
        カレンダーのセルを描画する
        """
        # デフォルトのセル描画
        super().paintCell(painter, rect, date)
        
        # 今日の日付の場合は特別な色で囲む
        if date == QDate.currentDate():
            painter.setPen(QPen(QColor(255, 0, 0), 2))
            painter.drawRect(rect.adjusted(1, 1, -1, -1))
        
        # 日記がある日付の場合は背景色を変える
        if date in self.diary_dates:
            # 日記の数を数える
            diary_count = 0
            date_str = date.toString('yyyy-MM-dd')
            
            for file_name in os.listdir(self.parent.diary_folder):
                if file_name.startswith(date_str) and file_name.endswith('.json') and file_name != "metadata.json":
                    diary_count += 1
            
            # 背景色を設定
            if date in self.favorite_dates:
                # お気に入りがある日付
                painter.fillRect(rect.adjusted(2, 2, -2, -2), QColor(255, 182, 193, 150))  # 薄いピンク
            else:
                # 通常の日記がある日付
                painter.fillRect(rect.adjusted(2, 2, -2, -2), QColor(173, 216, 230, 150))  # 薄い青
            
            # 複数の日記がある場合は数を表示
            if diary_count > 1:
                painter.setPen(QPen(QColor(0, 0, 255)))
                painter.drawText(rect.adjusted(0, 0, -2, -int(rect.height() / 2)), 
                                Qt.AlignRight | Qt.AlignBottom, 
                                f"{diary_count}")

class CustomTextEdit(QTextEdit):
    """
    キーイベントをカスタマイズしたQTextEditのサブクラス
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.heading_applied = False  # 見出しが適用されたかどうかのフラグ
    
    def keyPressEvent(self, event):
        """
        キー入力イベントをオーバーライドして、Enterキーが押されたときに通常のテキストスタイルに戻す
        """
        if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            cursor = self.textCursor()
            
            # 親クラスのイベント処理を呼び出し（改行を実行）
            super().keyPressEvent(event)
            
            # 改行後に通常のテキストスタイルを適用
            if self.heading_applied:
                cursor = self.textCursor()
                normal_format = QTextCharFormat()
                normal_format.setFontPointSize(11)  # 通常のフォントサイズに戻す
                normal_format.setFontWeight(QFont.Normal)  # 通常の太さに戻す
                cursor.setCharFormat(normal_format)
                self.heading_applied = False  # フラグをリセット
                self.setTextCursor(cursor)  # カーソル位置を更新
        else:
            # その他のキーイベントは通常通り処理
            super().keyPressEvent(event)

class DiaryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 日記アプリ")
        self.setGeometry(100, 100, 1000, 700)
        
        # データの保存先
        self.diary_folder = "diary_entries"
        if not os.path.exists(self.diary_folder):
            os.makedirs(self.diary_folder)
        
        # 画像保存フォルダ
        self.images_folder = os.path.join(self.diary_folder, "images")
        if not os.path.exists(self.images_folder):
            os.makedirs(self.images_folder)
            
        # メタデータファイル
        self.metadata_file = os.path.join(self.diary_folder, "metadata.json")
        if os.path.exists(self.metadata_file):
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                self.metadata = json.load(f)
        else:
            self.metadata = {
                "tags": [],
                "moods": ["楽しい", "普通", "悲しい", "疲れた", "興奮", "不安", "満足"],
                "favorites": [],
                "theme": "light"
            }
            self.save_metadata()
        
        # 現在の日付と選択された日付
        self.current_date = QDate.currentDate()
        self.selected_date = self.current_date
        
        # フォント設定
        self.current_font = QFont("Yu Gothic", 11)
        
        # メインウィジェットの設定
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)
        
        # スプリッターを追加
        self.splitter = QSplitter(Qt.Horizontal)
        self.main_layout.addWidget(self.splitter)
        
        # 左側のウィジェット
        self.left_widget = QWidget()
        self.left_layout = QVBoxLayout(self.left_widget)
        self.splitter.addWidget(self.left_widget)
        
        # カレンダーの設定
        self.calendar = CustomCalendar(self)
        self.calendar.setSelectedDate(self.current_date)
        self.calendar.selectionChanged.connect(self.date_selected)
        self.left_layout.addWidget(self.calendar)
        
        # 日付表示ラベル
        self.date_label = QLabel()
        self.date_label.setAlignment(Qt.AlignCenter)
        self.left_layout.addWidget(self.date_label)
        
        # 日記数表示ラベル
        self.entry_count_label = QLabel()
        self.entry_count_label.setAlignment(Qt.AlignCenter)
        self.left_layout.addWidget(self.entry_count_label)
        self.entry_count_label.setVisible(False)
        
        # タグリスト
        self.left_layout.addWidget(QLabel("タグ一覧"))
        self.tag_list = QListWidget()
        self.tag_list.itemClicked.connect(self.filter_by_tag)
        self.update_tag_list()
        self.left_layout.addWidget(self.tag_list)
        
        # お気に入りリスト
        self.left_layout.addWidget(QLabel("お気に入り"))
        self.favorites_list = QListWidget()
        self.favorites_list.itemClicked.connect(self.open_favorite)
        self.update_favorites_list()
        self.left_layout.addWidget(self.favorites_list)
        
        # 右側のウィジェット
        self.right_widget = QWidget()
        self.right_layout = QVBoxLayout(self.right_widget)
        self.splitter.addWidget(self.right_widget)
        
        # 日付表示
        self.date_label = QLabel()
        self.update_date_label()
        self.right_layout.addWidget(self.date_label)
        
        # タイトル入力
        self.title_layout = QHBoxLayout()
        self.title_layout.addWidget(QLabel("タイトル:"))
        self.title_edit = QLineEdit()
        self.title_layout.addWidget(self.title_edit)
        self.right_layout.addLayout(self.title_layout)
        
        # 気分選択
        self.mood_layout = QHBoxLayout()
        self.mood_layout.addWidget(QLabel("今日の気分:"))
        self.mood_combo = QComboBox()
        self.mood_combo.addItems(self.metadata["moods"])
        self.mood_layout.addWidget(self.mood_combo)
        self.right_layout.addLayout(self.mood_layout)
        
        # タグ入力
        self.tag_layout = QHBoxLayout()
        self.tag_layout.addWidget(QLabel("タグ:"))
        self.tag_edit = QLineEdit()
        self.tag_edit.setPlaceholderText("カンマ区切りでタグを入力")
        self.tag_layout.addWidget(self.tag_edit)
        self.right_layout.addLayout(self.tag_layout)
        
        # テキストエディタのツールバー
        self.format_toolbar = QToolBar("書式設定")
        self.format_toolbar.setIconSize(QSize(16, 16))
        
        # フォント選択アクション
        font_action = QAction(QIcon.fromTheme("format-text-bold"), "フォント選択", self)
        font_action.triggered.connect(self.change_font)
        self.format_toolbar.addAction(font_action)
        
        # 太字アクション
        bold_action = QAction(QIcon.fromTheme("format-text-bold"), "太字", self)
        bold_action.triggered.connect(self.format_bold)
        bold_action.setShortcut("Ctrl+B")
        self.format_toolbar.addAction(bold_action)
        
        # 斜体アクション
        italic_action = QAction(QIcon.fromTheme("format-text-italic"), "斜体", self)
        italic_action.triggered.connect(self.format_italic)
        italic_action.setShortcut("Ctrl+I")
        self.format_toolbar.addAction(italic_action)
        
        # 下線アクション
        underline_action = QAction(QIcon.fromTheme("format-text-underline"), "下線", self)
        underline_action.triggered.connect(self.format_underline)
        underline_action.setShortcut("Ctrl+U")
        self.format_toolbar.addAction(underline_action)
        
        # 文字色アクション
        color_action = QAction(QIcon.fromTheme("format-text-color"), "文字色", self)
        color_action.triggered.connect(self.change_text_color)
        self.format_toolbar.addAction(color_action)
        
        # 見出しセレクトボックス
        self.heading_label = QLabel("見出し:")
        self.format_toolbar.addWidget(self.heading_label)
        
        self.heading_combo = QComboBox()
        # シンプルなテキストアイテム
        self.heading_combo.addItem("テキストスタイル")
        self.heading_combo.addItem("通常テキスト")
        self.heading_combo.addItem("見出し 1")
        self.heading_combo.addItem("見出し 2")
        self.heading_combo.addItem("見出し 3")
        
        # 各アイテムのスタイル設定
        self.heading_combo.setItemData(0, "適用するテキストスタイルを選択", Qt.ToolTipRole)
        
        self.heading_combo.setToolTip("テキストを選択してスタイルを適用\nCtrl+1, Ctrl+2, Ctrl+3 で見出しを直接適用")
        self.heading_combo.setMinimumWidth(120)  # 適切な幅に調整
        self.heading_combo.currentIndexChanged.connect(self.apply_heading_from_combo)
        self.format_toolbar.addWidget(self.heading_combo)
        
        # ショートカットキーの設定（メインウィンドウに追加）
        h1_shortcut = QAction("H1見出し", self)
        h1_shortcut.setShortcut("Ctrl+1")
        h1_shortcut.triggered.connect(lambda: self.apply_heading_shortcut(1))
        self.addAction(h1_shortcut)
        
        h2_shortcut = QAction("H2見出し", self)
        h2_shortcut.setShortcut("Ctrl+2")
        h2_shortcut.triggered.connect(lambda: self.apply_heading_shortcut(2))
        self.addAction(h2_shortcut)
        
        h3_shortcut = QAction("H3見出し", self)
        h3_shortcut.setShortcut("Ctrl+3")
        h3_shortcut.triggered.connect(lambda: self.apply_heading_shortcut(3))
        self.addAction(h3_shortcut)
        
        # 画像挿入アクション
        image_action = QAction(QIcon.fromTheme("insert-image"), "画像挿入", self)
        image_action.triggered.connect(self.insert_image)
        image_action.setShortcut("Ctrl+P")
        self.format_toolbar.addAction(image_action)
        
        # 箇条書きアクション
        bullet_action = QAction(QIcon.fromTheme("format-justify-fill"), "箇条書き", self)
        bullet_action.triggered.connect(self.insert_bullet_list)
        self.format_toolbar.addAction(bullet_action)
        
        self.right_layout.addWidget(self.format_toolbar)
        
        # テキストエディタ
        self.text_edit = CustomTextEdit()
        self.text_edit.setFont(self.current_font)
        self.right_layout.addWidget(self.text_edit)
        
        # ボタン配置
        self.button_layout = QHBoxLayout()
        
        self.save_button = QPushButton("保存")
        self.save_button.clicked.connect(self.save_entry)
        self.button_layout.addWidget(self.save_button)
        
        self.delete_button = QPushButton("削除")
        self.delete_button.clicked.connect(self.delete_entry)
        self.button_layout.addWidget(self.delete_button)
        
        self.favorite_button = QPushButton("お気に入り登録")
        self.favorite_button.clicked.connect(self.toggle_favorite)
        self.button_layout.addWidget(self.favorite_button)
        
        self.export_button = QPushButton("エクスポート")
        self.export_button.clicked.connect(self.export_entry)
        self.button_layout.addWidget(self.export_button)
        
        self.right_layout.addLayout(self.button_layout)
        
        # メニューバーの設定
        self.create_menu_bar()
        
        # ステータスバーの設定
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("準備完了", 3000)
        
        # 自動保存タイマー
        self.autosave_timer = QTimer()
        self.autosave_timer.timeout.connect(self.auto_save)
        self.autosave_timer.start(60000)  # 1分ごとに自動保存
        
        # 初期表示
        self.load_entry(self.selected_date)
        
        # テーマ適用
        self.apply_theme()
    
    def create_menu_bar(self):
        menu_bar = self.menuBar()
        
        # ファイルメニュー
        file_menu = menu_bar.addMenu("ファイル")
        
        new_action = QAction("新規", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self.new_entry)
        file_menu.addAction(new_action)
        
        save_action = QAction("保存", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_entry)
        file_menu.addAction(save_action)
        
        export_action = QAction("エクスポート", self)
        export_action.triggered.connect(self.export_entry)
        file_menu.addAction(export_action)
        
        import_action = QAction("インポート", self)
        import_action.triggered.connect(self.import_entry)
        file_menu.addAction(import_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("終了", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # 編集メニュー
        edit_menu = menu_bar.addMenu("編集")
        
        undo_action = QAction("元に戻す", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.triggered.connect(self.text_edit.undo)
        edit_menu.addAction(undo_action)
        
        redo_action = QAction("やり直し", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.triggered.connect(self.text_edit.redo)
        edit_menu.addAction(redo_action)
        
        edit_menu.addSeparator()
        
        cut_action = QAction("切り取り", self)
        cut_action.setShortcut("Ctrl+X")
        cut_action.triggered.connect(self.text_edit.cut)
        edit_menu.addAction(cut_action)
        
        copy_action = QAction("コピー", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.triggered.connect(self.text_edit.copy)
        edit_menu.addAction(copy_action)
        
        paste_action = QAction("貼り付け", self)
        paste_action.setShortcut("Ctrl+V")
        paste_action.triggered.connect(self.text_edit.paste)
        edit_menu.addAction(paste_action)
        
        # 履歴・検索メニュー
        history_menu = menu_bar.addMenu("履歴と検索")
        
        diary_list_action = QAction("日記一覧", self)
        diary_list_action.setShortcut("Ctrl+L")
        diary_list_action.triggered.connect(self.show_diary_list)
        history_menu.addAction(diary_list_action)
        
        advanced_search_action = QAction("詳細検索", self)
        advanced_search_action.setShortcut("Ctrl+F")
        advanced_search_action.triggered.connect(self.show_advanced_search)
        history_menu.addAction(advanced_search_action)
        
        quick_search_action = QAction("クイック検索", self)
        quick_search_action.setShortcut("Ctrl+K")
        quick_search_action.triggered.connect(self.search_entries)
        history_menu.addAction(quick_search_action)
        
        # 表示メニュー
        view_menu = menu_bar.addMenu("表示")
        
        theme_menu = view_menu.addMenu("テーマ")
        
        light_theme_action = QAction("ライト", self)
        light_theme_action.triggered.connect(lambda: self.change_theme("light"))
        theme_menu.addAction(light_theme_action)
        
        dark_theme_action = QAction("ダーク", self)
        dark_theme_action.triggered.connect(lambda: self.change_theme("dark"))
        theme_menu.addAction(dark_theme_action)
        
        # 統計メニュー
        stats_menu = menu_bar.addMenu("統計")
        
        month_stats_action = QAction("月間統計", self)
        month_stats_action.triggered.connect(self.show_month_stats)
        stats_menu.addAction(month_stats_action)
        
        year_stats_action = QAction("年間統計", self)
        year_stats_action.triggered.connect(self.show_year_stats)
        stats_menu.addAction(year_stats_action)
        
        # ヘルプメニュー
        help_menu = menu_bar.addMenu("ヘルプ")
        
        about_action = QAction("このアプリについて", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def update_date_label(self):
        """
        選択された日付のラベルを更新する
        """
        if self.selected_date:
            date_str = self.selected_date.toString('yyyy年MM月dd日(ddd)')
            self.date_label.setText(f"<h2>{date_str}</h2>")
            
            # 選択された日付の日記数を表示
            date_str_iso = self.selected_date.toString('yyyy-MM-dd')
            entry_count = 0
            
            for file_name in os.listdir(self.diary_folder):
                if file_name.startswith(date_str_iso) and file_name.endswith('.json') and file_name != "metadata.json":
                    entry_count += 1
            
            if entry_count > 0:
                self.entry_count_label.setText(f"この日付の日記: {entry_count}件")
                self.entry_count_label.setVisible(True)
            else:
                self.entry_count_label.setVisible(False)
    
    def date_selected(self, date):
        """
        カレンダーで日付が選択された時に呼ばれるメソッド
        選択された日付の日記があれば読み込む
        """
        self.selected_date = date
        
        # 選択された日付の日記ファイルをチェック
        date_str = date.toString('yyyy-MM-dd')
        
        # 同じ日付で複数の日記がある場合は選択ダイアログを表示
        diary_files = []
        for file_name in os.listdir(self.diary_folder):
            if file_name.startswith(date_str) and file_name.endswith('.json') and file_name != "metadata.json":
                try:
                    file_path = os.path.join(self.diary_folder, file_name)
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        title = data.get("title", "無題")
                        diary_files.append({
                            "file_path": file_path,
                            "title": title,
                            "file_name": file_name
                        })
                except:
                    continue
        
        if diary_files:
            if len(diary_files) == 1:
                # 1つしかない場合は直接読み込む
                self.load_entry(diary_files[0]["file_path"])
            else:
                # 複数ある場合は選択ダイアログを表示
                dialog = QDialog(self)
                dialog.setWindowTitle("日記を選択")
                dialog.setMinimumWidth(400)
                
                layout = QVBoxLayout(dialog)
                layout.addWidget(QLabel(f"<b>{date.toString('yyyy年MM月dd日')}の日記が複数あります。</b>"))
                layout.addWidget(QLabel("読み込む日記を選択してください："))
                
                diary_list = QListWidget()
                for diary in diary_files:
                    item = QListWidgetItem(diary["title"])
                    item.setData(Qt.UserRole, diary["file_path"])
                    diary_list.addItem(item)
                
                layout.addWidget(diary_list)
                
                button_layout = QHBoxLayout()
                
                open_button = QPushButton("開く")
                open_button.setDefault(True)
                button_layout.addWidget(open_button)
                
                new_button = QPushButton("新規作成")
                button_layout.addWidget(new_button)
                
                cancel_button = QPushButton("キャンセル")
                button_layout.addWidget(cancel_button)
                
                layout.addLayout(button_layout)
                
                # イベント接続
                open_button.clicked.connect(lambda: self.load_entry(diary_list.currentItem().data(Qt.UserRole)) if diary_list.currentItem() else dialog.reject())
                new_button.clicked.connect(lambda: self.new_entry_with_date(date, dialog))
                cancel_button.clicked.connect(dialog.reject)
                diary_list.itemDoubleClicked.connect(lambda item: self.load_entry(item.data(Qt.UserRole)) and dialog.accept())
                
                dialog.exec_()
        else:
            # 日記が存在しない場合は新しい日記を作成
            self.new_entry()
    
    def load_entry(self, file_path=None):
        """
        指定されたファイルから日記を読み込む
        """
        if not file_path:
            return
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                
                # メタデータを取得
                title = data.get("title", "")
                self.title_edit.setText(title)
                
                # HTMLコンテンツ内の画像パスを絶対パスに変換
                content = data.get("content", "")
                content = self.convert_image_paths_to_absolute(content)
                self.text_edit.setHtml(content)
                
                self.mood_combo.setCurrentText(data.get("mood", "普通"))
                self.tag_edit.setText(", ".join(data.get("tags", [])))
                
                # お気に入りボタンの更新
                file_name = os.path.basename(file_path)
                file_key = file_name[:-5]  # .jsonを除去
                is_favorite = file_key in self.metadata["favorites"]
                self.favorite_button.setText("お気に入り解除" if is_favorite else "お気に入り登録")
                
                # 変更フラグをリセット
                self.text_edit.document().setModified(False)
                
                self.statusBar().showMessage(f"日記を読み込みました: {title}", 5000)
                return True
        except Exception as e:
            self.statusBar().showMessage(f"日記の読み込みに失敗しました: {str(e)}", 5000)
            return False
    
    def new_entry(self):
        """
        編集中の内容を確認した後、新しい日記を作成する
        """
        # 未保存の変更がある場合は確認
        if self.text_edit.document().isModified():
            reply = QMessageBox.question(self,
                                         '確認',
                                         '未保存の変更があります。保存しますか？',
                                         QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            
            if reply == QMessageBox.Yes:
                self.save_entry()
            elif reply == QMessageBox.Cancel:
                return
        
        # 新しい日記を作成
        self.title_edit.clear()
        self.text_edit.clear()
        self.mood_combo.setCurrentText("普通")
        self.tag_edit.clear()
        
        # 変更フラグをリセット
        self.text_edit.document().setModified(False)
        
        self.statusBar().showMessage("新しい日記を作成しました", 5000)
    
    def new_entry_with_date(self, date, dialog=None):
        """
        特定の日付で新しい日記を作成する
        """
        if dialog:
            dialog.accept()
            
        self.new_entry()
        self.calendar.setSelectedDate(date)
    
    def save_entry(self):
        """
        現在の日記を保存する
        タイトルが同じ場合は上書き、異なる場合は新規作成
        """
        # 各情報を取得
        title = self.title_edit.text().strip()
        if not title:
            title = "無題"
            self.title_edit.setText(title)
            
        content = self.text_edit.toHtml()
        
        # HTMLコンテンツ内の画像パスを相対パスに変換
        content = self.convert_image_paths_to_relative(content)
        
        mood = self.mood_combo.currentText()
        
        # タグをリストに変換
        tags_text = self.tag_edit.text().strip()
        tags = []
        if tags_text:
            tags = [tag.strip() for tag in tags_text.split(",") if tag.strip()]
        
        # メタデータに追加
        for tag in tags:
            if tag not in self.metadata["tags"]:
                self.metadata["tags"].append(tag)
        
        # 日付形式の文字列を取得
        date_str = self.selected_date.toString('yyyy-MM-dd')
        
        # スラグの作成（タイトルをURL安全な形式に変換）
        # スペースをハイフンに、特殊文字を削除
        title_slug = re.sub(r'[^\w\s-]', '', title.lower())
        title_slug = re.sub(r'[\s]+', '-', title_slug)
        
        # ファイル名を決定（重複を避けるため）
        base_file_key = f"{date_str}_{title_slug}"
        file_key = base_file_key
        counter = 1
        
        # 既存のファイルをチェック
        existing_file_path = None
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                if file_name.startswith(date_str):
                    file_path = os.path.join(self.diary_folder, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            # 同じタイトルが見つかった場合、上書きする
                            if data.get("title") == title:
                                existing_file_path = file_path
                                file_key = file_name[:-5]  # .jsonを除去
                                break
                    except:
                        continue
        
        # 同じタイトルがない場合かつ同じslugのファイルが存在する場合、連番を付ける
        if not existing_file_path:
            while os.path.exists(os.path.join(self.diary_folder, f"{file_key}.json")):
                file_key = f"{base_file_key}-{counter}"
                counter += 1
                
        # 保存するデータを構築
        data = {
            "title": title,
            "content": content,
            "mood": mood,
            "tags": tags,
            "date": date_str,
            "last_modified": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # ファイルパスを構築
        file_path = os.path.join(self.diary_folder, f"{file_key}.json")
        
        # ファイルに保存
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            # メタデータを更新
            self.save_metadata()
            self.update_tag_list()
            
            # 変更フラグをリセット
            self.text_edit.document().setModified(False)
            
            # お気に入りボタンの更新
            is_favorite = file_key in self.metadata["favorites"]
            self.favorite_button.setText("お気に入り解除" if is_favorite else "お気に入り登録")
            
            self.statusBar().showMessage(f"日記を保存しました: {title}", 5000)
            return True
        except Exception as e:
            self.statusBar().showMessage(f"日記の保存に失敗しました: {str(e)}", 5000)
            return False
    
    def delete_entry(self):
        """
        現在の日記を削除する
        """
        if not self.title_edit.text().strip():
            return
            
        reply = QMessageBox.question(self,
                                     '確認',
                                     '現在の日記を削除してもよろしいですか？',
                                     QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            date_str = self.selected_date.toString('yyyy-MM-dd')
            title = self.title_edit.text().strip()
            
            # 該当タイトルの日記ファイルを検索
            found = False
            for file_name in os.listdir(self.diary_folder):
                if file_name.startswith(date_str) and file_name.endswith('.json') and file_name != "metadata.json":
                    file_path = os.path.join(self.diary_folder, file_name)
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            if data.get("title") == title:
                                # ファイルを削除
                                os.remove(file_path)
                                
                                # お気に入りから削除
                                file_key = file_name[:-5]
                                if file_key in self.metadata["favorites"]:
                                    self.metadata["favorites"].remove(file_key)
                                    self.save_metadata()
                                    self.update_favorites_list()
                                
                                found = True
                                break
                    except:
                        continue
            
            if found:
                self.new_entry()
                self.statusBar().showMessage("日記を削除しました", 5000)
            else:
                self.statusBar().showMessage("削除する日記が見つかりませんでした", 5000)
    
    def toggle_favorite(self):
        """
        お気に入り状態を切り替える
        """
        date_str = self.selected_date.toString('yyyy-MM-dd')
        title = self.title_edit.text().strip()
        
        if not title:
            return
            
        # 該当タイトルの日記ファイルを検索
        for file_name in os.listdir(self.diary_folder):
            if file_name.startswith(date_str) and file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        if data.get("title") == title:
                            file_key = file_name[:-5]  # .jsonを除去
                            
                            if file_key in self.metadata["favorites"]:
                                self.metadata["favorites"].remove(file_key)
                                self.favorite_button.setText("お気に入り登録")
                                self.statusBar().showMessage("お気に入りから削除しました", 5000)
                            else:
                                self.metadata["favorites"].append(file_key)
                                self.favorite_button.setText("お気に入り解除")
                                self.statusBar().showMessage("お気に入りに追加しました", 5000)
                            
                            self.save_metadata()
                            self.update_favorites_list()
                            return
                except:
                    continue
        
        # 保存されていない場合、保存してからお気に入りに追加
        if self.save_entry():
            # 再度トグル処理を行う
            self.toggle_favorite()
    
    def update_calendar_marks(self):
        """
        カレンダーのマークを更新
        """
        # カレンダーウィジェットのupdateCellsメソッドを呼び出す
        # これによりCustomCalendarクラスのupdateCellsが実行される
        self.calendar.updateCells()
    
    def convert_image_paths_to_relative(self, html_content):
        """
        HTML内の画像パスを相対パスに変換する
        
        Args:
            html_content (str): 変換するHTML文字列
            
        Returns:
            str: 変換後のHTML文字列
        """
        import re
        
        # 画像フォルダの絶対パス
        images_abs_path = os.path.abspath(self.images_folder)
        
        # src="[絶対パス]" のパターンを検索
        def replace_path(match):
            path = match.group(1)
            # 絶対パスが画像フォルダ内を指している場合は相対パスに変換
            if os.path.abspath(path).startswith(images_abs_path):
                return f'src="{os.path.relpath(path, os.path.abspath(self.diary_folder))}"'
            return match.group(0)
        
        # 正規表現で置換
        pattern = r'src="([^"]+)"'
        result = re.sub(pattern, replace_path, html_content)
        
        return result
    
    def convert_image_paths_to_absolute(self, html_content):
        """
        HTML内の相対画像パスを絶対パスに変換する
        
        Args:
            html_content (str): 変換するHTML文字列
            
        Returns:
            str: 変換後のHTML文字列
        """
        import re
        
        # 日記フォルダの絶対パス
        diary_abs_path = os.path.abspath(self.diary_folder)
        
        # src="[相対パス]" のパターンを検索
        def replace_path(match):
            path = match.group(1)
            # 相対パスの場合は絶対パスに変換
            if not os.path.isabs(path) and not path.startswith(("http:", "https:")):
                # URIエンコーディングされたパスを戻す場合もある
                import urllib.parse
                path = urllib.parse.unquote(path)
                abs_path = os.path.join(diary_abs_path, path)
                return f'src="{abs_path}"'
            return match.group(0)
        
        # 正規表現で置換
        pattern = r'src="([^"]+)"'
        result = re.sub(pattern, replace_path, html_content)
        
        return result
    
    def auto_save(self):
        if self.text_edit.document().isModified():
            self.save_entry()
            self.status_bar.showMessage("自動保存しました", 2000)
    
    def export_entry(self):
        # 保存ダイアログを表示
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "日記をエクスポート", 
                                                f"diary_{self.selected_date.toString('yyyy-MM-dd')}.html", 
                                                "HTMLファイル (*.html);;テキストファイル (*.txt);;JSONファイル (*.json);;画像付きHTMLフォルダ (*.zip)", 
                                                options=options)
        
        if file_name:
            if file_name.endswith('.html'):
                # HTMLとしてエクスポート
                title = self.title_edit.text()
                content = self.text_edit.toHtml()
                date_str = self.selected_date.toString('yyyy年MM月dd日')
                
                html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{ font-family: "Yu Gothic", sans-serif; margin: 20px; }}
        .diary-date {{ color: #555; }}
        .diary-title {{ font-size: 1.5em; margin: 10px 0; }}
        .diary-content {{ margin-top: 20px; }}
    </style>
</head>
<body>
    <div class="diary-date">{date_str}</div>
    <h1 class="diary-title">{title}</h1>
    <div class="diary-content">{content}</div>
</body>
</html>"""
                
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(html_content)
            
            elif file_name.endswith('.txt'):
                # テキストとしてエクスポート
                title = self.title_edit.text()
                content = self.text_edit.toPlainText()
                date_str = self.selected_date.toString('yyyy年MM月dd日')
                mood = self.mood_combo.currentText()
                tags = self.tag_edit.text()
                
                text_content = f"""日付: {date_str}
タイトル: {title}
気分: {mood}
タグ: {tags}

{content}"""
                
                with open(file_name, 'w', encoding='utf-8') as f:
                    f.write(text_content)
            
            elif file_name.endswith('.json'):
                # JSONとしてエクスポート
                title = self.title_edit.text()
                content = self.text_edit.toHtml()
                plain_content = self.text_edit.toPlainText()
                mood = self.mood_combo.currentText()
                tags = [tag.strip() for tag in self.tag_edit.text().split(",") if tag.strip()]
                
                json_data = {
                    "date": self.selected_date.toString('yyyy-MM-dd'),
                    "title": title,
                    "content": content,
                    "plain_content": plain_content,
                    "mood": mood,
                    "tags": tags,
                    "export_time": datetime.datetime.now().isoformat()
                }
                
                with open(file_name, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, ensure_ascii=False, indent=2)
                    
            elif file_name.endswith('.zip'):
                # 画像付きHTMLフォルダとしてエクスポート
                self.export_with_images(file_name)
            
            self.status_bar.showMessage(f"日記を {file_name} にエクスポートしました", 3000)
    
    def export_with_images(self, zip_file_path):
        """
        日記を画像付きでZIPファイルにエクスポートする
        
        Args:
            zip_file_path (str): 出力するZIPファイルのパス
        """
        import zipfile
        import tempfile
        import re
        
        try:
            # 一時ディレクトリを作成
            with tempfile.TemporaryDirectory() as temp_dir:
                # 画像を保存するディレクトリを作成
                images_dir = os.path.join(temp_dir, "images")
                os.makedirs(images_dir)
                
                # 日記のタイトルと内容を取得
                title = self.title_edit.text()
                content = self.text_edit.toHtml()
                date_str = self.selected_date.toString('yyyy年MM月dd日')
                
                # HTML内の画像パスを抽出
                image_paths = []
                pattern = r'src="([^"]+)"'
                for match in re.finditer(pattern, content):
                    path = match.group(1)
                    # 絶対パスを処理（相対パスは日記フォルダからの相対パス）
                    if not os.path.isabs(path) and not path.startswith(("http:", "https:")):
                        # 相対パスを絶対パスに変換
                        abs_path = os.path.join(os.path.abspath(self.diary_folder), path)
                        if os.path.exists(abs_path):
                            image_paths.append((path, abs_path))
                    elif os.path.exists(path) and not path.startswith(("http:", "https:")):
                        image_paths.append((os.path.basename(path), path))
                
                # 画像をコピーしてHTMLを更新
                for rel_path, abs_path in image_paths:
                    # 画像ファイル名だけを取得
                    img_filename = os.path.basename(rel_path)
                    # 一時ディレクトリにコピー
                    dest_path = os.path.join(images_dir, img_filename)
                    import shutil
                    shutil.copy2(abs_path, dest_path)
                    
                    # HTMLのパスを更新
                    content = content.replace(f'src="{abs_path}"', f'src="images/{img_filename}"')
                    if 'images/' in rel_path:  # 既に相対パスの場合
                        content = content.replace(f'src="{rel_path}"', f'src="images/{img_filename}"')
                
                # HTMLファイルを作成
                html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{ font-family: "Yu Gothic", sans-serif; margin: 20px; }}
        .diary-date {{ color: #555; }}
        .diary-title {{ font-size: 1.5em; margin: 10px 0; }}
        .diary-content {{ margin-top: 20px; }}
        img {{ max-width: 100%; height: auto; }}
    </style>
</head>
<body>
    <div class="diary-date">{date_str}</div>
    <h1 class="diary-title">{title}</h1>
    <div class="diary-content">{content}</div>
</body>
</html>"""
                
                html_path = os.path.join(temp_dir, "index.html")
                with open(html_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                
                # ZIPファイルを作成
                with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    # HTMLファイルを追加
                    zipf.write(html_path, arcname="index.html")
                    
                    # 画像ファイルを追加
                    for root, _, files in os.walk(images_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.join("images", file)
                            zipf.write(file_path, arcname=arcname)
                
                self.status_bar.showMessage(f"画像付き日記を {zip_file_path} にエクスポートしました", 3000)
        
        except Exception as e:
            QMessageBox.warning(self, "エクスポートエラー", f"画像付きエクスポート中にエラーが発生しました: {str(e)}")
    
    def import_entry(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "日記をインポート", "", 
                                                  "JSONファイル (*.json);;HTMLファイル (*.html);;テキストファイル (*.txt)", 
                                                  options=options)
        
        if file_name:
            try:
                if file_name.endswith('.json'):
                    with open(file_name, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    if "date" in data:
                        # 日付の設定
                        qdate = QDate.fromString(data["date"], 'yyyy-MM-dd')
                        self.selected_date = qdate
                        self.calendar.setSelectedDate(qdate)
                        self.update_date_label()
                    
                    # データの設定
                    if "title" in data:
                        self.title_edit.setText(data["title"])
                    if "content" in data:
                        self.text_edit.setHtml(data["content"])
                    if "mood" in data:
                        self.mood_combo.setCurrentText(data["mood"])
                    if "tags" in data:
                        self.tag_edit.setText(", ".join(data["tags"]))
                
                elif file_name.endswith('.html'):
                    with open(file_name, 'r', encoding='utf-8') as f:
                        html_content = f.read()
                    
                    self.text_edit.setHtml(html_content)
                    
                    # タイトルを抽出してみる
                    import re
                    title_match = re.search(r"<title>(.*?)</title>", html_content)
                    if title_match:
                        self.title_edit.setText(title_match.group(1))
                
                elif file_name.endswith('.txt'):
                    with open(file_name, 'r', encoding='utf-8') as f:
                        txt_content = f.read()
                    
                    self.text_edit.setPlainText(txt_content)
                
                self.status_bar.showMessage(f"{file_name} をインポートしました", 3000)
            
            except Exception as e:
                QMessageBox.warning(self, "インポートエラー", f"ファイルのインポート中にエラーが発生しました: {str(e)}")
    
    def update_tag_list(self):
        """
        メタデータからタグリストを更新する
        """
        self.tag_list.clear()
        
        # タグを追加（アルファベット順）
        for tag in sorted(self.metadata["tags"]):
            self.tag_list.addItem(tag)
    
    def update_favorites_list(self):
        self.favorites_list.clear()
        
        for date_title_str in sorted(self.metadata["favorites"], reverse=True):
            # ファイル名を構築
            file_path = os.path.join(self.diary_folder, f"{date_title_str}.json")
            
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    title = data.get("title", "無題")
                    
                    # 日付部分を抽出（yyyy-MM-dd）
                    date_str = date_title_str.split('_')[0]
                    date_obj = QDate.fromString(date_str, 'yyyy-MM-dd')
                    display_date = date_obj.toString('yyyy/MM/dd')
                    
                    item = QListWidgetItem(f"{display_date}: {title}")
                    item.setData(Qt.UserRole, date_title_str)
                    self.favorites_list.addItem(item)
    
    def open_favorite(self, item):
        date_title_str = item.data(Qt.UserRole)
        # 日付部分のみを抽出
        date_str = date_title_str.split('_')[0]
        date = QDate.fromString(date_str, 'yyyy-MM-dd')
        
        # カレンダーの日付を変更
        self.calendar.setSelectedDate(date)
        
        # 特定のタイトルの日記を開く
        file_path = os.path.join(self.diary_folder, f"{date_title_str}.json")
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                title = data.get("title", "")
                
                # load_entry で日記を読み込む
                for file_name in os.listdir(self.diary_folder):
                    if file_name.startswith(date_str) and file_name.endswith('.json') and file_name != "metadata.json":
                        diary_file_path = os.path.join(self.diary_folder, file_name)
                        try:
                            with open(diary_file_path, 'r', encoding='utf-8') as f:
                                diary_data = json.load(f)
                                if diary_data.get("title", "") == title:
                                    # カレンダーを更新した後、ロード処理を手動で行う
                                    self.title_edit.setText(title)
                                    
                                    # HTMLコンテンツ内の画像パスを絶対パスに変換
                                    content = diary_data.get("content", "")
                                    content = self.convert_image_paths_to_absolute(content)
                                    self.text_edit.setHtml(content)
                                    
                                    self.mood_combo.setCurrentText(diary_data.get("mood", "普通"))
                                    self.tag_edit.setText(", ".join(diary_data.get("tags", [])))
                                    
                                    # お気に入りボタンの更新
                                    is_favorite = file_name[:-5] in self.metadata["favorites"]
                                    self.favorite_button.setText("お気に入り解除" if is_favorite else "お気に入り登録")
                                    
                                    # 変更フラグをリセット
                                    self.text_edit.document().setModified(False)
                                    return
                        except:
                            continue

    def filter_by_tag(self, item):
        selected_tag = item.text()
        
        matching_entries = []
        
        # すべての日記ファイルを検索
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    try:
                        data = json.load(f)
                        if selected_tag in data.get("tags", []):
                            # 日付とタイトル情報を保持
                            file_key = file_name[:-5]  # .jsonを除去
                            title = data.get("title", "無題")
                            
                            # 日付部分を抽出
                            date_str = file_key.split('_')[0]
                            date_obj = QDate.fromString(date_str, 'yyyy-MM-dd')
                            
                            matching_entries.append({
                                "file_key": file_key,
                                "title": title,
                                "date": date_obj,
                                "date_str": date_str
                            })
                    except:
                        continue
        
        if matching_entries:
            # 結果表示ダイアログ
            result_dialog = QDialog(self)
            result_dialog.setWindowTitle(f"タグ '{selected_tag}' の検索結果")
            result_dialog.setMinimumWidth(500)
            
            layout = QVBoxLayout(result_dialog)
            
            result_list = QListWidget()
            layout.addWidget(result_list)
            
            # 結果をリストに追加（日付の新しい順）
            for entry in sorted(matching_entries, key=lambda x: x["date"], reverse=True):
                display_date = entry["date"].toString('yyyy/MM/dd')
                title = entry["title"]
                
                item = QListWidgetItem(f"{display_date}: {title}")
                item.setData(Qt.UserRole, {"date_str": entry["date_str"], "title": title})
                result_list.addItem(item)
            
            # アイテムクリック時の処理
            result_list.itemDoubleClicked.connect(lambda item: self.open_tag_search_result(item, result_dialog))
            
            close_button = QPushButton("閉じる")
            close_button.clicked.connect(result_dialog.accept)
            layout.addWidget(close_button)
            
            result_dialog.exec_()
        else:
            QMessageBox.information(self, "検索結果", f"タグ '{selected_tag}' が付いた日記はありません。")
    
    def open_tag_search_result(self, item, dialog):
        data = item.data(Qt.UserRole)
        date_str = data["date_str"]
        title = data["title"]
        date = QDate.fromString(date_str, 'yyyy-MM-dd')
        
        # ダイアログを閉じる
        dialog.accept()
        
        # カレンダーの日付を変更
        self.calendar.setSelectedDate(date)
        
        # 特定のタイトルの日記を開く（date_selectedイベントが発生した後）
        # 少し遅延を入れて確実に日付選択処理が完了してから実行
        QTimer.singleShot(100, lambda: self.select_diary_by_title(title))
    
    def select_diary_by_title(self, title):
        """
        指定されたタイトルの日記を選択して表示する
        """
        date_str = self.selected_date.toString('yyyy-MM-dd')
        
        for file_name in os.listdir(self.diary_folder):
            if file_name.startswith(date_str) and file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        if data.get("title", "") == title:
                            # 日記を読み込む
                            self.title_edit.setText(title)
                            
                            # HTMLコンテンツ内の画像パスを絶対パスに変換
                            content = data.get("content", "")
                            content = self.convert_image_paths_to_absolute(content)
                            self.text_edit.setHtml(content)
                            
                            self.mood_combo.setCurrentText(data.get("mood", "普通"))
                            self.tag_edit.setText(", ".join(data.get("tags", [])))
                            
                            # お気に入りボタンの更新
                            file_key = file_name[:-5]
                            is_favorite = file_key in self.metadata["favorites"]
                            self.favorite_button.setText("お気に入り解除" if is_favorite else "お気に入り登録")
                            
                            # 変更フラグをリセット
                            self.text_edit.document().setModified(False)
                            return
                except:
                    continue
    
    def show_diary_list(self):
        """
        すべての日記一覧をシンプルに表示するダイアログ
        """
        # 日記一覧ダイアログの作成
        diary_list_dialog = QDialog(self)
        diary_list_dialog.setWindowTitle("日記一覧")
        diary_list_dialog.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(diary_list_dialog)
        
        # ラベル
        title_label = QLabel("<h2>日記一覧</h2>")
        layout.addWidget(title_label)
        
        # 検索フィルターレイアウト
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("フィルター:"))
        filter_edit = QLineEdit()
        filter_edit.setPlaceholderText("タイトルや日付でフィルター")
        filter_layout.addWidget(filter_edit)
        layout.addLayout(filter_layout)
        
        # ソート機能
        sort_layout = QHBoxLayout()
        sort_layout.addWidget(QLabel("並べ替え:"))
        sort_combo = QComboBox()
        sort_combo.addItems(["新しい順", "古い順", "タイトル順"])
        sort_layout.addWidget(sort_combo)
        layout.addLayout(sort_layout)
        
        # 日記リスト（シンプルなリスト表示）
        diary_list = QListWidget()
        diary_list.setAlternatingRowColors(True)  # 行の背景色を交互に変える
        diary_list.setSelectionMode(QListWidget.SingleSelection)  # 単一選択モード
        diary_list.setStyleSheet("""
            QListWidget { 
                font-size: 11pt;
            }
            QListWidget::item { 
                padding: 8px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected { 
                background-color: #e6f2ff;
                color: black;
            }
        """)
        layout.addWidget(diary_list)
        
        # 全ての日記ファイルを取得
        diary_entries = []
        
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        file_key = file_name[:-5]  # .jsonを除去
                        
                        # 日付情報を取得
                        date = None
                        date_str = ""
                        jp_date = ""
                        
                        # ファイル名から日付を抽出
                        if '_' in file_key:
                            date_parts = file_key.split('_')[0]
                            date = QDate.fromString(date_parts, 'yyyy-MM-dd')
                            if date.isValid():
                                date_str = date_parts
                                jp_date = date.toString('yyyy年MM月dd日(ddd)')
                        
                        # ファイル名が直接日付形式の場合（旧形式）
                        if not date or not date.isValid():
                            date = QDate.fromString(file_key, 'yyyy-MM-dd')
                            if date.isValid():
                                date_str = file_key
                                jp_date = date.toString('yyyy年MM月dd日(ddd)')
                        
                        # 日付が取得できなかった場合はファイルの更新日時を使用
                        if not date or not date.isValid():
                            if "last_modified" in data:
                                try:
                                    modified_str = data["last_modified"]
                                    modified_date = datetime.datetime.strptime(modified_str, "%Y-%m-%d %H:%M:%S")
                                    date = QDate(modified_date.year, modified_date.month, modified_date.day)
                                    date_str = date.toString('yyyy-MM-dd')
                                    jp_date = date.toString('yyyy年MM月dd日(ddd)')
                                except:
                                    # 現在の日付を使用
                                    date = QDate.currentDate()
                                    date_str = date.toString('yyyy-MM-dd')
                                    jp_date = date.toString('yyyy年MM月dd日(ddd)')
                            else:
                                # 現在の日付を使用
                                date = QDate.currentDate()
                                date_str = date.toString('yyyy-MM-dd')
                                jp_date = date.toString('yyyy年MM月dd日(ddd)')
                        
                        title = data.get("title", "無題")
                        # タイトルが空の場合
                        if title.strip() == "":
                            title = "無題"
                        
                        # 気分、タグなどの情報取得
                        mood = data.get("mood", "")
                        tags = data.get("tags", [])
                        tags_str = ", ".join(tags) if tags else "タグなし"
                        modified = data.get("last_modified", "")
                        
                        # テキスト内容のプレビュー
                        content = data.get("content", "")
                        plain_content = self.html_to_plain(content)
                        preview = plain_content[:100] + "..." if len(plain_content) > 100 else plain_content
                        
                        # 日記エントリを追加
                        diary_entries.append({
                            "date": date,
                            "date_str": date_str,
                            "jp_date": jp_date,
                            "title": title,
                            "mood": mood,
                            "tags": tags,
                            "tags_str": tags_str,
                            "modified": modified,
                            "preview": preview,
                            "file_key": file_key,
                            "file_path": file_path
                        })
                except Exception as e:
                    # 読み込みエラーの場合でも最低限の情報を表示
                    print(f"ファイル読み込みエラー: {file_name} - {str(e)}")
                    file_key = file_name[:-5]
                    date = QDate.currentDate()
                    diary_entries.append({
                        "date": date,
                        "date_str": date.toString('yyyy-MM-dd'),
                        "jp_date": date.toString('yyyy年MM月dd日(ddd)'),
                        "title": f"[読み込みエラー] {file_key}",
                        "mood": "",
                        "tags": [],
                        "tags_str": "タグなし",
                        "modified": "",
                        "preview": "ファイルの読み込みに失敗しました。",
                        "file_key": file_key,
                        "file_path": file_path
                    })
        
        # 日付順に並べ替え（新しい順）
        diary_entries.sort(key=lambda x: x["date"], reverse=True)
        original_entries = diary_entries.copy()  # 元のリストを保持
        
        # エントリ数を表示
        title_label.setText(f"<h2>日記一覧 ({len(diary_entries)}件)</h2>")
        
        # リストに追加する関数
        def update_list_items():
            diary_list.clear()
            for entry in diary_entries:
                # シンプルにタイトルと日付のみ表示
                display_text = f"{entry['title']} - {entry['jp_date']}"
                
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, {"date_str": entry['date_str'], "title": entry['title'], "file_path": entry.get('file_path', '')})
                diary_list.addItem(item)
        
        # 初期表示
        update_list_items()
        
        # フィルタリング機能
        def filter_list():
            filter_text = filter_edit.text().lower()
            if not filter_text:
                # フィルターなしの場合、元のリストを復元
                nonlocal diary_entries
                diary_entries = original_entries.copy()
                sort_entries()  # 現在のソート順を適用
                return
            
            # フィルタリング
            filtered_entries = []
            for entry in original_entries:
                if (filter_text in entry['title'].lower() or 
                    filter_text in entry['jp_date'].lower() or
                    filter_text in entry.get('tags_str', '').lower() or
                    filter_text in entry.get('preview', '').lower()):
                    filtered_entries.append(entry)
            
            diary_entries = filtered_entries
            sort_entries()  # 現在のソート順を適用
            
            # フィルタリング結果の表示
            title_label.setText(f"<h2>日記一覧 ({len(diary_entries)}件 / 全{len(original_entries)}件)</h2>")
        
        # ソート機能
        def sort_entries():
            sort_type = sort_combo.currentText()
            
            if sort_type == "新しい順":
                diary_entries.sort(key=lambda x: x["date"], reverse=True)
            elif sort_type == "古い順":
                diary_entries.sort(key=lambda x: x["date"])
            elif sort_type == "タイトル順":
                diary_entries.sort(key=lambda x: x["title"].lower())
            
            update_list_items()
        
        # イベント接続
        filter_edit.textChanged.connect(filter_list)
        sort_combo.currentIndexChanged.connect(sort_entries)
        
        # アイテムクリック時の処理
        def open_diary(item):
            data = item.data(Qt.UserRole)
            date_str = data["date_str"]
            title = data["title"]
            file_path = data.get("file_path", "")
            
            # ファイルパスが直接指定されている場合はそれを使用
            if file_path and os.path.exists(file_path):
                self.load_entry(file_path)
                diary_list_dialog.accept()
                return
                
            # それ以外の場合は日付とタイトルから検索
            date = QDate.fromString(date_str, 'yyyy-MM-dd')
            
            # カレンダーの日付を変更して特定のタイトルの日記を開く
            self.calendar.setSelectedDate(date)
            diary_list_dialog.accept()
            
            # 少し遅延を入れて確実に日付選択処理が完了してから実行
            QTimer.singleShot(100, lambda: self.select_diary_by_title(title))
        
        diary_list.itemDoubleClicked.connect(open_diary)
        
        # 詳細表示領域
        detail_group = QWidget()
        detail_layout = QVBoxLayout(detail_group)
        detail_label = QLabel("<b>選択した日記の詳細：</b>")
        detail_layout.addWidget(detail_label)
        
        detail_content = QTextEdit()
        detail_content.setReadOnly(True)
        detail_layout.addWidget(detail_content)
        
        # 選択変更時に詳細を表示
        def show_detail():
            if not diary_list.currentItem():
                return
                
            data = diary_list.currentItem().data(Qt.UserRole)
            date_str = data["date_str"]
            title = data["title"]
            
            selected_entry = None
            
            for entry in diary_entries:
                if entry['date_str'] == date_str and entry['title'] == title:
                    selected_entry = entry
                    break
            
            if selected_entry:
                # 日記の詳細を表示
                html = f"""
                <h3>{selected_entry['title']}</h3>
                <p><b>日付:</b> {selected_entry['jp_date']}</p>
                <p><b>気分:</b> {selected_entry['mood']}</p>
                <p><b>タグ:</b> {selected_entry['tags_str']}</p>
                <hr>
                <p>{selected_entry['preview']}</p>
                """
                detail_content.setHtml(html)
        
        diary_list.currentItemChanged.connect(show_detail)
        
        # レイアウト設定（詳細とボタン）
        bottom_layout = QHBoxLayout()
        
        # ボタンレイアウト
        button_layout = QVBoxLayout()
        
        open_button = QPushButton("開く")
        open_button.clicked.connect(lambda: open_diary(diary_list.currentItem()) if diary_list.currentItem() else None)
        button_layout.addWidget(open_button)
        
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(diary_list_dialog.reject)
        button_layout.addWidget(close_button)
        
        bottom_layout.addWidget(detail_group, 2)  # 2:1の比率
        bottom_layout.addLayout(button_layout, 1)
        
        layout.addLayout(bottom_layout)
        
        # ダイアログを表示
        diary_list_dialog.exec_()

    def html_to_plain(self, html_content):
        """
        HTMLコンテンツをプレーンテキストに変換する
        """
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QTextDocument
        
        doc = QTextDocument()
        doc.setHtml(html_content)
        return doc.toPlainText()
    
    def change_font(self):
        """
        フォント選択ダイアログを表示し、選択されたフォントをテキストエディタに適用する
        """
        font, ok = QFontDialog.getFont(self.current_font, self)
        if ok:
            self.current_font = font
            self.text_edit.setFont(font)
            self.statusBar().showMessage(f"フォントを変更しました: {font.family()} {font.pointSize()}pt", 3000)
    
    def format_bold(self):
        """
        選択されたテキストを太字にする
        """
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        if format.fontWeight() == QFont.Bold:
            format.setFontWeight(QFont.Normal)
        else:
            format.setFontWeight(QFont.Bold)
        cursor.mergeCharFormat(format)
        self.text_edit.setTextCursor(cursor)
        self.statusBar().showMessage("スタイルを適用しました", 2000)
    
    def format_italic(self):
        """
        選択されたテキストを斜体にする
        """
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        format.setFontItalic(not format.fontItalic())
        cursor.mergeCharFormat(format)
        self.text_edit.setTextCursor(cursor)
        self.statusBar().showMessage("スタイルを適用しました", 2000)
    
    def format_underline(self):
        """
        選択されたテキストに下線を引く
        """
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        format.setFontUnderline(not format.fontUnderline())
        cursor.mergeCharFormat(format)
        self.text_edit.setTextCursor(cursor)
        self.statusBar().showMessage("スタイルを適用しました", 2000)
    
    def change_text_color(self):
        """
        選択されたテキストの色を変更する
        """
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        current_color = format.foreground().color()
        color = QColorDialog.getColor(current_color, self)
        if color.isValid():
            format.setForeground(color)
            cursor.mergeCharFormat(format)
            self.text_edit.setTextCursor(cursor)
            self.statusBar().showMessage("文字色を変更しました", 2000)
    
    def insert_bullet_list(self):
        """
        箇条書きリストを挿入/解除する
        """
        cursor = self.text_edit.textCursor()
        
        # リスト形式を取得
        list_format = QTextListFormat()
        list_format.setStyle(QTextListFormat.ListDisc)  # ディスク型（黒丸）
        list_format.setIndent(1)  # インデントレベル
        
        # リストを作成または解除
        cursor.createList(list_format)
        self.statusBar().showMessage("箇条書きを適用しました", 2000)
    
    def insert_image(self):
        """
        画像を挿入する
        """
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, "画像を選択", "", 
            "画像ファイル (*.png *.jpg *.jpeg *.bmp *.gif)",
            options=options
        )
        
        if file_name:
            # 日記フォルダー内の画像フォルダにコピー
            import shutil
            import os
            from datetime import datetime
            
            # ファイル名を取得
            base_name = os.path.basename(file_name)
            # タイムスタンプを追加（重複防止）
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            new_filename = f"{timestamp}_{base_name}"
            
            # 保存先パス
            dest_path = os.path.join(self.images_folder, new_filename)
            
            # ファイルをコピー
            try:
                shutil.copy2(file_name, dest_path)
                
                # カーソル位置に画像を挿入
                cursor = self.text_edit.textCursor()
                image_format = QTextImageFormat()
                image_format.setName(dest_path)
                
                # 適切なサイズに調整（大きすぎる画像に対して）
                image = QImage(dest_path)
                max_width = self.text_edit.width() - 50  # マージンを考慮
                if image.width() > max_width:
                    image_format.setWidth(max_width)
                
                cursor.insertImage(image_format)
                self.statusBar().showMessage("画像を挿入しました", 2000)
            except Exception as e:
                QMessageBox.warning(self, "画像挿入エラー", f"画像の挿入中にエラーが発生しました: {str(e)}")
    
    def apply_heading_from_combo(self, index):
        """
        コンボボックスから選択された見出しスタイルを適用する
        """
        if index == 0:  # インデックス0は説明用なので何もしない
            return
        elif index == 1:  # 通常テキスト
            self.apply_normal_text()
        else:
            # 見出しレベル（2→1, 3→2, 4→3）
            level = index - 1
            self.apply_heading(level)
        
        # コンボボックスを最初の項目に戻す
        self.heading_combo.setCurrentIndex(0)
    
    def apply_normal_text(self):
        """
        選択したテキストを通常のテキストスタイルに戻す
        選択がない場合は現在の行全体を選択
        """
        cursor = self.text_edit.textCursor()
        
        # 選択がない場合は現在の行を選択
        if not cursor.hasSelection():
            cursor.movePosition(QTextCursor.StartOfBlock)
            cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            self.text_edit.setTextCursor(cursor)
        
        format = QTextCharFormat()
        format.setFontPointSize(11)  # 通常サイズ
        format.setFontWeight(QFont.Normal)  # 通常の太さ
        
        cursor.mergeCharFormat(format)
        self.text_edit.document().clearUndoRedoStacks()
        self.heading_applied = False
        
        self.statusBar().showMessage("通常テキストを適用しました", 2000)
    
    def apply_heading(self, level):
        """
        選択したテキストに見出しスタイルを適用する
        選択がない場合は現在の行全体を選択
        
        Args:
            level (int): 見出しレベル（1, 2, 3）
        """
        cursor = self.text_edit.textCursor()
        
        # 選択がない場合は現在の行を選択
        if not cursor.hasSelection():
            cursor.movePosition(QTextCursor.StartOfBlock)
            cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            self.text_edit.setTextCursor(cursor)
        
        format = QTextCharFormat()
        
        # 見出しレベルに応じてフォントサイズと太さを設定
        if level == 1:
            format.setFontPointSize(24)  # 大見出し
        elif level == 2:
            format.setFontPointSize(18)  # 中見出し
        elif level == 3:
            format.setFontPointSize(14)  # 小見出し
        
        format.setFontWeight(QFont.Bold)  # 太字
        
        cursor.mergeCharFormat(format)
        self.text_edit.document().clearUndoRedoStacks()
        self.heading_applied = True
        
        self.statusBar().showMessage(f"見出し {level} を適用しました", 2000)
    
    def apply_heading_shortcut(self, level):
        """
        ショートカットキーから見出しスタイルを適用する
        
        Args:
            level (int): 見出しレベル（1, 2, 3）
        """
        self.apply_heading(level)
    
    def search_entries(self):
        """
        クイック検索ダイアログを表示する
        """
        # 検索ダイアログの作成
        search_dialog = QDialog(self)
        search_dialog.setWindowTitle("クイック検索")
        search_dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(search_dialog)
        
        # 検索ボックス
        layout.addWidget(QLabel("検索キーワード:"))
        keyword_edit = QLineEdit()
        keyword_edit.setPlaceholderText("検索キーワードを入力...")
        layout.addWidget(keyword_edit)
        
        # 検索ボタン
        search_button = QPushButton("検索")
        layout.addWidget(search_button)
        
        # 結果リスト
        layout.addWidget(QLabel("検索結果:"))
        result_list = QListWidget()
        layout.addWidget(result_list)
        
        # 閉じるボタン
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(search_dialog.accept)
        layout.addWidget(close_button)
        
        # 検索実行関数
        def perform_search():
            keyword = keyword_edit.text().strip().lower()
            if not keyword:
                return
                
            result_list.clear()
            matching_entries = []
            
            # 全ての日記ファイルを検索
            for file_name in os.listdir(self.diary_folder):
                if file_name.endswith('.json') and file_name != "metadata.json":
                    file_path = os.path.join(self.diary_folder, file_name)
                    
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                            # 日付部分を抽出
                            file_key = file_name[:-5]  # .jsonを除去
                            date_str = ""
                            
                            # 新しいファイル命名形式 (yyyy-MM-dd_title-slug.json)
                            if '_' in file_key:
                                date_str = file_key.split('_')[0]
                            else:
                                # 旧形式 (yyyy-MM-dd.json)
                                date_str = file_key
                                
                            # 有効な日付かチェック
                            date = QDate.fromString(date_str, 'yyyy-MM-dd')
                            if not date.isValid():
                                continue
                                
                            title = data.get("title", "無題")
                            content = self.html_to_plain(data.get("content", ""))
                            tags = ", ".join(data.get("tags", []))
                            
                            # キーワードが含まれているかチェック
                            if (keyword in title.lower() or 
                                keyword in content.lower() or 
                                keyword in tags.lower()):
                                matching_entries.append({
                                    "date": date,
                                    "title": title,
                                    "date_str": date_str
                                })
                    except:
                        continue
            
            # 日付で降順にソート
            matching_entries.sort(key=lambda x: x["date"], reverse=True)
            
            if matching_entries:
                # 結果リストに追加
                for entry in matching_entries:
                    display_date = entry["date"].toString('yyyy/MM/dd')
                    item = QListWidgetItem(f"{display_date}: {entry['title']}")
                    item.setData(Qt.UserRole, {"date_str": entry["date_str"], "title": entry["title"]})
                    result_list.addItem(item)
            else:
                result_list.addItem("検索結果がありません")
        
        # 検索結果アイテムがダブルクリックされた時の処理
        def open_search_result(item):
            data = item.data(Qt.UserRole)
            if not data:  # "検索結果がありません"の場合
                return
                
            date_str = data["date_str"]
            title = data["title"]
            date = QDate.fromString(date_str, 'yyyy-MM-dd')
            
            # ダイアログを閉じる
            search_dialog.accept()
            
            # カレンダーの日付を変更
            self.calendar.setSelectedDate(date)
            
            # 少し遅延を入れて確実に日付選択処理が完了してから実行
            QTimer.singleShot(100, lambda: self.select_diary_by_title(title))
        
        # イベント接続
        search_button.clicked.connect(perform_search)
        keyword_edit.returnPressed.connect(perform_search)
        result_list.itemDoubleClicked.connect(open_search_result)
        
        # ダイアログを表示
        search_dialog.exec_()
    
    def show_advanced_search(self):
        """
        詳細検索ダイアログを表示する
        """
        # 検索ダイアログの作成
        search_dialog = QDialog(self)
        search_dialog.setWindowTitle("詳細検索")
        search_dialog.setMinimumSize(700, 500)
        
        layout = QVBoxLayout(search_dialog)
        
        # 検索条件グループ
        search_group = QWidget()
        search_layout = QGridLayout(search_group)
        
        # キーワード検索
        search_layout.addWidget(QLabel("キーワード:"), 0, 0)
        keyword_edit = QLineEdit()
        keyword_edit.setPlaceholderText("検索キーワードを入力...")
        search_layout.addWidget(keyword_edit, 0, 1, 1, 3)
        
        # 日付範囲
        search_layout.addWidget(QLabel("日付範囲:"), 1, 0)
        date_from = QCalendarWidget()
        date_from.setMaximumHeight(200)
        date_from.setSelectedDate(QDate.currentDate().addMonths(-1))
        search_layout.addWidget(date_from, 1, 1)
        
        search_layout.addWidget(QLabel("～"), 1, 2, Qt.AlignCenter)
        
        date_to = QCalendarWidget()
        date_to.setMaximumHeight(200)
        date_to.setSelectedDate(QDate.currentDate())
        search_layout.addWidget(date_to, 1, 3)
        
        # タグ検索
        search_layout.addWidget(QLabel("タグ:"), 2, 0)
        tag_combo = QComboBox()
        tag_combo.addItem("すべて")
        tag_combo.addItems(sorted(self.metadata["tags"]))
        search_layout.addWidget(tag_combo, 2, 1)
        
        # 気分検索
        search_layout.addWidget(QLabel("気分:"), 2, 2)
        mood_combo = QComboBox()
        mood_combo.addItem("すべて")
        mood_combo.addItems(self.metadata["moods"])
        search_layout.addWidget(mood_combo, 2, 3)
        
        # 検索オプション
        search_layout.addWidget(QLabel("検索オプション:"), 3, 0)
        title_only_check = QCheckBox("タイトルのみ")
        search_layout.addWidget(title_only_check, 3, 1)
        
        case_sensitive_check = QCheckBox("大文字/小文字を区別")
        search_layout.addWidget(case_sensitive_check, 3, 2)
        
        exact_match_check = QCheckBox("完全一致")
        search_layout.addWidget(exact_match_check, 3, 3)
        
        layout.addWidget(search_group)
        
        # 検索ボタン
        search_button = QPushButton("検索")
        layout.addWidget(search_button)
        
        # 結果リスト
        layout.addWidget(QLabel("検索結果:"))
        result_list = QListWidget()
        layout.addWidget(result_list)
        
        # 閉じるボタン
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(search_dialog.accept)
        layout.addWidget(close_button)
        
        # 検索実行関数
        def perform_search():
            # 検索条件を取得
            keyword = keyword_edit.text().strip()
            from_date = date_from.selectedDate()
            to_date = date_to.selectedDate()
            selected_tag = tag_combo.currentText()
            selected_mood = mood_combo.currentText()
            title_only = title_only_check.isChecked()
            case_sensitive = case_sensitive_check.isChecked()
            exact_match = exact_match_check.isChecked()
            
            result_list.clear()
            matching_entries = []
            
            # キーワード処理
            if not case_sensitive and keyword:
                keyword = keyword.lower()
            
            # 全ての日記ファイルを検索
            for file_name in os.listdir(self.diary_folder):
                if file_name.endswith('.json') and file_name != "metadata.json":
                    file_path = os.path.join(self.diary_folder, file_name)
                    
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                            # 日付部分を抽出
                            file_key = file_name[:-5]  # .jsonを除去
                            date_str = ""
                            
                            # 新しいファイル命名形式 (yyyy-MM-dd_title-slug.json)
                            if '_' in file_key:
                                date_str = file_key.split('_')[0]
                            else:
                                # 旧形式 (yyyy-MM-dd.json)
                                date_str = file_key
                            
                            # 有効な日付かチェック
                            date = QDate.fromString(date_str, 'yyyy-MM-dd')
                            if not date.isValid():
                                continue
                            
                            # 日付範囲チェック
                            if date < from_date or date > to_date:
                                continue
                            
                            title = data.get("title", "無題")
                            content = self.html_to_plain(data.get("content", ""))
                            tags = data.get("tags", [])
                            mood = data.get("mood", "")
                            
                            # タグフィルター
                            if selected_tag != "すべて" and selected_tag not in tags:
                                continue
                            
                            # 気分フィルター
                            if selected_mood != "すべて" and selected_mood != mood:
                                continue
                            
                            # キーワード検索
                            if keyword:
                                # 検索対象の文字列
                                if title_only:
                                    search_text = title
                                else:
                                    search_text = f"{title} {content} {' '.join(tags)}"
                                
                                if not case_sensitive:
                                    search_text = search_text.lower()
                                
                                # 検索方法
                                if exact_match:
                                    if keyword not in search_text.split():
                                        continue
                                else:
                                    if keyword not in search_text:
                                        continue
                            
                            # 条件に一致
                            matching_entries.append({
                                "date": date,
                                "title": title,
                                "date_str": date_str
                            })
                    except:
                        continue
            
            # 日付で降順にソート
            matching_entries.sort(key=lambda x: x["date"], reverse=True)
            
            if matching_entries:
                # 結果リストに追加
                for entry in matching_entries:
                    display_date = entry["date"].toString('yyyy/MM/dd')
                    item = QListWidgetItem(f"{display_date}: {entry['title']}")
                    item.setData(Qt.UserRole, {"date_str": entry["date_str"], "title": entry["title"]})
                    result_list.addItem(item)
                
                # 結果数を表示
                result_count = len(matching_entries)
                result_list.insertItem(0, f"-- 検索結果: {result_count}件 --")
            else:
                result_list.addItem("検索結果がありません")
        
        # 検索結果アイテムがダブルクリックされた時の処理
        def open_search_result(item):
            # 最初の項目（結果概要行）の場合は何もしない
            if item.text().startswith("-- 検索結果"):
                return
                
            # "検索結果がありません"の場合も何もしない
            if item.text() == "検索結果がありません":
                return
            
            data = item.data(Qt.UserRole)
            date_str = data["date_str"]
            title = data["title"]
            date = QDate.fromString(date_str, 'yyyy-MM-dd')
            
            # ダイアログを閉じる
            search_dialog.accept()
            
            # カレンダーの日付を変更
            self.calendar.setSelectedDate(date)
            
            # 少し遅延を入れて確実に日付選択処理が完了してから実行
            QTimer.singleShot(100, lambda: self.select_diary_by_title(title))
        
        # イベント接続
        search_button.clicked.connect(perform_search)
        result_list.itemDoubleClicked.connect(open_search_result)
        
        # ダイアログを表示
        search_dialog.exec_()

    def show_month_stats(self):
        """
        月間統計を表示する
        """
        current_month = self.calendar.selectedDate().toString('yyyy-MM')
        
        # データ収集
        diary_count = 0
        mood_counts = {}
        tag_counts = {}
        
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                
                try:
                    # 日付チェック
                    file_key = file_name[:-5]  # .jsonを除去
                    date_str = ""
                    
                    # 新しいファイル命名形式 (yyyy-MM-dd_title-slug.json)
                    if '_' in file_key:
                        date_str = file_key.split('_')[0]
                    else:
                        # 旧形式 (yyyy-MM-dd.json)
                        date_str = file_key
                    
                    # 月が一致するかチェック
                    if not date_str.startswith(current_month):
                        continue
                    
                    # ファイルを読み込む
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        
                        # 日記数をカウント
                        diary_count += 1
                        
                        # 気分をカウント
                        mood = data.get("mood", "不明")
                        if mood in mood_counts:
                            mood_counts[mood] += 1
                        else:
                            mood_counts[mood] = 1
                        
                        # タグをカウント
                        for tag in data.get("tags", []):
                            if tag in tag_counts:
                                tag_counts[tag] += 1
                            else:
                                tag_counts[tag] = 1
                except:
                    continue
        
        # 統計ダイアログの作成
        stats_dialog = QDialog(self)
        stats_dialog.setWindowTitle("月間統計")
        stats_dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(stats_dialog)
        
        # 期間表示
        month_label = QLabel(f"<h2>{current_month}の統計</h2>")
        layout.addWidget(month_label)
        
        # 日記数
        diary_count_label = QLabel(f"日記数: {diary_count}件")
        layout.addWidget(diary_count_label)
        
        # 気分グラフ
        if mood_counts:
            layout.addWidget(QLabel("<h3>気分の分布</h3>"))
            
            mood_list = QListWidget()
            for mood, count in sorted(mood_counts.items(), key=lambda x: x[1], reverse=True):
                item = QListWidgetItem(f"{mood}: {count}件")
                mood_list.addItem(item)
            
            layout.addWidget(mood_list)
        
        # タググラフ
        if tag_counts:
            layout.addWidget(QLabel("<h3>よく使われたタグ</h3>"))
            
            tag_list = QListWidget()
            for tag, count in sorted(tag_counts.items(), key=lambda x: x[1], reverse=True)[:10]:  # 上位10件
                item = QListWidgetItem(f"{tag}: {count}件")
                tag_list.addItem(item)
            
            layout.addWidget(tag_list)
        
        # 閉じるボタン
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(stats_dialog.accept)
        layout.addWidget(close_button)
        
        # ダイアログを表示
        stats_dialog.exec_()
    
    def show_year_stats(self):
        """
        年間統計を表示する
        """
        current_year = self.calendar.selectedDate().toString('yyyy')
        
        # データ収集
        diary_count = 0
        month_counts = {f"{current_year}-{str(month).zfill(2)}": 0 for month in range(1, 13)}
        mood_counts = {}
        tag_counts = {}
        
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                
                try:
                    # 日付チェック
                    file_key = file_name[:-5]  # .jsonを除去
                    date_str = ""
                    
                    # 新しいファイル命名形式 (yyyy-MM-dd_title-slug.json)
                    if '_' in file_key:
                        date_str = file_key.split('_')[0]
                    else:
                        # 旧形式 (yyyy-MM-dd.json)
                        date_str = file_key
                    
                    # 年が一致するかチェック
                    if not date_str.startswith(current_year):
                        continue
                    
                    # ファイルを読み込む
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        
                        # 日記数をカウント
                        diary_count += 1
                        
                        # 月別カウント
                        month_str = date_str[:7]  # yyyy-MM
                        if month_str in month_counts:
                            month_counts[month_str] += 1
                        
                        # 気分をカウント
                        mood = data.get("mood", "不明")
                        if mood in mood_counts:
                            mood_counts[mood] += 1
                        else:
                            mood_counts[mood] = 1
                        
                        # タグをカウント
                        for tag in data.get("tags", []):
                            if tag in tag_counts:
                                tag_counts[tag] += 1
                            else:
                                tag_counts[tag] = 1
                except:
                    continue
        
        # 統計ダイアログの作成
        stats_dialog = QDialog(self)
        stats_dialog.setWindowTitle("年間統計")
        stats_dialog.setMinimumWidth(500)
        
        layout = QVBoxLayout(stats_dialog)
        
        # 期間表示
        year_label = QLabel(f"<h2>{current_year}年の統計</h2>")
        layout.addWidget(year_label)
        
        # 日記数
        diary_count_label = QLabel(f"日記数: {diary_count}件")
        layout.addWidget(diary_count_label)
        
        # 月別グラフ
        layout.addWidget(QLabel("<h3>月別の日記数</h3>"))
        
        month_list = QListWidget()
        for month_str, count in sorted(month_counts.items()):
            # 日本語の月名に変換
            month = int(month_str.split('-')[1])
            month_name = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"][month-1]
            
            item = QListWidgetItem(f"{month_name}: {count}件")
            month_list.addItem(item)
        
        layout.addWidget(month_list)
        
        # 気分グラフ
        if mood_counts:
            layout.addWidget(QLabel("<h3>気分の分布</h3>"))
            
            mood_list = QListWidget()
            for mood, count in sorted(mood_counts.items(), key=lambda x: x[1], reverse=True):
                item = QListWidgetItem(f"{mood}: {count}件")
                mood_list.addItem(item)
            
            layout.addWidget(mood_list)
        
        # タググラフ
        if tag_counts:
            layout.addWidget(QLabel("<h3>よく使われたタグ</h3>"))
            
            tag_list = QListWidget()
            for tag, count in sorted(tag_counts.items(), key=lambda x: x[1], reverse=True)[:10]:  # 上位10件
                item = QListWidgetItem(f"{tag}: {count}件")
                tag_list.addItem(item)
            
            layout.addWidget(tag_list)
        
        # 閉じるボタン
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(stats_dialog.accept)
        layout.addWidget(close_button)
        
        # ダイアログを表示
        stats_dialog.exec_()

    def show_about(self):
        """
        アプリの情報を表示する
        """
        QMessageBox.about(self, "このアプリについて",
                      """<h1>PyQt5 日記アプリ</h1>
<p>PyQt5で作成されたシンプルな日記アプリケーションです。</p>
<p>機能：</p>
<ul>
    <li>日記の作成、保存、編集</li>
    <li>日付ごとの管理</li>
    <li>タグ付け</li>
    <li>お気に入り機能</li>
    <li>カレンダー表示</li>
    <li>テキスト書式設定</li>
    <li>画像挿入</li>
    <li>エクスポート/インポート</li>
    <li>検索機能</li>
    <li>統計表示</li>
</ul>
<p>バージョン: 1.0</p>
<p>© 2023 All Rights Reserved</p>""")
    
    def save_metadata(self):
        """
        メタデータをJSONファイルに保存する
        """
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(self.metadata, f, ensure_ascii=False, indent=4)
    
    def change_theme(self, theme):
        """
        アプリケーションのテーマを変更する
        """
        self.metadata["theme"] = theme
        self.save_metadata()
        self.apply_theme()
    
    def apply_theme(self):
        """
        選択されたテーマを適用する
        """
        if self.metadata["theme"] == "dark":
            # ダークテーマ
            app = QApplication.instance()
            app.setStyleSheet("""
                QMainWindow, QDialog {
                    background-color: #2D2D30;
                    color: #F1F1F1;
                }
                QWidget {
                    background-color: #2D2D30;
                    color: #F1F1F1;
                }
                QTextEdit, QLineEdit, QComboBox, QCalendarWidget {
                    background-color: #1E1E1E;
                    color: #F1F1F1;
                    border: 1px solid #3F3F46;
                }
                QPushButton {
                    background-color: #0E639C;
                    color: #FFFFFF;
                    border: none;
                    padding: 5px;
                    min-height: 25px;
                }
                QPushButton:hover {
                    background-color: #1177BB;
                }
                QPushButton:pressed {
                    background-color: #0D5487;
                }
                QHeaderView::section {
                    background-color: #2D2D30;
                    color: #F1F1F1;
                }
                QListWidget, QListWidget::item {
                    background-color: #1E1E1E;
                    color: #F1F1F1;
                }
                QListWidget::item:selected {
                    background-color: #264F78;
                    color: #FFFFFF;
                }
                QToolBar {
                    background-color: #2D2D30;
                    border: none;
                }
            """)
        else:
            # ライトテーマ
            app = QApplication.instance()
            app.setStyleSheet("")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DiaryApp()
    window.show()
    sys.exit(app.exec_())