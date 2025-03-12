import sys
import os
import json
import datetime
import calendar
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QTextEdit, QPushButton, QLabel, QCalendarWidget, QComboBox, 
                            QLineEdit, QMessageBox, QTabWidget, QGridLayout, QListWidget,
                            QListWidgetItem, QFileDialog, QColorDialog, QFontDialog, QMenu,
                            QAction, QToolBar, QStatusBar, QSplitter, QDialog, QCheckBox)
from PyQt5.QtGui import QFont, QIcon, QTextCharFormat, QColor, QTextCursor, QTextListFormat, QTextBlockFormat, QImage, QTextImageFormat
from PyQt5.QtCore import Qt, QDate, QTimer, QSize, QUrl
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent

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
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.setMinimumWidth(300)
        self.calendar.clicked.connect(self.date_selected)
        
        # カレンダーに日記のある日をマーク
        self.update_calendar_marks()
        
        self.left_layout.addWidget(QLabel("カレンダー"))
        self.left_layout.addWidget(self.calendar)
        
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
        day_names = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"]
        day_of_week = day_names[self.selected_date.dayOfWeek() - 1]
        date_str = f"{self.selected_date.year()}年 {self.selected_date.month()}月 {self.selected_date.day()}日 ({day_of_week})"
        self.date_label.setText(f"<h2>{date_str}</h2>")
    
    def date_selected(self, date):
        # 保存確認
        if self.text_edit.document().isModified():
            reply = QMessageBox.question(self, '確認', 
                                        '変更を保存しますか？',
                                        QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            
            if reply == QMessageBox.Save:
                self.save_entry()
            elif reply == QMessageBox.Cancel:
                # キャンセルの場合は日付選択を無効にする
                self.calendar.setSelectedDate(self.selected_date)
                return
        
        self.selected_date = date
        self.update_date_label()
        self.load_entry(date)
    
    def load_entry(self, date):
        # 日付に対応するファイル名
        file_name = f"{date.toString('yyyy-MM-dd')}.json"
        file_path = os.path.join(self.diary_folder, file_name)
        
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.title_edit.setText(data.get("title", ""))
                
                # HTMLコンテンツ内の画像パスを絶対パスに変換
                content = data.get("content", "")
                content = self.convert_image_paths_to_absolute(content)
                self.text_edit.setHtml(content)
                
                self.mood_combo.setCurrentText(data.get("mood", "普通"))
                self.tag_edit.setText(", ".join(data.get("tags", [])))
                
                # お気に入りボタンの更新
                is_favorite = file_name[:-5] in self.metadata["favorites"]
                self.favorite_button.setText("お気に入り解除" if is_favorite else "お気に入り登録")
        else:
            self.title_edit.clear()
            self.text_edit.clear()
            self.mood_combo.setCurrentIndex(0)
            self.tag_edit.clear()
            self.favorite_button.setText("お気に入り登録")
        
        # 変更フラグをリセット
        self.text_edit.document().setModified(False)
    
    def save_entry(self):
        # 現在の入力を保存
        title = self.title_edit.text()
        content = self.text_edit.toHtml()
        
        # HTMLコンテンツ内の画像パスを相対パスに変換
        content = self.convert_image_paths_to_relative(content)
        
        mood = self.mood_combo.currentText()
        tags = [tag.strip() for tag in self.tag_edit.text().split(",") if tag.strip()]
        
        # 日付に対応するファイル名
        file_name = f"{self.selected_date.toString('yyyy-MM-dd')}.json"
        file_path = os.path.join(self.diary_folder, file_name)
        
        # データの準備
        data = {
            "title": title,
            "content": content,
            "mood": mood,
            "tags": tags,
            "last_modified": datetime.datetime.now().isoformat()
        }
        
        # ファイルに保存
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # メタデータの更新
        for tag in tags:
            if tag not in self.metadata["tags"]:
                self.metadata["tags"].append(tag)
        
        self.save_metadata()
        self.update_tag_list()
        self.update_calendar_marks()
        
        # 変更フラグをリセット
        self.text_edit.document().setModified(False)
        
        # ステータスバーに保存メッセージを表示
        self.status_bar.showMessage(f"{self.selected_date.toString('yyyy年MM月dd日')}の日記を保存しました", 3000)
    
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
    
    def delete_entry(self):
        reply = QMessageBox.question(self, '確認', 
                                    'この日の日記を削除してもよろしいですか？',
                                    QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # 日付に対応するファイル名
            file_name = f"{self.selected_date.toString('yyyy-MM-dd')}.json"
            file_path = os.path.join(self.diary_folder, file_name)
            
            if os.path.exists(file_path):
                os.remove(file_path)
                
                # お気に入りから削除
                date_str = self.selected_date.toString('yyyy-MM-dd')
                if date_str in self.metadata["favorites"]:
                    self.metadata["favorites"].remove(date_str)
                    self.save_metadata()
                    self.update_favorites_list()
                
                # カレンダーマークを更新
                self.update_calendar_marks()
                
                # エディタをクリア
                self.title_edit.clear()
                self.text_edit.clear()
                self.mood_combo.setCurrentIndex(0)
                self.tag_edit.clear()
                
                # ステータスバーにメッセージを表示
                self.status_bar.showMessage(f"{self.selected_date.toString('yyyy年MM月dd日')}の日記を削除しました", 3000)
    
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
    
    def new_entry(self):
        # 保存確認
        if self.text_edit.document().isModified():
            reply = QMessageBox.question(self, '確認', 
                                        '変更を保存しますか？',
                                        QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            
            if reply == QMessageBox.Save:
                self.save_entry()
            elif reply == QMessageBox.Cancel:
                return
        
        # 今日の日付を選択
        today = QDate.currentDate()
        self.selected_date = today
        self.calendar.setSelectedDate(today)
        self.update_date_label()
        
        # エディタをクリア
        self.title_edit.clear()
        self.text_edit.clear()
        self.mood_combo.setCurrentIndex(0)
        self.tag_edit.clear()
        
        # 変更フラグをリセット
        self.text_edit.document().setModified(False)
    
    def toggle_favorite(self):
        date_str = self.selected_date.toString('yyyy-MM-dd')
        
        if date_str in self.metadata["favorites"]:
            self.metadata["favorites"].remove(date_str)
            self.favorite_button.setText("お気に入り登録")
            self.status_bar.showMessage("お気に入りから削除しました", 2000)
        else:
            # まず日記が存在するか確認
            file_name = f"{date_str}.json"
            file_path = os.path.join(self.diary_folder, file_name)
            
            if not os.path.exists(file_path):
                # 存在しない場合は保存
                self.save_entry()
            
            self.metadata["favorites"].append(date_str)
            self.favorite_button.setText("お気に入り解除")
            self.status_bar.showMessage("お気に入りに追加しました", 2000)
        
        self.save_metadata()
        self.update_favorites_list()
    
    def update_favorites_list(self):
        self.favorites_list.clear()
        
        for date_str in sorted(self.metadata["favorites"], reverse=True):
            file_path = os.path.join(self.diary_folder, f"{date_str}.json")
            
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    title = data.get("title", "無題")
                    
                    date_obj = QDate.fromString(date_str, 'yyyy-MM-dd')
                    display_date = date_obj.toString('yyyy/MM/dd')
                    
                    item = QListWidgetItem(f"{display_date}: {title}")
                    item.setData(Qt.UserRole, date_str)
                    self.favorites_list.addItem(item)
    
    def open_favorite(self, item):
        date_str = item.data(Qt.UserRole)
        date = QDate.fromString(date_str, 'yyyy-MM-dd')
        
        # カレンダーの日付を変更
        self.calendar.setSelectedDate(date)
        
        # 日記を読み込む（date_selectedイベントが発生する）
    
    def update_tag_list(self):
        self.tag_list.clear()
        
        for tag in sorted(self.metadata["tags"]):
            self.tag_list.addItem(tag)
    
    def filter_by_tag(self, item):
        selected_tag = item.text()
        
        matching_dates = []
        
        # すべての日記ファイルを検索
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    try:
                        data = json.load(f)
                        if selected_tag in data.get("tags", []):
                            date_str = file_name[:-5]  # .jsonを除去
                            matching_dates.append(date_str)
                    except:
                        continue
        
        if matching_dates:
            # 結果表示ダイアログ
            result_dialog = QDialog(self)
            result_dialog.setWindowTitle(f"タグ '{selected_tag}' の検索結果")
            result_dialog.setMinimumWidth(400)
            
            layout = QVBoxLayout(result_dialog)
            
            result_list = QListWidget()
            layout.addWidget(result_list)
            
            # 結果をリストに追加
            for date_str in sorted(matching_dates, reverse=True):
                file_path = os.path.join(self.diary_folder, f"{date_str}.json")
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    title = data.get("title", "無題")
                    
                    date_obj = QDate.fromString(date_str, 'yyyy-MM-dd')
                    display_date = date_obj.toString('yyyy/MM/dd')
                    
                    item = QListWidgetItem(f"{display_date}: {title}")
                    item.setData(Qt.UserRole, date_str)
                    result_list.addItem(item)
            
            # アイテムクリック時の処理
            result_list.itemDoubleClicked.connect(lambda item: self.open_search_result(item, result_dialog))
            
            close_button = QPushButton("閉じる")
            close_button.clicked.connect(result_dialog.accept)
            layout.addWidget(close_button)
            
            result_dialog.exec_()
        else:
            QMessageBox.information(self, "検索結果", f"タグ '{selected_tag}' が付いた日記はありません。")
    
    def open_search_result(self, item, dialog):
        date_str = item.data(Qt.UserRole)
        date = QDate.fromString(date_str, 'yyyy-MM-dd')
        
        # ダイアログを閉じる
        dialog.accept()
        
        # カレンダーの日付を変更
        self.calendar.setSelectedDate(date)
    
    def search_entries(self):
        # 検索ダイアログ
        search_dialog = QDialog(self)
        search_dialog.setWindowTitle("日記を検索")
        search_dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(search_dialog)
        
        layout.addWidget(QLabel("検索キーワード:"))
        search_edit = QLineEdit()
        layout.addWidget(search_edit)
        
        button_layout = QHBoxLayout()
        search_button = QPushButton("検索")
        cancel_button = QPushButton("キャンセル")
        button_layout.addWidget(search_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        result_list = QListWidget()
        layout.addWidget(result_list)
        result_list.hide()
        
        # ボタンのイベント
        cancel_button.clicked.connect(search_dialog.reject)
        search_button.clicked.connect(lambda: self.perform_search(search_edit.text(), result_list))
        
        # アイテムクリック時の処理
        result_list.itemDoubleClicked.connect(lambda item: self.open_search_result(item, search_dialog))
        
        search_dialog.exec_()
    
    def perform_search(self, keyword, result_list):
        if not keyword:
            return
        
        keyword = keyword.lower()
        matching_entries = []
        
        # すべての日記ファイルを検索
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                file_path = os.path.join(self.diary_folder, file_name)
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    try:
                        data = json.load(f)
                        title = data.get("title", "").lower()
                        content = data.get("content", "").lower()
                        plain_content = self.html_to_plain(content).lower()
                        
                        if (keyword in title or 
                            keyword in plain_content or 
                            any(keyword in tag.lower() for tag in data.get("tags", []))):
                            
                            date_str = file_name[:-5]  # .jsonを除去
                            matching_entries.append((date_str, data.get("title", "無題")))
                    except:
                        continue
        
        # 結果を表示
        result_list.clear()
        
        if matching_entries:
            result_list.show()
            
            for date_str, title in sorted(matching_entries, key=lambda x: x[0], reverse=True):
                date_obj = QDate.fromString(date_str, 'yyyy-MM-dd')
                display_date = date_obj.toString('yyyy/MM/dd')
                
                item = QListWidgetItem(f"{display_date}: {title}")
                item.setData(Qt.UserRole, date_str)
                result_list.addItem(item)
        else:
            QMessageBox.information(self, "検索結果", f"'{keyword}' に一致する日記はありません。")
    
    def html_to_plain(self, html):
        """
        HTMLからプレーンテキストに変換する
        
        Args:
            html (str): 変換するHTML文字列
            
        Returns:
            str: 変換後のプレーンテキスト
        """
        import re
        # タグを削除
        text = re.sub(r'<[^>]+>', '', html)
        # HTML特殊文字をデコード
        text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"').replace('&apos;', "'")
        return text
    
    def update_calendar_marks(self):
        # カレンダーのマークをリセット
        self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
        
        # 日記のある日にマークを付ける
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                date_str = file_name[:-5]  # .jsonを除去
                date = QDate.fromString(date_str, 'yyyy-MM-dd')
                
                if date.isValid():
                    format = QTextCharFormat()
                    format.setBackground(QColor(173, 216, 230))  # 薄い青色
                    self.calendar.setDateTextFormat(date, format)
        
        # お気に入りの日にはさらに目立つマークを付ける
        for date_str in self.metadata["favorites"]:
            date = QDate.fromString(date_str, 'yyyy-MM-dd')
            
            if date.isValid():
                format = QTextCharFormat()
                format.setBackground(QColor(255, 182, 193))  # 薄いピンク色
                self.calendar.setDateTextFormat(date, format)
    
    def save_metadata(self):
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(self.metadata, f, ensure_ascii=False, indent=2)
    
    def format_bold(self):
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        format.setFontWeight(QFont.Bold if format.fontWeight() != QFont.Bold else QFont.Normal)
        cursor.mergeCharFormat(format)
        self.text_edit.setTextCursor(cursor)
    
    def format_italic(self):
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        format.setFontItalic(not format.fontItalic())
        cursor.mergeCharFormat(format)
        self.text_edit.setTextCursor(cursor)
    
    def format_underline(self):
        cursor = self.text_edit.textCursor()
        format = cursor.charFormat()
        format.setFontUnderline(not format.fontUnderline())
        cursor.mergeCharFormat(format)
        self.text_edit.setTextCursor(cursor)
    
    def change_font(self):
        current = self.text_edit.currentFont()
        font, ok = QFontDialog.getFont(current, self)
        if ok:
            self.current_font = font
            self.text_edit.setCurrentFont(font)
    
    def change_text_color(self):
        cursor = self.text_edit.textCursor()
        current_color = cursor.charFormat().foreground().color()
        color = QColorDialog.getColor(current_color, self)
        if color.isValid():
            format = cursor.charFormat()
            format.setForeground(color)
            cursor.mergeCharFormat(format)
            self.text_edit.setTextCursor(cursor)
    
    def insert_bullet_list(self):
        cursor = self.text_edit.textCursor()
        
        # リストフォーマットを作成
        list_format = QTextListFormat()
        list_format.setStyle(QTextListFormat.ListDisc)
        list_format.setIndent(1)
        
        cursor.createList(list_format)
        self.text_edit.setTextCursor(cursor)
    
    def apply_heading_shortcut(self, level):
        """
        ショートカットキーで見出しスタイルを適用する
        
        Args:
            level (int): 見出しレベル（1=H1, 2=H2, 3=H3）
        """
        # 見出しを適用（選択範囲の確認はapply_headingメソッド内で行う）
        self.apply_heading(level)
    
    def apply_heading_from_combo(self, index):
        """
        セレクトボックスから選択した見出しスタイルを適用する
        
        Args:
            index (int): コンボボックスのインデックス（0=選択肢, 1=通常, 2=H1, 3=H2, 4=H3）
        """
        if index == 0:  # 「スタイルを選択」の場合は何もしない
            return
            
        cursor = self.text_edit.textCursor()
        
        # 選択範囲がない場合は、現在の行を選択
        if not cursor.hasSelection():
            cursor.movePosition(QTextCursor.StartOfBlock)
            cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            self.text_edit.setTextCursor(cursor)
        
        if index == 1:  # 通常テキスト
            self.apply_normal_text()
        elif index >= 2:  # 見出し（インデックスを調整）
            self.apply_heading(index - 1)
            
        # コンボボックスを初期状態に戻す
        self.heading_combo.setCurrentIndex(0)
    
    def apply_normal_text(self):
        """
        選択したテキストに通常のスタイルを適用する
        選択範囲がない場合は、現在の行を選択
        """
        cursor = self.text_edit.textCursor()
        
        # 選択範囲がない場合は、現在の行を選択
        if not cursor.hasSelection():
            cursor.movePosition(QTextCursor.StartOfBlock)
            cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            self.text_edit.setTextCursor(cursor)
        
        # 通常のテキスト書式
        char_format = QTextCharFormat()
        char_format.setFontPointSize(11)
        char_format.setFontWeight(QFont.Normal)
        
        # 選択範囲に適用
        cursor.beginEditBlock()
        cursor.mergeCharFormat(char_format)
        cursor.endEditBlock()
        
        # カーソル位置を維持
        position = cursor.position()
        cursor.clearSelection()
        cursor.setPosition(position)
        self.text_edit.setTextCursor(cursor)
        
        # 見出しフラグをリセット
        self.text_edit.heading_applied = False
        
        # ステータスバーに通知
        self.status_bar.showMessage("通常のテキストスタイルを適用しました", 2000)
    
    def apply_heading(self, level):
        """
        選択したテキストに見出しスタイルを適用する
        選択範囲がない場合は、現在の行を選択
        
        Args:
            level (int): 見出しレベル（1=H1, 2=H2, 3=H3）
        """
        cursor = self.text_edit.textCursor()
        
        # 選択範囲がない場合は、現在の行を選択
        if not cursor.hasSelection():
            cursor.movePosition(QTextCursor.StartOfBlock)
            cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)
            self.text_edit.setTextCursor(cursor)
            cursor = self.text_edit.textCursor()  # 更新したカーソルを取得
        
        # 文字書式を作成
        char_format = QTextCharFormat()
        
        # 見出しレベルに基づいてフォントサイズを設定
        if level == 1:
            char_format.setFontPointSize(24)
            char_format.setFontWeight(QFont.Bold)
        elif level == 2:
            char_format.setFontPointSize(18)
            char_format.setFontWeight(QFont.Bold)
        elif level == 3:
            char_format.setFontPointSize(14)
            char_format.setFontWeight(QFont.Bold)
        
        # 選択範囲に見出しスタイルを適用
        cursor.beginEditBlock()
        cursor.mergeCharFormat(char_format)
        cursor.endEditBlock()
        
        # カーソル位置を維持しながら選択を解除
        position = cursor.position()
        cursor.clearSelection()
        cursor.setPosition(position)
        self.text_edit.setTextCursor(cursor)
        
        # 見出しフラグを設定
        self.text_edit.heading_applied = True
        
        # ステータスバーに通知
        self.status_bar.showMessage(f"見出し{level}を適用しました", 2000)
    
    def insert_image(self):
        """
        画像をテキストエディタに挿入する
        """
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self, "画像を選択", "", 
            "画像ファイル (*.png *.jpg *.jpeg *.bmp *.gif);;すべてのファイル (*)", 
            options=options
        )
        
        if file_name:
            cursor = self.text_edit.textCursor()
            
            # 画像を読み込む
            image = QImage(file_name)
            
            if image.isNull():
                QMessageBox.warning(self, "画像挿入エラー", "選択された画像を読み込めませんでした。")
                return
            
            # 画像が大きすぎる場合はリサイズする
            max_width = 600
            if image.width() > max_width:
                image = image.scaledToWidth(max_width, Qt.SmoothTransformation)
            
            # 画像ファイルを日記アプリの画像フォルダにコピーする
            from pathlib import Path
            import shutil
            
            # オリジナルファイル名を取得
            original_filename = os.path.basename(file_name)
            # 現在の日付とタイムスタンプを加えて一意なファイル名を作成
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            new_filename = f"{timestamp}_{original_filename}"
            new_filepath = os.path.join(self.images_folder, new_filename)
            
            # 画像を新しい場所にコピー
            try:
                # 画像が大きい場合は新しいサイズで保存
                if image.width() > max_width:
                    image.save(new_filepath)
                else:
                    shutil.copy2(file_name, new_filepath)
                
                # 画像フォーマットを作成（相対パスを使用）
                image_format = QTextImageFormat()
                image_format.setName(new_filepath)  # 保存済み画像へのパスを設定
                image_format.setWidth(image.width())
                image_format.setHeight(image.height())
                
                # 画像を挿入
                cursor.insertImage(image_format)
                
                # ステータスバーに通知
                self.status_bar.showMessage("画像を挿入しました", 2000)
            except Exception as e:
                QMessageBox.warning(self, "画像挿入エラー", f"画像の保存中にエラーが発生しました: {str(e)}")
    
    def change_theme(self, theme):
        self.metadata["theme"] = theme
        self.save_metadata()
        self.apply_theme()
    
    def apply_theme(self):
        if self.metadata["theme"] == "dark":
            # ダークテーマ
            self.setStyleSheet("""
                QMainWindow, QDialog {
                    background-color: #2b2b2b;
                    color: #e0e0e0;
                }
                QWidget {
                    background-color: #2b2b2b;
                    color: #e0e0e0;
                }
                QTextEdit, QLineEdit, QComboBox, QCalendarWidget {
                    background-color: #3c3c3c;
                    color: #e0e0e0;
                    border: 1px solid #555555;
                }
                QPushButton {
                    background-color: #4b4b4b;
                    color: #e0e0e0;
                    border: 1px solid #555555;
                    padding: 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #555555;
                }
                QLabel {
                    color: #e0e0e0;
                }
                QListWidget {
                    background-color: #3c3c3c;
                    color: #e0e0e0;
                    border: 1px solid #555555;
                }
                QListWidget::item:selected {
                    background-color: #4b6eaf;
                }
                QMenuBar, QMenu {
                    background-color: #2b2b2b;
                    color: #e0e0e0;
                }
                QMenuBar::item:selected, QMenu::item:selected {
                    background-color: #4b6eaf;
                }
                QToolBar {
                    background-color: #323232;
                    border: 1px solid #555555;
                }
                QStatusBar {
                    background-color: #323232;
                    color: #e0e0e0;
                }
            """)
        else:
            # ライトテーマ
            self.setStyleSheet("")
    
    def show_month_stats(self):
        # 現在の月の統計を表示
        year = self.selected_date.year()
        month = self.selected_date.month()
        
        # その月の日記エントリ数をカウント
        entry_count = 0
        mood_counts = {}
        tag_counts = {}
        
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                date_str = file_name[:-5]  # .jsonを除去
                date = QDate.fromString(date_str, 'yyyy-MM-dd')
                
                if date.year() == year and date.month() == month:
                    entry_count += 1
                    
                    # 気分と使用されたタグをカウント
                    try:
                        with open(os.path.join(self.diary_folder, file_name), 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                            mood = data.get("mood", "不明")
                            if mood in mood_counts:
                                mood_counts[mood] += 1
                            else:
                                mood_counts[mood] = 1
                            
                            for tag in data.get("tags", []):
                                if tag in tag_counts:
                                    tag_counts[tag] += 1
                                else:
                                    tag_counts[tag] = 1
                    except:
                        continue
        
        # 統計情報を表示
        month_names = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
        month_name = month_names[month - 1]
        
        stats_dialog = QDialog(self)
        stats_dialog.setWindowTitle(f"{year}年{month_name}の統計")
        stats_dialog.setMinimumWidth(400)
        
        layout = QVBoxLayout(stats_dialog)
        
        # 基本情報
        layout.addWidget(QLabel(f"<h3>{year}年{month_name}の統計</h3>"))
        layout.addWidget(QLabel(f"日記エントリ数: {entry_count}"))
        
        # 月のカレンダー日数
        days_in_month = calendar.monthrange(year, month)[1]
        completion_rate = round(entry_count / days_in_month * 100, 1)
        layout.addWidget(QLabel(f"月の記入率: {completion_rate}%"))
        
        # 気分の分布
        if mood_counts:
            layout.addWidget(QLabel("<h4>気分の分布:</h4>"))
            for mood, count in sorted(mood_counts.items(), key=lambda x: x[1], reverse=True):
                percentage = round(count / entry_count * 100, 1)
                layout.addWidget(QLabel(f"{mood}: {count}回 ({percentage}%)"))
        
        # タグの分布
        if tag_counts:
            layout.addWidget(QLabel("<h4>よく使われたタグ:</h4>"))
            for tag, count in sorted(tag_counts.items(), key=lambda x: x[1], reverse=True)[:5]:
                layout.addWidget(QLabel(f"{tag}: {count}回"))
        
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(stats_dialog.accept)
        layout.addWidget(close_button)
        
        stats_dialog.exec_()
    
    def show_year_stats(self):
        # 現在の年の統計を表示
        year = self.selected_date.year()
        
        # 月ごとの日記エントリ数をカウント
        monthly_counts = [0] * 12
        mood_counts = {}
        tag_counts = {}
        
        for file_name in os.listdir(self.diary_folder):
            if file_name.endswith('.json') and file_name != "metadata.json":
                date_str = file_name[:-5]  # .jsonを除去
                date = QDate.fromString(date_str, 'yyyy-MM-dd')
                
                if date.year() == year:
                    monthly_counts[date.month() - 1] += 1
                    
                    # 気分と使用されたタグをカウント
                    try:
                        with open(os.path.join(self.diary_folder, file_name), 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                            mood = data.get("mood", "不明")
                            if mood in mood_counts:
                                mood_counts[mood] += 1
                            else:
                                mood_counts[mood] = 1
                            
                            for tag in data.get("tags", []):
                                if tag in tag_counts:
                                    tag_counts[tag] += 1
                                else:
                                    tag_counts[tag] = 1
                    except:
                        continue
        
        # 統計情報を表示
        stats_dialog = QDialog(self)
        stats_dialog.setWindowTitle(f"{year}年の統計")
        stats_dialog.setMinimumSize(500, 400)
        
        layout = QVBoxLayout(stats_dialog)
        
        # 基本情報
        layout.addWidget(QLabel(f"<h3>{year}年の統計</h3>"))
        total_entries = sum(monthly_counts)
        layout.addWidget(QLabel(f"日記エントリ総数: {total_entries}"))
        
        # 月ごとの統計
        layout.addWidget(QLabel("<h4>月ごとの日記数:</h4>"))
        
        month_grid = QGridLayout()
        month_names = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
        
        for i, (month, count) in enumerate(zip(month_names, monthly_counts)):
            month_grid.addWidget(QLabel(month), i // 3, (i % 3) * 2)
            month_grid.addWidget(QLabel(f"{count}日"), i // 3, (i % 3) * 2 + 1)
        
        layout.addLayout(month_grid)
        
        # 気分の分布
        if mood_counts:
            layout.addWidget(QLabel("<h4>年間の気分分布:</h4>"))
            mood_layout = QGridLayout()
            
            for i, (mood, count) in enumerate(sorted(mood_counts.items(), key=lambda x: x[1], reverse=True)):
                if total_entries > 0:
                    percentage = round(count / total_entries * 100, 1)
                    mood_layout.addWidget(QLabel(mood), i // 2, (i % 2) * 2)
                    mood_layout.addWidget(QLabel(f"{count}回 ({percentage}%)"), i // 2, (i % 2) * 2 + 1)
            
            layout.addLayout(mood_layout)
        
        # タグの分布
        if tag_counts:
            layout.addWidget(QLabel("<h4>よく使われたタグ:</h4>"))
            for tag, count in sorted(tag_counts.items(), key=lambda x: x[1], reverse=True)[:8]:
                percentage = round(count / total_entries * 100, 1)
                layout.addWidget(QLabel(f"{tag}: {count}回 ({percentage}%)"))
        
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(stats_dialog.accept)
        layout.addWidget(close_button)
        
        stats_dialog.exec_()
    
    def show_about(self):
        about_text = """
        <h3>PyQt5 日記アプリ</h3>
        <p>バージョン 1.0</p>
        <p>このアプリケーションは、日々の出来事や思いを記録するためのシンプルな日記アプリケーションです。</p>
        <p>主な機能:</p>
        <ul>
            <li>日記の作成、編集、保存</li>
            <li>タグ付け、検索機能</li>
            <li>お気に入り登録</li>
            <li>テキスト書式設定</li>
            <li>統計情報の表示</li>
            <li>エクスポート/インポート機能</li>
            <li>ライト/ダークテーマ切り替え</li>
        </ul>
        """
        
        QMessageBox.about(self, "このアプリについて", about_text)
    
    def closeEvent(self, event):
        # 終了前に保存確認
        if self.text_edit.document().isModified():
            reply = QMessageBox.question(self, '確認', 
                                        '変更を保存しますか？',
                                        QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
            
            if reply == QMessageBox.Save:
                self.save_entry()
            elif reply == QMessageBox.Cancel:
                event.ignore()
                return
        
        event.accept()

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
                        date_str = file_name[:-5]  # .jsonを除去
                        date = QDate.fromString(date_str, 'yyyy-MM-dd')
                        
                        if date.isValid():
                            title = data.get("title", "無題")
                            mood = data.get("mood", "")
                            tags = data.get("tags", [])
                            tags_str = ", ".join(tags) if tags else "タグなし"
                            modified = data.get("last_modified", "")
                            
                            # テキスト内容のプレビュー
                            content = data.get("content", "")
                            plain_content = self.html_to_plain(content)
                            preview = plain_content[:100] + "..." if len(plain_content) > 100 else plain_content
                            
                            # 日本語形式の日付表示
                            jp_date = date.toString('yyyy年MM月dd日(ddd)')
                            
                            diary_entries.append({
                                "date": date,
                                "date_str": date_str,
                                "jp_date": jp_date,
                                "title": title,
                                "mood": mood,
                                "tags": tags,
                                "tags_str": tags_str,
                                "modified": modified,
                                "preview": preview
                            })
                except:
                    continue
        
        # 日付順に並べ替え（新しい順）
        diary_entries.sort(key=lambda x: x["date"], reverse=True)
        original_entries = diary_entries.copy()  # 元のリストを保持
        
        # リストに追加する関数
        def update_list_items():
            diary_list.clear()
            for entry in diary_entries:
                # シンプルにタイトルと日付のみ表示
                display_text = f"{entry['title']} - {entry['jp_date']}"
                
                item = QListWidgetItem(display_text)
                item.setData(Qt.UserRole, entry['date_str'])
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
                    filter_text in entry['jp_date'].lower()):
                    filtered_entries.append(entry)
            
            diary_entries = filtered_entries
            sort_entries()  # 現在のソート順を適用
        
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
            date_str = item.data(Qt.UserRole)
            date = QDate.fromString(date_str, 'yyyy-MM-dd')
            self.calendar.setSelectedDate(date)
            diary_list_dialog.accept()
        
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
                
            date_str = diary_list.currentItem().data(Qt.UserRole)
            selected_entry = None
            
            for entry in diary_entries:
                if entry['date_str'] == date_str:
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

    def show_advanced_search(self):
        """
        詳細検索ダイアログを表示する
        """
        # 検索ダイアログの作成
        search_dialog = QDialog(self)
        search_dialog.setWindowTitle("詳細検索")
        search_dialog.setMinimumSize(700, 600)
        
        layout = QVBoxLayout(search_dialog)
        
        # タイトル
        title_label = QLabel("<h2>日記の詳細検索</h2>")
        layout.addWidget(title_label)
        
        # 検索条件入力フォーム
        form_layout = QGridLayout()
        form_layout.setVerticalSpacing(10)
        form_layout.setHorizontalSpacing(15)
        
        # キーワード検索
        form_layout.addWidget(QLabel("<b>キーワード:</b>"), 0, 0)
        keyword_edit = QLineEdit()
        keyword_edit.setPlaceholderText("検索キーワードを入力（複数キーワードはスペース区切り）")
        form_layout.addWidget(keyword_edit, 0, 1, 1, 3)
        
        # 検索オプション
        option_layout = QHBoxLayout()
        
        title_only_check = QCheckBox("タイトルのみ検索")
        option_layout.addWidget(title_only_check)
        
        case_sensitive_check = QCheckBox("大文字/小文字を区別")
        option_layout.addWidget(case_sensitive_check)
        
        exact_match_check = QCheckBox("完全一致")
        option_layout.addWidget(exact_match_check)
        
        option_layout.addStretch()
        form_layout.addLayout(option_layout, 1, 1, 1, 3)
        
        # 日付範囲
        form_layout.addWidget(QLabel("<b>日付範囲:</b>"), 2, 0)
        
        date_range_layout = QHBoxLayout()
        
        # 開始日
        from_layout = QVBoxLayout()
        from_layout.addWidget(QLabel("開始日:"))
        date_from = QCalendarWidget()
        date_from.setMaximumHeight(200)
        from_layout.addWidget(date_from)
        
        # 終了日
        to_layout = QVBoxLayout()
        to_layout.addWidget(QLabel("終了日:"))
        date_to = QCalendarWidget()
        date_to.setMaximumHeight(200)
        date_to.setSelectedDate(QDate.currentDate())
        to_layout.addWidget(date_to)
        
        # 一ヶ月前の日付をデフォルトの開始日に設定
        default_from_date = QDate.currentDate().addMonths(-1)
        date_from.setSelectedDate(default_from_date)
        
        date_range_layout.addLayout(from_layout)
        date_range_layout.addLayout(to_layout)
        
        form_layout.addLayout(date_range_layout, 3, 0, 1, 4)
        
        # 日付ショートカットボタン
        date_shortcuts = QHBoxLayout()
        
        today_btn = QPushButton("今日")
        today_btn.clicked.connect(lambda: set_date_range(0))
        date_shortcuts.addWidget(today_btn)
        
        week_btn = QPushButton("過去1週間")
        week_btn.clicked.connect(lambda: set_date_range(7))
        date_shortcuts.addWidget(week_btn)
        
        month_btn = QPushButton("過去1ヶ月")
        month_btn.clicked.connect(lambda: set_date_range(30))
        date_shortcuts.addWidget(month_btn)
        
        three_month_btn = QPushButton("過去3ヶ月")
        three_month_btn.clicked.connect(lambda: set_date_range(90))
        date_shortcuts.addWidget(three_month_btn)
        
        year_btn = QPushButton("過去1年")
        year_btn.clicked.connect(lambda: set_date_range(365))
        date_shortcuts.addWidget(year_btn)
        
        all_btn = QPushButton("すべて")
        all_btn.clicked.connect(lambda: set_date_range(-1))
        date_shortcuts.addWidget(all_btn)
        
        form_layout.addLayout(date_shortcuts, 4, 0, 1, 4)
        
        # 日付範囲を設定する関数
        def set_date_range(days):
            end_date = QDate.currentDate()
            date_to.setSelectedDate(end_date)
            
            if days == -1:
                # すべての期間（古い日付）
                start_date = QDate(2000, 1, 1)
            else:
                # 指定日数前
                start_date = end_date.addDays(-days)
            
            date_from.setSelectedDate(start_date)
        
        # タグ検索
        form_layout.addWidget(QLabel("<b>タグ:</b>"), 5, 0)
        tag_combo = QComboBox()
        tag_combo.addItem("すべてのタグ")
        tag_combo.addItems(sorted(self.metadata["tags"]))
        form_layout.addWidget(tag_combo, 5, 1)
        
        # 気分検索
        form_layout.addWidget(QLabel("<b>気分:</b>"), 5, 2)
        mood_combo = QComboBox()
        mood_combo.addItem("すべての気分")
        mood_combo.addItems(self.metadata["moods"])
        form_layout.addWidget(mood_combo, 5, 3)
        
        layout.addLayout(form_layout)
        
        # 検索ボタン
        search_layout = QHBoxLayout()
        search_button = QPushButton("検索")
        search_button.setMinimumHeight(35)
        search_button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        search_layout.addWidget(search_button)
        
        reset_button = QPushButton("条件リセット")
        reset_button.clicked.connect(lambda: reset_search_form())
        search_layout.addWidget(reset_button)
        
        layout.addLayout(search_layout)
        
        # 条件リセット関数
        def reset_search_form():
            keyword_edit.clear()
            title_only_check.setChecked(False)
            case_sensitive_check.setChecked(False)
            exact_match_check.setChecked(False)
            set_date_range(30)  # デフォルトの日付範囲に戻す
            tag_combo.setCurrentIndex(0)
            mood_combo.setCurrentIndex(0)
            result_list.clear()
            result_label.setText("検索結果:")
        
        # 結果表示リスト
        result_label = QLabel("<b>検索結果:</b>")
        layout.addWidget(result_label)
        
        result_list = QListWidget()
        result_list.setAlternatingRowColors(True)
        result_list.setStyleSheet("""
            QListWidget::item { 
                border-bottom: 1px solid #ddd; 
                padding: 5px;
            }
            QListWidget::item:selected { 
                background-color: #e6f2ff;
                color: black;
            }
        """)
        layout.addWidget(result_list)
        
        # 検索実行関数
        def perform_search():
            keyword = keyword_edit.text()
            from_date = date_from.selectedDate()
            to_date = date_to.selectedDate()
            selected_tag = tag_combo.currentText()
            selected_mood = mood_combo.currentText()
            title_only = title_only_check.isChecked()
            case_sensitive = case_sensitive_check.isChecked()
            exact_match = exact_match_check.isChecked()
            
            # 結果リストをクリア
            result_list.clear()
            
            # キーワードの処理
            keywords = []
            if keyword:
                if not case_sensitive:
                    keyword = keyword.lower()
                    
                if exact_match:
                    keywords = [keyword]
                else:
                    keywords = keyword.split()
            
            # 全ての日記ファイルを検索
            matching_entries = []
            
            # 全ての日記ファイルを検索
            matching_entries = []
            
            for file_name in os.listdir(self.diary_folder):
                if file_name.endswith('.json') and file_name != "metadata.json":
                    file_path = os.path.join(self.diary_folder, file_name)
                    
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            
                            # 日付チェック
                            date_str = file_name[:-5]  # .jsonを除去
                            date = QDate.fromString(date_str, 'yyyy-MM-dd')
                            
                            if not (from_date <= date <= to_date):
                                continue
                            
                            # タグチェック
                            if selected_tag != "すべてのタグ" and selected_tag not in data.get("tags", []):
                                continue
                            
                            # 気分チェック
                            if selected_mood != "すべての気分" and selected_mood != data.get("mood", ""):
                                continue
                            
                            # キーワードチェック
                            if keyword:
                                title = data.get("title", "").lower()
                                content = data.get("content", "").lower()
                                plain_content = self.html_to_plain(content).lower()
                                tags = " ".join(data.get("tags", [])).lower()
                                
                                if (keyword not in title and 
                                    keyword not in plain_content and 
                                    keyword not in tags):
                                    continue
                            
                            # すべての条件をパス
                            title = data.get("title", "無題")
                            jp_date = date.toString('yyyy年MM月dd日')
                            
                            matching_entries.append({
                                "date": date,
                                "date_str": date_str,
                                "jp_date": jp_date,
                                "title": title
                            })
                    except:
                        continue
            
            # 日付順に並べ替え（新しい順）
            matching_entries.sort(key=lambda x: x["date"], reverse=True)
            
            # 結果をリストに追加
            if matching_entries:
                for entry in matching_entries:
                    item = QListWidgetItem(f"{entry['jp_date']} - {entry['title']}")
                    item.setData(Qt.UserRole, entry['date_str'])
                    result_list.addItem(item)
                
                result_label.setText(f"検索結果: {len(matching_entries)}件見つかりました")
            else:
                result_label.setText("検索結果: 0件")
        
        # 検索ボタンに関数を接続
        search_button.clicked.connect(perform_search)
        
        # アイテムクリック時の処理
        def open_diary(item):
            date_str = item.data(Qt.UserRole)
            date = QDate.fromString(date_str, 'yyyy-MM-dd')
            self.calendar.setSelectedDate(date)
            search_dialog.accept()
        
        result_list.itemDoubleClicked.connect(open_diary)
        
        # ボタンレイアウト
        button_layout = QHBoxLayout()
        
        open_button = QPushButton("開く")
        open_button.clicked.connect(lambda: open_diary(result_list.currentItem()) if result_list.currentItem() else None)
        button_layout.addWidget(open_button)
        
        close_button = QPushButton("閉じる")
        close_button.clicked.connect(search_dialog.reject)
        button_layout.addWidget(close_button)
        
        layout.addLayout(button_layout)
        
        # ダイアログを表示
        search_dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DiaryApp()
    window.show()
    sys.exit(app.exec_())