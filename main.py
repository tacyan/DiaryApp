import sys
import os
import json
import datetime
import calendar
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QTextEdit, QPushButton, QLabel, QCalendarWidget, QComboBox, 
                            QLineEdit, QMessageBox, QTabWidget, QGridLayout, QListWidget,
                            QListWidgetItem, QFileDialog, QColorDialog, QFontDialog, QMenu,
                            QAction, QToolBar, QStatusBar, QSplitter, QDialog)
from PyQt5.QtGui import QFont, QIcon, QTextCharFormat, QColor, QTextCursor, QTextListFormat
from PyQt5.QtCore import Qt, QDate, QTimer, QSize, QUrl
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent

class DiaryApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 日記アプリ")
        self.setGeometry(100, 100, 1000, 700)
        
        # データの保存先
        self.diary_folder = "diary_entries"
        if not os.path.exists(self.diary_folder):
            os.makedirs(self.diary_folder)
            
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
        
        # 箇条書きアクション
        bullet_action = QAction(QIcon.fromTheme("format-justify-fill"), "箇条書き", self)
        bullet_action.triggered.connect(self.insert_bullet_list)
        self.format_toolbar.addAction(bullet_action)
        
        self.right_layout.addWidget(self.format_toolbar)
        
        # テキストエディタ
        self.text_edit = QTextEdit()
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
        
        # 表示メニュー
        view_menu = menu_bar.addMenu("表示")
        
        theme_menu = view_menu.addMenu("テーマ")
        
        light_theme_action = QAction("ライト", self)
        light_theme_action.triggered.connect(lambda: self.change_theme("light"))
        theme_menu.addAction(light_theme_action)
        
        dark_theme_action = QAction("ダーク", self)
        dark_theme_action.triggered.connect(lambda: self.change_theme("dark"))
        theme_menu.addAction(dark_theme_action)
        
        # 検索メニュー
        search_menu = menu_bar.addMenu("検索")
        
        search_action = QAction("日記を検索", self)
        search_action.setShortcut("Ctrl+F")
        search_action.triggered.connect(self.search_entries)
        search_menu.addAction(search_action)
        
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
                self.text_edit.setHtml(data.get("content", ""))
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
                                                "HTMLファイル (*.html);;テキストファイル (*.txt);;JSONファイル (*.json)", 
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
            
            self.status_bar.showMessage(f"日記を {file_name} にエクスポートしました", 3000)
    
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
        # 簡易的なHTMLからプレーンテキスト変換
        import re
        text = re.sub(r'<[^>]+>', '', html)
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


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DiaryApp()
    window.show()
    sys.exit(app.exec_())