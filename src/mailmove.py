# 必要なモジュールをインポート
import pythoncom
import signal
import sys
import time
import win32com.client
import json
import re
import os
from datetime import datetime
from collections import defaultdict
from watchfiles import watch  # ファイル監視用のライブラリ
import threading # スレッド処理用のライブラリ
from flask import Flask

# ローカルフォルダへ保存するためのベースディレクトリ(** 適宜書き換えてください **)
BASE_DIR = R"C:\Users\pcrik\Documents\メールバックアップ"
data_ready = threading.Event()

app = Flask(__name__)

@app.route('/')
def index():
    """標準出力を表示するページ"""
    return "メール監視中..."

# メールを保存するためのクラス
class MailSaver:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        self.counters = defaultdict(int)  # フォルダごとのカウンターを管理

    def save_message(self, message, folder_path):
        """メールをローカルフォルダに保存"""
        self.counters[folder_path] += 1
        folder_name = message.ReceivedTime.strftime("%Y_%m%d") + f"_{self.counters[folder_path]:05d}"
        full_path = os.path.join(folder_path, folder_name)
        os.makedirs(full_path, exist_ok=True)

        # 添付ファイルの保存
        if message.Attachments.Count > 0:
            self.save_attachments(os.path.join(full_path, "attachments"), message)

        # メール本文の保存
        self.save_email_content(os.path.join(full_path, "本文.txt"), message)

    @staticmethod
    def save_attachments(folder_path, message):
        """添付ファイルを保存"""
        os.makedirs(folder_path, exist_ok=True)
        for attachment in message.Attachments:
            attachment_file_path = os.path.join(folder_path, attachment.FileName)
            attachment.SaveAsFile(attachment_file_path)
            print(f"添付ファイルを保存しました: {attachment_file_path}")

    @staticmethod
    def save_email_content(file_path, message):
        """メール内容をファイルに書き込む"""
        header_fields = [
            ("送信日時", message.SentOn),
            ("送信者", f"{message.SenderName} <{message.SenderEmailAddress}>"),
            # ("宛先 (To)", message.To or "宛先なし"),
            ("CC", message.CC or "CCなし"),
            ("BCC", message.BCC or "BCCなし"),
            ("件名", message.Subject or "件名なし")
        ]
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("=== メール基本情報 ===\n")
            for label, value in header_fields:
                if value:
                    f.write(f"・{label}: {value}\n")
            f.write("\n=== 本文 ===\n")
            f.write(f"{message.Body or '本文なし'}\n")
        print(f"メールを保存しました: {file_path}")

# メールを仕分けして保存するためのクラス
class MailProcessor:
    def __init__(self, base_dir=BASE_DIR, json_file=R"json/mail.json", root=None):
        self.root = root
        self.base_dir = base_dir
        self.json_file = json_file
        self.folders = {}
        self.rules = {}
        self.local_paths = {}
        self.saver = MailSaver(base_dir)

    def load_rules(self):
        """JSONファイルから仕分けルールをロード"""
        try:
            with open(self.json_file, "r", encoding="utf-8") as f:
                self.rules = json.load(f)
                print("仕分けルールを読み込みました。")
                return 0
        except Exception as e:
            print(f"JSONファイルの読み込み中にエラーが発生しました: {e}")
            self.rules = {}
            return 1

    def find_folder(self, parent_folder, folder_name):
        """指定されたフォルダ名を持つフォルダを検索（再帰処理対応）"""
        try:
            folders = folder_name.split("/")
            current_folder_name = folders[0]
            remaining_folders = "/".join(folders[1:]) if len(folders) > 1 else None

            for folder in parent_folder.Folders:
                if folder.Name == current_folder_name:
                    print(f"フォルダ '{current_folder_name}' は既に存在します。")
                    if remaining_folders:
                        return self.find_folder(folder, remaining_folders)
                    return folder

            new_folder = parent_folder.Folders.Add(current_folder_name)
            print(f"フォルダ '{current_folder_name}' を作成しました。")
            if remaining_folders:
                return self.find_folder(new_folder, remaining_folders)
            return new_folder

        except Exception as e:
            print(f"フォルダ作成中にエラーが発生しました: {e}")
            return None

    def setup_local_paths(self):
        """ローカルフォルダを作成"""
        for category, rule in self.rules.items():
            if category == "del":
                continue
            folder_structure = rule["folder"]
            print(f"folder_structure is: {folder_structure}")
            path = os.path.join(self.base_dir, *folder_structure.split("/"))
            self.folders[category] = self.find_folder(self.root, folder_structure)
            if category == "uncategorized" or category == "archive":
                continue
            if not os.path.isdir(path):
                os.makedirs(path, exist_ok=True)
                print(f"{folder_structure} フォルダを作成しました: {path}")
            self.local_paths[category] = path
        # for category, folder in self.folders.items():
        #     print(f"カテゴリ: {category}, フォルダ名: {folder.Name}, フォルダパス: {folder.FolderPath}")

    def categorize_email(self, message, dic):
        """メールを仕分けルールに基づいて分類"""
        for category, rule in dic.items():
            if self.match_rule(message, rule):
                return category
        return None

    @staticmethod
    def match_rule(message, rule):
        """メールがルールに一致するか判定"""
        for subject in rule["subject"]:
            if re.search(subject, message.Subject or ""):
                return True
        for address in rule["address"]:
            if address in (message.SenderEmailAddress or ""):
                return True
        return False

    def move_mail(self, dic, folders, target_folder, view_none=True, view_move=True, view_delete=True):
        """メールを移動、削除、または未分類"""
        counter_move = 0
        counter_remain = 0
        counter_delete = 0
        list_move = []

        sorted_messages = sorted(target_folder.Items, key=lambda msg: msg.ReceivedTime)
        for message in sorted_messages:
            key = self.categorize_email(message, dic)
            if key == "del":
                counter_delete += 1
                list_move.append((message, None))
            elif key is None or folders[key] is None or folders[key] == target_folder.FolderPath:
                counter_remain += 1
                list_move.append((message, folders["uncategorized"]))
                if view_none:
                    print(f"{counter_remain}, none, {folders["uncategorized"].Name} , {message.Subject}")
            elif dic[key]["unread"] or not message.UnRead:
                counter_move += 1
                if view_move:
                    print(folders[key].Name, message.Subject)
                list_move.append((message, folders[key]))
            # else:
            #     counter_remain += 1
            #     if view_none:
            #         print("unread", message.Subject)

        for message, dest in list_move:
            if dest is None:
                if view_delete:
                    print("delete", message.Subject)
                message.Delete()
            else:
                # print(dest, message.Subject)
                message.UnRead = False
                message.Move(dest)
                if dest != folders["uncategorized"] and dest != folders["archive"]:
                    folder_path = self.local_paths.get(key, None)
                    if folder_path:
                        self.saver.save_message(message, folder_path)

        print("moved:", counter_move, "delete:", counter_delete, "archive:", counter_remain)

def signal_handler(signum, frame):
    """シグナルハンドラ"""
    sys.exit(0)

def initialize_signal_handler():
    """シグナルハンドラを初期化"""
    signal.signal(signal.SIGINT, signal_handler)

def watch_json_folder():
    print("jsonファイルの監視開始")
    for changes in watch(R"C:\Users\pcrik\my_code\sort_emails_v2\src\json"): # 適宜書き換えてください
        print("jsonファイルの変更を検出")
        data_ready.set()  # JSONファイルの変更を検出

class OutlookEventHandler:
    """Outlookのイベントを処理するクラス"""
    def __init__(self):
        self.processor = None
        self.inbox = None

    def set_context(self, processor, inbox):
        self.processor = processor
        self.inbox = inbox

    def OnNewMailEx(self, entry_id_collection):
        """新しいメールを受信したときに呼び出されるイベント"""
        print("新しいメールを受信しました。")
        try:
            # メールを仕分け
            self.processor.move_mail(self.processor.rules, self.processor.folders, self.inbox)
        except Exception as e:
            print(f"メール処理中にエラーが発生しました: {e}")

# スクリプトのエントリーポイント
if __name__ == "__main__":
    initialize_signal_handler()  # シグナルハンドラを初期化

    try:
        mail_app = win32com.client.DispatchWithEvents("Outlook.Application", OutlookEventHandler)
        root = mail_app.Session.DefaultStore.GetRootFolder()
        inbox = mail_app.GetNamespace("MAPI").GetDefaultFolder(6)

        processor = MailProcessor(root=root)
        if processor.load_rules():
            print("仕分けルールの読み込みに失敗しました。")
            sys.exit(1)

        processor.setup_local_paths()
        processor.move_mail(processor.rules, processor.folders, inbox)

        json_watch_thread = threading.Thread(target=watch_json_folder, daemon=True) # JSONファイル監視スレッド    
        json_watch_thread.start()  # JSONファイル監視スレッド開始
        
        handler = mail_app
        handler.set_context(processor, inbox)  # イベントハンドラにコンテキストを設定

         # Flask アプリケーションを別スレッドで実行
        flask_thread = threading.Thread(target=lambda: app.run(debug=True, use_reloader=False), daemon=True)
        flask_thread.start()
        
        print("メール監視を開始します...")
        while True:
            pythoncom.PumpWaitingMessages()  # COMメッセージを処理
            
            if data_ready.is_set():
                data_ready.clear()  # イベントをリセット
                if processor.load_rules():
                    print("仕分けルールの読み込みに失敗しました。")
                    sys.exit(1)
                processor.setup_local_paths()
                # print(processor.rules)

            time.sleep(0.1)  # 0.1秒スリープしてCPU負荷を抑える
    except KeyboardInterrupt:
        print("CTRL+C による割り込みを検出しました。プログラムを終了します。")
    finally:
        print("リソースを解放しています...")
        # 必要なリソース解放処理をここに記述
        try:
            del mail_app
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"リソース解放中にエラーが発生しました: {e}")
        finally:
            print("リソース解放完了")
        # プログラム終了メッセージ
        print("プログラムを終了しました。")
