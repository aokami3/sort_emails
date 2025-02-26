import win32com.client
import json
import re
import os

# Outlook関係のオブジェクト初期化
app = win32com.client.Dispatch("Outlook.Application")
root = app.Session.DefaultStore.GetRootFolder() #アプリに設定されているパス
ns = app.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)
messages = inbox.Items

# ローカルフォルダへ保存するためのパスを定義
dir_path = R"C:\Users\pcrik\OneDrive\ドキュメント\メールバックアップ" # Rはパスのエスケープシーケンスに必要
local_path = {}

# Outlookのフォルダ検索
def findfolder(root, name):
    for folder in root.Folders:
        # フォルダ名の部分一致
        if name in folder.name:
            # print(folder.folderpath)
            return folder
        # フォルダにない場合、サブフォルダも検索(再起関数)
        ret = findfolder(folder, name)
        if ret is not None:
            return ret
    return None

# 条件適合判定
# mail.jsonのアドレスには差出人名または差出人アドレス、宛先メールアドレスが入っている(差出人メールアドレスは必要か？後で検討)
def isit(message, subjects, addresses=[]): # addressesには何が入ってるん？mail.jsonで定義したフォルダのメールアドレス（リスト）が入る
    # 件名（正規表現）での判定
    for subject in subjects:
        # reライブラリのsearchメソッドを使用
        if re.search(subject, message.subject)!= None:
            return True
    # メールアドレス（部分一致）判定
    for address in addresses:
        # 差出人名
        if address in message.sendername: # 送り主の中に事前に定義した名前があれば真
            return True
        # 差出人アドレス
        if address in message.senderemailaddress:
            return True
        # 宛先
        # for recip in message.recipients:
        #     if address in recip.name or address in recip.address:
        #         return True
    return False

# アーカイブ先フォルダ検索
def whichFolder(message, dic):
    # 会議案内などは除外(IPM.Noteのみに処理を行う。IPM.Noteは電子メールメッセージ意味する)
    if message.messageClass == "IPM.Note":
        for key in dic:
            if isit(message, dic[key]["subject"], dic[key]["address"]): # isit()関数の中でもループしているため、フォルダとフォルダのアドレスと件名が多いほど実行時間がかかる
                return key
    return None

# JSONファイルから移動条件をロードする(最初の一回のみ実行すれば良い, 追加した場合のアップデートが必要になる可能性あり)
def load_json(filename="mail.json"):
    with open("mail.json", "r", encoding = "utf-8") as f:
        dic = json.load(f)
        folders = {}
        for k in dic:
            # 識別名の先頭が$ならコメントとして扱う
            if not k.startswith('$'):
                folders[k] = findfolder(root, dic[k]["folder"]) # Noneの場合はフォルダが見つからなかったことを表す
    return dic, folders

# メールのアーカイブ処理メイン部
def move_mail(
    dic, folders, target_folder=inbox, view_none=True, view_move=True, view_delete=True
):
    i = 1
    counter_move = 0
    counter_remain = 0
    counter_delete = 0
    list_move = list()
    for message in target_folder.Items:#target_folderは受信トレイを指定
        print(message)
        # print(len(target_folder.Items)) # これで受信フォルダにあるメールの件数が分かる
        key = whichFolder(message, dic) # dicはload_json()で定義済み
        # print(key) 
        if key == "del": # keyはフォルダ名。messageに適したフォルダがあったらフォルダ名が返ってきて、無かったらNoneが帰ってくる
            counter_delete += 1
            list_move.append((message, None))
        elif ( # 事前に設定していないメールの種類はそのまま残しておき、出力する
            key is None
            or folders[key] is None
            or folders[key].folderpath == target_folder.folderpath
        ):
            counter_remain += 1
            if view_none: # 通常時はTrue, do_all_folder関数の時のみFalse
                print(counter_remain, "none", message.subject) # 仕分け未登録の件名を出力
        elif dic[key]["unread"] or not message.unread: # 移動の際、未読を対象にするか？未読を対象にしない場合、message.unread == Flase
            counter_move += 1
            if view_move: #移動するものを出力するか? ログ用に使用
                print(folders[key].name, message.subject)
            list_move.append((message, folders[key]))
        else: # 未読を対象にしない場合かつ、メールが未読の場合は残します
            counter_remain += 1
            if view_none: # Noneとなったメールの件名を出力します
                print("unread", message.subject)
        i += 1
    for item in list_move:
        message = item[0]
        dest = item[1]
        if dest is None:
            if view_delete:
                print("delete", message.subject)
            message.delete()
        else:
            print(dest.name, message.subject)
            message.unread = False
            message.move(dest)
            # メールの処理
            if message.Attachments.Count > 0:
                # 添付ファイルがある場合、フォルダを作成して保存
                file_name = sanitize_filename(message.Subject)
                attachment_folder_path = os.path.join(local_path[str(dest)], file_name)
                save_attachments(attachment_folder_path, message)
                
                # メール内容をフォルダ内のテキストファイルに保存
                file_name = file_name + ".txt"
                file_path = os.path.join(attachment_folder_path, file_name)
            else:
                # 添付ファイルがない場合、通常のテキストファイルに保存
                file_name = sanitize_filename(message.Subject) + ".txt"
                file_path = os.path.join(local_path[str(dest)], file_name)

            # メール内容を保存
            save_email_content(file_path, message)
    print("moved:", counter_move, "delete:", counter_delete, "remain:", counter_remain)
    
# アーカイブ処理を全アーカイブ対象フォルダに対して実行
def do_all_folder(dic, folders):
    for k in dic:
        print(k)
        if k != "del":
            move_mail(dic, folders, target_folder=folders[k], view_none=False)
    print("do all done.")

# jsonからフォルダと
dic, folders = load_json()

def check_dir():
    """ローカルに保存するためのフォルダがあるか確認し、無い場合は新たに作成します。"""
    for category, folder_name in folders.items():
        # "del" カテゴリはスキップ
        if category == "del":
            continue

        # 親フォルダかどうかの判定
        if "#" in str(folder_name) or category == "archive":
            path = os.path.join(dir_path, str(folder_name))
        else:
        #子フォルダの場合
            path = os.path.join(path, str(folder_name))

        # フォルダが存在しない場合は作成
        if not os.path.isdir(path):
            os.makedirs(path)
            print(f"{folder_name} フォルダを作成しました: {path}")
        
        # 作成したフォルダのパスを local_path に保存
        local_path[str(folder_name)] = path

def sanitize_filename(filename):
    # 使用できない文字のパターン
    invalid_chars = r'[\\/:*?"<>| ]'
    # 不正な文字を "_" に置き換える
    sanitized = re.sub(invalid_chars, '_', filename)
    # 先頭と末尾のスペースやドットを削除
    sanitized = sanitized.strip(' .')
    return sanitized

def save_email_content(file_path, message):
    """メール内容をファイルに書き込む"""
    with open(file_path, "w", encoding='utf-8') as f:
        f.write(f"送信日時: {message.SentOn}\n")
        f.write(f"送信者: {message.Sender}\n")
        f.write(f"送信者のメールアドレス: {message.SenderEmailAddress}\n")
        f.write(f"宛先 (To): {message.To}\n")
        f.write(f"CC: {message.CC}\n")
        f.write(f"BCC: {message.BCC}\n")
        f.write(f"件名: {message.Subject}\n")
        f.write("本文 (プレーンテキスト):\n")
        f.write(message.Body)
        f.write("\n\n本文 (HTML):\n")
        f.write(message.HTMLBody)
    print(f"メール内容を保存しました: {file_path}")

def save_attachments(attachment_folder_path, message):
    """添付ファイルを保存"""
    os.makedirs(attachment_folder_path, exist_ok=True)
    for attachment in message.Attachments:
        attachment_file_path = os.path.join(attachment_folder_path, attachment.FileName)
        attachment.SaveAsFile(attachment_file_path)
        print(f"添付ファイルを保存しました: {attachment_file_path}")

# ここからメイン関数となる部分
check_dir()
# print(local_path)
# 受信フォルダに対して処理を行う場合
move_mail(dic, folders)