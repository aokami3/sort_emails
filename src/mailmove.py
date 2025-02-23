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
            file_name = sanitize_filename(message.subject) + ".txt"
            # print(os.path.join(local_path[str(dest)], file_name))            
            with open(os.path.join(local_path[str(dest)], file_name), "w", encoding='utf-8') as f:
                f.write(str(message.SentOn))
                f.write(str(message.Sender))
                f.write(str(message.subject))
                f.write(str(message.body))
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
    print("***** ローカルに保存するためのフォルダがあるか確認します。無い場合は新たに作成します。 *****")
    for cat, folder in folders.items():
        if '#' not in str(folder): # サブフォルダの場合, 親フォルダの後ろにくっつける
            path = os.path.join(path, str(folder))
        # print(os.path.join(dir_path, str(folder)))
        else:
            path = os.path.join(dir_path, str(folder))
        if not os.path.isdir(path):
            os.mkdir(path)
            print(str(folder) + "フォルダを作成しました：" + path)
            local_path[str(folder)] = path
    print("**********")

def sanitize_filename(filename):
    # 使用できない文字のパターン
    invalid_chars = r'[\\/:*?"<>| ]'
    # 不正な文字を "_" に置き換える
    sanitized = re.sub(invalid_chars, '_', filename)
    # 先頭と末尾のスペースやドットを削除
    sanitized = sanitized.strip(' .')
    return sanitized


# ここからメイン関数となる部分
check_dir()
# print(local_path)
# 受信フォルダに対して処理を行う場合
move_mail(dic, folders)