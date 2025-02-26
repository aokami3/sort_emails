# メール自動仕分け 手順書

## 使用方法

### 環境構築

1. 仮想環境の構築：```python3 -m venv myenv```
2. 仮想環境の起動```.\myenv\Scripts\activate```
3. (権限エラーになった場合、以下を実行し実行権限を変更)```Set-ExecutionPolicy RemoteSigned```
4. 必要なライブラリをインストール
	```pip install -r requirements.txt```
### 準備と実行

1. outlookアカウントの下にメールを振り分けたいフォルダを作成する。
<img width="1470" alt="スクリーンショット 2025-02-26 18 06 14" src="https://github.com/user-attachments/assets/3f5a57c0-fd70-4531-bb38-fa8637540947" />

2. mail_json_説明資料.txtを参考にして、mail.jsonを記述(テストの場合、少量のメールを対象にしてください)

3. mail_move.pyの14行目dir_pathを変更(場所に指定なし)

4. python ./mailmove.pyを実行

## 基本機能

- outlookで受信したメールの本文と添付ファイルをpythonでフォルダ(プロジェクト単位)に振り分け社内サーバへ保存
- 受信したメールと事前定義した仕分けのルールは部分一致で比較し、仕分け先フォルダを決めている（完全一致ではない）

### 追加機能

- プロジェクト単位でチャネルを分けて本文をteamsに時系列的に流す。
- ファイルはBoxなどのクラウドに保存し、リンクを流す（teamsとbox連携）

## Plan

### 実装順序

1. outlook操作
  件名から保存するフォルダを判別する。既存のフォルダがない場合は新たに作成する。
2. flaskサーバ上での実装
  セキュリティ面にも配慮する必要がある。
3. teams workflowを用いた連携
4. BOXクラウドからの通知
5. 

### 基本機能

以下の2つのライブラリを用いてメールをフォルダに振り分ける。
- pywin32：WindowsのCOMオブジェクトを操作するためのライブラリです。Outlookの操作に必要な機能を提供しています。メールの読み込みに使用。
- win32com.client：pywin32のサブモジュールで、COMオブジェクトを操作するための機能を提供しています。
(参考 : [PythonでOutlookを操作する：メール送信から予定抽出まで | Pythonの学習帳](https://beginner-engineers.com/python-outlook/#Python%E3%81%A8Outlook%E3%81%AE%E9%80%A3%E6%90%BA%E3%81%95%E3%81%9B%E3%82%8B%E3%81%A8))

常時稼働させるため、flaskサーバを構築し、その上でプログラムを動作させる。

### 追加機能

1.workflowのwebhook要求を受信するとチャンネルに投函する機能を用いてteamsへメッセージを送信。
<img width="1466" alt="スクリーンショット 2025-01-23 20 34 34" src="https://github.com/user-attachments/assets/839e6b15-2e5e-4855-9d56-771792d436de" />
<img width="1470" alt="スクリーンショット 2025-01-23 20 36 47" src="https://github.com/user-attachments/assets/f3ebf4b7-7ba7-4c87-a12b-5690b3facae4" />

黒塗りされている部分にはURLが表示されており、そこへPOSTリクエストを送信することで、以下のようにメッセージカード(内容は例)が送信される
<img width="645" alt="スクリーンショット 2025-01-23 20 40 16" src="https://github.com/user-attachments/assets/44b2e0ef-2431-4658-abbc-996561a0448f" />
（参考）[Microsoft TeamsのIncoming Webhookが廃止になるので、Workflows(Power Automate)で通知する方法を調べた - Devplatform blog](https://blog.devplatform.techmatrix.jp/blog/teams_workflows_notification/)

2.ファイルはBoxなどのクラウドに保存し、リンクを流す（teamsとbox連携）
pythonからboxへファイルをアップロードし、workflowの中にある機能を用いて通知とリンクを送信する。
<img width="1085" alt="スクリーンショット 2025-01-23 20 45 32" src="https://github.com/user-attachments/assets/a2686630-e08e-4737-8a89-07a8ddbda49b" />

（参考）[Box連携講座「組み合わせるとこんなに便利！Office365編」（2021.9.17） | 株式会社 Box Japan](https://www.boxsquare.jp/videos/admin/box-online-seminar-20210917)
また、pythonからBOXへのアップロードはBox Python SDKを利用して行う。

（参考）[Pythonを使用してBoxにコンテンツをアップロードする3つの方法 | by Yuko Taniguchi | Box Developer Japan Blog | Medium](https://medium.com/box-developer-japan-blog/python%E3%82%92%E4%BD%BF%E7%94%A8%E3%81%97%E3%81%A6box%E3%81%AB%E3%82%B3%E3%83%B3%E3%83%86%E3%83%B3%E3%83%84%E3%82%92%E3%82%A2%E3%83%83%E3%83%97%E3%83%AD%E3%83%BC%E3%83%89%E3%81%99%E3%82%8B3%E3%81%A4%E3%81%AE%E6%96%B9%E6%B3%95-db0e26ec3d74)

## 質問事項
- workflowの機能を利用するためにはteamsプランのアップグレードが必要であるが、問題ないでしょうか？
## 今後の計画
1. 基本機能を自身のPCを用いて試作
2. 追加機能を実装
