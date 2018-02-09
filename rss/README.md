## 概要
Googleスプレッドシート上でRSSを使いWEBサイトの更新情報を取得・管理します。また、取得した情報をNatural Language Classifier(NLC)を利用して任意のクラスタに分類するツールです。

<img src="https://github.com/softbank-developer/gsuite_with_watson/blob/master/rss/readme_images/logo.png" width="60px">

&emsp;↓

![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/rss/readme_images/data.png)

- `ソース〜サマリ`: RSS取得情報
- `分類器[1-3]:手入力`: 手動による分類(NLCの学習対象列)
- `分類器[1-3]:Watson`: NLCによる分類
- `分類器[1-3]:確信度`: NLCの確信度(0.0〜1.0)
- `分類器[1-3]:処理日時`: NLCによる分類を実施した日時  
1つのNLCで3つまでの分類器を利用します。


## 利用条件
- [Google](https://accounts.google.com/)アカウントを持っていること
  - Google スプレッドシート用
- [IBM Bluemix](https://accounts.google.com/)アカウントを持っていること
  - Natural Language Classifier用(一つ以上のインスタンスを用意しておくこと)


## 準備
Google ドライブにてスプレッドシートとGaoogle スプレッドシートとGoogle Apps Scriptを用意しています。ドライブからコピーして利用する場合は、本準備手順の7番のみの対応で準備は終わります。

1. スプレッドシートの作成  
任意の名称でスプレッドシートを作成します。

2. "設定"シートの作成  
作成したスプレッドシート内に、各種設定を管理するシートを作成します。  
[Googleドライブから環境をコピー](https://github.com/softbank-developer/gsuite_with_watson#環境構築)より、サンプルのシートを利用できます。

	(シート設定)
	- `データシート名`: データの格納シート
	- `開始列`: データを格納する開始列 
	- `開始行`: データを格納する開始行
	- `学習・分類対象`: NLCによる学習・分類の対象
	- `分類器[1-3]:手入力`: 手動で分類させる列
	- `分類器[1-3]:watson`: 分類結果を記載する列
	- `分類日時[1-3]`: 分類された日時を記載する列
	- `ログシート名`:  NLCの学習と分類ログの保存シート
	- `メール通知`:  メール通知機能の切り替え
	- `メール通知設定シート名`: メール通知条件の指定  

	(RSS設定)
	- `ソース`: 情報を取得源
	- `フィードURL`：フィード用のURL


	(分類器)
	- `Classifier ID`: NLCの分類器のID(学習後自動で挿入されます)
	- `ステータス`: NLCのステータス

	![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/rss/readme_images/config.png)

3. "データ"シートの作成   
作成したスプレッドシート内に、データ管理用のシートを作成します。  
[Googleドライブから環境をコピー](https://github.com/softbank-developer/gsuite_with_watson#環境構築)より、サンプルのシートを利用できます。  

4. "ログ"シートの作成  
作成したスプレッドシート内に、ログ保存用のシートを作成します。  
[Googleドライブから環境をコピー](https://github.com/softbank-developer/gsuite_with_watson#環境構築)より、サンプルのシートを利用できます。

5. "通知"シートの作成  
作成したスプレッドシート内に、メールで通知を行う設定用のシートを作成します。  
[Googleドライブから環境をコピー](https://github.com/softbank-developer/gsuite_with_watson#環境構築)より、サンプルのシートを利用できます。

6. GASスクリプトの読み込み  
スクリプトエディタを起動(ツール -> スクリプト エディタ)し、すべての.gsファイルをインポートします。
	- main.gs
	- NLCLIB.gs
	- RSSLIB.gs

7. NLCの属性値の設定  
	NLCの属性値をScript propertiesとして設定します(ファイル -> プロパティ)。
	- `CREDS_URL`: [NLC用のURL]
	- `CREDS_USERNAME`: [NLC用のユーザ名]
	- `CREDS_PASSWORD`: [NLC用のパスワード]


## 使い方
1. RSSによる情報取得  
スプレッドシートのメニューにWatsonが追加されています。
	- `Watson` -> `データ取得`を実行します。

2. NLCの分類器の学習  
	- データシートの`分類器[1-3]:手入力`へ学習させたいインテント名を付与(空白を含んだ行は無視)
	- `Watson` -> `学習`を実行  
	![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/rss/readme_images/menu.png)
	
	学習したNLCの分類IDは設定用のシートに自動で記録されます。**Training**中は利用できません。**Available**になるまで待つ必要があります。

3. データの分類
学習させたNLCの分類器を利用して分類させます。
	- `Watson` -> `分類`  

	分類結果はデータを記録するシートの分類器[1-3]:Watsonに記録されます。


## 詳細資料
https://softbank-developer.github.io/docs/


## ライセンス
[MIT](https://accounts.google.com/https://github.com/softbank-developer/gsuite_with_watson/blob/master/rss/LICENSE)
