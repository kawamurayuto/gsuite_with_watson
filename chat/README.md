## 概要
Googl Apps Scriptsで構築した問い合わせ対応を行うWEBアプリケーションです。問い合わせデータはスプレッドシートで管理されます。各データはNatural Language Classifier(NLC)を利用して任意のクラスタに分類させることができます。分類結果に任意の応答メッセージを結びつけることもできるため、自然な会話を演出すことが可能となります。

![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/chat/readme_images/web.png)

&emsp;↓

![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/chat/readme_images/data.png)

- `BOT受信日時`: 質問を受けた日時
- `入力メッセージ(学習・分類対象)`: 受け取った質問 
- `BOT応答メッセージ`: BOTの応答
- `分類器[1-3]:手入力`: Watsonに学習させるの手動での分類
- `分類器[1-3]:Watson`: Watsonによる分類
- `分類器[1-3]:処理日時`: Watsonが分類した日時
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
sheetsディレクトリ内に、サンプルのシートを置いています。

	(シート設定)
	- `データシート名`: 会話データの格納シート
	- `開始列`: データを格納する開始列 
	- `開始行`: データを格納する開始行
	- `学習・分類対象`: NLCによる学習・分類の対象
	- `分類器[1-3]:手入力`: 手動で分類させる列
	- `分類器[1-3]:watson`: 分類結果を記載する列
	- `分類日時[1-3]`: 分類された日時を記載する列
	- `ログシート名`:  NLCの学習と分類ログの保存シート
	- `応答設定シート名`: 分類結果に応じてで応答する文言を管理するシート
	- `該当なしメッセージ`: クラスタに分類されなかった質問への応答メッセージ 
	- `例外メッセージ`: 例外発生時にシステムが返す応答メッセージ
	- `アバターURL`: BOTのアイコンURL

	(分類器)
	- `Classifier ID`: NLCの分類器のID(学習後自動で挿入されます)
	- `ステータス`: NLCのステータス  
	![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/chat/readme_images/config.png)

3. "応答"シートの作成  
作成したスプレッドシート内に、応答メッセージを管理するシートを作成します。
sheetsディレクトリ内に、サンプルのシートを置いています。
	- `分類器[1-3]:Watson`: 応答メッセージを決める条件となるWatsonの分類
	- `分類器[1-3]:確信度`: 分類対象と判定するための確信度(この数値以上が対象)
	- `応答メッセージ`: 条件が合致した場合の応答に使うメッセージ        
	![img](https://github.com/softbank-developer/gsuite_with_watson/blob/master/chat/readme_images/answer.png)

4. "データ"シートの作成   
作成したスプレッドシート内に、会話のデータを管理するシートを作成します。  
sheetsディレクトリ内に、サンプルのシートを置いています。

5. "ログ"シートの作成  
作成したスプレッドシート内に、ログ保存用のシートを作成します。  
sheetsディレクトリ内に、サンプルのシートを置いています。

6. GASスクリプトの読み込み  
スクリプトエディタを起動(ツール -> スクリプト エディタ)し、すべての.gsファイルをインポートします。
	- main.gs
	- NLCLIB.gs
	- CHATLIB.gs
	- index.html

7. NLCの属性値の設定  
	NLCの属性値をScript propertiesとして設定します(ファイル -> プロパティ)。
	- `CREDS_URL`: [NLC用のURL]
	- `CREDS_USERNAME`: [NLC用のユーザ名]
	- `CREDS_PASSWORD`: [NLC用のパスワード]


## 使い方
1. WEBアプリケーションの起動  
スプレッドシートのメニューからスクリプト エディタを開き、スクリプト エディタからウェブアプリケーションとして公開します。
	- `ツール` -> `スクリプト エディタ`
	- `公開` -> `ウェブアプリケーションとして導入`  
	このWEBアプリケーションでの質問と回答がスプレッドシートのデータシートで管理されます。

2. NLCの分類器の学習  
質問に対してインテントを付与し、Watsonへ学習させます。
	- スプレッドシートの`データシート`を開き、`分類器[1-3]:手入力`へ学習させたいインテント名を付与(空白行は無視)
	- スプレッドシートのメニューから`Watson` -> `学習`  
	
	学習したNLCの分類IDは設定用のシートに自動で記録されます。**Training**中は利用できません。**Available**になるまで待つ必要があります。

3. データの分類
学習させたNLCの分類器を利用して分類させます。
	- `Watson` -> `分類`  
  分類結果はデータを記録するシートの分類器[1-3]:Watsonに記録されます。

4. 応答メッセージの作成  
分類器の結果に応答するメッセージを作成します。
	- スプレッドシートの`データ`シートを開き、`応答メッセージ`を作成します。  
WEBから質問を投げ、意図した応答メッセージが返されるか確認します。



## ライセンス
[MIT](https://accounts.google.com/https://github.com/softbank-developer/gsuite_with_watson/blob/master/rss/LICENSE)
