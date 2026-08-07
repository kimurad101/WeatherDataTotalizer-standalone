本リポジトリは就職活動の際のポートフォリオ用として作成したWebアプリのGASスクリプト(スタンドアロン版)のものです。<br>
(技術スタック:GAS,JS,html)<br>
<br>
[本Webアプリを開く](https://script.google.com/macros/s/AKfycbzE9-rrBWczBvXDD-1HJWN0mDYXWsXIzSbTtXKkbNTVFC3X8ZmqRs6dWWP2N4BY9bSRMQ/exec)<br>
<br>
[本Webアプリの使用方法(マニュアル)](https://drive.google.com/file/d/1XYsmolWC9d6blSSIVc_AoIbU6NjlsqMd/view?usp=sharing)<br>
<br>
本リポジトリのファイルはGAS(Google Apps Script)のプロジェクト内のファイル(「README.md」と「.gitignore」を除く)となっています。<br>
<br>
[各種ファイル解説]<br>
「README.md」：本ファイル。<br>
「appsscript.json」:本プロジェクトの設定ファイル。<br>
「wa_env_config.js(gs)」:本プロジェクトの環境変数ファイル。<br>
「weather-totalize_standalone.js(gs)」:バックエンド(Google Apps Script)のソースコード。<br>
「standalone.html」:トップページのベースとなるHTMLファイル(バックエンド側の関数より各種ページ内要素を描写。)<br>
「style_sheet.html」:トップページのCSSにあたるファイル。(「standalone.html」に依存。)<br>
「csv_upload.html」:トップページの内容(入力フォーム)の要素を含むHTMLファイル。(「standalone.html」に依存。)<br>
「main_process」:処理中のプログレスバーを表示する要素を含むHTMLファイル。(「standalone.html」に依存。)<br>
「java_script.html」:フロントエンド(JavaScript)のソースコード(「standalone.html」に依存。)<br>
<br>
※本リポジトリのファイルを元にプロジェクトを作成する際には必ず、プロジェクト作成元ユーザーのドライブ内に「アップロードファイル一時保管用フォルダ」と「処理中または処理後のスプレッドシートを格納する処理用フォルダ」を作成しつつ、本プロジェクトの環境変数ファイル「wa_env_config.js(gs)」の「wheather_csv_temp_forder_id」変数に作成したアップロードファイル一時保管用フォルダのフォルダIDと「wheather_totalize_files_folder_id」変数に作成した処理用フォルダのフォルダIDを指定してから実行すること。
