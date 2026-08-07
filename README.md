本リポジトリは就職活動の際のポートフォリオ用として作成したGASスクリプト(スタンドアロン版)のものです。
(技術スタック:GAS,JS,html)

本リポジトリのファイルはGAS(Google Apps Script)のプロジェクト内のファイル(「README.md」と「.gitignore」を除く)となっています。
[各種ファイル解説]
「README.md」：本ファイル。
「appsscript.json」:本プロジェクトの設定ファイル。
「wa_env_config.js(gs)」:本プロジェクトの環境変数ファイル。
「weather-totalize_standalone.js(gs)」:バックエンド(Google Apps Script)のソースコード
「standalone.html」:トップページのベースとなるHTMLファイル(バックエンド側の関数より各種ページ内要素を描写。)
「style_sheet.html」:トップページのCSSにあたるファイル。(「standalone.html」に依存。)
「csv_upload.html」:トップページの内容(入力フォーム)の要素を含むHTMLファイル。(「standalone.html」に依存。)
「main_process」:処理中のプログレスバーを表示する要素を含むHTMLファイル。(「standalone.html」に依存。)
「java_script.html」:フロントエンド(JavaScript)のソースコード(「standalone.html」に依存。)

※本リポジトリのファイルを元にプロジェクトを作成する際には必ず、プロジェクト作成元ユーザーのドライブ内に「アップロードファイル一時保管用フォルダ」と「処理中または処理後のスプレッドシートを格納する処理用フォルダ」を作成しつつ、本プロジェクトの環境変数ファイル「wa_env_config.js(gs)」の「wheather_csv_temp_forder_id」変数に作成したアップロードファイル一時保管用フォルダのフォルダIDと「wheather_totalize_files_folder_id」変数に作成した処理用フォルダのフォルダIDを指定してから実行すること。
