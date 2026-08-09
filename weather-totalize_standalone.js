/*
weather-totalize_standalone(backend) ver.0.0.2

ver.0.0.2[2026/8/5]:コンテナバインド版からスタンドアロン版を制作。
ver.0.0.1[2026/05/12~2026/8/5]:初版(コンテナバインド版)。

目標:必要最低限に動作できる状態を目指す。

参考元1:https://qiita.com/dadadaiiiiiii/items/beb33b838d0516cade05
参考元2:https://zenn.dev/flyoutstudio/articles/d2a308223b3815

*/

//トップページ表示
function doGet(){
  let index = HtmlService.createTemplateFromFile('standalone');
  return index.evaluate().setTitle('気象庁 過去の気象データ集計アプリ');
}

//ページ要素後読み込み
function include_file(filename){
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

//CSVファイルからスプレッドシートに変換。
//[旧最速記録(10年分)]1:10:59
//[完成時記録(10年分)]1:15:67
function ConvJmaCsv(file_name,totalize_book_fileid,add_yearly_table = false){
  let process_is_failed = false;
  let error_msg = "";
  let process_sheet_name = ""; 
  let conv_filename = undefined;
  let ob_area_name = "";
  let uniq_date_list = undefined;
  let process_lock = undefined;

  try{
    const uploaded_csv_file = DriveApp.getFilesByName(file_name);

    //ディレクトリ内ファイルの有無チェック
    if(uploaded_csv_file.hasNext() == false){
      //フォルダ内空っぽ。
      //エラーメッセージを表示して処理を中止する。
      throw "指定されたCSVファイルが見つかりませんでした。";
    }

    //CSVファイル取得
    const process_csv_file = uploaded_csv_file.next();
    //const process_csv_file_id = process_csv_file.getId();
    const process_csv_file_name = process_csv_file.getName();
    const process_csv_file_mimetype = process_csv_file.getMimeType();
    //const get_csv_file_blob = DriveApp.getFileById(process_csv_file_id).getBlob();
    const get_csv_file_blob = process_csv_file.getBlob();

    //アップロード済みのCSVファイルの本来のファイル名(セッションNo&拡張子抜き)をシート名とする。
    //サンプルファイル名「341260629104314_temp_data(1).csv」
    process_sheet_name = process_csv_file_name.substring(21,process_csv_file_name.lastIndexOf("."));

    //CSVファイルの中身をエンコード形式「ShiftJIS」で読み込む。
    const csv_blob = get_csv_file_blob.getDataAsString('Shift_JIS');
    let csv_values = undefined;

    //中身チェック
    if(csv_blob.trim() === "" || process_csv_file_mimetype != "text/csv"){
      //空ファイルまたはCSVファイルではない場合は読み込めなかったファイル名として記録し、現在のシートを消して次のファイルを読み込む。
      //load_failed_files.push(process_sheet_name);
      throw "ファイルを読み込めませんでした。";

    }else{
      //CSVデータをスプレッドシートに転写
      csv_values = Utilities.parseCsv(csv_blob);
      if(csv_values.length == 0){
        throw "無効なCSVデータです。処理を中止します。"
      }else{
        csv_parse_is_done = true;
      }

    }

    //最後行削除(内容が重複する可能性があるのと、最後の日付のデータが1時間分しかないため)
    csv_values.pop();

    //年月範囲をユニークに取得
    const date_regex = /^\d{4}\/\d{1,2}\/\d{1,2} \d{1,2}:\d{2}:\d{2}/; //「YYYY/MM/DD HH:mm:ss」
    const date_sort = csv_values.filter(v => date_regex.test(v[0]));

    //観測地点名をユニーク取得
    ob_area_name = [...new Set(csv_values[2])].filter((v) => v !== "").join(",");
    Logger.log(ob_area_name);

    //全般的な年月範囲のユニーク出し&配列の先頭と末尾のみを具体的な日時のものに書き換え。
    uniq_date_list = [...new Set(date_sort.map(s => s[0].match(/^\d{4}\/\d{1,2}/)[0]))].filter((v) => v !== null); //一旦ユニーク化して絞り込む。
    Logger.log(uniq_date_list);
    uniq_date_list[uniq_date_list.length - 1] = date_sort[date_sort.length - 1][0];
    uniq_date_list[0] = date_sort[0][0];

    const uniq_date_count = uniq_date_list.length

    //一時ブック作成
    let conv_book = SpreadsheetApp.create(process_csv_file_name.substring(0,process_csv_file_name.lastIndexOf(".")));
    const conv_fileid = conv_book.getId();
    conv_filename = conv_book.getName();

    let conv_csv_sheet = conv_book.getActiveSheet();

    //CSV内容貼り付け
    conv_csv_sheet.getRange(1, 1, csv_values.length,csv_values[0].length).setValues(csv_values);

    //ブック場所移動
    const book_file = DriveApp.getFileById(conv_fileid);
    const temp_forder = DriveApp.getFolderById(wheather_csv_temp_forder_id);
    book_file.moveTo(temp_forder);

    //本シートに貼り付け
    //ロック取得&待機処理
    process_lock = LockService.getScriptLock();
    while(process_lock.tryLock(1000) === false){
      Utilities.sleep(1000);
      Logger.log("待機中...");
    }
    Logger.log("実行");
    //集計ブックに貼り付け&シート名設定
    let totalize_book = SpreadsheetApp.openById(totalize_book_fileid);
    const csv_sheet = conv_csv_sheet.copyTo(totalize_book).setName(process_sheet_name);
    const result_sheet = totalize_book.getSheetByName("結果");
    result_sheet.activate();
    
    //集計
    //CSVシートの最後行数を取得。
    const csv_sheet_max_row = csv_sheet.getLastRow();

    //数式自動生成
    function generate_jma_totalize_cell_formulas(column_type,func_name,date_str,sheet_name,max_row){
      const datetime = new Date(date_str);//「YYYY/MM/DD HH:mm:ss」or「YYYY/MM」
      const year_str = datetime.getFullYear();
      const month_str = datetime.getMonth() + 1;
      const day_str = datetime.getDate();
      const hour_str = datetime.getHours();
      const e_hour_str = 23;
      const end_of_month = new Date(year_str,month_str,0,e_hour_str,0,0);
      let s_day_str = day_str;
      let s_hour_str = hour_str;
      const e_day_str = end_of_month.getDate();

      let totalize_position_num = 0;
      let args_str = "";
      let iferror_func_strs = ["",""];

      //最終年月のデータ行の場合、月初日の日時を反映。
      if(datetime.toDateString() === end_of_month.toDateString()){
        const start_of_month = new Date(year_str,datetime.getMonth(),1,0,0,0);
         s_day_str = start_of_month.getDate();
         s_hour_str = 0;
      }

      if(column_type === "temp"){
        totalize_position_num = 2; //
      }else if(column_type === "rainfall"){
        totalize_position_num = 5;
      }else if(column_type === "sunshine"){
        totalize_position_num = 9;
      }else if(column_type === "airpressure"){
        totalize_position_num = 13;
      }else{
        throw "[Conv]: タイプが不正です。"
      }

      //平均値を集計する際に「値:0」を集計対象外にし、ゼロ除算の場合は0を返すようにする。
      if(func_name.toUpperCase() === "AVERAGEIF"){
        //0を集計しないようにするため
        args_str = ',\"<>0\"';
        //ゼロ除算対策
        iferror_func_strs = ["IFERROR(",",0)"];
      }

     //数式を出力
      let function_str = "=" + iferror_func_strs[0] + func_name.toUpperCase() +"(INDIRECT(\"\'" + sheet_name + "\'\!\"&ADDRESS(MATCH(DATE(" + year_str + "," + month_str + "," + s_day_str + ")+TIME("+ s_hour_str +",0,0),\'" + sheet_name + "\'!$A$1:$A$" + max_row + "),"+ totalize_position_num +")&\":\"&ADDRESS(MATCH(DATE(" + year_str + "," + month_str + "," + e_day_str + ")+TIME("+ e_hour_str +",0,0),\'" + sheet_name + "\'!$A$1:$A$" + max_row + "),"+ totalize_position_num +"))"+ args_str +")"+ iferror_func_strs[1];
 
      return function_str;
    }

    //セル反映
    //行生成
    const totalize_cells = uniq_date_list.map((d) => {
      return [d.match(/^(\d{4})\/(\d{1,2})/)[0].replace(/^(\d{4})\/(\d{1,2})/,'$1年$2月'),generate_jma_totalize_cell_formulas("temp","MAX",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("temp","MIN",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("temp","AVERAGEIF",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("rainfall","MAX",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("rainfall","AVERAGEIF",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("sunshine","AVERAGEIF",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("airpressure","MAX",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("airpressure","MIN",d,process_sheet_name,csv_sheet_max_row),generate_jma_totalize_cell_formulas("airpressure","AVERAGEIF",d,process_sheet_name,csv_sheet_max_row)]
    });

    //書き出し
    const append_start_cell_row = result_sheet.getLastRow() + 1;
    Logger.log("追記開始行: " + append_start_cell_row);
    result_sheet.getRange(append_start_cell_row,1,totalize_cells.length,totalize_cells[0].length).setValues(totalize_cells);
    
    //年別集計処理(ファイルタイプがExcel[xlsx]の場合のみ。)
    if(add_yearly_table){
      Logger.log("年別シート作成中...");

      //年別集計シート作成
      const yearly_start_date = new Date(uniq_date_list[0]);
      const yearly_end_date = new Date(uniq_date_list[uniq_date_list.length -1]);
      const yearly_start_year_str = yearly_start_date.getFullYear();
      const yearly_start_month_str = yearly_start_date.getMonth() +1;
      const yearly_end_year_str =  yearly_end_date.getFullYear();
      const yearly_end_month_str = yearly_end_date.getMonth() +1;
      const yearly_table_name = yearly_start_year_str + "/" + yearly_start_month_str + "~" + yearly_end_year_str + "/" + yearly_end_month_str;
      const yearly_table_sheet_name = yearly_start_year_str + "." + yearly_start_month_str + "~" + yearly_end_year_str + "." + yearly_end_month_str;

      let yearly_table_sheet = totalize_book.insertSheet(yearly_table_sheet_name,1);
      yearly_table_sheet.activate();

      //結果シートから必要なもの(ヘッダー&集計表)部分をコピー
      result_sheet.getRange(1,1,1,10).copyTo(yearly_table_sheet.getRange(1,1,1,10)); //グラフ部分ヘッダー
      const totalize_start_cell_row = 31;
      result_sheet.getRange(totalize_start_cell_row,1,2,totalize_cells[0].length).copyTo(yearly_table_sheet.getRange(totalize_start_cell_row,1,2,totalize_cells[0].length));//集計ページヘッダー部分
      result_sheet.getRange(append_start_cell_row,1,totalize_cells.length,totalize_cells[0].length).copyTo(yearly_table_sheet.getRange(totalize_start_cell_row + 2,1,totalize_cells.length,totalize_cells[0].length)); //集計ページ表部分

      //結果シート内の集計結果にフィルタをつけて、昇順に並べ替える
      yearly_table_sheet.getRange(totalize_start_cell_row + 1,1,totalize_cells.length,totalize_cells[0].length).createFilter();

      //グラフページ - ヘッダー行結合(タイトル&対象期間&処理日時)
      yearly_table_sheet.getRange(1,1,1,2).mergeAcross();
      yearly_table_sheet.getRange(1,4,1,2).mergeAcross();
      yearly_table_sheet.getRange(1,7,1,2).mergeAcross().setHorizontalAlignment("left");

      //グラフページ - ヘッダーテキスト位置調整
      yearly_table_sheet.getRange(1,3).setHorizontalAlignment("right");
      yearly_table_sheet.getRange(1,6).setHorizontalAlignment("right");
      yearly_table_sheet.getRange(1,9).setHorizontalAlignment("right");

      //集計ページ - ヘッダー行結合(タイトル&対象期間&処理日時)
      yearly_table_sheet.getRange(totalize_start_cell_row,1,1,2).mergeAcross();
      yearly_table_sheet.getRange(totalize_start_cell_row,4,1,2).mergeAcross();
      yearly_table_sheet.getRange(totalize_start_cell_row,7,1,2).mergeAcross().setHorizontalAlignment("left");

      //グラフページ - ヘッダーテキスト位置調整
      yearly_table_sheet.getRange(totalize_start_cell_row,3).setHorizontalAlignment("right");
      yearly_table_sheet.getRange(totalize_start_cell_row,6).setHorizontalAlignment("right");
      yearly_table_sheet.getRange(totalize_start_cell_row,9).setHorizontalAlignment("right");

      //書式動的変更
      const yearly_totalize_end_cell_row = yearly_table_sheet.getLastRow();
      yearly_table_sheet.getRange(totalize_start_cell_row + 2,2,yearly_totalize_end_cell_row,4).setNumberFormat('0.0℃'); //気温
      yearly_table_sheet.getRange(totalize_start_cell_row + 2,5,yearly_totalize_end_cell_row,6).setNumberFormat('0.0"mm"'); //降水量
      yearly_table_sheet.getRange(totalize_start_cell_row + 2,7,yearly_totalize_end_cell_row,7).setNumberFormat('0.0時間'); //日照時間
      yearly_table_sheet.getRange(totalize_start_cell_row + 2,8,yearly_totalize_end_cell_row,10).setNumberFormat('0.0"hPa"'); //気圧

      //罫線動的描写
      yearly_table_sheet_last_rows = yearly_table_sheet.getLastRow();
      yearly_table_sheet.getRange(totalize_start_cell_row +1,1,uniq_date_count +1,10).setBorder(true,true,true,true,true,true);

      const send_config_obj = {"download_filetype": "xlsx","total_data_count": uniq_date_list.length,"page_title": yearly_table_name,"append_verbose_graphs_pages":false};
      GenerateTotalizeGraphsPage(yearly_table_sheet,send_config_obj);
      
      //観測地表記引用
      yearly_table_sheet.getRange(1,10).setFormula("='" + result_sheet.getSheetName() +"'!J1");//='結果'!J1

      //観測地表記引用
      yearly_table_sheet.getRange(totalize_start_cell_row,10).setFormula("='" + result_sheet.getSheetName() +"'!J1");//='結果'!J1
    }

  }catch(e){
    console.log(e.message);
    Logger.log(e.message);
    Logger.log(e.stack);
    error_msg = e.message;
    process_is_failed = true;

  }finally{
    //アップロード済みのCSVファイルを削除する。
    const csv_file_delete_result = DeleteUploadFile(file_name);
    if(!csv_file_delete_result){
      Logger.log("Failed delete. Filename: " + file_name);
    }

    //変換後のスプレッドシートも削除する。
    if(conv_filename !== undefined){
      const conv_file_delete_result = DeleteUploadFile(conv_filename);
      if(!conv_file_delete_result){
        Logger.log("Failed delete. Filename: " + conv_filename);
      }  
    }

    //ロック解放
    process_lock.releaseLock();
    Logger.log("完了");

    //結果出力
    if(process_is_failed){
      const e_msg = "[" + file_name + "]: "+ error_msg;
      throw e_msg;
    }else{
      return {"observation_point": ob_area_name,"totalize_period_count": uniq_date_list.length};
    }

  }
}

//加工処理
//処理後のファイル作成
function MakeTotalizationWheaterBook(session_num){
  Logger.log("session_number: " + session_num);

  //集計ファイル作成
  let main_book = SpreadsheetApp.create("temp_" + session_num);

  //一時処理フォルダに移動
  const templete_ss_file = DriveApp.getFileById(main_book.getId());
  const process_temp_forder = DriveApp.getFolderById(wheather_totalize_files_folder_id); //一時処理フォルダ
  //const process_ss_file = templete_ss_file.makeCopy("temp_" + session_num,process_temp_forder);
  const process_ss_file = templete_ss_file.moveTo(process_temp_forder);
  
  //結果シート作成
  let totalize_result_sheet = main_book.getSheets()[0].activate();
  totalize_result_sheet.setName("結果");

  //セル反映
  const result_sheet_graph_page_last_row = 30; //グラフページ最終行
  const result_sheet_first_totalize_page_last_row = 60; //集計表ページ 1ページ目最終行
  const result_sheet_totalize_page_max_row = 30; //1ページに収まる最大行数

  const process_datetime = new Date();
  const wheather_analysis_datetime_str = process_datetime.getFullYear() + "/" + (process_datetime.getMonth() + 1) + "/" + process_datetime.getDate() + " " + process_datetime.getHours() + ":" + process_datetime.getMinutes() + ":" + process_datetime.getSeconds();
  const totalize_book_header = ["気象庁 - 気象データ集計結果","","対象期間:","","","処理日時:",wheather_analysis_datetime_str,"","観測地点名:",""]; //共通ヘッダー
  const result_sheet_first_totailze_page_headers = [totalize_book_header,["年/月","最高気温","最低気温","平均気温","最大降水量","平均降水量","平均日照時間","最高気圧","最低気圧","平均気圧"]];//集計ページ
  totalize_result_sheet.getRange(1,1,1,10).setValues([totalize_book_header]); //グラフページ用

  //グラフページ - ヘッダー行結合(タイトル&対象期間&処理日時)
  totalize_result_sheet.getRange(1,1,1,2).mergeAcross();
  totalize_result_sheet.getRange(1,4,1,2).mergeAcross();
  totalize_result_sheet.getRange(1,7,1,2).mergeAcross().setHorizontalAlignment("left");

  //グラフページ - ヘッダーテキスト位置調整
  totalize_result_sheet.getRange(1,3).setHorizontalAlignment("right");
  totalize_result_sheet.getRange(1,6).setHorizontalAlignment("right");
  totalize_result_sheet.getRange(1,9).setHorizontalAlignment("right");

  //集計ページ - 専用ヘッダー代入
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,1,2,10).setValues(result_sheet_first_totailze_page_headers); //集計ページ1ページ目

  //集計ページ - ヘッダー行結合(タイトル&対象期間&処理日時)
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,1,1,2).mergeAcross();
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,4,1,2).mergeAcross();
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,7,1,2).mergeAcross().setHorizontalAlignment("left");

  //グラフページ - ヘッダーテキスト位置調整
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,3).setHorizontalAlignment("right");
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,6).setHorizontalAlignment("right");
  totalize_result_sheet.getRange(result_sheet_graph_page_last_row + 1,9).setHorizontalAlignment("right");

  return process_ss_file.getId();
}

//グラフ生成
function GenerateTotalizeGraphsPage(sheet,recv_config_obj){
  const download_filetype = recv_config_obj.download_filetype;
  const total_data_count = recv_config_obj.total_data_count;
  const page_title = recv_config_obj.page_title;
  const append_verbose_graphs_pages = recv_config_obj.append_verbose_graphs_pages;

  const result_sheet_totalize_page_max_row = 30; //１ページにつき使用できる行数は30行(集計1ページ目で31~60)、ヘッダー2行分を除くと28行。
  const max_totalize_pages = Math.ceil(total_data_count / (result_sheet_totalize_page_max_row - 2)); //分割が必要な集計ページ総数(小数点切り上げ[小数点分=+1ページ])
  let graph_dynamic_params_list = []; //分割後各ページ最終行リスト

  Logger.log("total pages: " + max_totalize_pages);

  try{
    const last_row_number = sheet.getLastRow();
    //31ほどずれてしまう原因は、最終行番号を相対行数に当ててしまっているため。実データの長さを相対行数に当てるようにする。
    const date_cell_values = sheet.getRange(33,1,last_row_number,1).getValues().filter(v => v[0].length > 0 || v[0] !== "");
    let max_data_count = date_cell_values.length;

    //集計期間表記
    Logger.log(date_cell_values[date_cell_values.length - 1][0]);
    const start_cell_datetime_str = date_cell_values[0][0].getFullYear() + "/" + (date_cell_values[0][0].getMonth() + 1);
    const last_cell_datetime_str = date_cell_values[date_cell_values.length - 1][0].getFullYear() + "/" + (date_cell_values[date_cell_values.length - 1][0].getMonth() + 1);
    const totalize_date_str = start_cell_datetime_str + "~" + last_cell_datetime_str;

    sheet.getRange(1,4).setValue(totalize_date_str); //グラフページ
    sheet.getRange(31,4).setValue(totalize_date_str); //集計ページ

    //グラフページタイトル
    sheet.getRange(2,1,1,10).mergeAcross().setHorizontalAlignment('center').setFontWeight('bold').setBorder(true,true,true,true,null,null).setValue(page_title); 

    //出力種別がPDFの場合　毎ページヘッダー動的生成
    if(download_filetype === "pdf"){
      //デザインサンプルたたき台:https://docs.google.com/spreadsheets/d/1spLYrciog96K60JFFEAHOASGAxarMJICb6Q5l7KHZ5E/edit?gid=161207033#gid=161207033
      //集計ページ分割行数算出(グラフ範囲指定時に使用。)
      /*
      １ページにつき使用できる行数は30行(集計１ページ目で31~60)、ヘッダー2行分を除くと28行。
      各ページのセル開始行は31行。
      各ページのセル開始行ごとにページ数ごとに30を加算すれば良い。(1ページごとの行数は30)
      */

      //1ページに収まらない(27行以上の)場合、次の集計ページを生成。
      if(max_totalize_pages > 1){
        let left_move_param_cell_count = total_data_count - (result_sheet_totalize_page_max_row - 3);//残り移動セル数
        const copy_header_cells_range = sheet.getRange(result_sheet_totalize_page_max_row + 1,1,2,10); //コピー対象のヘッダーセル範囲
        const max_data_cell_count = 28; //ヘッダー行分を除いたデータ行の最大行数

        //各ページの開始行算出のためにあえて0スタート。
        for(page_cnt = 0; page_cnt < max_totalize_pages; page_cnt++){
          const start_dynamic_row_count = result_sheet_totalize_page_max_row * (1 + page_cnt);
          const range_last_cell_row_number_offset = 1;
          const move_start_cell_row_number_offset = 3;
          
          const move_params_src_cells_range = sheet.getRange(start_dynamic_row_count + range_last_cell_row_number_offset,1,left_move_param_cell_count + 1,10); //移動対象の集計セル範囲

          const move_params_dst_cells_range = sheet.getRange(start_dynamic_row_count + move_start_cell_row_number_offset,1); //移動先のセル

          //グラフ 分割後集計ページ位置記録
          const now_start_row_number = start_dynamic_row_count + move_start_cell_row_number_offset;
          const dynamic_last_row_number = last_row_number + (page_cnt * 2); //2はページごとに追記されるヘッダー行分
          if(page_cnt == max_totalize_pages - 1){
            //最終ページ
            Logger.log("dynamic_last_row_number: "+ dynamic_last_row_number);
            Logger.log("now_start_row_number: " + now_start_row_number);
            graph_dynamic_params_list.push({"start_row_number":now_start_row_number,"max_data_count": (dynamic_last_row_number - now_start_row_number) + 1});
            
          }else{
            graph_dynamic_params_list.push({"start_row_number":now_start_row_number,"max_data_count": (now_start_row_number + (result_sheet_totalize_page_max_row - 2) - now_start_row_number)});
          }

          //集計2ページ目から
          if(page_cnt > 0){
            //既に記述されているセルを新たなヘッダー行分ずらして移動する。
            move_params_src_cells_range.moveTo(move_params_dst_cells_range);
            if(left_move_param_cell_count > max_data_cell_count){
              left_move_param_cell_count = left_move_param_cell_count - max_data_cell_count; //相対行数を移動した分減らす。
            }
            //ヘッダー行を新しいページにコピー
            const range_copy_header_row_offset = 1;
            copy_header_cells_range.copyTo(sheet.getRange(start_dynamic_row_count + range_copy_header_row_offset,1));
          }
        }

      }else{
        graph_dynamic_params_list.push({"start_row_number":result_sheet_totalize_page_max_row,"max_data_count": total_data_count});
      }

    }

    //結果シート内の集計結果の書式を変更。
    const sheet_last_columns = sheet.getLastColumn();
    Logger.log(graph_dynamic_params_list);

    //セル幅調整
    sheet.setColumnWidths(1,sheet_last_columns,90);
    sheet.setColumnWidth(1,127);
    sheet.setColumnWidth(2,80);
    sheet.setColumnWidth(7,107);

    //各種折れ線グラフシート
    //参考元:https://developers.google.com/apps-script/reference/spreadsheet/embedded-chart-builder?hl=ja

    //グラフ生成下準備
    let date_graph_x_range_lists = undefined;
    let temp_graph_y_range_lists = undefined;
    let rainfall_graph_y_range_lists = undefined;
    let sunshine_graph_y_range_lists = undefined;
    let airpressure_graph_y_range_lists = undefined;
    let graph_size_width = undefined;
    let graph_size_height = undefined;

    //ファイル種別ごとのグラフサイズ調整
    if(download_filetype === "xlsx"){
      graph_size_width = 423; //-10%
      graph_size_height = 225;  //-10%
    }else{
      graph_size_width = 470;
      graph_size_height = 250;
    }

    let prepare_temp_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(4,1,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1);
    let prepare_rainfall_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(4,6,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1);
    let prepare_sunshine_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(18,1,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1);
    let prepare_airpressure_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(18,6,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1).setOption('vAxis.maxValue',1050).setOption('vAxis.minValue',970);

    if(download_filetype === "pdf"){ 
      //グラフ範囲動的取得(詳細ページオプション有効時のみ)
      if(append_verbose_graphs_pages && max_data_count > result_sheet_totalize_page_max_row - 2){
        //PDF処理時オプション動作時(グラフページ追記用下処理)
        Logger.log("グラフ - 分割範囲指定中。"); //A33とB33しか指定されていない

        Logger.log("動的取得開始行: " + result_sheet_totalize_page_max_row);
        //対象範囲を配列にして、mapメソッドで代入&範囲を追加させる。
      
        const start_param_cell_row = result_sheet_totalize_page_max_row + 2; //先頭(1ページ目)の範囲
        let date_graph_x_row_list = ["R" + start_param_cell_row + "C" + 1]; //先頭(1ページ目)の範囲
        let temp_graph_y_row_list = ["R" + start_param_cell_row + "C" + 2]; //先頭(1ページ目)の範囲
        let rainfall_graph_y_row_list = ["R" + start_param_cell_row + "C" + 5]; //先頭(1ページ目)の範囲
        let sunshine_graph_y_row_list = ["R" + start_param_cell_row + "C" + 7]; //先頭(1ページ目)の範囲
        let airpressure_graph_y_row_list = ["R" + start_param_cell_row + "C" + 8]; //先頭(1ページ目)の範囲
        Logger.log("各開始行: " + start_param_cell_row);

        let append_list_count = 0;
        graph_dynamic_params_list.forEach((r) => {
          const start_row_position = r.start_row_number - 1;
          const data_count = r.max_data_count;

          if(append_list_count == 0){
            //初回の1ページ分のみ、範囲を追記。
            date_graph_x_row_list[append_list_count] = date_graph_x_row_list[append_list_count] + ":" + "R" + (start_row_position + data_count) + "C" + 1;
            temp_graph_y_row_list[append_list_count] = temp_graph_y_row_list[append_list_count] + ":" + "R" + (start_row_position + data_count) + "C" + 4;
            rainfall_graph_y_row_list[append_list_count] = rainfall_graph_y_row_list[append_list_count] + ":" + "R" + (start_row_position + data_count) + "C" + 6;
            sunshine_graph_y_row_list[append_list_count] = sunshine_graph_y_row_list[append_list_count] + ":" + "R" + (start_row_position + data_count) + "C" + 7;
            airpressure_graph_y_row_list[append_list_count] = airpressure_graph_y_row_list[append_list_count] + ":" + "R" + (start_row_position + data_count) + "C" + 10;
          }else{
            //2ページ以降は範囲を追加。
            date_graph_x_row_list.push("R" + start_row_position + "C" + 1 + ":" + "R" + (start_row_position + data_count) + "C" + 1);
            temp_graph_y_row_list.push("R" + start_row_position + "C" + 2 + ":" + "R" + (start_row_position + data_count) + "C" + 4);
            rainfall_graph_y_row_list.push("R" + start_row_position + "C" + 5 + ":" + "R" + (start_row_position + data_count) + "C" + 6);
            sunshine_graph_y_row_list.push("R" + start_row_position + "C" + 7 + ":" + "R" + (start_row_position + data_count) + "C" + 7);
            airpressure_graph_y_row_list.push("R" + start_row_position + "C" + 8 + ":" + "R" + (start_row_position + data_count) + "C" + 10);
          }
          append_list_count++;
        });

        date_graph_x_range_lists = sheet.getRangeList(date_graph_x_row_list).getRanges();
        temp_graph_y_range_lists = sheet.getRangeList(temp_graph_y_row_list).getRanges();
        rainfall_graph_y_range_lists = sheet.getRangeList(rainfall_graph_y_row_list).getRanges();
        sunshine_graph_y_range_lists = sheet.getRangeList(sunshine_graph_y_row_list).getRanges();
        airpressure_graph_y_range_lists = sheet.getRangeList(airpressure_graph_y_row_list).getRanges();
      }

      //PDF時のみ、分割後各ページヘッダー分(2行分)を加算。
      max_data_count = max_data_count + (2 * (max_totalize_pages - 1) + 1); //末尾の「+1」は誤差調整。
    
    }else{
      //Excel(xlsx)選択時、グラフ範囲が1足りずにズレてしまうため調整。
      max_data_count++;
    }

    //共通グラフ
    const date_graph_x_range = sheet.getRange(32,1,max_data_count);
    const temp_graph_y_range = sheet.getRange(32,2,max_data_count,3);
    const rainfall_graph_y_range = sheet.getRange(32,5,max_data_count,2);
    const sunshine_graph_y_range = sheet.getRange(32,7,max_data_count);
    const airpressure_graph_y_range = sheet.getRange(32,8,max_data_count,3);

    prepare_temp_graph_chart.addRange(date_graph_x_range).addRange(temp_graph_y_range);
    prepare_rainfall_graph_chart.addRange(date_graph_x_range).addRange(rainfall_graph_y_range);
    prepare_sunshine_graph_chart.addRange(date_graph_x_range).addRange(sunshine_graph_y_range);
    prepare_airpressure_graph_chart.addRange(date_graph_x_range).addRange(airpressure_graph_y_range);

    const temp_graph_chart = prepare_temp_graph_chart.build();
    const rainfall_graph_chart = prepare_rainfall_graph_chart.build();
    const sunshine_graph_chart = prepare_sunshine_graph_chart.build();
    const airpressure_graph_chart = prepare_airpressure_graph_chart.build();

    //各グラフタイトル
    sheet.getRange(3,1,1,4).mergeAcross().setHorizontalAlignment('center').setFontWeight('bold').setBorder(true,true,true,null,null,null).setValue("「気温」");
    sheet.getRange(3,6,1,5).mergeAcross().setHorizontalAlignment('center').setFontWeight('bold').setBorder(true,null,true,true,null,null).setValue("「降水量」");

    sheet.getRange(17,1,1,4).mergeAcross().setHorizontalAlignment('center').setFontWeight('bold').setBorder(true,true,true,null,null,null).setValue("「平均日照時間」");
    sheet.getRange(17,6,1,5).mergeAcross().setHorizontalAlignment('center').setFontWeight('bold').setBorder(true,null,true,true,null,null).setValue("「気圧」");

    //各グラフタイトル中間罫線
    sheet.getRange(3,5).setBorder(true,null,true,null,null,null);
    sheet.getRange(17,5).setBorder(true,null,true,null,null,null);

    //グラフ生成
    sheet.insertChart(temp_graph_chart);
    sheet.insertChart(rainfall_graph_chart);
    sheet.insertChart(sunshine_graph_chart);
    sheet.insertChart(airpressure_graph_chart);
    
    //PDF指定でなおかつオプション指定があったとき。(ただし集計ページが1ページのみの場合は除く。)
    Logger.log(append_verbose_graphs_pages);
    if(download_filetype === "pdf" && append_verbose_graphs_pages){
      if(max_totalize_pages > 1){
        //2ページ以上の場合のみ実行。
        const send_param_obj ={
          "file_type": download_filetype,
          "page_list": graph_dynamic_params_list,
          "append_date_x_ranges": date_graph_x_range_lists,
          "append_temp_y_ranges": temp_graph_y_range_lists,
          "append_rainfall_y_ranges": rainfall_graph_y_range_lists,
          "append_sunshine_y_ranges": sunshine_graph_y_range_lists,
          "append_airpressure_y_ranges":airpressure_graph_y_range_lists
        };
        GenerateAppendGraphsPage(sheet,send_param_obj);
      }
    }
    
  }catch(e){
    Logger.log(e.message);
    throw "[GAS_Graph]: " + e.message;
  }
}

//グラフページ追加(PDF指定時のみのオプション)
function GenerateAppendGraphsPage(sheet,recv_config_obj){
  const page_max_row = 30;
  const download_filetype = recv_config_obj.file_type;
  const split_page_list = recv_config_obj.page_list;
  const append_date_graph_x_range_lists = recv_config_obj.append_date_x_ranges;
  const append_temp_graph_y_range_lists = recv_config_obj.append_temp_y_ranges;
  const append_rainfall_graph_y_range_lists = recv_config_obj.append_rainfall_y_ranges;
  const append_sunshine_graph_y_range_lists = recv_config_obj.append_sunshine_y_ranges;
  const append_airpressure_graph_y_range_lists = recv_config_obj.append_airpressure_y_ranges;

  try{
    //コピー元となるヘッダーとグラフタイトルセルの範囲を指定。
    const header_and_titles_copy_from_range = sheet.getRange(1,1,30,10);
    let now_last_row_number = sheet.getLastRow();
    Logger.log("処理前の末尾行: " + now_last_row_number);

    //分割した集計ページ分、処理を実行。
    Logger.log("最大処理数: " + split_page_list.length);
    for(generate_graph_page_count = 0; generate_graph_page_count < split_page_list.length; generate_graph_page_count++){
      const start_row_number = split_page_list[generate_graph_page_count].start_row_number;
      const max_data_count = split_page_list[generate_graph_page_count].max_data_count;

      //現在の最終行を指定
      now_last_row_number = (page_max_row * (Math.floor(now_last_row_number / page_max_row) + 1))+1;

      Logger.log("処理中の末尾行: " + now_last_row_number + "(" + generate_graph_page_count + ")");
      //現在の末尾行より1つ下にヘッダーや各グラフタイトルをコピー。
      header_and_titles_copy_from_range.copyTo(sheet.getRange(now_last_row_number +1 ,1));

      //集計期間表記
      const date_cell_values = sheet.getRange(start_row_number,1,28,1).getValues().filter(v => v[0].length > 0 || v[0] !== "");
      const start_cell_datetime_str = date_cell_values[0][0].getFullYear() + "/" + (date_cell_values[0][0].getMonth() + 1);
      const last_cell_datetime_str = date_cell_values[date_cell_values.length - 1][0].getFullYear() + "/" + (date_cell_values[date_cell_values.length - 1][0].getMonth() + 1);

      const totalize_date_str = start_cell_datetime_str + "~" + last_cell_datetime_str;

      //シートタイトル変更
      sheet.getRange(now_last_row_number + 1,1).setValue(totalize_date_str); // YYYY/MM~YYYY/MM

      //ファイル種別ごとのグラフサイズ調整
      if(download_filetype === "xlsx"){
        graph_size_width = 423; //-10%
        graph_size_height = 225;  //-10%
      }else{
        graph_size_width = 470;
        graph_size_height = 250;
      }

      //各種折れ線グラフシート
      //グラフ生成下準備
      let prepare_temp_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(now_last_row_number + 4,1,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1);
      let prepare_rainfall_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(now_last_row_number + 4,6,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1);
      let prepare_sunshine_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(now_last_row_number + 18,1,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1);
      let prepare_airpressure_graph_chart = sheet.newChart().setChartType(Charts.ChartType.LINE).setPosition(now_last_row_number + 18,6,0,0).setOption('width',graph_size_width).setOption('height',graph_size_height).setNumHeaders(1).setOption('vAxis.maxValue',1050).setOption('vAxis.minValue',970);

      //グラフ横範囲指定
      prepare_temp_graph_chart.addRange(append_date_graph_x_range_lists[generate_graph_page_count]);
      prepare_rainfall_graph_chart.addRange(append_date_graph_x_range_lists[generate_graph_page_count]);
      prepare_sunshine_graph_chart.addRange(append_date_graph_x_range_lists[generate_graph_page_count]);
      prepare_airpressure_graph_chart.addRange(append_date_graph_x_range_lists[generate_graph_page_count]);

      //グラフ縦範囲指定
      prepare_temp_graph_chart.addRange(append_temp_graph_y_range_lists[generate_graph_page_count]);
      prepare_rainfall_graph_chart.addRange(append_rainfall_graph_y_range_lists[generate_graph_page_count]);
      prepare_sunshine_graph_chart.addRange(append_sunshine_graph_y_range_lists[generate_graph_page_count]);
      prepare_airpressure_graph_chart.addRange(append_airpressure_graph_y_range_lists[generate_graph_page_count]);

      const temp_graph_chart = prepare_temp_graph_chart.build();
      const rainfall_graph_chart = prepare_rainfall_graph_chart.build();
      const sunshine_graph_chart = prepare_sunshine_graph_chart.build();
      const airpressure_graph_chart = prepare_airpressure_graph_chart.build();

      //グラフ生成
      sheet.insertChart(temp_graph_chart);
      sheet.insertChart(rainfall_graph_chart);
      sheet.insertChart(sunshine_graph_chart);
      sheet.insertChart(airpressure_graph_chart);

    }
  }catch(e){
    Logger.log(e.message);
    throw "[GAS_Graph]: " + e;
  }
}

//総集計処理
function TotalizationWheatherParameters(recv_param_json_obj){
  //やる処理1:テンプレファイルからコピーして、コピーしたものを主に加工する。
  //やる処理2:年と月の範囲を特定し、年と月ごとの範囲を記録する。
  //やる処理3:上記の処理で取得した範囲を元に、「結果」シートに各種項目を年と月ごとに載せる。
  //やる処理4:「結果」シートにまとめた内容を、ダウンロードフォームにて事前に指定された指示を元にグラフを表示する。(PDFの場合は1ページに収めなければならない。)
  
  //パース
  const recv_params = recv_param_json_obj;
  
  Logger.log(recv_params);
  const filetype_list = ["xlsx","pdf"];
  const download_filetype = recv_params.file_type
  const session_num = recv_params.session_number;
  const totalize_book_fileid = recv_params.totalize_fileid;
  const ob_point_name = recv_params.observation_point_name;
  const total_period_count = recv_params.total_period_count;
  const option_append_pages = recv_params.append_pages;

  let print_sheet = undefined; //PDF指定時のみ
  const result_sheet_graph_page_last_row = 30; //グラフページ最終行
  const result_sheet_first_totalize_page_last_row = 60; //集計表ページ 1ページ目最終行
  const result_sheet_totalize_page_max_row = 29; //ヘッダー行(2行)を含む1ページに収まる最大行数(27行=2年3ヶ月分書き込み可能)

  try{
    //結果シート
    let totalize_book = SpreadsheetApp.openById(totalize_book_fileid);
    const result_sheet = totalize_book.getSheetByName("結果");
    result_sheet.activate;

    //ヘッダー行代入
    result_sheet.getRange(1,10).setValue(ob_point_name); //グラフページ 観測地点名
    result_sheet.getRange(31,10).setValue(ob_point_name); //集計1ページ目　観測地点名

    //結果シート内の集計結果の書式を変更。
    let result_sheet_last_rows = result_sheet.getLastRow();
    const result_sheet_last_columns = result_sheet.getLastColumn();
    Logger.log(result_sheet_last_rows);
    Logger.log(result_sheet_last_columns);

    //書式動的変更
    result_sheet.getRange(result_sheet_graph_page_last_row + 3,2,result_sheet_last_rows,4).setNumberFormat('0.0℃'); //気温
    result_sheet.getRange(result_sheet_graph_page_last_row + 3,5,result_sheet_last_rows,6).setNumberFormat('0.0"mm"'); //降水量
    result_sheet.getRange(result_sheet_graph_page_last_row + 3,7,result_sheet_last_rows,7).setNumberFormat('0.0時間'); //日照時間
    result_sheet.getRange(result_sheet_graph_page_last_row + 3,8,result_sheet_last_rows,10).setNumberFormat('0.0"hPa"'); //気圧

    //集計表に書き込まれたデータをソーティングして整える。
    result_sheet.getRange(result_sheet_graph_page_last_row + 3,1,result_sheet_last_rows,result_sheet_last_columns).sort({column:1,ascending: true});

    //罫線動的描写
    result_sheet_last_rows = result_sheet.getLastRow();
    result_sheet.getRange(result_sheet_graph_page_last_row + 2,1,result_sheet_last_rows - 31,10).setBorder(true,true,true,true,true,true);

    //グラフ生成
    const send_config_obj = {"download_filetype": download_filetype,"total_data_count": total_period_count,"page_title": "全期間","append_verbose_graphs_pages": option_append_pages};
    GenerateTotalizeGraphsPage(result_sheet,send_config_obj);

    //Excel(xlsx)向け処理
    if(download_filetype === "xlsx"){
      //結果シート内の集計結果にフィルタをつけて、昇順に並べ替える
      result_sheet.getRange(result_sheet_graph_page_last_row + 2,1,result_sheet_last_rows -1,result_sheet_last_columns).createFilter();

      //シート順番ソーティング
      const totalize_book_sheet_list = totalize_book.getSheets().map((s) => s.getSheetName());
      //Logger.log(totalize_book_sheet_list);
      const sorted_sheet_list = totalize_book_sheet_list.toSorted();
      //Logger.log(sorted_sheet_list);
      for(s = 0; s < sorted_sheet_list.length;s++){
        totalize_book.getSheetByName(sorted_sheet_list[s]).activate();
        totalize_book.moveActiveSheet(s + 1);
      }
      //結果シートのみ先頭に移動
      result_sheet.activate;
      totalize_book.moveActiveSheet(1);
    }

  }catch(err){
    Logger.log(err.message);
    throw "[GAS_Total]" + err;
  }
}

//アップロード(開始)画面表示処理
function ShowUploadStartPage(){
  //本処理
  let uploader_html = HtmlService.createTemplateFromFile('index');

  SpreadsheetApp.getUi().showModalDialog(uploader_html.evaluate(),"気象データCSV集計");
}

//アップロード処理時確認用(csv_upload.html内JS用)
function CheckUploadCsvFileCount(recv_file_name){
  //アップロード確認用ハンドラ
  try{
    const g_drive_temp_dir = DriveApp.getFolderById(wheather_csv_temp_forder_id);
    const g_drive_temp_dir_files_list = g_drive_temp_dir.getFiles();
    let hit_filename_ids = [];
    let file_count = 0;

    //ファイル検索(空であれば処理をスルー。)
    if(g_drive_temp_dir_files_list.hasNext() == true){
      //ファイルがあれば処理を続行。
      while(g_drive_temp_dir_files_list.hasNext() == true){
        const check_file = g_drive_temp_dir_files_list.next();
        const check_file_id = check_file.getId();
        const check_file_name = check_file.getName();

        if(check_file_name == recv_file_name){
          //同名のファイルの個数を記録する
          hit_filename_ids.push(check_file_id);
        }
      }

      //ヒットした個数を返す。
      if(hit_filename_ids.length > 0){
        file_count = hit_filename_ids.length;
      }

    }

    //なければ0を返し、あればヒットした個数を返す。
    console.log(file_count);
    return file_count;

  }catch(err){
    Logger.log(err);
    throw err;
  }

}

//日付チェック処理
function DateStringValidationCheck(datetime_str){
  const date_regex = /^\d{4}\/\d{1,2}\/\d{1,2} \d{1,2}:\d{2}:\d{2}/; //「YYYY/MM/DD HH:mm:ss」
  const date_regex_result = date_regex.test(datetime_str);
  const check_current_date = new Date(datetime_str);
  let is_valid = false;

  if(date_regex_result == true && !Number.isNaN(check_current_date.getDate())){
    //実在する日付であるかのチェック
    is_valid = true;      
  }else{
    Logger.log("不正な日付の文字列です。" + datetime_str);
  }
  return is_valid;
}

//CSVファイルチェック処理
function JmaCsvValidationCheck(csv_file_name){
  //やらない処理:CSVの行や列を編集(追加、削除)すること(時間がかかりすぎることが予想されるため。)
  //やる処理:気象庁のCSVファイルであるかの確認、異常値の検出(違うCSVファイルや異常値があればエラーを返す。問題がなければ取り込む。)
  //参考元:https://www.data.jma.go.jp/risk/obsdl/top/help3

  //しきい値設定
  //しきい値を超えている値があるCSVファイルは対象外ファイルとして扱う。
  const jma_csv_file_top_str = "ダウンロードした時刻："; //気象庁の先頭行文字列
  const max_temp = 50.0; //最大気温
  const min_temp = -100.0; //最低気温
  const max_rainfall = 100.0; //最大降水量(前1時間あたり)
  const min_rainfall = 0.0; //最大降水量(前1時間あたり)
  const max_sunshine_hour = 400.0; //最大日照時間(前1時間あたり)
  const min_sunshine_hour = 0.0; //最大日照時間(前1時間あたり)
  const max_air_pressure = 1100; //最大現地気圧
  const min_air_pressure = 860; //最低現地気圧

  //クライアントに結果を送信(JSON)
  let result_json = undefined;

  try{
    const check_csv_file = DriveApp.getFilesByName(csv_file_name);
    const check_csv_file_blob = check_csv_file.next().getBlob().getDataAsString('SHIFT_JIS');

    //空チェック
    if(check_csv_file_blob.trim() === ""){
      //「Error:」という文字が冗長になるため、throwに「new Error」をつけない。
      throw "空のCSVファイルを含んでいます。";
    }

    //先頭行の文字列チェック(部分一致)
    const parse_check_csv_data_row = check_csv_file_blob.split(/\r\n|\n/); //分割
    if(parse_check_csv_data_row[0].includes(jma_csv_file_top_str) == false){
      throw "気象庁の過去の気象データではないCSVファイルです。";
    }

    //ヘッダーチェック
    //現時点では「年月日時,気温,降水量,日照時間,現地気圧」のみ対応。
    const jma_header_keywords = ["年月日時","気温","降水量","日照時間","現地気圧"];
    const jma_header = parse_check_csv_data_row[3];
    for(header_check_count = 0; header_check_count < jma_header_keywords.length;header_check_count++){
      if(!jma_header.includes(jma_header_keywords[header_check_count])){
        throw "処理非対応の項目を含んでいます。";
      }
    }
    //値チェック前定義
    let param_is_invalid = false;

    //各種データ注意表記記号の正規表現パターン
    //参照:https://www.data.jma.go.jp/risk/obsdl/top/help4.html#data_hyouki
    const caution_date_regex_pattern = /\s\*/ //「発生日時 *」
    const caution_data_regex_pattern = /\s\)|\s\]/; //

    //処理対象列
    const date_column_num = 0; //日付
    const temp_column_num = 1; //気温
    const rainfall_column_num = 4; //降水量
    const shunshine_hour_column_num = 8; //日照時間
    const air_pressure_column_num = 12; //現地気圧
    const weather_column_num = 15; //天気
    let failure_reason = "";

    //重複行チェック
    const uniq_check = new Set(parse_check_csv_data_row.map(c => JSON.stringify(c)));
    //CSVに元々ある空白行が重複として消されてしまい一致しなくなるため、元の方(length)を-1にして調整。
    if(uniq_check.size !== parse_check_csv_data_row.length - 1){
      throw "重複データを含んでいます。"
    }

    //値チェック(各項目全てがtrueであれば正常。1つでもfalseが出たら異常と判断。ただし、空白の値[欠落データ]はスルーとする。])
    for(row_count = 6;row_count < parse_check_csv_data_row.length - 1 ;row_count++){
      const parse_check_csv_data_columns = parse_check_csv_data_row[row_count].split(",");
      //CSV行を配列に変換
      //日付チェック(YYYY/MM/DD HH:mm:ss)
      let date_cell_param = parse_check_csv_data_columns[date_column_num];
      const check_date_regex_test = caution_date_regex_pattern.test(date_cell_param);
      if(check_date_regex_test){
        date_cell_param = date_cell_param.replace(caution_date_regex_pattern,""); //値についている注意表記記号を取り除く
      }
      const check_date_param = DateStringValidationCheck(date_cell_param);

      //気温チェック
      let check_temp_param = undefined;
      let temp_param = parse_check_csv_data_columns[temp_column_num];
      const check_temp_regex_test = caution_data_regex_pattern.test(temp_param);
      
      if(check_temp_regex_test && (temp_param !== "" && temp_param !== "--" && temp_param !== "///" && temp_param !== "#")){
        temp_param = temp_param.replace(caution_data_regex_pattern,""); //値についている注意表記記号を一時的に取り除く
      }

      //注意表記記号または空文字でないかをチェック。(事前に分割しないと、文字列と数字を比較[文字コード]することになりバグの元になるため。)
      if(temp_param !== "" && temp_param !== "--" && temp_param !== "///" && temp_param !== "#"){
        check_temp_param = (temp_param <= max_temp) && (temp_param >= min_temp);
      }else{
        check_temp_param = true; //注意表記記号のみの場合はスルーする。(SSのaverage関数等は文字列を無視するため。)
      }

      //降水量チェック
      let check_rainfall_param = undefined;
      let rainfall_param = parse_check_csv_data_columns[rainfall_column_num];
      const check_rainfall_regex_test = caution_data_regex_pattern.test(rainfall_param);
      
      if(check_rainfall_regex_test && (rainfall_param !== "" && rainfall_param !== "--" && rainfall_param !== "///" && rainfall_param !== "#")){
        rainfall_param = rainfall_param.replace(caution_data_regex_pattern,"");
      }
      
      if(rainfall_param !== "" && rainfall_param !== "--" && rainfall_param !== "///" && rainfall_param !== "#"){
        check_rainfall_param = (rainfall_param  <= max_rainfall) && (rainfall_param >= min_rainfall);
      }else{
        check_rainfall_param = true;
      }

      //日照時間チェック
      let check_sunshine_hour_param = undefined;
      let sunshine_hour_param = parse_check_csv_data_columns[shunshine_hour_column_num];
      const check_sunshine_hour_regex_test = caution_data_regex_pattern.test(sunshine_hour_param);
      
      if(check_sunshine_hour_regex_test && (sunshine_hour_param !== "" || sunshine_hour_param !== "--" || sunshine_hour_param !== "///" || sunshine_hour_param !== "#")){
        sunshine_hour_param = sunshine_hour_param.replace(caution_data_regex_pattern,"");
      }

      if(sunshine_hour_param !== "" && sunshine_hour_param !== "--" && sunshine_hour_param !== "///" && sunshine_hour_param !== "#"){
        check_sunshine_hour_param = (sunshine_hour_param <= max_sunshine_hour) && (sunshine_hour_param >= min_sunshine_hour);
      }else{
        check_sunshine_hour_param = true;
      }

      //気圧チェック
      let check_air_pressure_param = undefined;
      let air_pressure_param = parse_check_csv_data_columns[air_pressure_column_num];

      const check_air_pressure_regex_test = caution_data_regex_pattern.test(air_pressure_param);

      if(check_air_pressure_regex_test && (air_pressure_param !== "" && air_pressure_param !== "--" && air_pressure_param !== "///" || air_pressure_param !== "#")){
        air_pressure_param = air_pressure_param.replace(caution_data_regex_pattern,"");
      }

      if(air_pressure_param !== "" && air_pressure_param !== "--" && air_pressure_param !== "///" && air_pressure_param !== "#"){
        check_air_pressure_param = (air_pressure_param <= max_air_pressure) && (air_pressure_param >= min_air_pressure);
      }else{
        check_air_pressure_param = true;
      }

      //チェック結果
      if(check_date_param == false || check_temp_param == false || check_rainfall_param == false || check_sunshine_hour_param == false || check_air_pressure_param == false){
        const print_rowcount = row_count + 1;
        if(check_date_param == false){
          failure_reason = "日付の値が不正です。\n" + "[位置]:"+ print_rowcount +"行目/" + "[値]:"+ parse_check_csv_data_columns[date_column_num];
        }

        if(check_temp_param == false){
          failure_reason = "気温の値が不正です。\n" + "[位置]:"+ print_rowcount +"行目/" + "[値]:"+ parse_check_csv_data_columns[temp_column_num];
        }

        if(check_rainfall_param == false){
          failure_reason = "降水量の値が不正です。\n" + "[位置]:"+ print_rowcount +"行目/" + "[値]:"+ parse_check_csv_data_columns[rainfall_column_num];
        }

        if(check_sunshine_hour_param == false){
          failure_reason = "日照時間の値が不正です。\n" + "[位置]:"+ print_rowcount +"行目/" + "[値]:"+ parse_check_csv_data_columns[shunshine_hour_column_num];
        }

        if(check_air_pressure_param == false){
          failure_reason = "現地気圧の値が不正です。\n" + "[位置]:"+ print_rowcount +"行目/" + "[値]:"+ parse_check_csv_data_columns[air_pressure_column_num];
        }

        Logger.log(failure_reason);
        param_is_invalid = true;
        //不正な値が混じっているCSVファイルは使用しないため、強制的にループを抜けた後にエラーを出す。
        break;
      }
    }

    //処理結果分岐
    if(param_is_invalid == true){
      throw "不正な値を含んでいるCSVファイルです。" + failure_reason;
    }else{
      Logger.log("done.");
    }

    //結果を生成
    result_json = {"filename": csv_file_name,"msg": "done"};

  }catch(err){
    Logger.log("対象ファイル: " + csv_file_name);
    Logger.log(err);
    result_json = {"filename": csv_file_name,"msg": String(err)}; //errをString関数で文字列化しないとJS側では[Object]になって返してしまう。

  }finally{
    //結果をJSONとして返す。
    return JSON.stringify(result_json);
  }

}

//受信(アップロード)処理
function UploadCsvFiles(recv_b64_data,recv_file_name){
  //アップロード用ハンドラ
  //アップロード画面(HTML)を表示させる。->ファイルを選択後にアップロードボタンを押すことで、アップロードが開始される。
  //->GASスクリプトがあるドライブに保存する。->アップロードされて閉じた際にアップロードされたファイルのパスを戻す。

  //参考1:https://tonari-it.com/gas-dialog-file-blob/
  //参考2:https://gist.github.com/tanaikech/3e44c779f05374d19333444c9a4dd5ba
  //参考3:https://webird-programming.tech/archives/1431

  try{
    let recv_decode_data = Utilities.base64Decode(recv_b64_data);
    let upload_blob_obj = Utilities.newBlob(recv_decode_data,'text/csv',recv_file_name);
  
    const g_drive_save_temp_dir = DriveApp.getFolderById(wheather_csv_temp_forder_id);
    const g_drive_save_temp_dir_files_list = g_drive_save_temp_dir.getFiles();
    let duplicate_name_file_ids = [];

    //保険としてファイル重複チェック(空であれば処理をスルー。)
    let apply_file_name = "";

    if(g_drive_save_temp_dir_files_list.hasNext() == true){
      //ファイルがあれば処理を続行。
      while(g_drive_save_temp_dir_files_list.hasNext() == true){
        const check_file = g_drive_save_temp_dir_files_list.next();
        const check_file_id = check_file.getId();
        const check_file_name = check_file.getName();

        if(check_file_name == recv_file_name){
          //重複があれば、重複ファイルIDを記録する。
          duplicate_name_file_ids.push(check_file_id);
        }
      }

      //重複があれば変更する。
      if(duplicate_name_file_ids.length > 0){
        apply_file_name = recv_file_name + "_" + duplicate_name_file_ids.length;
      }else{
        apply_file_name = recv_file_name;
      }

    }else{
      apply_file_name = recv_file_name;
    }

    //ファイルをドライブ内に保存
    const blob_to_file = g_drive_save_temp_dir.createFile(upload_blob_obj);
    blob_to_file.setName(apply_file_name);

    return true;
    
  }catch(err){
    Logger.log(err);
    return false;
  }
}

//送信(ダウンロード)処理
function SendCompletedFile(download_fileid,download_filetype){
  //ダウンロード用ハンドラ
  //ダウンロード画面(HTML)を表示させる。->保存先を選択後、ファイル種類(xlsx、PDF等)を指定する。
  //->一時保存された処理済みのスプレッドシートをコピーした上で、指定された種類と保存先を元に指定通りに処理してダウンロード用ファイルを生成。
  //->指定されたファイル種類で保存し、ダウンロードを開始する。
  //->一時保存された処理済みのスプレッドシートを削除する。
  //->処理結果をBase64の文字列で戻す。(失敗した場合はエラーをクライアント側にエスカレーションする。)

  //参考元1:https://qiita.com/chovin/items/4647e4b62df9579996f0
  //参考元2:https://qiita.com/rc_code/items/150003f016287750cf34

  let process_is_failed = false;
  let download_blob = null;
  let whather_analysis_ss_filename = "";
  let temp_file_id = "";
  let temp_filename = "";
  let error_msg = "";
  const filetype_list = ["xlsx","pdf"];

  try{
    //PDF形式のページ設定(gidは後で代入。)
    /*印刷設定一覧
      "size": "A4",
      "portrait": false 横向き
      "fitw": true ページ幅に合わせる?いる？
      "scale":2,
      "gridlines": false 表示しない
      "top_margin":	1.905 上の余白(プリセット「狭い」相当)
      "bottom_margin": 1.905 下の余白(プリセット「狭い」相当)
      "left_margin": 0.635 左の余白(プリセット「狭い」相当)
      "right_margin":	0.635 右の余白(プリセット「狭い」相当)
      "pagenumbers": true

    */
    let pdf_page_config = {
      "exportFormat": "pdf",
      "format": "pdf",
      "gid": null,
      "size": "A4",
      "portrait": false,
      "fitw": true,
      "scale": 2,
      "gridlines": false,
      "pagenumbers": true
    }

    //アドレス生成
    const download_base_url = "https://docs.google.com/spreadsheets/d/"
    let get_blob_url = "";

    //引数チェック
    if(download_filetype.length === 0){
      //空チェック
      throw Error("ファイルタイプが指定されていません。");
    }else if(!filetype_list.includes(download_filetype)){
      //ファイルタイプチェック
      throw Error("不正なファイルタイプです。");
    }

    //フォルダ指定
    const process_completed_folder = DriveApp.getFolderById(wheather_totalize_files_folder_id);

    //ダウンロード対象のスプレッドシートを展開
    let base_book = SpreadsheetApp.openById(download_fileid);

    //ダウンロード対象のスプレッドシートファイルのBlobオブジェクト
    const wheather_analysis_ss_file = DriveApp.getFileById(download_fileid);
    whather_analysis_ss_filename = wheather_analysis_ss_file.getName();
    base_file_id = wheather_analysis_ss_file.getId();

    //ファイル名:download-処理時刻-気象データCSV分析
    const temp_ss_filename = "download-" + download_filetype + "-" + wheather_analysis_ss_file.getName().substring(5);

    //処理前下準備
    //ダウンロードするファイルタイプ別に処理。
    if(download_filetype == "pdf"){
      //PDFファイルの場合(印刷用結果ページのみを残す。)
      base_book.getSheetByName("結果").activate();
      const base_book_sheet_id = base_book.getSheetId();

      //エクスポート対象のIDを代入。
      pdf_page_config.gid = base_book_sheet_id;
      
      //クエリ部分生成
      const query = Object.keys(pdf_page_config).map((key) => {
        return encodeURIComponent(key) + "=" + encodeURIComponent(pdf_page_config[key]);
      }).join('&');

      //スプレッドシートからPDF形式に変換するためのURLを生成。
      get_blob_url = download_base_url + base_file_id + "/export?" + query;
      

    }else{
      //Excelファイルの場合は、本ファイルをコピーした上でメニュー、varシートを消す。

      //ファイルコピー
      const copy_ss_file = wheather_analysis_ss_file.makeCopy(temp_ss_filename,process_completed_folder);
      temp_file_id = copy_ss_file.getId();
      temp_filename = copy_ss_file.getName();
      const copy_book = SpreadsheetApp.openById(copy_ss_file.getId());

      //シート削除
      const delete_sheet_lists = copy_book.getSheets();
      const delete_sheet_name_list = ["メニュー"];

      for(let del_count = 0; del_count < delete_sheet_lists.length; del_count++){
        if(!delete_sheet_name_list.includes(delete_sheet_lists[del_count].getName())){
          //削除対象のシートでなければスルーして次の処理に進む。
          continue;
        }else{
          //対象シートを削除
          copy_book.deleteSheet(delete_sheet_lists[del_count]);
        }
      }

      //スプレッドシートからExcel形式に変換するためのURLを生成。
      get_blob_url = download_base_url + temp_file_id + "/export?format=xlsx";
    }

    //逆アップロードしてBlobオブジェクトを生成する。
    let token = ScriptApp.getOAuthToken();
    let response = UrlFetchApp.fetch(get_blob_url,{
      headers: {'Authorization': 'Bearer ' + token,'MuteHttpExceptions': true}
    });
    download_blob = response.getBlob().getBytes();

    //blobオブジェクトをbase64に変換
    send_b64_data = Utilities.base64Encode(download_blob);

  }catch(err){
    Logger.log(err);
    process_is_failed = true;
    error_msg = err.message

  }finally{
    //本処理で作成した生成ファイル全てを削除
    if(temp_file_id.length > 0){
      DeleteUploadFile(temp_filename);
      DeleteUploadFile(whather_analysis_ss_filename);
    }

    //失敗した場合は、エスカレーションする。
    if(process_is_failed == true){
      throw "[GAS_download]:" + error_msg;

    }else{
      //base64データを返す。
      return send_b64_data;
    }
  }

}

//アップロードファイル削除
function DeleteUploadFile(delete_filename){
  try{
    //対象ファイルの存在チェック
    const check_file_exist_result = CheckUploadCsvFileCount(delete_filename);

    if(check_file_exist_result == 0){
      throw new Error("削除対象のファイルが見つかりませんでした。");
    }else if(typeof(check_file_exist_result) == "null"){
      throw new Error("例外が発生しました。");
    }

    //削除処理    
    const delete_file = DriveApp.getFilesByName(delete_filename);
    if(!delete_file.hasNext()){
      throw new Error("指定された削除対象のファイルは存在しません。");

    }
    delete_file.next().setTrashed(true); //実際はゴミ箱に移動されるだけ。
    return true;
    
  }catch(err){
    Logger.log(err);
    return false;
  }
}
