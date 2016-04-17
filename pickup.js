//-------------------------------------------------------------------
// 複数Excelファイルから任意の行をピックアップ
// 使い方: 抽出したいファイルをドラグ＆ドロップします
//-------------------------------------------------------------------
var keyword = "東京都三鷹市";
var result_file = getAppPath() + "result.csv";
//-------------------------------------------------------------------
var result = "";
var match_count = 0;
var fso = WScript.CreateObject("Scripting.FileSystemObject");
// ドラグ&ドロップされたファイルを確認する
var file_count = WScript.Arguments.Count();
if (file_count == 0) {
  WScript.echo("ファイルがありません。");
  WScript.Quit();
}
// Excelを起動する
var excel = WScript.CreateObject("Excel.Application");
excel.Visible = true;
// ドラグ&ドロップされたファイルを1つずつ処理する
for (var i = 0; i < file_count; i++) {
  pickupData(WScript.Arguments.Item(i));
}
excel.Quit();
// 抽出結果を保存する
var csv = fso.CreateTextFile(result_file, true);
csv.Write(result);
csv.Close();
WScript.echo(match_count+"件見つけました。");
//-------------------------------------------------------------------
// データの抽出処理
function pickupData(fname) {
  var book = excel.Workbooks.Open(fname); // ブックを開く
  // 各シートを調べていく
  for (var bi = 1; bi <= book.Worksheets.Count; bi++) {
    var sheet = book.Worksheets(bi);
    // 最終行を調べる
    var used = sheet.UsedRange;
    if (used.Count <= 1) continue; // 使用中セルが1以下なら処理しない
    var lastrow = used.Cells(used.Count).Row;
    var lastcol = used.Cells(used.Count).Column;
    // 各行を調べていく
    for (var row = 1; row <= lastrow; row++) {
      var flag_find = false;
      var r = [];
      for (var col = 1; col <= lastcol; col++) {
        var cell = sheet.Cells(row, col).Value;
        if (!cell) continue; // 空ならばチェックしない
        if (checkKeyword(cell)) flag_find = true; // チェック
        cell = cell.replace(/\"/g, "”"); // エスケープ
        r.push('"' + cell + '"'); // CSVとして記録する
      }
      if (flag_find) {
        result += r.join(",") + "\n";
        match_count++;
      }
    }
  }
  book.Close();
}
// キーワードにマッチするかどうか調べる
function checkKeyword(cell) {
  //indexOfのメソッドはjavaScriptのもののようである？
  return (cell.indexOf(keyword) >= 0);
}
//-------------------------------------------------------------------
// アプリケーションのパスを調べる
function getAppPath() {
  return WScript.ScriptFullName.substr(0, 
    WScript.ScriptFullName.length - WScript.ScriptName.length);
}
