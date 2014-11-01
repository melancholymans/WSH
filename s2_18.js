//エクセルのobjectをつくる
var excel = WScript.CreateObject("Excel.Application");
//新しいブックを追加する
excel.WorkBooks.Add();
excel.Cells(1,1).Formula = "=sum(23,37)";
WScript.Echo(excel.Cells(1,1).Value);
excel.DisplayAlerts = false;
//Excellを終了させる
excel.Quit()
