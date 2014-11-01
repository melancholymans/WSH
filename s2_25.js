//実行できなかった、いろいろ調べたがわからない
var obj = WScript.GetObject("C:\myproject\learningWSH\sample.xls");
obj.Application.Visible = true;
obj.Parent.Windows("sample.xlsx").Visible = true;

