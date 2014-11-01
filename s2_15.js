var today;
var d = new Date();
//このスクリプトはコマンドプロンプトで実行すること
//>cscript s2_15.js
today = d.getYear() + "/" + (d.getMonth() + 1) + "/" + d.getDate();
WScript.StdOut.WriteLine(today);

