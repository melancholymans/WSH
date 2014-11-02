var shell;
shell = WScript.CreateObject("WScript.Shell");
//ノートパッドを起動
shell.Run("notepad.exe",1,0);
//タイムラグを入れる
WScript.Sleep(500);
//START:というkey入力をノートパッドに送信
shell.Sendkeys("START:");
//%はAlt keyなので日付けと時刻を送信する
shell.Sendkeys("%ED");
