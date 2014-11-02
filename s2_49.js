var shell,intRet;
shell = WScript.CreateObject("WScript.Shell");
intRet = shell.Popup("自爆シーケンスを開始しますか",0,"自爆開始",1);
//WScript.Echo(intRet);
if(intRet == 1){
	WScript.Echo("ワープコア自爆まで３０秒");
}else{
	WScript.Echo("自爆シーケンスを中止しました");
}
