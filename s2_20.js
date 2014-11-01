//MyProc_DownLoadBeginの関数が呼ばれない、ブラウザー自体は動く
var ie = WScript.CreateObject("InternetExplorer.Application","MyProc_");
ie.Visible = 1;
//指定したURLにジャンプ
ie.Navigate("http://www.oreilly.co.jp/index.shtml");

function MyProc_DownLoadBegin()
{
	WScript.Echo("aaa");
}
