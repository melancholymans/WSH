var shell,intRet;
shell = WScript.CreateObject("WScript.Shell");
intRet = shell.Popup("�����V�[�P���X���J�n���܂���",0,"�����J�n",1);
//WScript.Echo(intRet);
if(intRet == 1){
	WScript.Echo("���[�v�R�A�����܂łR�O�b");
}else{
	WScript.Echo("�����V�[�P���X�𒆎~���܂���");
}
