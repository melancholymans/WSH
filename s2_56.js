var shell;
shell = WScript.CreateObject("WScript.Shell");
//�m�[�g�p�b�h���N��
shell.Run("notepad.exe",1,0);
//�^�C�����O������
WScript.Sleep(500);
//START:�Ƃ���key���͂��m�[�g�p�b�h�ɑ��M
shell.Sendkeys("START:");
//%��Alt key�Ȃ̂œ��t���Ǝ����𑗐M����
shell.Sendkeys("%ED");
