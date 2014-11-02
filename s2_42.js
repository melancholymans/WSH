var shell,env;
shell = WScript.CreateObject("WScript.Shell");
env = shell.Environment("Process");
WScript.Echo(env.Item("OS"));
/*
環境変数

NUMBER_OF_PROCESSORS	プロセッサの数
PROCESSOR_ARCHITECTURE	プロセッサの種類
PROCESSOR_IDENTIFIER		プロセッサの識別コード
PROCESSOR_LEVEL			プロセッサのレベル
PROCESSOR_REVISION		プロセッサのバージョン
OS							オペレーションシステム
COMSPEC					コマンドプロセッサの名前
HOMEDRIVE					プライマリードライブ
HOMEPATH					ユーザー用の既定フォルダ
PATH						ＰＡＴＨの内容
PATHEXT					実行可能ファイルの拡張子
PROMPT						プロンプトの文字列
SYSTEMDRIVE				システムフォルダが存在するローカルドライブ
SYSTEMROOT				システムフォルダ
WINDIR						システムフォルダ
TEMP						一時ファイル用のフォルダ
TMP							一時ファイル用のフォルダ
*/
