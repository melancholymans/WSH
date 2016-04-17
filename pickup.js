//-------------------------------------------------------------------
// ����Excel�t�@�C������C�ӂ̍s���s�b�N�A�b�v
// �g����: ���o�������t�@�C�����h���O���h���b�v���܂�
//-------------------------------------------------------------------
var keyword = "�����s�O��s";
var result_file = getAppPath() + "result.csv";
//-------------------------------------------------------------------
var result = "";
var match_count = 0;
var fso = WScript.CreateObject("Scripting.FileSystemObject");
// �h���O&�h���b�v���ꂽ�t�@�C�����m�F����
var file_count = WScript.Arguments.Count();
if (file_count == 0) {
  WScript.echo("�t�@�C��������܂���B");
  WScript.Quit();
}
// Excel���N������
var excel = WScript.CreateObject("Excel.Application");
excel.Visible = true;
// �h���O&�h���b�v���ꂽ�t�@�C����1����������
for (var i = 0; i < file_count; i++) {
  pickupData(WScript.Arguments.Item(i));
}
excel.Quit();
// ���o���ʂ�ۑ�����
var csv = fso.CreateTextFile(result_file, true);
csv.Write(result);
csv.Close();
WScript.echo(match_count+"�������܂����B");
//-------------------------------------------------------------------
// �f�[�^�̒��o����
function pickupData(fname) {
  var book = excel.Workbooks.Open(fname); // �u�b�N���J��
  // �e�V�[�g�𒲂ׂĂ���
  for (var bi = 1; bi <= book.Worksheets.Count; bi++) {
    var sheet = book.Worksheets(bi);
    // �ŏI�s�𒲂ׂ�
    var used = sheet.UsedRange;
    if (used.Count <= 1) continue; // �g�p���Z����1�ȉ��Ȃ珈�����Ȃ�
    var lastrow = used.Cells(used.Count).Row;
    var lastcol = used.Cells(used.Count).Column;
    // �e�s�𒲂ׂĂ���
    for (var row = 1; row <= lastrow; row++) {
      var flag_find = false;
      var r = [];
      for (var col = 1; col <= lastcol; col++) {
        var cell = sheet.Cells(row, col).Value;
        if (!cell) continue; // ��Ȃ�΃`�F�b�N���Ȃ�
        if (checkKeyword(cell)) flag_find = true; // �`�F�b�N
        cell = cell.replace(/\"/g, "�h"); // �G�X�P�[�v
        r.push('"' + cell + '"'); // CSV�Ƃ��ċL�^����
      }
      if (flag_find) {
        result += r.join(",") + "\n";
        match_count++;
      }
    }
  }
  book.Close();
}
// �L�[���[�h�Ƀ}�b�`���邩�ǂ������ׂ�
function checkKeyword(cell) {
  //indexOf�̃��\�b�h��javaScript�̂��̂̂悤�ł���H
  return (cell.indexOf(keyword) >= 0);
}
//-------------------------------------------------------------------
// �A�v���P�[�V�����̃p�X�𒲂ׂ�
function getAppPath() {
  return WScript.ScriptFullName.substr(0, 
    WScript.ScriptFullName.length - WScript.ScriptName.length);
}
