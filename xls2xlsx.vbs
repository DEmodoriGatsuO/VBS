' �萔
Const XlFileFormat_xlOpenXMLWorkbook = 51 ' .xlsx : Excel �u�b�N

' �h���b�O�h���b�v�œn���ꂽ�t�@�C���𕶎���Ŏ�荞�ށB���b�Z�[�W�p�B
set args = WScript.Arguments
fileList = ""
for each arg in args
  fileList = fileList & vbNewLine & arg
next

' �����̃`�F�b�N�B�Ώۃt�@�C���ȊO���������Ă���ꍇ�I���B
set fobj = CreateObject("Scripting.FileSystemObject")
for each arg in args
    ext = fobj.GetextensionName(arg)
    if ext <> "xls" then
        msgbox "xls�t�@�C���ȊO���w�肳��܂����B�I�����܂��B" & vbNewLine & fileList
        WScript.Quit
    end if
next

' �����Ŗ�����eExcel�t�@�C�����ŐV�̌`���ŕۑ�����B
set oXlsApp = CreateObject("Excel.Application")
for each path in args
    oXlsApp.Application.Visible = true
    set book = oXlsApp.Application.Workbooks.Open(path)
    book.SaveAs Replace(path, ".xls", ".xlsx"), XlFileFormat_xlOpenXMLWorkbook
    book.Close
next
oXlsApp.Quit
set oXlsApp = nothing

msgbox "�ϊ��������܂����B"