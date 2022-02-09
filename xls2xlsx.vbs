' 定数
Const XlFileFormat_xlOpenXMLWorkbook = 51 ' .xlsx : Excel ブック

' ドラッグドロップで渡されたファイルを文字列で取り込む。メッセージ用。
set args = WScript.Arguments
fileList = ""
for each arg in args
  fileList = fileList & vbNewLine & arg
next

' 引数のチェック。対象ファイル以外が混ざっている場合終了。
set fobj = CreateObject("Scripting.FileSystemObject")
for each arg in args
    ext = fobj.GetextensionName(arg)
    if ext <> "xls" then
        msgbox "xlsファイル以外が指定されました。終了します。" & vbNewLine & fileList
        WScript.Quit
    end if
next

' 引数で貰った各Excelファイルを最新の形式で保存する。
set oXlsApp = CreateObject("Excel.Application")
for each path in args
    oXlsApp.Application.Visible = true
    set book = oXlsApp.Application.Workbooks.Open(path)
    book.SaveAs Replace(path, ".xls", ".xlsx"), XlFileFormat_xlOpenXMLWorkbook
    book.Close
next
oXlsApp.Quit
set oXlsApp = nothing

msgbox "変換完了しました。"