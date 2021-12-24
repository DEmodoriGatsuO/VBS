Option Explicit

Call Main

Sub Main()

	Dim fso: Set fso = CreateObject ("Scripting.FileSystemObject") 
	Dim arg
	For Each arg In WScript.Arguments 
		Dim f: Set f = fso.GetFile(arg) 
		Dim newName: newName = fso.GetBaseName (f.Name) & "_" & CreateDateTimeString(Now) & "." & fso.GetExtensionName(f.Name)
		f.Name = newName
	Next 

End Sub

Function CreateDateTimeString(dt) 
'引数dtはDate型を取る
'dt をyyyy-MM-dd_HH-mm-ss にフォーマッティングした文字列を返す。
'※VBSではFormat関数が使えない
	Dim yyyy: yyyy = Year(now)
	Dim MM: MM = Month(now)
	Dim dd: dd = Day(now)
	if len(MM) = 1 then MM = "0" & MM
	if len(dd) = 1 then dd = "0" & dd
	CreateDateTimeString = yyyy & MM & dd
End Function 