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
	CreateDateTimeString = Replace(Replace(Replace(dt, "/", "-"), ":", "_"), " ", "_")
End Function 