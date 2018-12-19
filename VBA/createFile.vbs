'根据输入的文件路径，产生一个指定的文件，如果指定的路径下存在一个同名的文件，就先删除该文件
'@ FilePath 文件绝对路径
Sub createFile(FilePath As String)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(FilePath) Then
    fso.DeleteFile (FilePath)
End If

Dim oFile As Object
Set oFile = fso.CreateTextFile(FilePath)

oFile.WriteLine "--'** author:kursk.ye@gmail.com **"
oFile.Close

Set fso = Nothing
Set oFile = Nothing

End Sub