''return array of files in a input folder
''@path     :   path of folder
''@search   :   if "" then return all files in folder, or return file contain searched string
''@return   :   array contain files ,variant data type
Sub getSpeicalFolderFiles(ByRef filesArr() As Variant, path As String, search As String)

Set fso = CreateObject("Scripting.FileSystemObject")

For Each f In fso.GetFolder(path).Files
    If search = "" Then
       pushArray f, filesArr
    ElseIf InStr(f.Name, search) > 0 Then
       pushArray f, filesArr
    End If
Next

End Sub