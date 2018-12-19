dim grade

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "D:\HdkjWok"

Set objFolder = objFSO.GetFolder(objStartFolder)
grade = 0
Wscript.Echo "--" & grade & "--  "  & objFolder.Path

Set colFiles = objFolder.Files

For Each objFile in colFiles
    Wscript.Echo objFile.Name
Next
Wscript.Echo

ShowSubfolders objFSO.GetFolder(objStartFolder), grade+1

Sub ShowSubFolders(Folder, grade)
    For Each Subfolder in Folder.SubFolders
        echoGrade grade
        Wscript.Echo Subfolder.Name
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            Wscript.Echo objFile.Name
        Next
        Wscript.Echo
        ShowSubFolders Subfolder, grade+1
    Next
End Sub

Sub echoGrade(grade)
Wscript.Stdout.Write "FG:" & grade
For i = 1 to grade
    Wscript.Stdout.Write "--"
next
End Sub