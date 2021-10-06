directoryPaths = Array(".\")
daysToRetain = 45

Set FSO = CreateObject("Scripting.FileSystemObject")
Set logFile = FSO.OpenTextFile(".\directoryDeletion.log", 8, true, 0)
strUser = CreateObject("WScript.Network").UserName
logFile.WriteLine strUser & "	" & now

Function DirectoryClean (Directory)
	For Each Folder in FSO.GetFolder(Directory).SubFolders
		If now - Folder.DateCreated > daysToRetain Then
			FSO.DeleteFolder(Folder.Path)
			Deleted = Deleted + 1
			logFile.WriteLine "Deleted Folder " & Target.Path
		End If
	Next
End Function

Function DirectoryDFS (Directory)
	For Each Folder in FSO.GetFolder(Directory).SubFolders
		If Folder.Name = "toClean" Then
			DirectoryClean Folder.Path
		Else
			DirectoryDFS Folder.Path
		End If
	Next
End Function

For i = 0 To UBound(directoryPaths)
	Deleted = 0
	DirectoryDFS directoryPaths(i)
	logFile.WriteLine "Deleted " & Deleted & " Directories in " & directoryPaths(i)s
Next

logFile.WriteLine
logFile.Close