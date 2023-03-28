Dim directoryPaths, FSO, logFile, strUser, Deleted

drectoryPaths = Array(".\",".\foobar")
daysToRetain = 45

Set FSO = CreateObject("Scripting.FileSystemObject")
Set logFile = FSO.OpenTextFile(".\directoryDeletion.log", 8, true, 0)

strUser = CreateObject("WScript.Network").UserName

logFile.WriteLine strUser & "	" & now

Function DirectoryCleanByDate (Directory)
	For Each Folder in FSO.GetFolder(Directory).SubFolders
		If now - Folder.DateCreated > daysToRetain Then
			logFile.WriteLine "Deleted Folder " & Chr(34) & Folder.Path & Chr(34)
			FSO.DeleteFolder(Folder.Path)
			Deleted = Deleted + 1
		End If
	Next
End Function

Function DirectoryCleanByName (Directory)
	For Each Folder in FSO.GetFolder(Directory).SubFolders
		tokens = Split(Folder.Name,"-")
		compareSum = (year(now)-tokens(0))*365 + (month(now)-tokens(1))*30 + (day(now)-tokens(2))
		If compareSum > daysToRetain Then
			logFile.WriteLine " Deleted Folder " & Chr(34) & Folder.Path & Chr(34)
			FSO.DeleteFolder(Folder.Path)
			Deleted = Deleted + 1
		End If
	Next
End Function

Function DirectoryDFS (Directory)
	For Each Folder in FSO.GetFolder(Directory).SubFolders
		If Folder.Name = "toClean" Then
			' Use DirectoryCleanByName if you want to delete based on Name'
			' DirectoryCleanByName Folder.Path '
			' Use DirectoryCleanByName if you want to delete based on Date'
			DirectoryCleanByDate Folder.Path
		Else
			DirectoryDFS Folder.Path
		End If
	Next
End Function

For i = 0 To UBound(directoryPaths)
	If FSO.FolderExists(directoryPaths(i)) Then
		Deleted = 0
		DirectoryDFS directoryPaths(i)
		logFile.WriteLine "Deleted " & Deleted & " Directories in " & Chr(34) & directoryPaths(i) & Chr(34)
	Else
		logFile.WriteLine "Directory " & Chr(34) & directoryPaths(i) & Chr(34) & " can't be found"
	End If
Next

logFile.WriteLine
logFile.Close
