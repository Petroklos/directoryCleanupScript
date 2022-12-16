' Declare the variables
Dim directoryPaths, FSO, logFile, strUser, Deleted

' Set the paths to be checked and the number of days to retain folders
drectoryPaths = Array(".\",".\foobar")
daysToRetain = 45

' Set the paths to be checked and the number of days to retain folders
Set FSO = CreateObject("Scripting.FileSystemObject")
Set logFile = FSO.OpenTextFile(".\directoryDeletion.log", 8, true, 0)

' Get the current user's username
strUser = CreateObject("WScript.Network").UserName

' Write the user's username and the current date and time to the log file
logFile.WriteLine strUser & "	" & now

' Define the DirectoryCleanByDate function, which takes a directory path as an argument and deletes any subfolders in that directory that are older than the specified number of days
Function DirectoryCleanByDate (Directory)
	For Each Folder in FSO.GetFolder(Directory).SubFolders
		If now - Folder.DateCreated > daysToRetain Then
			logFile.WriteLine "Deleted Folder " & Chr(34) & Folder.Path & Chr(34)
			FSO.DeleteFolder(Folder.Path)
			Deleted = Deleted + 1
		End If
	Next
End Function

' Define the DirectoryCleanByName function, which takes a directory path as an argument and deletes any subfolders in that directory whose names meet a certain criteria
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

' Define the DirectoryDFS function, which takes a directory path as an argument, recursively searches for subfolders named "toClean" and calls the DirectoryCleanByDate function on them
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

' Loop through the list of directory paths
For i = 0 To UBound(directoryPaths)
	' Check if the directory exists
	If FSO.FolderExists(directoryPaths(i)) Then
		 ' Set the counter for deleted directories to 0
		Deleted = 0
		' Call the DirectoryDFS function on the current directory
		DirectoryDFS directoryPaths(i)
		' Write the number of deleted directories to the log file
		logFile.WriteLine "Deleted " & Deleted & " Directories in " & Chr(34) & directoryPaths(i) & Chr(34)
	Else
		' Write a message to the log file if the directory does not exist
		logFile.WriteLine "Directory " & Chr(34) & directoryPaths(i) & Chr(34) & " can't be found"
	End If
Next

' Add a blank line to the log file and close it
logFile.WriteLine
logFile.Close
