On Error Resume Next

Dim cFSO, cFile, sThisFolder, sTargetFolder

Function timeStampForFolder (cTime)
    timeStampForFolder = Year(cTime) & "." & _
		Right ("0" & Month(cTime),2) & "." & _
		Right ("0" & Day(cTime),2)
End Function


Function CreateDateFolder ()
	Dim sTargetName
	
	If Not cFile Is Nothing Then
		sTargetName = cFile.ParentFolder.Path & "\" & timeStampForFolder (cFile.DateLastModified)

		If sTargetName <> "" Then
			If NOT cFSO.FileExists (sTargetName) AND _
			   NOT cFSO.FolderExists (sTargetName) Then
				cFSO.CreateFolder (sTargetName)
			End If
			If cFSO.FolderExists (sTargetName) Then
				CreateDateFolder = sTargetName
			End If
		End If
	End If
End Function

Set cFSO= Wscript.CreateObject("Scripting.FileSystemObject")   

For Each Arg In WScript.Arguments
	If cFSO.FileExists (Arg) Then
		Set cFile = cFSO.GetFile (Arg)
		sTargetFolder = CreateDateFolder ()
		sThisFolder = cFile.Path
		If sTargetFolder <> "" Then
			If NOT cFSO.FileExists (sTargetFolder & "\" & cFile.Name) AND _
			   NOT cFSO.FolderExists (sTargetFolder & "\" & cFile.Name) Then
				Set cFile = Nothing
				cFSO.MoveFile sThisFolder, sTargetFolder & "\"
			End If
		End If
	End If
Next

Set cFSO = Nothing
