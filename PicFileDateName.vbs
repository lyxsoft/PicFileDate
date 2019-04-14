On Error Resume Next

Dim cFSO, cFile, cThisFolder

Function timeStamp (cTime)
    timeStamp = Year(cTime) & _
		Right ("0" & Month(cTime),2) & _
		Right ("0" & Day(cTime),2)  & "_" & _
		Right ("0" & Hour(cTime),2) & _
		Right ("0" & Minute(cTime),2) & _
		Right ("0" & Second(cTime),2) 
End Function

Function SetFileDateName ()
	Dim sTargetName
	Dim nIndex
	
	If Not cFile Is Nothing Then
		SetFileDateName = timeStamp (cFile.DateLastModified)

		If SetFileDateName <> "" Then
			Select Case UCase (cFSO.GetExtensionName (cFile.Name))
			Case "JPG", "PNG", "GIF"
				If Left (cFile.Name, 4) = "IMG_" Then
					sTargetName = "IMG_" & SetFileDateName
				End If
			Case "MP4", "MOV"
				If Left (cFile.Name, 4) = "DIV_" OR Left (cFile.Name, 4) = "IMG_" Then
					sTargetName = "DIV_" & SetFileDateName 
				End If
			End Select

			If sTargetName <> "" And Left (cFile.Name, Len(sTargetName)) <> sTargetName Then
				nIndex = 0
				If cFSO.FileExists (cFile.ParentFolder.Path & "\" & sTargetName & "." & cFSO.GetExtensionName (cFile.Name)) Then
					nIndex = 2
					Do While cFSO.FileExists (cFile.ParentFolder.Path & "\" & sTargetName & "_" & nIndex & "." & cFSO.GetExtensionName (cFile.Name))
						nIndex = nIndex + 1
					Loop
					sTargetName = sTargetName & "_" & nIndex
				End If
				cFile.Name = sTargetName & "." & cFSO.GetExtensionName (cFile.Name)
			End If
		End If
		Set cFile=Nothing
	End If
End Function

Function DoFolder (cFolder)
	Dim cSubFolder
	
	If Not cFolder Is Nothing Then
		For Each cFile in cFolder.Files
			SetFileDateName
		Next
		For Each cSubFolder in cFolder.SubFolders
			DoFolder cSubFolder
		Next
	End If
End Function

Set cFSO= Wscript.CreateObject("Scripting.FileSystemObject")   

For Each Arg In WScript.Arguments
	
	If cFSO.FileExists (Arg) Then
		Set cFile = cFSO.GetFile (Arg)
		SetFileDateName
	ElseIf cFSO.FolderExists (Arg) Then
		set cThisFolder = cFSO.GetFolder (Arg)
		DoFolder cThisFolder
	End If
Next

Set cFSO = Nothing

