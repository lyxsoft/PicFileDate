On Error Resume Next

Dim cFSO, cFile

Function yearForFolder (cTime)
    yearForFolder = "" & Year(cTime)
End Function


Function CreateDateFolder ()
	Dim sTargetName

	If Not cFile Is Nothing Then
		sTargetName = cFile.ParentFolder.Path & "\" & yearForFolder (cFile.DateLastModified)

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

Function MoveFileToDateFolder ()
	Dim sTargetFolder, sThisFolder

	If Not cFile Is Nothing Then
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
End Function

Function StatusWindow ()
	On Error Resume Next

	Dim cHTML
	
	Set cHTML = wscript.Createobject("htmlfile")
	
	Set cIE = wscript.CreateObject("internetexplorer.application")
	With cIE
		.MenuBar=0 
		.AddressBar=0
		.ToolBar=0 
		.StatusBar=0
		.Resizable=0 

		.Width=450
		.Height=180
		.Left = Fix((cHTML.ParentWindow.Screen.AvailWidth-.Width)/2)
		.Top = Fix((cHTML.ParentWindow.Screen.AvailHeight-.Height)/2)

		.Navigate "about:blank"

		with .Document
			.Write "<html><title>Status</title>" & vbCr
			.write "<body scroll=no>" & vbCr
			.write "<font color=#0066ff size=2 face=""Arial""><div id=StatusText align=center>Please wait...</div></font>" & vbCr
			.write "</body></html>"
		End with
		
		.visible=1	
	End With
End Function

Function ShowStatus (sStatus)
	On Error Resume Next

	'If cIE.HWND = 0 Then
	'	Set cIE = Nothing
	'	StatusWindow
	'End If
	
	'wscript.CreateObject("Wscript.Shell").AppActivate cIE.HWND '"Status - Internet Explorer"
	cIE.Document.getElementById("StatusText").innerText = sStatus
End Function

Function CloseStatus ()
	On Error Resume Next

	If NOT cIE is Nothing Then
		cIE.Quit
	End If
	
	Set cIE = Nothing
End Function

Set cFSO= Wscript.CreateObject("Scripting.FileSystemObject")   

For Each Arg In WScript.Arguments

	If cFSO.FileExists (Arg) Then
		Set cFile = cFSO.GetFile (Arg)
		MoveFileToDateFolder
	ElseIf cFSO.FolderExists (Arg) Then
		StatusWindow
		set cThisFolder = cFSO.GetFolder (Arg)
		For Each cFile in cThisFolder.Files
			ShowStatus "Moving File [" & cFile.Name & "]"
			MoveFileToDateFolder
		Next
		CloseStatus
	End If
Next

Set cFSO = Nothing
