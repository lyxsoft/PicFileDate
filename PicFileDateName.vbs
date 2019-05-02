On Error Resume Next

Dim cIE
Dim cFSO, cFile, cThisFolder

Function timeStamp (cTime)
    timeStamp = Year(cTime) & _
		Right ("0" & Month(cTime),2) & _
		Right ("0" & Day(cTime),2)  & "_" & _
		Right ("0" & Hour(cTime),2) & _
		Right ("0" & Minute(cTime),2) & _
		Right ("0" & Second(cTime),2) 
End Function

Function isImgFile ()
	isImgFile = ""
	sFileName = cFile.Name
	
	If UCASE (Left (sFileName, 5)) = "PANO_" Then
		isImgFile = "PANO_"
	ElseIf UCASE (Left (sFileName, 4)) = "IMG_" Then
		isImgFile = "IMG_"
	ElseIf UCASE (Left (sFileName, 4)) = "DIV_" Then
		isImgFile = "DIV_"
	ElseIf (UCASE (Left (sFileName, 11)) = "SCREENSHOT_") Then 
		isImgFile = "SCREENSHOT_"
	ElseIf (UCASE (Left (sFileName, 8)) = "MMEXPORT") And IsNumeric (MID (sFileName, 9, 13)) Then
		isImgFile = "MMEXPORT_"
	ElseIf (UCASE (Left (sFileName, 9)) = "MICROMSG.") And IsNumeric (MID (sFileName, 10, 13)) Then
		isImgFile = "MICROMSG_"
	ElseIf (UCASE (Left (sFileName, 10)) = "WX_CAMERA_") And IsNumeric (MID (sFileName, 11, 13)) Then
		isImgFile = "WX_CAMERA_"
	ElseIf IsNumeric (Left (sFileName, 13)) And (Mid (sFileName, 14, 1) = "." OR Mid (sFileName, 14, 1) = "(") Then
		isImgFile = "IMG_"
	ElseIf IsNumeric (Left (sFileName, 4)) And Mid (sFileName, 5, 1) = "_" And IsNumeric (Mid(sFileName, 6, 1)) Then
		isImgFile = "IMG_"
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
		.Height=100
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

	If cIE.HWND = 0 Then
		Set cIE = Nothing
		StatusWindow
	End If
	
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

Function SetFileDateName (bShowStatus)
	Dim sTargetName
	Dim nIndex
	
	If Not cFile Is Nothing Then
		SetFileDateName = timeStamp (cFile.DateLastModified)

		If SetFileDateName <> "" Then
			Select Case UCase (cFSO.GetExtensionName (cFile.Name))
			Case "JPG", "PNG", "GIF"
				sTargetName = isImgFile ()
				If sTargetName <> "" Then
					sTargetName = sTargetName & SetFileDateName
				End If
			Case "MP4", "MOV"
				sTargetName = isImgFile ()
				If sTargetName <> "" Then
					sTargetName = sTargetName & SetFileDateName
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
				If bShowStatus Then
					ShowStatus "Set File name [" & cFile.Name & "]."
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
			SetFileDateName (True)
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
		SetFileDateName (False)
	ElseIf cFSO.FolderExists (Arg) Then
		StatusWindow
		set cThisFolder = cFSO.GetFolder (Arg)
		DoFolder cThisFolder
	End If
Next

Set cFSO = Nothing

