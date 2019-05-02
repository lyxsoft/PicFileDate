On Error Resume Next

Dim cIE
Dim cFSO, cFile, cShell, cShellApp, cThisFolder
Dim sFileName, sFileDate
Dim sFDatePath
Dim nPos

Function timeStamp (cTime)
    timeStamp = Year(cTime) & _
		Right ("0" & Month(cTime),2) & _
		Right ("0" & Day(cTime),2)  & "_" & _
		Right ("0" & Hour(cTime),2) & _
		Right ("0" & Minute(cTime),2) & _
		Right ("0" & Second(cTime),2) 
End Function

Function timeString (cTime)
    timeString = Year(cTime) & "/" & _
		Right ("0" & Month(cTime),2) & "/" & _
		Right ("0" & Day(cTime),2)  & " " & _
		Right ("0" & Hour(cTime),2) & ":" & _
		Right ("0" & Minute(cTime),2) & ":" & _
		Right ("0" & Second(cTime),2) 
End Function

Function longToTime (nLong)
	longToTime = DateSerial(1970, 1, 1) + (nLong / 86400000)
End Function

Function getNextNumber ()
	If nPos > 0 Then
		If Mid (sFileName, nPos + 1, 1) = "_" OR Mid (sFileName, nPos + 1, 1) = " " Then
			getNextNumber = Mid (sFileName, nPos, 1)
			nPos = nPos + 2		
		ElseIf IsNumeric (Mid (sFileName, nPos, 2)) And (Mid (sFileName, nPos + 2, 1) = "_" OR Mid (sFileName, nPos + 2, 1) = " ") Then
			getNextNumber = Mid (sFileName, nPos, 2)
			nPos = nPos + 3
		Else
			nPos = 0
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
		.Width=450
		.Height=100
		.Left = Fix((cHTML.ParentWindow.Screen.AvailWidth-.Width)/2)
		.Top = Fix((cHTML.ParentWindow.Screen.AvailHeight-.Height)/2)

		.Resizable=0 
		.Navigate "about:blank"

		.visible=1
		
		with .Document
			.Write "<html><title>Status</title>" & vbCr
			.write "<body scroll=no>" & vbCr
			.write "<font color=#0066ff size=2 face=""Arial""><div id=StatusText align=center>Please wait...</div></font>" & vbCr
			.write "</body></html>"
		End with
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

Function getFileDate ()
	Dim myResult
	
	myResult = ""
	sFileName = cFile.Name
	
	If (UCASE (Left (sFileName, 5)) = "PANO_") And IsNumeric (Mid (sFileName, 6, 8)) AND IsNumeric (MID (sFileName, 15,6)) Then
		myResult = Mid (sFileName, 6, 4) & "/" & Mid(sFileName,10, 2) & "/" & Mid(sFileName, 12, 2) & " " & Mid(sFileName, 15, 2) & ":" & Mid(sFileName, 17, 2) & ":" & Mid(sFileName, 19, 2)
	ElseIf (UCASE (Left (sFileName, 4)) = "IMG_" OR UCASE (Left (sFileName, 4)) = "VID_") And IsNumeric (Mid (sFileName, 5, 8)) AND IsNumeric (MID (sFileName, 14,6)) Then
		myResult = Mid (sFileName, 5, 4) & "/" & Mid(sFileName, 9, 2) & "/" & Mid(sFileName, 11, 2) & " " & Mid(sFileName, 14, 2) & ":" & Mid(sFileName, 16, 2) & ":" & Mid(sFileName, 18, 2)
	ElseIf (UCASE (Left (sFileName, 11)) = "SCREENSHOT_") Then 'And IsNumeric (Mid (sFileName, 12, 4)) AND IsNumeric (MID (sFileName, 17,2)) AND IsNumeric (MID (sFileName, 20,2)) AND IsNumeric (MID (sFileName, 23,2)) AND IsNumeric (MID (sFileName, 26,2)) AND IsNumeric (MID (sFileName, 29,2)) Then
		'myResult = Mid (sFileName, 12, 4) & "/" & Mid(sFileName, 17, 2) & "/" & Mid(sFileName, 20, 2) & " " & Mid(sFileName, 23, 2) & ":" & Mid(sFileName, 26, 2) & ":" & Mid(sFileName, 29, 2)
		myResult = Mid (sFileName, 12, 4) & "/"
		nPos = 17
		myResult = myResult & getNextNumber () & "/"
		myResult = myResult & getNextNumber () & " "
		myResult = myResult & getNextNumber () & ":"
		myResult = myResult & getNextNumber ()
		If nPos = 0 Then
			myResult = ""
		End If		
	ElseIf IsNumeric (Left (sFileName, 13)) And (Mid (sFileName, 14, 1) = "." OR Mid (sFileName, 14, 1) = "(") Then
		myResult = timeString(longToTime(Left (sFileName, 13)))
	ElseIf (UCASE (Left (sFileName, 8)) = "MMEXPORT") And IsNumeric (MID (sFileName, 9, 13)) Then
		myResult = timeString(longToTime(MID (sFileName, 9, 13)))
	ElseIf (UCASE (Left (sFileName, 9)) = "MICROMSG.") And IsNumeric (MID (sFileName, 10, 13)) Then
		myResult = timeString(longToTime(MID (sFileName, 10, 13)))
	ElseIf (UCASE (Left (sFileName, 10)) = "WX_CAMERA_") And IsNumeric (MID (sFileName, 11, 13)) Then
		myResult = timeString(longToTime(MID (sFileName, 11, 13)))
	ElseIf IsNumeric (Left (sFileName, 4)) And Mid (sFileName, 5, 1) = "_" And IsNumeric (Mid(sFileName, 6, 1)) Then
		myResult = Left (sFileName, 4) & "/"
		nPos = 6
		myResult = myResult & getNextNumber () & "/"
		myResult = myResult & getNextNumber () & " "
		myResult = myResult & getNextNumber () & ":"
		myResult = myResult & getNextNumber ()
		If nPos = 0 Then
			myResult = ""
		End If
	End If
	
	If myResult = "" Then
		Dim cFolder, cFileItem
		
		Set cFolder = cShellApp.NameSpace (cFile.ParentFolder.Path)
		Set cFileItem = cFolder.ParseName (sFileName)
		myDate = cFolder.GetDetailsOf (cFileItem, 12) 'PictureCreateDate
		If myDate = "" Then
			myDate = cFolder.GetDetailsOf (cFileItem, 209) 'PictureCreateDate
		End If
		If myDate = "" Then
			myDate = cFolder.GetDetailsOf (cFileItem, 208) 'MediaCreateDate
		End If
		myResult = ""
		For nByte = 1 to Len(myDate)
			If Asc(Mid(myDate, nByte, 1)) <> 63 Then
				myResult = myResult & Mid(myDate, nByte, 1)
			End If
		Next
		
		Set cFileItem = Nothing
		Set cFolder = Nothing
	End If
	
	getFileDate = myResult
End Function

Function DoFile (bShowStatus)
	If Not cFile Is Nothing Then
		sFileDate = getFileDate()
		If sFileDate <> "" And sFileDate <> timeString (cFile.DateLastModified) Then
			If bShowStatus Then
				ShowStatus "Set File Date of [" & cFile.Name & "]."
			End If
			cShell.Run """" & sFDatePath & """ """ & cFile.Path & """ """ & getFileDate() & """", 0, 0
		End If
	End If
End Function

Function DoFolder (cFolder)
	Dim cSubFolder
	
	If Not cFolder Is Nothing Then
		For Each cFile in cFolder.Files
			DoFile (True)
		Next
		For Each cSubFolder in cFolder.SubFolders
			DoFolder cSubFolder
		Next
	End If
End Function

Set cFSO = Wscript.CreateObject("Scripting.FileSystemObject")   
Set cShellApp = Wscript.CreateObject ("Shell.Application")
Set cShell = Wscript.CreateObject("WScript.Shell")

Set cFile = cFSO.GetFile (Wscript.ScriptFullname)
sFDatePath = cFile.ParentFolder.Path & "\FDate.EXE"
Set cFile = Nothing
If (cFSO.FileExists(sFDatePath)) Then
	For Each Arg In WScript.Arguments
		
		If (cFSO.FileExists(Arg)) Then
			Set cFile = cFSO.GetFile (Arg)
			DoFile (False)
		ElseIf (cFSO.FolderExists(Arg)) Then
			Set cThisFolder = cFSO.GetFolder (Arg)
			StatusWindow
			DoFolder cThisFolder
			CloseStatus
		End If

		Set cFile=Nothing
	Next
Else
	MsgBox "File not found: " & sFDatePath
End If

Set cShell = Nothing
Set cShellApp = Nothing
Set cFSO = Nothing

