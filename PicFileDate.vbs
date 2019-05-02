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

Function getNext2Number ()
	If nPos > 0 Then
		If IsNumeric (Mid (sFileName, nPos, 1)) Then
			If (NOT IsNumeric (Mid (sFileName, nPos + 1, 1))) Then
				getNext2Number = Mid (sFileName, nPos, 1)
				nPos = nPos + 2		
			Else
				getNext2Number = Mid (sFileName, nPos, 2)
				If (NOT IsNumeric (Mid (sFileName, nPos + 2, 1))) Then
					nPos = nPos + 3
				Else
					nPos = nPos + 2
				End If
			End If
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

Function isFileID (sID)
	If sID <> "" Then
		isFileID = UCase (Left (sFileName, Len(sID))) = UCase(sID)
		If isFileID Then
			nPos = Len (sID) + 1
		End If
	Else
		isFileID = True
		nPos = 1
	End If
End Function

Function getFileNameDateTime ()
	getFileNameDateTime = ""

	If nPos > 0 Then
		If IsNumeric (Mid (sFileName, nPos, 8)) AND IsNumeric (MID (sFileName, nPos + 9,6)) Then
			' Type: ID_YYYYMMDD_HHMMSS
			getFileNameDateTime = Mid (sFileName, nPos, 4) & "/" & Mid(sFileName, nPos + 4, 2) & "/" & Mid(sFileName, nPos + 6, 2) & " " & _
								  Mid (sFileName, nPos + 9, 2) & ":" & Mid(sFileName, nPos + 11, 2) & ":" & Mid(sFileName, nPos + 13, 2)
		ElseIf IsNumeric (Mid (sFileName, nPos, 4)) AND (NOT IsNumeric (Mid (sFileName, nPos + 4, 1))) Then
			' Type: ID_YYYY-MM-DD-HH-MM-SS
			getFileNameDateTime = Mid (sFileName, nPos, 4) & "/"
			nPos = nPos + 5
			getFileNameDateTime = getFileNameDateTime & getNext2Number () & "/"
			getFileNameDateTime = getFileNameDateTime & getNext2Number () & " "
			getFileNameDateTime = getFileNameDateTime & getNext2Number () & ":"
			getFileNameDateTime = getFileNameDateTime & getNext2Number () & ":"
			getFileNameDateTime = getFileNameDateTime & getNext2Number ()
			If nPos = 0 Then
				getFileNameDateTime = ""
			End If
		ElseIf IsNumeric (Mid (sFileName, nPos, 13)) And (NOT IsNumeric (Mid (sFileName, nPos + 14, 1))) Then
			' File name is the long number of seconds from 1970-1-1
			getFileNameDateTime = timeString(longToTime(MID (sFileName, nPos, 13)))
		End If
	End If
End Function

Function getFileDate ()
	getFileDate = ""
	sFileName = cFile.Name
	
	If isFileID ("PANO_") OR isFileID ("IMG_") OR isFileID ("DIV_") OR isFileID ("VID_") OR isFileID ("SCREENSHOT_") OR isFileID ("SCANNER_") OR isFileID ("MMEXPORT") OR isFileID ("MICROMSG.") OR isFileID ("WX_CAMERA_") Then
		getFileDate = getFileNameDateTime ()
	ElseIf isFileID ("") Then
		getFileDate = getFileNameDateTime ()
	End If

	If getFileDate = "" Then
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
		getFileDate = ""
		For nByte = 1 to Len(myDate)
			If Asc(Mid(myDate, nByte, 1)) <> 63 Then
				getFileDate = getFileDate & Mid(myDate, nByte, 1)
			End If
		Next
		
		Set cFileItem = Nothing
		Set cFolder = Nothing
	End If
End Function

Function DoFile (bShowStatus)
	If Not cFile Is Nothing Then
		sFileDate = getFileDate()
		If sFileDate <> "" And sFileDate <> timeString (cFile.DateLastModified) Then
			If bShowStatus Then
				ShowStatus "Set File Date of [" & cFile.Name & "]."
			End If
			cShell.Run """" & sFDatePath & """ """ & cFile.Path & """ """ & sFileDate & """", 0, 0
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

