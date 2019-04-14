On Error Resume Next

Dim cFSO, cFile

Function getFileInfos ()
	Set cShellApp = Wscript.CreateObject ("Shell.Application")
	Set cFolder = cShellApp.NameSpace (cFile.ParentFolder.Path)
	Set cFileItem = cFolder.ParseName (cFile.Name)

	sResult = ""
	For nIndex = 0 to 512
		sFileInfo = cFolder.GetDetailsOf (cFileItem, nIndex)
		If sFileInfo <> "" Then
			If sResult <> "" Then
				sResult = sResult & chr(10)
			End If
			sResult = sResult & "[" & nIndex & "]-" & cFolder.GetDetailsOf (0, nIndex) & ":" & sFileInfo
		End If
	Next

	Set cFileItem = Nothing
	Set cFolder = Nothing
	Set cShellApp = Nothing
	getFileInfos = sResult
End Function

Set cFSO= Wscript.CreateObject("Scripting.FileSystemObject")   

For Each Arg In WScript.Arguments
	
	If (cFSO.FileExists(Arg)) Then
		Set cFile=cFSO.GetFile (Arg)
	Else
		Set cFile=Nothing
	End If
	If Not cFile Is Nothing Then
		MsgBox getFileInfos ()
		Set cFile=Nothing
	End If
Next

Set cFSO = Nothing
Set cShell = Nothing

