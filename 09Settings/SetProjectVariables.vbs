Dim strProjectRoot
Dim strProjectName
Dim strTestCaseName

strProjectRoot = FindProjectRootFolderDrive & ":\sumbaxi"
strProjectName = "04SumbaxiProject"
strTestCaseName = "Random"

SetProjectVariables			'[Set User Project Variable before starting execution]

Private Function SetProjectVariables()
	Dim intSecondsToWait
	Dim wshShell
	Dim wshObject
	
	intSecondsToWait = 3
	
	Set wshShell = CreateObject("WScript.Shell")
	Set wshObject = wshShell.Environment("User")
	
	wshObject("OBTAFProjectRoot") = strProjectRoot
	wshShell.Popup "OBTAFProjectRoot = " & wshObject("OBTAFProjectRoot"), intSecondsToWait	'[Display the current value for Framework Root Folder]
	
	wshObject("OBTAFProjectName") = strProjectName
	wshShell.Popup "OBTAFProjectName = " & wshObject("OBTAFProjectName"), intSecondsToWait  '[Display the current value for Project name]
	
	wshObject("testcasename") = strTestCaseName
	wshShell.Popup "testcasename = " & wshObject("testcasename"), intSecondsToWait			'[Display the current value User Name]
	
	Set wshObject = Nothing
	Set wshShell = Nothing
End Function

Private Function FindProjectRootFolderDrive()
	On Error Resume Next
	bFolderExists = False
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set collDrives = fso.Drives
	For Each objDrive in collDrives
		For Each objFolder in fso.GetFolder(objDrive.RootFolder).SubFolders
			If Instr(1,ObjFolder.Name,"sumbaxi") > 0 Then
				FindProjectRootFolderDrive = objDrive.DriveLetter
				bFolderExists = True
				Exit For
			End If
		Next
		If bFolderExists Then Exit For
	Next
	On Error GoTo 0
	Set collDrives = Nothing
	Set fso = Nothing
End Function
