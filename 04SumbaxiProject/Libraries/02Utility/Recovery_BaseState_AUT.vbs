'  ******************************************** Global variables*************************************************
'''<summary>This file holds all the Global variables resused in OBTAF for specific projects</summary>

'''  ***************************************** Project Specific Variables*********************************************
Option explicit

Dim oApp
Dim thirdlevel


Public Function setBaseAppState()
	Set oApp = cBrowserApplication()
	oApp.setBaseAppState(gstrApplicationURL)
End Function

Public Function setBaseAppStateBDT()
	Set oApp = cBrowserApplication()
	oApp.setBaseAppState_BDT()
End Function

Public Sub LogOff()
	If (bLogOff = True) Then
	 Set oApp = cBrowserApplication()
	 oApp.logOff()
	End If
End sub

Public Function Launch(apptype, val)		'*** Function to Launch Application Under Test
	If "website" = apptype Then
		thirdlevel = ""
		LogMessage  "INFO", "Initialization", "Initializing Framework for  web Site", true
		level = split(webLevels, leveldelimiter, -1, 1)
		desc = split(webLevelsDesc, leveldescdelimiter, -1, 1)
		object = split(objects, objectdelimiter, -1, 1)
		objectDescription = split(objectsDescription, objectsDescriptiondelimiter, -1, 1)
		
		killAndOpenBrowserNavigate(val)
		
	ElseIf "window" = apptype then
		level = split(winLevels, leveldelimiter, -1, 1)
		desc = split(winLevelsDesc, leveldescdelimiter, -1, 1)
		object = split(objects, objectdelimiter, -1, 1)
		objectDescription = split(objectsDescription, objectsDescriptiondelimiter, -1, 1)
		
		InvokeApplication "C:\Program Files\HP\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe"
	End If
	
	initialized = true
	Launch = true
	
End Function

Public Sub killAndOpenBrowserNavigate(URL)
	
	Dim intMaxAttemptsToOpenBrowser
	Dim intBrowserOpeningMaxTimeOut
	Dim strApplication
	Dim strAppName
	Dim lstChilds
	Dim oDesc
	
	intMaxAttemptsToOpenBrowser = 2
	intBrowserOpeningMaxTimeOut = 5
	
	Select Case Ucase(gstrBrowser)
		Case "CHROME"
			strApplication = "chrome.exe"
			strAppName = "Google Chrome"
		Case "IE8"
			strApplication = "iexplore.exe"
			strAppName = "Internet Explorer"
	End Select
	
	If bReLaunchBrowser Then
		CloseBrowsers
		SystemUtil.Run strApplication,URL
		SyncBrowser
		bReLaunchBrowser = False
	ElseIf Environment.Value("intBrowserLaunchCounter") = 0 Then
	    SystemUtil.Run strApplication,URL
	    SyncBrowser
	End If
	
	If Not IsProcessRunning(strApplication) OR Not IsBrowserExist Then
		Print VbTab & VbTab & strAppName & " is not initialized in: 1 Attempt"
		BrowserRecovery intBrowserOpeningMaxTimeOut, intMaxAttemptsToOpenBrowser,strApplication,strAppName,URL
		Browser("index:=0").Sync
	Else
		Print VbTab & VbTab & strAppName & " initialized successfully at : "& Now
	End If
	
	'To close if chrome Application Error Dialog window
	Dim chromeDlg
	Set chromeDlg = Dialog("window id:=0","nativeclass:=#32770")
	If Not chromeDlg.Exist(0.5) Then
		Wait(0.5)
	else
		Set oDesc = Description.Create()
		oDesc("micclass").Value = "WinButton"
		oDesc("nativeclass").Value = "Button"
		set lstChilds = chromeDlg.ChildObjects(oDesc)
		If lstChilds.count <> 0 Then
			lstChilds(0).highlight
			lstChilds(0).click
			Wait(1)
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Click button Ok" ,false
	Else 
		LogMessage "RSLT","Verification","Dialog Window button Ok is clicked successfully",true
	End If
		End If			
	End if	
	Set chromeDlg = Nothing
	
	Wait(5)
	On Error Resume next
	Browser("index:=0").Sync
	On Error Goto 0
End Sub


Public Function BrowserRecovery(iTimeOut,iAttempt,sProcess,sName,strURL)
	Dim i,j,bBrowserOpened
	bBrowserOpened  = False
	For i = 1 To iAttempt
		SystemUtil.CloseProcessByName(sProcess)
		If IsProcessRunning(sProcess) Then Wait(0.5)
		
		SystemUtil.Run sProcess,strURL
		SyncBrowser
		
		For j = 1 To iTimeOut
			If IsProcessRunning(sProcess) And IsBrowserExist Then
				bBrowserOpened = True
				Print VbTab & VbTab & sName & " initialized successfully at : "& Now & ": in " & i+1 &" Attempt"
				Exit Function
			Else
				If j = iTimeOut  Then
					Print VbTab & VbTab & sName & " is not initialized within:" & iTimeOut & " seconds in " & i+1 &" Attempt"
				End If
			End If
		Next
	Next
	Wait(4)
	If Not bBrowserOpened Then
		Print VbTab & VbTab & sName & " is not initialized after trying to open "& iAttempt &" times within:" & iTimeOut*iAttempt & " seconds."_
		& " Stoping current test execution. Please check for any issue and restart the execution : " & Now
		Msgbox sName & " is not initialized after trying to open "& iAttempt &" times within:" & iTimeOut*iAttempt & " seconds" & VBCr &_
		"Stoping current test execution" & VBCr & "Please check for any issue and restart the execution",0,"Application Initialization"
		ExitTest
	End If
End Function

Public Function CloseBrowsers

'	If bSetBaseState = True And bLogOff = True Then
'		SystemUtil.CloseProcessByName("iexplore.exe")  'Satheesh
'		SystemUtil.CloseProcessByName("chrome.exe")
'	End If

	'[Modified by  - To control browsers closing ]
	If bCloseBrowsers Then
		SystemUtil.CloseProcessByName("chrome.exe")
		SystemUtil.CloseProcessByName("iexplore.exe")
		If IsProcessRunning("chrome.exe") Then Wait(0.5)
		If IsProcessRunning("iexplore.exe") Then Wait(0.5)
		bCloseBrowsers = False
	End If
	

	'******* Close all the browser except the QC
	
	'	Dim intIndex
	'	Set oDesc=Description.Create
	'	oDesc("micclass").Value="Browser"
	'	intIndex=0
	'	While Browser("micclass:=Browser","index:="&intIndex).exist(0) and intIndex<Desktop.ChildObjects(oDesc).count
	'		If instr(1, Browser("micclass:=Browser","index:="&intIndex).getRoProperty("name"),"HP Quality Center 10.00") = 0 Then
	'		Browser("micclass:=Browser","index:="&intIndex).Close
	'		'SystemUtil.CloseProcessByHwnd( Browser("micclass:=Browser","index:="&intIndex).getRoProperty("hwnd"))
	'		else
	'		intIndex=intIndex+1
	'		End if
	'	 Wend
	'**********
	
End Function

Public Function DisconnectAll()
   CloseBrowsers()	' close browser  at the end of test 
End Function

Public Function CloseAllOpenBrowsers		' *****Function to Close all open browsers
	If Dialog("Internet Explorer").Exist(1) Then
		Dialog("Internet Explorer").Close
	End If 

	While Browser("micclass:=Browser", "index:=1").Exist (1)
		Browser("index:=1").Close
	Wend
End Function

Function BrowserActivate(Object)
	Dim hWnd
	hWnd = Object.GetROProperty("hwnd")
	On Error Resume Next
	Window("hwnd:=" & hWnd).Activate
	
	If Err.Number <> 0 Then
		Window("hwnd:=" & Browser("hwnd:=" & hWnd).Object.hWnd).Activate
		Err.Clear
	End If
	'  On Error Goto 0
End Function

Public Function IsProcessRunning(ByVal strProcessName)
	Dim strComputerName
	Dim objWMIService
	Dim strWMIQuery
	
	IsProcessRunning = False
	strComputerName = "."
	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	Set objWMIService = GetObject("winmgmts:\\" & strComputerName & "\root\cimv2")
	If objWMIService.ExecQuery(strWMIQuery).Count > 0 Then 
		IsProcessRunning = True
	End If
End Function

Public Function IsBrowserExist()
	Dim i
	IsBrowserExist = False
	For i = 1 To 10 Step 1
		If Browser("index:=0").Exist(1) Then
			IsBrowserExist = True
			Window("regexpwndtitle:=Windows Internet Explorer|Google Chrome|IServe","regexpwndclass:=IEFrame|Chrome_WidgetWin_1").Maximize
			Exit For
		End If
	Next
End Function
