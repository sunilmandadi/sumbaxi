Dim arg
Dim TestPath
Dim strData
Dim strTestcase
Dim strUFTUpTimeinMinutes
strUFTUpTimeinMinutes = 120
arg = WScript.Arguments.Count
TestPath = WScript.Arguments(0)
strData = WScript.Arguments(1)
strTestcase = WScript.Arguments(2)
ExecuteTest TestPath,strData,strTestcase
Public Function setTestcaseName(strTest)
    set strTestcase = strTest
End Function
Public Function GetProcessRunningTime(strProcessName)
                Dim wShell,strText,arrTexts,strSystemRoot,strShellPath
                Set wShell = CreateObject("WScript.Shell")
                strSystemRoot = wShell.ExpandEnvironmentStrings("%SystemRoot%")
                strSystemRoot = Split(strSystemRoot,":")(0) &":\"& Split(strSystemRoot,":")(1)
                wShell.Run "powershell -noexit -command New-TimeSpan -Start (get-process "&strProcessName&").StartTime"
                Wait(3) 
				If wShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "AMD64" OR wShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "IA64" Then
                                strShellPath = strSystemRoot & "\\SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe"
                Else
                                strShellPath = strSystemRoot & "\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"
                End If
				wShell.AppActivate strShellPath
                '[regexpwndtitle value updated by  from strShellPath to Windows PowerShell]
                strText = Window( "object class:=ConsoleWindowClass","regexpwndtitle:=Windows PowerShell").GetVisibleText
                arrTexts = Split(strText,VBLf)
                SystemUtil.CloseProcessByName("powershell.exe")
                For i = 1 To UBound(arrTexts)
                                temp = Trim(UCase(Replace(arrTexts(i),vbCr,"")))
                                If temp = "TOTALMINUTES" Then
                                                GetProcessRunningTime = Trim(Split(Replace(arrTexts(i+11),vbCr,""),".")(0))
                                                Exit For
                                End If
                Next
                Set wShell = Nothing
End Function

Public Sub ExecuteTest(TestPath,strPath,strTestcaseName)
                Dim qtpApp
                Dim qtpTest
                Dim gstrTestDataPath
                Dim strUFTRunningTime
                Dim strProjectRoot
                strProjectRoot = getMachineEnviromentalVariable("User", "OBTAFProjectRoot")
                gstrTestDataPath = strPath
                Set qtpApp = CreateObject("QuickTest.Application")
               If Not qtpApp.Launched Then ' If QuickTest is not yet open
                                qtpApp.Launch ' Start QuickTest (with the correct add-ins loaded)
                End If
				qtpApp.Open TestPath, True              
                qtpApp.Options.Run.ImageCaptureForTestResults = "OnError"
                qtpApp.Options.Run.RunMode = "Fast"
                qtpApp.Options.Run.ViewResults = False
                qtpApp.Visible = false
                Set qtpTest = qtpApp.Test
                qtpTest.Environment.LoadFromFile strProjectRoot & "\04IServeProject\Config\env.xml"
              Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
                qtResultsOpt.ResultsLocation = strProjectRoot & "\05ResultLog\UFTResults\tempRes" ' Set the results location
				qtpTest.Run qtResultsOpt
				While qtpTest.IsRunning
                'Wait For Test To Finish
                Wend
                'qtpApp.Options.Run.ViewResults = false
                strUFTRunningTime = 121'GetProcessRunningTime("UFT")
                If strUFTRunningTime > strUFTUpTimeinMinutes Then
                                qtpTest.Close
                                qtpApp.Quit
                                Set qtResultsOpt = Nothing
                                Set qtpTest = Nothing
                                Set qtpApp = Nothing
                End If
                Dim strTestName
                strTestName = getMachineEnviromentalVariable("User", "testcasename")
End Sub

Public Function getMachineEnviromentalVariable(strVariableType, strVariableName)
                'Declare Variables
                Dim WshShl, Shell, UserVar              
                'Set objects
                Set WshShl = CreateObject("WScript.Shell")
                Set Shell = WshShl.Environment(strVariableType)
                getMachineEnviromentalVariable =  Shell(strVariableName)
               'Cleanup Objects
                Set WshShl = Nothing
				Set Shell = Nothing       
                Exit Function
End Function