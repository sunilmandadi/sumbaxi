dim Client, Process, Feature,TestScenario, KWSheetPath
                Dim strProjectRoot, strProjectName
                Client=WScript.Arguments(0)
                Country=WScript.Arguments(1)
                Process=WScript.Arguments(2)
                Feature=WScript.Arguments(3)
                KWSheetPath=WScript.Arguments(4)
                TestScenario=WScript.Arguments(5)
                               strProjectRoot = getMachineEnviromentalVariable("User", "OBTAFProjectRoot")
                strProjectName = getMachineEnviromentalVariable("User", "OBTAFProjectName")
                               Set FSO = CreateObject("Scripting.FileSystemObject")
                  Set SasPth = FSO.GetFolder(strProjectRoot & "\" & strProjectName & "\Config")
                  Set TextStream = SasPth.CreateTextFile("env.xml")
    'Count the Number of Total Rows in the Master Sheet
   ' LastRow = ActiveCell.SpecialCells(xlCellTypeLastCell).Row
      'Construct the Header Part of the "UAT.bat" Batch File
    TextStream.WriteLine ("<Environment>")
                TextStream.WriteLine ("<Variable>")
                TextStream.WriteLine ("<Name>Client</Name>")
                TextStream.WriteLine ("<Value>" &Client&"</Value>")
                TextStream.WriteLine ("</Variable>")
                TextStream.WriteLine ("<Variable>")
                TextStream.WriteLine ("<Name>Country</Name>")
                TextStream.WriteLine ("<Value>" &Country&"</Value>")
                TextStream.WriteLine ("</Variable>")
                TextStream.WriteLine ("<Variable>")
                TextStream.WriteLine ("<Name>Process</Name>")
                TextStream.WriteLine ("<Value>" &Process&"</Value>")
                TextStream.WriteLine ("</Variable>")
                TextStream.WriteLine ("<Variable>")
                TextStream.WriteLine ("<Name>Feature</Name>")
                TextStream.WriteLine ("<Value>" &Feature&"</Value>")
                TextStream.WriteLine ("</Variable>")
                TextStream.WriteLine ("<Variable>")
                TextStream.WriteLine ("<Name>KWSheetPath</Name>")
                TextStream.WriteLine ("<Value>" &KWSheetPath&"</Value>")
                TextStream.WriteLine ("</Variable>")
                TextStream.WriteLine ("<Variable>")
                TextStream.WriteLine ("<Name>Testcase</Name>")
                TextStream.WriteLine ("<Value>" &TestScenario&"</Value>")
                TextStream.WriteLine ("</Variable>")
                TextStream.WriteLine ("</Environment>")
				
				
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