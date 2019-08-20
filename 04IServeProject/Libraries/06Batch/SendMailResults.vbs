dim strSubject

dim strMailTO

dim strProjectRoot

 

strProjectRoot = getMachineEnviromentalVariable("User", "OBTAFProjectRoot")

 

executeExcelMacro(strProjectRoot & "\05ResultLog\TESTLOG.xls")

 

strSubject= WScript.Arguments(0)

strMailTO = WScript.Arguments(1)

strMailCC = WScript.Arguments(2)

 

SendAttach(strSubject)

 

Public Sub SendAttach(strSubject)

                dim objOutlk

                dim objMail

                dim strMsg

                dim strHtml

                const olMailItem = 0

               

                'Create a new message

                Set objOutlk = createobject("Outlook.Application")

                Set objMail = objOutlk.createitem(olMailItem)

                objMail.To = strMailTO                 

                objMail.cc = strMailCC

               

                objMail.subject = strSubject & "-" & Date

 

                objMail.attachments.add(strProjectRoot & "\05ResultLog\TestSuiteResult.html") ' Attachement in email

               

                strHtml = readTextFile(strProjectRoot & "\05ResultLog\TestSuiteResultSummary.html")

               

                objMail.HTMLBody  = HTMLHeader & vbcr & HTMLBody & vbcr & strHtml & vbcr & HTMLFooter

                objMail.Send

 

                Set objMail = Nothing

                Set objOutlk = Nothing

  End Sub

 

 Public Function HTMLHeader()

                HTMLHeader = "<!DOCTYPE html>"&VBCr&_

                                                                "<html lang=""en-US"">"&VBCr&_

                                                                                "<head>"&_

                                                                                "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html"">"&VBCr&_

                                                                                "<style>"&VBCr&_

                                                                                                "table {"&VBCr&_

                                                                                                "border: 1px ;"&VBCr&_

                                                                                    "font-family: arial;"&VBCr&_

                                                                                    "font-size: 14px;"&VBCr&_

                                                                                    "border-collapse: collapse;"&VBCr&_

                                                                                    "width: 20%;"&VBCr&_

                                                                                "}"&VBCr&_

                                                                                "th, td {"&VBCr&_

                                                                                    "border: 0px ;"&VBCr&_

                                                                                    "text-align: left;"&VBCr&_

                                                                                    "padding: 0px;"&VBCr&_

                                                                                "}"&VBCr&_

                                                                                "</style>"&VBCr&_

                                                                "</head>"&VBCr

End Function

 

Public Function HTMLBody()

                HTMLBody = "<body>"&VBCr&_

                                                                "<p style = 'font-size:15px'><strong>[This is an Auto Generated Email from I.Serve Automation Team]</strong></p>"&VBCr&_

                                                                "<p>Hi All,</p>"&VBCr&_

                                                                "<p>Please refer below table for Test execution summary. For failed testcase (if any) detailed analysis will be shared later.</p>"&VBCr

End Function

 

 Public Function HTMLFooter()

                strHostName = "Batch file executed in Host System: " & CreateObject("WScript.Shell").ExpandEnvironmentStrings("%COMPUTERNAME%") & vbcr

                HTMLFooter = "<p><strong>"&strHostName&"</strong></p>"&VBCr&_

                                                                "<p></p>"&VBCr&_

                                                                "<p>Thanks,</p>"&VBCr&_

                                                                "<p>I.Serve Automation Team</p>"&VBCr&_

                                                                "</body>"&VBCr&_

                                                                "</html>"

End Function

 

  Public Function readTextFile(strFileName)

 

                Dim strResult

                strResult=""

               

                Set fso=createobject("Scripting.FileSystemObject")

               

                'Open the file "qtptest.txt" in reading mode.

               

                Set qfile=fso.OpenTextFile(strFileName,1,True)

               

                'Read  the entire contents of  priously written file

               

                Do while qfile.AtEndOfStream <> true

                'Output --> The file will contain the above written content  in  single line.

                                strResult= strResult &" "&  qfile.ReadLine

                Loop

               

                'Close the files

                qfile.Close

                readTextFile = strResult

               

                'Release the allocated objects

                Set qfile=nothing

                Set fso=nothing

 

End Function

 

Public Function executeExcelMacro(strExcelPath)

                Set objExcel = CreateObject("Excel.Application")

                Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)

 

                objExcel.Application.Visible = True

 

                objExcel.Application.Run "ThisWorkBook.SelectTestcasesForExecution"

               

                objExcel.ActiveWorkbook.Save

                objExcel.ActiveWorkbook.Close

                objExcel.Application.Quit

               

                Set objExcel = Nothing

                Set objWorkbook = Nothing

End Function

 

Private Function getMachineEnviromentalVariable(strVariableType, strVariableName)

                'Declare Variables

                Dim WshShl, Shell, UserVar

                'Set objects

                Set WshShl = CreateObject("WScript.Shell")

                Set Shell = WshShl.Environment(strVariableType)

                getMachineEnviromentalVariable =  Shell(strVariableName)

                               

                'Cleanup Objects

                Set WshShl = Nothing

                Set Shell = Nothing

 

End Function