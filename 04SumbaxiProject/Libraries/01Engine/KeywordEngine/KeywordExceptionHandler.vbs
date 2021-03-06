'+++++++++++++++++++++++++++++++++++ File Header Information ++++++++++++++++++++++++++++++++++++++++++++++
	'<Summary>  This file contains all the class declarations and 
								'functions for the  Keyword Exception Handling. 
                                '</summary>


   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
	Public Function cExceptionHandling()
		set cExceptionHandling=  new clsKeywordExceptionHandling
	End Function

Class clsKeywordExceptionHandling

   Public Function HandleExpectedException()
	  Print("Desired Exception was "& gstrExceptionType)

	  If  checkIfNull(gstrExceptionType) Then 'Exception Not Expected

			 If  ( err.number <> 0)  Then 'Exception Not Expected But Came

				LogMessage "WARN","Exception", " Error Reported in the Execution of the Keyword:  " &  " is  " & Err.Number & " " & Err.description & " " & Err.Source &  " " & Err.HelpContext , false
			    HandleExpectedException = false

			 else 'Exception Not Expected and didn't Come
			     HandleExpectedException = true		 

			End If
		  
	 else '//Exception  Expected

			 If  ( err.number <> 0)  Then 'Exception  Expected But Came
				
			    HandleExpectedException=true

			 else 'Exception  Expected But  Didn' Come
				LogMessage "RSLT", "Exception Handling", "Specified Exception " & gstrExceptionType & " did not get raised", false
			     HandleExpectedException = false		 

			End If	
	
	  End If

   End Function

   Public Function handleException()

		  Dim intErrNum,strErrDesc,strErrSource,bHandleException,bExpHandler

			intErrNum = err.number
			strErrDesc = err.Description
			strErrSource =  err.source
			strErrHelpContext = Err.HelpContext
			err.clear
	
		  If   checkIfNull(gstrExceptionType) Then 'Exception not expected
	
				  'Exception not expected but Came
				  If intErrNum <>0 Then 
					  
					LogMessage "RSLT","Exception Handling", "Final Error Reported in the Execution of the Keyword:  " & gstrKeyword & " is  " & intErrNum & " " & strErrDesc & " " & strErrSource &  " " & strErrHelpContext , False
					handleException = false
	
				'Exception not Expected Not came
				 else
					handleException = true
				  End If
				  
		else 'Exception Expected and Came now Execute Exception Keywords
					Print("Handling Exception")
					 Dim arrExceptArg (2)
	
					 arrExceptArg (0) = gstrExceptionType
					 arrExceptArg(1) = gstrExpDetails
					 arrExceptArg(2) = gstrExpAction
	
					strSQLStatement="Select  ExceptionNumber, ExceptionClass  from [ExceptionClassMap$] where   Exception ='"& gstrExceptionType &"' "
					Set clsObj = cExcelDataEngine()
					
					strData=clsObj.FetchExcelValue (strSQLStatement,gstrExceptionMap)
					'Print("strData:"&strData(0,1))
					strExceptClassName=strData(0,1)
					strExceptNum=strData(0,0)
					
	
			
					If strExceptNum = cStr(intErrNum) Then ' Expected and Observed Exception Matched

							'Execute Exception Keyword
							print("Going for the Execution of  Exception Keyword  "  & gstrExceptionType)
							logMessage "WARN",  "Exception Handling",  "Going for the Execution of  Exception Keyword  "  & gstrExceptionType , True

							Set clsExecuter = cKeywordEngine()
							Dim bException
							bExceptionKeywordExec = false
		
							If clsExecuter.executeRuntimeClass(strExceptClassName, arrExceptArg) then
									bExceptionKeywordExec = true
									print("Successfully executed  the Execution of  Exception Keyword  "  & gstrExceptionType)
									logMessage "WARN",  "Exception Handling",  "Successfully executed  the Execution of  Exception Keyword  "  & gstrExceptionType , True
							else
								handleException = false
							End if 							
		
							If bExceptionKeywordExec Then

									print("Now Going for the Execution of  Exception Action "& gstrExpAction)
									logMessage "WARN",  "Exception Handling",  "Now Going for the Execution of  Exception Action "& gstrExpAction , True

									'Execute Exception Action
		
									If  not isnull (gstrExpAction)Then
										bStatus = clsExecuter.executeRuntimeClass(arrExceptArg)
										handleException = bStatus
									else
										 'logMessage "RSLT","Exception Verification","Expected exceptionnot raised, Actual Exception Raised was  "& intErrNum & ": "&strErrDesc,false
										handleException = true
									End If

							   End If
		
							   If err.number <> 0 Then
                                   Print(" An Exception occured in Exception handler")
								    logMessage "WARN",  "Exception Handling",  "Exception Reported  in the Exception Handling of Keyword " & gstrKeyword & " is  " & Err.Number & " " & Err.description & " " & Err.Source &  " " & Err.HelpContext , False
								   handleException=false
							   End If
					  else' Expected and Observed Exception Not Matched
							  logMessage "RSLT",  "Exception Verification",  "Expected exception " & gstrExceptionType  & " not raised,  Actual Exception Raised was  " & intErrNum & " " & strErrDesc & " " & strErrSource &  " " & strErrHelpContext, false
							  handleException=false
					  End If
		 End If

   End Function

End Class
