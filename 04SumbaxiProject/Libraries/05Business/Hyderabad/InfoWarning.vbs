
		Public Function VerifyInfoWarning(strInfoWarnLocation, strStatus, strInfoMessage, strPageName, strPageNumber, strLinkToClick)
			   Dim bVerifyInfoWarning:bVerifyInfoWarning = True

				If Not IsNull(strLinkToClick) Then
					Set objLinkToCLick = Description.Create()
					objLinkToCLick("innerhtml").Value = Trim(strLinkToClick)
					objLinkToCLick("class").value =  "v-button-caption|"

					Browser("Browser_iCall_Home").Page("ICall_InfoWarn").WebElement(objLinkToCLick).highlight
					Browser("Browser_iCall_Home").Page("ICall_InfoWarn").WebElement(objLinkToCLick).Click
					WaitForICallLoading
					Wait 2
				End If

				If Ucase(Trim(strInfoWarnLocation)) = "PAGE" Then
					Set objInfoWarnBtn = btn_InfoWarn_Page
					Set objInfoWarnMsg = txt_InfoMsg_Page
					Set objInfoWarnOK  = btn_InfoWarn_Page_OK
					Set objLnk_Next = lnk_Next_Page
					Set objLnk_Previous = lnk_Previous_Page
					Set objPageNumber = txt_PageNumber_Page
		
				ElseIf Ucase(Trim(strInfoWarnLocation)) = "POPUP" Then
					Set objInfoWarnBtn = btn_InfoWarn_Popup
					Set objInfoWarnMsg = txt_InfoMsg_Popup
					Set objInfoWarnOK = btn_InfoWarn_Page_OK
					Set objLnk_Next = lnk_Next_Popup
					Set objLnk_Previous = lnk_Previous_Popup
					Set objPageNumber = txt_PageNumber_Popup
				End If
		
				If Ucase(Trim(strStatus)) = "ENABLED" Then
					Dim bStatus:bStatus = False
'					Check the info icon is enabled in the first page  - Click verify the message is matching

					GoToFirstRecord objLnk_Previous, objPageNumber
	
					Print "Check the info icon is enabled in the first page. Page -  "&strPageName
					bStatus = VerifyInfoEnabledAndMsg(objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "MATCH")
					If Not bStatus Then
						bVerifyInfoWarning = False
					End If
			
					'Check the info icon is enabled after reaching the expected page  - Click and verify the message is matching
					If Not IsNull(strPageNumber) Then
						Print "Check the info icon is enabled after reaching the penultimate Service call's last record. Page -  "&strPageName
		
						bStatus = GoToExpectedRow(strPageNumber, objPageNumber, objLnk_Next)
						If Not bStatus Then
							bVerifyInfoWarning = False
						End If
	
						'Check enabled and message
						bStatus = VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "MATCH")
						If Not bStatus Then
							bVerifyInfoWarning = False
						End If
		
						'Click the Next  once and  the check  the icon should be disabled 
						Print "Click the Next  once and  the check  the icon should be disabled. Page -  "&strPageName
						objLnk_Next.RefreshObject
						Wait 1
						objLnk_Next.WebElement("class:=v-button-caption").click	
						WaitForICallLoading
						bIconEnabled = isInfoWarnEnabled(objInfoWarnBtn, strPageName)
		
						If  bIconEnabled Then 'Expected is disabled
							'If the icon is enabled - Check the Info message should not match with the existing
							Print "If the icon is enabled - Check the Info message should not match with the existing. Page -  "&strPageName
							bMessageExist = VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "NOMATCH")
							If bMessageExist = True Then 'As this is No Match, True is Pass case
								bVerifyInfoWarning = True
							Else ' Pass
								bVerifyInfoWarning = False
							End If
						End If

						' Click until you reach Next button is disabled
						Print "Click until you reach Next button is disabled and  check info is not enabled, if enabled message should not match. Page -  "&strPageName
						GoToLastRecord(objLnk_Next)

						'Check  the icon should be disabled 
						bIconEnabled = isInfoWarnEnabled(objInfoWarnBtn, strPageName)

						If  bIconEnabled Then 'Expected is disabled

							'If the icon is enabled - Check the Info message should not match with the existing
							Print "If the icon is enabled - Check the Info message should not match with the existing. Page -  "&strPageName

							bMessageExist = VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "NOMATCH")
							If bMessageExist = False Then 'As this is No Match, True is Pass case
								bVerifyInfoWarning = False
							End If
						End If
					Else ' For Alerts Page

						'Click until you reach Next button is disabled
						Print "Click until you reach Next button is disabled and  check info is enabled, if enabled message should match. Page -  "&strPageName
						GoToLastRecord(objLnk_Next)

						'Check  the icon should be disabled 
							bIconEnabled = isInfoWarnEnabled(objInfoWarnBtn, strPageName)
							If  bIconEnabled Then 'Expected is enabled
							'If the icon is enabled - Check the Info message should not match with the existing
								Print "If the icon is enabled - Check the Info message is matching  with the existing. Page -  "&strPageName
	
								bMessageExist = VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "MATCH")
								If bMessageExist = False Then 'As this is Match, True is Pass case
									bVerifyInfoWarning = False
								End If
							End If							
					End If
			
			ElseIf Ucase(Trim(strStatus)) = "DISABLED"  Then

					'Check  the icon should be disabled 
					bIconEnabled = isInfoWarnEnabled(objInfoWarnBtn, strPageName)

					 If  bIconEnabled Then 'Expected is disabled
						'If the icon is enabled - Check the Info message should not match with the existing
						Print "If the icon is enabled - Check the Info message should not match with the existing. Page -  "&strPageName
						LogMessage "WARN","Verification","Info warn in not disabled as expected.and Pop up is displayed with message: " & strCurrentInfoWarnMessage ,False
						bMessageExist = VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "NOMATCH")
						If bMessageExist = False Then 'As this is No Match, True is Pass case
							bVerifyInfoWarning = False
						End If
					Else
						LogMessage "RSLT","Verification","Info icon is disabled as expected",True
					End If	

					'Click until you reach Next button is disabled
					Print "Click until you reach Next button is disabled and  check info is enabled, if enabled message should match. Page -  "&strPageName
					GoToLastRecord(objLnk_Next)

					'Check  the icon should be disabled 
						bIconEnabled = isInfoWarnEnabled(objInfoWarnBtn, strPageName)
	
					 If  bIconEnabled Then 'Expected is disabled

						'f the icon is enabled - Check the Info message should not match with the existing
							Print "If the icon is enabled - Check the Info message should not match with the existing. Page -  "&strPageName
							LogMessage "WARN","Verification","Info warn in not disabled as expected.and Pop up is displayed with message: " & strCurrentInfoWarnMessage ,False
							bMessageExist = VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, "NOMATCH")
							If bMessageExist = False Then 'As this is No Match, True is Pass case
								bVerifyInfoWarning = False
							End If
					Else
								LogMessage "RSLT","Verification","Info icon is disabled as expected",True
					End If
			End If
	
			If objInfoWarnOK.Exist(0) Then
				objInfoWarnOK.Click
			End If
	
			If bVerifyInfoWarning Then
						LogMessage "RSLT","Verification","Info Warn details  are working as expected",True
						VerifyInfoWarning = True
			Else					    
						LogMessage "WARN","Verification","Info Warn details are not working as expected",False								
						VerifyInfoWarning = False
			End If
		End Function
		
		
		Public Function GetPageNumber(objPageNumber)
			Dim intCurrentPageNumber:intCurrentPageNumber = 0
			If objPageNumber.Exist(0) Then
					intCurrentPageNumber = Cint(Replace(objPageNumber.GetROProperty("innertext"),":",""))
			End If
			GetPageNumber = intCurrentPageNumber
		End Function
		
		
		Public Function isInfoWarnEnabled(objInfoWarnBtn, strPage)
			 WaitForICallLoading
			  objInfoWarnBtn.RefreshObject
			  If objInfoWarnBtn.Exist(0) Then
				  bEnabled =matchStr(objInfoWarnBtn.GetROProperty("innerhtml"),"INFO.GIF")
		
				  If bEnabled Then
						isInfoWarnEnabled = True
						Print  "Info Icon is enabled in page "&strPage
				  Else
						isInfoWarnEnabled = False
						Print "Info Icon is disabled in page "&strPage
				  End If
			  Else
					LogMessage "WARN","Verification","Object doesnot exist in page "&strPage&". Check the info icon property",False
					isInfoWarnEnabled = False	  
			  End If
		End Function
		
		
		Public Function VerifyInfoEnabledAndMsg (objInfoWarnBtn, objInfoWarnMsg, objInfoWarnOK, strInfoMessage, strInfoWarnLocation, strPageName, strMatch)
'			  WaitForICallLoading
			   strCurrentInfoWarnStatus = isInfoWarnEnabled(objInfoWarnBtn, strPageName)
			   If strCurrentInfoWarnStatus = True Then
				   'pass 1
					LogMessage "RSLT","Verification","Info Icon is enabled in page "&strPageName,True
				   objInfoWarnBtn.RefreshObject
				   objInfoWarnBtn.Click ' Click the Info icon button
				   WaitForICallLoading
'					   Wait 1
				   objInfoWarnMsg.RefreshObject
				   strCurrentInfoWarnMessage = objInfoWarnMsg.GetROProperty("innertext") 'Get the info message
					Wait 1
					
	
				   bInfoMsgMatch = matchStr(Ucase(Trim(strCurrentInfoWarnMessage)), Ucase(Trim(strInfoMessage))) 'check the expected message exist in the info warn message text
	
				   If strMatch = "MATCH" Then
						   If bInfoMsgMatch Then  'Info message matches
							   'Pass 2
								LogMessage "RSLT","Verification","Expected Info Warn message exists. Actual: " & strCurrentInfoWarnMessage & " | Expected: " & strInfoMessage,True
								VerifyInfoEnabledAndMsg = True
						   Else 'Info message doesnot  match
								'Fail	
								LogMessage "WARN","Verification","Info Warn message doesn't match. Actual: " & strCurrentInfoWarnMessage & " | Expected: " & strInfoMessage,False								
								VerifyInfoEnabledAndMsg = False
						   End If
				   ElseIf strMatch = "NOMATCH" Then
							'LogMessage "WARN","Verification","Info warn in not disabled as expected.and Pop up is displayed with message: " & strCurrentInfoWarnMessage ,False
						   If bInfoMsgMatch Then  'Info message matches
								LogMessage "WARN","Verification","Message displayed in popup is: " & strCurrentInfoWarnMessage & " | should not match with : " & strInfoMessage,False
								VerifyInfoEnabledAndMsg = False
						   Else 'Info message doesnot  match
								'Pass
								LogMessage "RSLT","Verification","Info Warn message doesn't match with actaul. Actual: " & strCurrentInfoWarnMessage & " | Expected: " & strInfoMessage,True								
								VerifyInfoEnabledAndMsg = True
						   End If
				   End If
				   If UCase(Trim(strInfoWarnLocation)) = "PAGE" Then
						objInfoWarnOK.RefreshObject
						objInfoWarnOK.Click
					End If
				Else  ' Info icon disabled
						'Fail
						LogMessage "WARN","Verification","Info Icon is disabled or Does not exist or Property changed in page "&strPageName,False
						VerifyInfoEnabledAndMsg = False
				End If
		End Function
		
		
		Public Function GoToFirstRecord(objLnk_Previous, objPageNumber)
				If objLnk_Previous.Exist(0) Then
						Dim bPreviousPageExist, intPageCount
						Do 
							objLnk_Previous.RefreshObject
							bPreviousPageExist =matchStr(objLnk_Previous.GetROProperty("outerhtml"),"v-disabled")	
									If Not bPreviousPageExist Then
										objLnk_Previous.WebElement("class:=v-button-caption").click	
										WaitForICallLoading
										intPageCount = intPageCount + 1
									End If
						Loop while Not  ( bPreviousPageExist Or  intPageCount > 100)  'To avoid looping
						If  intPageCount > 100 Then
							Print  "Previous Link was clicked more than 100 times"
							GoToFirstRecord = False					
						Else
									objPageNumber.RefreshObject
									intPage = GetPageNumber(objPageNumber)
									If  intPage = 1 Then
										LogMessage "RSLT","Verification", "Page number in first page is equal to 1", True
									Else
										LogMessage "WARN","Warning", "Page number not equals to 1 or Check Pagination control does not exists", False
									End If
							GoToFirstRecord = True
						End If            
				Else
		'			LogMessage "WARN","Verification", "Previous Link was not available in the screen or property changed", False
					Print "Previous Link was not available in the screen or property changed"
					GoToFirstRecord = False			   
				End If
		End Function
		
		
		Public Function GoToExpectedRow(strPageNumber, objPageNumber, objLnk_Next)
				Dim bGoToExpectedRow:bGoToExpectedRow = True
		
				If objLnk_Next.Exist(0) Then
						Dim bNextPageExist, intCurrentPageNumber
						intCurrentPageNumber = GetPageNumber(objPageNumber)
		
						Do While(Cint(intCurrentPageNumber) < Cint(strPageNumber))
								  
								   bNextPageExist =matchStr(objLnk_Next.GetROProperty("outerhtml"),"v-disabled")
									If Not bNextPageExist Then
										objLnk_Next.RefreshObject
										objLnk_Next.WebElement("class:=v-button-caption").click	
										WaitForICallLoading
										intCurrentPageNumber = GetPageNumber(objPageNumber)
									Else
										If (intCurrentPageNumber =   strPageNumber) Then
											bGoToExpectedRow = True
											Exit Do
										Else
											bGoToExpectedRow = False
											Exit Do
										End If
									End If
						Loop

						If  (intCurrentPageNumber <   strPageNumber)  Then
							LogMessage "WARN","Verification","The Next link is disabled before reaching the expected page - pls verify the data, Expected Page: "& strPageNumber &" Actual Page: "&intCurrentPageNumber, True	
						End If           
												
						If bGoToExpectedRow = True Then
							LogMessage "RSLT","Verification","Reached the expected Page :  "& strPageNumber, True	
							GoToExpectedRow = bGoToExpectedRow
						Else
							LogMessage "WARN","Verification","Next Link disabled and less number of records are shown than expected", False
							GoToExpectedRow = bGoToExpectedRow
						End If
				Else
					LogMessage "WARN","Verification", "Next Link was not available in the screen or property changed", False
					GoToExpectedRow = False						   
				End If     
		End Function
		
		Public Function GoToLastRecord(objLnk_Next)
				If objLnk_Next.Exist(0) Then
						Dim bNextPageExist, intPageCount
						Do 
							objLnk_Next.RefreshObject
							bNextPageExist =matchStr(objLnk_Next.GetROProperty("outerhtml"),"v-disabled")	
									If Not bNextPageExist Then
										objLnk_Next.WebElement("class:=v-button-caption").click	 
										WaitForICallLoading
										intPageCount = intPageCount + 1
									End If
						Loop while Not  ( bNextPageExist Or  intPageCount > 100)  'To avoid looping
						If  intPageCount > 100 Then
							Print "Next Link was clicked more than 100 times"
							GoToLastRecord = False					
						Else
							Print "Reached the Last page"
							GoToLastRecord = True
						End If            
				Else
					Print "Next Link was not available in the screen or property changed"
		'			LogMessage "WARN","Verification", "Next Link was not available in the screen or property changed", False
					GoToLastRecord = False			   
				End If
		End Function

End Class
