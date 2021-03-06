Dim gWaitTime:gWaitTime = 10
'###############################################################################################
'# Name: VerifyMaxLength()
'# Description: Function to verify an object is clickable/Not and then write the result to log file.
'# Author: 
'# Date: 17-November-2016
'# Input Parameters: oCurrentObj,intMaxLengthExpected,strObjectName
'#	oCurrentObj = Object path
'# 	intMaxLengthExpected = Expected max length of the edit box
'# 	strObjectName = Maker comment box /Checker comment box  etc...
'# Output Parameters: True/False
'###############################################################################################

Public Function VerifyMaxLength(oCurrentObj,intMaxLengthExpected,strObjectName) 
	bverifyMaxLength=True
    If oCurrentObj.Exist(gWaitTime) Then
        intMaxLengthActual = Trim(oCurrentObj.GetROProperty("max length"))
        If intMaxLengthActual = Trim(intMaxLengthExpected) Then
            LogMessage "RSLT","Verification",strObjectName & "Max character limit is as expected: " & intMaxLengthActual ,True     
        Else      
        	LogMessage "WARN","Verification",strObjectName & "Max character limit is not as expected, " & "Actual is " & intMaxLengthActual & "Expected is " & intMaxLenthExpected ,False
        	bverifyMaxLength=False
        End If
    Else
        LogMessage "WARN","Step Failure","Expected object " & strObjectName &" does not exist in current Webpage.",False
		bverifyMaxLength=False        
    End If
    
    VerifyMaxLength = bverifyMaxLength
End Function

'###############################################################################################
'# Name: SetValue()
'# Description: Function to set text on object and verify the same
'# Author: 
'# Date: 17-November-2016
'# Input Parameters: oCurrentObj,intMaxLengthExpected,strObjectName
'#	oCurrentObj = Object path
'# 	strTextToSet = text to set in the object
'# 	strObjectName = Maker comment box /Checker comment box  etc...
'# Output Parameters: True/False
'###############################################################################################

Public Function SetValue(oCurrentObj,strTextToSet,strObjectName)

	blnsetTextEditField = True
	
	If Not IsNull(strTextToSet)  Then		
		If oCurrentObj.Exist(gWaitTime) Then
	        oCurrentObj.Init
	        oCurrentObj.Set Trim(Cstr(strTextToSet))
	        strTextSet = oCurrentObj.getROProperty("value")
	        
	        strTextToSet = Replace(strTextToSet,",","")
	        strTextSet = Replace(strTextSet,",","")
	        
	        If Trim(strTextToSet) = Trim(strTextSet) Then
	            LogMessage "RSLT","Verification",strObjectName & " Field : Set the text as expected: " & strTextSet ,True 
	        Else      
	        	LogMessage "WARN","Verification",strObjectName & " Field : Failed to set the text , " & "Actual is: " & strTextSet & ",Expected is: " & strTextToSet ,False
	        	blnsetTextEditField = False
	        End If
	        
	    Else
	        LogMessage "WARN","Step Failure","Expected object " & strObjectName & ", Field: does not exist in current Webpage.",False
			blnsetTextEditField = False  
	    End If
	Else
		LogMessage "WARN","Verification","Input Value is Null",False
	End If
	
	SetValue = blnsetTextEditField
	
	Set oCurrentObj = Nothing
End Function

'###############################################################################################
'# Name: SelectCheckBoxAndVerify()
'# Description: Function to select checkbox and verify the same
'# Author: 
'# Date: 17-November-2016
'# Input Parameters: oCurrentObj,strhtmlid,strObjectName
'#	oCurrentObj = Object path
'# 	strhtmlid = unique html id of the checkbox/webelement
'# 	strObjectName = Reporting text
'# Output Parameters: True/False
'###############################################################################################

Public Function SelectCheckBoxAndVerify(oCurrentObj,strhtmlid,strObjectName)
	bselectCheckBoxFromOptions=True
	Set chkBox = Description.Create
	chkBox("html id").value = strhtmlid
	Set chkBoxObject = oCurrentObj.childObjects(chkBox)
	For i = 0 To chkBoxObject.Count-1 Step 1
	
		chkBoxObject(i).Click				
		If instr(1,chkBoxObject(i).getROProperty("class"),"checked") <> "0" Then
			LogMessage "RSLT","Verification",strObjectName & " checkbox : checked successfully" ,True 
		Else
			LogMessage "WARN","Verification",strObjectName & " checkbox : Failed to select the checkbox",False
			bselectCheckBoxFromOptions=False
		End If 
	Next
	SelectCheckBoxAndVerify = bselectCheckBoxFromOptions
	Set oCurrentObj = Nothing
End Function

'###############################################################################################
'# Name: VerifyObjectEnabledDisabled()
'# Description: Function to verify an object is enable/disable state and then write the result to log file.
'# Author: 
'# Date: 17-November-2016
'# Input Parameters: strChildClass,strChildName,strCheckFlag,strObjectName
'# 		oCurrentObj = ObjectName in OR / pass the object properties through DP
'#						e.g: lnkPhoneBanking / "innertext:= Phone Banking"
'# 		strCheckFlag = Enable/Disable
'# 		strObjectName = Phone Banking link /Create Profile button /Delete Profile button etc....
'# Output Parameters: None
'# E.g.: VerifyObjectEnabledDisabled("WebElement","lnkPhoneBanking","Enable","Phone banking link")
'###############################################################################################
Public Function VerifyObjectEnabledDisabled(oCurrentObj,strCheckFlag,strObjectName)
	
	bverifyObjectEnabledDisabled=True
	
	If oCurrentObj.Exist(gWaitTime) Then
		If oCurrentObj.GetROProperty("disabled") = 0 Then
			bverifyObjectEnabledDisabled = True
		Else
			bverifyObjectEnabledDisabled = False
		End If
		
		Select Case strCheckFlag
			Case "Enable"
				If bverifyObjectEnabledDisabled Then
					LogMessage "RSLT","Verification",strObjectName & " is enabled as expected.",True
					bverifyObjectEnabledDisabled = True
				Else
					LogMessage "WARN","Verification",strObjectName & " is disabled. Expected enabled.",False
					bverifyObjectEnabledDisabled = False
				End If
			Case "Disable"
				If Not(bverifyObjectEnabledDisabled) Then
					bverifyObjectEnabledDisabled = True
					LogMessage "RSLT","Verification",strObjectName & " is disabled as expected.",True
				Else
					LogMessage "WARN","Verification",strObjectName & " is enabled. Expected disabled.",False
					bverifyObjectEnabledDisabled = False
				End If
		End Select
	Else
		LogMessage "WARN","Step Failure","Expected object " & strObjectName &" does not exist in current Webpage.",False
		bverifyObjectEnabledDisabled = False
	End If
	
	Set oCurrentObj = Nothing
	VerifyObjectEnabledDisabled = bverifyObjectEnabledDisabled
	
End Function
'###############################################################################################
'# Name: VerifyObjectEnabledDisabled()
'# Description: Function to verify an object is enable/disable state and then write the result to log file.
'# Author: 
'# Date: 19-October-2017
'# Input Parameters: strchildobject,strCheckFlag,strObjectName
'# 		oCurrentObj = ObjectName in OR / pass the object properties through DP
'#						e.g: lnkPhoneBanking / "innertext:= Phone Banking"
'# 		strCheckFlag = Enable/Disable
'# 		strObjectName = Submit biutton 
'# Output Parameters: None
'# E.g.: VerifyObjectEnabledDisabled("WebElement","lnkPhoneBanking","Enable","Phone banking link")
'###############################################################################################
Public Function VerifyObjectDisabled(oCurrentObj,strCheckFlag,strObjectName)	
	bVerifyObjectDisabled = True	
	If oCurrentObj.Exist(gWaitTime) Then
		If Instr(oCurrentObj.GetROProperty("Outerhtml"),"disabled") <> 0 Then
			bSetObjdisabled = True
		Else
			bSetObjdisabled = False
		End If		
		Select Case strCheckFlag
			Case "Enable"
				If bSetObjdisabled Then
					LogMessage "WARN","Verification",strObjectName & " is disabled. Expected enabled.",False
					bVerifyObjectDisabled = False
				Else
					LogMessage "RSLT","Verification",strObjectName & " is enabled as expected.",True
					bVerifyObjectDisabled = True
				End If
			Case "Disable"
				If bSetObjdisabled Then
					LogMessage "RSLT","Verification",strObjectName & " is disabled as expected.",True
					bVerifyObjectDisabled = True
				Else
					LogMessage "WARN","Verification",strObjectName & " is enabled. Expected disabled.",False
					bVerifyObjectDisabled = False
				End If
		End Select
	Else
		LogMessage "WARN","Step Failure","Expected object " & strObjectName &" does not exist in current Webpage.",False
		bVerifyObjectDisabled = False
	End If
	
	Set oCurrentObj = Nothing
	VerifyObjectDisabled = bVerifyObjectDisabled	
End Function
'###############################################################################################
'# Name: ClickOnObject()
'# Description: Function to click on the object
'# Author: 
'# Date: 17-November-2016
'# Input Parameters: oCurrentObj,objChildClass,objName,strObjectName
'#	oCurrentObj = Object path
'#	objChildClass = WebElement/WebButton/Link
'# 	objName = Object name as per object repository
'# 	strObjectName = Reporting text
'# Output Parameters: True/False
'###############################################################################################

Function ClickOnObject(oCurrentObj,strObjectName)
	Dim  bgetCurrentObject
	bgetCurrentObject=True
	
	If oCurrentObj.Exist(gWaitTime) Then
		oCurrentObj.Click
		LogMessage "RSLT","Verification",strObjectName & " is clicked as expected",True
	Else
		bgetCurrentObject = False
		LogMessage "RSLT","Verification",strObjectName & " is not clicked as expected",False
	End If
	
	ClickOnObject = bgetCurrentObject
	Set oCurrentObj = Nothing
End Function



Public Function signIn(strUserType, strUserName, strPassword,strExpectedStatus,strMessage )
	WaitForBrowserReady
	If Not Browser("title:=I.*").Exist(1) Then
		If gstrExecutionFramework = "OBTAF" Then
			setBaseAppState()
		else
			setBaseAppStateBDT()
		End If
	End If
	
	Window("BrowserWin").Maximize
	
	If Not (bcCustomerSearch.pageExists()) AND Not (cUserProfile.pageExists()) Then
		If Not IsNull(strUserName) Then
			bcLoginIServe_Page.txtUserId().set strUserName
		End If
		
		If Not IsNull(strPassword) Then
			bcLoginIServe_Page.txtPassword().set strPassword
		End If
		bcLoginIServe_Page.btnLogin().Click
	End If 
	
	WaitForIServeLoading
	
	If Ucase(trim(strExpectedStatus)) = "ERROR"  Then

		gObjIServePage.WebElement("eleErrorMsg").WaitProperty "visible","True",15000
		strActualMessage = Browser("Browser_IServe").Page("Page_Login").WebElement("eleErrorMsg").GetROProperty("innertext")
		Wait(2)
		If  Trim(strMessage) = trim(strActualMessage) Then
			LogMessage "RSLT","Verification"," Error message is displayed. Expected :- "&strMessage&", Actual:- "&strActualMessage,True
			signIn = True
		Else
			LogMessage "WARN","Verification"," Error message is not displayed. Expected :- "&strMessage&", Actual:- "&strActualMessage,False
			signIn = False
		End If
	Else
		signIn = UpdateLoginStatus(strUserType)	
		
		If (bcIServeHome_Page.lblWelcome().WaitProperty("innertext","Welcome",30000)) Then
			LogMessage "RSLT","Verification","Successful Logon",True
			signIn = True
		Else
			LogMessage "WARN","Verification"," Not Successful Logon or page is loading for more than 30 seconds",False
			signIn = False  
		End If
						
	End If	
End Function

Public Function UpdateLoginStatus(strUserType)
	If  (bcIServeHome_Page.lblWelcome().WaitProperty("innertext","Welcome",30000)) Then
		LogMessage "RSLT","Verification","User " &strUserType& " Successfully able to login into Application.",True
		UpdateLoginStatus = true					
	else
		LogMessage "RSLT","Verification","User " &strUserType& " Successfully able to login into Application.",True
		UpdateLoginStatus = false							  
	End If
End Function

Public Function IServeLogout()
	btnLogout().Click
	If btnConfirmationYes().Exist(gWaitTime) Then
		btnConfirmationYes().Click
		Wait(2)
		LogMessage "RSLT","Verification"," Logout successful from I.Serve application.",True
		IServeLogout = True
	Else	
		LogMessage "WARN","Verification"," Failed to Logout from I.Serve application.",False
		IServeLogout = False
	End If
	CloseBrowsers
End Function
'###############################################################################################
'# Name: verifyCrossMarkAndCloseTab()
'# Description: Function to verify cross(x) mark icon & to close the opened tab.
'# Author: 
'# Date: 21-October-2016
'# Input Parameters: strTabName, needToWriteLog
'# 		strTabName = Phone Banking/Create Profile/Delete Profile etc....
'#		needToWriteLog = True/False
'# Output Parameters: None
'# e.g: verifyCrossMarkAndCloseTab "Phone Banking",True
'###############################################################################################
Public Function verifyCrossMarkAndCloseTab(strTabName,needToWriteLog)
	Dim intTabCount,strCrossIconXpath
	verifyCrossMarkAndCloseTab = False
	
	Set oDesc = Description.Create
	oDesc("xpath").Value = "(//*[@class='md-tabs-canvas'])[1]/md-pagination-wrapper/md-tab-item"
	Set objChild = gObjIServePage.ChildObjects(oDesc)
	intTabCount = objChild.Count
	
	strCrossIconXpath = "(//i[contains(@class,'lnr-cross tab-close-icon')])"
	
	For i = 1 To intTabCount-1
		strCurrentTab = Trim(objChild(i).GetROProperty("innertext"))
		If (strCurrentTab <> "") AND (Trim(strTabName) = strCurrentTab) Then
			If gObjIServePage.WebElement("xpath:=" & strCrossIconXpath &"["& i & "]").Exist(gWaitTime) Then
				gObjIServePage.WebElement("xpath:=" & strCrossIconXpath &"["& i & "]").click
				Wait(2)
				If Not(gObjIServePage.WebElement("xpath:=" & strCrossIconXpath &"["& i & "]").Exist(gWaitTime)) Then 
					verifyCrossMarkAndCloseTab = True
				    Exit For
			    End If
		    End If
		End If
	Next
	
	If needToWriteLog Then
		If verifyCrossMarkAndCloseTab Then
			LogMessage "Result","VERIFICATION"," On Clicking cross mark in "&strTabName&" page, it closed successfully." ,True     
		Else      	
			LogMessage "Warning","VERIFICATION"," On Clicking cross mark in "&strTabName&" page, it does not closed successfully.",False     
		End If
	End If
	
	Set oDesc = Nothing
	Set oCurrentObj = Nothing
End Function

'###############################################################################################
'# Name: VerifyFieldExistenceInPage()
'# Description: Function to verify an particular field is availble or not in a page.
'# Author: 
'# Date: 20-April-2017
'# Input Parameters: oFieldObj,strPageName,strFieldName
'# 		oFieldObj = Field object
'#		strPageName = Name of the page ,where we need to check the field
'#		strFieldName = Name of the Web field to verify
'# Output Parameters: None
'# e.g: VerifyFieldExistenceInPage lstPreferedLanguage "Account Details" "Prefered Language"
'###############################################################################################
Public Function VerifyFieldExistenceInPage(oFieldObj,strPageName,strFieldName)
	bgetCurrentObject = True
	If oFieldObj.Exist(5) Then
		LogMessage "RSLT","Verification",strFieldName & " is exists as expected in "&strPageName,True
	Else
		bgetCurrentObject = False
		LogMessage "WARN","Verification",strFieldName & " is not exists in "&strPageName,False
	End If
	VerifyFieldExistenceInPage = bgetCurrentObject
End Function

'###############################################################################################
'# Name: VerifyFieldRemovedFromPage()
'# Description: Function to verify an particular field is removed or not from a page.
'# Author: 
'# Date: 20-April-2017
'# Input Parameters: oFieldObj,strPageName,strFieldName
'# 		oFieldObj = Field object
'#		strPageName = Name of the page ,where we need to check the field
'#		strFieldName = Name of the Web field to verify
'# Output Parameters: None
'# e.g: VerifyFieldRemovedFromPage lstCustomerSegment "Account Details" "Customer Segment"
'###############################################################################################
Public Function VerifyFieldRemovedFromPage(oFieldObj,strPageName,strFieldName)
	bgetCurrentObject = True
	If Not oFieldObj.Exist(5) Then
		LogMessage "RSLT","Verification",strFieldName & " is not exists as expected in "&strPageName,True
	Else
		bgetCurrentObject = False
		LogMessage "WARN","Verification",strFieldName & " is exists in "&strPageName,False
	End If
	VerifyFieldRemovedFromPage = bgetCurrentObject
End Function
'###############################################################################################
'# Name: verifyLabelValuePairs()
'# Description: Function to verify label and value combinations
'# Author: 
'# Date: 12-April-2017
'# Input Parameters: strLabelValuePair
'# Output Parameters: None
'###############################################################################################
Public Function verifyLabelValuePairs(strLabelValuePair)
	Dim objDiv
	Dim blnIndicator:blnIndicator=true
	
	For Iterator2 = 0 To Ubound(strLabelValuePair) Step 1
		strValues1 = strLabelValuePair(Iterator2)
		
		Set objDiv=Description.Create
		objDiv("class").value="layout-padding layout-column flex"
		
		Set objDivChild = gObjIServePage.ChildObjects(objDiv)
		
		For k = 0 To objDivChild.Count-1 Step 1
			
			Set objLabel=Description.Create
			objLabel("class").value="ng-scope flex"
			objLabel("html tag").value="LABEL"
			Set obj = gObjIServePage.WebElement("xpath:=//div[@class='layout-padding layout-column flex']["&k+1&"]").ChildObjects(objLabel)
			
			Set objValue=Description.Create
			objValue("class").value="md-body-2 ng-binding.*"
			objValue("html tag").value="SPAN"
			Set obj1 = gObjIServePage.WebElement("xpath:=//div[@class='layout-padding layout-column flex']["&k+1&"]").ChildObjects(objValue)
			
			intCount = obj.Count			
			For i = 0 To intCount-1 Step 1
				strInnertextLable = Trim(obj(i).GetROProperty("innertext"))
				strInnertextValue = Trim(obj1(i).GetROProperty("innertext"))
				 
				strValue1 = strInnertextLable &":"&strInnertextValue
				strValues = strValues &"|"& strValue1
			Next
		Next
		
		If instr(1,strValues,strValues1,1)>0 Then
			LogMessage "RSLT","Verification","Text is displayed as expected." & strValues1,True
			blnIndicator=true
		Else
			LogMessage "WARN","Verifiation","Failed to display text as expected" & strValues1  ,false
			blnIndicator=false
			
		End If
	Next	
	verifyLabelValuePairs=blnIndicator
End Function

'[Verify the lists of the Dropdown box]
Public Function verifyDropdownListValues(txtComboObj,lstCombo,strDropdownName)
	verifyDropdownListValues = true
	
	If IsArray(lstCombo) Then
		ctList = ubound(lstCombo)
	Else
		ctList = 0
	End If	
	For it = 0 To ctList Step 1	
		If IsArray(lstCombo) Then
			strValue = lstCombo(it)
		Else
			strValue = lstCombo
		End If		
		'check if the list has a value
		If strValue <> "" Then
			'the move forward
			'Now use the for loop to set the value and after setting get it and compare
			txtComboObj.set strValue
			WaitForIServeLoading
			actComboList = txtComboObj.GetRoProperty("value")
			If strValue = actComboList Then
				LogMessage "RSLT", "Verification","The value "&strValue&" exists in the ("&strDropdownName&") Dropdown List",True
			else
				LogMessage "WARN", "Verification","The value "&strValue&" does not exist in ("&strDropdownName&") Dropdown List",False
				verifyDropdownListValues = false
			End If
		End If
	Next
End Function

Public Function VerifyDropdownDefaultValue(txtComboObj,strDefaultVal,strDropdownName)
	bComDefaultVal = False
	strActualVal = txtComboObj.GetRoProperty("value")
	If strDefaultVal = strActualVal Then
		LogMessage "RSLT", "Verification","As expected, Default value for ("&strDropdownName&") Dropdown is : "&strDefaultVal,True
		bComDefaultVal = True
	else
		LogMessage "WARN", "Verification","Failed,Default value for ("&strDropdownName&") Dropdown is incorrect, Expected : "&strDefaultVal&", Actual : "&strActualVal,False
	End If
	VerifyDropdownDefaultValue = bComDefaultVal
End Function

Public Function VerifyWebListDefaultValue(txtComboObj,strDefaultVal,strDropdownName)
	bComDefaultVal = False
	strActualVal = Trim(txtComboObj.GetRoProperty("outertext"))
	If strDefaultVal = strActualVal Then
		LogMessage "RSLT", "Verification","As expected, Default value for ("&strDropdownName&") Dropdown is : "&strDefaultVal,True
		bComDefaultVal = True
	else
		LogMessage "WARN", "Verification","Failed,Default value for ("&strDropdownName&") Dropdown is incorrect, Expected : "&strDefaultVal&", Actual : "&strActualVal,False
	End If
	VerifyWebListDefaultValue = bComDefaultVal
End Function



