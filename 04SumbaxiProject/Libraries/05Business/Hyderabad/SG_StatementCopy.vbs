'[Click on Button Statement Copy for statement page]
Public Function clickbtnStatementCopy()
	clickbtnStatementCopy = True
	For iCounti = 1 To 180 Step 1
		If Not bcSTPStatementCopy.btnStatementCopy.Exist(0.5) Then
			Wait(0.5)
		else
			 bcSTPStatementCopy.btnStatementCopy.click
			Exit for
		End if
	Next 
   
   If Err.Number<>0 Then
		clickbtnStatementCopy = false
		LogMessage "WARN","Verification","Failed to Click button Statement Copy" ,false
		Exit Function
   	Else 
       LogMessage "RSLT","Verification","The button Statement Copy is clicked successfully",true
   End If
   WaitForIcallLoading
End Function

'[Verify Validation Message displayed on Statement Copy as]
Public Function verifyValidationMessage_StatmntCopy(strValidationMsgStmntCopy)
	verifyValidationMessage_StatmntCopy = True
	For iCountj = 1 To 180 Step 1
		If Not bcSTPStatementCopy.popupPreValidateStatementCopy.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(strValidationMsgStmntCopy) Then
				If Not VerifyInnerText (bcSTPStatementCopy.lblPopUpPreValidContent(), strValidationMsgStmntCopy, "Validation Message")Then
					verifyValidationMessage_StatmntCopy = false
				End If
			End If
			Exit for
		End if
	Next
	
	For icountl = 1 To 180 Step 1
		If Not bcSTPStatementCopy.btnPopupOkbtn.Exist(0.5) Then
			Wait(0.5)
		else
			bcSTPStatementCopy.btnPopupOkbtn.Click
			Exit for
		End If
	Next
	
	If Err.Number<>0 Then
		verifyValidationMessage_StatmntCopy = false
		LogMessage "WARN","Verification","Failed to Click button Ok" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","The button Ok is clicked successfully",true
	End If
	WaitForIcallLoading
End Function

'[Verify row Data in Table SelectedCards for Statement Copy STP]
Public Function verifytblSelectedCardsContent_StatmntCopy(arrRowDataList)
   bDevPending=false
   verifytblSelectedCardsContent_StatmntCopy = True
   For icountm = 1 To 180 Step 1
		If Not bcSTPStatementCopy.tblSelectedCardsHeader.Exist(0.5) Then
			Wait(0.5)
		else
			wait(8)
			verifytblSelectedCardsContent_StatmntCopy = verifyTableContentList(bcSTPStatementCopy.tblSelectedCardsHeader,bcSTPStatementCopy.tblSelectedCardsContent,arrRowDataList,"SelectedCardsContent" , false,null ,null,null)
			Exit for
		End If
	Next
	
	If Err.Number<>0 Then
		verifytblSelectedCardsContent_StatmntCopy = false
		LogMessage "WARN","Verification","Failed to Navigate to Statement Copy Page" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","Successfully Navigate to Statement Copy Page",true
	End If
	WaitForIcallLoading  
End Function

'[Get AddressLine1 Label Text on Statement Copy]
Public Function getAddressLine1Text_SC()
   bDevPending = false
   getAddressLine1Text_SC = bcSTPStatementCopy.lblAddressLine1.GetRoProperty("innertext")
End Function

'[Verify Field Address Line1 on Statement Copy displayed as]
Public Function verifyAddressLine1Text_SC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine1Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblAddressLine1(), strExpectedText, "AddressLine1")Then
           bVerifyAddressLine1Text=false
       End If
   End If
   verifyAddressLine1Text_SC=bVerifyAddressLine1Text
End Function

'[Get Address Line2 Label Text on Statement Copy]
Public Function getAddressLine2Text_SC()
   bDevPending=false
   getAddressLine2Text_SC=bcSTPStatementCopy.lblAddressLine2.GetRoProperty("innertext")
End Function

'[Verify Field Address Line2 on Statement Copy displayed as]
Public Function verifyAddressLine2Text_sC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine2Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblAddressLine2(), strExpectedText, "AddressLine2")Then
           bVerifyAddressLine2Text=false
       End If
   End If
   verifyAddressLine2Text_sC=bVerifyAddressLine2Text
End Function

'[Get AddressLine3 Label Text on Statement Copy]
Public Function getAddressLine3Text_SC()
   bDevPending=false
   getAddressLine3Text_SC=bcSTPStatementCopy.lblAddressLine3.GetRoProperty("innertext")
End Function

'[Verify Field Address Line3 on Statement Copy displayed as]
Public Function verifyAddressLine3Text_SC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine3Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblAddressLine3(), strExpectedText, "AddressLine3")Then
           bVerifyAddressLine3Text=false
       End If
   End If
   verifyAddressLine3Text_SC=bVerifyAddressLine3Text
End Function

'[Get AddressLine4 Label Text on Statement Copy]
Public Function getAddressLine4Text_SC()
   bDevPending=false
   getAddressLine4Text_SC=bcSTPStatementCopy.lblAddressLine4.GetRoProperty("innertext")
End Function

'[Verify Field Address Line4 on Statement Copy displayed as]
Public Function verifyAddressLine4Text_SC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine4Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblAddressLine4(), strExpectedText, "AddressLine4")Then
           bVerifyAddressLine4Text=false
       End If
   End If
   verifyAddressLine4Text_SC=bVerifyAddressLine4Text
End Function

'[Get AddressLine5 Label Text on Statement Copy]
Public Function getAddressLine5Text_SC()
   bDevPending=false
   getAddressLine5Text_SC=bcSTPStatementCopy.lblAddressLine5.GetRoProperty("innertext")
End Function

'[Verify Field Address Line5 on Statement Copy displayed as]
Public Function verifyAddressLine5Text_SC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine5Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblAddressLine5(), strExpectedText, "AddressLine5")Then
           bVerifyAddressLine5Text=false
       End If
   End If
   verifyAddressLine5Text_SC=bVerifyAddressLine5Text
End Function

'[Get PostalCode Label Text on Statement Copy screen]
Public Function getPostalCodeText_SC()
   bDevPending=false
   getPostalCodeText_SC=bcSTPStatementCopy.lblPostalCode.GetRoProperty("innertext")
End Function

'[Verify Field PostalCode on Satement Copy displayed as]
Public Function verifyPostalCodeText_SC(strExpectedText)
   bDevPending=false
   bVerifyPostalCodeText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblPostalCode(), strExpectedText, "PostalCode")Then
           bVerifyPostalCodeText=false
       End If
   End If
   verifyPostalCodeText_SC=bVerifyPostalCodeText
End Function


'[Verify Field eStatement Enrollment Status on Statement Copy displayed as]
Public Function verifyeStatementEnrollmentStatus(strExpectedText)
   bDevPending=false
   bVerifyStatementEnrollmentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lbleStatementEnrollmentStatus(), strExpectedText, "eStatement Enrollment Status")Then
           bVerifyStatementEnrollmentText=false
       End If
   End If
   verifyeStatementEnrollmentStatus=bVerifyStatementEnrollmentText
End Function

'[Verify Button Submit is enabled on Statement Copy SR Screen]
Public Function VerifybtnSubmit_SmtCopy(bEnabled)
	bDevPending=False
    bVerifybtnSubmit_SmtCopy=true
	intBtnSubmit=Instr(bcSTPStatementCopy.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
	If bEnabled Then
		If  intBtnSubmit=0 Then
			LogMessage "RSLT","Verification","Submit button is enable as per expectation.",True
			bVerifybtnSubmit_SmtCopy=true
		Else
			LogMessage "WARN","Verifiation","Submit button is disable. Expected to be enable.",false
			bVerifybtnSubmit_SmtCopy=false
		End If
	else
		If  intBtnSubmit<>0 Then
			LogMessage "RSLT","Verification","Submit button is disabled as per expectation.",True
			bVerifybtnSubmit_SmtCopy=true
		Else
			LogMessage "WARN","Verifiation","Submit button is Enabled. Expected to be disabled.",false
			bVerifybtnSubmit_SmtCopy=false
		End If
	End If
	VerifybtnSubmit_SmtCopy=bVerifybtnSubmit_SmtCopy
End Function

'[Select Archived Statements Duration in Statement Copy SR Screen]
Public Function selectArchivedStatementsDuration(strFromDate,strToDate)
   selectArchivedStatementsDuration = true
	If Not IsNull (strFromDate) Then
		If strFromDate = "RUNTIME" Then
			strFromDate=Date
			strFromDate=dateAdd("d",1,strFromDate)	
			'If len(Day(CDate(strFromDate)))=1 Then
				'strDay="0"&Day(CDate(strFromDate))
			'else
				strDay=""&Day(CDate(strFromDate))
			'End If
			strFromDate=""&strDay & " "&monthName(Month(CDate(strFromDate)),true) &" " &Year(CDate(strFromDate))&""
		End If
		selectArchivedStatementsDuration = SelectDatePicker_FromDate(strFromDate)
		'bcSTPStatementCopy.txtArchivedStatementsFromDate.set strFromDate
	End If
	If Not IsNull (strToDate) Then
		If strToDate = "RUNTIME" Then
			strToDate=Date
			strToDate=dateAdd("d",2,strToDate)	
			'If len(Day(CDate(strToDate)))=1 Then
				'strDay="0"&Day(CDate(strToDate))
			'else
				strDay=""&Day(CDate(strToDate))
			'End If
			strToDate=""&strDay & " "&monthName(Month(CDate(strToDate)),true) &" " &Year(CDate(strToDate))&""
		End If
		selectArchivedStatementsDuration = SelectDatePicker_TODate(strToDate)
		'bcSTPStatementCopy.txtArchivedStatementsToDate.set strToDate
	End If	
	If Err.Number<>0 Then
       selectArchivedStatementsDuration = false
       LogMessage "WARN","Verification","Failed to enter Archived Statements Duration" ,false
       Exit Function
   End If
End Function

'[Verify Field Inline Message for Error displayed on Statement Copy Screen as]
Public Function verifyInlineMessageArchStm_StmCopy(strExpectedText)
   verifyInlineMessageArchStm_StmCopy = true
   wait(1)
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblInlineMessage_Error(), strExpectedText, "Archived Statements")Then
           verifyInlineMessageArchStm_StmCopy = false
       End If
   End If
End Function

'[Select the Requested Dates based on the Available Statements on Statement Copy Screen]
Public Function selectAvailableStatementCopy(lstAvilSmtDates)
	selectAvailableStatementCopy = True
	If Not IsNull(lstAvilSmtDates) Then	
		'split the columns and values
		intSize = Ubound(lstAvilSmtDates) + 1
		ReDim arrCol(intSize)
		ReDim arrVal(intSize)
		For iterator = 0 To Ubound(lstAvilSmtDates)  Step 1
			tempVal = split(lstAvilSmtDates(iterator),":")
			arrCol(iterator) = tempVal(0)
			arrVal(iterator) = tempVal(1)
		Next	
		'Fetch the total no of rows
		Set objAllRows=getAllRows(bcSTPStatementCopy.tblRadioButAvailableStatements)
		intRow=objAllRows.Count  
		For colLoop = 0 To intSize-1 Step 1
			'Select the checkbox only if arrval() is True
			strColName = arrCol(colLoop)
			If arrval(colLoop) = "True" Then
				selectChk = SelectRadioButtonGrp(strColName,bcSTPStatementCopy.tblRadioButAvailableStatements,"Date Selection")
			End If		
		Next
	End If
	
	If Err.Number<>0 Then
		selectAvailableStatementCopy = false
		LogMessage "WARN","Verification","Failed to Select Available Statements Check Box" ,false
		Exit Function
	End If
End Function

'[Verify Field Fee displayed as]
Public Function verifydFeeDisplayed_StmCopy(strExpectedText)
   verifydFeeDisplayed_StmCopy = true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblsmtCopyFeeDisplayed(), strExpectedText, "Statement Copy Fee")Then
           verifydFeeDisplayed_StmCopy = false
       End If
   End If
End Function

'[Select statement Waive Fee in Statement Copy Screen as]
Public Function selectStatementWaiveFee(strWaiveFee)
	bDevPending=false
	selectStatementWaiveFee = true
	If Not IsNull (strWaiveFee) Then
		selectStatementWaiveFee = SelectRadioButtonGrp(strWaiveFee,bcSTPStatementCopy.rbtnWaiveFee, "Yes or No")
	End If
	If Err.Number<>0 Then
       selectStatementWaiveFee = false
       LogMessage "WARN","Verification","Failed to Select Statement Waive Fee: Yes or No" ,false
       Exit Function
   End If
End Function

'[Select the Account or Card to be Charged as]
Public Function selectAccOrCardToBeCharged(strAccountOrCard)
	selectAccOrCardToBeCharged = true
	If Not IsNull(strAccountOrCard) Then
       If Not (selectItem_Combobox (bcSTPStatementCopy.lstAccountCardCharged(), strAccountOrCard))Then
            LogMessage "WARN","Verification","Failed to select :"&strAccountOrCard&" From Account/Card to be Charged drop down list" ,false
           selectAccOrCardToBeCharged = false
       End If
   End If
   WaitForICallLoading
End Function

'[Verify Field Description on Statement Copy displayed as]
Public Function verifydDescriptionDisplayed_StmCopy(strExpectedText)
   verifydDescriptionDisplayed_StmCopy = true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblsmtCopyDescriptionDisplayed(), strExpectedText, "Statement Copy Fee")Then
           verifydDescriptionDisplayed_StmCopy = false
       End If
   End If
End Function

'[Verify Knowledge base link is enabled on Statement Copy SR Screen]
Public Function VerifyKnowledgebaselinkEnabled_SC()
    bDevPending=false
   	Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
     strKBLink=bcSTPStatementCopy.lnkKnowledgeBase.GetROProperty("Outerhtml")	
    If inStr(strKBLink,"v-disabled") = 0 Then
		LogMessage "RSLT","Verification","Knowledge base Link  enabled successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not enabledas expected",false
		bVerifyKnowledgebaselink=false
	End If
	VerifyKnowledgebaselinkEnabled_SC=bVerifyKnowledgebaselink
End Function

'[Verify Field KnowledgeBase on Statement Copy SR Screen displayed as]
Public Function verifyKnowledgeBase_SC(strExpectedLink)
   bDevPending=false
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strExpectedLink) Then
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"
			strKBLink=bcSTPStatementCopy.lnkKnowledgeBase.GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_SC=bVerifyKnowledgeBaseText
End Function

'[Perform Add Notes by clicking Add Notes Button on Statement Copy SR Screen]
Public Function addNote_SC(strNote)
   bDevPending = false
   bVerifypopupNotes = true
	Dim bVerifypopupNotes:VerifypopupNotes=true	
	If not isNull(strNote) Then
		bcSTPStatementCopy.btnAddNotes.click
		WaitForICallLoading
            If not ServiceRequest.popupVerification.exist(5)Then
				LogMessage "WARN","Verification","New Note dialog did not displayed",false
				bVerifypopupNotes=false
			 else
			 strMessage=ServiceRequest.lblMaxAllowed.GetROProperty("innerText")
				If not strMessage="Max allowed - 3000" Then
					LogMessage "WARN","Verification","Add New Comment popup dialog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage,false
					bVerifypopupNotes=false
				End If
			   	ServiceRequest.txtNewComment.set strNote			  
				ServiceRequest.clickSave_Popup
			  	WaitForIcallLoading
		   End If 
		End If 
	addNote_SC = bVerifypopupNotes
End Function

'[Set TextBox Comment on Statement Copy SR Screen to]
Public Function setCommentTextbox_SC(strComment)
	bDevPending=False
	strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	'gstrRuntimeCommentStep="Set TextBox Comment on Statement Copy SR Screen to"
	'insertDataStore "SRComment", strComment
	
	'Insert the TimeStamp in the datastore
	'strTimeStamp = convertDateTime_WithoutSec(Now)
	gstrRuntimeCommentStep="Set TextBox Comment on Statement Copy SR Screen to"
	gstrParameterNameStep = "TimeStamp"&replace((replace((replace(now,"/","-"))," ","-")),":","-")
	insertDataStore gstrParameterNameStep, strComment
	bcSTPStatementCopy.txtComment.Set(strComment )
	If Err.Number<>0 Then
		setCommentTextbox_SC = false
		LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
		Exit Function
	End If
	setCommentTextbox_SC = true
End Function

'[Click Button Submit on Statement Copy SR Screen]
Public Function clickButtonSubmit_SC()
   bDevPending=False
   bcSTPStatementCopy.btnSubmit.click
   
   'Capturing time stamp to open Memo for this SR
	strRunTimeTimeStamp_Step="Click Button Submit on Statement Copy SR Screen"
	
		StrDateFormat  = FormatDateTime(now, 1)
        StrDateFormat1  = Split (Trim(StrDateFormat),",",-1,1)
        StrDt  = Right(Trim(StrDateFormat1(1)),2)
        StrMonth  = Left(Trim(StrDateFormat1(1)),3)
        strDate  = StrDt&" "&StrMonth&" "&Trim(StrDateFormat1(2))
        
	strTempTime=FormatDateTime(now,4)	
 	strTimeStamp=strDate&" "&strTempTime
	insertDataStore "TimeStamp", strTimeStamp
	
	If Err.Number<>0 Then
		clickButtonSubmit_SC = false
		LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
		Exit Function
	End If
   WaitForICallLoading
   clickButtonSubmit_SC = true
End Function

'[Verify Popup Request Submitted exist for Statement Copy]
Public Function verifyPopupRequestSubmitted_SmtCopy(bExist)
	bDevPending=false
	WaitForICallLoading
	For icounty = 1 To 180 Step 1
	    If Not bcSTPStatementCopy.popupRequestSubmitted.Exist(0.5) Then
		   Wait(0.5)
		else
		    WaitForICallLoading
		    bActualExist=bcSTPStatementCopy.popupRequestSubmitted.Exist(4)
		    Exit for
		End If
	Next
	
	If Err.Number<>0 Then
		verifyPopupRequestSubmitted_SmtCopy = false
		LogMessage "WARN","Verification","Failed to Submit Statement Copy SR" ,false
		Exit Function
	End If

	If bExist And  bActualExist  Then
		LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
		verifyPopupRequestSubmitted_SmtCopy=True
	ElseIf not bExist And  not bActualExist  Then
		LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
		verifyPopupRequestSubmitted_SmtCopy=True
	ElseIf bExist And  not bActualExist  Then
		LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
		verifyPopupRequestSubmitted_SmtCopy=False
	ElseIf not bExist And   bActualExist  Then
		LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
		verifyPopupRequestSubmitted_SmtCopy=False
	End If
End Function

'[Verify Field CardNumber on Request Submitted Popup for Statement Copy displayed as]
Public Function verifyCardNumber_RequestSubmitted_SmtCopy(strCardNumber)
   bDevPending=false
   bverifyCardNumber_RequestSubmitted=true
   WaitForICallLoading
   insertDataStore "NewSCUsedCard", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       WaitForICallLoading
       If Not VerifyInnerText (bcSTPStatementCopy.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber_RequestSubmittedText=false
       End If
   End If
   verifyCardNumber_RequestSubmitted_SmtCopy=bVerifyCardNumber_RequestSubmittedText
End Function

'[Verify Field ProductDescription on Request Submitted Popup for Statement Copy displayed as]
Public Function verifyProductDescription_RequestSubmitted_SmtCopy(strProductDescription)
   bDevPending=false
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (bcSTPStatementCopy.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_SmtCopy=bVerifyProductDescription_RequestSubmittedText
End Function

'[Click Close button on Request Submitted Popup for Statement Copy]
Public Function verifybtnClose_RequestSubmitted_SmtCopy()
    bverifybtnClose_RequestSubmitted_CBR=true
    WaitForIcallLoading
    bcSTPStatementCopy.btnCancel_RequestSubmitted.click
    If Err.Number<>0 Then
       bverifybtnClose_RequestSubmitted_CBR=false
       LogMessage "WARN","Verification","Failed to Click Close Button : Yes on Confirmation popup" ,false
       Exit Function
    End If
    WaitForICallLoading
    verifybtnClose_RequestSubmitted_SmtCopy=bverifybtnClose_RequestSubmitted_CBR
End Function

'[Verify Field TM Approval Message on Statement Copy Screen displayed as]
Public Function verifyTMApprovalMessage_SC(strValidationMessage)
   bDevPending=False
   bverifyTMApprovalMessage_SC=true
   WaitForIcallLoading
   If Not IsNull(strValidationMessage) Then
       WaitForIcallLoading
       If Not VerifyInnerText (bcSTPStatementCopy.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bverifyTMApprovalMessage_SC=false
       End If
   End If
   bcSTPStatementCopy.btnOK_ValidationPopup.Click
   WaitForIcallLoading
   verifyTMApprovalMessage_SC=bverifyTMApprovalMessage_SC
End Function
