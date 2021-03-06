'[Click the Amendment STP shortcut button from the enquiry screen]
Public Function clickShortcutButton_Amendment()
	bclickShortcutButton_Amendment = true
	BankAndEarnEnrollment.btnAmendment.click
	If Err.Number<>0 Then
       clickShortcutButton_Amendment=false
            LogMessage "RSLT","Verification","Failed to Click shortcut button : Amendment" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickShortcutButton_Amendment=true	
End Function

'[Verify the row data for the current bank and earn programs table displayed]
Public Function verifyrowdata_CurrBankEarnProg(arrRowDataList)
	bverifyrowdata_CurrBankEarnProg = true
	verifyrowdata_CurrBankEarnProg = verifyTableContentList(OLS_Amendment.tblCurrentBankEarnProg_header,OLS_Amendment.tblCurrentBankEarnProg_content,arrRowDataList,"Current Bank & Earn Program",false,Null,Null,Null)
	verifyrowdata_CurrBankEarnProg = bverifyrowdata_CurrBankEarnProg
End Function

'[Verify the row data for the settings table displayed]
Public Function verifyrowdata_Settings(arrRowDataList)
	bverifyrowdata_Settings = true
	verifyrowdata_Settings = verifyTableContentList(OLS_Amendment.tblSettings_header,OLS_Amendment.tblSettings_content,arrRowDataList,"Settings",false,Null,Null,Null)
	verifyrowdata_Settings = bverifyrowdata_CurrBankEarnProg
End Function

'[Select the value from combobox bank and earn program]
Public Function selectBankandEarnProgCombobox(strBankandEarnProg)
   'bDevPending=false
   bselectBankandEarnProgCombobox=true
   obj_enabled = OLS_Amendment.lstBankandEarnSummary.GetROProperty("disabled")
   If obj_enabled <>0 Then
   	If Not IsNull(strBankandEarnProg) Then
       If Not (selectItem_Combobox (OLS_Amendment.lstBankandEarnSummary(),strBankandEarnProg))Then
            LogMessage "RSLT","Verification","Failed to select :"&strBankandEarnProg&" From ApprovalLevel drop down list" ,false
           bselectBankandEarnProgCombobox=false
       End If
   End If
   WaitForIcallLoading
   else
    LogMessage "RSLT","Verification","The dropdown is disabled and the value cannot be selected", true
  End If
   selectBankandEarnProgCombobox=bselectBankandEarnProgCombobox
End Function

'[Select the value from combobox crediting Account]
Public Function selectCreditingAccCombobox(strCreditingAcc)
   bDevPending=false
   bselectCreditingAccCombobox=true
   If Not IsNull(strCreditingAcc) Then
       If Not (selectItem_Combobox (OLS_Amendment.lstCreditingAccount(),strCreditingAcc))Then
            LogMessage "RSLT","Verification","Failed to select :"&strCreditingAcc&" From ApprovalLevel drop down list" ,false
           bselectCreditingAccCombobox=false
       End If
   End If
   WaitForIcallLoading
   selectCreditingAccCombobox=bselectCreditingAccCombobox
End Function

'[Verify the inline message for crediting account displayed as]
Public Function verifyInlineMessage_CreAcc(strInlineMesg)
	bDevPending=false
   bverifyInlineMessage_CreAcc=true
   If Not IsNull(strInlineMesg) Then
       If Not VerifyInnerText (OLS_Amendment.lblCreditingAccInlineMessage(), strInlineMesg, "InLine Message")Then
           bverifyInlineMessage_CreAcc=false
       End If
   End If
   verifyInlineMessage_CreAcc=bverifyInlineMessage_CreAcc	
End Function

'[Verify the Description for Amendment displayed as]
Public Function verifyDesc_OLSAmend(strDesc)
	bDevPending=false
   bverifyDesc_OLSAmend=true
   If Not IsNull(strDesc) Then
       If Not VerifyInnerText (OLS_Amendment.lblDescription_OLSAmendment(), strDesc, "Description")Then
           bverifyDesc_OLSAmend=false
       End If
   End If
   verifyDesc_OLSAmend=bverifyDesc_OLSAmend	
End Function

'[Verify the click for link of non eligible accounts is enabled and popup is displayed]
Public Function verifylink_NonEligibleAcc()
	bverifylink_NonEligibleAcc = true
	enabled_Obj = OLS_Amendment.lnkNonEligibleAcc.GetROProperty("visible")
	If enabled_Obj = true  Then
	OLS_Amendment.lnkNonEligibleAcc.click	
		If OLS_Amendment.popupNonEligibleAccounts.exist Then
			LogMessage "RSLT","Verification","The link is enabled and clicked successfully",True
			else
			LogMessage "RSLT","Verification","The link is enabled but not clicked successfully",False
		End If
		'OLS_Amendment.btnOk_popupNonEligibleAcc.click
		verifylink_NonEligibleAcc = bverifylink_NonEligibleAcc
		else
		LogMessage "RSLT","Verification","The link is disabled and cannot be clicked",False
	End If	
End Function

'[Verify the row data for non eligible account table displayed as]
Public Function verifyrowdata_NonEligibleAcc(arrRowDataList)
	bverifyrowdata_NonEligibleAcc = true
	verifyrowdata_NonEligibleAcc = verifyTableContentList(OLS_Amendment.tblNonEligibleAccounts_header,OLS_Amendment.tblNonEligibleAccounts_content,arrRowDataList,"Non Eligible Account",false,Null,Null,Null)
	OLS_Amendment.btnOk_popupNonEligibleAcc.click
	WaitForIcallLoading
	verifyrowdata_NonEligibleAcc = bverifyrowdata_NonEligibleAcc
End Function

'[Perform Add Notes by clicking Add Notes Button on OLS Amendment screen]
Public Function addNote_OLSAmendment(strNote)
   bDevPending=false
   bVerifypopupNotes=true
	Dim bVerifypopupNotes:VerifypopupNotes=true
	
	If not isNull(strNote) Then
		OLS_Amendment.btnAddNotes_OLSAmend.click
		WaitForICallLoading
            If not OLS_Amendment.popupAddNotes_OLS.exist(5)Then
				LogMessage "WARN","Verification","New Note dialog did not displayed",false
				bVerifypopupNotes=false
			 else
			 strMessage=OLS_Amendment.lblMaxAllowed_OLSAmend.GetROProperty("innerText")
				If not strMessage="Max allowed - 3000" Then
					LogMessage "WARN","Verification","Add New Comment popup dislog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage,false
					bVerifypopupNotes=false
				End If
			   ServiceRequest.txtNewComment.set (strNote)
				   ServiceRequest.clickSave_Popup
			  WaitForIcallLoading
		   End If 
		End If 
	addNote_OLSAmendment=bVerifypopupNotes
End Function

'[Set TextBox Comments to OLS Amendment]
Public Function setCommentsTextbox_OLSAmend(strComment)
   bDevPending=false
   If not isNull(strComment) Then
	   OLS_Amendment.txtComments_OLSAmend.set strComment
   End If
   If Err.Number<>0 Then
       setCommentsTextbox_OLSAmend=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentsTextbox_OLSAmend=true
End Function

'[Verify the submit button is enabled for Amendment STP]
Public Function verifySubmitEnabled_Amendment()
	bverifySubmitEnabled_Amendment = true
	enabled_Obj = OLS_Amendment.btnSubmit_OLSAmend.GetROProperty("disabled")
	If enabled_Obj = 0 Then
 	LogMessage "RSLT","Verification","The button Save is enabled as expected",True
 	else
 	LogMessage "RSLT","Verification","The button Save is disabled",False
	 End If
	verifySubmitEnabled_Amendment = bverifySubmitEnabled_Amendment
End Function

'[Click Button Submit on OLS amendment]
Public Function clickButtonSubmit_OLSAmendment()
   bDevPending=false
   OLS_Amendment.btnSubmit_OLSAmend.click
   If Err.Number<>0 Then
       clickButtonSubmit_OLSAmendment=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   clickButtonSubmit_OLSAmendment=true
End Function

'[Verify the prevalidation popup for Amendment exist is true]
Public Function verifypopupPreVal_OLSAmendment(bExist)
	 bDevPending=False
   bActualExist=OLS_Amendment.popupPreVal_OLSAmend.Exist(1)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage Exists As Expected" ,true
       verifypopupPreVal_OLSAmendment=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage does not Exists As Expected" ,true
       verifypopupPreVal_OLSAmendment=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage does not Exists As Expected" ,False
       verifypopupPreVal_OLSAmendment=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage Still Exists" ,False
       verifypopupPreVal_OLSAmendment=False
   End If
End Function

'[Verify the pre validation message for OLS Amendment displayed as]
Public Function verifyPreVal_OLSAmendment(strExpectedText)
'bDevPending=False
   bverifyPreVal_OLSAmendment=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (OLS_Amendment.lblValMesg_OLS(), strExpectedText, "ValidationMessage")Then
           bverifyPreVal_OLSAmendment=false
       End If
       OLS_Amendment.btnOk_PreVal.click
   End If
   
   verifyPreVal_OLSAmendment=bverifyPreVal_OLSAmendment
End Function

'[Verify the mailing address for the crediting account selected]
Public Function verifymailingaddress_OLSAmendment(strAddressline1,strAddressline2,strAddressline3,strAddressline4)
	bverifymailingaddress_OLSAmendment = true
	
	'I.Serve address values displayed as 
	strIServeAddressline1 = OLS_Amendment.lblAddress1.GetROProperty("value")
	strIServeAddressline2 = OLS_Amendment.lblAddress2.GetROProperty("value")
	strIServeAddressline3 = OLS_Amendment.lblAddress3.GetROProperty("value")
	strIServeAddressline4 = OLS_Amendment.lblAddress4.GetROProperty("value")
	
	If strIServeAddressline1 =strAddressline1 Then
		LogMessage "RSLT","Verification","Mailing Adress :Address line1 is as Expected" ,true
		else
		LogMessage "RSLT","Verification","Mailing Adress :Address line1 is not as Expected" ,false
	End If
	
	If strIServeAddressline2 =strAddressline2 Then
		LogMessage "RSLT","Verification","Mailing Adress :Address line2 is as Expected" ,true
		else
		LogMessage "RSLT","Verification","Mailing Adress :Address line2 is not as Expected" ,false
	End If
	
	If strIServeAddressline3 =strAddressline3 Then
		LogMessage "RSLT","Verification","Mailing Adress :Address line3 is as Expected" ,true
		else
		LogMessage "RSLT","Verification","Mailing Adress :Address line3 is not as Expected" ,false
	End If
	
	If strIServeAddressline4 =strAddressline4 Then
		LogMessage "RSLT","Verification","Mailing Adress :Address line4 is as Expected" ,true
		else
		LogMessage "RSLT","Verification","Mailing Adress :Address line4 is not as Expected" ,false
	End If
	verifymailingaddress_OLSAmendment = bverifymailingaddress_OLSAmendment
End Function

'[Click on the Cancel button for OLS Amendment]
Public Function clickCancel_OlsAmendment()
	bclickCancel_OlsAmendment = true
	OLS_Amendment.btnCancel_OLSAmend.click
	If Err.Number<>0 Then
       clickCancel_OlsAmendment=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickCancel_OlsAmendment=bclickCancel_OlsAmendment
End Function

'[Click on the Submit button for OLS Amendment]
Public Function clickSubmit_OlsAmendment()
	bclickSubmit_OlsAmendment = true
	OLS_Amendment.btnSubmit_OLSAmend.click
	If Err.Number<>0 Then
       clickSubmit_OlsAmendment=false
            LogMessage "WARN","Verification","Failed to Click Button :Submit" ,false
       Exit Function
   End If
   clickSubmit_OlsAmendment=bclickSubmit_OlsAmendment
End Function

'[Verify the confirmation message on the popup displayed as]
Public Function verifyConfirmationMesg_OLS(strExpectedText)
	bverifyConfirmationMesg_OLS=true
   bActualExist=OLS_Amendment.popupConfirmation_OLS.Exist(3)
   If bActualExist Then

	   If Not IsNull(strExpectedText) Then
		   If Not VerifyInnerText(OLS_Amendment.lblConfirmationMesg_OLS(),strExpectedText, "Confirmation Message")Then
			   bverifyConfirmationMesg_OLS=false
		   End If
	   End If
		OLS_Amendment.btnYes_popupConfirmation.click
	else
		Logmessage "RSLT","Verification","OLS Amendment confirmation popup does not displayed",false
		bverifyConfirmationMesg_OLS=false
   End If
   verifyConfirmationMesg_OLS=bverifyConfirmationMesg_OLS
End Function

'[Click Button Close_RequestSubmitted on Amendment SR Screen]
Public Function clickCloseButton_RequestSubmitted_Amendment()
   bDevPending=False
   OLS_Amendment.btnClose_ReqSubmitted_Amendment.click
   If Err.Number<>0 Then
       clickCloseButton_RequestSubmitted_Amendment=false
            LogMessage "WARN","Verification","Failed to Click Button : Close_RequestSubmitted" ,false
       Exit Function
   End If
   clickCloseButton_RequestSubmitted_Amendment=true
End Function

'[Verify the additional info fields on view SR page for STP Amendment]
Public Function verifyfields_OLSAmendement_SR(lstAdditionalDetails)
	bverifyfields_OLSAmendement_SR = true
	intSize = Ubound(lstAdditionalDetails)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstAdditionalDetails(Iterator),":")(0))
		arrValue = trim(Split(lstAdditionalDetails(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "Bank - Earn Program"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblBankEarnProg_SR(), arrValue, "Bank - Earn Program")Then
				LogMessage "RSLT","Verification","Additional Details - Bank - Earn Program:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Crediting Account"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblCredAcc_SR(), arrValue, "Crediting Account")Then
				LogMessage "RSLT","Verification","Additional Details - Crediting Account:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Mailing Address"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblMailingAddr_SR(), arrValue, "Mailing Address")Then
				LogMessage "RSLT","Verification","Additional Details - Mailing Address:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Address Line 1"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblAddress1_SR(), arrValue, "Address Line 1")Then
				LogMessage "RSLT","Verification","Additional Details - Address Line 1:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Address Line 2"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblAddress2_SR(), arrValue, "Address Line 2")Then
				LogMessage "RSLT","Verification","Additional Details - Address Line 2:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	     Case "Address Line 3"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblAddress3_SR(), arrValue, "Address Line 3")Then
				LogMessage "RSLT","Verification","Additional Details - Address Line 3:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Address Line 4"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblAddress4_SR(), arrValue, "Address Line 4")Then
				LogMessage "RSLT","Verification","Additional Details - Address Line 4:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Postal Code"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblPostalCode_SR(), arrValue, "Postal Code")Then
				LogMessage "RSLT","Verification","Additional Details - Postal Code:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	     Case "Enrolment Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblEnrolmentDate_SR(), arrValue, "Enrolment Date")Then
				LogMessage "RSLT","Verification","Additional Details - Enrolment Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Current Crediting Account"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblCurrCredAcc_SR(), arrValue, "Current Crediting Account")Then
				LogMessage "RSLT","Verification","Additional Details - Current Crediting Account:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Email"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblEmail_SR(), arrValue, "Email")Then
				LogMessage "RSLT","Verification","Additional Details - Email:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    
	    Case "Statement Option"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (ViewSR.lblStatementOpt_SR(), arrValue, "Statement Option")Then
				LogMessage "RSLT","Verification","Additional Details - Statement Option:"&arrValue&" is not displayed as expected",false
				bverifyfields_OLSAmendement_SR=false
			End If
	    End If
	    End select
	Next
	verifyfields_OLSAmendement_SR = bverifyfields_OLSAmendement_SR
End Function
