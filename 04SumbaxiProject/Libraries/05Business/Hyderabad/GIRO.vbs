'*****This is auto generated code using code generator please Re-validate ****************

'[Verify row Data in Table SelectedCards for GIRO STP]
Public Function verifytblSelectedCardsContent_GIRO(arrRowDataList)
   bDevPending=false
   Wait(60)
   For iCounter1 = 1 To 180 Step 1
		If Not GIRO.tblSelectedCardsHeader.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
	
	For iCounter2 = 1 To 180 Step 1
		If Not GIRO.tblSelectedCardsContent.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
	verifytblSelectedCardsContent_GIRO=verifyTableContentList(GIRO.tblSelectedCardsHeader,GIRO.tblSelectedCardsContent,arrRowDataList,"SelectedCardsContent" , false,null ,null,null)
End Function

'[Verify row Data in Table Current Statement Cycle for GIRO STP]
Public Function verifytblCurrentStatementCycleContent_GIRO(arrRowDataList)
	Wait(10)
   bDevPending=false
   For iCounter3 = 1 To 180 Step 1
		If Not GIRO.tblCurrentStatementCycleHeader.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
	
	For iCounter4 = 1 To 180 Step 1
		If Not GIRO.tblCurrentStatementCycleContent.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
   verifytblCurrentStatementCycleContent_GIRO=verifyTableContentList(GIRO.tblCurrentStatementCycleHeader,GIRO.tblCurrentStatementCycleContent,arrRowDataList,"Current Statement Cycle" , false,null ,null,null)
End Function

'[Verify row Data in Table Selected GIRO Account for GIRO STP]
Public Function verifytblSelectedGIROAccountContent_GIRO(arrRowDataList)
   bDevPending=false
   For iCounter5 = 1 To 180 Step 1
		If Not GIRO.tblSelectedGIROAccountHeader.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
	
	For iCounter6 = 1 To 180 Step 1
		If Not GIRO.tblSelectedGIROAccountContent.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
   verifytblSelectedGIROAccountContent_GIRO=verifyTableContentList(GIRO.tblSelectedGIROAccountHeader,GIRO.tblSelectedGIROAccountContent,arrRowDataList,"SelectedGIROAccountContent" , false,null ,null,null)
End Function

'[Verify row Data in Table Existing GIRO Details for GIRO STP]
Public Function verifytblExistingGIRODetailsContent_GIRO(arrRowDataList)
   bDevPending=false
   For iCounter7 = 1 To 180 Step 1
		If Not GIRO.tblExistingGIRODetailsHeader.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
	
	For iCounter8 = 1 To 180 Step 1
		If Not GIRO.tblExistingGIRODetailsContent.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
   verifytblExistingGIRODetailsContent_GIRO=verifyTableContentList(GIRO.tblExistingGIRODetailsHeader,GIRO.tblExistingGIRODetailsContent,arrRowDataList,"ExistingGIRODetailsContent" , false,null ,null,null)
End Function

'[Verify row Data in Table GIRO Suspension Details for GIRO STP]
Public Function verifytblExistingGIROSuspensionDetailsContent_GIRO(arrRowDataList)
   bDevPending=false   
   If Not IsNull (arrRowDataList) Then
   For iCounter9 = 1 To 180 Step 1
		If Not GIRO.tblGIROSuspensionDetailsHeader.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
	
	For iCounter9a = 1 To 180 Step 1
		If Not GIRO.tblGIROSuspensionDetailsContent.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
   	verifytblExistingGIROSuspensionDetailsContent_GIRO=verifyTableContentList(GIRO.tblGIROSuspensionDetailsHeader,GIRO.tblGIROSuspensionDetailsContent,arrRowDataList,"GIROSuspensionDetails" , false,null ,null,null)
   End If
   verifytblExistingGIROSuspensionDetailsContent_GIRO=true
End Function

'[Select Radio Button of Payment on GIRO Screen]
Public Function selectPayment_GIRO(strPayment)
	bDevPending=false
	bselectPayment_GIRO=true
	If Not IsNull (strPayment) Then
		bselectPayment_GIRO=SelectRadioButtonGrp(strPayment,GIRO.rbtnPayment, Array("Full Payment","Minimum Payment"))
	End If
	If Err.Number<>0 Then
       bselectPayment_GIRO=false
       LogMessage "WARN","Verification","Failed to Click Button : Payment" ,false
       Exit Function
   End If
   selectPayment_GIRO=bselectPayment_GIRO
End Function

'[Verify Radio Button Payment for Cashline on GIRO Screen]
Public Function verifyPayment_GIRO()
	bDevPending=false
	bverifyPayment_GIRO=true
	If (GIRO.rbtnPayment.Exist()) Then		
       LogMessage "WARN","Verification","Payment radio button is available for Cashline. Expected to be disable." ,false
       bverifyPayment_GIRO=false
	End If	
   verifyPayment_GIRO=bverifyPayment_GIRO
End Function

'[Verify Combobox Debiting Account in GIRO Screen displayed as]
Public Function verifyDebitingAccount_Default(strDebitingAccount)
   bDevPending=false
   bverifyDebitingAccount_Default=true
   If Not IsNull(strDebitingAccount) Then
       If Not verifyComboSelectItem (GIRO.lstDebitingAccount(),strDebitingAccount, "Debiting Account")Then
           bverifyDebitingAccount_Default=false
       End If
   End If
   verifyDebitingAccount_Default=bverifyDebitingAccount_Default
End Function

'[Select Combobox Debiting Account in GIRO Screen]
Public Function selectDebitingAccount_GIRO(strDebitingAccount)
	bselectDebitingAccount_GIRO=true
	If Not IsNull(strDebitingAccount) Then
	For iCounteri = 1 To 180 Step 1
		If Not GIRO.lstDebitingAccount.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next	
   If Not (selectItem_Combobox (GIRO.lstDebitingAccount(), strDebitingAccount))Then
        LogMessage "WARN","Verification","Failed to select :"&strDebitingAccount&" From Debiting Account drop down list" ,false
       bselectDebitingAccount_GIRO=false
   End If
   End If
   WaitForICallLoading
   selectDebitingAccount_GIRO=bselectDebitingAccount_GIRO
End Function

'[Select Action Menu Maintenance from GIRO Details on GIRO Screen]
Public Function selectMaintenanceAction_GIRO(lstGIRODetails)
   bDevPending=False
   bselectMaintenanceAction_GIRO=true
 	With GIRO
		  bselectMaintenanceAction_GIRO= selectTableSubMenu(.tbGIRODetailsHeader,.tbGIRODetailsContent,lstGIRODetails,"GIRO Details","Actions",False,NULL,NULL,NULL,"Maintenance",bDisabled)
	End With
	If bDisabled Then
		LogMessage "RSLT", "Verification","Maintenance action menu is not enabled",false
		bselectMaintenanceAction_GIRO=false
	End If
	WaitForICallLoading
	Wait 1
    selectMaintenanceAction_GIRO=bselectMaintenanceAction_GIRO
End Function

'[Verify row Data in Table GIRO Details on GIRO Screen]
Public Function verifytblGIRODetails_GIRO(arrRowDataList)
   bDevPending=false 
   verifytblGIRODetails_GIRO=verifyTableContentList(GIRO.tbGIRODetailsHeader,GIRO.tbGIRODetailsContent,arrRowDataList,"GIRO Details" , false,null ,null,null)
End Function

'[Select Combobox Account in GIRO Screen]
Public Function selectAccount_GIRO(strAccount)
	bselectAccount_GIRO=true
	If Not IsNull(strAccount) Then
       If Not (selectItem_Combobox (GIRO.lstAccount_GIRO(), strAccount))Then
            LogMessage "WARN","Verification","Failed to select :"&strAccount&" From Account drop down list" ,false
           bselectAccount_GIRO=false
       End If
   End If
   WaitForICallLoading
   selectAccount_GIRO=bselectAccount_GIRO
End Function

'[Verify Request Type Combobox has Items]
Public Function verifyRequestTypeComboboxItems(lstItems)
   bDevPending=false
   bverifyRequestTypeComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (GIRO.lstRequestType(),lstItems, "request Type")Then
           bverifyRequestTypeComboboxItems=false
       End If
   End If
   verifyRequestTypeComboboxItems=bverifyRequestTypeComboboxItems
End Function

'[Select Combobox Request Type in GIRO Screen]
Public Function selectRequestType_GIRO(strRequestType)
	bselectRequestType_GIRO=true
	wait(5)
	If Not IsNull(strRequestType) Then
       If Not (selectItem_Combobox (GIRO.lstRequestType(), strRequestType))Then
            LogMessage "WARN","Verification","Failed to select :"&strRequestType&" From Request Type drop down list" ,false
           bselectRequestType_GIRO=false
       End If
   End If
   wait(20)
   WaitForICallLoading
   selectRequestType_GIRO=bselectRequestType_GIRO
End Function

'[Verify Field Validation Message For GIRO STP displayed as]
Public Function verifyValidationMessage_GIRO(strExpectedText)
   bDevPending=False
   bverifyValidationMessage_GIRO=true
   wait(5)
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (GIRO.lblConfirmationMessage(), strExpectedText, "Validation Message")Then
           bverifyValidationMessage_GIRO=false
       End If
   End If
   GIRO.btnOK_ConfirmationMsg.Click
   wait(13)
   WaitForICallLoading
   verifyValidationMessage_GIRO=bverifyValidationMessage_GIRO
End Function

'[Click on SR Shortcut Button GIRO Setup]
Public Function clickBtnGIROSetup()
	bclickBtnGIROSetup=true
	Wait(5)
	GIRO.btnGIROSetup.click
	wait(5)
	If Err.Number<>0 Then
       clickBtnGIROSetup=false
       LogMessage "WARN","Verification","Failed to Click Button : GIRO Setup" ,false
       Exit Function
   End If
    clickBtnGIROSetup=true
	WaitForICallLoading
End Function

'[Click on SR Shortcut Button GIRO Maintenance]
Public Function clickBtnGIROMaintenance()
	bclickBtnGIROMaintenance=true
	wait(5)
	GIRO.btnGIROMaintenance.click
	wait(10)
	If Err.Number<>0 Then
       clickBtnGIROMaintenance=false
       LogMessage "WARN","Verification","Failed to Click Button : GIRO Maintenance" ,false
       Exit Function
   End If
    clickBtnGIROMaintenance=true
	WaitForICallLoading
End Function

'[Click Button Cancel on GIRO SR Screen]
Public Function clickButtonCancel_GIRO()   
   GIRO.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_GIRO=false
       LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel_GIRO=true
End Function

'[Verify Confirmation Popup on GIRO SR Screen]
Public Function verifyConfirmationPopup_GIRO(strConfirmationMsg)
	bverifyConfirmationPopup_GIRO=true
	If Not IsNull (strConfirmationMsg) Then	
		If Not verifyInnerText(GIRO.lblConfirmationMessage(), strConfirmationMsg, "Confirmation Message") Then
			bverifyConfirmationPopup_GIRO=false
		End If
	End If
	GIRO.btnYes_Confirmation.click
	  If Err.Number<>0 Then
       bverifyConfirmationPopup_GIRO=false
            LogMessage "WARN","Verification","Failed to Click Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
	verifyConfirmationPopup_GIRO=bverifyConfirmationPopup_GIRO
End Function

'[Verify Button Submit is enabled on GIRO SR Screen]
Public Function VerifybtnSubmit_GIRO(bEnabled)
	bDevPending=False
   bVerifybtnSubmit_GIRO=true
   Wait(5)
	intBtnSubmit=Instr(GIRO.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
	If bEnabled Then
		If  intBtnSubmit=0 Then
			LogMessage "RSLT","Verification","Submit button is enable as per expectation.",True
			bVerifybtnSubmit_GIRO=true
		Else
			LogMessage "WARN","Verifiation","Submit button is disable. Expected to be enable.",false
			bVerifybtnSubmit_GIRO=false
		End If
	else
		If  intBtnSubmit<>0 Then
			LogMessage "RSLT","Verification","Submit button is disabled as per expectation.",True
			bVerifybtnSubmit_GIRO=true
		Else
			LogMessage "WARN","Verifiation","Submit button is Enabled. Expected to be disabled.",false
			bVerifybtnSubmit_GIRO=false
		End If
	End If
	VerifybtnSubmit_GIRO=bVerifybtnSubmit_GIRO
End Function

'[Click Button Submit on GIRO SR Screen]
Public Function clickButtonSubmit_GIRO()
   bDevPending=true
   GIRO.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_GIRO=false
       LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonSubmit_GIRO=true
End Function

'[Verify Field Inline Message displayed on GIRO Screen as]
Public Function verifyInlineMessageText_GIRO(strExpectedText)
   bverifyInlineMessageText_GIRO=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (GIRO.lblInlineMessage(), strExpectedText, "Inline Message")Then
           bverifyInlineMessageText_GIRO=false
       End If
   End If
   verifyInlineMessageText_GIRO=bverifyInlineMessageText_GIRO
End Function

'[Verify Field Inline Message for Error displayed on GIRO Screen as]
Public Function verifyInlineMessagePaymentLimit_GIRO(strExpectedText)
   bverifyInlineMessagePaymentLimit_GIRO=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (GIRO.lblInlineMessage_Error(), strExpectedText, "Payment Limit")Then
           bverifyInlineMessagePaymentLimit_GIRO=false
       End If
   End If
   verifyInlineMessagePaymentLimit_GIRO=bverifyInlineMessagePaymentLimit_GIRO
End Function

'[Verify Field Description displayed on GIRO Screen as]
Public Function verifyDescription_GIRO(strExpectedText)
   bverifyDescription_GIRO=true
   If Not IsNull(strExpectedText) Then
   
   For iCountero = 1 To 180 Step 1
		If Not GIRO.lblDescription.Exist(0.5) Then
			Wait(0.5)
		else
			Exit for
		End if
	Next
       If Not VerifyInnerText (GIRO.lblDescription(), strExpectedText, "Description")Then
           bverifyDescription_GIRO=false
       End If
   End If
   verifyDescription_GIRO=bverifyDescription_GIRO
End Function

'[Verify Field KnowledgeBase on GIRO SR Screen displayed as]
Public Function verifyKnowledgeBase_GIRO(strExpectedLink)
   bDevPending=false
   bverifyKnowledgeBase_GIRO=true
   If Not IsNull(strExpectedLink) Then		
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"		
			'strKBLink=GIRO.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
			strKBLink=GIRO.lnkKnowledgeBase.GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bverifyKnowledgeBase_GIRO=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_GIRO=bverifyKnowledgeBase_GIRO
End Function

'[Perform Add Notes by clicking Add Notes Button on GIRO SR Screen]
Public Function addNote_GIRO(strNote)
   bDevPending=false
   baddNote_GIRO=true	
	If not isNull(strNote) Then
		GIRO.btnAddNotes.click
		WaitForICallLoading
           If Not GIRO.popupValidationMessage.exist(5)Then
			  LogMessage "WARN","Verification","Add New Comment action failed"
			  baddNote_GIRO=false
		   else
			  LogMessage "RSLT","Verification","Add New Comment performed successfully" ,true
			  baddNote_GIRO=True
	  	   End If
		GIRO.txtNotes_Comment.set strNote
		GIRO.btnOK_ConfirmationMsg.Click
		WaitForIcallLoading
	End If		
	addNote_GIRO=baddNote_GIRO
End Function

'[Verify Combobox Reason in GIRO Screen displayed as]
Public Function verifyReason_Default(strReason)
   bDevPending=false
   bverifyReason_Default=true
   If Not IsNull(strReason) Then
       If Not verifyComboSelectItem (GIRO.lstReason(),strReason, "Reason")Then
           bverifyReason_Default=false
       End If
   End If
   verifyReason_Default=bverifyReason_Default
End Function

'[Verify Reason Combobox has Items]
Public Function verifyReasonComboboxItems(lstItems)
   bDevPending=false
   bverifyReasonComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (GIRO.lstReason(),lstItems, "Reason")Then
           bverifyReasonComboboxItems=false
       End If
   End If
   verifyReasonComboboxItems=bverifyReasonComboboxItems
End Function

'[Select Combobox Reason in GIRO Screen]
Public Function selectReason_GIRO(strReason)
	bselectReason_GIRO=true
	If Not IsNull(strRequestType) Then
       If Not (selectItem_Combobox (GIRO.lstReason(), strReason))Then
            LogMessage "WARN","Verification","Failed to select :"&strReason&" From Reason drop down list" ,false
           bselectReason_GIRO=false
       End If
   End If
   WaitForICallLoading
   selectReason_GIRO=bselectReason_GIRO
End Function

'[Select Suspension Duration in GIRO SR Screen]
Public Function selectSuspensionDuration(strFromDate,strToDate)
	WaitForICallLoading
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
		'GIRO.txtFromDate.set strFromDate
		selectSuspensionDuration = SelectDatePicker_FromDate(strFromDate)
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
		'GIRO.txtToDate.set strToDate
		selectSuspensionDuration = SelectDatePicker_TODate(strToDate)
	End If	
	If Err.Number<>0 Then
       selectSuspensionDuration=false
       LogMessage "WARN","Verification","Failed to enter Suspension Duration" ,false
       Exit Function
   End If
   WaitForICallLoading
   selectSuspensionDuration=true
End Function

'[Set TextBox Comments to GIRO]
Public Function setCommentsTextbox_GIRO(strComment)
   bDevPending=false
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comments to GIRO"
	gstrParameterNameStep = "TimeStamp"&replace((replace((replace(now,"/","-"))," ","-")),":","-")
	insertDataStore gstrParameterNameStep, strComment
	'insertDataStore "SRComment", strComment
   GIRO.txtComment.Set strComment
   If Err.Number<>0 Then
       setCommentsTextbox_GIRO=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comments" ,false
       Exit Function
   End If
   setCommentsTextbox_GIRO=true
End Function

'[Set text in Payment Limit TextBox in GIRO SR Screen]
Public Function setPaymentLimit_GIRO(strPaymentLimit)	
	If Not Isnull (strPaymentLimit) Then
		GIRO.txtPaymentLimit.set strPaymentLimit
		If Err.Number<>0 Then
			setPaymentLimit_GIRO = False
			LogMessage "WARN","Verification","Failed to Set Text Box :Payment Limit" ,False
			Exit Function
		End If
	End If
	setPaymentLimit_GIRO = True
End Function

'[Select Radio Button of Amendment on GIRO Screen]
Public Function selectAmendment_GIRO(strAmendment)
	bDevPending=false
	bselectAmendment_GIRO=true
	If Not IsNull (strAmendment) Then
		bselectAmendment_GIRO=SelectRadioButtonGrp(strAmendment,GIRO.rbtnAmendment, Array("Payment Limit","Debiting Account"))
	End If
	If Err.Number<>0 Then
       bselectAmendment_GIRO=false
       LogMessage "WARN","Verification","Failed to Click Button : Amendment" ,false
       Exit Function
   End If
   selectAmendment_GIRO=bselectAmendment_GIRO
End Function

'[Verify Popup Request Submitted exist for GIRO]
Public Function verifyPopupRequestSubmitted_GIRO(bExist)
   bDevPending=false
   bActualExist=GIRO.popupRequestSubmitted.Exist(4)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupRequestSubmitted_GIRO=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupRequestSubmitted_GIRO=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupRequestSubmitted_GIRO=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupRequestSubmitted_GIRO=False
   End If
End Function

'[Verify Field CardNumber on Request Submitted Popup for GIRO displayed as]
Public Function verifyCardNumber_RequestSubmitted_GIRO(strCardNumber)
   bDevPending=false
   bverifyCardNumber_RequestSubmitted=true
   insertDataStore "NewSAUsedCard", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       If Not VerifyInnerText (GIRO.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bverifyCardNumber_RequestSubmitted=false
       End If
   End If
   verifyCardNumber_RequestSubmitted_GIRO=bverifyCardNumber_RequestSubmitted
End Function

'[Verify Field ProductDescription on Request Submitted Popup for GIRO displayed as]
Public Function verifyProductDescription_RequestSubmitted_GIRO(strProductDescription)
   bDevPending=false
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (GIRO.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_GIRO=bVerifyProductDescription_RequestSubmittedText
End Function

'[Click Close button on Request Submitted Popup for GIRO]
Public Function verifybtnClose_RequestSubmitted_GIRO()
	bverifybtnClose_RequestSubmitted_CBR=true
	GIRO.btnCancel_RequestSubmitted.click
   If Err.Number<>0 Then
       bverifybtnClose_RequestSubmitted_CBR=false
       LogMessage "WARN","Verification","Failed to Click Close Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
   WaitForICallLoading
	verifybtnClose_RequestSubmitted_GIRO=bverifybtnClose_RequestSubmitted_CBR
End Function

'[Verify GIRO Setpup postvalidation from ARMB screen]
Public Function verifyGIROSetpup(strCardNumber,strBankAccountType,strBankID,strAccount,strStatus,strRequestDay,strPayment,strNominalAmount)
	bverifyGIROSetpup=true
	verifyGIROSetup_ARMB(strCardNumber)
	If  Ucase(Trim(strRunTimeBankAccountType)) = UCase(Trim(strBankAccountType)) Then
		LogMessage "RSLT", "Verification","Bank Account Type successfully matched with the expected value. Expected: "+ strBankAccountType &" , Actual: "& strRunTimeBankAccountType, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Bank Account Type not matching with the expected value. Expected: "+ strBankAccountType &" , Actual: "& strRunTimeBankAccountType, False
		bverifyGIROSetpup = False
	End If
	If  Ucase(Trim(strRunTimeBankID)) = UCase(Trim(strBankID)) Then
		LogMessage "RSLT", "Verification","Bank ID successfully matched with the expected value. Expected: "+ strBankID &" , Actual: "& strRunTimeBankID, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Bank ID not matching with the expected value. Expected: "+ strBankID &" , Actual: "& strRunTimeBankID, False
		bverifyGIROSetpup = False
	End If
	If  Ucase(Trim(strRunTimeAccount)) = UCase(Trim(strAccount)) Then
		LogMessage "RSLT", "Verification","Account successfully matched with the expected value. Expected: "+ strAccount &" , Actual: "& strRunTimeAccount, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Account not matching with the expected value. Expected: "+ strAccount &" , Actual: "& strRunTimeAccount, False
		bverifyGIROSetpup = False
	End If
	If  Ucase(Trim(strRunTimeStatus_FTSP)) = UCase(Trim(strStatus)) Then
		LogMessage "RSLT", "Verification","Status successfully matched with the expected value. Expected: "+ strStatus &" , Actual: "& strRunTimeStatus_FTSP, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Status not matching with the expected value. Expected: "+ strStatus &" , Actual: "& strRunTimeStatus_FTSP, False
		bverifyGIROSetpup = False
	End If
	If  Ucase(Trim(strRunTimeRequestDay)) = UCase(Trim(strRequestDay)) Then
		LogMessage "RSLT", "Verification","Request Day successfully matched with the expected value. Expected: "+ strRequestDay &" , Actual: "& strRunTimeRequestDay, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Request Day not matching with the expected value. Expected: "+ strRequestDay &" , Actual: "& strRunTimeRequestDay, False
		bverifyGIROSetpup = False
	End If
	If  Ucase(Trim(strRunTimePayment)) = UCase(Trim(strPayment)) Then
		LogMessage "RSLT", "Verification","Payment successfully matched with the expected value. Expected: "+ strPayment &" , Actual: "& strRunTimePayment, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Payment not matching with the expected value. Expected: "+ strPayment &" , Actual: "& strRunTimePayment, False
		bverifyGIROSetpup = False
	End If
	If  Ucase(Trim(strRunTimeNominalAmount)) = UCase(Trim(strNominalAmount)) Then
		LogMessage "RSLT", "Verification","Payment successfully matched with the expected value. Expected: "+ strPayment &" , Actual: "& strNominalAmount, True
		bverifyGIROSetpup = True
	else
		LogMessage "WARN", "Verification","Payment not matching with the expected value. Expected: "+ strPayment &" , Actual: "& strNominalAmount, False
		bverifyGIROSetpup = False
	End If
	verifyGIROSetpup=bverifyGIROSetpup
End Function


'[Click on GIRO link under Banking Facilities in Overview Page]
Public Function ClickOnGIROlinkEQ()
    ClickOnGIROlinkEQ = true
    bcCustomerOverview.lnkGIRO.Click
	Wait(20)    
	If Err.Number<>0 Then
       ClickOnGIROlinkEQ=false
       LogMessage "WARN","Verification","Failed to Click Button : GIRO Link" ,false
       Exit Function
   	End If
	WaitForICallLoading
End Function

'[Verify the GIRO Tab]
Public Function VerifyGIROTabEQ(strVerifyGIROTabEQ)
	VerifyGIROTabEQ = false	
	If Not IsNull (strVerifyGIROTabEQ) Then
		VerifyGIROTabEQ = verifyInnerText(GIRO.lblGIROTabEQ(),strVerifyGIROTabEQ, "GIROTab")
	End If
End Function

'[Verify default Account Number in GIRO Page displayed as]
Public Function VrfyDefValAccountDropdownEQ(strDefValAccDropDownEQ)
   VrfyDefValAccountDropdownEQ = true
   If Not IsNull(strDefValAccDropDownEQ) Then
       If Not verifyComboSelectItem (GIRO.lstAccount_GIRO(),strDefValAccDropDownEQ, "Account")Then
    	  LogMessage "WARN","Verification","Expected Default Account type:"&strDefValAccDropDownEQ&" not displayed in the Account field" ,false
          VrfyDefValAccountDropdownEQ = false
       End If
   End If
End Function

'[Verify list of values displayed in Account Number dropdown]
Public Function VerifylistAccountNo_EQ(lstAccNoEQ) 
	VerifylistAccountNo_EQ = True 
	If Not IsNull(lstAccNoEQ) Then
	 If IsArray(lstAccNoEQ) Then
		If Not verifyComboboxItems(GIRO.lstAccount_GIRO(),lstAccNoEQ, "Account")Then
       	   LogMessage "WARN","Verification","List of Account No displayed in the combox box is not as expected" ,false
           VerifylistAccountNo_EQ = false
       End If
     Else
        VerifylistAccountNo_EQ = verifyComboSelectItem(GIRO.lstAccount_GIRO(),lstAccNoEQ, "Account")
	 End If
    End If
End Function

'[Select Account Number combox box as]
Public Function SelectAccountNoEQ(strAccNoEQ)
	SelectAccountNoEQ = true
	If Not IsNull(strAccNoEQ) Then
       If Not (selectItem_Combobox (GIRO.lstAccount_GIRO(), strAccNoEQ))Then
           LogMessage "WARN","Verification","Failed to select :"&strAccNoEQ&" From Account No dropdown list" ,false
           SelectAccountNoEQ = false
       End If
   End If
   WaitForICallLoading
End Function

'[Verify GIRO Details table details displayed based on the selected Account type from the dropdown]
Public Function verifyGIROdetails_EQ(lstlstGIRODetailsEQ)
   verifyGIROdetails_EQ = verifyTableContentList(GIRO.tbGIRODetailsHeader, GIRO.tbGIRODetailsContent,lstlstGIRODetailsEQ,"GIRO Details",false,NULL,NULL,NULL)
End Function

'[Click on View hyperlink]
Public Function ClickOnViewHyperlink_EQ(lstGIROdetailsEQ)
	ClickOnViewHyperlink_EQ = selectTableLink(GIRO.tbGIRODetailsHeader, GIRO.tbGIRODetailsContent, lstGIROdetailsEQ, "GIRO Details", "Transaction Details", false, null, null, null)
End Function

'[Verify GIRO Transaction details displayed for the selected Account No in the table displayed]
Public Function verifyGIROTransactionDetails_EQ(lstlstGIROTransactionDetailsEQ)
   verifyGIROTransactionDetails_EQ = verifyTableContentList(GIRO.lblgiroViewDatatableheader, GIRO.lblgiroViewDatatablecontent,lstlstGIROTransactionDetailsEQ,"GIRO Transaction Details",false,NULL,NULL,NULL)
   WaitForICallLoading
   ClickOnGIROOKButton
End Function

'[Click on GIRO Ok Button]
Public Function ClickOnGIROOKButton()	
	ClickOnGIROOKButton=true
	GIRO.lblOKButton.Click
	If Err.Number<>0 Then
       ClickOnGIROOKButton=false
       LogMessage "WARN","Verification","Failed to Click Button : OK" ,false
       Exit Function
   	End If
	WaitForICallLoading
End Function

'[Click on Status hyperlink]
Public Function ClickOnStatusHyperlink_EQ(lstGIROdetailsEQ)
	ClickOnStatusHyperlink_EQ = selectTableLink(GIRO.tbGIRODetailsHeader, GIRO.tbGIRODetailsContent, lstGIROdetailsEQ, "GIRO Details", "Status", false, null, null, null)
End Function

'[Verify Additional Details displayed for the selected Account No in the table displayed]
Public Function verifyGIROAdditionalDetails_EQ(lstGIROAdditionalDetailsEQ)
	verifyGIROAdditionalDetails_EQ = true	
	intSize = Ubound(lstGIROAdditionalDetailsEQ)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstGIROAdditionalDetailsEQ(Iterator),":")(0))
		arrValue = trim(Split(lstGIROAdditionalDetailsEQ(Iterator),":")(1))
		arrValue = checknull(arrValue)
		Select Case (arrLabel)
			Case "Suspension Date"
				If Not IsNull(arrValue) Then
				   verifyGIROAdditionalDetails_EQ = VerifyInnerText (GIRO.lblSuspensionDateEQ(), arrValue, "Suspension Date")
			    End If
			Case "Suspension Expiry Date"
				If Not IsNull(arrValue) Then
				   verifyGIROAdditionalDetails_EQ = VerifyInnerText (GIRO.lblSuspensionExpiryDateEQ(), arrValue, "Suspension Expiry Date")
			    End If
		    Case "Termination Date"
		      If Not IsNull(arrValue)  Then
		      	verifyGIROAdditionalDetails_EQ = VerifyInnerText (GIRO.lblTerminationDateEQ(), arrValue, "Termination Date")
		      End If		    
		End Select
	Next
	WaitForICallLoading
	verifyGIROAdditionalDetails_EQ = ClickOnGIROOKButton()
	WaitForICallLoading
End Function

'[Click on Account No hyperlink]
Public Function ClickOnAccountNoHyperlink_EQ(lstGIROAccountNoEQ)
	ClickOnAccountNoHyperlink_EQ = selectTableLink(GIRO.tbGIRODetailsHeader, GIRO.tbGIRODetailsContent, lstGIROAccountNoEQ, "GIRO Details", "Account No.", false, null, null, null)
End Function

'[Verify Info Warn message on GIRO screen]
Public Function verifyInfoWarnGIROpopup(strInfoWarn)
	bverifyInfoWarn = true
	strInfoWarn = Replace(strInfoWarn,"@","=")
	strInfoWarn = Replace(strInfoWarn,"#",";")
	If not (bcInfoWarning.Popup.Exist(1)) Then
		LogMessage "WARN","Verification","Failed to open Info Warn popup." ,false
		bverifyInfoWarn=false
		Exit Function
	End If
	If Not VerifyInnerText(bcInfoWarning.lblMessage_InfoWarn(),strInfoWarn,"Info Warn Message") Then
		bverifyInfoWarn = False
	End If
	bcInfoWarning.btnOK.click
	verifyInfoWarnGIROpopup = bverifyInfoWarn
End Function
