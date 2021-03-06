'*****This is auto generated code using code generator please Re-validate ****************

'[Verify Tab Fee Adjustment is displayed]
Public Function verifyTabFeeAdjustmentExist()
   bDevPending=false
   verifyTabFeeAdjustmentExist=verifyTabExist("Fee Adjustment")
End Function

'[Verify Table SelectedCards displayed on Fee Adjustment SR Screen]
Public Function verifySelectedCardsTabledisplayed_FA()
   bDevPending=False
   verifySelectedCardsTabledisplayed_FA= FeeAdjustment.tblSelectedCardsHeader.Exist(1)
End Function
'[Verify Table SelectedCards on Fee Adjustment SR Screen has following Columns]
Public Function verifySelectedCardsTableColumns_FA(arrColumnNameList)
   bDevPending=False
   verifySelectedCardsTableColumns_FA=verifyTableColumns(FeeAdjustment.tblSelectedCardsHeader,arrColumnNameList)
End Function
'[Verify row Data in Table SelectedCards on Fee Adjustment SR Screen]
Public Function verifytblSelectedCards_RowData_FA(arrRowDataList)
   bDevPending=False
   verifytblSelectedCards_RowData_FA=verifyTableContentList(FeeAdjustment.tblSelectedCardsHeader,FeeAdjustment.tblSelectedCardsContent,arrRowDataList,"SelectedCards" , False,Null ,Null,Null)
End Function

'[Verify Table SelectedTransaction displayed on Fee Adjustment SR Screen]
Public Function verifySelectedTransactionTabledisplayed_FA()
   bDevPending=False
   verifySelectedTransactionTabledisplayed_FA= FeeAdjustment.tblSelectedTransactionHeader.Exist(1)
End Function

'[Verify Table SelectedTransaction on Fee Adjustment SR Screen has following Columns]
Public Function verifySelectedTransactionTableColumns_FA(arrColumnNameList)
   bDevPending=False
   verifySelectedTransactionTableColumns_FA=verifyTableColumns(FeeAdjustment.tblSelectedTransactionHeader,arrColumnNameList)
End Function
'[Verify row Data in Table SelectedTransaction on Fee Adjustment SR Screen]
Public Function verifytblSelectedTransaction_RowData_FA(arrRowDataList)
   bDevPending=False
   verifytblSelectedTransaction_RowData_FA=verifyTableContentList(FeeAdjustment.tblSelectedTransactionHeader,FeeAdjustment.tblSelectedTransactionContent,arrRowDataList,"SelectedTransaction" , False,Null ,Null,Null)
End Function

'[Verify Field FeeType on Fee Adjustment SR Screen displayed as]
Public Function verifyFeeTypeText(strExpectedText)
   bDevPending=False
   bVerifyFeeTypeText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblFeeType(), strExpectedText, "Fee Type")Then
           bVerifyFeeTypeText=false
       End If
   End If
   verifyFeeTypeText=bVerifyFeeTypeText
End Function

'[Select Combobox FeeType on Fee Adjustment SR Screen as]
Public Function selectFeeTypeComboBox_FA(strFeeType)
   bDevPending=False
   bSelectFeeTypeComboBox=true
   If Not IsNull(strFeeType) Then
       If Not (selectItem_Combobox (FeeAdjustment.lstFeeType(), strFeeType))Then
            LogMessage "WARN","Verification","Failed to select :"&strFeeType&" From FeeType drop down list" ,false
           bSelectFeeTypeComboBox=false
       End If
   End If
   selectFeeTypeComboBox_FA=bSelectFeeTypeComboBox
End Function

'[Verify Combobox FeeType on Fee Adjustment SR Screen displayed as]
Public Function verifyFeeTypeCombo_FA(strExpectedText)
   bDevPending=False
   bVerifyFeeTypeText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (FeeAdjustment.lstFeeType(), strExpectedText, "FeeType")Then
           bVerifyFeeTypeText=false
       End If
   End If
   verifyFeeTypeCombo_FA=bVerifyFeeTypeText
End Function

'[Verify Combobox Fee Type on Fee Adjustment SR Screen has items]
Public Function verifyFeeType_ItemList(lstItems)
   bDevPending=false
   bVerifyFeeType=true
   If Not IsNull(lstItems) Then
	
       If Not verifyComboboxItems (FeeAdjustment.lstFeeType, lstItems, "Urgency")Then
           bVerifyFeeType=false
       End If
   End If
   verifyFeeType_ItemList=bVerifyFeeType
End Function
'[Verify If Combobox Fee Type is enabled on Fee Adjustment SR Screen]
Public Function VerifyFeeTypeCombo_enabled(bEnabled)
	bDevPending=False
	If isnull(bEnabled) Then
		'If not FeeAdjustment.lstFeeType.exist(1) Then
			LogMessage "RSLT","Verification","Fee Type Combobox is not displayed",true
			VerifyFeeTypeCombo_enabled=true
			Exit Function
		'End If
	End If
   Dim bVerifyFeeType:bVerifyFeeType=true
	intBtnLookUp=Instr(FeeAdjustment.lblFeeType.Object.GetAttribute("disabled"),"disabled")

	If bEnabled Then
		If  intBtnLookUp=0 Then
			LogMessage "RSLT","Verification","Combobox Fee Type is enable as per expectation.",True
			bVerifyFeeType=true
		Else
			LogMessage "WARN","Verifiation","Combobox Fee Type is disable. Expected to be enable.",false
			bVerifyFeeType=false
		End If
	else
		If  intBtnLookUp<>0 Then
			LogMessage "RSLT","Verification","Combobox Fee Type is disabled as per expectation.",True
			bVerifyFeeType=true
		Else
			LogMessage "WARN","Verifiation","Combobox Fee Type is Enabled. Expected to be disabled.",false
			bVerifyFeeType=false
		End If
	End If
    
	VerifyFeeTypeCombo_enabled=bVerifyFeeType
End Function

'[Select Combobox ApprovalLevel on Fee Adjustment SR Screen as]
Public Function selectApprovalLevelComboBox_FA(strApprovalLevel)
   bDevPending=False
   bSelectApprovalLevelComboBox=true
   If Not IsNull(strApprovalLevel) Then
       If Not (selectItem_Combobox (FeeAdjustment.lstApprovalLevel(), strApprovalLevel))Then
            LogMessage "WARN","Verification","Failed to select :"&strApprovalLevel&" From ApprovalLevel drop down list" ,false
           bSelectApprovalLevelComboBox=false
       End If
   End If
   selectApprovalLevelComboBox_FA=bSelectApprovalLevelComboBox
End Function

'[Verify Combobox ApprovalLevel on Fee Adjustment SR Screen displayed as]
Public Function verifyApprovalLevelCombo_FA(strExpectedText)
   bDevPending=False
   bVerifyApprovalLevelText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (FeeAdjustment.lstApprovalLevel(), strExpectedText, "ApprovalLevel")Then
           bVerifyApprovalLevelText=false
       End If
   End If
   verifyApprovalLevelCombo_FA=bVerifyApprovalLevelText
End Function

'[Select Combobox AdjustmentReason on Fee Adjustment SR Screen as]
Public Function selectAdjustmentReasonComboBox_FA(strAdjustmentReason)
   bDevPending=False
   bSelectAdjustmentReasonComboBox=true
   If Not IsNull(strAdjustmentReason) Then
       If Not (selectItem_Combobox (FeeAdjustment.lstAdjustmentReason(), strAdjustmentReason))Then
            LogMessage "WARN","Verification","Failed to select :"&strAdjustmentReason&" From AdjustmentReason drop down list" ,false
           bSelectAdjustmentReasonComboBox=false
       End If
   End If
   selectAdjustmentReasonComboBox_FA=bSelectAdjustmentReasonComboBox
End Function

'[Verify Combobox AdjustmentReason on Fee Adjustment SR Screen displayed as]
Public Function verifyAdjustmentReasonText_FA(strExpectedText)
   bDevPending=False
   bVerifyAdjustmentReasonText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (FeeAdjustment.lstAdjustmentReason(), strExpectedText, "AdjustmentReason")Then
           bVerifyAdjustmentReasonText=false
       End If
   End If
   verifyAdjustmentReasonText_FA=bVerifyAdjustmentReasonText
End Function
'[Verify If Combobox Adjustment Reason is enabled on Fee Adjustment SR Screen]
Public Function VerifyAdjReasonCombo_enabled(bEnabled)
	bDevPending=False
	If isNull(bEnabled) Then
		VerifyAdjReasonCombo_enabled=true
		Exit Function
	End If
   Dim bVerifyAdjReason:bVerifyAdjReason=true
	intBtnLookUp=Instr(FeeAdjustment.lstAdjustmentReason.GetROproperty("outerhtml"),"v-disabled")

	If bEnabled Then
		If  intBtnLookUp=0 Then
			LogMessage "RSLT","Verification","Combobox Adjustment Reason is enable as per expectation.",True
			bVerifyAdjReason=true
		Else
			LogMessage "WARN","Verifiation","Combobox Adjustment Reason is disable. Expected to be enable.",false
			bVerifyAdjReason=false
		End If
	else
		If  intBtnLookUp<>0 Then
			LogMessage "RSLT","Verification","Combobox Adjustment Reason is disabled as per expectation.",True
			bVerifyAdjReason=true
		Else
			LogMessage "WARN","Verifiation","Combobox Adjustment Reason is Enabled. Expected to be disabled.",false
			bVerifyAdjReason=false
		End If
	End If
    
	VerifyAdjReasonCombo_enabled=bVerifyAdjReason
End Function


'[Verify Field EffectiveDate on Fee Adjustment SR Screen displayed as]
Public Function verifyEffectiveDateText_FA(strExpectedText)
   bDevPending=False
   bVerifyEffectiveDateText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( FeeAdjustment.txtEffectiveDate(), strExpectedText, "EffectiveDate")Then
           bVerifyEffectiveDateText=false
       End If
   End If
   verifyEffectiveDateText_FA=bVerifyEffectiveDateText
End Function

'[VerifyTextBox EffectiveDate is enabled on Fee Adjustment SR Screen]
Public Function VerifyEffectiveDate_enabled(bEnabled)
	bDevPending=False
   Dim bVerifyEffectiveDate:bVerifyEffectiveDate=true
	'intBtnLookUp=Instr(FeeAdjustment.txtEffectiveDate.Object.GetAttribute("disabled"),"disabled")
	intBtnLookUp=Instr(FeeAdjustment.lblEffectiveDate.Object.GetAttribute("class"),"disabled-area ")
	If bEnabled Then
		If  intBtnLookUp=0 Then
			LogMessage "RSLT","Verification","Textbox Effective Date is enable as per expectation.",True
			bVerifyEffectiveDate=true
		Else
			LogMessage "WARN","Verifiation","Textbox Effective Date is disable. Expected to be enable.",false
			bVerifyFeeType=false
		End If
	else
		If  intBtnLookUp<>0 Then
			LogMessage "RSLT","Verification","Textbox Effective Date is disabled as per expectation.",True
			bVerifyEffectiveDate=true
		Else
			LogMessage "WARN","Verifiation","Textbox Effective Date is Enabled. Expected to be disabled.",false
			bVerifyEffectiveDate=false
		End If
	End If
    
	VerifyEffectiveDate_enabled=bVerifyEffectiveDate
End Function

'[Set TextBox EffectiveDate on Fee Adjustment SR Screen to]
Public Function setEffectiveDateTextbox_FA(strEffectiveDate)
   bDevPending=False
   If not isNull(strEffectiveDate) Then
	   FeeAdjustment.txtEffectiveDate.Set(strEffectiveDate)
	   If Err.Number<>0 Then
		   setEffectiveDateTextbox_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :EffectiveDate" ,false
		   Exit Function
	   End If
   End If
   setEffectiveDateTextbox_FA=true
End Function

'[Verify Field RequestedAmount on Fee Adjustment SR Screen displayed as]
Public Function verifyRequestedAmountText_FA(strExpectedText)
   bDevPending=False
   bVerifyRequestedAmountText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( FeeAdjustment.txtRequestedAmount(), strExpectedText, "RequestedAmount")Then
           bVerifyRequestedAmountText=false
       End If
   End If
   verifyRequestedAmountText_FA=bVerifyRequestedAmountText
End Function

'[Set TextBox RequestedAmount on Fee Adjustment SR Screen to]
Public Function setRequestedAmountTextbox_FA(strRequestedAmount)
   bDevPending=False
   FeeAdjustment.txtRequestedAmount.Set(strRequestedAmount)
   If Err.Number<>0 Then
       setRequestedAmountTextbox_FA=false
            LogMessage "WARN","Verification","Failed to Set Text Box :RequestedAmount" ,false
       Exit Function
   End If
   setRequestedAmountTextbox_FA=true
End Function

'[Verify Field Description displayed on Fee Adjustment SR Screen as]
Public Function verifyDescriptionText_FA(strExpectedText)
   bDevPending=False
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_FA=bVerifyDescriptionText
End Function

'[Verify Field Error Message on Fee Adjustment SR Screen displayed as]
Public Function verifyErrorMessage_FA(strExpectedText)
   bDevPending=False
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblErrorMsg(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   Else
		If FeeWaiver.lblErrorMsg.Exist(1) Then
			LogMessage "RSLT","Verification","Unexpected Error message displayed",true
			bVerifyDescriptionText=false
		End If
   End If
   verifyErrorMessage_FA=bVerifyDescriptionText
End Function

'[Verify Knowledge base link is enabled on Fee Adjustment SR Screen]
Public Function VerifyKnowledgebaselinkEnabled_FA()
      bDevPending=false
   Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
     strKBLink=FeeAdjustment.lnkKnowledgeBase.GetROProperty("Outerhtml")
	
    If inStr(strKBLink,"v-disabled") = 0Then
		LogMessage "RSLT","Verification","Knowledge base Link  enabled successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not enabled as expected",false
		bVerifyKnowledgebaselink=false
	End If
	VerifyKnowledgebaselinkEnabled_FA=bVerifyKnowledgebaselink
End Function

'[Verify Field KnowledgeBase on Fee Adjustment SR Screen displayed as]
Public Function verifyKnowledgeBase_FA(strExpectedLink)
   bDevPending=False
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strExpectedText) Then
		
	Set oDesc_KB = Description.Create()
	oDesc_KB("micclass").Value = "Link"
	'strKBLink=FeeAdjustment.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
	strKBLink=FeeAdjustment.lnkKnowledgeBase.GetROProperty("href")
	strExpectedLink=Replace(strExpectedLink,"@","=")			
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_FA=bVerifyKnowledgeBaseText
End Function

'[Verify Field Comment displayed on Fee Adjustment SR Screen as]
Public Function verifyCommentText_FA(strExpectedText)
   bDevPending=False
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( FeeAdjustment.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText_FA=bVerifyCommentText
End Function


'[Set TextBox Comment on Fee Adjustment SR Screen to]
Public Function setCommentTextbox_FA(strComment)
   bDevPending=False
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comment on Fee Adjustment SR Screen to"
	insertDataStore "SRComment", strComment
	
   FeeAdjustment.txtComment.Set(strComment )
   If Err.Number<>0 Then
       setCommentTextbox_FA=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_FA=true
End Function

'[Click Button Submit on Fee Adjustment SR Screen]
Public Function clickButtonSubmit_FA()
   bDevPending=False
   'intBtnSubmit=Instr(FeeAdjustment.btnSubmit.GetROproperty("outerhtml"),("v-disabled"))
   intBtnSubmit=Instr(FeeAdjustment.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
   If intBtnSubmit<>0 Then
		LogMessage "RSLT", "Verification","Submit button is disabled",false
	   clickButtonSubmit_FA=false
	   Exit Function
   End If
   FeeAdjustment.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_FA=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   '*************** Capturing time stamp to open Memo for this SR by Manish
	'strRunTimeTimeStamp_Step="Click Button Submit on Fee Adjustment SR Screen"
	'strDate="9 Apr 2014"
	'strTempTime=FormatDateTime(now,4)
	'call VPlusLogin_DateTime	
 	'strTimeStamp=strDate&" "&strTempTime
	'insertDataStore "TimeStamp", strTimeStamp
	WaitForIcallLoading
   clickButtonSubmit_FA=true
End Function

'[Verify Button Submit is enabled on Fee Adjustment SR Screen]
Public Function VerifybtnSubmit_FA(bEnabled)
	bDevPending=False
   Dim bVerifybtnSubmit:bVerifybtnSubmit=true
	intBtnSubmit=Instr(FeeAdjustment.btnSubmit.Object.GetAttribute("disabled"),("disabled"))

	If bEnabled Then
		If  intBtnSubmit=0 Then
			LogMessage "RSLT","Verification","Submit button is enable as per expectation.",True
			bVerifyButtonSubmit=true
		Else
			LogMessage "WARN","Verifiation","Submit button is disable. Expected to be enable.",false
			bVerifyButtonSubmit=false
		End If
	else
		If  intBtnSubmit<>0 Then
			LogMessage "RSLT","Verification","Submit button is disabled as per expectation.",True
			bVerifyButtonSubmit=true
		Else
			LogMessage "WARN","Verifiation","Submit button is Enabled. Expected to be disabled.",false
			bVerifyButtonSubmit=false
		End If
	End If
	VerifybtnSubmit_FA=bVerifyButtonSubmit
End Function

'[Click Button Cancel on Fee Adjustment SR Screen]
Public Function clickButtonCancel_FA()
   bDevPending=False
   FeeAdjustment.btnCancel.click

   If Err.Number<>0 Then
       clickButtonCancel_FA=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickButtonCancel_FA=true
End Function

'[Verify Popup RequestSubmitted on Fee Adjustment SR Screen exist]
Public Function verifyPopupRequestSubmittedexist_FA(bExist)
   bDevPending=False
   bActualExist=FeeAdjustment.popupRequestSubmitted.Exist(1)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupRequestSubmittedexist_FA=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupRequestSubmittedexist_FA=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupRequestSubmittedexist_FA=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupRequestSubmittedexist_FA=False
   End If
End Function

'[Verify Field CardNumber_RequestSubmitted on Fee Adjustment SR Screen displayed as]
Public Function verifyCardNumber_RequestSubmittedText_FA(strExpectedText)
   bDevPending=False
   bVerifyCardNumber_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblCardNumber_RequestSubmitted(), strExpectedText, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber_RequestSubmittedText=false
       End If
   End If
   verifyCardNumber_RequestSubmittedText_FA=bVerifyCardNumber_RequestSubmittedText
End Function

'[Verify Field ProductDescription_RequestSubmitted on Fee Adjustment SR Screen displayed as]
Public Function verifyProductDescription_RequestSubmittedText_FA(strExpectedText)
   bDevPending=False
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblProductDescription_RequestSubmitted(), strExpectedText, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmittedText_FA=bVerifyProductDescription_RequestSubmittedText
End Function

'[Verify Link SRNumber available on Request Submitted popup for Fee Adjustment]
Public Function verifyLinkSRNumber_RequestSubmitted_FA()
   bDevPending=False
   bverifyLinkSRNumber_RequestSubmitted=true
	strSelectedSR=FeeAdjustment.lnkSRNumber_RequestSubmitted.GetRoProperty("innerText")
	If instr(FeeAdjustment.lnkSRNumber_RequestSubmitted.GetRoProperty("class"),"link")=0 Then
		bverifyLinkSRNumber_RequestSubmitted=false
	else
		bverifyLinkSRNumber_RequestSubmitted=true
	end If
	LogMessage "RSLT","Verification","SR Number link "& strSelectedSR &" displayed on Request Submitted popup",true
	If IsNull(strSRNumber) Then
		LogMessage "WARN","Verification", "SR Number not available with link on Request Submitted popup.",false
		bverifyLinkSRNumber_RequestSubmitted=false
	End If

   verifyLinkSRNumber_RequestSubmitted_FA=bverifyLinkSRNumber_RequestSubmitted
End Function

'[Click Link SRNumber_RequestSubmitted on Fee Adjustment SR Screen]
Public Function clickLinkSRNumber_RequestSubmitted_FA()
   bDevPending=False
   gstrRuntimeSRNumStep="Click Link SRNumber_RequestSubmitted on Fee Adjustment SR Screen"
	strSelectedSR=FeeAdjustment.lnkSRNumber_RequestSubmitted.GetRoProperty("innerText")
   If strSelectedSR<>"" Then
	 insertDataStore "SelectedSRLink", strSelectedSR
	  FeeAdjustment.lnkSRNumber_RequestSubmitted.click
	 else
   		LogMessage "RSLT","Verification","SR Number did not displayed on Request Submitted pop up",false
	End If
   WaitForIcallLoading
   If Err.Number<>0 Then
       clickLinkSRNumber_RequestSubmitted_FA=false
            LogMessage "WARN","Verification","Failed to Click Link : SRNumber_RequestSubmitted" ,false
       Exit Function
   End If
   clickLinkSRNumber_RequestSubmitted_FA=true
End Function

'[Verify Field Status_RequestSubmitted on Fee Adjustment SR Screen displayed as]
Public Function verifyStatus_RequestSubmittedText_FA(strExpectedText)
   bDevPending=False
   bVerifyStatus_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblStatus_RequestSubmitted(), strExpectedText, "Status_RequestSubmitted")Then
           bVerifyStatus_RequestSubmittedText=false
       End If
   End If
   verifyStatus_RequestSubmittedText_FA=bVerifyStatus_RequestSubmittedText
End Function

'[Click Button RefreshStatus on Fee Adjustment SR Screen]
Public Function clickButtonRefreshStatus_FA()
   bDevPending=False
   wait 1
   FeeAdjustment.btnRefreshStatus.click
	WaitForICallLoading
    		'Get Status
		If FeeAdjustment.lblStatus_RequestSubmitted.getROProperty("innertext")="In Progress" then 
			bStatus=true
		 else
			bStatus=false
		End If
	While  bStatus AND (iCount<60)
		FeeAdjustment.btnRefreshStatus.click
		wait 1
        	'Get Status
			strStatus=FeeAdjustment.lblStatus_RequestSubmitted.getROProperty("innertext")
			If Trim(strStatus)="In Progress" then 
				bStatus=true
			 else
				LogMessage "WARN","Verification","Status displayed as  :"&strStatus ,true
				bStatus=false
			End If
		wait 5
		intBtnRefreshStatus=Instr(FeeAdjustment.btnRefreshStatus.GetROproperty("outerhtml"),"v-disabled")
		If intBtnRefreshStatus<>0 Then
			LogMessage "WARN","Verification","Button : RefreshStatus is disabled" ,true
			bStatust=true
		End If
		iCount=iCount+1
	  Wend	
      If Err.Number<>0 Then
       clickButtonRefreshStatus_FA=false
            LogMessage "WARN","Verification","Failed to Click Button : RefreshStatus" ,false
       Exit Function
   End If
   clickButtonRefreshStatus_FA=true
End Function

'[Click Button OK_RequestSubmitted on Fee Adjustment SR Screen]
Public Function clickButtonOK_RequestSubmitted_FA()
   bDevPending=False
   FeeAdjustment.btnOK_RequestSubmitted.click
   If Err.Number<>0 Then
       clickButtonOK_RequestSubmitted_FA=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_RequestSubmitted" ,false
       Exit Function
   End If
   clickButtonOK_RequestSubmitted_FA=true
End Function

'[Perform Add Notes by clicking Add Notes Button on Fee Adjustment SR Screen]
Public Function addNote_FA(strNote)
   bDevPending=false
   bVerifypopupNotes=true
	Dim bVerifypopupNotes:VerifypopupNotes=true
	
	If not isNull(strNote) Then
		FeeAdjustment.btnAddNotes.click
		WaitForICallLoading
            If not   ServiceRequest.popupVerification.exist(5)Then
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
	addNote_FA=bVerifypopupNotes
End Function

'[Verify Popup ValidationMessage exist For Fee Adjustment]
Public Function verifyPopupValidation_FA(bExist)
   bDevPending=False
   bActualExist=FeeAdjustment.popupValidationMessage.Exist(1)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage Exists As Expected" ,true
       verifyPopupValidation_FA=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage does not Exists As Expected" ,true
       verifyPopupValidation_FA=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage does not Exists As Expected" ,False
       verifyPopupValidation_FA=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage Still Exists" ,False
       verifyPopupValidation_FA=False
   End If
End Function

'[Click Button OK_ValidationPopup For Fee Adjustment]
Public Function clickButtonOK_ValidationPopup_FA()
   bDevPending=False
   
   FeeAdjustment.btnOK_ValidationPopup.click
   If Err.Number<>0 Then
       clickButtonOK_ValidationPopup_FA=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationPopup" ,false
       Exit Function
   End If
   clickButtonOK_ValidationPopup_FA=true
End Function

'[Verify Field ValidationMessage For Fee Adjustment displayed as]
Public Function verifyValidationMessage_FA(strExpectedText)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
           bVerifyValidationMessageText=false
       End If
   End If
   verifyValidationMessage_FA=bVerifyValidationMessageText
End Function

'[Verify Confirmation Message For Fee Adjustment displayed as]
Public Function verifyConfirmationMessage_FA(strExpectedText)
   bDevPending=False
   bVerifyValidationMessageText=true
   bActualExist=FeeAdjustment.popupValidationMessage.Exist(3)
   If bActualExist Then

	   If Not IsNull(strExpectedText) Then
		   If Not VerifyInnerText (FeeAdjustment.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
			   bVerifyValidationMessageText=false
		   End If
	   End If
		FeeAdjustment.btnOK_ConfirmationPopup.click
	else
		Logmessage "RSLT","Verification","Fee Adjustment confirmation popup does not displayed",false
		bVerifyValidationMessageText=false
   End If
   verifyConfirmationMessage_FA=bVerifyValidationMessageText
End Function


'[Select Action Menu Adjust Fee from Statement Transaction table on Statement Screen]
Public Function selectAdjustFee_Statement(lstTransactionsData)
   bDevPending=False
   bSelectAdjustFee=true
 	With bcStatements
		  bSelectAdjustFee= selectTableSubMenu(.tblStatementTransactionHeader,.tblStatementTransactionContent,lstTransactionsData,"Statement Transaction","Actions",True,.btnNext,.lnkNext,.btnPrevious,"Adjust Fee",bDisabled)
	End With
	If bDisabled Then
		LogMessage "RSLT", "Verification","Adjust fee action menu is not enabled",false
		bSelectAdjustFee=false
	End If	
    selectAdjustFee_Statement=bSelectAdjustFee
End Function

'[Select Action Menu Adjust Fee from Unbilled Transaction table on Transaction History Screen]
Public Function selectAdjustFee_TransactionHistory(lstTransactionsData)
   bDevPending=False
   bSelectAdjustFee=true
 	With TransactionHistory
		  bSelectAdjustFee= selectTableSubMenu(.tblTransactionsHeader_UB,.tblTransactions_UB,lstTransactionsData,"Transaction History","Actions",True,.lnkNext1_UB,.lnkNext_UB,.lnkPrevious_UB,"Adjust Fee",bDisabled)
	End With
	If bDisabled Then
		LogMessage "RSLT", "Verification","Adjust fee action menu is not enabled",false
		bSelectAdjustFee=false
	End If	
    selectAdjustFee_TransactionHistory=bSelectAdjustFee
End Function

'[Verify if Action Menu Adjust Fee Enabled from Unbilled Transaction table on Transaction History Screen]
Public Function VerifyAdjustFee_TransactionHistory_Enabled(lstTransactionsData,bEnabled)
   bDevPending=False
   bSelectAdjustFee=true
 	With TransactionHistory
		  bSelectAdjustFee= selectTableSubMenu(.tblTransactionsHeader_UB,.tblTransactions_UB,lstTransactionsData,"Transaction History","Actions",True,.lnkNext1_UB,.lnkNext_UB,.lnkPrevious_UB,"Adjust Fee",bDisabled)
	End With
	If bEnabled Then
		If bDisabled Then
			LogMessage "RSLT", "Verification","Adjust fee action menu is not enabled",false
			bSelectAdjustFee=false
		 else
			LogMessage "RSLT", "Verification","Adjust fee action menu is enabled as expected",true
		End If
	else
		If bDisabled Then
			LogMessage "RSLT", "Verification","Adjust fee action menu is disabled as expected",true
		 else
			LogMessage "RSLT", "Verification","Adjust fee action menu is not disabled",false
			bSelectAdjustFee=false
		End If

	End If
    VerifyAdjustFee_TransactionHistory_Enabled=bSelectAdjustFee
End Function

'[Verify if Action Menu Adjust Fee Enabled from Statement Transaction table on Statement Screen]
Public Function VerifyAdjustFee_Statement_Enabled(lstTransactionsData,bEnabled)
   bDevPending=False
   bSelectAdjustFee=true
 	With bcStatements
		  bSelectAdjustFee= selectTableSubMenu(.tblStatementTransactionHeader,.tblStatementTransactionContent,lstTransactionsData,"Statement Transaction","Actions",True,.btnNext,.lnkNext,.btnPrevious,"Adjust Fee",bDisabled)
		  'bSelectAdjustFee= selectTableSubMenu(.tblStatementTransactionHeader,.tblStatementTransactionContent,lstTransactionsData,"Statement Transaction","Actions",True,.lnkNext,.btnNext,.btnPrevious,"Adjust Fee",bDisabled)
	End With
	
	If bEnabled Then
		If bDisabled Then
			LogMessage "RSLT", "Verification","Adjust fee action menu is not enabled",false
			bSelectAdjustFee=false
		 else
			LogMessage "RSLT", "Verification","Adjust fee action menu is enabled as expected",true
		End If
	else
		If bDisabled Then
			LogMessage "RSLT", "Verification","Adjust fee action menu is disabled as expected",true
		 else
			LogMessage "RSLT", "Verification","Adjust fee action menu is not disabled",false
			bSelectAdjustFee=false
		End If

	End If
    VerifyAdjustFee_Statement_Enabled=bSelectAdjustFee
End Function

'[Verify if Action Menu Adjust Fee Enabled on Other Plans Screen]
Public Function VerifyAdjustFee_OtherPlans_Enabled(lstTransactionsData,bEnabled)
   bDevPending=False
   bSelectAdjustFee=true
 	With FeeAdjustment
		  bSelectAdjustFee= selectTableSubMenu(.tblOtherPlansTransactionHeader,.tblOtherPlansTransactionContent,lstTransactionsData,"Other Plans Transaction","Actions",True,.btnNext,.lnkNext,.btnPrevious,"Adjust Fee",bDisabled)
		  'bSelectAdjustFee= selectTableSubMenu(.tblStatementTransactionHeader,.tblStatementTransactionContent,lstTransactionsData,"Statement Transaction","Actions",True,.lnkNext,.btnNext,.btnPrevious,"Adjust Fee",bDisabled)
	End With
	
	If bEnabled Then
		If bDisabled Then
			LogMessage "RSLT", "Verification","Adjust fee action menu is not enabled",false
			bSelectAdjustFee=false
		 else
			LogMessage "RSLT", "Verification","Adjust fee action menu is enabled as expected",true
		End If
	else
		If bDisabled Then
			LogMessage "RSLT", "Verification","Adjust fee action menu is disabled as expected",true
		 else
			LogMessage "RSLT", "Verification","Adjust fee action menu is not disabled",false
			bSelectAdjustFee=false
		End If

	End If
    VerifyAdjustFee_OtherPlans_Enabled=bSelectAdjustFee
End Function

'[Select Action Menu Adjust Fee from Transaction table on Other Plans Screen]
Public Function selectAdjustFee_otherPlans(lstTransactionsData)
   bDevPending=False
   bSelectAdjustFee=true
 	With FeeAdjustment
		  bSelectAdjustFee= selectTableSubMenu(.tblOtherPlansTransactionHeader,.tblOtherPlansTransactionContent,lstTransactionsData,"Other Plans","Actions",True,.btnNext,.lnkNext,.btnPrevious,"Adjust Fee",bDisabled)
	End With
	If bDisabled Then
		LogMessage "RSLT", "Verification","Adjust fee action menu is not enabled",false
		bSelectAdjustFee=false
	End If	
    selectAdjustFee_otherPlans=bSelectAdjustFee
End Function

'[Verify Adjustment Type Combo box has Items]
Public Function verifyAdjTypeComboboxItems(strlstAdjustmentType)
   bDevPending=false
   bverifyAdjTypeComboboxItems=true
   If Not IsNull(strlstAdjustmentType) Then
       If Not verifyComboboxItems (FeeAdjustment.lstAdjustmentType(), strlstAdjustmentType, "Adjustment Type has items")Then
           bverifyAdjTypeComboboxItems=false
       End If
   End If
   verifyAdjTypeComboboxItems=bverifyAdjTypeComboboxItems
End Function

'[Verify Adjustment Reason Combo box has Items]
Public Function verifyAdjReasonComboboxItems(strlstAdjustmentReason)
   bDevPending=false
   bverifyAdjReasonComboboxItems=true
   If Not IsNull(strlstAdjustmentReason) Then
       If Not verifyComboboxItems (FeeAdjustment.lstAdjustmentReason(), strlstAdjustmentReason, "Adjustment reason has items")Then
           bverifyAdjReasonComboboxItems=false
       End If
   End If
   verifyAdjReasonComboboxItems=bverifyAdjReasonComboboxItems
End Function

'[Set the adjustment reason value for the fee adjustment]
Public Function setAdjustmentReason_FA(strAdjReason)
	bsetAdjustmentReason_FA = true
	If not isNull(strAdjReason) Then
	   FeeAdjustment.txtAdjustmentReason.Set strAdjReason
	If Err.Number<>0 Then
		   bsetAdjustmentReason_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :Adjustment Reason" ,false
		   Exit Function
	   End If
   End If
	setAdjustmentReason_FA = bsetAdjustmentReason_FA
End Function

'[Verify the PopUp on the basis of the minimum payment made by the Customer]
Public Function verifyPopUp_MimPayAmt()
	bverifyPopup = true
	'bExist = verifyPopUpExist_VPlusValidation(strCardNumber)
	
	bActualExist=FeeAdjustment.popupValidationMessage.Exist(0)
	If bActualExist = true Then
		'both of them match (either both true or false)
		LogMessage "RSLT","Verification","Popup matching as expected. Expected: " &bExist&" Actual: "&bActualExist  ,True
	Else
		'mismatch
		LogMessage "WARN","Verification","Popup not matching as expected. Expected: " &bExist&" Actual: "&bActualExist  ,False
        verifyPopUp_MimPayAmt=False
	End If
   
End Function

'[Set TextBox Adjustment Type on Fee Adjustment SR Screen to]
Public Function setAdjustmentTypeTextbox_FA(strAdjustmentType)
   bDevPending=False
   If not isNull(strAdjustmentType) Then
	   FeeAdjustment.txtAdjustmentType.Set(strAdjustmentType)
	   If Err.Number<>0 Then
		   setAdjustmentTypeTextbox_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :Adjustment Type" ,false
		   Exit Function
	   End If
   End If
   setAdjustmentTypeTextbox_FA=true
End Function

'[Set TextBox Plan No on Fee Adjustment SR Screen to]
Public Function setPlanNoTextbox_FA(strPlanNo)
   bDevPending=False
   If not isNull(strPlanNo) Then
	   FeeAdjustment.txtPlanNo.Set(strPlanNo)
	   If Err.Number<>0 Then
		   setPlanNoTextbox_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :Plan No" ,false
		   Exit Function
	   End If
   End If
   setPlanNoTextbox_FA=true
End Function

'[Set TextBox sequence No on Fee Adjustment SR Screen to]
Public Function setSequenceNoTextbox_FA(strSequenceNo)
   bDevPending=False
   If not isNull(strSequenceNo) Then
	   FeeAdjustment.txtsequenceNo.Set(strSequenceNo)
	   If Err.Number<>0 Then
		   setSequenceNoTextbox_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :Sequence No" ,false
		   Exit Function
	   End If
   End If
   setSequenceNoTextbox_FA=true
End Function

'[Set TextBox Adjustment Amount on Fee Adjustment SR Screen to]
Public Function setAdjAmtTextbox_FA(strAdjustmentAmt)
   bDevPending=False
   If not isNull(strAdjustmentAmt) Then
	   FeeAdjustment.txtAdjustmentAmt.Set(strAdjustmentAmt)
	   If Err.Number<>0 Then
		   setAdjAmtTextbox_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :Adjustment Amount" ,false
		   Exit Function
	   End If
   End If
   setAdjAmtTextbox_FA=true
End Function

'[Set TextBox Other Reason on Fee Adjustment SR Screen to]
Public Function setOtherReasonTextbox_FA(strOtherReason)
   bDevPending=False
   If not isNull(strOtherReason) Then
	   FeeAdjustment.txtAdjustmentAmt.Set(strOtherReason)
	   If Err.Number<>0 Then
		   setOtherReasonTextbox_FA=false
				LogMessage "WARN","Verification","Failed to Set Text Box :Other Reason" ,false
		   Exit Function
	   End If
   End If
   setOtherReasonTextbox_FA=true
End Function

'[Verify Sequence Error Message on Fee Adjustment SR Screen displayed as]
Public Function verifySeqErrorMessage_FA(strExpectedText)
   bDevPending=False
   verifySeqErrorMessage_FA=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblSeqErrorMsg(), strExpectedText, "Sequence No")Then
           verifySeqErrorMessage_FA=false
       End If
   End If
End Function

'[Verify Adjustment Type on Fee Adjustment SR Screen displayed as]
Public Function verifyAdjustmentType_FA(strExpectedText)
   bDevPending=False
   verifyAdjustmentType_FA=true
   If Not IsNull(strExpectedText) Then
   	   'Get the actual input from the FE
   	   strActualText = FeeAdjustment.txtAdjustmentType().GetRoProperty("value")
   	   If strExpectedText <> strActualText Then
   	   		verifyAdjustmentType_FA=false
   	   		LogMessage "WARN","Verification","Adjustment Type did not match as expected. Expected: "&strExpectedText& " Actual: " &strActualText ,False
   	   Else
   	   		verifyAdjustmentType_FA=true
   	   		LogMessage "RSLT","Verification","Adjustment Type matched as expected." ,True
   	   End If
   End If
End Function

'[Verify Effective Date Error Message on Fee Adjustment SR Screen displayed as]
Public Function verifyEffDateErrorMessage_FA(strExpectedText)
   bDevPending=False
   verifyEffDateErrorMessage_FA=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeAdjustment.lblEffDateErrorMsg(), strExpectedText, "Effective Date")Then
           verifyEffDateErrorMessage_FA=false
       End If
   End If
End Function

'[Verify the visibility of the Effective date in Fee Adjustment SR Page]
Public Function verifyVisibilityEffDate_FA(bEnabled)
	verifyVisibilityEffDate_FA = true
	strClass = FeeAdjustment.txtEffectiveDate().GetRoProperty("class")
	If instr(1,strClass,"disabled") Then
		'then the field is disabled
		If bEnabled Then
			LogMessage "WARN","Verification","Effective Date field is disabled but was expected to be enable." ,false
			verifyVisibilityEffDate_FA = false
		else
			LogMessage "RSLT","Verification","Effective Date field is disabled as expected." ,True
		End If
	Else
		If bEnabled Then
			LogMessage "RSLT","Verification","Effective Date field is enabled as expected." ,True
		else
			LogMessage "WARN","Verification","Effective Date field is enabled but was expected to be disable." ,false
			verifyVisibilityEffDate_FA = false
		End If
	End If
End Function

'[Verify text for Memo Details on Notes Screen For Fee Adjustment]
Public Function verifyMemo_FA(strPlanNo,strSequenceNo,strAmount)
   Dim bverifyMemo:bverifyMemo=true

    MessageDetailsText=bcVerify_Notes.lblMemoDetailsText.GetROProperty("innertext")
    
	If Not IsNull(strPlanNo) Then
	   If Not matchStr(MessageDetailsText,strPlanNo) Then
		bverifyMemo = False
	   End If
	End If

	
	If Not IsNull(strSequenceNo) Then
	   If Not matchStr(MessageDetailsText,strSequenceNo) Then
		bverifyMemo = False
	   End If
	End If
	
	
	If Not IsNull(strAmount) Then
	   If Not matchStr(MessageDetailsText,strAmount) Then
		bverifyMemo = False
	   End If
	End If
	
	bcVerify_Notes.btnOKMessageDetails.Click
	
	verifyMemo_FA=bverifyMemo
End Function
