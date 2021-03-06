'*****This is auto generated code using code generator please Re-validate ****************
'[Verify Tab Fee Waiver is displayed]
Public Function verifyTabFeeWaiverExist()
   bDevPending=false
   verifyTabFeeWaiverExist=verifyTabExist("Fee Waiver")
End Function

'[Verify Field FeeWaiverType displayed as]
Public Function verifyFeeWaiverTypeText(strExpectedText)
   bDevPending=False
   bVerifyFeeWaiverTypeText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblFeeWaiverType(), strExpectedText, "FeeWaiverType")Then
           bVerifyFeeWaiverTypeText=false
       End If
   End If
   verifyFeeWaiverTypeText=bVerifyFeeWaiverTypeText
End Function

'[Verify Table SelectedCards on Fee Waiver SR Screen displayed]
Public Function verifySelectedCardsTabledisplayed_FW()
   bDevPending=False
   verifySelectedCardsdisplayed_FW= FeeWaiver.tblSelectedCards.Exist(1)
End Function
'[Verify Table SelectedCards has on Fee Waiver SR Screen following Columns]
Public Function verifySelectedCardsTableColumns_FW(arrColumnNameList)
   bDevPending=False
   verifySelectedCardsTableColumns_FW=verifyTableColumns(FeeWaiver.tblSelectedCards,arrColumnNameList)
End Function

'[Verify row Data in Table SelectedCards on Fee Waiver SR Screen]
Public Function verifytblSelectedCard_RowData_FW(arrRowDataList)
   bDevPending=False
   verifytblSelectedCard_RowData_FW=verifyTableContentList(FeeWaiver.tblSelectedCardsHeader,FeeWaiver.tblSelectedCardsContent,arrRowDataList,"SelectedCards", False,Null ,Null,Null)
End Function

'[Verify Table SelectedTransaction displayed on Fee Waiver SR Screen]
Public Function verifySelectedTransactionTabledisplayed_FW()
   bDevPending=False
   verifySelectedTransactiondisplayed_FW= FeeWaiver.tblSelectedTransaction.Exist(1)
End Function
'[Verify Table SelectedTransaction on Fee Waiver SR Screen has following Columns]
Public Function verifySelectedTransactionTableColumns_FW(arrColumnNameList)
   bDevPending=False
   verifySelectedTransactionTableColumns_FW=verifyTableColumns(FeeWaiver.tblSelectedTransaction,arrColumnNameList)
End Function
'[Verify row Data in Table SelectedTransaction on Fee Waiver SR Screen]
Public Function verifytblSelectedTransaction_RowData_FW(arrRowDataList)
   bDevPending=False
   verifytblSelectedTransaction_RowData_FW=verifyTableContentList(FeeWaiver.tblSelectedTransactionHeader,FeeWaiver.tblSelectedTransactionContent,arrRowDataList,"SelectedTransaction", False,Null ,Null,Null)
End Function

'[Verify Field RequestedAmount on Fee Waiver SR Screen displayed as]
Public Function verifyRequestedAmountText_FW(strExpectedText)
   bDevPending=False
   bVerifyRequestedAmountText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( FeeWaiver.txtRequestedAmount(), strExpectedText, "RequestedAmount")Then
           bVerifyRequestedAmountText=false
       End If
   End If
   verifyRequestedAmountText_FW=bVerifyRequestedAmountText
End Function

'[Set TextBox RequestedAmount on Fee Waiver SR Screen to]
Public Function setRequestedAmountTextbox_FW(strRequestedAmount)
   bDevPending=False
   FeeWaiver.txtRequestedAmount.Set(strRequestedAmount)
   If Err.Number<>0 Then
       setRequestedAmountTextbox_FW=false
            LogMessage "WARN","Verification","Failed to Set Text Box :RequestedAmount" ,false
       Exit Function
   End If
   setRequestedAmountTextbox_FW=true
End Function
'[Verify Field Description on Fee Waiver SR Screen displayed as]
Public Function verifyDescriptionText_FW(strExpectedText)
   bDevPending=False
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_FW=bVerifyDescriptionText
End Function

'[Verify Field Error Message on Fee Waiver SR Screen displayed as]
Public Function verifyErrorMessage_FW(strExpectedText)
   bDevPending=False
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblErrorMsg(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   Else
		If FeeWaiver.lblErrorMsg.Exist(1) Then
			LogMessage "RSLT","Verification","Unexpected Error message displayed",true
			bVerifyDescriptionText=false
		End If
   End If
   verifyErrorMessage_FW=bVerifyDescriptionText
End Function
'[Verify Knowledge base link is enabled on Fee Waiver SR Screen]
Public Function VerifyKnowledgebaselinkEnabled_FW()
      bDevPending=false
   Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
     strKBLink=FeeWaiver.lnkKnowledgeBase.GetROProperty("Outerhtml")
	
    If inStr(strKBLink,"v-disabled") = 0 Then
		LogMessage "RSLT","Verification","Knowledge base Link  enabled successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not enabledas expected",false
		bVerifyKnowledgebaselink=false
	End If
	VerifyKnowledgebaselinkEnabled_FW=bVerifyKnowledgebaselink
End Function

'[Verify Field KnowledgeBase on Fee Waiver SR Screen displayed as]
Public Function verifyKnowledgeBase_FW(strExpectedLink)
   bDevPending=false
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strExpectedLink) Then
		
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"
		
			'strKBLink=SpendingLimit.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
			strKBLink=FeeWaiver.lnkKnowledgeBase.GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_FW=bVerifyKnowledgeBaseText
End Function

'[Verify Field Comment on Fee Waiver SR Screen displayed as]
Public Function verifyCommentText_FW(strExpectedText)
   bDevPending=False
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( FeeWaiver.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText_FW=bVerifyCommentText
End Function

'[Set TextBox Comment on Fee Waiver SR Screen to]
Public Function setCommentTextbox_FW(strComment)
   bDevPending=False
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comment on Fee Waiver SR Screen to"
	gstrParameterNameStep = "TimeStamp"&replace((replace((replace(now,"/","-"))," ","-")),":","-")
	insertDataStore gstrParameterNameStep, strComment
	'insertDataStore "SRComment", strComment
	
   FeeWaiver.txtComment.Set(strComment )
   If Err.Number<>0 Then
       setCommentTextbox_FW=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_FW=true
End Function

'[Click Button Cancel on Fee Waiver SR Screen]
Public Function clickButtonCancel_FW()
   bDevPending=False
   FeeWaiver.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_FW=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickButtonCancel_FW=true
End Function

'[Click Button AddNotes on Fee Waiver SR Screen]
Public Function clickButtonAddNotes_FW()
   bDevPending=False
   FeeWaiver.btnAddNotes.click
   If Err.Number<>0 Then
       clickButtonAddNotes_FW=false
            LogMessage "WARN","Verification","Failed to Click Button : AddNotes" ,false
       Exit Function
   End If
   clickButtonAddNotes_FW=true
End Function

'[Perform Add Notes by clicking Add Notes Button on Fee Waiver SR Screen]
Public Function addNote_FW(strNote)
   bDevPending=false
   bVerifypopupNotes=true
	Dim bVerifypopupNotes:VerifypopupNotes=true
	
	If not isNull(strNote) Then
		FeeWaiver.btnAddNotes.click
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
	addNote_FW=bVerifypopupNotes
End Function

'[Click Button Submit on Fee Waiver SR Screen]
Public Function clickButtonSubmit_FW()
   bDevPending=False
   FeeWaiver.btnSubmit.click
   
   '*************** Capturing time stamp to open Memo for this SR by Manish
	strRunTimeTimeStamp_Step="Click Button Submit on Fee Waiver SR Screen"
	strDate="31 Oct 2017"
	strTempTime=FormatDateTime(now,4)	
 	strTimeStamp=strDate&" "&strTempTime
	insertDataStore "TimeStamp", strTimeStamp
	
   If Err.Number<>0 Then
       clickButtonSubmit_FW=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonSubmit_FW=true
End Function

'[Verify Button Submit is enabled on Replace Card Screen]
Public Function VerifybtnSubmit_FW(bEnabled)
	bDevPending=False
   Dim bVerifybtnSubmit:bVerifybtnSubmit=true
	intBtnSubmit=Instr(FeeWaiver.btnSubmit.Object.GetAttribute("disabled"),("disabled"))

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
	VerifybtnSubmit_FW=bVerifyButtonSubmit
End Function

'[Click Button AssignToTM]
Public Function clickButtonAssignToTM()
   bDevPending=False
   FeeWaiver.btnAssignToTM.click
   If Err.Number<>0 Then
       clickButtonAssignToTM=false
            LogMessage "WARN","Verification","Failed to Click Button : AssignToTM" ,false
       Exit Function
   End If
   clickButtonAssignToTM=true
End Function

'[Click Button Reject]
Public Function clickButtonReject()
   bDevPending=False
   FeeWaiver.btnReject.click
   If Err.Number<>0 Then
       clickButtonReject=false
            LogMessage "WARN","Verification","Failed to Click Button : Reject" ,false
       Exit Function
   End If
   clickButtonReject=true
End Function

'[Click Button Approve]
Public Function clickButtonApprove()
   bDevPending=False
   FeeWaiver.btnApprove.click
   If Err.Number<>0 Then
       clickButtonApprove=false
            LogMessage "WARN","Verification","Failed to Click Button : Approve" ,false
       Exit Function
   End If
   clickButtonApprove=true
End Function

'[Verify Field CreatedBy displayed as]
Public Function verifyCreatedByText(strExpectedText)
   bDevPending=False
   bVerifyCreatedByText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblCreatedBy(), strExpectedText, "CreatedBy")Then
           bVerifyCreatedByText=false
       End If
   End If
   verifyCreatedByText=bVerifyCreatedByText
End Function

'[Verify Field CreatedDate displayed as]
Public Function verifyCreatedDateText(strExpectedText)
   bDevPending=False
   bVerifyCreatedDateText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblCreatedDate(), strExpectedText, "CreatedDate")Then
           bVerifyCreatedDateText=false
       End If
   End If
   verifyCreatedDateText=bVerifyCreatedDateText
End Function

'[Verify Popup ValidationMessage exist For Fee Waiver]
Public Function verifyPopupValidation_FW(bExist)
   bDevPending=False
   bActualExist=FeeWaiver.popupValidationMessage.Exist(1)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage Exists As Expected" ,true
       verifyPopupValidation_FW=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage does not Exists As Expected" ,true
       verifyPopupValidation_FW=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage does not Exists As Expected" ,False
       verifyPopupValidation_FW=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage Still Exists" ,False
       verifyPopupValidation_FW=False
   End If
End Function

'[Click Button OK_ValidationPopup For Fee Waiver]
Public Function clickButtonOK_ValidationPopup_FW()
   bDevPending=False
   
   FeeWaiver.btnOK_ValidationPopup.click
   If Err.Number<>0 Then
       clickButtonOK_ValidationPopup_FW=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationPopup" ,false
       Exit Function
   End If
   clickButtonOK_ValidationPopup_FW=true
End Function

'[Verify Field ValidationMessage For Fee Waiver displayed as]
Public Function verifyValidationMessage_FW(strExpectedText)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
           bVerifyValidationMessageText=false
       End If
   End If
   verifyValidationMessage_FW=bVerifyValidationMessageText
End Function


'[Verify Field CardNumber_RequestSubmitted displayed as]
Public Function verifyCardNumber_RequestSubmittedText(strExpectedText)
   bDevPending=False
   bVerifyCardNumber_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblCardNumber_RequestSubmitted(), strExpectedText, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber_RequestSubmittedText=false
       End If
   End If
   verifyCardNumber_RequestSubmittedText=bVerifyCardNumber_RequestSubmittedText
End Function

'[Verify Field ProductDescription_RequestSubmitted displayed as]
Public Function verifyProductDescription_RequestSubmittedText(strExpectedText)
   bDevPending=False
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblProductDescription_RequestSubmitted(), strExpectedText, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmittedText=bVerifyProductDescription_RequestSubmittedText
End Function

'[Click Link SRNumber_RequestSubmitted]
Public Function clickLinkSRNumber_RequestSubmitted()
   bDevPending=False
   FeeWaiver.lnkSRNumber_RequestSubmitted.click
   If Err.Number<>0 Then
       clickLinkSRNumber_RequestSubmitted=false
            LogMessage "WARN","Verification","Failed to Click Link : SRNumber_RequestSubmitted" ,false
       Exit Function
   End If
   clickLinkSRNumber_RequestSubmitted=true
End Function

'[Verify Field Status_RequestSubmitted displayed as]
Public Function verifyStatus_RequestSubmittedText(strExpectedText)
   bDevPending=False
   bVerifyStatus_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (FeeWaiver.lblStatus_RequestSubmitted(), strExpectedText, "Status_RequestSubmitted")Then
           bVerifyStatus_RequestSubmittedText=false
       End If
   End If
   verifyStatus_RequestSubmittedText=bVerifyStatus_RequestSubmittedText
End Function

'[Click Button RefreshStatus]
Public Function clickButtonRefreshStatus()
   bDevPending=False
   FeeWaiver.btnRefreshStatus.click
   If Err.Number<>0 Then
       clickButtonRefreshStatus=false
            LogMessage "WARN","Verification","Failed to Click Button : RefreshStatus" ,false
       Exit Function
   End If
   clickButtonRefreshStatus=true
End Function

'[Click Button Cancel_RequestSubmitted]
Public Function clickButtonCancel_RequestSubmitted()
   bDevPending=False
   FeeWaiver.btnCancel_RequestSubmitted.click
   If Err.Number<>0 Then
       clickButtonCancel_RequestSubmitted=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel_RequestSubmitted" ,false
       Exit Function
   End If
   clickButtonCancel_RequestSubmitted=true
End Function

'[Verify Field Comment_Popup displayed as]
Public Function verifyComment_PopupText(strExpectedText)
   bDevPending=False
   bVerifyComment_PopupText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( FeeWaiver.txtComment_Popup(), strExpectedText, "Comment_Popup")Then
           bVerifyComment_PopupText=false
       End If
   End If
   verifyComment_PopupText=bVerifyComment_PopupText
End Function

'[Set TextBox Comment_Popup on Fee Waiver SR screen to]
Public Function setComment_PopupTextbox(strComment_Popup)
   bDevPending=False
   FeeWaiver.txtComment_Popup.Set(strComment_Popup)
   If Err.Number<>0 Then
       setComment_PopupTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment_Popup" ,false
       Exit Function
   End If
   setComment_PopupTextbox=true
End Function

'[Select Action Menu Waive Fee from Statement Transaction table on Statement Screen]
Public Function selectWaiveFee_Statement(lstTransactionsData)
	WaitForIcallLoading
   bDevPending=False
   bSelectWaiveFee=true
   Wait(10)
   WaitForIcallLoading
 	With bcStatements
		  bSelectWaiveFee= selectTableSubMenu(.tblStatementTransactionHeader,.tblStatementTransactionContent,lstTransactionsData,"Statement Transaction","Actions",True,.btnNext,.lnkNext,.btnPrevious,"Waive Fee",bDisabled)
	End With
	
    selectWaiveFee_Statement=bSelectWaiveFee
End Function

'[Select Action Menu Waive Fee from Unbilled Transaction table on Transaction History Screen]
Public Function selectWaiveFee_TransactionHistory(lstTransactionsData)
   bDevPending=False
   bSelectWaiveFee=true
   WaitForIcallLoading
 	With TransactionHistory
		  bSelectWaiveFee= selectTableSubMenu(.tblTransactionsHeader,.tblTransactionsContent,lstTransactionsData,"Transaction History","Actions",True,.lnkNext1_UB,.lnkNext_UB,.lnkPrevious_UB,"Waive Fee",bDisabled)
		  ' bSelectWaiveFee= selectTableSubMenu(.tblTransactionsHeader,.tblTransactions,lstTransactionsData,"Transaction History","Actions",True,.lnkNext1,.lnkNext,.lnkPrevious,"Waive Fee",bDisabled)
	End With
	
    selectWaiveFee_TransactionHistory=bSelectWaiveFee
End Function

'[Verify Field TM Approval Message on Fee Waiver Screen displayed as]
Public Function verifyTMApprovalMessage_FW(strValidationMessage)
   bDevPending=False
   bverifyTMApprovalMessage_FW=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (FeeWaiver.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bverifyTMApprovalMessage_FW=false
       End If
   End If
   FeeWaiver.btnOK_ValidationPopup.Click
   WaitForIcallLoading
   verifyTMApprovalMessage_FW=bverifyTMApprovalMessage_FW
End Function

'[Click button Close on Request Submitted Popup on Fee Waiver Screen]
Public Function clickBtnCloseFW_RequestSubmitted()
	bDevPending=false
	For iCountp = 1 To 180 Step 1
		If Not FeeWaiver.btnOK_RequestSubmitted.Exist(0.5) Then
			Wait(0.5)
		else
			FeeWaiver.btnOK_RequestSubmitted.click
			Exit for
		End if
	Next   
   If Err.Number<>0 Then
       clickBtnCloseFW_RequestSubmitted=false
            LogMessage "WARN","Verification","Failed to Click Button : Close_RequestSubmitted" ,false
       Exit Function
   End If
   WaitForICallLoading
   wait (10)
   clickBtnCloseFW_RequestSubmitted=true
End Function

'******************************************************** New functions added as a part of 1602 ***********************
'[Verify row Data in Table SummaryOfFinanceWaiver on Fee Waiver SR Screen]
Public Function verifytblSummaryOfFinanceWaiver_RowData_FW(arrRowDataList)
   bDevPending=False
   verifytblSummaryOfFinanceWaiver_RowData_FW=verifyTableContentList(FeeWaiver.tblSummaryOfFinanceWaiverHeader,FeeWaiver.tblSummaryOfFinanceWaiverContents,arrRowDataList,"Summary of Finance Waiver", False,Null ,Null,Null)
End Function

'[Verify row Data in Table SelectedWaiverDetails on Fee Waiver SR Screen]
Public Function verifytblSelectedWaiverDetails_RowData_FW(arrRowDataList)
   bDevPending=False
   verifytblSelectedWaiverDetails_RowData_FW=verifyTableContentList(FeeWaiver.tblSelectedWaiverDetailsHeader,FeeWaiver.tblSelectedWaiverDetailsContent,arrRowDataList,"Selected Waiver Details", False,Null ,Null,Null)
End Function

'[Select the Requested Amount based on the Waive Cycle]
Public Function selectRequestedAmt_WaiveCycle(lstWaiveCycle)
	selectRequestedAmt_WaiveCycle = True
	'split the columns and values
	intSize = Ubound(lstWaiveCycle) + 1
	ReDim arrCol(intSize)
	ReDim arrVal(intSize)
	For iterator = 0 To Ubound(lstWaiveCycle)  Step 1
		tempVal = split(lstWaiveCycle(iterator),":")
		arrCol(iterator) = tempVal(0)
		arrVal(iterator) = tempVal(1)
	Next	
	'Fetch the total no of rows
	Set objAllRows=getAllRows(FeeWaiver.tblSelectedWaiverDetailsContent)
	intRow=objAllRows.Count  
	Dim totalAmt
	totalAmt = 0.00
	For colLoop = 0 To intSize-1 Step 1
		'Select the checkbox only if arrval() is True
		strColName = arrCol(colLoop)
		If arrval(colLoop) = "True" Then
			'write the function to select the checkbox
			selectChk = selectCheckBox(FeeWaiver.tblSelectedWaiverDetailsHeader,strColName)
			'write the function to read the value
			strCellVal=getCellTextFor(FeeWaiver.tblSelectedWaiverDetailsHeader,objAllRows(0),rowLoop,strColName)
			If strCellVal="" Then
				strCellVal = 0
			End If
			totalAmt = totalAmt + strCellVal
		End If		
	Next	
	'now check if totalAmt is getting displayed in Requested Amount
	intActualVal = cdbl(FeeWaiver.txtRequestedAmount().GetRoProperty("value"))
	If totalAmt = intActualVal Then
		LogMessage "RSLT","Verification","Total amount matching as expected. Expected" &totalAmt& " Actual: "&intActualVal ,True
	Else
		LogMessage "WARN","Verification","Total amount not matching as expected. Expected" &totalAmt& " Actual: "&intActualVal  ,false
		selectRequestedAmt_WaiveCycle = False
	End If
	
End Function

Public Function selectCheckBox(objTableHeader,strColName)
	selectCheckBox = true
	Set oDesc = Description.Create
	oDesc("xpath").value =".//div[contains(@class,'dt-header-cell ng-scope')]"
	Set tableColumnsObj = objTableHeader.childobjects(oDesc)
	intCol=tableColumnsObj.Count
	For it = 0 To intCol-1 Step 1
		'check which childobject contains the class "csat-icon-checkbox ng-binding"
		Dim strColHeader
		strColHeader=tableColumnsObj(it).GetROProperty("innertext")
		If trim(strColHeader) = trim(strColName) Then
			'search for the checkbox and click
			Set chkBox = Description.Create
			chkBox("xpath").value = ".//div[contains(@class,'md-container md-ink-ripple')]"
			Set chkBoxChildObj = tableColumnsObj(it).childobjects(chkBox)
			countChk = chkBoxChildObj.Count
			print countChk
			chkBoxChildObj(0).click
		End If
	Next
	
End Function
