'*****This is auto generated code using code generator please Re-validate ****************

'[Select Combobox DeliveryMode on New PIN Screen as]
Public Function selectDeliveryModeComboBox(strDeliveryMode)
   bDevPending=false
   bSelectDeliveryModeComboBox=true
   If Not IsNull(strDeliveryMode) Then
       If Not (selectItem_Combobox (NewPin.lstDeliveryMode(), strDeliveryMode))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From DeliveryMode drop down list" ,false
           bSelectDeliveryModeComboBox=false
       Else
			  LogMessage "RSLT","Verification","Selected :"&strControlName&" From DeliveryMode drop down list" ,true
       End If
   End If
   selectDeliveryModeComboBox=bSelectDeliveryModeComboBox
End Function

'[Verify Combobox DeliveryMode on New PIN Screen displayed as]
Public Function verifyDeliveryModeText(strExpectedText)
   bDevPending=false
   bVerifyDeliveryModeText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (NewPin.lstDeliveryMode(), strExpectedText, "DeliveryMode")Then
           bVerifyDeliveryModeText=false
       End If
   End If
   verifyDeliveryModeText=bVerifyDeliveryModeText
End Function
'[Verify DeliveryMode Combobox has Items]
Public Function verifyDeliveryModeComboboxItems(lstItems)
   bDevPending=false
   bVerifyDeliveryModeItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (NewPin.lstDeliveryMode, lstItems, "Delivery Mode")Then
           bVerifyDeliveryModeItems=false
       End If
   End If
   verifyDeliveryModeComboboxItems=bVerifyDeliveryModeItems
End Function

'[Select Combobox DeliveryInstruction on New PIN Screen as]
Public Function selectDeliveryInstructionComboBox(strDeliveryInstruction)
   bDevPending=false
   bSelectDeliveryInstructionComboBox=true
   If Not IsNull(strDeliveryInstruction) Then
       If Not (selectItem_Combobox (NewPin.lstDeliveryInstruction(), strDeliveryInstruction))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From DeliveryInstruction drop down list" ,false
           bSelectDeliveryInstructionComboBox=false
		Else
			  LogMessage "RSLT","Verification","Selected :"&strControlName&" From DeliveryInstruction drop down list" ,true
       End If
   End If
   selectDeliveryInstructionComboBox=bSelectDeliveryInstructionComboBox
End Function

'[Verify DeliveryInstruction Combobox has Items]
Public Function verifyDeliveryInstructionComboboxItems(lstItems)
   bDevPending=false
   bVerifyDeliveryInstructionItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (NewPin.lstDeliveryInstruction, lstItems, "Delivery Mode")Then
           bVerifyDeliveryInstructionItems=false
       End If
   End If
   verifyDeliveryInstructionComboboxItems=bVerifyDeliveryInstructionItems
End Function

'[Verify Combobox DeliveryInstruction on New PIN Screen displayed as]
Public Function verifyDeliveryInstructionText(strExpectedText)
   bDevPending=false
   bVerifyDeliveryInstructionText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (NewPin.lstDeliveryInstruction(), strExpectedText, "DeliveryInstruction")Then
           bVerifyDeliveryInstructionText=false
       End If
   End If
   verifyDeliveryInstructionText=bVerifyDeliveryInstructionText
End Function

'[Verify Textbox ContactNo on New PIN Screen displayed as]
Public Function verifyContactNo_NewPin(strContactNum)
   bDevPending=false
   bVerifyContactNoText=true
   If Not IsNull(strContactNum) Then
       If Not VerifyField( NewPin.txtContactNo(), strContactNum, "ContactNo")Then
           bVerifyContactNoText=false
       End If
   End If
   verifyContactNo_NewPin=bVerifyContactNoText
End Function


'[Set TextBox ContactNo on New PIN Screen to]
Public Function setContactNoTextbox_NewPin(strContactNo)
   bDevPending=false
   If not isNull(strContactNo) Then
	   NewPin.txtContactNo.Set(strContactNo)
	   If Err.Number<>0 Then
		   setContactNoTextbox_NewPin=false
				LogMessage "WARN","Verification","Failed to Set Text Box :ContactNo" ,false
		   Exit Function
	   End If
   End If
   setContactNoTextbox_NewPin=true
End Function


'[Verify TextBox Comment on New PIN Screen displayed as]
Public Function verifyComment_NewPin(strExpectedText)
   bDevPending=false
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( NewPin.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyComment_NewPin=bVerifyCommentText
End Function


'[Set TextBox Comment on New PIN Screen to]
Public Function setCommentTextbox_NewPin(strComment)
   bDevPending=false
   If not isNull(strComment) Then
	   NewPin.txtComment.Set(strComment)
   End If
   If Err.Number<>0 Then
       setCommentTextbox_NewPin=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_NewPin=true
End Function

'[Click Button AddNotes on New PIN Screen]
Public Function clickButtonAddNotes_NewPin()
   bDevPending=false
   NewPin.btnAddNotes.click
   If Err.Number<>0 Then
       clickButtonAddNotes_NewPin=false
            LogMessage "WARN","Verification","Failed to Click Button : AddNotes" ,false
       Exit Function
   End If
   clickButtonAddNotes_NewPin=true
End Function

'[Perform Add Notes by clicking Add Notes Button]
Public Function addNote_NewPin(strNote)
   bDevPending=false
   bVerifypopupNotes=true
	Dim bVerifypopupNotes:VerifypopupNotes=true
	
	If not isNull(strNote) Then
		NewPin.btnAddNotes.click
		WaitForICallLoading
            If not   NewPin.popupValidationMessage.exist(5)Then
				LogMessage "WARN","Verification","New Note dialog did not displayed",false
				bVerifypopupNotes=false
			 else
			 strMessage=NewPin.lblMaxAllowed.GetROProperty("innerText")
				If not strMessage="Max allowed - 3000" Then
					LogMessage "WARN","Verification","Add New Comment popup dislog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage,false
					bVerifypopupNotes=false
				End If
			   NewPin.txtComment_Notes.set strNote
				NewPin.btnOK_ValidationPopup.Click
			  WaitForIcallLoading
		   End If 
		End If 
	addNote_NewPin=bVerifypopupNotes
End Function

'[Verify Button Submit on New PIN Screen is enabled]
Public Function VerifybtnSubmit_NewPIN(bEnabled)
	bDevPending=false
   Dim bVerifyButtonSubmit:bVerifyButtonSubmit=true
	intBtnSubmit=Instr(NewPin.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
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
    VerifybtnSubmit_NewPIN=bVerifyButtonSubmit
End Function

'[Click Button Submit on New PIN Screen]
Public Function clickButtonSubmit_NewPin()
   bDevPending=false
   intBtnSubmit=Instr(NewPin.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
   If  intBtnSubmit<>0 Then
		LogMessage "WARN","Verification","Submit button is disabled as per expectation.",True
		clickButtonSubmit_NewPin=true
		Exit Function
   End If
	NewPin.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_NewPin=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonSubmit_NewPin=true
End Function

'[Click Button Cancel on New PIN Screen]
Public Function clickButtonCancel_NewPin()
   bDevPending=false
   intBtnSubmit=Instr(NewPin.btnCancel.Object.GetAttribute("disabled"),("disabled"))
   If  intBtnSubmit<>0 Then
		LogMessage "WARN","Verification","Calcel button is disabled as per expectation.",True
		clickButtonCancel_NewPin=true
		Exit Function
   End If
	NewPin.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_NewPin=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonCancel_NewPin=true
End Function

'[Click Button OK on Confirmation Popup]
Public Function clickButtonOk_ConfirmationNewPin()
   bDevPending=false
   intBtnSubmit=Instr(NewPin.btnOK_Confirmation.Object.GetAttribute("disabled"),("disabled"))
   If  intBtnSubmit<>0 Then
		LogMessage "WARN","Verification","Calcel button is disabled as per expectation.",True
		clickButtonOk_ConfirmationNewPin=true
		Exit Function
   End If
	NewPin.btnOK_Confirmation.click
   If Err.Number<>0 Then
       clickButtonOk_ConfirmationNewPin=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonOk_ConfirmationNewPin=true
End Function

'[Verify Mandatory Field error Message on New PIN Screen displayed as]
Public Function verifyErrorMessage_NewPin(strExpectedMessage)
   bDevPending=false
   bVerifyErrorMessage=true
   If Not IsNull(strExpectedMessage) Then
	   If not NewPin.lblErrorMessage.Exist(1) Then
		   LogMessage "WARN","Verification","Madatory Field Valication message does not displayed",false
		   verifyErrorMessage_NewPin=false
		   Exit Function
	   End If
       If Not VerifyInnerText (NewPin.lblErrorMessage, strExpectedMessage, "Error Message")Then
           bVerifyErrorMessage=false
       End If
   End If
   verifyErrorMessage_NewPin=bVerifyErrorMessage
End Function


'[Verify Field Description on New PIN Screen displayed as]
Public Function verifyDescription_NewPin(strDescription)
   bDevPending=false
   bVerifyDescriptionText=true
   If Not IsNull(strDescription) Then
       If Not VerifyInnerText (NewPin.lblDescription(), strDescription, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescription_NewPin=bVerifyDescriptionText
End Function

'[Verify Popup Validation Message exist]
Public Function verifyPopupValidationMessageexist(bExist)
   bDevPending=false
   WaitForIcallLoading
   bActualExist=NewPin.popupValidationMessage.Exist(1)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage Exists As Expected" ,true
       verifyPopupValidationMessageexist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage does not Exists As Expected" ,true
       verifyPopupValidationMessageexist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage does not Exists As Expected" ,False
       verifyPopupValidationMessageexist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage Still Exists" ,False
       verifyPopupValidationMessageexist=False
   End If
End Function

'[Click Button OK on Validation Message Popup]
Public Function clickButtonOK_ValidationPopup()
   bDevPending=false
   NewPin.btnOK_ValidationPopup.click   
   If Err.Number<>0 Then
       clickButtonOK_ValidationPopup=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationPopup" ,false
       Exit Function
   End If
   waitForIcallLoading
   clickButtonOK_ValidationPopup=true
End Function

'[Verify ValidationMessage displayed as]
Public Function verifyValidationMessageText(strExpectedText)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (NewPin.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
           bVerifyValidationMessageText=false
       End If
   End If
   'NewPin.btnOK_ValidationPopup.Click
   verifyValidationMessageText=bVerifyValidationMessageText
End Function

'[Verify Knowledge base link is enabled on New PIN  Screen]
Public Function VerifyKnowledgebaselinkEnabled_NewPin()
      bDevPending=False
   Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
     strKBLink=NewPin.lnkKnowledgeBase.GetROProperty("Outerhtml")
	
    If inStr(strKBLink,"disabled") = 0 Then
		LogMessage "RSLT","Verification","Knowledge base Link  enabled successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not enabledas expected",false
		bVerifyKnowledgebaselink=false
	End If
	VerifyKnowledgebaselinkEnabled_NewPin=bVerifyKnowledgebaselink
End Function

'[Verify Tab New PIN displayed]
Public Function verifyTabNewPINexist()
   bDevPending=false
   verifyTabNewPINexist=verifyTabExist("New PIN")
End Function

'[Verify Table Selected Cards displayed on New PIN Screen]
Public Function verifySelectedCardsTableNewPIN()
   bDevPending=false
   bverifySelectedCardsTabledisplayed=true
   If not (NewPin.tblSelectedCardsHeader.Exist(1)) Then
		bverifySelectedCardsTabledisplayed=false
   End If
	verifySelectedCardsTableNewPIN=bverifySelectedCardsTabledisplayed
End Function

'[Verify Table Selected Cards on New PIN Screen has following Columns]
Public Function verifySelectedCardsTblColumns_NewPIN(arrColumnNameList)
   bDevPending=false
   verifySelectedCardsTblColumns_NewPIN=verifyTableColumns(NewPin.tblSelectedCardsHeader,arrColumnNameList)
End Function

'[Verify row Data in Table Selected Cards on New PIN Screen]
Public Function verifytblSelectedCards_RowData_NewPIN(lstlstSelectedCards)
   'bDevPending=false
   bverifytblSelectedCards_RowData_NewPIN = true
   verifytblSelectedCards_RowData_NewPIN=verifyTableContentList(NewPin.tblSelectedCardsHeader,NewPin.tblSelectedCardsContent,lstlstSelectedCards,"Selected Cards",False,null,null,null)
	verifytblSelectedCards_RowData_NewPIN = bverifytblSelectedCards_RowData_NewPIN
End Function


'[Verify Knowledge base link on New PIN Screen is enabled]
Public Function VerifyKnowledgebaselinkEnabled_NewPin()
      bDevPending=false
   Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
     'strKBLink=NewPin.lnkKnowledgeBase.GetROProperty("Outerhtml")
     strKBLink=NewPin.lnkKnowledgeBase.object.GetAttribute("disabled")
	
    If inStr(strKBLink,"disabled") = 0 Then
		LogMessage "RSLT","Verification","Knowledge base Link  enabled successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not enabledas expected",false
		bVerifyKnowledgebaselink=false
	End If
	VerifyKnowledgebaselinkEnabled_NewPin=bVerifyKnowledgebaselink
End Function

'[Verify Knowledge base link on New PIN Screen is disabled]
Public Function VerifyKnowledgebaselinkDisabled_NewPin()
      bDevPending=false
   Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
     strKBLink=NewPin.lnkKnowledgeBase.GetROProperty("Outerhtml")
	
    If not inStr(strKBLink,"disabled") = 0 Then
		LogMessage "RSLT","Verification","Knowledge base Link  disabled successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not disabled as expected",false
		bVerifyKnowledgebaselink=false
	End If
	VerifyKnowledgebaselinkEnabled_NewPin=bVerifyKnowledgebaselink
End Function

'[Verify Knowledge base link]
Public Function VerifyKnowledgebaselink_NewPin(strKnoledgeBase)
      bDevPending=false
   Dim bVerifyKnowledgebaselink:bVerifyKnowledgebaselink=true
   Set oDesc_KB = Description.Create()
	oDesc_KB("micclass").Value = "Link"
    strKBLink=NewPin.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
	
	If isNull(strKnoledgeBase) Then '*********************** NOT COMPLETED. Related To, Type and Sub Type have to define
		strQuery_KB="Select Distinct d.Knowledge_Base from cca_prm_sr_relto a, cca_prm_sr_type b,cca_prm_sr_subtype c,cca_prm_sr_other d where a.related_to='"&strRelatedTo&"' and a.rlt_id= b.rlt_id AND b.req_type='"&strType&"'and b.rt_id=c.rt_id and c.req_sub_type='"&strSubType&"'and d.OT_id= c.otherparameo_ot_id"
		strKnoledgeBase=getDBValForColumn(strQuery_KB)(0)
	End If
	If Trim(strKBLink)=Trim(strKnoledgeBase) Then
		LogMessage "RSLT","Verification","Knowledge base Link  "& strKBLink &" displayed successfully as expected",true
	else
		LogMessage "WARN","Verification","Knowledge base Link  does not displayed correctly. Expected : "& strKnoledgeBase &" Actual : " &strKBLink,false
		bCreateServiceRequest=false
	End If
	VerifyKnowledgebaselink_NewPin=bVerifyKnowledgebaselink
End Function


'[Verify Popup Request Submitted exist for New PIN]
Public Function verifyPopupRequestSubmitted_NewPIN(bExist)
   bDevPending=false
   bActualExist=NewPin.popupRequestSubmitted.Exist(4)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupRequestSubmitted_NewPIN=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupRequestSubmitted_NewPIN=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupRequestSubmitted_NewPIN=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupRequestSubmitted_NewPIN=False
   End If
End Function

'[Verify Field CardNumber on Request Submitted Popup for New Pin displayed as]
Public Function verifyCardNumber_RequestSubmitted_NewPIN(strCardNumber)
   bDevPending=false
   bVerifyCardNumber_RequestSubmittedText=true
   insertDataStore "NewPINUsedCard", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       If Not VerifyInnerText (NewPin.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber_RequestSubmittedText=false
       End If
   End If
   verifyCardNumber_RequestSubmitted_NewPIN=bVerifyCardNumber_RequestSubmittedText
End Function

'[Verify Field ProductDescription on Request Submitted Popup for New Pin displayed as]
Public Function verifyProductDescription_RequestSubmitted_NewPin(strProductDescription)
   bDevPending=false
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (NewPin.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_NewPin=bVerifyProductDescription_RequestSubmittedText
End Function

'[Click Link SRNumber on Request Submitted Popup for New Pin]
Public Function clickLinkSRNumber_RequestSubmitted_NewPIN()
   bDevPending=false
   strSelectedSR=NewPin.lnkSRNumber_RequestSubmitted.GetRoProperty("innerText")
	If strSelectedSR<>"" Then
		 insertDataStore "SelectedSRLink", strSelectedSR
	
	   NewPin.lnkSRNumber_RequestSubmitted.click
	 else
   		LogMessage "RSLT","Verification","SR Number did not displayed on Request Submitted pop up",false
	End If
   WaitForIcallLoading
   If Err.Number<>0 Then
       clickLinkSRNumber_RequestSubmitted_NewPIN=false
            LogMessage "WARN","Verification","Failed to Click Link : SRNumber_RequestSubmitted" ,false
       Exit Function
   End If
   clickLinkSRNumber_RequestSubmitted_NewPIN=true
End Function

'[Verify Field Status_RequestSubmitted For New Pin displayed as]
Public Function verifyStatus_RequestSubmittedINewPIN(strExpectedText)
   bDevPending=false
   bVerifyStatus_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (NewPin.lblStatus_RequestSubmitted(), strExpectedText, "Status_RequestSubmitted")Then
           bVerifyStatus_RequestSubmittedText=false
       End If
   End If
   verifyStatus_RequestSubmittedINewPIN=bVerifyStatus_RequestSubmittedText
End Function

'[Click Button RefreshStatus For New Pin]
Public Function clickButtonRefreshStatus_NewPIN()
   bDevPending=false
   NewPin.btnRefreshStatus.click
	WaitForICallLoading
    		'Get Status
		If NewPin.lblStatus_RequestSubmitted.getROProperty("innertext")="In Progress" then 
			bStatus=true
		 else
			bStatus=false
		End If
		
	
	While  bStatus AND (iCount<60)
		NewPin.btnRefreshStatus.click
		wait 1
        	'Get Status
			strStatus=NewPin.lblStatus_RequestSubmitted.getROProperty("innertext")
			If Trim(strStatus)="In Progress" then 
				bStatus=true
			 else
				LogMessage "WARN","Verification","Status displayed as  :"&strStatus ,true
				bStatus=false
			End If
		wait 5
		intBtnRefreshStatus=Instr(NewPin.btnRefreshStatus.Object.GetAttribute("disabled"),("disabled"))
		If intBtnRefreshStatus<>0 Then
			LogMessage "WARN","Verification","Button : RefreshStatus is disabled" ,true
			bStatust=true
			If  NewPin.lblStatus_RequestSubmitted.getROProperty("innertext")="In Progress" Then
				bStatust=false
				LogMessage "RSLT","Verification","Status is In Progress but Refresh Status button disabled",false
			End If
		End If
		iCount=iCount+1
	  Wend	

   If Err.Number<>0 Then
       
            LogMessage "WARN","Verification","Failed to Click Button : RefreshStatus" ,false
			clickButtonRefreshStatus_NewPIN=false
       Exit Function
   End If
   WaitForICallLoading
  
   clickButtonRefreshStatus_NewPIN=true
End Function

'[Click Button Cancel_Request Submitted for New Pin]
Public Function clickButtonCancel_RequestSubmitted()
   bDevPending=false
   NewPin.btnCancel_RequestSubmitted.click
   waitForIcallLoading
   If Err.Number<>0 Then
       clickButtonCancel_RequestSubmitted=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel_RequestSubmitted" ,false
       Exit Function
   End If
   clickButtonCancel_RequestSubmitted=true
End Function

'[Verify Field SRNumber displayed on View SR for New Pin as]
Public Function verifySRNumber_NewPin(strSRNumber)
   bDevPending=false
   bVerifySRNumberText=true
    ' If SR link clicked from popup lable link
   If Ucase(strSRNumber)="SELECTED SR LINK" Then
		strSRNumber=fetchFromDataStore("User click Unknown SR Number form Service Request Tab","BLANK","SelectedSRLink")(0)
   End If
   If Not IsNull(strSRNumber) Then
       If Not VerifyInnerText (ViewSR.lblSRNumber(), strSRNumber, "SRNumber")Then
           bVerifySRNumberText=false
       End If
   End If
   verifySRNumber_NewPin=bVerifySRNumberText
End Function

'[Perform Additional Verification for TPIN call from New PIN]
Public Function performTPINPlusOne_NewPIN(bExpectedVerification, strValidationMessage)
   Dim bPerformTPINPlusOneVerification:bPerformTPINPlusOneVerification=true
   bActualExist=NewPin.popupValidationMessage.Exist(1)
	'bActualExist=true	
   If bExpectedVerification Then
		If not bActualExist Then
			bPerformTPINPlusOneVerification=false
		 else
			strActualValidationMessage=NewPin.lblValidationMessage.GetRoProperty("innertext")
			If Not IsNull(strActualValidationMessage) Then
				If Not VerifyInnerText (NewPin.lblValidationMessage(), strValidationMessage, "ValidationMessage")Then
					bPerformTPINPlusOneVerification=false
				End If
			End If
			NewPin.btnOK_ValidationPopup.Click
			VerifyCustomer.btnVerify.Click
			waitForIcallLoading
			
			'Check Required Verification met Check box
			Set oDesc= Description.Create()
			oDesc("micclass").Value="WebCheckBox"
			'CardActivation.cbTPINVerification.ChildObjects(oDesc)(0).Set "ON" '******** Check box removed
			CardActivation.rbAdditionalAnswer.click
			'Wait 1
			VerifyCustomer.btnVerifyCustomer().click
			waitForIcallLoading
			If err.number<>0 Then
				bPerformTPINPlusOneVerification=false
			End If
		End If	 
   End If
   performTPINPlusOne_NewPIN=bPerformTPINPlusOneVerification
End Function

'[Click Close On Request Submit Button FOR NewPin]
Public Function clickCloseNewPin()
   bDevPending=false
   NewPin.btnCancel_RequestSubmitted.click
   If Err.Number<>0 Then
       clickCloseNewPin=false
            LogMessage "WARN","Verification","Failed to Click Button : Temporary Limit Increase" ,false
       Exit Function
   End If
   waitForIcallLoading
   clickCloseNewPin=true
End Function

'[Verify Field InLine Message on New PIN Screen displayed as]
Public Function verifyInLineMessage_NewPin(strInLineMessage)
   bDevPending=false
   bverifyInLineMessage_NewPin=true
   If Not IsNull(strInLineMessage) Then
       If Not VerifyInnerText (NewPin.lblInLineMessage(), strInLineMessage, "InLine Message")Then
           bverifyInLineMessage_NewPin=false
       End If
   End If
   verifyInLineMessage_NewPin=bverifyInLineMessage_NewPin
End Function

'**********************1601 added new radio button***********************
'[Select Radio Button Customer agree to reveal contact number on New Pin Screen]
Public Function selectCustomerAgreeRevealConctNumberNewPin(strCustomerAgree)
	bDevPending=false
	bselectCustomerAgreeRevealConctNumberNewPin=true
	bselectCustomerAgreeRevealConctNumberNewPin=SelectRadioButtonGrp(strCustomerAgree, NewPin.rbtnNewPinCustomerAgreeRevealConctNo, Array("Yes","No"))
	WaitForICallLoading
	If not IsNull (strCustomerAgree) Then
		Select Case strCustomerAgree
			Case "Yes"
                'strPostalCode=CardReplacement.txtPostalCode.GetROProperty("Outerhtml")
                strContactNumber=NewPin.txtContactNo.Object.GetAttribute("disabled")
                strcompare=inStr(strContactNumber,"disabled") 
              
                Print "Value is "&strcompare
                If inStr(strContactNumber,"disabled") = 0 Then
                'If not Ucase(Trim(strContactNumber))="" Then
					 LogMessage "RSLT","Verification","Contact Number is enabled as expected",true
					 bselectCustomerAgreeRevealConctNumberNewPin=true
				Else
					LogMessage "WARN","Verification","Contact Number is not enabled",false
					bselectCustomerAgreeRevealConctNumberNewPin=false						
				 End If
				 'strContactNumber=NewPin.txtContactNumber.GetROProperty("value")
				 'If  strContactNumber = NA Then
					'LogMessage "WARN","Verification","Contact Number field value displaying incorrectly ",false
					'bselectCustomerAgreeRevealConctNumberNewPin=false
				'End If
				
			Case "No"
            	'strContactNumber=NewPin.txtContactNumber.GetROProperty("value")
            	strContactNumber=CardReplacement.txtContactNumber.Object.GetAttribute("disabled")
				 If inStr(strContactNumber,"disabled") = 0 Then
					 LogMessage "WARN","Verification","Contact Number is not disabled",false
					 bselectCustomerAgreeRevealConctNumberNewPin=false
				Else
					LogMessage "RSLT","Verification","Contact Number is disabled as expected",true
					bselectCustomerAgreeRevealConctNumberNewPin=true					
					 
				 End If
				 'If  Not strContactNumber = NA Then
					'LogMessage "WARN","Verification","Contact Number field value displaying incorrectly ",false
					'bselectCustomerAgreeRevealConctNumberNewPin=false
				'End If
		End Select
	End If
	If Err.Number<>0 Then
	   bselectCustomerAgreeRevealConctNumberNewPin=false
		  LogMessage "WARN","Verification","Failed to Click Button : CustomerAgreeRevealConctNumber" ,false
	   Exit Function
   End If
   selectCustomerAgreeRevealConctNumberNewPin=bselectCustomerAgreeRevealConctNumberNewPin
End Function

