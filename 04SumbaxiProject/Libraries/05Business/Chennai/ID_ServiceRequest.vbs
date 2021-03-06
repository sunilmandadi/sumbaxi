'################################# FUNCTIONS FOR [NEW SR] STARTS ########################################################


'[Perfrom Customer Verification from overview page]
Public Function VerifyCustomerFromOverViewPage(iNoOfIdentQues,iNoOfAuthQues)
	IndonesiaCustomerVerification iNoOfIdentQues,iNoOfAuthQues
End Function

'[Verify fields displayed in New SR Grey panel section]
Public Function verifyNewSRGreyPanel(arrLblValPairs)
	verifyNewSRGreyPanel = VerifyIDLabelValuePairs(coServiceRequest_Page.lblNewSRHeader,arrLblValPairs,"Service Request","Grey Panel")
End Function

'[Verify Default Status button selected in New SR Page]
Public Function verifySRStatusButton(strExpStatus)
	verifySRStatusButton = getDefaultSelectedRadioButton(coServiceRequest_Page.eleStatusRdioGrp,strExpStatus,"Status")
End Function

'[Verify Default value displayed in Sub Status dropdown in New SR Page]
Public Function verifySRSubStatusValue(strExpSubStatus)
	verifySRSubStatusValue = verifyFieldValue(coServiceRequest_Page.txtSRSubStatus,strExpSubStatus,"Sub Status")
End Function

'[Verify Default Priority button selected in New SR Page]
Public Function verifySRPriorityButton(strExpStatus)
	verifySRPriorityButton = getDefaultSelectedRadioButton(coServiceRequest_Page.elePriorityRdioGrp,strExpStatus,"Priority")
End Function

'[Verify Default Followup Required button selected in New SR Page]
Public Function verifySRFollowupRequiredButton(strExpStatus)
	verifySRFollowupRequiredButton = getDefaultSelectedRadioButton(coServiceRequest_Page.eleFollowUpRdioGrp,strExpStatus,"Followup Required")
End Function

'[Verify Default display of Go button in New SR Page]
Public Function VerifyButtondisplayGO_SR(strCheckFlag)
	VerifyButtondisplayGO_SR = VerifyObjectDisabled(coServiceRequest_Page.btnGOSR,strCheckFlag,"Go Button")
End Function

'[Verify Default Manual Verification message in New SR Page]
Public Function verifyManualVerificationMsg_SR(strMsg)
	verifyManualVerificationMsg_SR =  VerifyInnerText(coServiceRequest_Page.eleVerifcationMsg(),strMsg,"Manual Verification message")
End Function

'[Click GO Button in New SR Page]
Public Function ClickButtonGO_SR()
	coServiceRequest_Page.btnGOSR.click
	WaitForIServeLoading
	If Err.Number <> 0 Then
		ClickButtonGO_SR = False
		LogMessage "WARN","Verification","Failed to Click Button: GO", False
		Exit Function
	End If
	ClickButtonGO_SR = True
End Function

'[Verify ID Customer verification message]
Public Function Customer_VerificationMsg(strExpctdMsg)
	Wait(2)
	strActMsg = Trim(coServiceRequest_Page.eleVerifcationMsg.GetROProperty("innertext"))
	If strActMsg = strExpctdMsg Then
		Customer_VerificationMsg = True
		LogMessage "RSLT","Verification","Customer verification alert message displayed as expected", True
	Else
		Customer_VerificationMsg = False
		LogMessage "WARN","Verification","Customer verification alert message not displayed as expected", False	
	End If
	coServiceRequest_Page.btnOKVerifcationMsg.Click
End Function

'[Click OK Button in verification Message popup]
Public Function ClickButtonOK_SR_VerificationMsg()
	coServiceRequest_Page.btnOKVerifcationMsg.click
	Wait(2)
	If Err.Number <> 0 Then
	  ClickButtonOK_SR_VerificationMsg = False
	  LogMessage "WARN","Verification","Failed to Click Button: OK", False
	  Exit Function
	End If
	ClickButtonOK_SR_VerificationMsg = True
End Function

'[Verify Default display of Knowledge base link in New SR Page]
Public Function VerifyKnowledgeBase_SR(strCheckFlag)
	VerifyKnowledgeBase_SR = VerifyObjectDisabled(coServiceRequest_Page.lnkKnowledgeBase,strCheckFlag,"Knowledge Base Link")
End Function

'[Verify Default display of Manager Approval YesNo button in New SR Page]
Public Function verifyManagerApprovalButtonState_SR()
	blnButton = False
	Set oDesc = Description.Create
	oDesc("html tag").value = "md-radio-button"
	Set oChild = coServiceRequest_Page.eleManagerApprovalRdioGrp.ChildObjects(oDesc)
	iCount = oChild.Count-1
	For i = 0 To iCount Step 1
		intClass = oChild(i).GetROProperty("class")
		If Instr(1,intClass,"disabled") > 0 Then
			blnButton = True
		End If
	Next
	
	If blnButton Then
		LogMessage "RSLT","Verification","As Expected By Default: Both Yes and No button is disabled under Manager Approval", True
	Else
		LogMessage "WARN","Verification","Failed: By Default: Both Yes and No button is not disabled under Manager Approval", False
	End If
	verifyManagerApprovalButtonState_SR = blnButton
	Set oChild = Nothing
	Set oDesc =	Nothing
End Function

'[Click SR Status button in New SR Page]
Public Function ClickSRStatusBtn(strButtonName)
	ClickSRStatusBtn = ClickSingleRadionButton(coServiceRequest_Page.eleStatusRdioGrp,strButtonName)
End Function

'[Click Manager Approval YesNo button in New SR Page]
Public Function ClickYesNoMangerApprovalButton(strButtonName)
	'bClick = False
	WaitForIServeLoading
	'bClick = verifyManagerApprovalButtonState_SR
	'If bClick Then
		ClickYesNoMangerApprovalButton = ClickSingleRadionButton(coServiceRequest_Page.eleManagerApprovalRdioGrp,strButtonName)
	'End If
End Function

'[Click Customer Status button in New SR Page]
Public Function ClickCustomerStatusBtn(strButtonName)
	ClickCustomerStatusBtn = ClickSingleRadionButton(coServiceRequest_Page.eleCustomerStatusRdioGrp,strButtonName)
End Function

'[Click on Submit Button in New SR Page]
Public Function clickButtonSubmitViewPage_SR()
	coServiceRequest_Page.btnSubmitSR.click 
	If Err.Number <> 0 Then
	  clickButtonSubmitViewPage_SR = False
	  LogMessage "WARN","Verification","Failed to Click Button: Submit", False
	  Exit Function
	End If
	WaitForIServeLoading
	clickButtonSubmitViewPage_SR = True
End Function

'[Enter Comments textbox in New SR Page]
Public Function SetComments_SR(strComment)
	bVerifytext = True
	coServiceRequest_Page.txtComments().Set StrComment
	If Err.Number <> 0 Then
	  bVerifytext = False
	  LogMessage "WARN","Verification","Failed to Set Comments in text box", False
	  Exit Function
	End If
	SetComments_SR = bVerifytext
End Function

'[Enter Approver comments textbox in SR Page]
Public Function SetApproverComments_SR(strComment)
	bVerifytext = True
	coServiceRequest_Page.txtApproveRejectComments().Set StrComment
	If Err.Number <> 0 Then
	  bVerifytext = False
	  LogMessage "WARN","Verification","Failed to Set Comments in text box", False
	  Exit Function
	End If
	SetApproverComments_SR = bVerifytext
End Function

'[Click on Approve OR Reject Button in SR Page]
Public Function clickButtonApproveReject_SR(sButtonFlag)
	If sButtonFlag = "Approve" Then
		clickButtonApproveReject_SR = clickButtonApprove_SR
	ElseIf  sButtonFlag = "Reject" Then
		clickButtonApproveReject_SR = clickButtonReject_SR
	End If
End Function

'[Click on Approve Button in SR Page]
Public Function clickButtonApprove_SR()
	coServiceRequest_Page.btnSRApprove.click 
	If Err.Number <> 0 Then
	  clickButtonApprove_SR = False
	  LogMessage "WARN","Verification","Failed to Click Button: Approve", False
	  Exit Function
	End If
	WaitForIServeLoading
	clickButtonApprove_SR = True
End Function

'[Click on Reject Button in SR Page]
Public Function clickButtonReject_SR()
	coServiceRequest_Page.btnSRReject.click 
	If Err.Number <> 0 Then
	  clickButtonReject_SR = False
	  LogMessage "WARN","Verification","Failed to Click Button: Reject", False
	  Exit Function
	End If
	WaitForIServeLoading
	clickButtonReject_SR = True
End Function

'[Verify SR submission messsage]
Public Function VerifySRSubmissionMessage(strMsg)
	VerifySRSubmissionMessage = VerifyInnerText(coServiceRequest_Page.eleSRSubmissionPopUpMsg,strMsg,"SR Submission message")
End Function

'[Click OK button in Submission Message popup in SR Page]
Public Function ClickButtonOK_SR()
	coServiceRequest_Page.btnOKPopUpMsg.click 
	If Err.Number <> 0 Then
	  ClickButtonOK_SR = False
	  LogMessage "WARN","Verification","Failed to Click Button: OK", False
	  Exit Function
	End If
	WaitForIServeLoading
	ClickButtonOK_SR = True
End Function

Public Function ClickSingleRadionButton(objRadioGrp,strButtonName)
	Setting.WebPackage("ReplayType") = 2
	scrollPageDown 5
	Set oDesc = Description.Create
	oDesc("html tag").value = "md-radio-button"
	Set oChild = objRadioGrp.ChildObjects(oDesc)
	iCount = oChild.Count-1
	For i = 0 To iCount Step 1
		intBtnName = Trim(oChild(i).GetROProperty("innertext"))
		If intBtnName = strButtonName Then
			ClickSingleRadionButton = ClickOnObject(oChild(i),"Radio Button : " & strButtonName)
		End If
	Next
	Set oDesc = Nothing
	Setting.WebPackage("ReplayType") = 1
End Function

'[Verify SR Status and Substatus]
Public Function verifyStatusSubStatusAfterApprovalOrRejection(sStatus,sSubStatus)
	bStatus = VerifyInnerText(coServiceRequest_Page.eleSRStatus,sStatus,"SR Status")
	bSubStatus = VerifyInnerText(coServiceRequest_Page.eleSRSubStatus,sSubStatus,"SR Sub Status")
	If bStatus And bSubStatus Then
		verifyStatusSubStatusAfterApprovalOrRejection = True
	Else
		verifyStatusSubStatusAfterApprovalOrRejection = False
	End If
End Function

'[Verify display of Add Attachment button in New SR Page]
Public Function VerifyButtonAddAttachment_SR(strCheckFlag)
	VerifyButtonAddAttachment_SR = VerifyObjectEnabledDisabled(coDashboard_IA_Page.btnAddAttachments,strCheckFlag,"Add Attachment Button")
End Function

'[Verify text displayed below Attachment section in New SR Page]
Public Function VerifyAttachmentText_SR(StrExpText1,StrExpText2,StrExpText3)
	bverifyAttachmenttext = False 
	StrActText1 = Trim(coDashboard_IA_Page.txtAttachment1.GetROProperty("innertext"))
	StrActText2 = Trim(coDashboard_IA_Page.txtAttachment2.GetROProperty("innertext"))
	StrActText3 = Trim(coDashboard_IA_Page.txtAttachment3.GetROProperty("innertext"))

	If (StrExpText1 = StrActText1) AND (StrExpText2 = StrActText2) AND (StrExpText3 = StrActText3) Then
	   LogMessage "RSLT","Verification","Text displayed below the Attachment section is displayed as expected",True
	   bverifyAttachmenttext = True
	Else 
	   LogMessage "WARN","Verification","Text displayed below the Attachment section is not displayed as expected", False
	End If
	VerifyAttachmentText_SR  = bverifyAttachmenttext
End Function

'[Verify display of OnceDone checkbox in New SR Page]
Public Function VerifyOnceDoneCheckbox_SR(strCheckFlag)
	VerifyOnceDoneCheckbox_SR = VerifyObjectDisabled(coServiceRequest_Page.chkOnceAndDone,strCheckFlag,"Once and Done checkbox")
End Function

'[Verify checkbox OnceDone checked or Unchecked in SR Page]
Public Function VerifyOnceDoneCheckboxChecked_SR(strCheckFlag)
	VerifyOnceDoneCheckboxChecked_SR = VerifyObjectCheckedUnchecked(coServiceRequest_Page.chkOnceAndDone,strCheckFlag,"Once and Done checkbox")
End Function

'[Verify display of Submit Button in New SR Page]
Public Function VerifyButtondisplaySubmit_SR(strCheckFlag)
	VerifyButtondisplaySubmit_SR = VerifyObjectDisabled(coServiceRequest_Page.btnSubmitSR,strCheckFlag,"Submit Button")
End Function

'[Click on Cancel Button in New SR page]
Public Function clickCancelButton_SR()
	coServiceRequest_Page.btnCancelSR.click 
	If Err.Number <> 0 Then
		clickCancelButton_SR = False
		LogMessage "WARN","Verification","Failed to Click Button: Cancel", False
		Exit Function
	End If
	WaitForIServeLoading
	clickCancelButton_SR = True
End Function

'[Verify the Cancel Confirmation message in SR Page]
Public Function VerifyCancelMessage_SR(strCancelMsg)
	bVerifyCancelMessage = True
	If Not IsNull(strCancelMsg) Then
	   If Not verifyInnerText(coServiceRequest_Page.lblCancelMsgSR,strCancelMsg,"Cancel Message") Then
			bVerifyCancelMessage = False
		End If
	End If
	VerifyCancelMessage_SR = bVerifyCancelMessage
End Function

'[Click Yes Or No Button in cancel Message displayed in New SR]
Public Function SelectButtonCancel_SR(strSelect)
	bClickButton = True
	WaitForIServeLoading
	If Not IsNull(strSelect) Then
	   If strSelect = "YES" Then
	   	 coServiceRequest_Page.btnYesSR.click
	   ElseIf strSelect = "NO" Then
	     coServiceRequest_Page.btnNOSR.click
	   End If 	   
	   If Err.Number<>0 Then
		   LogMessage "WARN","Verification","Failed to Click Button YES/NO", False
		   bClickButton = False
	   End If
	End If
	WaitForIServeLoading
	SelectButtonCancel_SR = bClickButton
End Function

'[Attach one file to New SR by Clicking Attachment Button]
Public Function AddAttachments_SR(strFileName)
	bverifyAddedAttachment = False
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 1000")
	Wait 2
	Setting.WebPackage("ReplayType") = 2
	
	coDashboard_IA_Page.btnAddAttachments.click   
	
	If Err.Number <> 0 Then
		LogMessage "WARN","Verification","Failed to Click Button : NEW SR_Add Attachment", False
		Exit Function
	End If  
	
	'Get the gstrAttachmentsPath folder path from the OBTAF_Config
	coDashboard_IA_Page.txtFileName.Set gstrAttachmentsPath + "\" + strFileName
	
	coDashboard_IA_Page.btnOpen.Click
	
	Setting.WebPackage("ReplayType") = 1   
	
	strAddedFileName = coDashboard_IA_Page.lblFileName.GetROProperty("innertext")  
	If Len(strFileName)<=75 Then
		If Trim(strFileName) = Trim(strAddedFileName) Then
			LogMessage "RSLT","Verification","File Added is displayed in SR Page as expected",True
			bverifyAddedAttachment = True
		Else 
			LogMessage "WARN","Verification","File Added is not displayed in SR Page", False
			bverifyAddedAttachment = False
		End If
	Else
  	 bverifyAddedAttachment = True
	End If
	AddAttachments_SR = bverifyAddedAttachment
End Function

'[Verify Inline error message displayed related to Attachments added in New SR]
Public Function VerifyAttachmentInlineMsg_SR(strErrorMsg) 
	bverifyInlineErrorMsg = False
	If VerifyInnerText(coDashboard_IA_Page.lblInlineErrortxt(), strErrorMsg, "Invalid Attachment Error") Then
	   bverifyInlineErrorMsg = True
	End If
	VerifyAttachmentInlineMsg_SR = bverifyInlineErrorMsg
End Function

'[Verify CreatedBy and CreatedOn displayed for Added Attachments in New SR Page]
Public Function VerifyAttachmentCreatedInfo_SR(strCreatedBy, strCreatedOn)
	bVerifyCreatedInfo = False
	strActCreatedBy = coDashboard_IA_Page.lblCreatedBy.GetROProperty("innertext")
	If Ucase(Trim(strActCreatedBy)) = Ucase(Trim(strCreatedBy)) Then
	   LogMessage "RSLT","Verification","CreatedBy "&strCreatedBy&" displayed as expected in Attachments section",True
	   bVerifyCreatedInfo = True
	   If Not IsNull(strCreatedOn) Then
	   	  strActCreatedDate = coDashboard_IA_Page.lblCreatedOn.GetROProperty("innertext")
	   	  If strActCreatedDate = strCreatedOn  Then
	   	  	 LogMessage "WARN","Verification","CreatedOn "&strCreatedOn&" displayed as expected in Attachments section",True
	         bVerifyCreatedInfo = True
	      Else 
	   	  	 LogMessage "WARN","Verification","CreatedOn "&strCreatedOn&" displayed as expected in Attachments section",False
	         bVerifyCreatedInfo = False			         
	   	  End If
	   End If
	Else
	   LogMessage "WARN","Verification","CreatedBy "&strCreatedBy&" not displayed in Attachments section as expected",False
	End If
	VerifyAttachmentCreatedInfo_SR = bVerifyCreatedInfo
End Function

'[Verify display of description field in New SR Page]
Public Function VerifyFieldDescription_SR(strCheckFlag)
	VerifyFieldDescription_SR = VerifyObjectDisabled(coDashboard_IA_Page.txtAttachmentComment,strCheckFlag,"Description or Comments Field")
End Function

'[Click button Remove Attachment in SR Page]
Public Function clickButtonRemoveAttachment_SR()
	coDashboard_IA_Page.btnRemoveAttachment.click 
	If Err.Number <> 0 Then
	  clickButtonRemoveAttachment_SR = False
	  LogMessage "WARN","Verification","Failed to Click Button: Remove Attachment in New SR page", False
	  Exit Function
	End If
	clickButtonRemoveAttachment_SR = True
End Function

'[Select triplet RelatedTo_Type_SubType in New SR Page]
Public Function selectTripletIn_SR(strTriplet)
	bRelatedTo = SetValue(coServiceRequest_Page.lstRelatedTo,strTriplet(0),"Related To")
	bType = SetValue(coServiceRequest_Page.lstType,strTriplet(1),"Type")
	bSubType = SetValue(coServiceRequest_Page.lstSubType,strTriplet(2),"Sub Type")
	If bRelatedTo And bType And  bSubType Then
		selectTripletIn_SR = True
	Else
		selectTripletIn_SR = False
	End If
End Function

'[Select Triplet RelatedTo_Type_SubType in ID New SR Page]
Public Function selectTriplet_SR(strTriplet)
	bRelatedTo = SelectComboBoxItem(coServiceRequest_Page.lstSRRelatedTo,Trim(strTriplet(0)),"Related To")
	bType = SelectComboBoxItem(coServiceRequest_Page.lstSRType,Trim(strTriplet(1)),"Type")
	bSubType = SelectComboBoxItem(coServiceRequest_Page.lstSRSubType,Trim(strTriplet(2)),"Sub Type")
	If bRelatedTo And bType And  bSubType Then
		selectTriplet_SR = True
	Else
		selectTriplet_SR = False
	End If
End Function


'[Verify Created By and Created On text displayed in View SR Page]
Public Function VerifyCreatedInfo_ViewSR(strCreatedBy)
	bVerifyCreatedInfo = False
	If Not IsNull(strCreatedBy) Then
	   strActCreatedBy = coServiceRequest_Page.lblIACreatedBy.GetROProperty("innertext")
	End If

	strActCreatedDate = coServiceRequest_Page.lblIACreatedDate.GetROProperty("innertext")	
	strActCreatedInfo = "Created By "&UCase(Trim(strActCreatedBy))&" on "
	strExpCreatedInfo = "Created By "&UCase(Trim(strCreatedBy))&" on "
	
	If Trim(strActCreatedInfo) = Trim(strExpCreatedInfo) AND Not IsNull(strActCreatedDate) Then
		LogMessage "RSLT","Verification","Created By :"&strCreatedBy&" and Created Date is displayed as expected",True
		bVerifyCreatedInfo = True
	End If
	
	VerifyCreatedInfo_ViewSR = bVerifyCreatedInfo
End Function

'[Verify fields displayed in View SR Grey panel section]
Public Function verifyGreyPanel_ViewSR(arrLblValPairs)
	verifyGreyPanel_ViewSR = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleGreyPanel,arrLblValPairs,"Service Request","Grey Panel")
End Function

'[Verify fields displayed in View SR CustomerInfo section]
Public Function verifyCustomerInfo_ViewSR(arrLblValPairs)
	verifyCustomerInfo_ViewSR = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleCustomerInfo,arrLblValPairs,"Service Request","Customer Information")
End Function

'[Verify list of fields displayed in View SR Page]
Public Function VerifyFields_ViewSR(arrLblValPairs)
	scrollPageDown 5
	VerifyFields_ViewSR = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.lblViewSR,arrLblValPairs,"Service Request","View SR Page")
End Function

'[Verify fields below Additional Info displayed in View SR Page]
Public Function VerifyAdditionalInfo_ViewSR(arrLblValPairs)
	scrollPageDown 5
	VerifyAdditionalInfo_ViewSR = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleAdditionalInfo,arrLblValPairs,"Service Request","View SR AdditionalInfo")
End Function


'################################# FUNCTIONS FOR [NEW SR] ENDS ########################################################




'################################# FUNCTIONS FOR [VIEW SR] STARTS ########################################################

'[Navigate to SR page in Dashboard]
Public Function NavigateToDashboardSRPage()
	NavigateToDashboardSRPage = ClickTab("SERVICE REQUEST")
End Function

'[Verify existenence of fields displayed in SR Search panel]
Public Function VerifyFieldExistenece_SR()
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblStatusSR(),"Dashboard SR","Status")
		If Not bVerifyfields Then
		   LogMessage "WARN","Verification","Field label: Status not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblGroupSR(),"Dashboard SR","Group")
		If Not bVerifyfields  Then
		   LogMessage "WARN","Verification","Field label: Group not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblSelectSR(),"Dashboard SR","Select")
		If Not bVerifyfields  Then
		   LogMessage "WARN","Verification","Field label: Select not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblStaffSR(),"Dashboard SR","Staff")
		If Not bVerifyfields  Then
		   LogMessage "WARN","Verification","Field label: Staff not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblChannelSR(),"Dashboard SR","Channel")
		If Not bVerifyfields  Then
		   LogMessage "WARN","Verification","Field label: Channel not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblFromDateSR(),"Dashboard SR","From Date")
		If Not bVerifyfields Then
		   LogMessage "WARN","Verification","Field label: From Date not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		bVerifyfields = VerifyFieldExistenceInPage(coServiceRequest_Page.lblToDateSR(),"Dashboard SR","To Date")
		If Not bVerifyfields Then
		   LogMessage "WARN","Verification","Field label: To Date not displayed as expected", False
		   VerifyFieldExistenece_SR = bVerifyfields
		End If
		
		VerifyFieldExistenece_SR =  bVerifyfields
End Function


'[Verify default value displayed in the dropdowns of Dashboard SR]
Public Function VerifyDefaultDropdown_SR(StrStatus,StrGroup,StrSelect,StrStaff,StrChannel)
	bVerifyDefaultvalue = False
	If Not IsNull(StrStatus) Then
		bVerifyDefaultvalue = VerifyDropdownDefaultValue(coServiceRequest_Page.txtComboBoxStatusSR,StrStatus,"Status")
	End If
	If Not IsNull(StrGroup) Then
		bVerifyDefaultvalue = VerifyDropdownDefaultValue(coServiceRequest_Page.txtComboBoxGroupSR,StrGroup,"Group")
	End If
	If Not IsNull(StrSelect) Then
		bVerifyDefaultvalue = VerifyDropdownDefaultValue(coServiceRequest_Page.txtComboBoxSelectSR,StrSelect,"Select")
	End If
	If Not IsNull(StrStaff) Then
		bVerifyDefaultvalue = VerifyDropdownDefaultValue(coServiceRequest_Page.txtComboBoxStaffSR,StrStaff,"Staff")
	End If
	If Not IsNull(StrChannel) Then
		bVerifyDefaultvalue = VerifyDropdownDefaultValue(coServiceRequest_Page.txtComboBoxChannelSR,StrChannel,"Channel")
	End If 
	VerifyDefaultDropdown_SR = bVerifyDefaultvalue	
End Function

'[Select Combobox Status in Dashboard SR Search Panel]
Public Function SelectStatuscombobox_DashboardSR(strItem)
	wait(2)
	WaitForIServeLoading
	bVerify = True
	strExpitem = coServiceRequest_Page.lstComboBoxStatusSR.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coServiceRequest_Page.lstComboBoxStatusSR,strItem,"Status")
	End If
	SelectStatuscombobox_DashboardSR = bVerify
End Function

'[Verify list of values in Status dropdown displayed in SR Search panel]
Public Function VerifyStatusDropdown_SR(lstStatus)
	bVerifyValues = True
	If Not IsNull(lstStatus) Then
		bVerifyValues = verifyComboboxItems(coServiceRequest_Page.lstComboBoxStatusSR,lstStatus,"Status")		
	End If
	VerifyStatusDropdown_SR = bVerifyValues
End Function

'[Verify list of values in Group dropdown displayed in SR Search panel]
Public Function VerifyGroupDropdown_SR(lstGroup)
	bVerifyValues = True
	If Not IsNull(lstGroup) Then
		bVerifyValues = verifyComboboxItems(coServiceRequest_Page.lstComboBoxGroupSR,lstGroup,"Group")		
	End If
	VerifyGroupDropdown_SR = bVerifyValues
End Function

'[Verify list of values in Select dropdown displayed in SR Search panel]
Public Function VerifySelectDropdown_SR(lstSelect)
	bVerifyValues = True
	If Not IsNull(lstSelect) Then
		bVerifyValues = verifyComboboxItems(coServiceRequest_Page.lstComboBoxSelectSR,lstSelect,"Select")		
	End If
	VerifySelectDropdown_SR = bVerifyValues
End Function

'[Set Staff Combobox in ServiceRequest Search panel]
Public Function SetStaffcombobox_SR(strItem)
	WaitForIServeLoading
	bVerify = True
	strExpitem = coServiceRequest_Page.txtComboBoxStaffSR.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coServiceRequest_Page.txtComboBoxStaffSR,strItem,"Staff")
	End If	
	SetStaffcombobox_SR = bVerify
End Function

'[Verify display of channel field in SR Search panel]
Public Function VerifyDisplayOfSRChannel(strFlag)
	bVerifyValues = True
	If Not IsNull(strFlag) Then
		bVerifyValues = VerifyObjectEnabledDisabled(coServiceRequest_Page.txtComboBoxChannelSR,strFlag,"Dashboard-SR-Channel Combobox")		
	End If
	VerifyDisplayOfSRChannel = bVerifyValues
End Function

'[Verify default From and To date displayed in Dashboard SR Search panel]
Public Function verifyDefaultDateRange_SR(strDateRange,strToDate)	
	verifyDefaultDateRange_SR = VerifyDateRange(coServiceRequest_Page.txtFromDate,coServiceRequest_Page.txtToDate,strDateRange,strToDate)
End Function

'[Verify display of Submit button in SR Search panel]
Public Function VerifyButtonSubmit_SR(strCheckFlag)
	VerifyButtonSubmit_SR = VerifyObjectEnabledDisabled(coServiceRequest_Page.btnSubmitSR,strCheckFlag,"Submit Button")
End Function

'[Select FROM Date using Date Picker in SR Search Panel]
Public Function SelectFromDate_SR(strFromDate)
	bverifyDate = True

	If Not IsNull(strFromDate) Then
	
		If Trim(strFromDate) = "TODAY" Then
			strFromDate = Day(Now) & " " & MonthName(Month(Now),True) &" "& Year(Now)
		End If
		
		SelectFromDate_SR = SelectDateFromIDCalendar(coServiceRequest_Page.txtFromDate,strFromDate)
		strExpFromDate = Right("0" & Datepart("d",strFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strFromDate),2))&" " & Year(strFromDate)
		
		If SelectFromDate_SR Then
		
			strActFromDate = coServiceRequest_Page.txtFromDate.GetROProperty("value")
			strActFromDate = Right("0" & Datepart("d",strActFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strActFromDate),2))&" " & Year(strActFromDate)
			
			If Trim(strActFromDate) = Trim(strExpFromDate) Then
			   LogMessage "RSLT","Verification","Selected date "&strFromDate&" in FROM date text box is displayed as expected", True
			   bverifyDate = True 
			Else
				LogMessage "WARN","Verification","As expected, Selected date "&strFromDate&" in FROM date text box is not displayed.", False
			   bverifyDate = False 
			End If	
			
		End If
		
	End If
	
	SelectFromDate_SR = bverifyDate
End Function

'[Select TO Date using Date Picker in SR Search Panel]
Public Function SelectTODate_SR(strTODate)
	bverifyDate = True 
	
	If Not IsNull(strTODate) Then
	
		If Trim(strTODate) = "TODAY" Then
		   strTODate = strFromDate = Day(Now) & " " & MonthName(Month(Now),True) &" "& Year(Now)
		End If
		
		SelectTODate_SR =  SelectDateFromIDCalendar(coServiceRequest_Page.txtToDate,strTODate)
		StrExpToDate = Right("0" & Datepart("d",strTODate),2) &" "& MonthName(Right("0" & Datepart("m",strTODate),2))&" " & Year(strTODate)
		
		If SelectTODate_SR Then
		
			strActTODate = coServiceRequest_Page.txtToDate.GetROProperty("value")
			strActTODate = Right("0" & Datepart("d",strActTODate),2) &" "& MonthName(Right("0" & Datepart("m",strActTODate),2))&" " & Year(strActTODate)
			
			If Trim(strActTODate) = Trim(StrExpToDate) Then
			   LogMessage "RSLT","Verification","Selected date "&strTODate&" in TO date text box is displayed as expected", True
			   bverifyDate = True 
			Else
				LogMessage "WARN","Verification","As expected, Selected date "&strTODate&" in TO date text box is not displayed.",False
			  bverifyDate = False
			End If	
			
		End If
	
	End IF
	
	SelectTODate_SR = bverifyDate
End Function

'[Click on Submit Button in SR Search Panel]
Public Function clickButtonSubmitSRSearchPanel()
	coServiceRequest_Page.btnSubmitSR.click 
	If Err.Number <> 0 Then
	  clickButtonSubmitSRSearchPanel = False
	  LogMessage "WARN","Verification","Failed to Click Button : Submit", False
	  Exit Function
	End If
	WaitForIServeLoading
	clickButtonSubmitSRSearchPanel = True
End Function

'[Verify Inline error message displayed in SR Search panel]
Public Function VerifyInlineErrorMsg_SR(strErrorMsg)
	bverifyInlineErrorMsg = True
	If Not VerifyInnerText(coServiceRequest_Page.lblInlineMessage(), strErrorMsg, "Inline Date Error") Then
	   bverifyInlineErrorMsg = False
	End If
	VerifyInlineErrorMsg_SR = bverifyInlineErrorMsg
End Function

'[Verify records displayed in SR Summary table based on Selected From and To Date]
Public Function VerifyRecordDisplayedBasedonDates_SR()
	strFromDate = coServiceRequest_Page.txtFromDate.GetROProperty("value")
	strToDate = coServiceRequest_Page.txtToDate.GetROProperty("value")
	VerifyRecordDisplayedBasedonDates_SR = VerifyDateSearchRecordsdisplayed(coServiceRequest_Page.tblDashboardSRHeader,coServiceRequest_Page.tblDashboardSRBody,strFromDate,strToDate,"CREATED ON")
End Function

'[Verify Pagination for table displayed in Dashboard SR Page]
Public Function VerifyPagination_SR(NoOfRows)
	 bVerifyPagination = False  
	 
	 Set objFristPage = SetObjectFirstPage(coServiceRequest_Page.tblSRPager)
	 Set objPreviousPage = SetObjectPreviousPage(coServiceRequest_Page.tblSRPager)
	 Set objNextPage = SetObjectNextPage(coServiceRequest_Page.tblSRPager)
	 Set objLastPage = SetObjectLastPage(coServiceRequest_Page.tblSRPager)
	 
	 bVerifyPagination = VerifytablePagination(coServiceRequest_Page.tblDashboardSRHeader,coServiceRequest_Page.tblDashboardSRBody,objFristPage,objPreviousPage,objNextPage,objLastPage,"CREATED ON",NoOfRows)
	 
	 VerifyPagination_SR = bVerifyPagination
	 
	 Set objFristPage = Nothing
	 Set objPreviousPage = Nothing
	 Set objNextPage = Nothing
	 Set objLastPage = Nothing
End Function

'[Verify Record count displayed based on Selected Status in SR Summary table]
Public Function VerifySRRecordCount_DashboardSR(strStatus)
	bVerifyRecordCount = False
	
	strDisplayedMsgtext = coServiceRequest_Page.lblRecordCountSR.GetROProperty("innertext")
	strMsgText = "Service Request "&strStatus
	
	If Instr(1,strDisplayedMsgtext,strMsgText,1) > 0 Then 
		LogMessage "WARN","Verification","SR Record Count text message is displayed as expected", True
		bVerifyRecordCount = True
	Else 
		LogMessage "WARN","Verification","SR Record Count text message is not displayed as expected", False
		bVerifyRecordCount = False
	End If
	VerifySRRecordCount_DashboardSR = bVerifyRecordCount
End Function

'[Enter Customer Name OR CIN Number in SR Search textbox]
Public Function SetCustomerNametext_SR(strCustCIN)
	coServiceRequest_Page.txtCustomerCINSR.Set strCustCIN
	StrActText = coServiceRequest_Page.txtCustomerCINSR.GetROProperty("value")
		
	If Ucase(Trim(strCustCIN)) = Ucase(Trim(StrActText)) Then
	   LogMessage "RSLT","Verification","Textbox inside the table is set with value "&strCustCIN&" as Expected", True
	   SetCustomerNametext_SR = True
	Else
	   LogMessage "RSLT","Verification","Textbox inside the table is doesnt set with value "&strCustCIN&" as Expected", False
	   SetCustomerNametext_SR = False
	End If
End Function

'[Verify records displayed in SR Summary table based on Customer Name or CIN Number search]
Public Function VerifyResultsdisplayed_SR(strExpValue)
 	VerifyResultsdisplayed_SR = VerifySearchRecordsdisplayed(coServiceRequest_Page.tblDashboardSRHeader,coServiceRequest_Page.tblDashboardSRBody,"Customer Name / CIF",strExpValue)
End Function

'[Verify Infowarn Message displayed in Dashboard SR Page]
Public Function VerifyInfowan_DashboardSR(strInfoMsgtext)
	VerifyInfowan_DashboardSR = VerifyInfowarntext(coServiceRequest_Page.lblInfowarnSR,strInfoMsgtext)
End Function

'[Verify No Data Message displayed in Dashboard SR]
Public Function VerifyNoDataSRDisplayMessage(strMsg)
	VerifyNoDataSRDisplayMessage = VerifyInnerText(coServiceRequest_Page.lblNoDataSR,strMsg,"No Data Display message")
End Function

'[Select on row displayed in SR results table]
Public Function ClickTableRow_SR(lstRowData)
	ClickTableRow_SR = SelectTableRow(coServiceRequest_Page.tblDashboardSRHeader,coServiceRequest_Page.tblDashboardSRBody,lstRowData,"Dashboard SR Summary","CREATED ON",False,False)
End Function

'[Verify Customer Information in view SR page]
Public Function VerifySRCustomerInformation(arrCustomerInfo)
	VerifySRCustomerInformation = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleViewSRCustInfo,arrCustomerInfo,"View SR Page","Customer Information")
End Function

'[Verify SR Related Information in view SR page upper]
Public Function VerifySRRelatedInfromationU(arrSRRelatedInfo)
	VerifySRRelatedInfromationU = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleViewSRRelatedInfo,arrSRRelatedInfo,"View SR Page","SR Related Information")
End Function

'[Verify SR Related Information in view SR page middle]
Public Function VerifySRRelatedInfromationM(arrSRRelatedInfo)
	VerifySRRelatedInfromationM = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleViewSRRelatedInfo,arrSRRelatedInfo,"View SR Page","SR Related Information")
End Function

'[Verify SR Related Information in view SR page lower]
Public Function VerifySRRelatedInfromationL(arrSRRelatedInfo)
	VerifySRRelatedInfromationL = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleViewSRRelatedInfo,arrSRRelatedInfo,"View SR Page","SR Related Information")
End Function

'[Verify list of accordions present in view SR page]
Public Function VerifySRAccordionList(arrAccordions)
	VerifySRAccordionList = VerifyAccordionHeader(coServiceRequest_Page.eleViewSRAccordionObject,arrAccordions)
End Function

'[Scroll down iserve page]
Public Function ScrollDownIservePage()
	scrollPageDown 20
End Function

'[Verify Additional SR Info accordion data]
Public Function VerifyAdditionalSRInfo(strStatus,arrAdditionalSRInfoData)
	ExpandSingleAccordion coServiceRequest_Page.eleViewSRAccordionObject,"Additional SR Info"
	If UCase(strStatus) <> "OPEN" Then
		VerifyAdditionalSRInfo = VerifyIDLabelValuePairsRandom(coServiceRequest_Page.eleViewSRAdditionalSRInfo,arrAdditionalSRInfoData,"View SR Page","Additional SR Info")
	End If
	CollapseSingleAccordion coServiceRequest_Page.eleViewSRAccordionObject,"Additional SR Info"
End Function

Public Function verifyAccordionDataViewSR(objHeader,objBody,arrData,strAccordName)
	ExpandSingleAccordion coServiceRequest_Page.eleViewSRAccordionObject,strAccordName
	verifyAccordionDataViewSR = VerifyTableSingleRowData(objHeader,objBody,arrData,strAccordName)
	CollapseSingleAccordion coServiceRequest_Page.eleViewSRAccordionObject,strAccordName
End Function

'[Verify SR View Activity Details accordion data]
Public Function VerifyViewActivityDetails(arrActivityData)
	VerifyViewActivityDetails = verifyAccordionDataViewSR(coServiceRequest_Page.tblViewActivityDetailsSRHeader,coServiceRequest_Page.tblViewActivityDetailsSRBody,arrActivityData,"View Activity Details")
End Function

'[Verify View SR Attachment accordion data]
Public Function VerifySRAttachment(arrAttachmentData)
	VerifySRAttachment = verifyAccordionDataViewSR(coServiceRequest_Page.tblAttachmentSRHeader,coServiceRequest_Page.tblAttachmentSRBody,arrAttachmentData,"Attachment")
End Function

'[Verify View SR Step Details accordion data]
Public Function VerifySRStepDetails(arrStepDetailsData)
	VerifySRStepDetails = verifyAccordionDataViewSR(coServiceRequest_Page.tblStepDetailsSRHeader,coServiceRequest_Page.tblStepDetailsSRBody,arrStepDetailsData,"Step Details")
End Function

'[Verify View SR Workflow Comments History accordion data]
Public Function VerifySRWorkflowCommentsHistory(arrWorkFlowHistData)
	VerifySRWorkflowCommentsHistory = verifyAccordionDataViewSR(coServiceRequest_Page.tblWorkflowCommentsHistorySRHeader,coServiceRequest_Page.tblWorkflowCommentsHistorySRBody,arrWorkFlowHistData,"Workflow Comments History")
End Function

'[Verify View SR Workflow Case History accordion data]
Public Function VerifySRWorkflowCaseHistory(arrWorkFlowCaseHistData)
	VerifySRWorkflowCaseHistory = verifyAccordionDataViewSR(coServiceRequest_Page.tblWorkflowCaseHistorySRHeader,coServiceRequest_Page.tblWorkflowCaseHistorySRBody,arrWorkFlowCaseHistData,"Workflow Case History")
End Function

'[Verify No Data Messages displayed in Attachment Accordion]
Public Function VerifyAttachementNoDataMessage(strFilenetError,strCRMError)
	ExpandSingleAccordion coServiceRequest_Page.eleViewSRAccordionObject,"Attachment"
	VerifyAttachementNoDataMessage = VerifyInnerText(coServiceRequest_Page.lblNoDataSR,strFilenetError,"Filenet No Data message")
	VerifyAttachementNoDataMessage = VerifyInnerText(coServiceRequest_Page.eleCRMNoAttachmentMsg,strCRMError,"CRM No Data message")
End Function
	
	
'################################# FUNCTIONS FOR [VIEW SR] ENDS ########################################################



'################################# FUNCTIONS FOR [EDIT SR] STARTS ########################################################

'[Click on Edit link in View SR Page]
Public Function ClickEditSRLink()
	ClickEditSRLink = ClickOnObject(coServiceRequest_Page.lnkEditSR,"Top Edit SR Link")
End Function

'[Expand single accordion in Edit SR Page]
Public Function ExpandAccordionEditSR(strAccordName)
	ExpandAccordionEditSR = False
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 800")
	Wait 2
	intCount = ExpandSingleAccordion(coServiceRequest_Page.eleViewSRAccordionObject,strAccordName)
	If intCount > 0 Then
		ExpandAccordionEditSR = True
	End If
End Function

'################################# FUNCTIONS FOR [EDIT SR] ENDS ########################################################
