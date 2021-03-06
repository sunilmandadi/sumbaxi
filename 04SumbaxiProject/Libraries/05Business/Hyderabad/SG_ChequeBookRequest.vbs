'*****This is auto generated code using code generator please Re-validate ****************

'[Verify New Cheque Book button Exist in Cheque Info Page]
Public Function verifyBtnNewChequeBook()
	bverifyBtnNewChequeBook=true
	If Not (CBRequest.btnNewChequeBook().Exist) Then
		LogMessage "WARN", "Verification", "Unable to find New Cheque Book button in Cheque Info Page.", False
		bverifyBtnNewChequeBook=false
	End If
	verifyBtnNewChequeBook=bverifyBtnNewChequeBook
End Function

'[Click on New Cheque Book Button]
Public Function clickBtnNewChequeBook()
	CBRequest.btnNewChequeBook.click
	If Err.Number<>0 Then
       clickBtnNewChequeBook=false
       LogMessage "WARN","Verification","Failed to Click Button : New Cheque Book" ,false
       Exit Function
   End If
    clickBtnNewChequeBook=true
	WaitForICallLoading
End Function

'[Verify row Data in Table SelectedCards for Cheque Book Request]
Public Function verifytblSelectedCardsContent_CBR(arrRowDataList)
	WaitForICallLoading
   verifytblSelectedCardsContent_CBR=verifyTableContentList(CBRequest.tblSelectedCardsHeader,CBRequest.tblSelectedCardsContent,arrRowDataList,"SelectedCardsContent" , false,null ,null,null)
End Function

'[Verify Field Validation Message For Cheque Book Request displayed as]
Public Function verifyValidationMessage_CBR(strExpectedText)
   bverifyValidationMessage_CBR=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblValidationMessage(), strExpectedText, "Validation Message")Then
           bverifyValidationMessage_CBR=false
       End If
   End If
   CBRequest.btnOK_ValidationMsg.Click
   WaitForICallLoading
   verifyValidationMessage_CBR=bverifyValidationMessage_CBR
End Function

'[Verify Available Balance for Cheque Book Request dispalyed as]
Public Function verifyAvailableBalance_CBR(strAvailableBalance)
	bverifyAvailableBalance=true
	If Not IsNull(strAvailableBalance) Then
       If Not VerifyInnerText (CBRequest.lblAvailableBalance(), strAvailableBalance, "Available Balance")Then
           bverifyAvailableBalance=false
       End If
    End If
    verifyAvailableBalance_CBR=bverifyAvailableBalance
End Function

'[Verify Number Of Cheque Book Combobox on Cheque Book Request has Items]
Public Function verifyNoofChequeBookComboboxItems_CBR(lstItems)
   bverifyNoofChequeBookComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems(CBRequest.lstNoChequeBook(),lstItems, "Number of Cheque Book")Then
           bverifyNoofChequeBookComboboxItems=false
       End If
   End If
   verifyNoofChequeBookComboboxItems_CBR=bverifyNoofChequeBookComboboxItems
End Function

'[Verify Combobox Number of Cheque Book displayed as]
Public Function verifyNoofChequeBook_Default(strChequeBookCount)
   bverifyNoofChequeBook_Default=true
   If Not IsNull(strChequeBookCount) Then
       If Not verifyComboSelectItem(CBRequest.lstNoChequeBook(),strChequeBookCount, "Number of Cheque Book")Then
           bverifyNoofChequeBook_Default=false
       End If
   End If
   verifyNoofChequeBook_Default=bverifyNoofChequeBook_Default
End Function

'[Select Combobox Number of Cheque Book in Cheque Book Request]
Public Function selectNoofChequeBook_CBR(strNoofChequeBook)
	bselectNoofChequeBook=true
	If Not IsNull(strNoofChequeBook) Then
       If Not (selectItem_Combobox(CBRequest.lstNoChequeBook(), strNoofChequeBook))Then
            LogMessage "WARN","Verification","Failed to select :"&strNoofChequeBook&" From No of Cheque Book drop down list" ,false
           bselectNoofChequeBook=false
       End If
   End If
   WaitForICallLoading
   selectNoofChequeBook_CBR=bselectNoofChequeBook
End Function

'[Verify Urgency Combobox has Items]
Public Function verifyUrgencyComboboxItems(lstItems)
   bverifyUrgencyComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems(CBRequest.lstUrgency(),lstItems, "Urgency")Then
           bverifyUrgencyComboboxItems=false
       End If
   End If
   verifyUrgencyComboboxItems=bverifyUrgencyComboboxItems
End Function

'[Verify Combobox Urgency displayed as]
Public Function verifyUrgency_Default(strUrgency_Default)
   bverifyUrgency_Default=true
   If Not IsNull(strUrgency_Default) Then
       If Not verifyComboSelectItem(CBRequest.lstUrgency(),strUrgency_Default, "Urgency")Then
           bverifyUrgency_Default=false
       End If
   End If
   verifyUrgency_Default=bverifyUrgency_Default
End Function

'[Select Combobox Urgency in Cheque Book Request]
Public Function selectUrgency_CBR(strUrgency)
	bselectUrgency=true
	If Not IsNull(strUrgency) Then
       If Not (selectItem_Combobox(CBRequest.lstUrgency(), strUrgency))Then
            LogMessage "WARN","Verification","Failed to select :"&strUrgency&" From Urgency drop down list" ,false
           bselectUrgency=false
       End If
   End If
   WaitForICallLoading
   selectUrgency_CBR=bselectUrgency
End Function

'[Verify Dispatched Mode Combobox has Items]
Public Function verifyDispatchedModeComboboxItems(lstItems)
   bverifyDispatchedModeComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems(CBRequest.lstDispatchedMode(),lstItems, "Dispatched Mode")Then
           bverifyDispatchedModeComboboxItems=false
       End If
   End If
   verifyDispatchedModeComboboxItems=bverifyDispatchedModeComboboxItems
End Function

'[Verify Combobox Dispatched Mode displayed as]
Public Function verifyDispatchedMode_Default(strDispatchedMode_Default)
   bverifyDispatchedMode_Default=true
   If Not IsNull(strDispatchedMode_Default) Then
       If Not verifyComboSelectItem (CBRequest.lstDispatchedMode(),strDispatchedMode_Default, "Dispatched Mode")Then
           bverifyDispatchedMode_Default=false
       End If
   End If
   verifyDispatchedMode_Default=bverifyDispatchedMode_Default
End Function

'[Select Combobox Dispatched Mode in Cheque Book Request]
Public Function selectDispatchedMode_CBR(strDispatchedMode)
	bselectDispatchedMode=true
	If Not IsNull(strDispatchedMode) Then
       If Not (selectItem_Combobox (CBRequest.lstDispatchedMode(), strDispatchedMode))Then
            LogMessage "WARN","Verification","Failed to select :"&strDispatchedMode&" From Dispatched Mode drop down list" ,false
           bselectDispatchedMode=false
       End If
   End If
   WaitForICallLoading
   selectDispatchedMode_CBR=bselectDispatchedMode
End Function

'[Verify Delivery Mode Combobox on Cheque Book Request has Items]
Public Function verifyDeliveryModeComboboxItems_CBR(lstItems)
   bverifyDeliveryModeComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CBRequest.lstDeliveryMode(),lstItems, "Delivery Mode")Then
           bverifyDeliveryModeComboboxItems=false
       End If
   End If
   verifyDeliveryModeComboboxItems_CBR=bverifyDeliveryModeComboboxItems
End Function

'[Verify Combobox Delivery Mode displayed as]
Public Function verifyDeliveryMode_Default(strDeliveryMode_Default)
   bverifyDeliveryMode_Default=true
   If Not IsNull(strDeliveryMode_Default) Then
       If Not verifyComboSelectItem (CBRequest.lstDeliveryMode(),strDeliveryMode_Default, "Delivery Mode")Then
           bverifyDeliveryMode_Default=false
       End If
   End If
   verifyDeliveryMode_Default=bverifyDeliveryMode_Default
End Function

'[Select Combobox Delivery Mode in Cheque Book Request]
Public Function selectDeliveryMode_CBR(strDeliveryMode)
	bselectDeliveryMode=true
	If Not IsNull(strDeliveryMode) Then
       If Not (selectItem_Combobox (CBRequest.lstDeliveryMode(), strDeliveryMode))Then
            LogMessage "WARN","Verification","Failed to select :"&strDeliveryMode&" From Delivery Mode drop down list" ,false
           bselectDeliveryMode=false
       End If
   End If
   WaitForICallLoading
   selectDeliveryMode_CBR=bselectDeliveryMode
End Function

'[Verify Collection Branch Combobox has Items]
Public Function verifyCollectionBranchComboboxItems(lstItems)
   bverifyCollectionBranchComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CBRequest.lstCollectionBranch(),lstItems, "Collection Branch")Then
           bverifyCollectionBranchComboboxItems=false
       End If
   End If
   verifyCollectionBranchComboboxItems=bverifyCollectionBranchComboboxItems
End Function

'[Verify Combobox Collection Branch displayed as]
Public Function verifyCollectionBranch_Default(strCollectionBranch_Default)
   bverifyCollectionBranch_Default=true
   If Not IsNull(strCollectionBranch_Default) Then
       If Not verifyComboSelectItem (CBRequest.lstCollectionBranch(),strCollectionBranch_Default, "Collection Branch")Then
           bverifyCollectionBranch_Default=false
       End If
   End If
   verifyCollectionBranch_Default=bverifyCollectionBranch_Default
End Function

'[Select Combobox Collection Branch displayed in Cheque Book Request]
Public Function selectCollectionBranch_CBR(strCollectionBranch)
	bselectCollectionBranch_CBR=true
	If Not IsNull(strCollectionBranch) Then
       If Not (selectItem_Combobox (CBRequest.lstCollectionBranch(), strCollectionBranch))Then
            LogMessage "WARN","Verification","Failed to select :"&strCollectionBranch&" From Collection Branch drop down list" ,false
           bselectCollectionBranch_CBR=false
       End If
   End If
   WaitForICallLoading
   selectCollectionBranch_CBR=bselectCollectionBranch_CBR
End Function

'[Verify Courier Fee Charge Combobox has Items]
Public Function verifyCourierFeeChargeComboboxItems(lstItems)
   bverifyCourierFeeChargeComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CBRequest.lstCourierFeeCharge(),lstItems, "Courier Fee Charge")Then
           bverifyCourierFeeChargeComboboxItems=false
       End If
   End If
   verifyCourierFeeChargeComboboxItems=bverifyCourierFeeChargeComboboxItems
End Function

'[Verify Combobox Courier Fee Charge displayed as]
Public Function verifyCourierFeeCharge_Default(strCourierFeeCharge_Default)
   bverifyCourierFeeCharge_Default=true
   If Not IsNull(strCourierFeeCharge_Default) Then
       If Not verifyComboSelectItem (CBRequest.lstCourierFeeCharge(),strCourierFeeCharge_Default, "Courier Fee Charge")Then
           bverifyCourierFeeCharge_Default=false
       End If
   End If
   verifyCourierFeeCharge_Default=bverifyCourierFeeCharge_Default
End Function

'[Select Combobox Courier Fee Charge in Cheque Book Request]
Public Function selectCourierFeeCharge_CBR(strCourierFeeCharge)
	bselectCourierFeeCharge_CBR=true
	If Not IsNull(strCourierFeeCharge) Then
       If Not (selectItem_Combobox (CBRequest.lstCourierFeeCharge(), strCourierFeeCharge))Then
            LogMessage "WARN","Verification","Failed to select :"&strCourierFeeCharge&" From Courier Fee Charge drop down list" ,false
           bselectCourierFeeCharge_CBR=false
       End If
   End If
   WaitForICallLoading   
   selectCourierFeeCharge_CBR=bselectCourierFeeCharge_CBR
End Function

'[Verify PC Code in Cheque Book Request displayed as]
Public Function verifyPCCode_CBR(strPCCode)
	bverifyPCCode_CBR=true
	If Not IsNull(strPCCode) Then
       If Not VerifyInnerText (CBRequest.lblPCCode(), strPCCode, "PC Code")Then
           bverifyPCCode_CBR=false
       End If
    End If
    verifyPCCode_CBR=bverifyPCCode_CBR
End Function

'[Verify default waive cheque book Fee Radio Button displayed on Cheque Book Request Screen]
Public Function verifydefaultWaiveChequeBookFee_CBR(strSelectedradioButton)	
	bverifydefaultWaiveChequeBookFee=true
	If Not IsNull(strSelectedradioButton) Then
		bverifydefaultWaiveChequeBookFee=VerifyRadioButtonGrpSelection(strSelectedradioButton,CBRequest.rbtnGroupWaiveChequeBookFee,Array("Yes","No"))
		If bverifydefaultWaiveChequeBookFee Then
			LogMessage "RSLT","Verification","Radio Button :NO selected by default. Selected value is "&strSelectedradioButton ,true
		else
			LogMessage "RSLT","Verification","Radio Button :NO is not selected by default. Selected value is "&strSelectedradioButton ,false
		End If
	    If Err.Number<>0 Then
	       bverifydefaultWaiveChequeBookFee=false
	       LogMessage "WARN","Verification","Failed to Verify Radio Button :Show" ,false
	       Exit Function
	   End If
	End IF 
   verifydefaultWaiveChequeBookFee_CBR=bverifydefaultWaiveChequeBookFee
End Function

'[Select Radio Button of Waive Cheque Book Fee on Cheque Book Request Screen]
Public Function selectWaiveChequeBookFeeRadio_CBR(strWaiveChequeBookFee)
	bselectWaiveChequeBookFeeRadio=true
	If Not IsNull(strWaiveChequeBookFee) Then
		bselectWaiveChequeBookFeeRadio_CBR=SelectRadioButtonGrp(strWaiveChequeBookFee,CBRequest.rbtnWaiveChequeBookFee, Array("Yes","No"))   
		If Err.Number<>0 Then
	       bselectWaiveChequeBookFeeRadio=false
	       LogMessage "WARN","Verification","Failed to Click Button : Waive Cheque Book Fee" ,false
	       Exit Function
	   End If
   End IF
   selectWaiveChequeBookFeeRadio_CBR=bselectWaiveChequeBookFeeRadio
End Function

'[Verify Radio Button of Waive Cheque Book Fee is disable on Cheque Book Request Screen]
Public Function verifyRadioBtnWaiveFee_Disable()
	bverifyRadioBtnWaiveFee_Disable=true
	WaitForICallLoading
	'intrBtnWaiveFee=Instr(CBRequest.rbtnWaiveChequeBookFee.GetROproperty("outerhtml"),"v-disabled")
	intrBtnWaiveFee =InStr(CBRequest.rbtnWaiveChequeBookFee.GetROProperty("class"),"disabled-area")
	If intrBtnWaiveFee <> 0 Then
		LogMessage "RSLT","Verification","Waive Fee Radio button is disable as expected.",True
	Else
		LogMessage "WARN","Verifiation","Waive Fee Radio button is enable. Expected to be disable.",false
		bverifyRadioBtnWaiveFee_Disable=false
	End If
	verifyRadioBtnWaiveFee_Disable=bverifyRadioBtnWaiveFee_Disable
End Function

'[Verify Total Fee Charge in Cheque Book Request displayed as]
Public Function verifyTotalFeeCharge_CBR(strTotalFeeCharge)
	bverifyTotalFeeCharge_CBR=true
	If Not IsNull(strTotalFeeCharge) Then
       If Not VerifyInnerText (CBRequest.lblTotalFeeCharge(),strTotalFeeCharge, "Total Fee Charge")Then
           bverifyTotalFeeCharge_CBR=false
       End If
    End If
    verifyTotalFeeCharge_CBR=bverifyTotalFeeCharge_CBR
End Function

'[Select Combobox Cheque Book Fee Waiver Reason in Cheque Book Request]
Public Function selectCBFeeWaiverReason_CBR(strCBFeeWaiverReason)
	bselectCBFeeWaiverReason_CBR=true
	If Not IsNull(strCBFeeWaiverReason) Then
       If Not (selectItem_Combobox (CBRequest.lstCBFeeWaiverReason(),strCBFeeWaiverReason))Then
           LogMessage "WARN","Verification","Failed to select :"&strCBFeeWaiverReason&" From Cheque Book Fee Waiver Reason drop down list" ,false
           bselectCBFeeWaiverReason_CBR=false
       End If
   End If
   WaitForICallLoading
   selectCBFeeWaiverReason_CBR=bselectCBFeeWaiverReason_CBR
End Function

'[Verify Cheque Book Fee Waiver Reason Combobox has Items]
Public Function verifyFeeWaiverReasonComboboxItems_CBR(lstItems)
   bverifyFeeWaiverReasonComboboxItems_CBR=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CBRequest.lstCBFeeWaiverReason(),lstItems, "Fee Waiver Reason")Then
           bverifyFeeWaiverReasonComboboxItems_CBR=false
       End If
   End If
   verifyFeeWaiverReasonComboboxItems_CBR=bverifyFeeWaiverReasonComboboxItems_CBR
End Function

'[Verify Combobox Cheque Book Fee Waiver Reason displayed as]
Public Function verifyFeeWaiverReason_Default(strFeeWaiverReason_Default)
   bverifyFeeWaiverReason_Default=true
   If Not IsNull(strFeeWaiverReason_Default) Then
       If Not verifyComboSelectItem (CBRequest.lstCBFeeWaiverReason(),strFeeWaiverReason_Default, "Fee Waiver Reason")Then
           bverifyFeeWaiverReason_Default=false
       End If
   End If
   verifyFeeWaiverReason_Default=bverifyFeeWaiverReason_Default
End Function

'[Verify Combobox Cheque Book Fee Waiver Reason is disable on Cheque Book Request Screen]
Public Function verifycbWaiveFeeReason_Disable()
	bverifycbWaiveFeeReason_Disable=true
	'intrcbWaiveFeeReason=Instr(CBRequest.lstCBFeeWaiverReason.GetROproperty("outerhtml"),"v-disabled")
	intrcbWaiveFeeReason =InStr(CBRequest.lstCBFeeWaiverReason.GetROProperty("class"),"disabled-area")
	If intrcbWaiveFeeReason <> 0 Then
		LogMessage "RSLT","Verification","Combobox Waive Fee reason is disable as expected.",True
	Else
		LogMessage "WARN","Verifiation","Combobox Waive Fee reason is enable. Expected to be disable.",false
		bverifycbWaiveFeeReason_Disable=false
	End If
	verifycbWaiveFeeReason_Disable=bverifycbWaiveFeeReason_Disable
End Function

'[Verify Field Description displayed on Cheque Book Request Screen as]
Public Function verifyDescriptionText_CBR(strExpectedText)
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_CBR=bVerifyDescriptionText
End Function

'[Verify Field Block displayed on Cheque Book Request Screen as]
Public Function verifyBlockText_CBR(strExpectedText)
   bVerifyBlockText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblBlock(), strExpectedText, "Block")Then
           bVerifyBlockText=false
       End If
   End If
   verifyBlockText_CBR=bVerifyBlockText
End Function

'[Verify Field Level displayed on Cheque Book Request Screen as]
Public Function verifyLevelText_CBR(strExpectedText)
   bVerifyLevelText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblLevel(), strExpectedText, "Level")Then
           bVerifyLevelText=false
       End If
   End If
   verifyLevelText_CBR=bVerifyLevelText
End Function

'[Verify Field Unit displayed on Cheque Book Request Screen as]
Public Function verifyUnitText_CBR(strExpectedText)
   bVerifyUnitText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblUnit(), strExpectedText, "Unit")Then
           bVerifyUnitText=false
       End If
   End If
   verifyUnitText_CBR=bVerifyUnitText
End Function

'[Verify Field Address Line1 displayed on Cheque Book Request Screen as]
Public Function verifyAddressLine1Text_CBR(strExpectedText)
   bVerifyAddressLine1Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblAddressLine1(), strExpectedText, "Address Line1")Then
           bVerifyAddressLine1Text=false
       End If
   End If
   verifyAddressLine1Text_CBR=bVerifyAddressLine1Text
End Function

'[Verify Field Address Line2 displayed on Cheque Book Request Screen as]
Public Function verifyAddressLine2Text_CBR(strExpectedText)
   bVerifyAddressLine2Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblAddressLine2(), strExpectedText, "Address Line2")Then
           bVerifyAddressLine2Text=false
       End If
   End If
   verifyAddressLine2Text_CBR=bVerifyAddressLine2Text
End Function

'[Verify Field Address Line3 displayed on Cheque Book Request Screen as]
Public Function verifyAddressLine3Text_CBR(strExpectedText)
   bVerifyAddressLine3Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblAddressLine3(), strExpectedText, "Address Line3")Then
           bVerifyAddressLine3Text=false
       End If
   End If
   verifyAddressLine3Text_CBR=bVerifyAddressLine3Text
End Function

'[Verify Field Address Line4 displayed on Cheque Book Request Screen as]
Public Function verifyAddressLine4Text_CBR(strExpectedText)
   bVerifyAddressLine4Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblAddressLine4(), strExpectedText, "Address Line4")Then
           bVerifyAddressLine4Text=false
       End If
   End If
   verifyAddressLine4Text_CBR=bVerifyAddressLine4Text
End Function

'[Verify Field Address Line5 displayed on Cheque Book Request Screen as]
Public Function verifyAddressLine5Text_CBR(strExpectedText)
   bVerifyAddressLine5Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblAddressLine5(), strExpectedText, "Address Line5")Then
           bVerifyAddressLine5Text=false
       End If
   End If
   verifyAddressLine5Text_CBR=bVerifyAddressLine5Text
End Function

'[Verify Field Postal Code displayed on Cheque Book Request Screen as]
Public Function verifyPostalCodeText_CBR(strExpectedText)
   bDevPending=false
   bVerifyPostalCodeText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblPostalCode(), strExpectedText, "Postal Code")Then
           bVerifyPostalCodeText=false
       End If
   End If
   verifyPostalCodeText_CBR=bVerifyPostalCodeText
End Function

'[Verify Field KnowledgeBase on Cheque Book Request SR Screen displayed as]
Public Function verifyKnowledgeBase_CBR(strExpectedLink)
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strExpectedLink) Then		
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"		
			'strKBLink=CBRequest.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
			strKBLink=CBRequest.lnkKnowledgeBase.GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_CBR=bVerifyKnowledgeBaseText
End Function

'[Set TextBox on Cheque Book Request Comment to]
Public Function setCommentTextbox_CBR(strComment)
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox on Cheque Book Request Comment to"
	insertDataStore "SRComment", strComment
   CBRequest.txtComment.Set(strComment)
   If Err.Number<>0 Then
       setCommentTextbox_CBR=false
       LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_CBR=true
End Function

'[Perform Add Notes by clicking Add Notes Button on Cheque Book Request Screen]
Public Function addNote_CBR(strNote)
   baddNote_CL=true	
	If not isNull(strNote) Then
		CBRequest.btnAddNotes.click
		WaitForICallLoading
           If Not CBRequest.popupValidationMessage.exist(5)Then
			  LogMessage "WARN","Verification","Add New Comment action failed"
			  baddNote_CL=false
		   else
			  LogMessage "RSLT","Verification","Add New Comment performed successfully" ,true
			  baddNote_CL=True
	  	   End If
		CBRequest.txtNotes.set strNote
		CBRequest.btnOK_ValidationMsg.Click
		WaitForIcallLoading
	End If		
	addNote_CBR=baddNote_CL
End Function

'[Verify Button Add Notes is disable on Cheque Book Request Screen]
Public Function verifybtnAddNotes_Disable()
	bverifybtnAddNotes_Disable=true
	intrBtnWaiveFee=Instr(CBRequest.btnAddNotes.Object.GetAttribute("disabled"),("disabled"))
	If not intrBtnWaiveFee=0 Then
		LogMessage "RSLT","Verification","Add Notes button is disable as expected.",True
		bverifybtnAddNotes_Disable=true
	Else
		LogMessage "WARN","Verifiation","Add Notes button is enable. Expected to be disable.",false
		bverifybtnAddNotes_Disable=false
	End If
	verifybtnAddNotes_Disable=bverifybtnAddNotes_Disable
End Function

'[Click Button Submit on Cheque Book Request]
Public Function clickButtonSubmit_CBR()
   CBRequest.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_CBR=false
       LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonSubmit_CBR=true
End Function

'[Verify Button Submit is disable on Cheque Book Request Screen]
  Public Function verifybtnSubmit_Disable()
	bverifybtnSubmit_Disable=true
	'intrBtnWaiveFee=Instr(CBRequest.btnSubmit.GetROproperty("outerhtml"),"v-disabled")
	intrBtnWaiveFee=Instr(CBRequest.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
	If not intrBtnWaiveFee = 0 Then
		LogMessage "RSLT","Verification","Submit button is disable as expected.",True
		bverifybtnSubmit_Disable=true
	Else
		LogMessage "WARN","Verifiation","Submit button is enable. Expected to be disable.",false
		bverifybtnSubmit_Disable=false
	End If
	verifybtnSubmit_Disable=bverifybtnSubmit_Disable
End Function

'[Verify Button Submit is enable on Cheque Book Request Screen]
Public Function verifybtnSubmit_enable()
	bverifybtnSubmit_enable=true
	intrBtnWaiveFee=Instr(CBRequest.btnSubmit.GetROproperty("outerhtml"),"v-disabled")
	If intrBtnWaiveFee=0 Then
		LogMessage "RSLT","Verification","Submit button is disable as expected.",True
		bverifybtnSubmit_enable=true
	Else
		LogMessage "WARN","Verifiation","Add Notes button is enable. Expected to be disable.",false
		bverifybtnSubmit_Disable=false
	End If
	verifybtnSubmit_enable=bverifybtnSubmit_enable
End Function

'[Click Button Cancel on Cheque Book Request]
Public Function clickButtonCancel_CBR()
   CBRequest.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_CBR=false
       LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel_CBR=true
End Function

'[Verify Confirmation Popup on Cheque Book Request]
Public Function verifyConfirmationPopup_CBR(strConfirmationMsg)
	bverifyConfirmationPopup_CBR=true
	If Not IsNull (strConfirmationMsg) Then	
		If Not verifyInnerText(CBRequest.lblValidationMessage(), strConfirmationMsg, "Confirmation Message") Then
			bverifyConfirmationPopup_CBR=false
		End If	
	CBRequest.btnYes.click
	End If
    If Err.Number<>0 Then
       bverifyConfirmationPopup_CBR=false
            LogMessage "WARN","Verification","Failed to Click Button : Yes on Confirmation popup" ,false
       Exit Function
    End If
	verifyConfirmationPopup_CBR=bverifyConfirmationPopup_CBR
End Function

'[Verify Popup Request Submitted exist for Cheque Book Request]
Public Function verifyPopupRequestSubmitted_CBR(bExist)
   bActualExist=CBRequest.popupRequestSubmitted.Exist(4)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupRequestSubmitted_CBR=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupRequestSubmitted_CBR=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupRequestSubmitted_CBR=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupRequestSubmitted_CBR=False
   End If
End Function

'[Verify Field CardNumber on Request Submitted Popup for Cheque Book Request displayed as]
Public Function verifyCardNumberReqSubmit_CBR(strCardNumber)
   bVerifyCardNumber = true
   'insertDataStore "NewSAUsedCard", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       If Not VerifyInnerText (CBRequest.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber=false
       End If
   End If
   verifyCardNumberReqSubmit_CBR = bVerifyCardNumber
End Function

'[Verify Field ProductDescription on Request Submitted Popup for Cheque Book Request displayed as]
Public Function verifyProductDescription_RequestSubmitted_CBR(strProductDescription)
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (CBRequest.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_CBR=bVerifyProductDescription_RequestSubmittedText
End Function

'[Verify Field Status on Request Submitted Popup for Cheque Book Request displayed as]
Public Function verifyStatusReqSubmit_CBR(strStatus)
   bVerifyStatus=true
   If Not IsNull(strStatus) Then
       If Not VerifyInnerText (CBRequest.lblStatus_RequestSubmitted(), strStatus, "Status_RequestSubmitted")Then
           bVerifyStatus=false
       End If
   End If
   verifyStatusReqSubmit_CBR = bVerifyStatus
End Function

'[Click Close button on Request Submitted Popup for Cheque Book Request]
Public Function verifybtnClose_RequestSubmitted_CBR()
	bverifybtnClose_RequestSubmitted_CBR=true
	CBRequest.btnClose_RequestSubmitted.click
   If Err.Number<>0 Then
       bverifybtnClose_RequestSubmitted_CBR=false
       LogMessage "WARN","Verification","Failed to Click Close Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
   WaitForICallLoading
	verifybtnClose_RequestSubmitted_CBR=bverifybtnClose_RequestSubmitted_CBR
End Function

'[Verify Field Inline Message for Total Fee Charge displayed on Cheque Book Request Screen as]
Public Function verifyInlineMessage_CBR(strExpectedText)
   bverifyInlineMessage_CBR=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CBRequest.lblInlineMessage(), strExpectedText, "Inline Message")Then
           bverifyInlineMessage_CBR=false
       End If
   End If
   verifyInlineMessage_CBR=bverifyInlineMessage_CBR
End Function

'[Click Link SR Status on Request Submitted popup for Cheque Book Request]
Public Function clickLinkSRStatus_RequestSubmitted_CBR()
   strSelectedSRStatus=CBRequest.lblStatus_RequestSubmitted.GetRoProperty("innerText")
	If strSelectedSRStatus = "Failed" Then
		CBRequest.lblStatus_RequestSubmitted.click
		WaitForICallLoading
		call clickUpdateRequest()
		WaitForICallLoading
		clickLinkSRStatus_RequestSubmitted_CBR=verifyStatus_ConfirmationPopup("Please Select")
		clickLinkSRStatus_RequestSubmitted_CBR=clickbtnCancel_ConfirmationPopup
		WaitForICallLoading
		Exit Function
	Else
		call verifybtnClose_RequestSubmitted_CBR()	
	End If
   If Err.Number<>0 Then
       clickLinkSRStatus_RequestSubmitted_CBR=false
            LogMessage "WARN","Verification","Failed to Click Link : SRStatus_RequestSubmitted" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickLinkSRStatus_RequestSubmitted_CBR=true
End Function
