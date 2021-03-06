'[Verify SPECIAL OFFER Link Enable Disable State]
Public Function verifySpecialOfferlinkState(strstatus)
bVerify = True
	If Not IsNull(strstatus) Then
		If coSpecialOffer_Page.lnkSpecialOffer.Exist(0) Then
			stractStatus = Instr(coSpecialOffer_Page.lnkSpecialOffer.GetROProperty("class"),"isrv-text-blink")
			If strstatus="Enable" Then
				If stractStatus>0 Then
					LogMessage "RSLT","Verification","Special Offer link is in Enabled Mode", True
				Else
					LogMessage "RSLT","Verification","Special Offer link is in Disabled Mode", False
					bVerify=False
				End If
			ElseIf strstatus="Disable" Then
				If stractStatus=0 Then
					LogMessage "RSLT","Verification","Special Offer link is in Disabled Mode", True
				Else
					LogMessage "RSLT","Verification","Special Offer link is in Enabled Mode", False
					bVerify=False
				End If
			Else
			bVerify=False
			End If
		Else
		   bVerify=False
		End If
	Else
	bVerify=False
	End If
	verifySpecialOfferlinkState = bVerify
End Function

'[Click Link Special Offer on Customer Overview Screen under Facilities]
Public Function clickSpecialOfferLink()
   bDevPending=false
   coSpecialOffer_Page.lnkSpecialOffer.click
   If Err.Number<>0 Then
       clickSpecialOfferLink=false
       LogMessage "WARN","Verification","Failed to Click Link : Special Offers" ,false
       Exit Function
   End If
   clickSpecialOfferLink=true
End Function

'[Verify Offer List Tab displayed]
Public Function verifyTabViewOfferListexist(strTabName)
   bDevPending=false
   verifyTabViewOfferListexist=verifyTabExist(strTabName)
End Function

'[Verify fields displayed in Special Offers Grey panel section]
Public Function verifySpecialOffersGreyPanel(arrLblValPairs)
	verifySpecialOffersGreyPanel = VerifyIDLabelValuePairs(coSpecialOffer_Page.lblOfferListHeader,arrLblValPairs,"OFFER LIST","Grey Panel")
End Function

'[Verify table row data for Special Offers List table]
Public Function VerifytblSelectedCard_CC(lstOfferList)
	VerifytblSelectedCard_CC = VerifyTableSingleRowData(coSpecialOffer_Page.tblSpecialOfferListHeader,coSpecialOffer_Page.tblSpecialOffersListBody,lstOfferList,"OFFER LIST")
End Function

'[Select Row from Special Offers List table displayed]
Public Function SelectRow_SpecialOffersLst(lstOffersList)
	SelectRow_SpecialOffersLst = SelectTableRow(coSpecialOffer_Page.tblSpecialOfferListHeader,coSpecialOffer_Page.tblSpecialOffersListBody,lstOffersList,"OFFER LIST","OFFER REFERENCE NO",False,False)
	'Wait(2)
	If VerifyTabExist("NEW OFFER") Then
		SelectRow_SpecialOffersLst = True
		LogMessage "RSLT","Verification","As expected On Clicking Offer List Result Row, New Offer page displayed",True
	Else
		SelectRow_SpecialOffersLst = False
		LogMessage "WARN","Verification","Failed: On Clicking Offer List Result Row, New Offer page is not displayed",False
	End If
End Function

'[Verify New Offer Label Reference Number]
Public Function VerifyNewOfferRefNumber(strRefNumber)
	bVerify=True	
	If Not IsNull(strRefNumber) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblRefNumber(),strRefNumber,"Reference Number label Name")Then
				bVerify = False
			End If
	End If				
	VerifyNewOfferRefNumber=bVerify	
End Function

'[Verify New Offer Label Value Reference Number]
Public Function VerifyNewOfferRefNumberVal(strRefNumberVal)
	bVerify=True	
	If Not IsNull(strRefNumberVal) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblRefNumberval(),strRefNumberVal,"Reference Number label Value")Then
				bVerify = False
			End If
	End If				
	VerifyNewOfferRefNumberVal=bVerify	
End Function

'[Verify New Offer Label Offer Type]
Public Function VerifyNewOfferType(strOfferType)
	bVerify=True	
	If Not IsNull(strOfferType) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblOfferType(),strOfferType,"Offer Type label Name")Then
				bVerify = False
			End If
	End If				
	VerifyNewOfferType=bVerify	
End Function

'[Verify New Offer Label Value Offer Type]
Public Function VerifyNewOfferTypeVal(strOfferTypeVal)
	bVerify=True	
	If Not IsNull(strOfferTypeVal) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblOfferTypeval(),strOfferTypeVal,"Offer Type label Value")Then
				bVerify = False
			End If
	End If				
	VerifyNewOfferTypeVal=bVerify	
End Function

'[Verify table row Product Account No Or Card No displayed for New Offer]
Public Function VerifytblSelectedNewOfferPage(lstNewOfferAccCardNo)
	VerifytblSelectedNewOfferPage = VerifyTableSingleRowData(coSpecialOffer_Page.tblNewOfferProductAccountCardNoHeader,coSpecialOffer_Page.tblNewOfferProductAccountCardNoBody,lstNewOfferAccCardNo,"New Offer Product")
End Function

'[Verify New Offer Label Customer Decision]
Public Function VerifyNewOfferCustomerDecision(strCustDecision)
	bVerify=True	
	If Not IsNull(strCustDecision) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblCustDecision(),strCustDecision,"Offer Customer Decision label Name")Then
				bVerify = False
			End If
	End If				
	VerifyNewOfferCustomerDecision=bVerify	
End Function

'[Verify list of values displayed in Customer Decision dropdown]
Public Function VerifylstCustDecisionDropDwn(lstCustomerDecision)
	bVerifyValues = True
	scrollPageDown 3
	Wait(2)
	If Not IsNull(lstCustomerDecision) Then
		bVerifyValues = verifyComboboxItems1(coSpecialOffer_Page.lstCustomerDecisionObj,coSpecialOffer_Page.lstCustomerDecisionItems,lstCustomerDecision,"Customer Decision")
		'bVerifyValues = verifyComboboxItems(coSpecialOffer_Page.lstCustomerDecisionObj,lstCustomerDecision,"Customer Decision")
	End If
	VerifylstCustDecisionDropDwn = bVerifyValues	
End Function

'[Set Customer Decision dropdown as]
Public Function SetCustDecisionDropDwn(strItem)
	Wait(2)
	SetCustDecisionDropDwn=SelectComboBoxItem(coSpecialOffer_Page.lstCustomerDecisionObj,strItem,"Customer Decision")
	If Err.Number <> 0 Then 
		LogMessage "WARN","Verification","Failed to Set Customer Decision in text box", False
		SetCustDecisionDropDwn = False
		Exit Function
	Else
		SetCustDecisionDropDwn = True
	End If
End Function

'[Verify New Offer Label Reason for Customer Decision]
Public Function VerifyNewOfferReasonCustomerDecision(strCustDecisionReason)
	bVerifyReason=True	
	If Not IsNull(strCustDecisionReason) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblCustDecisionReason(),strCustDecisionReason,"Offer Customer Decision Reason label Name")Then
				bVerifyReason = False
			End If
	End If				
	VerifyNewOfferReasonCustomerDecision=bVerifyReason	
End Function

'[Verify list of values displayed in Reason Customer Decision dropdown]
Public Function VerifylstCustReasonDecisionDropDwn(lstReasonCustomerDecision)
	bVerifyReasonValues = True
	scrollPageDown 2
	Wait(2)
	If Not IsNull(lstReasonCustomerDecision) Then
		bVerifyReasonValues = verifyComboboxItems1(coSpecialOffer_Page.lstCustrDecsionReasnObj,coSpecialOffer_Page.lstCustomerDecisionReasonItems,lstReasonCustomerDecision,"Reason Customer Decision")	
		'bVerifyReasonValues = verifyComboboxItems(coSpecialOffer_Page.lstCustomerDecisionReasonObj,lstReasonCustomerDecision,"Reason Customer Decision")
	End If
	VerifylstCustReasonDecisionDropDwn = bVerifyReasonValues	
End Function

'[Set Reason Customer Decision dropdown as]
Public Function SetCustReasonDecisionDropDwn(strReasonCustomerDecision)
	SetCustReasonDecisionDropDwn=True
	If Not IsNull(strReasonCustomerDecision) Then
		Wait(2)
		SetCustReasonDecisionDropDwn=SelectComboBoxItem(coSpecialOffer_Page.lstCustomerDecisionReasonObj,strReasonCustomerDecision,"Reason Customer Decision")
	End If
	If Err.Number <> 0 Then 
		LogMessage "WARN","Verification","Failed to Set Reason Customer Decision in text box", False
		SetCustReasonDecisionDropDwn=False
		'Exit Function
	End If
End Function

'[Verify New Offer Label Other Reason for Customer Decision]
Public Function VerifyNewOfferOtherReasonCustomerDecision(strCustOtherDecisionReason)
	bVerifyOtherReason=True	
	If Not IsNull(strCustOtherDecisionReason) Then	
		If Not verifyInnerText(coSpecialOffer_Page.lblCustDecisionOtherReason(),strCustOtherDecisionReason,"Offer Customer Decision Other Reason label Name")Then
				bVerifyOtherReason = False
			End If
	End If				
	VerifyNewOfferOtherReasonCustomerDecision=bVerifyOtherReason	
End Function

'[Enter Other Reason field displayed in New Offers Page]
Public Function SetOtherReasonNewOff(strOthers)
	bVerifyOthOff = True
	If Not IsNull(strOthers) Then
			If not SetValue(coSpecialOffer_Page.txtCustDecisionOtherReasonObj(),strOthers,"New Offers Other Reason Text box") Then
					bVerifyOthOff = False
			End If
	End If
	SetOtherReasonNewOff = bVerifyOthOff
End Function

'[Verify Comments Label for New Offers Page]
Public Function VerifyCommentslblOffers(strComments)
	bVerifyComments=True	
	If Not IsNull(strComments) Then	
		If Not verifyInnerText(coSpecialOffer_Page.txtComments(),strComments,"New Offer Comments label Name")Then
				bVerifyComments = False
			End If
	End If				
	VerifyCommentslblOffers=bVerifyComments	
End Function

'[Enter Comments Optional field displayed in New Offers]
Public Function SetTxtComments(strTxtComment)
	bVerifyComments=True
	If Not IsNull(strTxtComment) Then
			If not SetValue(coSpecialOffer_Page.txtCommentsOffer(),strTxtComment,"New Offer Comments Text box") Then
				bVerifyComments=False
			End If
	End If
	SetTxtComments=bVerifyComments
End Function

'[Click on Cancel button in New Offers]
Public Function ClickCancelNewOffers()
	bVerifyCancel=True
	If coSpecialOffer_Page.btnCancelOffers.Exist(0) Then
		coSpecialOffer_Page.btnCancelOffers.Click
		If Err.Number <> 0 Then
			bVerifyCancel=False
		End If
	Else
		bVerifyCancel=False
	End If
	ClickCancelNewOffers=bVerifyCancel
End Function

'[Verify the Cancellation message in New Offers]
Public Function VerifyCancelationMsgNewOffers(YesOrNo,strMsg)
	bVerify = False
	bVerify1 = True
	bVerify2 = True

	If coSpecialOffer_Page.btnCancelOffers.Exist(0) Then
		coSpecialOffer_Page.btnCancelOffers.Click
			If Err.Number <> 0 Then
				bVerify1 = False
			End If
	Else
	bVerify1 = False
	End If
	
	If Not IsNull(strMsg) Then
			If coSpecialOffer_Page.lblConfirm.Exist(0) Then
				strActMsg = coSpecialOffer_Page.lblConfirm.GetRoProperty("innertext")
				If Ucase(Trim(strMsg)) = Ucase(Trim(strActMsg)) Then
					bVerify = True
				End If
			End If
	End If
	
	If Not IsNull(YesOrNo) Then
		If Ucase(YesOrNo) = Ucase("Yes") Then
			If coSpecialOffer_Page.btnConfirmYes.Exist(0) Then
				coSpecialOffer_Page.btnConfirmYes.Click
				If Err.Number <> 0 Then
				bVerify2 = False
				End If
			End If
		ElseIf Ucase(YesOrNo) = Ucase("No") Then
			If coSpecialOffer_Page.btnConfirmNo.Exist(0) Then
				coSpecialOffer_Page.btnConfirmNo.Click
				If Err.Number <> 0 Then
				bVerify2 = False
				End If
			Else
			bVerify2 = False
			End If
		Else
			bVerify2 = False
		End If
	Else
			bVerify2 = False
	End If	
	If bVerify and bVerify1 and bVerify2 Then
		VerifyCancelationMsgNewOffers=True
	Else
		VerifyCancelationMsgNewOffers=False
	End If
End Function

'[Click on Submit button in New Offers]
Public Function ClickSubmitNewOffers(strSubmit)
	bVerifySubmit=True
	If Not IsNull(strSubmit) Then
		If coSpecialOffer_Page.btnSubmitOffers.Exist(0) Then
			coSpecialOffer_Page.btnSubmitOffers.Click
			If Err.Number <> 0 Then
				bVerifySubmit=False
			End If
		Else
			bVerifySubmit=False
		End If
	End IF
	ClickSubmitNewOffers=bVerifySubmit
End Function

'[Verify Submit Button Displayed in New Offers]
Public Function VerifySubmitNewOffers(strstatus)
    bVerify = True
	If Not IsNull(strstatus) Then
		If coSpecialOffer_Page.btnSubmitOffers.Exist(0) Then
			stractStatus = coSpecialOffer_Page.btnSubmitOffers.GetRoProperty("disabled")
			If strstatus="Enable" Then
				If stractStatus=0 Then
					LogMessage "RSLT","Verification","Submit Button is in Enabled Mode", True
				Else
					LogMessage "RSLT","Verification","Submit Button is in Disabled Mode", False
					bVerify=False
				End If
			ElseIf strstatus="Disable" Then
				If stractStatus=1 Then
					LogMessage "RSLT","Verification","Submit Button is in Disabled Mode", True
				Else
					LogMessage "RSLT","Verification","Submit Button is in Enabled Mode", False
					bVerify=False
				End If
			Else
				bVerify=False
			End If
		Else
			bVerify=False
		End If
	End If
	VerifySubmitNewOffers=bVerify
End Function

'[Verify Submission popup Message displayed in New Offers]
Public Function VerifySubmissionInNewOffers(SubMsg)
	bVerify = True
	If Not IsNull(SubMsg) Then
		If verifyInnerText(coSpecialOffer_Page.lblSubmissionMsg(),SubMsg,"Submission Message in Rewards Redemption") Then
			bVerify=True
		End If
		If coSpecialOffer_Page.btnOK.Exist(0) Then
			coSpecialOffer_Page.btnOK.Click
			If Err.Number <> 0 Then
				bVerify=False
			End If
		Else
			bVerify=False
		End If
	End If
	VerifySubmissionInNewOffers = bVerify
End Function

'[Change the contact CIF status in Database]
Public Function updateCIFDBCellValue(strGIFno) 
	bupdateDBCellValue=true         
	strUpdateQuery=updateCellValinDB("update iserve_offers_main set STATUS = 'OPEN' where CONTACT_CIF = '"&strGIFno&"'")
	strValue=getDBValForColumn_FE("select CONTACT_CIF from iserve_offers_main where CONTACT_CIF = '"&strGIFno&"'")(0)
	If not (strValue = strGIFno) Then 
		LogMessage "WARN","Verification","Failed to update cell value to change CONTACT_CIF STATUS." ,false 
		bupdateDBCellValue=false 
	End If 
	updateCIFDBCellValue=bupdateDBCellValue 
End Function

'[Verify Submission popup Message displayed for send SMS option in New Offers Page]
Public Function VerifySubmissionSMSOptionInNewOffers(SubMsg,SubMsgOption)
	'WaitForIServeLoading
	bVerify = True
	If Not IsNull(SubMsg) Then
		If Not verifyInnerText(coSpecialOffer_Page.lblSendSMSSubmissionMsg(),SubMsg,"Submission Message in Offers") and Not verifyInnerText(coSpecialOffer_Page.lblSendSMSOptionSubmissionMsg(),SubMsgOption,"Submission SMS Message in Offers") Then
			bVerify=false
		End If
		If coSpecialOffer_Page.btnSendSMSYes.Exist(0) Then
			coSpecialOffer_Page.btnSendSMSYes.Click
			If Err.Number <> 0 Then
				bVerify=False
			End If
		Else
			bVerify=False
		End If
	End If
	VerifySubmissionSMSOptionInNewOffers=bVerify
End Function
