'[Click Rewards Redemption link displayed in Rewards Accordion in Credit Cards]
Public Function ClickRRLink_CR()
bverify = True
	Wait 2
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
	WaitForIServeLoading
	If coRewardRedem_Page.lnkRewardRedemption.Exist(0) Then
		coRewardRedem_Page.lnkRewardRedemption.Click
			If Err.Number <> 0 Then
				bverify = False
			End If
	Else
	bverify = False
	End If
ClickRRLink_CR = bverify
End Function

'[Verify the Selected Card Details in the Reward Redemption Page]
Public Function VerifyTblSelectCardDetailsRR(lstCardDetails)
bVerify = False
If Not IsNull(lstCardDetails) Then
bVerify = VerifyTableSingleRowData(coRewardRedem_Page.tblSelectedCardHeader,coRewardRedem_Page.tblSelectedCardBody,lstCardDetails,"Rewards Redemption")		
End If
VerifyTblSelectCardDetailsRR = bVerify
End Function

'[Select option from CATAGORY Dropdown as]
Public Function SelectCategoryDropdown(strItem)
bVerifyValues = False
WaitForIServeLoading
If Not IsNull(strItem) Then
bVerifyValues = SelectComboBoxItem(coRewardRedem_Page.lstCategory,strItem,"Rewards Option")		
End If
SelectCategoryDropdown = bVerifyValues
End Function

'[Select option from PRODUCT Dropdown as]
Public Function SelectProductDropdown(strItem)
bVerifyValues = False
WaitForIServeLoading
If Not IsNull(strItem) Then
bVerifyValues = SelectComboBoxItem(coRewardRedem_Page.lstProduct,strItem,"Rewards Option")		
End If
SelectProductDropdown = bVerifyValues
End Function

'[Click on Submit button in RR]
Public Function ClickSubmitRR()
bVerify = True
	If coRewardRedem_Page.btnSubmit.Exist(0) Then
		coRewardRedem_Page.btnSubmit.Click
			If Err.Number <> 0 Then
				bVerify = False
			End If
	Else
	bVerify = False
	End If
		ClickSubmitRR = bVerify
End Function

'[Enter Comments field displayed in Rewards Redemption]
Public Function SetcommentsRR(strComment)
bVerify = False
	If Not IsNull(strComment) Then
			If SetValue(coRewardRedem_Page.txtComments(),strComment,"Rewards Redemption Comment Text box") Then
					bVerify = True
			End If
	End If
	SetcommentsRR = bVerify
End Function

'[Verify Points Or Rebates displayed as]
Public Function VerifyRebates(strVal)
bVerify = False
	If Not IsNull(strVal) Then
		If verifyInnerText(coRewardRedem_Page.lblRebates(),strVal,"Rewards Option Rebates") Then
			bVerify=True
		End If
	End If
VerifyRebates = bVerify
End Function

'[Verify field description displayed in RR]
Public Function VerifyDescriptionRR(strDesc)
bVerify = False
	If Not IsNull(strDesc) Then
		If verifyInnerText(coRewardRedem_Page.lblDescription(),strDesc,"Rewards Redemption") Then
			bVerify=True
		End If
	End If
VerifyDescriptionRR = bVerify
End Function

'[Verify link Knowledge Base displayed in RR]
Public Function VerifyKnowledgebaseRR(strstatus)
bVerify = True
	If Not IsNull(strstatus) Then
		If coRewardRedem_Page.lnkKnowledgeBase.Exist(0) Then
			stractStatus = coRewardRedem_Page.lnkKnowledgeBase.GetRoProperty("disabled")
			If strstatus="Enable" Then
				If stractStatus=0 Then
					LogMessage "RSLT","Verification","Knowledge Base Button is in Enabled Mode", True
				Else
					LogMessage "RSLT","Verification","Knowledge Base Button is in Disabled Mode", False
					bVerify=False
				End If
			ElseIf strstatus="Disable" Then
				If stractStatus=1 Then
					LogMessage "RSLT","Verification","Knowledge Base Button is in Disabled Mode", True
				Else
					LogMessage "RSLT","Verification","Knowledge Base Button is in Enabled Mode", False
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
	VerifyKnowledgebaseRR = bVerify
End Function

'[Verify Submit Button Displayed in RR]
Public Function VerifySubmit_RR(strstatus)
bVerify = True
	If Not IsNull(strstatus) Then
		If coRewardRedem_Page.btnSubmit.Exist(0) Then
			stractStatus = coRewardRedem_Page.btnSubmit.GetRoProperty("disabled")
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
	Else
		bVerify=False
	End If
	VerifySubmit_RR = bVerify
End Function

'[Verify Inline Message in Reward Redemption page]
Public Function VerifyMsgInRewardRedeem(inLineMsg)
bVerify = False
	If Not IsNull(inLineMsg) Then
		If verifyInnerText(coRewardRedem_Page.lblInlineMsg(),inLineMsg,"Rewards Redemption") Then
			bVerify=True
		End If
	End If
	VerifyMsgInRewardRedeem = bVerify
End Function

'[Verify the Delivary Address Details]
Public Function VerifyDAInRR(Address)
bVerify = False
	If Not IsNull(Address) Then
		If verifyInnerText(coRewardRedem_Page.lblAddressVal1(),Address,"Delivary Address in Rewards Redemption") Then
			bVerify=True
		End If
	End If
	VerifyDAInRR = bVerify
End Function

'[Verify Submission popup Message displayed in RR]
Public Function VerifySubmissionInRR(SubMsg)
bVerify = False
	If Not IsNull(SubMsg) Then
		If verifyInnerText(coRewardRedem_Page.lblSubmissionMsg(),SubMsg,"Submission Message in Rewards Redemption") Then
			bVerify=True
		End If
	End If
	If coRewardRedem_Page.btnOK.Exist(0) Then
		coRewardRedem_Page.btnOK.Click
		If Err.Number <> 0 Then
			bVerify=False
		End If
	Else
	  bVerify=False
	End If
	VerifySubmissionInRR = bVerify
End Function

'[Verify the Cancellation message in Rewards Redemption]
Public Function VerifyCancelationMsg_TL(YesOrNo,strMsg)
bVerify = False
bVerify1 = True
bVerify2 = True

	If coRewardRedem_Page.lnkCancel.Exist(0) Then
		coRewardRedem_Page.lnkCancel.Click
			If Err.Number <> 0 Then
				bVerify1 = False
			End If
	Else
	bVerify1 = False
	End If
	
	If Not IsNull(strMsg) Then
			If coRewardRedem_Page.lblConfirm.Exist(0) Then
				strActMsg = coRewardRedem_Page.lblConfirm.GetRoProperty("innertext")
				If Ucase(Trim(strMsg)) = Ucase(Trim(strActMsg)) Then
					bVerify = True
				End If
			End If
	End If
	
	If Not IsNull(YesOrNo) Then
		If Ucase(YesOrNo) = Ucase("Yes") Then
			If coRewardRedem_Page.btnConfirmYes.Exist(0) Then
				coRewardRedem_Page.btnConfirmYes.Click
				If Err.Number <> 0 Then
				bVerify2 = False
				End If
			Else
			bVerify2 = False
			End If
		ElseIf Ucase(YesOrNo) = Ucase("No") Then
			If coRewardRedem_Page.btnConfirmNo.Exist(0) Then
				coRewardRedem_Page.btnConfirmNo.Click
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
		VerifyCancelationMsg_TL = True
	Else
		VerifyCancelationMsg_TL = False
	End If

End Function

