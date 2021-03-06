'[Verify Left Panel details in Pin Reset section]
Public Function verifyPinResetLeftPanelDetailsSection(strCustCIF,strPhoneBankingNumber,strPINStatus)
	
	blnverifyPinResetLeftPanelDetailsSection=True
	

	If strCustCIF <>"" Then
			If Not verifyInnerText_Pattern(HK_PB_PinReset_Page.welePinResetCIFNo(), strCustCIF, "CIF No") Then
				blnverifyPinResetLeftPanelDetailsSection=False
			End If
	End If	
	

	If strPhoneBankingNumber <>"" Then
			If Not verifyInnerText_Pattern(HK_PB_PinReset_Page.welePinResetPhnBnkNumber(), strPhoneBankingNumber, "Phone Banking Number") Then
				blnverifyPinResetLeftPanelDetailsSection=False
			End If
	End If
	

	If strPINStatus <>"" Then
			If Not verifyInnerText_Pattern(HK_PB_PinReset_Page.welePinResetPinStatus(), strPINStatus, "PIN Status") Then
				blnverifyPinResetLeftPanelDetailsSection=False
			End If
	End If
	

	verifyPinResetLeftPanelDetailsSection=blnverifyPinResetLeftPanelDetailsSection

End Function
'[Set comments in Pin Reset comments box]
Public Function SetPinResetComments(strComments)
	blnSetPinResetComments=true
	HK_PB_PinReset_Page.txtPinResetComments().Set strComments
	If Err.Number<>0 Then
		blnSetPinResetComments=false
		LogMessage "WARN","Verification","Failed to set comments" ,false
	Else
		LogMessage "RSLT","Verification","set the comments as expected",True
		blnSetPinResetComments=true
	End If
	SetPinResetComments= blnSetPinResetComments
End function
'[Verify Max Length of Pin Reset comments box]
Public Function verifyPinResetCommentsMaxLength()
	blnverifyPinResetCommentsMaxLength=true
	blnverifyPinResetCommentsMaxLength= VerifyMaxLength (HK_PB_PinReset_Page.txtPinResetComments(),"1351","Pin Reset comments box")
	verifyPinResetCommentsMaxLength=blnverifyPinResetCommentsMaxLength
End Function
'[Verify Confirmation message is displayed Upon PIN Reset]
Public Function verifyConfirmationMessagePinReset()
	blnverifyConfirmationMessagePinReset = true

	If Not VerifyInnerText (HK_PB_CreateProfile_Page.welestpTMApproveCntspan(),"This SR will be routed to Team Manager for following reason(s): PIN Re-set.", "Confirmation  message PIN Reset") Then
		blnverifyConfirmationMessagePinReset = false
	End If

	verifyConfirmationMessagePinReset = blnverifyConfirmationMessagePinReset
End Function
'[Verify validation message is displayed Upon clicking Pin Reset Button]
Public Function verifyValidationMessagePinReset()
	blnverifyValidationMessagePinReset = true
	If Not IsNull(strErrorMessage) Then
		If Not VerifyInnerText (HK_PB_PinReset_Page.welePinResetValidationMsg(),"The request cannot be processed, as the PIN mailer is not yet dispatched", "Validation Message Pin Reset") Then
			blnverifyValidationMessagePinReset = false
		End If
	End If
	verifyValidationMessagePinReset = blnverifyValidationMessagePinReset
End Function
'[Click on SR Status Ok Button]
Public Function clickSRStatusOkButton()
	blnclickSRStatusOkButton=true
	HK_PB_PinReset_Page.btnSRStatusOk().Click
	WaitForICallLoading
	If Err.Number<>0 Then
		blnclickSRStatusOkButton=false
		LogMessage "WARN","Verification","Failed to Click Button :Ok" ,false
	Else
		LogMessage "RSLT","Verification","Clicked on Confirmation - Ok Button as expected.",True
		blnclickSRStatusOkButton=true
	End If
	clickSRStatusOkButton=blnclickSRStatusOkButton
End Function
