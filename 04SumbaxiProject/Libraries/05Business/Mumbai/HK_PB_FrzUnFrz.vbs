'[Click on Freeze Unfreeze Button]
Public Function clickFreezeUnFreezeButton()
	ClickOnObject HK_PB_CreateProfile_Page.btnFreezeUnfreezePin(),"Freeze Unfreeze Button"
	WaitForICallLoading
End Function
'[Verify Freeze/Unfreeze Action Combobox has Items]
Public Function verifyFreezeUnfreezeActionComboboxItems(lstItems)
	blnverifyFreezeUnfreezeActionComboboxItems=true
	
	If Not IsNull(lstItems) Then
		If Not verifyComboboxItems (HK_PB_FrzUnFrz_Page.lstFreezeUnfreezeAction(),lstItems,"Freeze/Unfreeze Action") Then
			blnverifyFreezeUnfreezeActionComboboxItems=false
		End If
	End If
	verifyFreezeUnfreezeActionComboboxItems=blnverifyFreezeUnfreezeActionComboboxItems
End Function
'[Select Combobox Freeze/Unfreeze Action]
Public Function selectFreezeUnfreezeActionComboBox(strShow)
	bDevPending=false

	blnselectFreezeUnfreezeActionComboBox=true
	If Not IsNull(strShow) Then
		If Not (selectItem_Combobox (HK_PB_FrzUnFrz_Page.lstFreezeUnfreezeAction(),strShow))Then
			LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Show drop down list" ,false
			blnselectFreezeUnfreezeActionComboBox=false
		End If
	End If
	WaitForICallLoading
	selectFreezeUnfreezeActionComboBox=blnselectFreezeUnfreezeActionComboBox
End Function
'[Verify Left Panel details in Freeze UnFreeze section]
Public Function verifyFrzUnFrzLeftPanelDetailsSection(strCIFNo,strPhoneBankingNumber,strPINStatus)
	
	blnverifyFrzUnFrzLeftPanelDetailsSection=True
	

	If strCIFNo <>"" Then
			If Not verifyInnerText_Pattern(HK_PB_FrzUnFrz_Page.weleFrUnFrzCIFNo(), strCIFNo, "CIF No") Then
				blnverifyPhoneBankingDetailsSection=False
			End If
	End If	
	

	If strPhoneBankingNumber <>"" Then
			If Not verifyInnerText_Pattern(HK_PB_FrzUnFrz_Page.weleFrUnFrzPhnBnkNumber(), strPhoneBankingNumber, "Phone Banking Number") Then
				blnverifyPhoneBankingDetailsSection=False
			End If
	End If
	

	If strPINStatus <>"" Then
			If Not verifyInnerText_Pattern(HK_PB_FrzUnFrz_Page.weleFrUnFrzPinStatus(), strPINStatus, "PIN Status") Then
				blnverifyPhoneBankingDetailsSection=False
			End If
	End If
	

	verifyFrzUnFrzLeftPanelDetailsSection=blnverifyFrzUnFrzLeftPanelDetailsSection

End Function
'[Verify error message is displayed Upon selecting Action dropdown in Freeze UnFreeze Pin Screen]
Public Function verifyErrorMessageActionFrzUnFrz()
	blnverifyErrorMessageActionFrzUnFrz = true
	If Not IsNull(strErrorMessage) Then
		If Not VerifyInnerText (HK_PB_FrzUnFrz_Page.weleFrUnFrzActionErrMessage(),"PIN Status is not valid for this request.", "Error Message FrzUnFrz Action") Then
			blnverifyErrorMessageActionFrzUnFrz = false
		End If
	End If
	verifyErrorMessageActionFrzUnFrz = blnverifyErrorMessageActionFrzUnFrz
End Function
'[Set comments in Freeze UnFreeze comments box]
Public Function SetFrzUnFrzComments(strComments)
	blnSetFrzUnFrzComments=true
	HK_PB_FrzUnFrz_Page.txtFrzUnFrzComments().Set strComments
	If Err.Number<>0 Then
		blnSetFrzUnFrzComments=false
		LogMessage "WARN","Verification","Failed to set comments" ,false
	Else
		LogMessage "RSLT","Verification","set the comments as expected",True
		blnSetFrzUnFrzComments=true
	End If
	SetFrzUnFrzComments= blnSetFrzUnFrzComments
End function
'[Verify Max Length of Freeze UnFreeze comments box]
Public Function verifyFrUnFrzCommentsMaxLength()
	blnverifyFrUnFrzCommentsMaxLength=true
	blnverifyFrUnFrzCommentsMaxLength= VerifyMaxLength (HK_PB_FrzUnFrz_Page.txtFrzUnFrzComments(),"1351","Freeze UnFreeze comments box")
	verifyFrUnFrzCommentsMaxLength=blnverifyFrUnFrzCommentsMaxLength
End Function
'[Verify Confirmation message is displayed Upon PIN Freeze]
Public Function verifyConfirmationMessagePinFreeze()
	blnverifyConfirmationMessagePinFreeze = true

	If Not VerifyInnerText (HK_PB_CreateProfile_Page.welestpTMApproveCntspan(),"This SR will be routed to Team Manager for following reason(s): Request for PIN Freeze.", "Confirmation  message PIN Freeze") Then
		blnverifyConfirmationMessagePinFreeze = false
	End If

	verifyConfirmationMessagePinFreeze = blnverifyConfirmationMessagePinFreeze
End Function
'[Select pending approval record from service request queue]
Public Function viewPendingSRDetailsFromOverviewPage(strlstOverviewPendingSRDetails)
	blnviewPendingSRDetailsFromOverviewPage = True
	selectItem_Combobox HK_PB_FrzUnFrz_Page.lstShowCmbBox(),"Service Request"
	WaitForICallLoading
	SelectRadioButtonGrp "Pending Approval", HK_PB_FrzUnFrz_Page.rbtnStatus, Array("Open","Pending","Failed","Pending Approval","Closed")
	WaitForICallLoading
	set objtblcontent=Browser("Browser_IServe").Page("Page_PhoneBanking_OldObjects").WebElement("tblPendingSRContent_2")

	blnviewPendingSRDetailsFromOverviewPage=selectTableLink(HK_PB_FrzUnFrz_Page.tblOverviewTablePendingSRHeader(),objtblcontent,strlstOverviewPendingSRDetails,"Pending SR Search" ,"Type",true,HK_PB_FrzUnFrz_Page.lnkNextOverviewTablePendingSR(),HK_PB_FrzUnFrz_Page.lnkNextOverviewTablePendingSR1(),HK_PB_FrzUnFrz_Page.lnkPrvOverviewTablePendingSR())
	WaitForICallLoading
	viewPendingSRDetailsFromOverviewPage=blnviewPendingSRDetailsFromOverviewPage
End Function
