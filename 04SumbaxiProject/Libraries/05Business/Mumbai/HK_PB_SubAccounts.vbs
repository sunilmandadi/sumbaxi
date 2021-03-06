'[Verify row Data in Table for Sub Accounts Tagged in Phone Banking Page]
Public Function verifytblContentSubAccountsTaggedPhoneBanking(arrSubAccountsTaggedDataList)

   verifytblContentSubAccountsTaggedPhoneBanking=verifyTableContentList(HK_PB_SubAccounts_Page.tblPBAccountsTaggedHeader(),HK_PB_SubAccounts_Page.tblPBAccountsTaggedContent(),arrSubAccountsTaggedDataList,"Sub Accounts Tagged - Phone Banking",false,NULL,NULL,NULL)
End Function
'[Verify txt CIFNo for Sub Accounts Mapped]
Public Function verifySubAccMappedCIFNoText(strCIFNoText)
	blnverifySubAccMappedCIFNoText=true
	If Not verifyInnerText_Pattern(HK_PB_SubAccounts_Page.weleSubAccCIFNO(), strCIFNoText, "Sub Accounts Mapped CIFNo Text") Then
		blnverifySubAccMappedCIFNoText=false
	End If
	verifySubAccMappedCIFNoText=blnverifySubAccMappedCIFNoText
End Function
'[Verify txt Phone Banking Number for Sub Accounts Mapped]
Public Function verifySubAccMappedhoneBankingNumberText(strPBNumberText)
	blnverifySubAccMappedhoneBankingNumberText=true
	If Not verifyInnerText_Pattern(HK_PB_SubAccounts_Page.weleSubAccPhonebankNumber(), strPBNumberText, "Sub Accounts Mapped Phone Banking Text") Then
		blnverifySubAccMappedhoneBankingNumberText=false
	End If
	verifySubAccMappedhoneBankingNumberText=blnverifySubAccMappedhoneBankingNumberText
End Function
'[Set comments in Maintain Sub Accounts Mapped comments box]
Public Function MaintainSubAccounts_setComments(strComments)
	blnMaintainSubAccounts_setComments=true
	HK_PB_CreateProfile_Page.txtDeleteProfileComments().Set strComments
	If Err.Number<>0 Then
		blnMaintainSubAccounts_setComments=false
		LogMessage "WARN","Verification","Failed to set comments" ,false
	Else
		LogMessage "RSLT","Verification","set the comments as expected",True
		blnMaintainSubAccounts_setComments=true
	End If
	MaintainSubAccounts_setComments= blnMaintainSubAccounts_setComments
End function
'[Select Combobox Opt-in/Opt-out in Sub Accounts]
Public Function selectOptinOputoutComboBox(strStatus)
	bDevPending=false

	blnselectOptinOputoutComboBox=true
	If Not IsNull(strStatus) Then
		If Not (selectItem_Combobox (HK_PB_SubAccounts_Page.lstOptinOptout(),strStatus))Then
			LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Show drop down list" ,false
			blnselectOptinOputoutComboBox=false
		End If
	End If
	WaitForICallLoading
	selectOptinOputoutComboBox=blnselectOptinOputoutComboBox
End Function
'[Verify Confirmation message is displayed Upon Sub Accounts Mapped Submission]
Public Function verifyConfirmationSubAccountsMapped()
	blnverifyConfirmationSubAccountsMapped = true
	If Not IsNull(strErrorMessage) Then
		If Not VerifyInnerText (HK_PB_CreateProfile_Page.welestpTMApproveCntspan(),"This SR will be routed to Team Manager for following reason(s): Request for Opt-in/Opt-out of Sub Accounts.", "Confirmation  message Sub Accounts") Then
			blnverifyConfirmationSubAccountsMapped = false
		End If
	End If
	verifyConfirmationSubAccountsMapped = blnverifyConfirmationSubAccountsMapped
End Function
'[Verify error message is displayed Upon opting out in Maintain Sub Accounts Screen]
Public Function verifyErrorMessageMaintainSubAcc(strErrorMessage)
	blnverifyErrorMessageMaintainSubAcc = true
	If Not IsNull(strErrorMessage) Then
		If Not VerifyInnerText (HK_PB_SubAccounts_Page.weleOptoutErrMessage(),strErrorMessage, "Error Message Opt out") Then
			blnverifyErrorMessageMaintainSubAcc = false
		End If
	End If
	verifyErrorMessageMaintainSubAcc = blnverifyErrorMessageMaintainSubAcc
End Function
