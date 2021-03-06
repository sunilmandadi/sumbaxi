'[Select Bank Account Type in Maintain Third Party Account]
Public Function setAccountTypeTxtBox(strType)

	blnsetAccountTypeTxtBox=true

	HK_PB_ThirdPartyAccount_Page.txtAccountType().Set strType
	If Err.Number<>0 Then
		blnsetAccountTypeTxtBox=false
		LogMessage "WARN","Verification","Failed to set Account Type" ,false
	Else
		LogMessage "RSLT","Verification","set the Account Type as expected",True
		blnsetAccountTypeTxtBox=true
	End If
	setAccountTypeTxtBox=blnsetAccountTypeTxtBox
End Function

'[Set Bank Account Number in Maintain Third Party Account]
Public Function setAccountNumberTxtBox(strAccountNumber)

	blnsetAccountNumberTxtBox=true

	HK_PB_ThirdPartyAccount_Page.txtAccountNumber().Set strAccountNumber
	If Err.Number<>0 Then
		blnsetAccountNumberTxtBox=false
		LogMessage "WARN","Verification","Failed to set Account Number" ,false
	Else
		LogMessage "RSLT","Verification","set the Account Number as expected",True
		blnsetAccountNumberTxtBox=true
	End If
	setAccountNumberTxtBox=blnsetAccountNumberTxtBox
End Function

'[Set Bank Account Name in Maintain Third Party Account]
Public Function setAccountNameTxtBox(strACName)

	blnsetAccountNameTxtBox=true

	HK_PB_ThirdPartyAccount_Page.txtAccountType().Set strACName
	If Err.Number<>0 Then
		blnsetAccountNameTxtBox=false
		LogMessage "WARN","Verification","Failed to set Account Name" ,false
	Else
		LogMessage "RSLT","Verification","set the Account Name as expected",True
		blnsetAccountNameTxtBox=true
	End If
	setAccountNameTxtBox=blnsetAccountNameTxtBox
End Function

'[Set Branch Code in Maintain Third Party Account]
Public Function setBranchCodeTxtBox(strBranchCode)

	blnsetBranchCodeTxtBox=true

	HK_PB_ThirdPartyAccount_Page.txtBranchCode().Set strBranchCode
	If Err.Number<>0 Then
		blnsetBranchCodeTxtBox=false
		LogMessage "WARN","Verification","Failed to set Branch code" ,false
	Else
		LogMessage "RSLT","Verification","set the Branch Code as expected",True
		blnsetBranchCodeTxtBox=true
	End If
	setBranchCodeTxtBox=blnsetBranchCodeTxtBox
End Function

'[Click on Add Button in Third Party Account Page]
Public Function clickAddButtonThirdPartyAcc()
	ClickOnObject HK_PB_ThirdPartyAccount_Page.btnAddThirdPartyAcc(),"Add Button"
End Function

'[Click on Remove Button in Third Party Account Page]
Public Function clickRemoveButtonThirdPartyAcc()
	ClickOnObject HK_PB_ThirdPartyAccount_Page.btnRemoveThirdPartyAcc(),"Remove Button"
End Function

'[Verify txt CIFNo for Third Party Account]
Public Function verifyThirdPartyAccountCIFNoText(strCIFNoText)
	blnverifyThirdPartyAccountCIFNoText=true
	If Not verifyInnerText_Pattern(HK_PB_ThirdPartyAccount_Page.weleThirdPartyAccountCIFNo(), strCIFNoText, "Third Party Account CIFNo Text") Then
		blnverifyThirdPartyAccountCIFNoText=false
	End If
	verifyThirdPartyAccountCIFNoText=blnverifyThirdPartyAccountCIFNoText
End Function
'[Verify txt Phone Banking Number for Third Party Account]
Public Function verifyThirdPartyAccountPhoneBankingNumberText(strPBNumberText)
	blnverifyThirdPartyAccountPhoneBankingNumberText=true
	If Not verifyInnerText_Pattern(HK_PB_ThirdPartyAccount_Page.weleThirdPartyAccountPhnBnkNumber(), strPBNumberText, "Third Party Account Phone Banking Text") Then
		blnverifyThirdPartyAccountPhoneBankingNumberText=false
	End If
	verifyThirdPartyAccountPhoneBankingNumberText=blnverifyThirdPartyAccountPhoneBankingNumberText
End Function
'[Verify txt Description for Third Party Account]
Public Function verifyThirdPartyAccountDescriptionText(strDescriptionText)
	blnverifyThirdPartyAccountDescriptionText=true
	If Not verifyInnerText_Pattern(HK_PB_ThirdPartyAccount_Page.weleThirdPartyAccountDesc(),strDescriptionText,"Third Party Account Description Text") Then
		blnverifyThirdPartyAccountDescriptionText=false
	End If
	verifyThirdPartyAccountDescriptionText=blnverifyThirdPartyAccountDescriptionText
End Function
'[Set comments in Third Party Account comments box]
Public Function ThirdPartyAccount_setComments(strComments)
	blnThirdPartyAccount_setComments=true
	HK_PB_CreateProfile_Page.txtDeleteProfileComments().Set strComments
	If Err.Number<>0 Then
		blnThirdPartyAccount_setComments=false
		LogMessage "WARN","Verification","Failed to set comments" ,false
	Else
		LogMessage "RSLT","Verification","set the comments as expected",True
		blnThirdPartyAccount_setComments=true
	End If
	ThirdPartyAccount_setComments= blnThirdPartyAccount_setComments
End function
'[Verify Confirmation message is displayed Upon Maintain Third Party Accounts Submission]
Public Function verifyConfirmationMaintainThirdPartyAccount()
	blnverifyConfirmationMaintainThirdPartyAccount = true
	If Not IsNull(strErrorMessage) Then
		If Not VerifyInnerText (HK_PB_CreateProfile_Page.welestpTMApproveCntspan(),"This SR will be routed to Team Manager for following reason(s): Third Party Account(s) Add / Delete.", "Confirmation  message Sub Accounts") Then
			blnverifyConfirmationMaintainThirdPartyAccount = false
		End If
	End If
	verifyConfirmationMaintainThirdPartyAccount = blnverifyConfirmationMaintainThirdPartyAccount
End Function
'[Verify row Data in Table for Third Party Accounts Linked in Phone Banking Page]
Public Function verifytblContentThirdPartyAccTaggedPhoneBanking(arrThirdPartyLinkedDataList)
   verifytblContentThirdPartyAccTaggedPhoneBanking=verifyTableContentList(HK_PB_ThirdPartyAccount_Page.tblThirdPartyAccHeader(),HK_PB_ThirdPartyAccount_Page.tblThirdPartyAccContent(),arrThirdPartyLinkedDataList,"Third Party Accounts Linked - Phone Banking",false,NULL,NULL,NULL)
End Function
