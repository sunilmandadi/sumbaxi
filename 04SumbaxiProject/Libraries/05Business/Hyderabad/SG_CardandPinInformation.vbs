'[Verify the card summary dislayed as on card and pin info page]
Public Function verifyCardSumm_CardPinInfo(lstCardSumm)
	bverifyCardSumm_CardPinInfo = true
	intSize = Ubound(lstCardSumm)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstCardSumm(Iterator),":")(0))
		arrValue = trim(Split(lstCardSumm(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "Card Number"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_CardNumber(), arrValue, "Card Number")Then
				LogMessage "RSLT","Verification","Card Summary - Card Number:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Embossed Name"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_EmbossedName(), arrValue, "Embossed Name")Then
				LogMessage "RSLT","Verification","Card Summary - Embossed Name:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Card Status"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_CardStatus(), arrValue, "Card Status")Then
				LogMessage "RSLT","Verification","Card Summary - Card Status:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Action Status"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_ActionStatus(), arrValue, "Action Status")Then
				LogMessage "RSLT","Verification","Card Summary - Action Status:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Reason"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_Reason(), arrValue, "Reason")Then
				LogMessage "RSLT","Verification","Card Summary - Reason:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Date and Time"
		If Not IsNull(arrValue) Then
		arrvalue_new = Replace(arrValue,"@",":")
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_DateTime(), arrvalue_new, "Date and Time")Then
				LogMessage "RSLT","Verification","Card Summary - Date and Time:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		Case "Tagged By"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_TaggedBy(), arrValue, "Tagged By")Then
				LogMessage "RSLT","Verification","Card Summary - Tagged By:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Brand"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_Brand(), arrValue, "Brand")Then
				LogMessage "RSLT","Verification","Card Summary - Brand:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "PIN Tries"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_PINTries(), arrValue, "PIN Tries")Then
				LogMessage "RSLT","Verification","Card Summary - PIN Tries:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "PIN Issued"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_PINIssued(), arrValue, "PIN Issued")Then
				LogMessage "RSLT","Verification","Card Summary - PIN Issued:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If
		
		Case "Last PIN Issued Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblSummary_LastPINIssuedDate(), arrValue, "Last PIN Issued Date")Then
				LogMessage "RSLT","Verification","Card Summary - Last PIN Issued Date:"&arrValue&" is not displayed as expected",false
				bverifyCardSumm_CardPinInfo=false
			End If
		End If	
		End Select
		Next
		verifyCardSumm_CardPinInfo = bverifyCardSumm_CardPinInfo
End Function

'[Verify the Card replacement history displayed as]
Public Function verifyCardReplaceHist_CardPinInfo(lstCardReplace)
	bverifyCardReplaceHist_CardPinInfo = true
	intSize = Ubound(lstCardReplace)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstCardReplace(Iterator),":")(0))
		arrValue = trim(Split(lstCardReplace(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "New Card Number"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblReplaceHist_NewCardNumber(), arrValue, "New Card Number")Then
				LogMessage "RSLT","Verification","Replacement History - New Card Number:"&arrValue&" is not displayed as expected",false
				bverifyCardReplaceHist_CardPinInfo=false
			End If
		End If
		
		Case "Old Card Number"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcCardAndInfoold.lblReplaceHist_OldCardNumber(), arrValue, "Old Card Number")Then
				LogMessage "RSLT","Verification","Replacement History - Old Card Number:"&arrValue&" is not displayed as expected",false
				bverifyCardReplaceHist_CardPinInfo=false
			End If
		End If
		End Select
		Next
		verifyCardReplaceHist_CardPinInfo = bverifyCardReplaceHist_CardPinInfo
End Function

'[Verify the card details displayed as for card and pin info]
Public Function verifyCardDetails_CardPinInfo(strDetails_OverseasWdl,strDetails_CashLineLink,strDetails_AccountNo,strDetails_FirstIssued,strDetails_ExpiryDate,strDetails_LastReplaced,strDetails_NoOfCardIssued,strDetails_IssuerID,strDetails_PINGenDate,strDetails_ActivationDate,strDetails_BOI,strDetails_LastTransactionDate,strDetails_LastUpdatedOn,strDetails_LastServiceType,strDetails_FPC,strCPFISLinkage)
	bverifyCardDetails_CardPinInfo = true

	If Not IsNull(strDetails_OverseasWdl) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_OverseasWdl.GetROProperty("innertext"))) = UCase(Trim(strDetails_OverseasWdl)) Then
				LogMessage "RSLT","Verification","Overseas Withdrawal matching with the expected value. Expected: "& strDetails_OverseasWdl &" , Actual: "& lblDetails_OverseasWdl.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "RSLT","Verification","Overseas Withdrawal not matching with the expected value. Expected: "& strDetails_OverseasWdl &" , Actual: "& lblDetails_OverseasWdl.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_CashLineLink) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_CashLineLink.GetROProperty("innertext"))) = UCase(Trim(strDetails_CashLineLink)) Then
				LogMessage "RSLT","Verification","CashLine Link matching with the expected value. Expected: "& strDetails_CashLineLink &" , Actual: "& lblDetails_CashLineLink.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "RSLT","Verification","CashLine Link not matching with the expected value. Expected: "& strDetails_CashLineLink &" , Actual: "& lblDetails_CashLineLink.GetROProperty("innertext") ,False
			End If
    End If
    
	If Not IsNull(strDetails_AccountNo) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_AccountNo.GetROProperty("innertext"))) = UCase(Trim(strDetails_AccountNo)) Then
				LogMessage "RSLT","Verification","Account Number matching with the expected value. Expected: "& strDetails_AccountNo &" , Actual: "& lblDetails_AccountNo.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Account No not matching with the expected value. Expected: "& strDetails_AccountNo &" , Actual: "& lblDetails_AccountNo.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_FirstIssued) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_FirstIssued.GetROProperty("innertext"))) = UCase(Trim(strDetails_FirstIssued)) Then
				LogMessage "RSLT","Verification","First issued matching with the expected value. Expected: "& strDetails_FirstIssued &" , Actual: "& lblDetails_FirstIssued.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","First issued not matching with the expected value. Expected: "& strDetails_FirstIssued &" , Actual: "& lblDetails_FirstIssued.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_ExpiryDate) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_ExpiryDate.GetROProperty("innertext"))) = UCase(Trim(strDetails_ExpiryDate)) Then
				LogMessage "RSLT","Verification","Expiry Date matching with the expected value. Expected: "& strDetails_ExpiryDate &" , Actual: "& lblDetails_ExpiryDate.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Expiry Date not matching with the expected value. Expected: "& strDetails_ExpiryDate &" , Actual: "& lblDetails_ExpiryDate.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_LastReplaced) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_LastReplaced.GetROProperty("innertext"))) = UCase(Trim(strDetails_LastReplaced)) Then
				LogMessage "RSLT","Verification","Last Replaced matching with the expected value. Expected: "& strDetails_LastReplaced &" , Actual: "& lblDetails_LastReplaced.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Last Replaced not matching with the expected value. Expected: "& strDetails_LastReplaced &" , Actual: "& lblDetails_LastReplaced.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_NoOfCardIssued) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_NoOfCardIssued.GetROProperty("innertext"))) = UCase(Trim(strDetails_NoOfCardIssued)) Then
				LogMessage "RSLT","Verification","Number of Cards Issued matching with the expected value. Expected: "& strDetails_NoOfCardIssued &" , Actual: "& lblDetails_NoOfCardIssued.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Number of Cards Issued not matching with the expected value. Expected: "& strDetails_NoOfCardIssued &" , Actual: "& lblDetails_NoOfCardIssued.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_IssuerID) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_IssuerID.GetROProperty("innertext"))) = UCase(Trim(strDetails_IssuerID)) Then
				LogMessage "RSLT","Verification","Issued ID matching with the expected value. Expected: "& strDetails_IssuerID &" , Actual: "& lblDetails_IssuerID.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Issued ID not matching with the expected value. Expected: "& strDetails_IssuerID &" , Actual: "& lblDetails_IssuerID.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_PINGenDate) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_PINGenDate.GetROProperty("innertext"))) = UCase(Trim(strDetails_PINGenDate)) Then
				LogMessage "RSLT","Verification","PIN Gen date matching with the expected value. Expected: "& strDetails_PINGenDate &" , Actual: "& lblDetails_PINGenDate.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","PIN Gen date not matching with the expected value. Expected: "& strDetails_PINGenDate &" , Actual: "& lblDetails_PINGenDate.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_ActivationDate) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_ActivationDate.GetROProperty("innertext"))) = UCase(Trim(strDetails_ActivationDate)) Then
				LogMessage "RSLT","Verification","Activation Date matching with the expected value. Expected: "& strDetails_ActivationDate &" , Actual: "& lblDetails_ActivationDate.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Activation Date not matching with the expected value. Expected: "& strDetails_ActivationDate &" , Actual: "& lblDetails_ActivationDate.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_BOI) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_BOI.GetROProperty("innertext"))) = UCase(Trim(strDetails_BOI)) Then
				LogMessage "RSLT","Verification","BOI matching with the expected value. Expected: "& strDetails_BOI &" , Actual: "& lblDetails_BOI.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","BOI not matching with the expected value. Expected: "& strDetails_BOI &" , Actual: "& lblDetails_BOI.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_LastTransactionDate) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_LastTransactionDate.GetROProperty("innertext"))) = UCase(Trim(strDetails_LastTransactionDate)) Then
				LogMessage "RSLT","Verification","Last Transaction Date matching with the expected value. Expected: "& strDetails_LastTransactionDate &" , Actual: "& lblDetails_LastTransactionDate.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Last Transaction Date not matching with the expected value. Expected: "& strDetails_LastTransactionDate &" , Actual: "& lblDetails_LastTransactionDate.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_LastUpdatedOn) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_LastUpdatedOn.GetROProperty("innertext"))) = UCase(Trim(strDetails_LastUpdatedOn)) Then
				LogMessage "RSLT","Verification","Last Upodated On matching with the expected value. Expected: "& strDetails_LastUpdatedOn &" , Actual: "& lblDetails_LastUpdatedOn.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Last Updated On not matching with the expected value. Expected: "& strDetails_LastUpdatedOn &" , Actual: "& lblDetails_LastUpdatedOn.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_LastServiceType) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_LastServiceType.GetROProperty("innertext"))) = UCase(Trim(strDetails_LastServiceType)) Then
				LogMessage "RSLT","Verification","Last Service Type matching with the expected value. Expected: "& strDetails_LastServiceType &" , Actual: "& lblDetails_LastServiceType.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","Last Service Type not matching with the expected value. Expected: "& strDetails_LastServiceType &" , Actual: "& lblDetails_LastServiceType.GetROProperty("innertext") ,False
			End If
    End If

    If Not IsNull(strDetails_FPC) Then
			If Ucase(Trim(bcCardAndInfoold.lblDetails_FPC.GetROProperty("innertext"))) = UCase(Trim(strDetails_FPC)) Then
				LogMessage "RSLT","Verification","FPC matching with the expected value. Expected: "& strDetails_FPC &" , Actual: "& lblDetails_FPC.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","FPC not matching with the expected value. Expected: "& strDetails_FPC &" , Actual: "& lblDetails_FPC.GetROProperty("innertext") ,False
			End If
    End If
	
	 If Not IsNull(strCPFISLinkage) Then
			If Ucase(Trim(bcCardAndInfoold.lblCPFISLinkage.GetROProperty("innertext"))) = UCase(Trim(strCPFISLinkage)) Then
				LogMessage "RSLT","Verification","CPFIS Linkage matching with the expected value. Expected: "& strCPFISLinkage &" , Actual: "& lblCPFISLinkage.GetROProperty("innertext") ,True
			Else
				bVerifyCardAndPINInfo = False
				LogMessage "WARN","Verification","CPFIS Linkage not matching with the expected value. Expected: "& strCPFISLinkage &" , Actual: "& lblCPFISLinkage.GetROProperty("innertext") ,False
			End If
    End If
End Function
