'[Verify the row data for unbilled transactions displayed as in transaction history page]
Public Function verifyrowdata_UnbilledTransactions(arrRowDataList)
	bverifyrowdata_UnbilledTransactions = true
	bverifyrowdata_UnbilledTransactions = verifyTableContentList(TransactionHistory.tblTransactionsHeader_UB,TransactionHistory.tblTransactions_UB,arrRowDataList,"Unbilled Transaction",true,TransactionHistory.lnkNext_UB,TransactionHistory.lnkNext1_UB,TransactionHistory.lnkPrevious_UB)
	verifyrowdata_UnbilledTransactions = bverifyrowdata_UnbilledTransactions
End Function

'[Verify and select the radiobutton for transaction history page]
Public Function selectRadioButton_TransactionHist(strTrnType)
	bselectRadioButton_TransactionHist=true
	intRadio_Rewards=Instr(TransactionHistory.rbtnTranType.GetROproperty("class"),"disabled-area")
	If intRadio_Rewards = 0 Then
		bselectRadioButton_rewards=SelectRadioButtonGrp(strTrnType,TransactionHistory.rbtnTranType, Array("Unbilled","Pending","Declined","Life To-Date","Earmark"))
	Else
		LogMessage "RSLT","Verifiation","Radio button is disabled by default.",True
	End If
	If Err.Number<>0 Then
       bselectRadioButton_TransactionHist=false
          LogMessage "WARN","Verification","Failed to select radiobutton or radiobutton is disabled" ,True
       Exit Function
   End If
   selectRadioButton_TransactionHist=bselectRadioButton_TransactionHist
End Function

'[Verify and click the transaction description link from unbilled transaction table]
Public Function clickTransDesc_UnbilledTransactions(lstRowData)
	bclickTransDesc_UnbilledTransactions = true
	bclickTransDesc_UnbilledTransactions = selectTableLink(TransactionHistory.tblTransactionsHeader_UB,TransactionHistory.tblTransactions_UB,lstRowData,"Unbilled transactions","Transaction Description",true,TransactionHistory.lnkNext_UB,TransactionHistory.lnkNext1_UB,TransactionHistory.lnkPrevious_UB)
	clickTransDesc_UnbilledTransactions = bclickTransDesc_UnbilledTransactions
End Function

'[Verify the transaction description popup and the fields displayed]
Public Function verifypopup_TransactionDesc(lstadditionalInfo)
	bverifypopup_TransactionDesc = False
	If TransactionHistory.popupAdditionalInfo.exist Then
	intSize = Ubound(lstadditionalInfo)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstadditionalInfo(Iterator),":")(0))
		arrValue = trim(Split(lstadditionalInfo(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "Merchant Category"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblMerchantCategory(), arrValue, "Merchant Category")Then
				LogMessage "RSLT","Verification","Additional Info - Merchant Category:"&arrValue&" is not displayed as expected",false
				arrValue1 = False
			End If
		End If
		
		Case "Authorization Code"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblAuthorizationCode(), arrValue, "Authorization Code")Then
				LogMessage "RSLT","Verification","Additional Info - Authorization Code:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Sequence No."
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblSequenceNo(), arrValue, "Sequence No.")Then
				LogMessage "RSLT","Verification","Additional Info - Sequence No.:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Points Earned"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblPointsEarned(), arrValue, "Points Earned")Then
				LogMessage "RSLT","Verification","Additional Info - Points Earned:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Cheque No."
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblCheckNo(), arrValue, "Cheque No.")Then
				LogMessage "RSLT","Verification","Additional Info - Cheque No.:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Merchant ID"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblMerchantID(), arrValue, "Merchant ID")Then
				LogMessage "RSLT","Verification","Additional Info - Merchant ID:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Merchant Name"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblMerchantName(), arrValue, "Merchant Name")Then
				LogMessage "RSLT","Verification","Additional Info - Merchant Name:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Merchant Organization"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblMerchantOrg(), arrValue, "Merchant Organization")Then
				LogMessage "RSLT","Verification","Additional Info - Merchant Organization:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Merchant Category Code"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblMerchantCatCode(), arrValue, "Merchant Category Code")Then
				LogMessage "RSLT","Verification","Additional Info - Merchant Category Code:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Card Acceptor Name"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblCardAcceptorName(), arrValue, "Card Acceptor Name")Then
				LogMessage "RSLT","Verification","Additional Info - Card Acceptor Name:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Card Acceptor City"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblCardAcceptorCity(), arrValue, "Card Acceptor City")Then
				LogMessage "RSLT","Verification","Additional Info - Card Acceptor City:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Card Acceptor Country"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblCardAcceptorCountry(), arrValue, "Card Acceptor Country")Then
				LogMessage "RSLT","Verification","Additional Info - Card Acceptor Country:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Acquiring Country"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblAquiringCountry(), arrValue, "Acquiring Country")Then
				LogMessage "RSLT","Verification","Additional Info - Acquiring Country:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "POS Entry Mode"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblPOSEntryMode(), arrValue, "POS Entry Mode")Then
				LogMessage "RSLT","Verification","Additional Info - POS Entry Mode:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "eCommerce Transaction Type"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lbleCommerceTranType(), arrValue, "eCommerce Transaction Type")Then
				LogMessage "RSLT","Verification","Additional Info - eCommerce Transaction Type:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Currency Code"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblCurrencyCode(), arrValue, "Currency Code")Then
				LogMessage "RSLT","Verification","Additional Info - Currency Code:"&arrValue&" is not displayed as expected",false
			End If
		End If
		
		Case "Additional Info"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (TransactionHistory.lblAdditionalIfno(), arrValue, "Additional Info")Then
				LogMessage "RSLT","Verification","Additional Info - Additional Info:"&arrValue&" is not displayed as expected",false
			End If
		End If
	   End Select
   Next
   	TransactionHistory.btnOK.click
	'verifypopup_TransactionDesc = bverifypopup_TransactionDesc
 else
    LogMessage "RSLT","Verification","The additional info link is disabled and the popup not displayed .",true
	verifypopup_TransactionDesc = bverifypopup_TransactionDesc
	End If	
End Function

'[Verify the row data for pending transactions displayed as in transaction history page]
Public Function verifyrowdata_PendingTransactions(arrRowDataList)
WaitForICallLoading
bPendingTransactions = verifyTableContentList(TransactionHistory.tblTransactionsHeader_Pend,TransactionHistory.tblTransactions_Pend,arrRowDataList,"Pending Transaction",true,TransactionHistory.lnkNext_Pend,TransactionHistory.lnkNext1_Pend,TransactionHistory.lnkPrevious_Pend)
verifyrowdata_PendingTransactions = bPendingTransactions
End Function

'[Verify and click the transaction type link from pending transaction table]
Public Function clickTransType_PendingTransactions(lstRowData)
WaitForICallLoading
bPendingTransactions = selectTableLink(TransactionHistory.tblTransactionsHeader_Pend,TransactionHistory.tblTransactions_Pend,lstRowData,"Pending Transaction","Transaction Type",true,TransactionHistory.lnkNext_Pend,TransactionHistory.lnkNext1_Pend,TransactionHistory.lnkPrevious_Pend)
clickTransType_PendingTransactions = bPendingTransactions
End Function

'[Verify the row data for Declined transactions displayed as in transaction history page]
Public Function verifyrowdata_DeclinedTransactions(arrRowDataList)
WaitForICallLoading
bDeclinedTransactions = verifyTableContentList(TransactionHistory.tblTransactionsHeader_Decl,TransactionHistory.tblTransactions_Decl,arrRowDataList,"Declined Transaction",true,TransactionHistory.lnkNext_Decl,TransactionHistory.lnkNext1_Decl,TransactionHistory.lnkPrevious_Decl)
verifyrowdata_DeclinedTransactions = bDeclinedTransactions
End Function

'[Verify and click the transaction type link from declined transaction table]
Public Function clickTransType_DeclinedTransactions(lstRowData)
WaitForICallLoading
bDeclinedTransactions = selectTableLink(TransactionHistory.tblTransactionsHeader_Decl,TransactionHistory.tblTransactions_Decl,lstRowData,"Declined Transaction","Transaction Type",true,TransactionHistory.lnkNext_Decl,TransactionHistory.lnkNext1_Decl,TransactionHistory.lnkPrevious_Decl)
clickTransType_DeclinedTransactions = bDeclinedTransactions
End Function

'[Verify the row data for Life to Date transactions displayed as in transaction history page]
Public Function verifyrowdata_LTDTransactions(strtablerow)
WaitForICallLoading
bLTDTransactions = True
'the below function of checking the column name doesnt work as the Column names are duplicated in the table. 
'bLTDTransactions = verifyTableContentList(TransactionHistory.tblTransactionHistoryHeader_LTD,TransactionHistory.tblTransactionHistory_LTD,arrRowDataList,"LTD Transaction",true,TransactionHistory.lnkNext_LTD,TransactionHistory.lnkNext1_LTD,TransactionHistory.lnkPrevious_LTD)
If Not IsNull(strtablerow) Then
	If Not VerifyInnerText (TransactionHistory.tblRowdata_LTD(), strtablerow, "Life to Mark Transaction table row") Then
	bLTDTransactions=False
	End If
End If 
verifyrowdata_LTDTransactions = bLTDTransactions
End Function

'[Verify and click the transaction desc link from LTD transaction table]
Public Function clickTransType_LTDTransactions(lstRowData)
WaitForICallLoading
bLTDTransactions = selectTableLink(TransactionHistory.tblTransactionHistoryHeader_LTD,TransactionHistory.tblTransactionHistory_LTD,lstRowData,"LTD Transaction","Transaction Description",true,TransactionHistory.lnkNext_LTD,TransactionHistory.lnkNext1_LTD,TransactionHistory.lnkPrevious_LTD)
clickTransType_LTDTransactions = bLTDTransactions
End Function

'[Verify the row data for Earmark transactions displayed as in transaction history page]
Public Function verifyrowdata_EarmarkTransactions(arrRowDataList)
WaitForICallLoading
bEarmarkTransactions = verifyTableContentList(TransactionHistory.tblTransactionHistoryHeader_Earmark,TransactionHistory.tblTransactionHistoryContent_Earmark,arrRowDataList,"Earmark Transaction",false,Null,Null,Null)
verifyrowdata_EarmarkTransactions = bEarmarkTransactions
End Function
