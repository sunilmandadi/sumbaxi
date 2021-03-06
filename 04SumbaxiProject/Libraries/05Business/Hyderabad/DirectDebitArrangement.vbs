'[Verify the fields of the Direct Debit Arrangement]
Public Function verifyfields_DirectDebit(lstDirectDebit)
	bverifyfields_DirectDebit = true
	intSize = Ubound(lstDirectDebit)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstDirectDebit(Iterator),":")(0))
		arrValue = trim(Split(lstDirectDebit(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "Bank Code"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblBankCode(), arrValue, "Bank Code")Then
				LogMessage "RSLT","Verification","Direct Debit details - Bank Code:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
			
		Case "Bank Account No"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblBankAccountNo(), arrValue, "Bank Account No")Then
				LogMessage "RSLT","Verification","Direct Debit details - Bank Account No:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
		
		Case "Payment Indicator"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblPaymentIndicator(), arrValue, "Payment Indicator")Then
				LogMessage "RSLT","Verification","Direct Debit details - Payment Indicator:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
		
		Case "Nominated Amount"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblNominatedAmount(), arrValue, "Nominated Amount")Then
				LogMessage "RSLT","Verification","Direct Debit details - Nominated Amount:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If

		Case "Nominated Percentage"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblNominatedPercentage(), arrValue, "Nominated Percentage")Then
				LogMessage "RSLT","Verification","Direct Debit details - Nominated Percentage:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
			
		Case "Start Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblStartDate(), arrValue, "Start Date")Then
				LogMessage "RSLT","Verification","Direct Debit details - Start Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
			
		Case "Expiry Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblExpiryDate(), arrValue, "Expiry Date")Then
				LogMessage "RSLT","Verification","Direct Debit details - Expiry Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
			
		Case "Status"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcDirectDebit_Arrangement.lblStatus(), arrValue, "Status")Then
				LogMessage "RSLT","Verification","Direct Debit details - Status:"&arrValue&" is not displayed as expected",false
				bverifyfields_DirectDebit=false
				End If
			End If
		End Select
	Next 
	verifyfields_DirectDebit = bverifyfields_DirectDebit
End Function

'[verify and Click on the link of Account Number]
Public Function clickaccno_DirectDebit(bDisabled)
	bclickaccno_DirectDebit = true
	  If bDisabled Then
	      LogMessage "RSLT","Verification","Link is disabled.",True
	  else
	      LogMessage "RSLT","Verification","Link is enabled.",True
		  bcDirectDebit_Arrangement.lblBankAccountNo.click
	WaitForICallLoading
	If not BalancesAndLimits.lblAccountBalance_AvailableBalance.Exist Then
		bclickaccno_DirectDebit =false
	   End if
	End If		
	clickaccno_DirectDebit = bclickaccno_DirectDebit
End Function
