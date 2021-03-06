'[Click on D2Pay link in Customer Overview Page]
Public Function clickD2Pay()	
	bcCustomerOverview.lnkD2Pay.Click
	If Err.Number<>0 Then
       clickD2Pay=false
       LogMessage "WARN","Verification","Failed to Click Button : D2Pay" ,false
       Exit Function
   End If
    clickD2Pay=true
	WaitForICallLoading
End Function

'[Verify Pink Panel in D2Pay enquiry screen]
Public Function verifyD2PayPinkPanel(strName,strCIN,strSegment)
	bverifyD2PayPinkPanel=true
	If Not IsNull (strName) Then
		If Not verifyInnerText(D2Pay.lblName(),strName, "Name") Then
			bverifyD2PayPinkPanel=false
		End If
	End If
	If Not IsNull (strCIN) Then
		If Not verifyInnerText(D2Pay.lblCIN(),strCIN, "CIN") Then
			bverifyD2PayPinkPanel=false
		End If
	End If
	If Not IsNull (strSegment) Then
		If Not verifyInnerText(D2Pay.lblSegment(),strSegment, "Segment") Then
			bverifyD2PayPinkPanel=false
		End If
	End If
	verifyD2PayPinkPanel=bverifyD2PayPinkPanel
End Function

'[Verify details displayed in D2Pay page for user with IB Access]
Public Function Verifydisplayinfo_D2Pay(strCustCIN, strDebitAccNo,strAccType, StrStatus, StrReason,StrEcommLimit,StrUsedLimit,StrRemLimit, StrRetryCount)
	bVerifydisplayinfo_D2Pay = True
	Call D2Pay_GPEP_GPSP(strCustCIN)
	'Values retrieved from the GPSP System 
	strExpDebitAccountNo =Environment.Value("strExpDebitAccountNo")
	strExpAccountType=Environment.Value("strExpAccountType")
	strExpIBstatus=Environment.Value("strExpIBstatus")
	strExpReason=Environment.Value("strExpReason")
	strExpEComLimit=Environment.Value("strExpEComLimit")
	strExpUsedLimit=Environment.Value("strExpUsedLimit")
	strExpRemLimit=Environment.Value("strExpRemLimit")
	strExpRetryCount=Environment.Value("strExpRetryCount")	
	
	'Values retrived from the Iserve screen
	strActualDebitAccountNo = D2Pay.lblDebitAccNo.GetROProperty("innertext")
	strActualAccountType = D2Pay.lblAccountType.GetROProperty("innertext")
	strActualIBStatus = D2Pay.lblStatus.GetROProperty("innertext")
	strActualReason = D2Pay.lblReason.GetROProperty("innertext")
	strActualEComLimit = Replace(D2Pay.lblEcommerceLimit.GetROProperty("innertext"),",","")
	strActualEComLimit = Replace(strActualEComLimit,"SGD","")
	strActualUsedLimit = Replace(D2Pay.lblUsedLimit.GetROProperty("innertext"),",","")
	strActualUsedLimit = Replace(strActualUsedLimit,"SGD","")
	strActualRemLimit =  Replace(D2Pay.lblRemainingLimit.GetROProperty("innertext"),",","")
	strActualRemLimit = Replace(strActualRemLimit,"SGD","")
	strActualRetryCount = D2Pay.lblRetryCount.GetROProperty("innertext")
	
 	If Not IsNull(strDebitAccNo) Then
		If strDebitAccNo = "RUNTIME" Then		
			If UCase(Trim(strExpDebitAccountNo)) = UCase(Trim(strActualDebitAccountNo)) Then
				LogMessage "RSLT","Verification","DebitAccount Number matched with the Expected value. Expected:" + strExpDebitAccountNo &",Actual:" + strActualDebitAccountNo,True
			Else
				LogMessage "WARN","Verification","DebitAccount Number not matched with the Expected value. Expected:" + strExpDebitAccountNo &",Actual:" + strActualDebitAccountNo,False
				bVerifydisplayinfo_D2Pay = False
			End If	
		End If 
	End If	
 
  If Not IsNull(strAccType) Then
	If strAccType = "RUNTIME" Then		
		If UCase(Trim(strExpAccountType)) = "010" AND UCase(Trim(strActualAccountType)) = "0010 - SA PLUS" Then
			LogMessage "RSLT","Verification","Account Type is matched with the Expected value. Expected:" + strExpAccountType &",Actual:" + strActualAccountType,True
			
		ElseIf UCase(Trim(strExpAccountType)) = "011" AND UCase(Trim(strActualAccountType)) = "0011 - POSB SA" Then
			LogMessage "RSLT","Verification","Account Type is matched with the Expected value. Expected:" + strExpAccountType &",Actual:" + strActualAccountType,True
			
		ElseIf UCase(Trim(strExpAccountType)) = "020" AND UCase(Trim(strActualAccountType)) = "0020 - CA" Then
			LogMessage "RSLT","Verification","Account Type is matched with the Expected value. Expected:" + strExpAccountType &",Actual:" + strActualAccountType,True
			
		ElseIf UCase(Trim(strExpAccountType)) = "021" AND UCase(Trim(strActualAccountType)) = "0021 - AUTOSAVE ACCOUNT" Then
			LogMessage "RSLT","Verification","Account Type is matched with the Expected value. Expected:" + strExpAccountType &",Actual:" + strActualAccountType,True
			
		ElseIf UCase(Trim(strExpAccountType)) = "025" AND UCase(Trim(strActualAccountType)) = "0025 - POSB CA" Then
			LogMessage "RSLT","Verification","Account Type is matched with the Expected value. Expected:" + strExpAccountType &",Actual:" + strActualAccountType,True
		Else
			LogMessage "WARN","Verification","Account Type is not matched with the Expected value. Expected:" + strExpAccountType &",Actual:" + strActualAccountType,False
			bVerifyIBStatus_D2Pay = False
		End If
	End If 
 End If 
 
   If Not IsNull(StrStatus) Then
	If StrStatus = "RUNTIME" Then	
		If Trim(strExpIBstatus) = "0" AND Ucase(Trim(strActualIBStatus)) = "0 - ACTIVE" Then
			LogMessage "RSLT","Verification","Status is matched with the Expected value. Expected:" + strExpIBstatus &",Actual:" + strActualIBStatus,True
		ElseIf Trim(strExpIBstatus) = "1" AND Ucase(Trim(strActualIBStatus)) = "1 - DENY" Then
			LogMessage "RSLT","Verification","Status is matched with the Expected value. Expected:" + strExpIBstatus &",Actual:" + strActualIBStatus,True
		ElseIf Trim(strExpIBstatus) = "2" AND Ucase(Trim(strActualIBStatus)) = "2 - HOT TAG" Then
			LogMessage "RSLT","Verification","Status is matched with the Expected value. Expected:" + strExpIBstatus &",Actual:" + strActualIBStatus,True
		ElseIf Trim(strExpIBstatus) = "8" AND Ucase(Trim(strActualIBStatus)) = "8 - CLOSED BY CUST" Then
			LogMessage "RSLT","Verification","Status is matched with the Expected value. Expected:" + strExpIBstatus &",Actual:" + strActualIBStatus,True
		ElseIf Trim(strExpIBstatus) = "9" AND Ucase(Trim(ActualIBStatus)) = "9 - CLOSED BY BANK" Then
			LogMessage "RSLT","Verification","Status is matched with the Expected value. Expected:" + strExpIBstatus &",Actual:" + strActualIBStatus,True
		Else
			LogMessage "WARN","Verification","Status is not matched with the Expected value. Expected:" + strExpIBstatus &",Actual:" + strActualIBStatus,False
			bVerifyIBStatus_D2Pay = False
		End If
	End If 
 End If 	

	If Not IsNull(StrReason) Then
		If StrReason = "RUNTIME" Then		
			If UCase(Trim(strExpReason)) = UCase(Trim(strActualReason)) Then
				LogMessage "RSLT","Verification","Reason is matched with the Expected value. Expected:" + strExpReason &",Actual:" + strActualReason,True
			Else
				LogMessage "WARN","Verification","Reason is not matched with the Expected value. Expected:" + strExpReason &",Actual:" + strActualReason,False
				bVerifyReason_D2Pay = False
			End If
		End If
	End If
	
 	If Not IsNull(StrEcommLimit) Then
		If StrEcommLimit = "RUNTIME" Then	
			If Trim(strExpEComLimit) = Trim(strActualEComLimit) Then
				LogMessage "RSLT","Verification","E-commerce Limit is matched with the Expected value. Expected:" + strExpEComLimit &",Actual:" + strActualEComLimit,True
			Else
				LogMessage "WARN","Verification","E-commerce Limit is not matched with the Expected value. Expected:" + strExpEComLimit &",Actual:" + strActualEComLimit,False
				bVerifylimits_D2Pay = False
			End If
		End If 
	End If	
			
 	If Not IsNull(StrUsedLimit) Then
		If StrUsedLimit = "RUNTIME" Then	
			If Trim(strExpUsedLimit) = Trim(strActualUsedLimit) Then
				LogMessage "RSLT","Verification","Used limit is matched with the Expected value. Expected:" + strExpUsedLimit &",Actual:" + strActualUsedLimit,True
			Else
				LogMessage "WARN","Verification","Used limit is not matched with the Expected value. Expected:" + strExpUsedLimit &",Actual:" + strActualUsedLimit,False
				bVerifylimits_D2Pay = False
			End If	
		End If	
	End If

 	If Not IsNull(StrRemLimit) Then
		If StrRemLimit = "RUNTIME" Then	
			If Trim(strExpremLimit) = Trim(strActualRemLimit) Then
				LogMessage "RSLT","Verification","Remaining limit is matched with the Expected value. Expected:" + strExpremLimit &",Actual:" + strActualRemLimit,True
			Else
				LogMessage "WARN","Verification","Remaining limit is not matched with the Expected value. Expected:" + strExpremLimit &",Actual:" + strActualRemLimit,False
				bVerifylimits_D2Pay = False
			End If
		End If	
	End If
	
	If Not IsNull(StrRetryCount) Then
		If StrRetryCount = "RUNTIME" Then
			If Trim(strExpRetryCount) = Trim(strActualRetryCount) Then
				LogMessage "RSLT","Verification","RetryCount is matched with the Expected value. Expected:" + strExpRetryCount &",Actual:" + strActualRetryCount,True
			Else
				LogMessage "WARN","Verification","RetryCount is not matched with the Expected value. Expected:" + strExpRetryCount &",Actual:" + strActualRetryCount,False
				bVerifyRetryCount_D2Pay = False
			End If
		End If	
	End If
End Function

'[LISA Verify details displayed in D2Pay page for user with IB Access]
Public Function VerifyProfileinfo_D2Pay(strDebitAccNo,strAccType,strStatus,strReason,strEcommLimit,strUsedLimit,strRemLimit,strRetryCount)
	
	' Get the Iserve fields related to terminal transaction details
	strIserveDebitAccountNo = D2Pay.lblDebitAccNo.GetROProperty("innertext")
	strIserveAccountType = D2Pay.lblAccountType.GetROProperty("innertext")
	strIserveIBStatus = D2Pay.lblStatus.GetROProperty("innertext")
	strIserveReason = D2Pay.lblReason.GetROProperty("innertext")
	strIserveEcommLimit = D2Pay.lblEcommerceLimit.GetROProperty("innertext")
	strIserveUsedLimit = D2Pay.lblUsedLimit.GetROProperty("innertext")
	strIserveRemLimit = D2Pay.lblRemainingLimit.GetROProperty("innertext")
	strIserveRetryCount = D2Pay.lblRetryCount.GetROProperty("innertext")
	
   bDevPending=false
   bVerifyProfileinfo_D2Pay=true		
			
	If strDebitAccNo = strIserveDebitAccountNo Then
		LogMessage "RSLT","Verification","The Iserve Debit Acc No. value is as expected: "&strDebitAccNo&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Debit Acc No. value is not as expected: "&strDebitAccNo&"",False
	End If
	
	If strAccType = strIserveAccountType Then
		LogMessage "RSLT","Verification","The Iserve Acc Type value is as expected: "&strAccType&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Acc Type value is not as expected: "&strAccType&"",False
	End If
	
	If strStatus = strIserveIBStatus Then
		LogMessage "RSLT","Verification","The Iserve Status value is as expected: "&strStatus&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Status value is not as expected: "&strStatus&"",False
	End If
	
	If strReason = strIserveReason Then
		LogMessage "RSLT","Verification","The Iserve Reason value is as expected: "&strReason&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Reason value is not as expected: "&strReason&"",False
	End If
	
	If strEcommLimit = strIserveEcommLimit Then
		LogMessage "RSLT","Verification","The Iserve Ecomm Limit value is as expected: "&strEcommLimit&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Ecomm Limit value is not as expected: "&strEcommLimit&"",False
	End If
	
	If strUsedLimit = strIserveUsedLimit Then
		LogMessage "RSLT","Verification","The Iserve Used Limit value is as expected: "&strUsedLimit&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Used Limit value is not as expected: "&strUsedLimit&"",False
	End If
	
	If strRemLimit = strIserveRemLimit Then
		LogMessage "RSLT","Verification","The Iserve Remaining Limit value is as expected: "&strRemLimit&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Remaining Limit value is not as expected: "&strRemLimit&"",False
	End If
	
	If strRetryCount = strIserveRetryCount Then
		LogMessage "RSLT","Verification","The Iserve Retry Count value is as expected: "&strRetryCount&"",True
		Else 
		LogMessage "WARN","Verification","The Iserve Retry Count value is not as expected: "&strRetryCount&"",False
	End If
	VerifyProfileinfo_D2Pay =bVerifyProfileinfo_D2Pay
End Function

'[Verify the click for field Account no]
Public Function verifyAccNo_D2Pay()
	bverifyAccNo_D2Pay=true
	D2Pay.lblDebitAccNo.click
	WaitForICallLoading
	'If not BalancesAndLimits.popupEarmarkDetails.Exist Then
	'	bverifyEarmarkPopup=false
	'End If
	verifyAccNo_D2Pay=bverifyAccNo_D2Pay
End Function

