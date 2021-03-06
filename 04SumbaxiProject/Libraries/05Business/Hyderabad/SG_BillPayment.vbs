'[Click on Bill Payment link under Banking Facilities in Overview Page]
Public Function ClickBillPayment_BP()		
	bcCustomerOverview.lnkBillPayment.Click
	ClickBillPayment_BP=true
	If Err.Number<>0 Then
       ClickBillPayment_BP=false
       LogMessage "WARN","Verification","Failed to Click Button : Bill Payment" ,false
       Exit Function
   	End If
	WaitForICallLoading	
End Function

'[Verify default Account Number in Bill Payment Page displayed as]
Public Function VerifyAccountNumber_BP(strAccNo)
   bVerifyAccountNumber=true
   If Not IsNull(strAccNo) Then
       If Not verifyComboSelectItem (BillPayment.lstAccountNo(),strAccNo, "Account/Card No.")Then
    	  LogMessage "WARN","Verification","Expected Default Account type:"&strAccNo&" not displayed in the Account/Card No field" ,false
          bVerifyAccountNumber=false
       End If
   End If
   VerifyAccountNumber_BP=bVerifyAccountNumber
End Function

'[Verify list of values displayed in Account Number dropdown]
Public Function VerifylistAccountNo_BP(lstAccNo) 
	bVerifylistAccountNo = True 
	If Not IsNull(lstAccNo) Then
       If Not verifyComboboxItems(BillPayment.lstAccountNo(),lstAccNo, "Account/Card No.")Then
       	   LogMessage "WARN","Verification","List of Account No displayed in the combox box is not as expected" ,false
           bVerifylistAccountNo=false
       End If
   End If
   VerifylistAccountNo_BP = bVerifylistAccountNo
End Function

'[Select Account Number combox box as]
Public Function SelectAccountNo(strAccNo)
	bSelectAccountNo=true
	If Not IsNull(strAccNo) Then
       If Not (selectItem_Combobox (BillPayment.lstAccountNo(), strAccNo))Then
           LogMessage "WARN","Verification","Failed to select :"&strAccNo&" From Account/Card No dropdown list" ,false
           bSelectAccountNo=false
       End If
   End If
   WaitForICallLoading
   SelectAccountNo=bSelectAccountNo
End Function

'[Verify Bill Payment table details displayed based on the selected Account type from the dropdown]
Public Function verifyBillPaymentdetails_BP(lstlstBillPaymentDetails)
   bverifyPaymentdetails=verifyTableContentList(BillPayment.tblBillPaymentDetailsHeader, BillPayment.tblBillPaymentDetailsContent,lstlstBillPaymentDetails,"Bill Payment Details",false,NULL,NULL,NULL)
   verifyBillPaymentdetails_BP = bverifyPaymentdetails
End Function

'[Click on View hyperlink]
Public Function verifyViewHyperlink_BP(lstBillPaymentsdetails)
	bverifyViewHyperlink = selectTableLink(BillPayment.tblBillPaymentDetailsHeader, BillPayment.tblBillPaymentDetailsContent, lstBillPaymentsdetails, "Bill Payment Details", "More Info.", false, null, null, null)
	verifyViewHyperlink_BP = bverifyViewHyperlink
End Function

'[Verify More details displayed for the selected Account No in the table displayed]
Public Function verifyMoreDetails_BP(lstMoreDetails)
	bverifyMoreDetails = true
	intSize = Ubound(lstMoreDetails)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstMoreDetails(Iterator),":")(0))
		arrValue = trim(Split(lstMoreDetails(Iterator),":")(1))
		Select Case (arrLabel)
			Case "Account Bill Limit"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (BillPayment.lblAccBillLimit(), arrValue, "Account Bill Limit")Then
			       LogMessage "WARN","Verification","More details - Account Bill Limit:"&arrValue&" is not displayed as expected",false
			       bverifyMoreDetails=false
			       End If
			    End If
			Case "Accumulated Bill Amount"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (BillPayment.lblAccumulatedBillAmt(), arrValue, "Accumulated Bill Amount")Then
			       	   LogMessage "WARN","Verification","More details - Accumulated Bill Amount:"&arrValue&" is not displayed as expected",false
			           bverifyMoreDetails=false
			       End If
			    End If
		End Select
	Next
	BillPayment.btnOK.click
	verifyMoreDetails_BP = bverifyMoreDetails
End Function

'[Click on Account No hyperlink]
Public Function verifyAccNoHyperlink_BP(lstBillPaymentsdetails)
	bverifyAccNoHyperlink = selectTableLink(BillPayment.tblBillPaymentDetailsHeader, BillPayment.tblBillPaymentDetailsContent, lstBillPaymentsdetails, "Bill Payment Details", "Account/Card No.", false, null, null, null)
	WaitForICallLoading
	verifyAccNoHyperlink_BP = bverifyAccNoHyperlink
End Function

'[Verify Info Warn message on Bill Payment Page]
Public Function VerifyInfowarn_BP(strInfoWarnMsg)
	bverifyInfoWarn = True
	strInfoWarnMsg= Replace(strInfoWarnMsg,"@","=")
	strInfoWarnMsg= Replace(strInfoWarnMsg,"-",";")
	bverifyInfoWarn = verifyInfoWarn_popup(strInfoWarnMsg)
	VerifyInfowarn_BP = bverifyInfoWarn
End Function


