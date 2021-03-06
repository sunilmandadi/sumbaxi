'[Verify the row data for Insurance Info Table]
Public Function verifyrowdata_InsuranceInfo(arrRowDataList)
	bverifyrowdata_InsuranceInfo = true
	verifyrowdata_InsuranceInfo = verifyTableContentList(bcVerify_Insurance.tblVerifyInsuranceHeader,bcVerify_Insurance.tblVerifyInsuranceTable,arrRowDataList,"Insurance Info",false,null,null,null)
	verifyrowdata_InsuranceInfo = bverifyrowdata_InstallmentPlan
End Function

'[Click on Insurance Details link from insurance info table]
Public Function clickView_InsuranceInfo(lstRowData)
	bclickView_InsuranceInfo = true
	clickView_InsuranceInfo = selectTableLink(bcVerify_Insurance.tblVerifyInsuranceHeader,bcVerify_Insurance.tblVerifyInsuranceTable,lstRowData,"Insurance Info","Insurance Details",false,null,null,null)
	clickView_InsuranceInfo = bclickView_InsuranceInfo
End Function

'[Verify the popup for insurance details exist]
Public Function verifyInsuranceDetails(bExist)
	bDevPending=false
   bActualExist=bcVerify_Insurance.popupInsuranceDetails.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Insurance Details Exists As Expected" ,true
       verifyInsuranceDetails=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Insurance Details does not Exists As Expected" ,true
       verifyInsuranceDetails=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Insurance Details does not Exists As Expected" ,False
       verifyInsuranceDetails=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Insurance Details Still Exists" ,False
       verifyInsuranceDetails=False
   End If
End Function

'[Verify the fields of the Insurance details]
Public Function verifyfields_InsuranceDetails(lstInsuranceDetails)
	bverifyfields_InsuranceDetails = true
	intSize = Ubound(lstInsuranceDetails)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstInsuranceDetails(Iterator),":")(0))
		arrValue = trim(Split(lstInsuranceDetails(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "Status Change Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblStatusChangeDate(), arrValue, "Status Change Date")Then
				LogMessage "RSLT","Verification","Insurance details - Status Change Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Effective Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblEffectiveDate(), arrValue, "Effective Date")Then
				LogMessage "RSLT","Verification","Insurance details - Effective Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Reinstatement Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblReinstatementDate(), arrValue, "Reinstatement Date")Then
				LogMessage "RSLT","Verification","Insurance details - Reinstatement Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Cancellation Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblCancellationDate(), arrValue, "Cancellation Date")Then
				LogMessage "RSLT","Verification","Insurance details - Cancellation Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Last Billed Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblLastBilledDate(), arrValue, "Last Billed Date")Then
				LogMessage "RSLT","Verification","Insurance details - Last Billed Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Premium Rate"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblPremiumRate(), arrValue, "Premium Rate")Then
				LogMessage "RSLT","Verification","Insurance details - Premium Rate:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		Case "Last"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblLastPremiumBilled(), arrValue, "Last")Then
				LogMessage "RSLT","Verification","Insurance details - Last:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		Case "Cycle To Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblCycleToDate(), arrValue, "Cycle To Date")Then
				LogMessage "RSLT","Verification","Insurance details - Cycle To Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		Case "Month To Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblMonthToDate(), arrValue, "Month To Date")Then
				LogMessage "RSLT","Verification","Insurance details - Month To Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Year To Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblYearToDate(), arrValue, "Year To Date")Then
				LogMessage "RSLT","Verification","Insurance details - Year To Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		
		Case "Life To Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcVerify_Insurance.lblLifeToDate(), arrValue, "Life To Date")Then
				LogMessage "RSLT","Verification","Insurance details - Life To Date:"&arrValue&" is not displayed as expected",false
			bverifyfields_InsuranceDetails=false
			End If
		End If
		End Select
	Next 
		bcVerify_Insurance.btnOK.click
		verifyfields_InsuranceDetails =bverifyfields_InsuranceDetails
End Function
