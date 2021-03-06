'[Verify the row data for installment plan details]
Public Function verifyrowdata_InstallmentPlan(arrRowDataList)
	bverifyrowdata_InstallmentPlan = true
	verifyrowdata_InstallmentPlan = verifyTableContentList(bcInstallmentPlan.tblInstallmentPlanHeader,bcInstallmentPlan.tblInstallmentPlanContent,arrRowDataList,"Installment Plan",false,null,null,null)
	verifyrowdata_InstallmentPlan = bverifyrowdata_InstallmentPlan
End Function

'[Click on the Merchant Detail from the installment plan details table]
Public Function clickmerchant_InstallmentPlan(lstRowData)
	bclickmerchant_InstallmentPlan = true
	clickmerchant_InstallmentPlan = selectTableLink(bcInstallmentPlan.tblInstallmentPlanHeader,bcInstallmentPlan.tblInstallmentPlanContent,lstRowData,"Installment Plan","Merchant Name",false,null,null,null)
	clickmerchant_InstallmentPlan = bclickmerchant_InstallmentPlan
End Function

'[Click on the sequence no. from the installment plan details table]
Public Function clicksequence_InstallmentPlan(lstRowData)
	bclicksequence_InstallmentPlan = true
	clicksequence_InstallmentPlan = selectTableLink(bcInstallmentPlan.tblInstallmentPlanHeader,bcInstallmentPlan.tblInstallmentPlanContent,lstRowData,"Installment Plan","Seq No.",false,null,null,null)
	clicksequence_InstallmentPlan = bclicksequence_InstallmentPlan
End Function

'[Verify the popup for Merchant Information exist]
Public Function verifyMerchantInfoExist(bExist)
	bDevPending=false
   bActualExist=bcInstallmentPlan.lblMerchantName_MerchantInfo.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Merchant Information Exists As Expected" ,true
       verifyMerchantInfoExist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Merchant Information does not Exists As Expected" ,true
       verifyMerchantInfoExist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Merchant Information does not Exists As Expected" ,False
       verifyMerchantInfoExist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Merchant Information Still Exists" ,False
       verifyMerchantInfoExist=False
   End If
End Function

'[Verify the fields for the merchant Information Installment Plan]
Public Function verifyfields_MerchantInfo(strMerchantName,strMerchantId,strCity,strAddr1,strAddr2,strAddr3,strAddr4,strPostal,strPhoneNo)

		'Getting the values of the fields from front end I.Serve
		strIserveMerchantName = bcInstallmentPlan.lblMerchantName_MerchantName.GetROProperty("innertext")
		strIserveMerchantId = bcInstallmentPlan.lblMerchantName_MerchantId.GetROProperty("innertext")
		strIserveCity = bcInstallmentPlan.lblMerchantName_City.GetROProperty("innertext")
		strIserveAddr1 = bcInstallmentPlan.lblMerchantName_Addr1.GetROProperty("innertext")
		strIserveAddr2 = bcInstallmentPlan.lblMerchantName_Addr2.GetROProperty("innertext")
		strIserveAddr3 = bcInstallmentPlan.lblMerchantName_Addr3.GetROProperty("innertext")
		strIserveAddr4 = bcInstallmentPlan.lblMerchantName_Addr4.GetROProperty("innertext")
		strIservePostal = bcInstallmentPlan.lblMerchantName_PostalCode.GetROProperty("innertext")
		strIservePhoneNo = bcInstallmentPlan.lblMerchantName_PhoneNum.GetROProperty("innertext")
		
		bverifyfields_MerchantInfo = true
		
		If strMerchantName = strIserveMerchantName Then
			LogMessage "RSLT","Verification","The Iserve Merchant name is as expected: "&strMerchantName&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Merchant name is not as expected: "&strMerchantName&"",False
		End If
		
		If strMerchantId = strIserveMerchantId Then
			LogMessage "RSLT","Verification","The Iserve Merchant Id is as expected: "&strMerchantId&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Merchant Id is not as expected: "&strMerchantId&"",False
		End If
		
		If strCity = strIserveCity Then
			LogMessage "RSLT","Verification","The Iserve Merchant city is as expected: "&strCity&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Merchant Id is not as expected: "&strCity&"",False
		End If
		
		If strAddr1 = strIserveAddr1 Then
			LogMessage "RSLT","Verification","The Iserve Address 1 is as expected: "&strAddr1&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Address 1 is not as expected: "&strAddr1&"",False
		End If
		
		If strAddr2 = strIserveAddr2 Then
			LogMessage "RSLT","Verification","The Iserve Address 2 is as expected: "&strAddr2&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Address 2 is not as expected: "&strAddr2&"",False
		End If
		
		If strAddr3 = strIserveAddr3 Then
			LogMessage "RSLT","Verification","The Iserve Address 3 is as expected: "&strAddr3&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Address 3 is not as expected: "&strAddr3&"",False
		End If
		
		If strAddr4 = strIserveAddr4 Then
			LogMessage "RSLT","Verification","The Iserve Address 4 is as expected: "&strAddr4&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Address 4 is not as expected: "&strAddr4&"",False
		End If
		
		If strPostal = strIservePostal Then
			LogMessage "RSLT","Verification","The Iserve Postal Code is as expected: "&strPostal&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Postal Code is not as expected: "&strPostal&"",False
		End If
		
		If strPhoneNo = strIservePhoneNo Then
			LogMessage "RSLT","Verification","The Iserve Phone Number is as expected: "&strPhoneNo&"",True
		Else
	  		LogMessage "RSLT","Verification","The Iserve Phone Number is not as expected: "&strPhoneNo&"",False
		End If
	  'bcInstallmentPlan.btnOK.click
	  'WaitForICallLoading
		verifyfields_MerchantInfo =bverifyfields_MerchantInfo
End Function

'[Verify the Additional info popup exist]
Public Function verifyAdditionalInfoExist(bExist)
	bDevPending=false
   bActualExist=bcInstallmentPlan.lblPrincipalAmt_AdditionalInfo.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Additional Info Exists As Expected" ,true
       verifyAdditionalInfoExist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup Additional Info does not Exists As Expected" ,true
       verifyAdditionalInfoExist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Additional Info does not Exists As Expected" ,False
       verifyAdditionalInfoExist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Additional Info Still Exists" ,False
       verifyAdditionalInfoExist=False
   End If
End Function

'[Verify the fields for the Additional Info popup]
Public Function verifyfields_AdditionalInfo(strProFee,strAdmnFee,strNoInstallments,strInstallmentsPaid,strTrnDate)
	
	'Get the value for the fields from the I.serve
	strIserveProFee = bcInstallmentPlan.lblPrincipalAmt_ProcessingFee.GetROProperty("innertext")
	strIserveAdmnFee = bcInstallmentPlan.lblPrincipalAmt_AdminFee.GetROProperty("innertext")
	strIserveNoInstallments = bcInstallmentPlan.lblPrincipalAmt_InstallmentPeriod.GetROProperty("innertext")
	strIserveInstallmentsPaid = bcInstallmentPlan.lblPrincipalAmt_InstallmentPaid.GetROProperty("innertext")
	strIserveTrnDate = bcInstallmentPlan.lblPrincipalAmt_StatementDate.GetROProperty("innertext")
	
	bverifyfields_AdditionalInfo = true
	
	If strProFee = strIserveProFee Then
		LogMessage "RSLT","Verification","The Iserve Processing Fee is as expected: "&strProFee&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Processing Fee is not as expected: "&strProFee&"",False
	End If
	
	If strAdmnFee = strIserveAdmnFee Then
		LogMessage "RSLT","Verification","The Iserve Admin Fee is as expected: "&strAdmnFee&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Admin Fee is not as expected: "&strAdmnFee&"",False
	End If
	
	If strNoInstallments = strIserveNoInstallments Then
		LogMessage "RSLT","Verification","The Iserve No of Installments is as expected: "&strNoInstallments&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve No of Installments is not as expected: "&strNoInstallments&"",False
	End If
	
	If strInstallmentsPaid = strIserveInstallmentsPaid Then
		LogMessage "RSLT","Verification","The Iserve Installments Paid is as expected:"&strInstallmentsPaid&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Installments Paid is not as expected:"&strInstallmentsPaid&"",False
	End If
	
	If strTrnDate = strIserveTrnDate Then
		LogMessage "RSLT","Verification","The Iserve Transaction Date is as expected:"&strTrnDate&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Transaction Date is not as expected:"&strTrnDate&"",False
	End If
	WaitForICallLoading
	bcInstallmentPlan.btnOK.click
	WaitForICallLoading
	verifyfields_AdditionalInfo = bverifyfields_AdditionalInfo
End Function

'[Verify Pink Panel in Installment enquiry screen]
Public Function verifyInstallmentPlanPinkPanel(strProduct,strSubProduct,strAccNo,strAccName,strStatus,strCurrency,strAccInd,strOpenDate)
	bverifyInstallmentPlanPinkPanel=true
	If Not IsNull (strProduct) Then
		If Not verifyInnerText(bcInstallmentPlan.lblPinkPanel_Product(),strProduct, "Product") Then
			bverifyInstallmentPlanPinkPanel=false
		End If
	End If
	
	If Not IsNull (strSubProduct) Then
		If Not verifyInnerText(bcInstallmentPlan.lblPinkPanel_SubProduct(),strSubProduct, "Sub Product") Then
			bverifyInstallmentPlanPinkPanel=false
		End If
	End If
	
	If Not IsNull (strAccNo) Then
		If Not verifyInnerText(bcInstallmentPlan.lblPinkPanel_AccountNo(),strAccNo, "Account/Card No.") Then
			bverifyInstallmentPlanPinkPanel=false
		End If
	End If
	
	If Not IsNull (strAccName) Then
		If Not verifyInnerText(bcInstallmentPlan.lblPinkPanel_Name(),strAccName, "Account Name") Then
			bverifyInstallmentPlanPinkPanel=false
		End If
	End If
	
	If Not IsNull (strStatus) Then
		If Not verifyInnerText(bcInstallmentPlan.lblPinkPanel_AccountNo(),strStatus, "Status") Then
			bverifyInstallmentPlanPinkPanel=false
		End If
	End If
	
	If Not IsNull (strCurrency) Then
		If Not verifyInnerText(bcInstallmentPlan.lblPinkPanel_Currency(),strCurrency, "Currency") Then
			bverifyInstallmentPlanPinkPanel=false
		End If
	End If
	verifyInstallmentPlanPinkPanel=bverifyInstallmentPlanPinkPanel
End Function

'[Click on OK button of the popup for installment plan]
Public Function clickOK_InstallmentPlan()
	bDevPending=true
   bcInstallmentPlan.btnOK.click
   If Err.Number<>0 Then
       clickOK_InstallmentPlan=false
            LogMessage "RSLT","Verification","Failed to Click Button : OK" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickOK_InstallmentPlan=true
End Function
