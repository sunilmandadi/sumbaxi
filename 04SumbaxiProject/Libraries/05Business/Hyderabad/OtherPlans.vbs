'[Verify and Click OtherPlans Link from leftMenu]
Public Function ClickLink_OtherPlans()
bClickLink_OtherPlans=true
	bcAccountOverview_LeftMenu.btnOtherPlans.Click
	WaitForIcallLoading
	If Err.Number<>0 Then
       bClickLink_OtherPlans=false
       LogMessage "RSLT","Verification","Failed to Click Link  : OtherPlans" ,false
       Exit Function
	End If
	Wait 1
	waitForIcallLoading	
ClickLink_OtherPlans = bClickLink_OtherPlans
End Function 

'[Verify Loan Information section displayed only for Loan Accounts]
Public Function VerifyLoanSectiondisplay(strProduct)
   bVerifyLoanSectiondisplay=True
   If Not IsNull(strProduct) Then
   		If strProduct = "LOAN" Then
   		   If Not VerifyInnerText (OtherPlans.lblLoanInfo(), "Loan Information", "Loan Section")Then
           bVerifyLoanSectiondisplay=False
       	   End If
       	ElseIf strProduct = "CRCARD" OR strProduct = "CASHLINE" Then
      	   If Not VerifyInnerText (OtherPlans.lblLoanInfo(), "Loan Information", "Loan Section")Then
           bVerifyLoanSectiondisplay=True
       	   End If    		
   		End If
   End If
   VerifyLoanSectiondisplay=bVerifyLoanSectiondisplay
End Function

'[Verify OtherPlan details for Credit Card displayed based on the Plan No Selected from the table]
Public Function VerifyOtherPlanDetails_CC(strCardNumber,strProduct,strCardType)
	'Get the Record count in the Plan details table
	intRecordCount = getRecordsCountForColumn(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent, "Plan No")
	'For i = 0 To intRecordCount-1
	For i = 0 To intRecordCount-3
		Set objAllRows=getAllRows(OtherPlans.tblPlanSummaryContent)
		strIserveSeqNumber = getCellTextFor(OtherPlans.tblPlanSummaryHeader,objAllRows(i),i, "Seq. No.")
		j=1			
		iPage = ((j*5) - (5-i))\3 ' Page is quotient (Page refers to Host system) 
		iRow = (((j-1)*5 + i) Mod 3) + 1   ' Row is remainder
		If strCardType = "Supplementry" Then
			iRow = iRow + 1 
		End If
		'Click on the Plan No based on selected row in the table
		Call clickVaddinLink_tblCell (OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,i, "Plan No")
		WaitForIcallLoading	
			'Get all the fields values below Plan Summary section in the Other Plan page 
			strIserveBaseRate = OtherPlans.lblBaseRate.GetROProperty("innertext")
			strIserveBaseRate = Replace(strIserveBaseRate,"%","")
			strIserveCalculatedRate = OtherPlans.lblCalulatedRate.GetROProperty("innertext")
			strIserveCalculatedRate = Replace(strIserveCalculatedRate,"%","")
			strIserveInterestStartDate = OtherPlans.lblInterestStartDate.GetROProperty("innertext")			
			strIserveLTDPrincipal = OtherPlans.lblLifeToDatePrincipal.GetROProperty("innertext")
			strIserveLTDInterest = OtherPlans.lblLifeToDateInterest.GetROProperty("innertext")			
			strIserveYTDPrincipal = OtherPlans.lblYeartoDatePrincipal.GetROProperty("innertext")
			strIserveYTDInterest = OtherPlans.lblYearToDateInterest.GetROProperty("innertext")			
			strIserveServiceCharges = OtherPlans.lblServiceCharges.GetROProperty("innertext")
			strIserveOpenDate = OtherPlans.lblOpenDate.GetROProperty("innertext")
			strIserveBalTransferMonthsRemaining = OtherPlans.lblBalanceTransferMonthsRemain.GetROProperty("innertext")		
			strIserveBalTransferExpiryDate = OtherPlans.lblBalanceTransferExpiry.GetROProperty("innertext")		
			strIserveAccuredInterest = OtherPlans.lblAccruedInterest.GetROProperty("innertext")
			strIservePerDiem = OtherPlans.lblPerDiem.GetROProperty("innertext")		
			strIserveNormalInterestBeginDate = OtherPlans.lblNormalInterestBeginDate.GetROProperty("innertext")	
	
			'Call function in Visionplus validation to validate the values in Host System
		Call getOtherDetails_CC_UL_ARQA_Vplus(strCardNumber,strIserveSeqNumber,i,strProduct,iRow)	

			'from screen1 	
			strVPStrPlan=Environment.Value("strPlan")
			strVPStrPlanDesc=Environment.Value("strPlanDesc")
			strVPStrCurrBalance=Environment.Value("strCurBalance")
			'from screen 2
			strVPBaseRate=Environment.Value("StrBaseRate")
			strVPCalculatedRate=Environment.Value("StrCalcRate") 
			strVPInterestStartDate=Environment.Value("StrBeginingDate")	
			'from screen 3
			strVPLTDPrincipal = Environment.Value("StrLTDPaidPrincipal")
			strVPLTDInterest = Environment.Value("StrLTDPaidInterest")
			'from screen 3	
			strVPYTDPrincipal = Environment.Value("StrYTDPaidPrincipal")
			strVPYTDInterest = Environment.Value("StrYTDPaidInterest")
			'from screen 4
			strVPServiceCharges = Environment.Value("StrServiceCharges")
			'from screen 5
			strVPOpenDate = Environment.Value("StrOpenDate")
			strVPBalTransferMonthsRemaining = Environment.Value("StrBalTransferMonthlyRemain")
			strVPBalTransferExpiryDate = Environment.Value("StrBalTransferExpDate")	
			strVPAccuredInterest = Environment.Value("StrAccuredInterest")
			strVPPerDiem = Environment.Value("StrPerDiem")	
			'from screen 7		
			strVPNormalInterestBeginDate = Environment.Value("StrNormalInterestBeginDate")
			'Defereable Information from Screen7 
			strVPDefferedInterestOrginial = Environment.Value("StrDeferredInterestORIG")
			strVPDefferedInsuranceOrginial = Environment.Value("StrDeferredInsuranceORIG")	
			strVPDefferedBillingOrginial = Environment.Value("StrDeferredBillingORIG")
			strVPDefferedPaymentOrginial = Environment.Value("StrDeferredPaymentORIG")	
			
			strVPDefferedInterestPeriod = Environment.Value("StrDeferredInterestPeriod")
			strVPDefferedInsurancePeriod = Environment.Value("StrDeferredInsurancePeriod")	
			strVPDefferedBillingPeriod = Environment.Value("StrDeferredBillingPeriod")
			strVPDefferedPaymentPeriod = Environment.Value("StrDeferredPaymentPeriod")	
			
			strVPDefferedInterestRemaining = Environment.Value("StrDeferredInterestREM")
			strVPDefferedInsuranceRemaining = Environment.Value("StrDeferredInsuranceREM")	
			strVPDefferedBillingRemaining = Environment.Value("StrDeferredBillingREM")
			strVPDefferedPaymentRemaining = Environment.Value("StrDeferredPaymentREM")	
			
			'Verify the table contents on Plan details table
			lstlstPlandetailstable= (checknull("(Plan No:"&strVPStrPlan&"|Plan Description:"&strVPStrPlanDesc&"|Current Balance:"&strVPStrCurrBalance&")|"))
			verifytablePlandetails_OtherPlan=verifyTableContentList(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,lstlstPlandetailstable,"Plan Details table",false,null,null,null)
				
		'Compare the values from Iserve below Plan Summary Section with the values retrived from Vision Plus System
		If  (Ucase(Trim(strVPBaseRate)) < 0 AND Ucase(Trim(strIserveBaseRate)) = "" ) OR  (Ucase(Trim(strIserveBaseRate)) =  Ucase(Trim(strVPBaseRate))) Then
			LogMessage "RSLT","Verification","Base Rate field value matched with the expected Value Expected: "&strVPBaseRate&" Actual: "&strIserveBaseRate&"",True
		Else
			LogMessage "RSLT","Verification","Base Rate field value doesnt match with the expected Value Expected: "&strVPBaseRate&" Actual: "&strIserveBaseRate&"",False	
		End IF 
		
		If  (Ucase(Trim(strVPCalculatedRate)) < 0 AND Ucase(Trim(strIserveCalculatedRate)) = "" ) OR  (Ucase(Trim(strVPCalculatedRate)) =  Ucase(Trim(strIserveCalculatedRate))) Then
			LogMessage "RSLT","Verification","Base Rate field value matched with the expected Value Expected: "&strVPCalculatedRate&" Actual: "&strIserveCalculatedRate&"",True
		Else
			LogMessage "RSLT","Verification","Base Rate field value doesnt match with the expected Value Expected: "&strVPCalculatedRate&" Actual: "&strIserveCalculatedRate&"",False	
		End IF 
				
		IF Ucase(Trim(strIserveInterestStartDate)) =  Ucase(Trim(strVPInterestStartDate)) Then
			LogMessage "RSLT","Verification","Interest Start Date field value matched with the expected Value Actual: "&strIserveInterestStartDate&" Expected: "&strVPInterestStartDate&"",True
		Else
			LogMessage "RSLT","Verification","Interest Start Date field value doesnt match with the expected Value Actual: "&strIserveInterestStartDate&" Expected: "&strVPInterestStartDate&"",False	
		End IF
		
		IF (Ucase(Trim(strVPLTDPrincipal)) < 0 AND Ucase(Trim(strIserveLTDPrincipal)) = "" ) OR  (Ucase(Trim(strIserveLTDPrincipal)) =  Ucase(Trim(strVPLTDPrincipal))) Then
			LogMessage "RSLT","Verification","Life to Date Principal value matched with the expected Value Actual: "&strIserveLTDPrincipal&" Expected: "&strVPLTDPrincipal&"",True
		Else
			LogMessage "RSLT","Verification","Life to Date Principal value doesnt match with the expected Value Actual: "&strIserveLTDPrincipal&" Expected: "&strVPLTDPrincipal&"",False	
		End IF 

		IF (Ucase(Trim(strVPLTDInterest)) < 0 AND Ucase(Trim(strIserveLTDInterest)) = "" ) OR (Ucase(Trim(strIserveLTDInterest)) =  Ucase(Trim(strVPLTDInterest))) Then
			LogMessage "RSLT","Verification","Life to Date Interest field value matched with the expected Value Actual: "&strIserveLTDInterest&" Expected: "&strVPLTDInterest&"",True
		Else
			LogMessage "RSLT","Verification","Life to Date Interest field value doesnt match with the expected Value Actual: "&strIserveLTDInterest&" Expected: "&strVPLTDInterest&"",False	
		End IF
	
		IF Ucase(Trim(strIserveYTDPrincipal)) =  Ucase(Trim(strVPYTDPrincipal)) Then
			LogMessage "RSLT","Verification","Year to Date Principal value matched with the expected Value Actual: "&strIserveYTDPrincipal&" Expected: "&strVPYTDPrincipal&"",True
		Else
			LogMessage "RSLT","Verification","Year to Date Principal value doesnt match with the expected Value Actual: "&strIserveYTDPrincipal&" Expected: "&strVPYTDPrincipal&"",False	
		End IF 

		IF Ucase(Trim(strIserveYTDInterest)) =  Ucase(Trim(strVPYTDInterest)) Then
			LogMessage "RSLT","Verification","Year to Date Interest field value matched with the expected Value Actual: "&strIserveYTDInterest&" Expected: "&strVPYTDInterest&"",True
		Else
			LogMessage "RSLT","Verification","Year to Date Interest field value doesnt match with the expected Value Actual: "&strIserveYTDInterest&" Expected: "&strVPYTDInterest&"",False	
		End IF

		IF Ucase(Trim(strIserveServiceCharges)) =  Ucase(Trim(strVPServiceCharges)) Then
			LogMessage "RSLT","Verification","Service Charges field value matched with the expected Value Actual: "&strIserveServiceCharges&" Expected: "&strVPServiceCharges&"",True
		Else
			LogMessage "RSLT","Verification","Service Charges field value doesnt match with the expected Value Actual: "&strIserveServiceCharges&" Expected: "&strVPServiceCharges&"",False	
		End IF 

		IF Ucase(Trim(strIserveOpenDate)) =  Ucase(Trim(strVPOpenDate)) Then
			LogMessage "RSLT","Verification","Open Date field value matched with the expected Value Actual: "&strIserveOpenDate&" Expected: "&strVPOpenDate&"",True
		Else
			LogMessage "RSLT","Verification","Open Date field value doesnt match with the expected Value Actual: "&strIserveOpenDate&" Expected: "&strVPOpenDate&"",False	
		End IF

		IF Ucase(Trim(strIserveBalTransferMonthsRemaining)) =  Ucase(Trim(strVPBalTransferMonthsRemaining)) Then
			LogMessage "RSLT","Verification","Balance Transfer Monthly Remaining field value matched with the expected Value Actual: "&strIserveBalTransferMonthsRemaining&" Expected: "&strVPBalTransferMonthsRemaining&"",True
		Else
			LogMessage "RSLT","Verification","Balance Transfer Monthly Remaining field value doesnt match with the expected Value Actual: "&strIserveBalTransferMonthsRemaining&" Expected: "&strVPBalTransferMonthsRemaining&"",False	
		End IF
	
		IF Ucase(Trim(strIserveBalTransferExpiryDate)) =  Ucase(Trim(strVPBalTransferExpiryDate)) Then
			LogMessage "RSLT","Verification","Balance transfer Expiry Date field value matched with the expected Value Actual: "&strIserveBalTransferExpiryDate&" Expected: "&strVPBalTransferExpiryDate&"",True
		Else
			LogMessage "RSLT","Verification","Balance transfer Expiry Date field value doesnt match with the expected Value Actual: "&strIserveBalTransferExpiryDate&" Expected: "&strVPBalTransferExpiryDate&"",False	
		End IF 

		IF (Ucase(Trim(strIserveAccuredInterest)) =  Ucase(Trim(strVPAccuredInterest))) OR (Ucase(Trim(strIserveAccuredInterest)) = 0.00 AND  Ucase(Trim(strVPAccuredInterest)) = 0) Then
			LogMessage "RSLT","Verification","Accured Interest field value matched with the expected Value Actual: "&strIserveAccuredInterest&" Expected: "&strVPAccuredInterest&"",True
		Else
			LogMessage "RSLT","Verification","Accured Interest field value doesnt match with the expected Value Actual: "&strIserveAccuredInterest&" Expected: "&strVPAccuredInterest&"",False	
		End IF

		IF (Ucase(Trim(strIservePerDiem)) =  Ucase(Trim(strVPPerDiem))) OR (Ucase(Trim(strIservePerDiem)) = 0.00 AND  Ucase(Trim(strVPPerDiem)) = 0 ) Then
			LogMessage "RSLT","Verification","Per Diem field value matched with the expected Value Actual: "&strIservePerDiem&" Expected: "&strVPPerDiem&"",True
		Else
			LogMessage "RSLT","Verification","Per Diem field value doesnt match with the expected Value Actual: "&strIservePerDiem&" Expected: "&strVPPerDiem&"",False	
		End IF 

		IF Ucase(Trim(strIserveNormalInterestBeginDate)) =  Ucase(Trim(strVPNormalInterestBeginDate)) Then
			LogMessage "RSLT","Verification","Normal Interest Begin Date field value matched with the expected Value Actual: "&strIserveNormalInterestBeginDate&" Expected: "&strVPNormalInterestBeginDate&"",True
		Else
			LogMessage "RSLT","Verification","Normal Interest Begin Date field value doesnt match with the expected Value Actual: "&strIserveNormalInterestBeginDate&" Expected: "&strVPNormalInterestBeginDate&"",False	
		End IF
		
		'Verify the Deferable Information 
		lstlstDeferableInformationtable= (checknull("(Deferral Information:Interest|Period:"&strVPDefferedInterestPeriod&"|Original:"&strVPDefferedInterestOrginial&"|Remaining:"&strVPDefferedInterestRemaining&")|(Deferral Information:Insurance|Period:"&strVPDefferedInsurancePeriod&"|Original:"&strVPDefferedInsuranceOrginial&"|Remaining:"&strVPDefferedInsuranceRemaining&")|(Deferral Information:Billing|Period:"&strVPDefferedBillingPeriod&"|Original:"&strVPDefferedBillingOrginial&"|Remaining:"&strVPDefferedBillingRemaining&")|(Deferral Information:Payment|Period:"&strVPDefferedPaymentPeriod&"|Original:"&strVPDefferedPaymentOrginial&"|Remaining:"&strVPDefferedPaymentRemaining&")|"))
		verifytableDefferalInfo_OtherPlan=verifyTableContentList(OtherPlans.tblDefferalInformationHeader,OtherPlans.tblDefferalInformationContent,lstlstDeferableInformationtable,"Deferal Information",false,null,null,null)		
	Next
End Function

'[Verify OtherPlan details for Product loan displayed based on the Plan No Selected from the table]
Public Function VerifyOtherPlanDetails_UL(strCardNumber,strProduct)
	'Get the Record count in the Plan details table
	intRecordCount = getRecordsCountForColumn(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent, "Plan No")
	For i = 0 To intRecordCount - 1
		Set objAllRows=getAllRows(OtherPlans.tblPlanSummaryContent)
		strIserveSeqNumber = getCellTextFor(OtherPlans.tblPlanSummaryHeader,objAllRows(i),i, "Seq. No.")	
		j=1	
		iPage = ((j*5) - (5-i))\3 ' Page is quotient (Page refers to Host system) 
		iRow = (((j-1)*5 + i) Mod 3) + 1   ' Row is remainder
		'Click on the Plan No based on selected row in the table
		Call clickVaddinLink_tblCell (OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,i, "Plan No")
		WaitForIcallLoading
			'Get all the fields values below Plan Summary section in the Other Plan page to be displayed for LOAN and Cashline Products
			strIserveBaseRate = OtherPlans.lblBaseRate.GetROProperty("innertext")
			strIserveCalculatedRate = OtherPlans.lblCalulatedRate.GetROProperty("innertext")
			strIserveCalculatedRate = Replace(strIserveCalculatedRate,"%","")
			strIserveInterestStartDate = OtherPlans.lblInterestStartDate.GetROProperty("innertext")
			strIserveLTDPrincipal = OtherPlans.lblLifeToDatePrincipal.GetROProperty("innertext")
			strIserveLTDInterest = OtherPlans.lblLifeToDateInterest.GetROProperty("innertext")			
			strIserveYTDPrincipal = OtherPlans.lblYeartoDatePrincipal.GetROProperty("innertext")
			strIserveYTDInterest = OtherPlans.lblYearToDateInterest.GetROProperty("innertext")			
			strIserveServiceCharges = OtherPlans.lblServiceCharges.GetROProperty("innertext")
			strIserveOpenDate = OtherPlans.lblOpenDate.GetROProperty("innertext")
			strIserveBalTransferMonthsRemaining = OtherPlans.lblBalanceTransferMonthsRemain.GetROProperty("innertext")		
			strIserveBalTransferExpiryDate = OtherPlans.lblBalanceTransferExpiry.GetROProperty("innertext")		
			strIserveAccuredInterest = OtherPlans.lblAccruedInterest.GetROProperty("innertext")
			strIservePerDiem = OtherPlans.lblPerDiem.GetROProperty("innertext")		
			strIserveNormalInterestBeginDate = OtherPlans.lblNormalInterestBeginDate.GetROProperty("innertext")	
		' LOAN Information sections 
			strIserveInitialTerm = OtherPlans.lblInitialTerm.GetROProperty("innertext")
			strIserveCurrentTerm = OtherPlans.lblCurrentTerm.GetROProperty("innertext")
			strIserveLoanAmount = OtherPlans.lblLoanAmount.GetROProperty("innertext")		
			strIservePrincipalAmount = OtherPlans.lblPrincipalAmount.GetROProperty("innertext")		
			strIserveInterestAmount = OtherPlans.lblInterestAmount.GetROProperty("innertext")
			strIserveFirstPaymentAmount = OtherPlans.lblFirstPaymentAmount.GetROProperty("innertext")		
			strIserveFirstPaymentDate= OtherPlans.lblFirstPaymentDate.GetROProperty("innertext")				
			strIserveFinalPaymentDate= OtherPlans.lblFinalPayementDate.GetROProperty("innertext")				
			strIserveTotalNoOfInstallments= OtherPlans.lblTotalInstallment.GetROProperty("innertext")				
			strIserveRemainingNoOfInstallments= OtherPlans.lblRemainingInstallment.GetROProperty("innertext")				
			strIserveAnnualPercentageCode= OtherPlans.lblAnnualPercentageCode.GetROProperty("innertext")
			strIserveAnnualPercentageCode = Replace(strIserveAnnualPercentageCode,"%","")			
			strIserveTotalDisbursedAmount= OtherPlans.lblTotalDisbursedAmount.GetROProperty("innertext")			
		'Call function in Visionplus validation to validate the values in Host System
	Call getOtherDetails_CC_UL_ARQA_Vplus(strCardNumber,strIserveSeqNumber,i,strProduct,iRow)	
		
			'from screen1 	
			strVPStrPlan=Environment.Value("strPlan")
			strVPStrPlanDesc=Environment.Value("strPlanDesc")
			strVPStrCurrBalance=Environment.Value("strCurBalance")
			'from screen 2
			strVPCalculatedRate=Environment.Value("StrCalcRate")
			strVPInterestStartDate=Environment.Value("StrBeginingDate")	
			'from screen 3
			strVPLTDPrincipal = Environment.Value("StrLTDPaidPrincipal")
			strVPLTDInterest = Environment.Value("StrLTDPaidInterest")
			'from screen 3	
			strVPYTDPrincipal = Environment.Value("StrYTDPaidPrincipal")
			strVPYTDInterest = Environment.Value("StrYTDPaidInterest")
			'from screen 5 
			strVPOpenDate = Environment.Value("StrOpenDate")
			strVPAccuredInterest = Environment.Value("StrAccuredInterest")
			strVPPerDiem = Environment.Value("StrPerDiem")	
			'Deferral Information from Screen7 
			strVPDefferedInterestOrginial = Environment.Value("StrDeferredInterestORIG")
			strVPDefferedInsuranceOrginial = Environment.Value("StrDeferredInsuranceORIG")	
			strVPDefferedBillingOrginial = Environment.Value("StrDeferredBillingORIG")
			strVPDefferedPaymentOrginial = Environment.Value("StrDeferredPaymentORIG")	
			strVPDefferedInterestPeriod = Environment.Value("StrDeferredInterestPeriod")
			strVPDefferedInsurancePeriod = Environment.Value("StrDeferredInsurancePeriod")	
			strVPDefferedBillingPeriod = Environment.Value("StrDeferredBillingPeriod")
			strVPDefferedPaymentPeriod = Environment.Value("StrDeferredPaymentPeriod")	
			strVPDefferedInterestRemaining = Environment.Value("StrDeferredInterestREM")
			strVPDefferedInsuranceRemaining = Environment.Value("StrDeferredInsuranceREM")	
			strVPDefferedBillingRemaining = Environment.Value("StrDeferredBillingREM")
			strVPDefferedPaymentRemaining = Environment.Value("StrDeferredPaymentREM")		
			'from screen 10		
			strVPTotalNoOfInstallments = Environment.Value("StrTotalNoOfInstallments")	
			'from screen 11				
			strVPInitialTerm = Environment.Value("StrInitialTerm")  
			strVPCurrentTerm = Environment.Value("StrCurrentTerm") 
			strVPLoanAmount = Environment.Value("StrLoanAmount")
			strVPRemainingNoOfInstallments =  Environment.Value("StrRemainingTerm")
			strVPFirstPaymentAmount = Environment.Value("StrFirstPaymentAmount")
			strVPFinalPaymentAmount = Environment.Value("StrFinalPaymentAmount")
			strVPFirstPaymentDate = Environment.Value("StrFirstPaymentDate")
			strVPFinalPaymentDate = Environment.Value("StrFinalPaymentDate") 			
			'from screen 12			
			strVPPrincipalAmount = Environment.Value("StrPrincipalAmount")
			strVPInterestAmount = Environment.Value("StrInterestAmount")
			strVPAnnualPercentageCode= Environment.Value("StrAnnualPercentageCode")
			'from screen 40		
			strVPTotalDisbursedAmount= Environment.Value("StrTotalDisbursableAmount")

			'Verify the table contents on Plan details table
			lstlstPlandetailstable= (checknull("(Plan No:"&strVPStrPlan&"|Plan Description:"&strVPStrPlanDesc&"|Current Balance:"&strVPStrCurrBalance&")|"))
			verifytablePlandetails_OtherPlan=verifyTableContentList(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,lstlstPlandetailstable,"Plan Details table",false,null,null,null)
			
			'Verify the Deferable Information 
			lstlstDeferableInformationtable= (checknull("(Deferral Information:Interest|Period:"&strVPDefferedInterestPeriod&"|Original:"&strVPDefferedInterestOrginial&"|Remaining:"&strVPDefferedInterestRemaining&")|(Deferral Information:Insurance|Period:"&strVPDefferedInsurancePeriod&"|Original:"&strVPDefferedInsuranceOrginial&"|Remaining:"&strVPDefferedInsuranceRemaining&")|(Deferral Information:Billing|Period:"&strVPDefferedPaymentPeriod&"|Original:"&strVPDefferedBillingOrginial&"|Remaining:"&strVPDefferedBillingRemaining&")|(Deferral Information:Payment|Period:"&strVPDefferedInterestOrginial&"|Original:"&strVPDefferedPaymentOrginial&"|Remaining:"&strVPDefferedPaymentRemaining&")|"))
			verifytableDefferalInfo_OtherPlan=verifyTableContentList(OtherPlans.tblDefferalInformationHeader,OtherPlans.tblDefferalInformationContent,lstlstDeferableInformationtable,"Deferal Information",false,null,null,null)		
	
			'Compare the values from Iserve below Plan Summary Section with the values retrived from Vision Plus System
			If  Trim(strIserveBaseRate) =  "" Then
				LogMessage "RSLT","Verification","Base Rate field displayed Null for LOAN or Cashline Products as Expected",True
			Else
				LogMessage "RSLT","Verification","Base Rate field is popualated for Loan and cashline Products ",False	
			End IF 
			
			If (Ucase(Trim(strIserveCalculatedRate)) < 0 AND Ucase(Trim(strVPCalculatedRate)) = "" ) OR  (Ucase(Trim(strIserveCalculatedRate)) =  Ucase(Trim(strVPCalculatedRate))) Then
				LogMessage "RSLT","Verification","Calculated Rate field value matched with the expected Value Expected: "&strVPCalculatedRate&" Actual: "&strIserveCalculatedRate&"",True
			Else
				LogMessage "RSLT","Verification","Calculated Rate field value doesnt match with the expected Value Expected: "&strVPCalculatedRate&" Actual: "&strIserveCalculatedRate&"",False	
			End IF 
			
			IF Ucase(Trim(strIserveInterestStartDate)) =  Ucase(Trim(strVPInterestStartDate)) Then
				LogMessage "RSLT","Verification","Interest Start Date field value matched with the expected Value Expected: "&strVPInterestStartDate&" Actual: "&strIserveInterestStartDate&"",True
			Else
				LogMessage "RSLT","Verification","Interest Start Date field value doesnt match with the expected Value Expected: "&strVPInterestStartDate&" Actual: "&strIserveInterestStartDate&"",False	
			End IF
			
			IF (Ucase(Trim(strVPLTDPrincipal)) < 0 AND Ucase(Trim(strIserveLTDPrincipal)) = "" ) OR  (Ucase(Trim(strIserveLTDPrincipal)) =  Ucase(Trim(strVPLTDPrincipal))) Then
				LogMessage "RSLT","Verification","Life to Date Principal value matched with the expected Value Expected: "&strVPLTDPrincipal&" Actual: "&strIserveLTDPrincipal&"",True
			Else
				LogMessage "RSLT","Verification","Life to Date Principal value doesnt match with the expected Value Expected: "&strVPLTDPrincipal&" Actual: "&strIserveLTDPrincipal&"",False	
			End IF 
	
			IF (Ucase(Trim(strVPLTDInterest)) < 0 AND Ucase(Trim(strIserveLTDInterest)) = "" ) OR (Ucase(Trim(strIserveLTDInterest)) =  Ucase(Trim(strVPLTDInterest))) Then
				LogMessage "RSLT","Verification","Life to Date Interest field value matched with the expected Value Expected: "&strVPLTDInterest&" Actual: "&strIserveLTDInterest&"",True
			Else
				LogMessage "RSLT","Verification","Life to Date Interest field value doesnt match with the expected Value Expected: "&strVPLTDInterest&" Actual: "&strIserveLTDInterest&"",False	
			End IF
		
			IF Ucase(Trim(strIserveYTDPrincipal)) =  Ucase(Trim(strVPYTDPrincipal)) Then
				LogMessage "RSLT","Verification","Year to Date Principal value matched with the expected Value Expected: "&strVPYTDPrincipal&" Actual: "&strIserveYTDPrincipal&"",True
			Else
				LogMessage "RSLT","Verification","Year to Date Principal value doesnt match with the expected Value Expected: "&strVPYTDPrincipal&" Actual: "&strIserveYTDPrincipal&"",False	
			End IF 
	
			IF Ucase(Trim(strIserveYTDInterest)) =  Ucase(Trim(strVPYTDInterest)) Then
				LogMessage "RSLT","Verification","Year to Date Interest field value matched with the expected Value Expected: "&strVPYTDInterest&" Actual: "&strIserveYTDInterest&"",True
			Else
				LogMessage "RSLT","Verification","Year to Date Interest field value doesnt match with the expected Value Expected: "&strVPYTDInterest&" Actual: "&strIserveYTDInterest&"",False	
			End IF
	
			IF Trim(strIserveServiceCharges)= "" Then
				LogMessage "RSLT","Verification","Service Charges are displayed Null for LOAN and cashline Products as expected",True
			Else
				LogMessage "RSLT","Verification","Service Charges are displayed for LOAN and cashline Products",False	
			End IF 
	
			IF Ucase(Trim(strIserveOpenDate)) =  Ucase(Trim(strVPOpenDate)) Then
				LogMessage "RSLT","Verification","Open Date field value matched with the expected Value Expected: "&strVPOpenDate&" Actual: "&strIserveOpenDate&"",True
			Else
				LogMessage "RSLT","Verification","Open Date field value doesnt match with the expected Value Expected: "&strVPOpenDate&" Actual: "&strIserveOpenDate&"",False	
			End IF
	
			IF Trim(strIserveBalTransferMonthsRemaining) =  "" Then		
				LogMessage "RSLT","Verification","Balance Transfer Monthly Remaining field are displayed Null for LOAN and cashline Products as expected",True
			Else
				LogMessage "RSLT","Verification","Balance Transfer Monthly Remaining fieldare displayed for LOAN and cashline Products",False	
			End IF 
		
			IF Trim(strIserveBalTransferExpiryDate) = "" Then
				LogMessage "RSLT","Verification","Balance transfer Expiry Date field value matched with the expected Value Expected: "&strVPBalTransferExpiryDate&" Actual: "&strIserveBalTransferExpiryDate&"",True
			Else
				LogMessage "RSLT","Verification","Balance transfer Expiry Date field value doesnt match with the expected Value Expected: "&strVPBalTransferExpiryDate&" Actual: "&strIserveBalTransferExpiryDate&"",False	
			End IF 
	
			IF (Ucase(Trim(strIserveAccuredInterest)) =  Ucase(Trim(strVPAccuredInterest)))  OR (Ucase(Trim(strIserveAccuredInterest)) = 0.00 AND  Ucase(Trim(strVPAccuredInterest)) = 0) Then 
				LogMessage "RSLT","Verification","Accured Interest field value matched with the expected Value Expected: "&strVPAccuredInterest&" Actual: "&strIserveAccuredInterest&"",True
			Else
				LogMessage "RSLT","Verification","Accured Interest field value doesnt match with the expected Value Expected: "&strVPAccuredInterest&" Actual: "&strIserveAccuredInterest&"",False	
			End IF
	
			IF (Ucase(Trim(strIservePerDiem)) =  Ucase(Trim(strVPPerDiem))) OR (Ucase(Trim(strIservePerDiem)) = 0.00 AND  Ucase(Trim(strVPPerDiem)) = 0 )Then
				LogMessage "RSLT","Verification","Per Diem field value matched with the expected Value Expected: "&strVPPerDiem&" Actual: "&strIservePerDiem&"",True
			Else
				LogMessage "RSLT","Verification","Per Diem field value doesnt match with the expected Value Expected: "&strVPPerDiem&" Actual: "&strIservePerDiem&"",False	
			End IF 
	
			IF Trim(strIserveNormalInterestBeginDate) = "" Then
				LogMessage "RSLT","Verification","Normal Interesting Begin Date field value matched with the expected Value Expected: "&strVPNormalInterestBeginDate&" Actual: "&strIserveNormalInterestBeginDate&"",True
			Else
				LogMessage "RSLT","Verification","Normal Interesting Begin Date field value doesnt match with the expected Value Expected: "&strVPNormalInterestBeginDate&" Actual: "&strIserveNormalInterestBeginDate&"",False	
			End IF
		
			IF (Ucase(Trim(strIserveInitialTerm)) =  Ucase(Trim(strVPInitialTerm))) OR (Ucase(Trim(strIserveInitialTerm)) = 0 AND   Ucase(Trim(strVPInitialTerm)) = 0) Then
				LogMessage "RSLT","Verification","InitialTerm displayed in Loan Information section matched with the expected Value Expected: "&strVPInitialTerm&" Actual: "&strIserveInitialTerm&"",True
			Else
				LogMessage "RSLT","Verification","InitialTerm displayed in Loan Information section doesnt match with the expected Value Expected: "&strVPInitialTerm&" Actual: "&strIserveInitialTerm&"",False	
			End IF
		
			IF (Ucase(Trim(strIserveCurrentTerm)) =  Ucase(Trim(strVPCurrentTerm)))  OR (Ucase(Trim(strIserveCurrentTerm)) = 0 AND  Ucase(Trim(strVPCurrentTerm))) = 0 Then
				LogMessage "RSLT","Verification","Current Term Installment field value matched with the expected Value Expected: "&strVPCurrentTerm&" Actual: "&strIserveCurrentTerm&"",True
			Else
				LogMessage "RSLT","Verification","Current Term Installment field value doesnt match with the expected Value Expected: "&strVPCurrentTerm&" Actual: "&strIserveCurrentTerm&"",False	
			End IF 
			
			IF (Ucase(Trim(strIserveLoanAmount)) =  Ucase(Trim(strVPLoanAmount))) OR (Ucase(Trim(strIserveLoanAmount)) = 0.00  AND Ucase(Trim(strVPLoanAmount))= 0 )Then
				LogMessage "RSLT","Verification","Loan Amount field value matched with the expected Value Expected: "&strVPLoanAmount&" Actual: "&strIserveLoanAmount&"",True
			Else
				LogMessage "RSLT","Verification","Loan Amount field value doesnt match with the expected Value Expected: "&strVPLoanAmount&" Actual: "&strIserveLoanAmount&"",False	
			End IF 
			
			IF (Ucase(Trim(strIservePrincipalAmount)) =  Ucase(Trim(strVPPrincipalAmount)))  OR (Ucase(Trim(strIservePrincipalAmount)) = 0.00 AND Ucase(Trim(strVPPrincipalAmount))= 0 )Then
				LogMessage "RSLT","Verification","Principal Amount field value matched with the expected Value Expected: "&strVPPrincipalAmount&" Actual: "&strIservePrincipalAmount&"",True
			Else
				LogMessage "RSLT","Verification","Principal Amount field value doesnt match with the expected Value Expected: "&strVPPrincipalAmount&" Actual: "&strIservePrincipalAmount&"",False	
			End IF	
			
			IF (Ucase(Trim(strIserveInterestAmount)) =  Ucase(Trim(strVPInterestAmount))) OR (Ucase(Trim(strIserveInterestAmount)) = 0.00 AND Ucase(Trim(strVPInterestAmount))= 0 )Then
				LogMessage "RSLT","Verification","Interest Amount field value matched with the expected Value Expected: "&strVPInterestAmount&" Actual: "&strIserveInterestAmount&"",True
			Else
				LogMessage "RSLT","Verification","Interest Amount field value doesnt match with the expected Value Expected: "&strVPInterestAmount&" Actual: "&strIserveInterestAmount&"",False	
			End IF	
			
			IF (Ucase(Trim(strIserveFirstPaymentAmount)) =  Ucase(Trim(strVPFirstPaymentAmount))) OR (Ucase(Trim(strIserveFirstPaymentAmount)) = 0.00 AND Ucase(Trim(strVPFirstPaymentAmount))= 0 )Then
				LogMessage "RSLT","Verification","First Payment Amount field value matched with the expected Value Expected: "&strVPFirstPaymentAmount&" Actual: "&strIserveFirstPaymentAmount&"",True
			Else
				LogMessage "RSLT","Verification","First Payment Amount field value doesnt match with the expected Value Expected: "&strVPFirstPaymentAmount&" Actual: "&strIserveFirstPaymentAmount&"",False	
			End IF	
			
			IF Ucase(Trim(strIserveFirstPaymentDate)) =  Ucase(Trim(strVPFirstPaymentDate)) Then 
				LogMessage "RSLT","Verification","First Payment Date value matched with the expected Value Expected: "&strVPFirstPaymentDate&" Actual: "&strIserveFirstPaymentDate&"",True
			Else
				LogMessage "RSLT","Verification","First Payment Date value doesnt match with the expected Value Expected: "&strVPFirstPaymentDate&" Actual: "&strIserveFirstPaymentDate&"",False	
			End IF	
			
			IF Ucase(Trim(strIserveFinalPaymentDate)) =  Ucase(Trim(strVPFinalPaymentDate)) Then
				LogMessage "RSLT","Verification","Final Payment Date field value matched with the expected Value Expected: "&strVPFinalPaymentDate&" Actual: "&strIserveFinalPaymentDate&"",True
			Else
				LogMessage "RSLT","Verification","Final Payment Date field value doesnt match with the expected Value Expected: "&strVPFinalPaymentDate&" Actual: "&strIserveFinalPaymentDate&"",False	
			End IF		
		
			IF (Ucase(Trim(strIserveTotalNoOfInstallments)) =  Ucase(Trim(strVPTotalNoOfInstallments)))  OR (Ucase(Trim(strIserveTotalNoOfInstallments)) = 0.00  AND Ucase(Trim(strVPTotalNoOfInstallments))= 0 ) Then
				LogMessage "RSLT","Verification","Total Number of Installaments value matched with the expected Value Expected: "&strVPTotalNoOfInstallments&" Actual: "&strIserveTotalNoOfInstallments&"",True
			Else
				LogMessage "RSLT","Verification","Total Number of Installaments value doesnt match with the expected Value Expected: "&strVPTotalNoOfInstallments&" Actual: "&strIserveTotalNoOfInstallments&"",False	
			End IF	
			
			IF (Ucase(Trim(strIserveRemainingNoOfInstallments)) =  Ucase(Trim(strVPRemainingNoOfInstallments))) OR (Ucase(Trim(strIserveRemainingNoOfInstallments)) = 0.00 AND Ucase(Trim(strVPRemainingNoOfInstallments))= 0 ) Then
				LogMessage "RSLT","Verification","Remaining number of Installments value matched with the expected Value Expected: "&strVPRemainingNoOfInstallments&" Actual: "&strIserveRemainingNoOfInstallments&"",True
			Else
				LogMessage "RSLT","Verification","Remaining number of Installments value doesnt match with the expected Value Expected: "&strVPRemainingNoOfInstallments&" Actual: "&strIserveRemainingNoOfInstallments&"",False	
			End IF	
			
			IF (Ucase(Trim(strIserveAnnualPercentageCode)) =  Ucase(Trim(strVPAnnualPercentageCode))) OR (Ucase(Trim(strIserveRemainingNoOfInstallments)) = 0.00 AND  Ucase(Trim(strVPRemainingNoOfInstallments))= 0 ) Then
				LogMessage "RSLT","Verification","Annual Perecentage Code value matched with the expected Value Expected: "&strVPAnnualPercentageCode&" Actual: "&strIserveAnnualPercentageCode&"",True
			Else
				LogMessage "RSLT","Verification","Annual Perecentage Code value doesnt match with the expected Value Expected: "&strVPAnnualPercentageCode&" Actual: "&strIserveAnnualPercentageCode&"",False	
			End IF	
			
			IF (Ucase(Trim(strIserveTotalDisbursedAmount)) =  Ucase(Trim(strVPTotalDisbursedAmount))) OR (Ucase(Trim(strIserveTotalDisbursedAmount)) = 0.00 AND Ucase(Trim(strVPTotalDisbursedAmount))= 0 ) Then
				LogMessage "RSLT","Verification","Total Disbursed Amount field value matched with the expected Value Expected: "&strVPTotalDisbursedAmount&" Actual: "&strIserveTotalDisbursedAmount&"",True
			Else
				LogMessage "RSLT","Verification","Total Disbursed Amount field value doesnt match with the expected Value Expected: "&strVPTotalDisbursedAmount&" Actual: "&strIserveTotalDisbursedAmount&"",False	
			End IF			
	Next
End Function

'[Verify the Plan description and current Balance for all the records displayed in the Plan details table]
Public Function Verifyrecorddetails_OtherPlans(strCardNumber,strProduct)
	bNextPageExists = True
	j=1	
	Do While bNextPageExists = True
	intRecordCount = getRecordsCountForColumn(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent, "Plan No")		
		For i = 0 To intRecordCount- 1 
			Set objAllRows=getAllRows(OtherPlans.tblPlanSummaryContent)
			strIserveSeqNumber = getCellTextFor(OtherPlans.tblPlanSummaryHeader,objAllRows(i),i, "Seq. No.") 
			iPage = ((j*5) - (5-i))\3 ' Page is quotient (Page refers to Host system) 
			iRow = (((j-1)*5 + i) Mod 3) + 1   ' Row is remainder
			Call GetPagination_ARQA_Vplus(strCardNumber,strSeqNumber,strProduct,iPage,iRow)
				strVPStrPlan=Environment.Value("strPlan")
				strVPStrPlanDesc=Environment.Value("strPlanDesc")
				strVPStrCurrBalance=Environment.Value("strCurBalance")		
				
			'Verify the table contents on Plan details table 
			'lstlstPlandetailstable= (checknull("(Plan No:"&strVPStrPlan&"|Plan Description:"&strVPStrPlanDesc&"|Current Balance:"&strVPStrCurrBalance&")|")), " commented this code becoz plan descrption is blank. Need to get the pagination record with all the datas filled
			lstlstPlandetailstable= (checknull("(Plan No:"&strVPStrPlan&"|Current Balance:"&strVPStrCurrBalance&")|"))
			verifytablePlandetails_OtherPlan=verifyTableContentList(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,lstlstPlandetailstable,"Plan Details",false,null,null,null)		
		Next		
		 bNextPageExists  = False
		If intRecordCount = 5 Then
			bNextPageExists =matchStr(OtherPlans.lnkNext.GetROProperty("class"),"enabled")
			If bNextPageExists Then
				OtherPlans.lnkNext.Click
				j = j+1
			End If
		End If			
	Loop 	
End Function

'[Verify total records and pagination in Plan details table in OtherPlans Page]
Public Function ValidatePagination_PlanDetailstable()
 bValidatePagination_PlanDetailstable=true
 bNextPageExist = True
	While bNextPageExist = True
	 intRecordCount = getRecordsCountForColumn(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent, "Plan No")	
	 iCheck = 5 
		If intRecordCount <=iCheck  Then
		     LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to "&iCheck, true   
		     bValidatePagination_PlanDetailstable=true
			 If intRecordCount < iCheck Then
			   	bNextPageExist =matchStr(OtherPlans.lnkNext.GetROProperty("class"),"enabled")
				If bNextPageExist Then
				LogMessage "RSLT","Verification","Next link expected to be disabled if record is less than "&iCheck&". Currently it is enabled.",false
				bvalidatePagination=false
				Else
				LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
				End If
			ElseIf intRecordCount = iCheck Then
				bNextPageExist = matchStr(OtherPlans.lnkNext.GetROProperty("class"),"enabled")
				If bNextPageExist Then
					OtherPlans.lnkNext.Click
				End If
			End If
		Else 
			LogMessage "RSLT","Verification","Number of records displayed per page not matched with expected. Expected Count is less than or equal to 5", false   
			bNextPageExist = False
		End If
   Wend
End Function

'[Verify settlement Quote values by clicking early redemeption link from plan details table]
Public Function VerifySettlementQuote_forUL(strCardNumber)
   bVerifySettlementQuote_forUL=true   
   intRecordCount = getRecordsCountForColumn(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent, "Plan No")
  ' For i = 0 To intRecordCount - 1  
   	  	Set objAllRows=getAllRows(OtherPlans.tblPlanSummaryContent)
		strIserveSeqNumber = getCellTextFor(OtherPlans.tblPlanSummaryHeader,objAllRows(i),i, "Seq. No.")	
  	   	Call getOtherDetails_UL_ARVV_Vplus(strCardNumber,strIserveSeqNumber,i)   	   
	 	'get the values from the V+ 
		   strVPPlan = Environment.Value("strPlan") 
		   strVPPlanDesc = Environment.Value("strPlanDesc")
		   strVPCurBalance = Environment.Value("strCurBalance")
	   	   strVPSettlementType = Environment.Value("strSettlementType")  
	   	   
	   	   strVPInsuranceAmt = Environment.Value("StrInitialInsurance") 
		   strVPOriginalTerm = Environment.Value("StrInitialTerm")
		   strVPOutstandingTerm = Environment.Value("StrRemainingTerm")
		   
		   strVPPayOffDate1 = Environment.Value("StrPOByStartDate")
		   strVPPayOffDate2 = Environment.Value("StrPOByEndDate")
		   
		   strVPCurrentOutstandingAmtPO1 = Environment.Value("StrCurrentOustandingBal_POStart")
		   strVPCurrentOutstandingAmtPO2 = Environment.Value("StrCurrentOustandingBal_POEnd")
		   
		   strVPInterestRebatePO1 = Environment.Value("StrInterestRebate_POStart")
		   strVPInterestRebatePO2 = Environment.Value("StrInterestRebate_POEnd")  	   
			
			strVPInsuranceRebatePO1 = Environment.Value("StrInsuranceRebate_POStart") 
			strVPInsuranceRebatePO2 = Environment.Value("StrInsuranceRebate_POEnd") 
			
			strVPInterestPenaltyPO1 = Environment.Value("StrInterestPenalty_POStart")
			strVPInterestPenaltyPO2 = Environment.Value("StrInterestPenalty_POEnd")  
			
			strVPInsurancePenaltyPO1 = Environment.Value("StrInsurancePenalty_POStart")
			strVPInsurancePenaltyPO2 = Environment.Value("StrInsurancePenalty_POEnd")
			
			strVPTerminationFeePO1 = Environment.Value("StrTerminationFee_POStart")
			strVPTerminationFeePO2 = Environment.Value("StrTerminationFee_POEnd")
								
			strVPCashRebatePO1 = Environment.Value("StrCashRebate_POStart")
			strVPCashRebatePO2 = Environment.Value("StrCashRebate_POEnd")
					
			strVPProjectedInterestPO1 =	Environment.Value("StrProjectedInterest_POStart")
			strVPProjectedInterestPO2 =	Environment.Value("StrProjectedInterest_POEnd")
		
			strVPPenaltyInterestMonthPO1 = Environment.Value("StrPenaltyInterestMonth_POStart")
			strVPPenaltyInterestMonthPO2 = Environment.Value("StrPenaltyInterestMonth_POEnd")
		
			strVPWithoutPaymentDuePO1 = Environment.Value("StrWithoutPaymentDue_POStart")
			strVPWithoutPaymentDuePO2 = Environment.Value("StrWithoutPaymentDue_POEnd")	
								
			strVPNetPaymentDuePO1 = Environment.Value("StrNetPaymentDue_POStart")
			strVPNetPaymentDuePO2 =	Environment.Value("StrNetPaymentDue_POEnd")							
	   	   
			strVPStrPlanDesc = Environment.Value("StrCurrentOustandingBal_POStart")
			strVPStrCurrBalance = Environment.Value("StrInterestRebate_POStart")
			strVPSettlementType = Environment.Value("strSettlementType")
   	
   	   ' collate all the column values in to list 
   	   lstPlanSummary = (checknull("Plan No:"&strVPPlan&"|Plan Description:"&strVPPlanDesc&"|Current Balance:"&strVPCurBalance))
   	   bselectEarlyRedemption_ActionMenu = selectTableSubMenu(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,lstPlanSummary,"PlanSummary","Actions",False,NULL,NULL,NULL,"Early Redemption",bDisabled)
   	   If bDisabled Then
			LogMessage "RSLT", "Verification","Early Redemeption action menu is not enabled",false
			bVerifySettlementQuote_forUL=false	   	
   	   Else
   	   		LogMessage "RSLT","Verification","Early Redemption Action Menu link enabled and able to Click on the link.", True   
			'Get all the fields values below Plan Summary section in the Other Plan page to be displayed for LOAN and Cashline Products
			strIserveInsuranceAmount = OtherPlans.lblSQInsuranceAmount.GetROProperty("innertext")
			strIserveOriginalAmount = OtherPlans.lblSQOrgDeferAmount.GetROProperty("innertext")
			strIserveOutstandingAmount = OtherPlans.lblSQOustandingDeferAmount.GetROProperty("innertext")
			strIserveRequestQuoteType = OtherPlans.lblSQRequestQuoteType.GetROProperty("innertext")
			strIserveSettlementType = OtherPlans.lblSQSettlementType.GetROProperty("innertext")	
			
			If strIserveRequestQuoteType = "P" And strIserveSettlementType = "P"  Then
				LogMessage "RSLT","Verification","Request Quote Type and SettlementType is displayed as P Actual: "&strIserveRequestQuoteType&"",True
			Else
				LogMessage "RSLT","Verification","Request Quote Type and SettlementType is not displayed as P Actual: "&strIserveRequestQuoteType&"",False
			End IF 

			IF (Ucase(Trim(strIserveInsuranceAmount)) = 0 And Ucase(Trim(strVPInsuranceAmt)) = "") OR Ucase(Trim(strIserveInsuranceAmount)) = Ucase(Trim(strVPInsuranceAmt))  Then
				LogMessage "RSLT","Verification","Insurance Amount in Settlement Quote Popup is displayed as expected. Actual: "&strIserveInsuranceAmount&" Expected: "&strVPInsuranceAmt&"",True
			Else
				LogMessage "RSLT","Verification","Insurance Amount in Settlement Quote Popup is not displayed as expected. Actual: "&strIserveInsuranceAmount&" Expected: "&strVPInsuranceAmt&"",False	
			End IF 
			
			IF (Ucase(Trim(strIserveOriginalAmount)) = 0 And Ucase(Trim(strVPOriginalTerm)) = "") OR (Ucase(Trim(strIserveOriginalAmount)) = Ucase(Trim(strVPOriginalTerm)))Then
				LogMessage "RSLT","Verification","Original Deferement Amount in Settlement Quote Popup is displayed as expected. Actual: "&strIserveOriginalAmount&" Expected: "&strVPOriginalTerm&"",True
			Else
				LogMessage "RSLT","Verification","Original Deferement Amount in Settlement Quote Popup is not displayed as expected. Actual: "&strIserveOriginalAmount&" Expected: "&strVPOriginalTerm&"",False	
			End IF 

			IF (Ucase(Trim(strIserveOutstandingAmount)) = 0  AND Ucase(Trim(strVPOutstandingTerm)) = "") OR (Ucase(Trim(strIserveOutstandingAmount)) = Ucase(Trim(strVPOutstandingTerm))) Then
				LogMessage "RSLT","Verification","Insurance Amount in Settlement Quote Popup is displayed as expected. Actual: "&strIserveOutstandingAmount&" Expected: "&strVPOutstandingTerm&"",True
			Else
				LogMessage "RSLT","Verification","Insurance Amount in Settlement Quote Popup is not displayed as expected. Actual: "&strIserveOutstandingAmount&" Expected: "&strVPOutstandingTerm&"",False	
			End IF 
			
		'Verify the table contents on Plan details table
		lstlstSettlementType= (checknull("(:Current Outstanding Balance|"&strVPPayOffDate1&":"&strVPCurrentOutstandingAmtPO1&"|"&strVPPayOffDate2&":"&strVPCurrentOutstandingAmtPO2&")|(:Interest Rebate|"&strVPPayOffDate1&":"&strVPInterestRebatePO1&"|"&strVPPayOffDate2&":"&strVPInterestRebatePO2&")|(:Insurance Rebate|"&strVPPayOffDate1&":"&strVPInsuranceRebatePO1&"|"&strVPPayOffDate2&":"&strVPInsuranceRebatePO2&")|(:Interest Penality|"&strVPPayOffDate1&":"&strVPInterestPenaltyPO1&"|"&strVPPayOffDate2&":"&strVPInterestPenaltyPO2&")|(:Insurance Penality|"&strVPPayOffDate1&":"&strVPInsurancePenaltyPO1&"|"&strVPPayOffDate2&":"&strVPInsurancePenaltyPO2&")|(:Termination Fee|"&strVPPayOffDate1&":"&strVPTerminationFeePO1&"|"&strVPPayOffDate2&":"&strVPTerminationFeePO2&")|(:ClawBack Of Cash Rebate|"&strVPPayOffDate1&":"&strVPCashRebatePO1&"|"&strVPPayOffDate2&":"&strVPCashRebatePO1&")|(:Projected Interest|"&strVPPayOffDate1&":"&strVPProjectedInterestPO1&"|"&strVPPayOffDate2&":"&strVPProjectedInterestPO2&")|(:Penality Interest Months|"&strVPPayOffDate1&":"&strVPPenaltyInterestMonthPO1&"|"&strVPPayOffDate2&":"&strVPPenaltyInterestMonthPO2&")|(:Without Payment Due|"&strVPPayOffDate1&":"&strVPWithoutPaymentDuePO1&"|"&strVPPayOffDate2&":"&strVPWithoutPaymentDuePO2&")|(:Net Payment Due|"&strVPPayOffDate1&":"&strVPNetPaymentDuePO1&"|"&strVPPayOffDate2&":"&strVPNetPaymentDuePO2&")|"))
		verifytablePlandetails_OtherPlan=verifyTableContentList(OtherPlans.tblSettlementQuoteHeader,OtherPlans.tblSettlementQuoteContent,lstlstSettlementType,"SettlementQuote",false,null,null,null)			   
	  End If    
 ' Next 
    VerifySettlementQuote_forUL=bVerifySettlementQuote_forUL
End Function

'[LISA Verify the plan details for Other Plans enquiry credit card]
Public Function verifyPlanDetailsCC_lisa(strCardnumber,strProduct,strType,strBaseRate,strCalculatedRate,strInterestStartDate,strLTDPrincipal,strLTDInterest,strYTDPrincipal,strYTDInterest,strServiceCharges,strOpenDate,strBalTransferMonthsRemaining,strBalTransferExpiryDate,strAccuredInterest,strPerDiem,strNormalInterestBeginDate)
	 		strIserveBaseRate = OtherPlans.lblBaseRate.GetROProperty("innertext")
			strIserveBaseRate = Replace(strIserveBaseRate,"%","")
			strIserveCalculatedRate = OtherPlans.lblCalulatedRate.GetROProperty("innertext")
			strIserveCalculatedRate = Replace(strIserveCalculatedRate,"%","")
			strIserveInterestStartDate = OtherPlans.lblInterestStartDate.GetROProperty("innertext")			
			strIserveLTDPrincipal = OtherPlans.lblLifeToDatePrincipal.GetROProperty("innertext")
			strIserveLTDInterest = OtherPlans.lblLifeToDateInterest.GetROProperty("innertext")			
			strIserveYTDPrincipal = OtherPlans.lblYeartoDatePrincipal.GetROProperty("innertext")
			strIserveYTDInterest = OtherPlans.lblYearToDateInterest.GetROProperty("innertext")
			strIserveServiceCharges = OtherPlans.lblServiceCharges.GetROProperty("innertext")
			strIserveOpenDate = OtherPlans.lblOpenDate.GetROProperty("innertext")
			strIserveBalTransferMonthsRemaining = OtherPlans.lblBalanceTransferMonthsRemain.GetROProperty("innertext")		
			strIserveBalTransferExpiryDate = OtherPlans.lblBalanceTransferExpiry.GetROProperty("innertext")		
			strIserveAccuredInterest = OtherPlans.lblAccruedInterest.GetROProperty("innertext")
			strIservePerDiem = OtherPlans.lblPerDiem.GetROProperty("innertext")		
			strIserveNormalInterestBeginDate = OtherPlans.lblNormalInterestBeginDate.GetROProperty("innertext")			
   bDevPending=false
   bverifyPlanDetailsCC_lisa=true
   
  ' If strBaseRate <> "" Then
   	If strBaseRate = strIserveBaseRate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Base rate is as expected. Actual: "&strBaseRate&" Expected: "&strIserveBaseRate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Base rate is not as expected. Actual:"&strBaseRate&" Expected:"&strIserveBaseRate&"", False
	End If
   'End If
   
   'If strCalculatedRate <> "" Then
   	If strCalculatedRate = strIserveCalculatedRate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Calculated rate is as expected. Actual: "&strCalculatedRate&" Expected: "&strIserveCalculatedRate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Calculated rate is not as expected. Actual: "&strCalculatedRate&" Expected: "&strIserveCalculatedRate&"", False
	End If
   'End If
  
  ' If strInterestStartDate <> "" Then
   	If strInterestStartDate = strIserveInterestStartDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Interest start date is as expected. Actual: "&strInterestStartDate&" Expected: "&strIserveInterestStartDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Interest start date is not as expected. Actual: "&strInterestStartDate&" Expected: "&strIserveInterestStartDate&"", False
	End If
   'End If
   
   'If strLTDPrincipal <> "" Then
   	If strLTDPrincipal = strIserveLTDPrincipal Then
	  	 LogMessage "RSLT", "Verification", "Iserve Life to date principal is as expected. Actual: "&strLTDPrincipal&" Expected: "&strIserveLTDPrincipal&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Life to date principal  is not as expected. Actual: "&strLTDPrincipal&" Expected: "&strIserveLTDPrincipal&"", False
	End If
   'End If
 
   'If  strLTDInterest <> "" Then
   	If  strLTDInterest = strIserveLTDInterest Then
	  	 LogMessage "RSLT", "Verification", "Iserve Life to date interest is as expected. Actual: "&strLTDInterest&" Expected: "&strIserveLTDInterest&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Life to date interest  is not as expected. "&strLTDInterest&" Expected: "&strIserveLTDInterest&"", False
	End If
   'End If
   
   'If  strYTDPrincipal <> "" Then
   	If  strYTDPrincipal = strIserveYTDPrincipal Then
	  	 LogMessage "RSLT", "Verification", "Iserve year to date principal is as expected. Actual: "&strYTDPrincipal&" Expected: "&strIserveYTDPrincipal&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve year to date principal  is not as expected. Actual: "&strYTDPrincipal&" Expected: "&strIserveYTDPrincipal&"", False
	End If
   'End If
   
   'If  strYTDInterest <> "" Then
   	If  strYTDInterest = strIserveYTDInterest Then
	  	 LogMessage "RSLT", "Verification", "Iserve year to date interest is as expected. Actual: "&strYTDInterest&" Expected: "&strIserveYTDInterest&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve year to date interest  is not as expected. Actual: "&strYTDInterest&" Expected: "&strIserveYTDInterest&"", False
	End If
   	'End If
   	
   	'If  strServiceCharges <> "" Then
   	If  strServiceCharges = strIserveServiceCharges Then
	  	 LogMessage "RSLT", "Verification", "Iserve Service Charges is as expected. Actual: "&strServiceCharges&" Expected: "&strIserveServiceCharges&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Service Charges is not as expected. Actual: "&strServiceCharges&" Expected: "&strIserveServiceCharges&"", False
	End If
   	'End If
   	
   	'If  strOpenDate <> "" Then
   	If  strOpenDate = strIserveOpenDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Open Date is as expected. Actual: "&strOpenDate&" Expected: "&strIserveOpenDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Open Date is not as expected. Actual: "&strOpenDate&" Expected: "&strIserveOpenDate&"", False
	End If
   	'End If
   	
   	'If  strBalTransferMonthsRemaining <> "" Then
   	If  strBalTransferMonthsRemaining = strIserveBalTransferMonthsRemaining Then
	  	 LogMessage "RSLT", "Verification", "Iserve Balance transfer month remaining is as expected. Actual: "&strBalTransferMonthsRemaining&" Expected: "&strIserveBalTransferMonthsRemaining&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Balance transfer month remaining is not as expected. Actual: "&strBalTransferMonthsRemaining&" Expected: "&strIserveBalTransferMonthsRemaining&"", False
	End If
   	'End If
   	
   	'If  strBalTransferExpiryDate <> "" Then
   	If  strBalTransferExpiryDate = strIserveBalTransferExpiryDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Balance transfer expiry date is as expected. Actual: "&strBalTransferExpiryDate&" Expected: "&strIserveBalTransferExpiryDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Balance transfer expiry date is not as expected. Actual: "&strBalTransferExpiryDate&" Expected: "&strIserveBalTransferExpiryDate&"", False
	End If
   	'End If
   	
   '	If  strAccuredInterest <> "" Then
   	If  strAccuredInterest = strIserveAccuredInterest Then
	  	 LogMessage "RSLT", "Verification", "Iserve Accured Interest is as expected. Actual: "&strAccuredInterest&" Expected: "&strIserveAccuredInterest&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Accured Interest is not as expected. Actual: "&strAccuredInterest&" Expected: "&strIserveAccuredInterest&"", False
	End If
   '	End If
   
 '  If  strPerDiem <> "" Then
   	If  strPerDiem = strIservePerDiem Then
	  	 LogMessage "RSLT", "Verification", "Iserve per diem is as expected. Actual: "&strPerDiem&" Expected: "&strIservePerDiem&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve per diem is not as expected. Actual: "&strPerDiem&" Expected: "&strIservePerDiem&"", False
	End If
  ' 	End If
   	
  ' 	If  strNormalInterestBeginDate <> "" Then
   	If  strNormalInterestBeginDate = strIserveNormalInterestBeginDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve normal interest begin date is as expected. Actual: "&strNormalInterestBeginDate&" Expected: "&strIserveNormalInterestBeginDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve normal interest begin date is not as expected. Actual: "&strNormalInterestBeginDate&" Expected: "&strIserveNormalInterestBeginDate&"", False
	End If
  ' 	End If
  verifyPlanDetailsCC_lisa = bverifyPlanDetailsCC_lisa
End Function

'[LISA Verify the plan details for Other Plans enquiry Loan]
Public Function verifyPlanDetailsLoan_lisa(strCardnumber,strProduct,strCardtype,strBaseRate,strCalculatedRate,strInterestStartDate,strLTDPrincipal,strLTDInterest,strYTDPrincipal,strYTDInterest,strServiceCharges,strOpenDate,strBalTransferMonthsRemaining,strBalTransferExpiryDate,strAccuredInterest,strPerDiem,strNormalInterestBeginDate,strInitialTerm,strCurrentTerm,strLoanAmount,strPrincipalAmount,strInterestAmount,strFirstPaymentAmount,strFirstPaymentDate,strFinalPaymentDate,strTotalNoOfInstallments,strRemainingNoOfInstallments,strAnnualPercentageCode,strTotalDisbursedAmount)
	
			strIserveBaseRate = OtherPlans.lblBaseRate.GetROProperty("innertext")
			strIserveCalculatedRate = OtherPlans.lblCalulatedRate.GetROProperty("innertext")
			strIserveCalculatedRate = Replace(strIserveCalculatedRate,"%","")
			strIserveInterestStartDate = OtherPlans.lblInterestStartDate.GetROProperty("innertext")
			strIserveLTDPrincipal = OtherPlans.lblLifeToDatePrincipal.GetROProperty("innertext")
			strIserveLTDInterest = OtherPlans.lblLifeToDateInterest.GetROProperty("innertext")			
			strIserveYTDPrincipal = OtherPlans.lblYeartoDatePrincipal.GetROProperty("innertext")
			strIserveYTDInterest = OtherPlans.lblYearToDateInterest.GetROProperty("innertext")			
			strIserveServiceCharges = OtherPlans.lblServiceCharges.GetROProperty("innertext")
			strIserveOpenDate = OtherPlans.lblOpenDate.GetROProperty("innertext")
			strIserveBalTransferMonthsRemaining = OtherPlans.lblBalanceTransferMonthsRemain.GetROProperty("innertext")		
			strIserveBalTransferExpiryDate = OtherPlans.lblBalanceTransferExpiry.GetROProperty("innertext")		
			strIserveAccuredInterest = OtherPlans.lblAccruedInterest.GetROProperty("innertext")
			strIservePerDiem = OtherPlans.lblPerDiem.GetROProperty("innertext")		
			strIserveNormalInterestBeginDate = OtherPlans.lblNormalInterestBeginDate.GetROProperty("innertext")	
			' LOAN Information sections 
			strIserveInitialTerm = OtherPlans.lblInitialTerm.GetROProperty("innertext")
			strIserveCurrentTerm = OtherPlans.lblCurrentTerm.GetROProperty("innertext")
			strIserveLoanAmount = OtherPlans.lblLoanAmount.GetROProperty("innertext")		
			strIservePrincipalAmount = OtherPlans.lblPrincipalAmount.GetROProperty("innertext")		
			strIserveInterestAmount = OtherPlans.lblInterestAmount.GetROProperty("innertext")
			strIserveFirstPaymentAmount = OtherPlans.lblFirstPaymentAmount.GetROProperty("innertext")		
			strIserveFirstPaymentDate= OtherPlans.lblFirstPaymentDate.GetROProperty("innertext")				
			strIserveFinalPaymentDate= OtherPlans.lblFinalPayementDate.GetROProperty("innertext")				
			strIserveTotalNoOfInstallments= OtherPlans.lblTotalInstallment.GetROProperty("innertext")				
			strIserveRemainingNoOfInstallments= OtherPlans.lblRemainingInstallment.GetROProperty("innertext")				
			strIserveAnnualPercentageCode= OtherPlans.lblAnnualPercentageCode.GetROProperty("innertext")
			strIserveAnnualPercentageCode = Replace(strIserveAnnualPercentageCode,"%","")			
			strIserveTotalDisbursedAmount= OtherPlans.lblTotalDisbursedAmount.GetROProperty("innertext")
	
   bDevPending=false
   bverifyPlanDetailsLoan_lisa=true
   
    If strBaseRate <> "" Then
   		LogMessage "RSLT", "Verification", "Base Rate field is popualated for Loan and cashline Products as expected. Actual: "&strBaseRate&" Expected: "&strIserveBaseRate&"", false
	 else
		LogMessage "RSLT", "Verification", "Base Rate field displayed Null for LOAN or Cashline Products as Expected. Actual: "&strBaseRate&" Expected: "&strIserveBaseRate&"", True
   End if 
   
  ' If strCalculatedRate <> "" Then
   	If strCalculatedRate = strIserveCalculatedRate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Calculated rate is as expected. Actual: "&strCalculatedRate&" Expected: "&strIserveCalculatedRate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Calculated rate is not as expected. Actual: "&strCalculatedRate&" Expected: "&strIserveCalculatedRate&"", False
	End If
   'End If
  
  ' If strInterestStartDate <> "" Then
   	If strInterestStartDate = strIserveInterestStartDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Interest start date is as expected. Actual: "&strInterestStartDate&" Expected: "&strIserveInterestStartDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Interest start date is not as expected. Actual: "&strInterestStartDate&" Expected: "&strIserveInterestStartDate&"", False
	End If
 '  End If
   
  ' If strLTDPrincipal <> "" Then
   	If strLTDPrincipal = strIserveLTDPrincipal Then
	  	 LogMessage "RSLT", "Verification", "Iserve Life to date principal is as expected. Actual: "&strLTDPrincipal&" Expected: "&strIserveLTDPrincipal&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Life to date principal  is not as expected. Actual: "&strLTDPrincipal&" Expected: "&strIserveLTDPrincipal&"", False
	End If
  ' End If
 
  ' If  strLTDInterest <> "" Then
   	If  strLTDInterest = strIserveLTDInterest Then
	  	 LogMessage "RSLT", "Verification", "Iserve Life to date interest is as expected. Actual: "&strLTDInterest&" Expected: "&strIserveLTDInterest&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Life to date interest  is not as expected. Actual: "&strLTDInterest&" Expected: "&strIserveLTDInterest&"", False
	End If
  ' End If
   
  ' If  strYTDPrincipal <> "" Then
   	If  strYTDPrincipal = strIserveYTDPrincipal Then
	  	 LogMessage "RSLT", "Verification", "Iserve year to date principal is as expected. Actual: "&strYTDPrincipal&" Expected: "&strIserveYTDPrincipal&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve year to date principal  is not as expected. Actual: "&strYTDPrincipal&" Expected: "&strIserveYTDPrincipal&"", False
	End If
   'End If
   
  ' If  strYTDInterest <> "" Then
   	If  strYTDInterest = strIserveYTDInterest Then
	  	 LogMessage "RSLT", "Verification", "Iserve year to date interest is as expected. Actual: "&strYTDInterest&" Expected: "&strIserveYTDInterest&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve year to date interest  is not as expected. Actual: "&strYTDInterest&" Expected: "&strIserveYTDInterest&"", False
	End If
   '	End If
   	
   	If  strServiceCharges <> "" Then
   	 LogMessage "RSLT","Verification","Service Charges are displayed for LOAN and cashline Products.",False
	else
	  	LogMessage "RSLT","Verification","Service Charges are displayed Null for LOAN and cashline Products as expected.",True
   	End If
   	
 '  	If  strOpenDate <> "" Then
   	If  strOpenDate = strIserveOpenDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Open Date is as expected. Actual: "&strOpenDate&" Expected: "&strIserveOpenDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Open Date is not as expected. Actual: "&strOpenDate&" Expected: "&strIserveOpenDate&"", False
	End If
 '  	End If
   	
  ' 	If  strBalTransferMonthsRemaining <> "" Then
   	If  strBalTransferMonthsRemaining = strIserveBalTransferMonthsRemaining Then
	  	 LogMessage "RSLT", "Verification", "Iserve Balance transfer month remaining is as expected. Actual: "&strBalTransferMonthsRemaining&" Expected: "&strIserveBalTransferMonthsRemaining&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Balance transfer month remaining is not as expected. Actual: "&strBalTransferMonthsRemaining&" Expected: "&strIserveBalTransferMonthsRemaining&"", False
	End If
 '  	End If
   	
  ' 	If  strBalTransferExpiryDate <> "" Then
   	If  strBalTransferExpiryDate = strIserveBalTransferExpiryDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve Balance transfer expiry date is as expected. Actual: "&strBalTransferExpiryDate&" Expected: "&strIserveBalTransferExpiryDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Balance transfer expiry date is not as expected. Actual: "&strBalTransferExpiryDate&" Expected: "&strIserveBalTransferExpiryDate&"", False
	End If
  ' 	End If
   	
  ' 	If  strAccuredInterest <> "" Then
   	If  strAccuredInterest = strIserveAccuredInterest Then
	  	 LogMessage "RSLT", "Verification", "Iserve Accured Interest is as expected. Actual: "&strAccuredInterest&" Expected: "&strIserveAccuredInterest&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Accured Interest is not as expected. Actual: "&strAccuredInterest&" Expected: "&strIserveAccuredInterest&"", False
	End If
  ' 	End If
   
 '  If  strPerDiem <> "" Then
   	If  strPerDiem = strIservePerDiem Then
	  	 LogMessage "RSLT", "Verification", "Iserve per diem is as expected. Actual: "&strPerDiem&" Expected: "&strIservePerDiem&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve per diem is not as expected. Actual: "&strPerDiem&" Expected: "&strIservePerDiem&"", False
	End If
 '  	End If
   	
  ' 	If  strNormalInterestBeginDate <> "" Then
   	If  strNormalInterestBeginDate = strIserveNormalInterestBeginDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve normal interest begin date is as expected. Actual: "&strNormalInterestBeginDate&" Expected: "&strIserveNormalInterestBeginDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve normal interest begin date is not as expected. Actual: "&strNormalInterestBeginDate&" Expected: "&strIserveNormalInterestBeginDate&"", False
	End If
 '  	End If
   	
  ' If  strInitialTerm <> "" Then
   	If  strInitialTerm = strIserveInitialTerm Then
	  	 LogMessage "RSLT", "Verification", "Iserve Initial term is as expected. Actual: "&strInitialTerm&" Expected: "&strIserveInitialTerm&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Initial term is not as expected. Actual: "&strInitialTerm&" Expected: "&strIserveInitialTerm&"", False
	End If
 '  	End If
   
   'If  strCurrentTerm <> "" Then
   	If  strCurrentTerm = strIserveCurrentTerm Then
	  	 LogMessage "RSLT", "Verification", "Iserve current term is as expected. Actual: "&strCurrentTerm&" Expected: "&strIserveCurrentTerm&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve current term is not as expected. Actual: "&strCurrentTerm&" Expected: "&strIserveCurrentTerm&"", False
	End If
   	'End If
   	
 '  	If  strLoanAmount <> "" Then
   	If  strLoanAmount = strIserveLoanAmount Then
	  	 LogMessage "RSLT", "Verification", "Iserve Loan Amount is as expected. Actual: "&strLoanAmount&" Expected: "&strIserveLoanAmount&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Loan Amount is not as expected. Actual: "&strLoanAmount&" Expected: "&strIserveLoanAmount&"", False
	End If
  ' 	End If
   	
  '		If  strPrincipalAmount <> "" Then
   	If  strPrincipalAmount = strIservePrincipalAmount Then
	  	 LogMessage "RSLT", "Verification", "Iserve Principal Amount is as expected. Actual: "&strPrincipalAmount&" Expected: "&strIservePrincipalAmount&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Principal Amount is not as expected. Actual: "&strPrincipalAmount&" Expected: "&strIservePrincipalAmount&"", False
	End If
 '  	End If
   	
   '	If  strInterestAmount <> "" Then
   	If  strInterestAmount = strIserveInterestAmount Then
	  	 LogMessage "RSLT", "Verification", "Iserve Interest Amount is as expected. Actual: "&strInterestAmount&" Expected: "&strIserveInterestAmount&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Interest Amount is not as expected. Actual: "&strInterestAmount&" Expected: "&strIserveInterestAmount&"", False
	End If
  ' 	End If
   	
  ' 	If  strFirstPaymentAmount <> "" Then
   	If  strFirstPaymentAmount = strIserveFirstPaymentAmount Then
	  	 LogMessage "RSLT", "Verification", "Iserve First Payment Amount is as expected. Actual: "&strFirstPaymentAmount&" Expected: "&strIserveFirstPaymentAmount&"", True
	else
	  	LogMessage "RSLT", "Verification", "IserveFirst Payment Amount is not as expected. Actual: "&strFirstPaymentAmount&" Expected: "&strIserveFirstPaymentAmount&"", False
	End If
'   	End If
   	
  ' 	If  strFirstPaymentDate <> "" Then
   	If  strFirstPaymentDate = strIserveFirstPaymentDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve first payment date is as expected. Actual: "&strFirstPaymentDate&" Expected: "&strIserveFirstPaymentDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve first payment date is not as expected. Actual: "&strFirstPaymentDate&" Expected: "&strIserveFirstPaymentDate&"", False
	End If
 '  	End If
   
 '  	If  strFirstPaymentDate <> "" Then
   	If  strFirstPaymentDate = strIserveFirstPaymentDate Then
	  	 LogMessage "RSLT", "Verification", "Iserve final payment date is as expected. Actual: "&strFirstPaymentDate&" Expected: "&strIserveFirstPaymentDate&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve final payment date is not as expected. Actual: "&strFirstPaymentDate&" Expected: "&strIserveFirstPaymentDate&"", False
	End If
  ' 	End If
   	
 '  	If  strTotalNoOfInstallments <> "" Then
   	If  strTotalNoOfInstallments = strIserveTotalNoOfInstallments Then
	  	 LogMessage "RSLT", "Verification", "Iserve Total no. of installment is as expected. Actual: "&strTotalNoOfInstallments&" Expected: "&strIserveTotalNoOfInstallments&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve Total no. of installment is not as expected. Actual: "&strTotalNoOfInstallments&" Expected: "&strIserveTotalNoOfInstallments&"", False
	End If
  ' 	End If
   	
  ' 	If  stremainingNoOfInstallments <> "" Then
   	If  strRemainingNoOfInstallments = strIserveRemainingNoOfInstallments Then
	  	 LogMessage "RSLT", "Verification", "Iserve remaining no. of installment is as expected. Actual: "&stremainingNoOfInstallments&" Expected: "&strIserveRemainingNoOfInstallments&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve remaining no. of installment is not as expected. Actual: "&stremainingNoOfInstallments&" Expected: "&strIserveRemainingNoOfInstallments&"", False
	End If
 '  	End If
   	
 '  	If  strAnnualPercentageCode <> "" Then
   	If  strAnnualPercentageCode = strIserveAnnualPercentageCode Then
	  	 LogMessage "RSLT", "Verification", "Iserve annual percentage code is as expected. Actual: "&strAnnualPercentageCode&" Expected: "&strIserveAnnualPercentageCode&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve annual percentage code is not as expected. Actual: "&strAnnualPercentageCode&" Expected: "&strIserveAnnualPercentageCode&"", False
	End If
  ' 	End If
   	
  ' 	If  strTotalDisbursedAmount <> "" Then
   	If  strTotalDisbursedAmount = strIserveTotalDisbursedAmount Then
	  	 LogMessage "RSLT", "Verification", "Iserve total disburse amount is as expected. Actual: "&strTotalDisbursedAmount&" Expected: "&strIserveTotalDisbursedAmount&"", True
	else
	  	LogMessage "RSLT", "Verification", "Iserve total disbursed amount is not as expected. Actual: "&strTotalDisbursedAmount&" Expected: "&strIserveTotalDisbursedAmount&"", False
	End If
 '  	End If
   	verifyPlanDetailsLoan_lisa = bverifyPlanDetailsLoan_lisa
End Function

'[LISA Verify the plan summary table]
Public Function verifyplansummarytablecontent(arrowDataList)
	bDevPending=false
	bverifyplansummarytablecontent = true
	verifyplansummarytable = verifyTableContentList(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,arrowDataList,"Summary" ,  false,null ,null,null)
	'If verifyplansummarytable = true Then
	'	clickPlanNumber=selectTableLink(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,arrowDataList,"Summary","Plan No",false,null,null,null)	
	'End If
	verifyplansummarytablecontent=bverifyplansummarytablecontent
End Function

'[LISA Verify Table Deferral information has following Columns]
Public Function verifyDeferralInfoTableColumns(arrColumnNameList)
   bDevPending=false
   verifyDeferralInfoTableColumns=verifyTableColumns(OtherPlans.tblDefferalInformationContent,arrColumnNameList)
End Function

'[LISA Verify row Data in Table Deferral Information]
Public Function verifyDeferralInfo_rowdata(arrRowDataList)
   bDevPending=false
   bverifyDeferralInfo_rowdata= true
  verifyDeferralInfo_rowdata =verifyTableContentList(OtherPlans.tblDefferalInformationHeader,OtherPlans.tblDefferalInformationContent,arrRowDataList,"Deferral info" ,  false,null ,null,null)
  verifyDeferralInfo_rowdata =bverifyDeferralInfo_rowdata
End Function

'[LISA Click on the Redemption link of Other Plans record]
Public Function selectRedemptionAction_OtherPlans(lstPlanDetails)
   bDevPending=False
   bselectRedemptionAction_OtherPlans=true
 '	With OtherPlans
		  bselectRedemptionAction_OtherPlans= selectTableSubMenu(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,lstPlanDetails,"PlanSummary","Actions",False,NULL,NULL,NULL,"Early Redemption",bDisabled)
'	End With
	If bDisabled Then
		LogMessage "RSLT", "Verification","Redemption action menu is not enabled",false
		bselectRedemptionAction_OtherPlans=false
	End If
	WaitForICallLoading
	Wait 1
    selectRedemptionAction_OtherPlans=bselectRedemptionAction_OtherPlans
End Function

'[LISA Verify Settlement Quote pop up exist]
Public Function verifyPopupSettlementQuoteexist(bExist)
   bDevPending=false
   bActualExist=OtherPlans.popupSettlementQuote.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Settlement Quoate Exists As Expected" ,true
       verifyPopupSettlementQuoteexist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Settlement Quoate does not Exists As Expected" ,true
       verifyPopupSettlementQuoteexist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Settlement Quoate does not Exists As Expected" ,False
       verifyPopupSettlementQuoteexist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Settlement Quoate Still Exists" ,False
       verifyPopupSettlementQuoteexist=False
   End If
End Function

'[LISA Verify the fields in Settlement Quoate for Loans]
Public Function verifySettlementQuoteDetails_lisa(strInsAmount,strOrgAmount,strOutStndAmount,strReqQuoType,strSettType)

	'Get all the fields values below Plan Summary section in the Other Plan page to be displayed for LOAN and Cashline Products
			strIserveInsuranceAmount = OtherPlans.lblSQInsuranceAmount.GetROProperty("innertext")
			strIserveOriginalAmount = OtherPlans.lblSQOrgDeferAmount.GetROProperty("innertext")
			strIserveOutstandingAmount = OtherPlans.lblSQOustandingDeferAmount.GetROProperty("innertext")
			strIserveRequestQuoteType = OtherPlans.lblSQRequestQuoteType.GetROProperty("innertext")
			strIserveSettlementType = OtherPlans.lblSQSettlementType.GetROProperty("innertext")
			
	bDevPending=false
   bverifySettlementQuoteDetails_lisa=true		
			
	If strReqQuoType = "P" and strIserveRequestQuoteType ="P" Then
		LogMessage "RSLT","Verification","Request Quote Type is displayed as P Actual: "&strReqQuoType&"",True
		Else 
		LogMessage "RSLT","Verification","Request Quote Type is not displayed as P Actual: "&strReqQuoType&"",False
	End If
	
	If strSettType= "P" and strIserveSettlementType ="P" Then
		LogMessage "RSLT","Verification","Settlement Quote Type is displayed as P Actual: "&strReqQuoType&"",True
		Else 
		LogMessage "RSLT","Verification","Settlement Quote Type is not displayed as P Actual: "&strReqQuoType&"",False
	End If

	'	If strInsAmount <> "" Then
   			If strInsAmount = strIserveInsuranceAmount Then
   	
	  	LogMessage "RSLT","Verification","The Iserve insurance amount is as expected: "&strInsAmount&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve insurance amount is not as expected: "&strInsAmount&"",False
		End if 
 '  End If
   
 '  If strOrgAmount <> "" Then
   	If strOrgAmount = strIserveOriginalAmount Then
   	
	  	LogMessage "RSLT","Verification","The Iserve original amount is as expected: "&strOrgAmount&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve original amount is not as expected: "&strOrgAmount&"",False
		End if 
  ' End If
   
  '  If strOutStndAmount <> "" Then
   	If strOutStndAmount = strIserveOutstandingAmount Then
   	
	  	LogMessage "RSLT","Verification","The Iserve outstanding amount is as expected: "&strOutStndAmount&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve outstanding amount is not as expected: "&strOutStndAmount&"",False
		End if 
  ' End If
   verifySettlementQuoteDetails_lisa = bverifySettlementQuoteDetails_lisa
End Function

'[LISA Click on plan from plan summmary table]
Public Function clickPlanNumber(arrowDataList)
	bDevPending=false
   	bclickPlanNumber=true
   	clickPlanNumber=selectTableLink(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,arrowDataList,"Summary","Plan No",false,null,null,null)
		clickPlanNumber = bclickPlanNumber
End Function

'[Lisa Verify row data in Table Settlement Quote]
Public Function verifySettlmentQuote_rowdata(arrRowDataList)
   bDevPending=false
   bverifySettlmentQuote_rowdata= true
  verifySettlmentQuote_rowdata =verifyTableContentList(OtherPlans.tblSettlementQuoteHeader,OtherPlans.tblSettlementQuoteContent,arrRowDataList,"Settlement Quote",false,null ,null,null)
  verifySettlmentQuote_rowdata =bverifySettlmentQuote_rowdata
End Function

'[LISA Click Button OK for Settlement Quote]
Public Function clickButtonOK_SettlementQuote()
   bDevPending=true
   OtherPlans.btnOK.click
   If Err.Number<>0 Then
       clickButtonOK_SettlementQuote=false
            LogMessage "RSLT","Verification","Failed to Click Button : OK" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonOK_SettlementQuote=true
End Function
