'[Click On Account number in customer overview page]
Public Function clickCurrentAccountNumber(strAccountDetails,strAccountType)
	Dim bAccountDetails:bAccountDetails = True
	bAccountDetails = selectTableLink(HK_CCTR_TDKeyInfo_Page.tblTDOverviewHeader(),HK_CCTR_TDKeyInfo_Page.tblTDOverviewContent(),strAccountDetails,strAccountType ,"Accounts",False,NULL,NULL,NULL)
	WaitForICallLoading
	
	If strAccountType = "Time Deposit Account" Then
		If HK_CCTR_ClosedAccounts_Page.elePlacementInfo().Exist(3) Then
			LogMessage "RSLT","Verification","Placement Info page displayed by default on clicking the account number from overview screen.",True
		Else
			LogMessage "WARN","Verification","Failed to display Placement Info page by default on clicking the account number from overview screen.",False
		End If
	Else
		If HK_CCTR_BalanceAndLimits_Page.eleAccountBalance().Exist(3) Then
			LogMessage "RSLT","Verification","Balances and Limits page displayed by default on clicking the account number from overview screen.",True
		Else
			LogMessage "WARN","Verification","Failed to display Balances and Limits page by default on clicking the account number from overview screen.",False
		End If
	
	End If
	
	clickCurrentAccountNumber = bAccountDetails
End Function

'[Verify Currency dropdown and Go button]
Public Function verifyCurrencyDropDownAndGoButton(strlstMultiCurrencies)
	VerifyFieldExistenceInPage HK_CCTR_BalanceAndLimits_Page.lblCurrency(),"Balance & Limits","Currency label"
	VerifyFieldExistenceInPage HK_CCTR_BalanceAndLimits_Page.eleCurrencyDropdown(),"Balance & Limits","Currency Drop Down"
	verifyComboboxItems HK_CCTR_BalanceAndLimits_Page.eleCurrencyDropdown(),strlstMultiCurrencies,"Multi Currency Dropdown"
	VerifyFieldExistenceInPage HK_CCTR_BalanceAndLimits_Page.btnGoButton(),"Balance & Limits","Go Button"
End Function

'[Verify Pink Panel Details of Left Menu link]
Public Function verifyPinkPanelDetailsOfLeftMenuLink(strLeftMenuLinkName,strProduct,strcardNumber,strStatus,strAccCardInd,strStaffIndc,strSubProduct,strOpenDate)
	bverifyPinkPanelDetailsOfLeftMenuLink = true
	
	If HK_CCTR_TDKeyInfo_Page.lblTDProduct().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Product Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Product Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	
	If HK_CCTR_TDKeyInfo_Page.lblTDAccountNumber().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Account Number Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Account Number Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDStatus().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Status Label is displayed as expected in " & strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Status Label in " &strLeftMenuLinkName &" Page",false
		bverifyLeftMenuLinksPinkPanelDetails=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDAccCardInd().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Acct/Card Ind Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Acct/Card Ind Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDStaffInd().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Staff Salary Crediting Indicator Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Staff Salary Crediting Indicator Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDSubProduct().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Sub Product Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Sub Product Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDBrnachCodeName().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Branch Code Name Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Branch Code Name Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDCurencyCode().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Currency Code Name Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Currency Code Name Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If HK_CCTR_TDKeyInfo_Page.lblTDOpenDate().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Open Date Label is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Open Date Label in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	
	If HK_CCTR_TDKeyInfo_Page.eleRefreshIcon().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Refresh symbol is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Refresh symbol in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	
	If HK_CCTR_TDKeyInfo_Page.eleWarningInfoIcon().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Warning Info symbol is displayed as expected in " &strLeftMenuLinkName &" Page",True
		bverifyPinkPanelDetailsOfLeftMenuLink=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Warning Info symbol in " &strLeftMenuLinkName &" Page",false
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDProduct(), strProduct, "Product Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDAccCardNo(), strcardNumber, "Card Number Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDStatus(), strStatus, "Status Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDAccCardInd(), strAccCardInd, "Account Card Indicator Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDStaffInd(), strStaffIndc, "Staff Indicator Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDSubProduct(), strSubProduct, "Sub Product Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_TDKeyInfo_Page.weleTDOpenDate(), strOpenDate, "Open Date Text in " &strLeftMenuLinkName &" Page") Then
		bverifyPinkPanelDetailsOfLeftMenuLink=false
	End If

	verifyPinkPanelDetailsOfLeftMenuLink = bverifyPinkPanelDetailsOfLeftMenuLink
	
End Function

'[Verify Left Panel Links]
Public Function verifyLeftPanelLinks(arrLeftPanelLinks)
   bverifyLeftPanelLinks = False
	For i = 0 To UBound(arrLeftPanelLinks) Step 1
		bverifyLeftPanelLinks = clickLefmenuLink(arrLeftPanelLinks(i))
		WaitForICallLoading
	Next
	verifyLeftPanelLinks = bverifyLeftPanelLinks
End Function

'[Verify Account Balance section Existence]
Public Function verifyAccountBalanceSection()
	bverifyAccountBalanceSection = False
	If HK_CCTR_BalanceAndLimits_Page.eleAccountBalance().Exist(3) Then
		LogMessage "RSLT","Verification","Account Balance Section is displayed as expected in Balance & Limts Page",True
		bverifyAccountBalanceSection = True
	Else
		LogMessage "WARN","Verification","Failed to display Account Balance Section in Balance & Limits Page.",False
	End If
	
	bverifyAccountBalanceSection = bverifyAccountBalanceSection
End Function

'[Verify Limits section Existence]
Public Function verifyLimitsSection()
	bverifyLimitsSection = False
	If HK_CCTR_BalanceAndLimits_Page.eleLimits().Exist(3) Then
		LogMessage "RSLT","Verification","Limits Section is displayed as expected in Balance & Limts Page",True
		bverifyLimitsSection = True
	Else
		LogMessage "WARN","Verification","Failed to display Limits Section in Balance & Limits Page.",False
	End If
	verifyLimitsSection = bverifyLimitsSection
End Function

'[Verify Account Balance section field labels and values]
Public Function verifyAccountBalanceLabelsAndValues(arrLabelValuePairs)

	bverifyAccountBalanceLabelsAndValues = True
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblPassbookBalance(), Split(arrLabelValuePairs(0),":")(0), "Passbook Balance Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblPassbookBalanceValue(), Split(arrLabelValuePairs(0),":")(1), "Passbook Balance Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblCurrentBalance(), Split(arrLabelValuePairs(1),":")(0), "Current Balance Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblCurrentBalanceValue(), Split(arrLabelValuePairs(1),":")(1), "Current Balance Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblAvailableBalance(), Split(arrLabelValuePairs(2),":")(0), "Available Balance Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblAvailableBalanceValue(), Split(arrLabelValuePairs(2),":")(1), "Available Balance Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblCurrentSignal(), Split(arrLabelValuePairs(3),":")(0), "Current Signal Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If HK_CCTR_BalanceAndLimits_Page.lblCurrentSignalValue().Exist(3) Then
		If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblCurrentSignalValue(), Split(arrLabelValuePairs(3),":")(1), "Current Signal Value") Then
			bverifyAccountBalanceLabelsAndValues = False
		End If
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblFloat1(), Split(arrLabelValuePairs(4),":")(0), "Float 1 Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblFloat1Value(), Split(arrLabelValuePairs(4),":")(1), "Float 1 Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblFloat2(), Split(arrLabelValuePairs(5),":")(0), "Float 2 Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblFloat2Value(), Split(arrLabelValuePairs(5),":")(1), "Float 2 Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblSamTrnTotal(), Split(arrLabelValuePairs(6),":")(0), "Sam Trn Total Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblSamTrnTotalValue(), Split(arrLabelValuePairs(6),":")(1), "Sam Trn Total Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblForeignCCYClearingChequeDeposit(), Split(arrLabelValuePairs(7),":")(0), "Foreign CCY Clearing Cheque Deposit Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblForeignCCYClearingChequeDepositValue(), Split(arrLabelValuePairs(7),":")(1), "Foreign CCY Clearing Cheque Deposit Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblForeignCCYChequeInterestAccrual(), Split(arrLabelValuePairs(8),":")(0), "Foreign CCY Cheque w/o Interest Accrual Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblForeignCCYChequeInterestAccrualValue(), Split(arrLabelValuePairs(8),":")(1), "Foreign CCY Cheque w/o Interest Accrual Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblHoldFundAmount(), Split(arrLabelValuePairs(9),":")(0), "Hold Fund Amount Label") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblHoldFundAmountValue(), Split(arrLabelValuePairs(9),":")(1), "Hold Fund Amount Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblInterestRate(), Split(arrLabelValuePairs(10),":")(0), "Interest Rate Label") Then
		bverifyAccountBalanceLabelsAndValues = FalseRate
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblInterestRateValue(), Split(arrLabelValuePairs(10),":")(1), "Interest Rate Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblInterestAmountEarned(), Split(arrLabelValuePairs(11),":")(0), "Interest Amount Earned (YTD) Label") Then
		bverifyAccountBalanceLabelsAndValues = FalseRate
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblInterestAmountEarnedValue(), Split(arrLabelValuePairs(11),":")(1), "Interest Amount Earned (YTD) Value") Then
		bverifyAccountBalanceLabelsAndValues = False
	End If
	
End Function

'[Verify Limits section field labels and values]
Public Function verifyLimitsLabelsAndValues(arrLimitLabelValuePairs)

	verifyLimitsLabelsAndValues = True
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblOverallOverdraftLimit(), Split(arrLimitLabelValuePairs(0),"#")(0), "Overall Overdraft Limit Label") Then
		verifyLimitsLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblOverallOverdraftLimitValue(), Split(arrLimitLabelValuePairs(0),"#")(1), "Overall Overdraft Limit Value") Then
		verifyLimitsLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblLimitUsage(), Split(arrLimitLabelValuePairs(1),"#")(0), "Limit Usage Label") Then
		verifyLimitsLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblLimitUsageValue(), Split(arrLimitLabelValuePairs(1),"#")(1), "Limit Usage Value") Then
		verifyLimitsLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblLimitRemaining(), Split(arrLimitLabelValuePairs(2),"#")(0), "Limit Remaining Label") Then
		verifyLimitsLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblLimitRemainingValue(), Split(arrLimitLabelValuePairs(2),"#")(1), "Limit Remaining Value") Then
		verifyLimitsLabelsAndValues = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblAccruedOverdraftInterest(), Split(arrLimitLabelValuePairs(3),"#")(0), "Accrued Overdraft Interest (Net Amount: Credit Interest minus Debit Interest MTD) Label") Then
		verifyLimitsLabelsAndValues = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.lblAccruedOverdraftInterestValue(), Split(arrLimitLabelValuePairs(3),"#")(1), "Accrued Overdraft Interest (Net Amount: Credit Interest minus Debit Interest MTD) Value") Then
		verifyLimitsLabelsAndValues = False
	End If
	
End Function

'[Verify Account tab name and close opened tab]
Public Function verifyTabName(strProduct,strAccountNumber)
	Wait(2)
	bverifyTabName = True
	If Not verifyInnerText_Pattern(HK_CCTR_BalanceAndLimits_Page.eleAccountTab(),strProduct &"-"&strAccountNumber , "Account Tab Name") Then
		bverifyTabName = False
	End If
	ClickOnObject HK_CCTR_BalanceAndLimits_Page.btnAccountTabCloseIcon(),"Account Tab Close Icon"
	verifyTabName = bverifyTabName
End Function







'[Navigate to Cheque Info Page]
Public Function NavigateToChequeInfoPage()
	bNavigateToChequeInfoPage = False
	bNavigateToChequeInfoPage = clickLefmenuLink("Cheque Info")
	WaitForICallLoading
	NavigateToChequeInfoPage = bNavigateToChequeInfoPage
End Function

'[Select currency to verify cheque info]
Public Function SelectCurrencyForChequeInfo(strCurrencyName)
	bCurrency = False
	'bCurrency = selectItem_Combobox(HK_CCTR_BalanceAndLimits_Page.eleCurrencyDropdown(),strCurrencyName)
	bCurrency = SetValue(HK_CCTR_BalanceAndLimits_Page.txtCurrencyDropdown(),strCurrencyName,"Currency Drop Down")
	SelectCurrencyForChequeInfo = bCurrency
End Function

'[Verify No Record Message and Info warn popup Message in cheque info page]
Public Function VerifyChequeInfoNoRecordMessage(strMsg)
	bMsgCheck = True
	If VerifyFieldExistenceInPage(HK_CCTR_ChequeInfo_Page.eleNoRecordErrorMsg(),"Cheque Info Page","No Record Found Error Message") Then
		If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.eleNoRecordErrorMsg(), strMsg, "No Record Found Message in Cheque Info Page") Then
			bMsgCheck = False
		End If
	Else
		LogMessage "WARN","Verification","Failed to display No Records Found message in Cheque Info Page.",False
	End If
	VerifyFieldExistenceInPage HK_CCTR_ChequeInfo_Page.btnWarningInfoEnable(),"Cheque Info Page","Enabled Warning Icon"
	
	If	HK_CCTR_ChequeInfo_Page.btnWarningInfoEnable().Exist(2) Then
		bTempClick = ClickOnObject(HK_CCTR_ChequeInfo_Page.btnWarningInfoEnable(),"Enabled Warning Icon in Cheque Info Page")
		If bTempClick And HK_CCTR_ChequeInfo_Page.eleNoRecordErrorMsgInPopUp().Exist(2) Then
			If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.eleNoRecordErrorMsgInPopUp(), strMsg, "No Records Found Message in Cheque Info Page ") Then
				bMsgCheck = False
			End If
		Else
			LogMessage "WARN","Verification","Failed to display No Records Found message in Cheque Info Page warning Popup.",False
		End If
	Else
		LogMessage "WARN","Verification","Failed to display Info warn icon which is blinkable and clickable",False
	End If
		
	Wait(1)
	ClickOnObject HK_CCTR_ChequeInfo_Page.btnOkBtnInErrorMsgPopUp(),"Ok button in Warning Popup"
End Function

'[Verify cheque info page details]
Public Function VerifyChequeInfoPageDetails(strlstChequeInfoDtls)
	bChequeInfoDetails = False
	bChequeInfoDetails = selectTableLink(HK_CCTR_ChequeInfo_Page.tblChequeInfoHeader(),HK_CCTR_ChequeInfo_Page.tblChequeInfoContent(),strlstChequeInfoDtls,"Cheque Info Details" ,"Cheque Status",False,NULL,NULL,NULL)
	WaitForICallLoading
	VerifyChequeInfoPageDetails = bChequeInfoDetails
End Function

'[Verify more cheque info details]
Public Function VerifyMoreChequeDetails(strlstChequeDetails,strlstTblDetails)
	bChequeInfoDetails = True
	bChequeInfoTblDetails = False
	
	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.eleChequeInfoDetails(), strlstChequeDetails(0), "Cheque Details") Then
		bChequeInfoDetails = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.lblChequeInfoIssueDate(), Split(strlstChequeDetails(1),":")(0), "Issue Date Label") Then
		bChequeInfoDetails = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.lblChequeInfoIssueDateVal(), Split(strlstChequeDetails(1),":")(1), "Issue Date Value") Then
		bChequeInfoDetails = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.lblChequeInfoPaidCheque(), Split(strlstChequeDetails(2),":")(0), "Paid Cheque Label") Then
		bChequeInfoDetails = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.lblChequeInfoPaidChequeVal(), Split(strlstChequeDetails(2),":")(1), "Paid Cheque Value") Then
		bChequeInfoDetails = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.lblChequeInfoNoOfCheques(), Split(strlstChequeDetails(3),":")(0), "Total No.of Cheques Label") Then
		bChequeInfoDetails = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ChequeInfo_Page.lblChequeInfoNoOfChequesVal(), Split(strlstChequeDetails(3),":")(1), "Total No.of Cheques Value") Then
		bChequeInfoDetails = False
	End If
	
	If Not selectTableLink(HK_CCTR_ChequeInfo_Page.tblChequeDetailsHeader(),HK_CCTR_ChequeInfo_Page.tblChequeDetailsContent(),strlstTblDetails,"Cheque Details" ,False,False,NULL,NULL,NULL) Then
		bChequeInfoDetails = False
	End If
	
	VerifyFieldExistenceInPage HK_CCTR_ChequeInfo_Page.btnStopCheque(),"Cheque Info Page","Stop Cheque Button"
	
	'VerifyFieldExistenceInPage HK_CCTR_ChequeInfo_Page.btnNewChequeBook(),"Cheque Info Page","New Cheque Book Button"
	
	VerifyFieldExistenceInPage HK_CCTR_ChequeInfo_Page.btnWarningInfoDisable(),"Cheque Info Page","Disabled Warning Icon "
	
	VerifyMoreChequeDetails = bChequeInfoDetails
End Function




'[Navigate to Transaction History Page]
Public Function NavigateToTransactionHistoryPage()
	bNavigateToTranHistPage = False
	bNavigateToTranHistPage = clickLefmenuLink("Transaction History")
	WaitForICallLoading
	NavigateToTransactionHistoryPage = bNavigateToTranHistPage
End Function

'[Verify Transaction Period fields]
Public Function VerifyTransactionPeriodFields(strStmntCycelDate)
	bVerifyTransactionPeriodFields = False
	
	bVerifyTransactionPeriodFields= VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.lblStatementCycleDate(),"Transaction History Page","Statement Cycle Date Label")
	bVerifyTransactionPeriodFields = VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.lblStatementCycleDateVal(),"Transaction History Page","Statement Cycle Date Value")
	
	If verifyInnerText_Pattern(HK_CCTR_TransactionHistory_Page.lblStatementCycleDateVal(), strStmntCycelDate, "Transaction History-Statement Cycle Date Value") Then
		bVerifyTransactionPeriodFields = True
	End If
	
	bVerifyTransactionPeriodFields = VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.lblTransPeriod(),"Transaction History Page","Transaction Period Label")
	
	bVerifyTransactionPeriodFields = VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.lblTransPeriodFromDate(),"Transaction History Page","Transaction Period From Date Label")
	bVerifyTransactionPeriodFields = VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.txtTransPeriodFromDate(),"Transaction History Page","Transaction Period From Date Value")
	
	bVerifyTransactionPeriodFields= VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.lblTransPeriodFromDate(),"Transaction History Page","Transaction Period To Date Label")
	bVerifyTransactionPeriodFields = VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.txtTransPeriodToDate(),"Transaction History Page","Transaction Period To Date Value")
	
	bVerifyTransactionPeriodFields = VerifyFieldExistenceInPage(HK_CCTR_TransactionHistory_Page.btnGoTransHistPage(),"Transaction History Page","Go Button in Transaction Period")
	
	VerifyTransactionPeriodFields = bVerifyTransactionPeriodFields
End Function

'[Verify Multicurrency dropdown values Transaction History]
Public Function VerifyTranHistoryMultiCurrDropDown(strMultiCurrencies)
	bVerifyTranHistoryMultiCurrDropDown = False
	bVerifyTranHistoryMultiCurrDropDown = verifyComboboxItems(HK_CCTR_TransactionHistory_Page.eleCurrencyDropdown(),strMultiCurrencies,"Multi Currency Dropdown Transaction History")
	VerifyTranHistoryMultiCurrDropDown = bVerifyTranHistoryMultiCurrDropDown
End Function
 
'[Verify Transaction History Table contents]
Public Function VerifyTransactioHistroyTable(strFromDate,strToDate,arrTransHistTableData)
	bVerifyTransactioHistroyTable = False
	SetValue HK_CCTR_TransactionHistory_Page.txtTransPeriodFromDate(),strFromDate,"Transaction History-From Date"
	SetValue HK_CCTR_TransactionHistory_Page.txtTransPeriodToDate(),strToDate,"Transaction History-To Date"
	ClickOnObject HK_CCTR_TransactionHistory_Page.btnGoTransHistPage(),"Go button in Transaction History page"
	WaitForICallLoading
	
	bVerifyTransactioHistroyTable = verifyTableContentList(HK_CCTR_TransactionHistory_Page.tblTransHistHeader(),HK_CCTR_TransactionHistory_Page.tblTransHistContent(),arrTransHistTableData,"Transaction History Table",False,NULL,NULL,NULL)
	VerifyTransactioHistroyTable = bVerifyTransactioHistroyTable
End Function

'[Verify error messages with invalid From and To date]
Public Function VerifyErrorMessages(arrDataToValidate)
	bVerifyErrorMessages = True
	For i = 0 To UBound(arrDataToValidate) Step 1
		strFromDate = Split(arrDataToValidate(i),"#")(0)
		strToDate = Split(arrDataToValidate(i),"#")(1)
		strMessage = Split(arrDataToValidate(i),"#")(2)
		
		SetValue HK_CCTR_TransactionHistory_Page.txtTransPeriodFromDate(),strFromDate,"Transaction History-From Date"
		SetValue HK_CCTR_TransactionHistory_Page.txtTransPeriodToDate(),strToDate,"Transaction History-To Date"
		ClickOnObject HK_CCTR_TransactionHistory_Page.btnGoTransHistPage(),"Go button in Transaction History page"
		WaitForICallLoading
		If Not verifyInnerText_Pattern(HK_CCTR_TransactionHistory_Page.lblErrorMessages(), strMessage, "Invalid Date Error message") Then
			bVerifyErrorMessages = False
		End If
	Next
	VerifyErrorMessages = bVerifyErrorMessages
End Function

'[Select currency to verify Transaction History]
Public Function SelectCurrencyForTranHistory(strCurrencyName)
	bCurrency = False
	'bCurrency = selectItem_Combobox(HK_CCTR_BalanceAndLimits_Page.eleCurrencyDropdown(),strCurrencyName)
	bCurrency = SetValue(HK_CCTR_TransactionHistory_Page.txtCurrencyDropdown(),strCurrencyName,"Currency Drop Down")
	SelectCurrencyForTranHistory = bCurrency
End Function

'[Verify Date default value and From To date difference]
Public Function VerifyDateDifference()
	bVerifyDateDifference = False
	
	strToDate = Trim (HK_CCTR_TransactionHistory_Page.txtTransPeriodToDate().GetROProperty("value"))
	strMonthYear = Date & FormatDateTime(Date, 1)
	If Instr(strMonthYear,Split(strToDate," ")(0))>0 And Instr(strMonthYear,Split(strToDate," ")(1))>0 And Instr(strMonthYear,Split(strToDate," ")(2))>0 Then
		LogMessage "RSLT","Verification","As Expected Current Date ["&strToDate&"] is displayed in To Date field.",True
		bVerifyDateDifference = True		
	Else
		LogMessage "WARN","Verification","Failed to display Current Date in To Date field.",False 
	End If 
	If bVerifyDateDifference Then
		strFromDate = Trim (HK_CCTR_TransactionHistory_Page.txtTransPeriodFromDate().GetROProperty("value"))
		If DateDiff("d",strFromDate,strToDate) = 30 Then
			LogMessage "RSLT","Verification","As Expected From date "&[strFromDate]&" displayed as [Current Date - 30 Days], Current Date ["&strToDate&"]",True
			bVerifyDateDifference = True
		Else
			LogMessage "WARN","Verification","Failed to display From date "&[strFromDate]&" as [Current Date - 30 Days], Current Date ["&strToDate&"]",False
			bVerifyDateDifference = False
		End If
	Else
		LogMessage "WARN","Verification","Failed to Calculate From date and To Date difference , As To Date is not current Date.",False
	End If
	VerifyDateDifference = bVerifyDateDifference
End Function








'[Select closed accounts radio button in Account Overview Page]
Public Function SelectClosedRadioButton()
	SelectRadioButtonGrp "Closed", HK_CCTR_ClosedAccounts_Page.radioGroupShowAccounts(), ""
	WaitForICallLoading
End Function 

'[Verify Left Menu Links for closed accounts]
Public Function VerifyClosedAccountLinks()
	bVerifyClosedAccountLinks = False
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleBalanceAndLimits(),"Account Details Page","Balance And Limits Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleTransHistory(),"Account Details Page","Transaction History Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleKeyInfo(),"Account Details Page","Key Info Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleChequeInfo(),"Account Details Page","Cheque Info Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleAddressAccountLinakge(),"Account Details Page","Address and Account Linkage Left Menu Link")
	VerifyClosedAccountLinks = bVerifyClosedAccountLinks
End Function

'[Verify Left Menu Links for TD closed account]
Public Function VerifyTDClosedAccountLinks()
	bVerifyClosedAccountLinks = False
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleBalanceAndLimits(),"Account Details Page","Balance And Limits Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleTransHistory(),"Account Details Page","Transaction History Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleKeyInfo(),"Account Details Page","Key Info Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleChequeInfo(),"Account Details Page","Cheque Info Left Menu Link")
	bVerifyClosedAccountLinks = VerifyFieldRemovedFromPage( HK_CCTR_ClosedAccounts_Page.eleAddressAccountLinakge(),"Account Details Page","Address and Account Linkage Left Menu Link")
	
	bVerifyClosedAccountLinks = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.elePlacementInfo(),"Account Details Page","Placement Info Left Menu Link")
	
	VerifyTDClosedAccountLinks = bVerifyClosedAccountLinks
End Function


'[Verify Account information details in key info page]
Public Function VerifyAccountLabelsAndValuesInKeyInfoPage(arrLabelValuePairs)
	bVerifyAccountLabelsAndValuesInKeyInfoPage = True
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.elePrimaryCIFlbl(), Split(arrLabelValuePairs(0),":")(0), "Primary CIF Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.elePrimaryCIFVal(), Split(arrLabelValuePairs(0),":")(1), "Primary CIF Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleAcountTypelbl(), Split(arrLabelValuePairs(1),":")(0), "Account Type Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleAcountTypeVal(), Split(arrLabelValuePairs(1),":")(1), "Account Type Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleAccountSignatoryTypelbl(), Split(arrLabelValuePairs(2),":")(0), "Account Signatory Type Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleAccountSignatoryTypeVal(), Split(arrLabelValuePairs(2),":")(1), "Account Signatory Type Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleStatuslbl(), Split(arrLabelValuePairs(3),":")(0), "Status Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleStatusVal(), Split(arrLabelValuePairs(3),":")(1), "Status Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleOpeningDatelbl(), Split(arrLabelValuePairs(4),":")(0), "Opening Date Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleOpeningDateVal(), Split(arrLabelValuePairs(4),":")(1), "Opening Date Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleClosedDatelbl(), Split(arrLabelValuePairs(5),":")(0), "Closed Date Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleClosedDateVal(), Split(arrLabelValuePairs(5),":")(1), "Closed Date Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleSalaryCreditingIndicatorlbl(), Split(arrLabelValuePairs(6),":")(0), "Salary Crediting Indicator Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleSalaryCreditingIndicatorVal(), Split(arrLabelValuePairs(6),":")(1), "Salary Crediting Indicator Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleLastTransactionDatelbl(), Split(arrLabelValuePairs(7),":")(0), "Last Transaction Date Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleLastTransactionDateVal(), Split(arrLabelValuePairs(7),":")(1), "Last Transaction Date Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleServChgWaiveIndlbl(), Split(arrLabelValuePairs(8),":")(0), "Serv Chg Waive Ind Label") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If

	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleServChgWaiveIndVal(), Split(arrLabelValuePairs(8),":")(1), "Serv Chg Waive Ind Value") Then
		bVerifyAccountLabelsAndValuesInKeyInfoPage = False
	End If
	
	VerifyAccountLabelsAndValuesInKeyInfoPage = bVerifyAccountLabelsAndValuesInKeyInfoPage
End Function

'[Verify Account holders detail]
Public Function VerifyAccountHoldersDetail(strAccountHolderDetails)
	bAccountHolder = False
	bAccountHolder = selectTableLink(HK_CCTR_ClosedAccounts_Page.tblAccountHolderHeader(),HK_CCTR_ClosedAccounts_Page.tblAccountHolderContent(),strAccountHolderDetails,"Account Holder(s) Detail" ,"CIF",False,NULL,NULL,NULL)
	WaitForICallLoading
	bAccountHolder = VerifyAccountHoldersDetail
End Function

'[Verify Customer Profile details for closed accounts]
Public Function VerifyClosedAccountsCustProfile(strTabName)
	Wait(2)
	bVerifyClosedAccountsCustProfile = True
	If Not verifyInnerText_Pattern(HK_CCTR_ClosedAccounts_Page.eleCustomerProfileTab(),strTabName, "Tab Name") Then
		bVerifyClosedAccountsCustProfile = False
		LogMessage "WARN","Verification","Failed to display Customer Profile page on clicking CIF hyperlink from Key Info page for closed accounts.",False
	Else
		LogMessage "RSLT","Verification","Customer Profile page displayed as expected on clicking CIF hyperlink from Key Info page for closed accounts.",True
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleCustomerProfileKeyInfo(),"Customer Profile-Closed Account","Key Info Section")
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleCustomerProfileFATCADetails(),"Customer Profile-Closed Account","FATCA Details Section")
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleCustomerProfilePersonalInfo(),"Customer Profile-Closed Account","Personal Info Section")
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleCustomerProfileEmploymentDetails(),"Customer Profile-Closed Account","Employment Details Section")
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleCustomerProfileAddress(),"Customer Profile-Closed Account","Address Section")
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.eleCustomerProfileMarketingPreferences(),"Customer Profile-Closed Account","Marketing Preferences Section")
		
		bVerifyClosedAccountsCustProfile = VerifyFieldExistenceInPage( HK_CCTR_ClosedAccounts_Page.btnMarketingPreferences(),"Customer Profile-Closed Account","Marketing Preferences Button")
	End If
	VerifyClosedAccountsCustProfile = bVerifyClosedAccountsCustProfile
End Function

'[Verify TD Deposit information details in Placement info page]
Public Function VerifyDepositInPlacementInfoPage(arrLabelValuePairs,arrLabelValuePairs1,arrLabelValuePairs2)
	bDeposit = False
	bDeposit = selectTableLink(HK_CCTR_ClosedAccounts_Page.tblTDPlaceInfoDepositHeader(),HK_CCTR_ClosedAccounts_Page.tblTDPlaceInfoDepositContent(),arrLabelValuePairs,"TD Deposit Details" ,"Deposit No.",False,NULL,NULL,NULL)
	'bDeposit = verifyTableContentList(HK_CCTR_ClosedAccounts_Page.tblTDPlaceInfoDepositHeader1(),HK_CCTR_ClosedAccounts_Page.tblTDPlaceInfoDepositContent1(),arrLabelValuePairs1,"TD Avl Balance",False,Null,Null,Null)
	'bDeposit = verifyTableContentList(HK_CCTR_ClosedAccounts_Page.tblTDPlaceInfoDepositHeader2(),HK_CCTR_ClosedAccounts_Page.tblTDPlaceInfoDepositContent2(),arrLabelValuePairs2,"TD Placement Status",False,Null,Null,Null)
	VerifyDepositInPlacementInfoPage = bDeposit
End Function
