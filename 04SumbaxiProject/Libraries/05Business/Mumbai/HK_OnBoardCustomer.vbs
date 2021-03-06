'[Enter Customer English Name after search]
Public Function EnterCustomerEnglishNameAfterSearch(strCustomerName)
	WaitForICallLoading
	strNow = Replace(Replace(Replace(Now," ", ""), "/",""), ":","")
	Environment.Value("ApplicantName") = strCustomerName&strNow
	SetValue HK_OnBoardCustomer_Page.txtCustNameInEnglishAfterSearch(),strCustomerName&strNow,"Customer Name in English"
End Function

'[Enter Customer English Name after search for Rejection]
Public Function EnterCustomerEnglishNameAfterSearchReject(strCustomerName)
	Environment.Value("ApplicantName") = strCustomerName
	SetValue HK_OnBoardCustomer_Page.txtCustNameInEnglishAfterSearch(),strCustomerName,"Customer Name in English"
End Function

'[Add Single Applicant for onboarding]
Public Function AddSingleApplicant()
	bAddSingleApplicant = False
	
	HK_OnBoardCustomer_Page.eleAddAsApplicant().Click
	Wait(3)
	If HK_OnBoardCustomer_Page.eleSelectedApplicant().Exist(5) Then
		LogMessage "RSLT","Verification","On Clicking Add as Applicant icon, Panel appeared with Name in English,ID Type and ID Number",True
		bAddSingleApplicant = true
	Else
		LogMessage "WARN","Verification","On Clicking Add as Applicant icon, Panel did not appear with Name in English,ID Type and ID Number",False
		bAddSingleApplicant = false		
	End If
	AddSingleApplicant = bAddSingleApplicant
End Function

'[Click on Account Overview Icon]
Public Function ClickAccountOverview()
	bClickAccountOverview = False
	ClickOnObject HK_OnBoardCustomer_Page.eleAccountOverview(),"Account Overview Icon"
	WaitForICallLoading
	If HK_OnBoardCustomer_Page.eleAccountOverviewHeaderName().Exist(2) Then
		LogMessage "RSLT","Verification","On Clicking Account Overview Icon, Account Overview Page is displayed",True
		bClickAccountOverview = true
	Else
		LogMessage "WARN","Verification","On Clicking Account Overview Icon, Account Overview Page is not displayed",False
	End If
	ClickAccountOverview = bClickAccountOverview
End Function

'[Verify Open Additional Account Presence and Status]
Public Function VerifyOpenAdditionalAccount(strEnDsFlag,strAccntSts)
	bTemp = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.btnOpenAddAccount(),"Account Overview","Open Additional Account")
	If bTemp Then
		LogMessage "RSLT","Verification","Open Additional Account Button exists in Account Overview Page.",True
		strBtnStatus = HK_OnBoardCustomer_Page.btnOpenAddAccount().GetROProperty ("disabled")
		If strEnDsFlag = "Yes" Then
			If strBtnStatus = 0 Then
				LogMessage "RSLT","Verification","As Expected , Open Additional Account Button is enabled for account Type: " &strAccntSts ,True
			Else
				LogMessage "WARN","Verification","Failed , Open Additional Account Button is disabled for account Type: " &strAccntSts ,False
			End If
		ElseIf strEnDsFlag = "No" Then
			If strBtnStatus = 1 Then
				LogMessage "RSLT","Verification","As Expected , Open Additional Account Button is disabled for account Type: " &strAccntSts ,True
			Else
				LogMessage "WARN","Verification","Failed , Open Additional Account Button is enabled for account Type: " &strAccntSts ,False
			End If
		End If
	Else
		LogMessage "WARN","Verification","Open Additional Account Button does not Exist in Account Overview Page.",False
	End If
	
End Function 

'[OnBoard Single Applicant]
Public Function OnBoardSingleApplicant()
	bOnBoardSingleApplicant = False
	HK_OnBoardCustomer_Page.btnStartOnBoarding().Click
	WaitForICallLoading
	If HK_OnBoardCustomer_Page.lnkEnityRelationship().Exist(5) Then
		LogMessage "RSLT","Verification","On Clicking Start OnBoarding button, Entity relationship page displayed as expected.",True
		bOnBoardSingleApplicant = True
	Else
		LogMessage "WARN","Verification","On Clicking Start OnBoarding button, Failed to display Entity Relationship page.",False
		bOnBoardSingleApplicant = False		
	End If
	OnBoardSingleApplicant = bOnBoardSingleApplicant
End Function

'[OnBoard Multiple Applicant]
Public Function OnBoardMultipleApplicant(strIDType,strIDNumber,strErrFlag,strMsg,strCustomerName)
	For i = 0 To UBound(strIDType) Step 1
		SearchCustomerByID strIDType(i),strIDNumber(i),strErrFlag(i),strMsg(i)
		EnterCustomerEnglishNameAfterSearchReject strCustomerName(i)
		AddSingleApplicant
	Next
	OnBoardSingleApplicant
	WaitForICallLoading
End Function

'[Enter Relationship Type]
Public Function EnterRelationshipType(strRelType)
	bRelation = False
	bRelation = SetValue (HK_OnBoardCustomer_Page.lstRelationshipType(),strRelType,"Relationship Type")
	WaitForICallLoading
	If strRelType = "Trust Minor" Then
		ClickOnObject HK_OnBoardCustomer_Page.chkTrustMinor(),"Trust Minor Check Box-Primary Applicant"
	End If
	
	If strRelType = "Other" Then
		ClickOnObject HK_OnBoardCustomer_Page.chkCanOperate(),"Can Operate On Account Check Box-Primary Applicant"
	End If
	
	ClickOnObject HK_OnBoardCustomer_Page.btnNext(),"Next Button"
	EnterRelationshipType = bRelation
End Function

'[Verify Entity Details for different Relationship Type]
Public Function VerifyMultiEntityDetails(lstRelTypes,lstJointAllRowVals,lstJointAnyRowVals,lstTrustMinRowVals,lstOtherRowVals)
	bVerifyEntityDetails = False
	For i = 0 To UBound(lstRelTypes) Step 1
		HK_OnBoardCustomer_Page.lstRelationshipType().Set lstRelTypes(i)
		Wait(2)
		Select Case lstRelTypes(i)
			Case "Joint - All to sign"
				bVerifyEntityDetails = CheckTableRowData( lstRelTypes(i),lstJointAllRowVals)
			Case "Joint - Any one to sign"
				bVerifyEntityDetails = CheckTableRowData(lstRelTypes(i),lstJointAnyRowVals)
			Case "Trust Minor"
				bVerifyEntityDetails = CheckTableRowData(lstRelTypes(i),lstTrustMinRowVals)
			Case "Other"
				bVerifyEntityDetails = CheckTableRowData(lstRelTypes(i),lstOtherRowVals)
		End Select
	Next
	VerifyMultiEntityDetails = bVerifyEntityDetails
End Function

Public Function CheckTableRowData(strRelType,lstRowVals)
		bCheckTableData = False
		Set tblObj = HK_OnBoardCustomer_Page.tblEntityDetails()
		
		For i = 0 To UBound(lstRowVals) Step 1
		
			strApp1 = tblObj.GetCellData(i+1,1)
			strTemp1 = Split(lstRowVals(i),":")(0)
			bCheckTableData = CheckInstrText(strRelType, strApp1,strTemp1)
			
			strApp2 = tblObj.GetCellData(i+1,2)
			strTemp2 = Split(lstRowVals(i),":")(1)
			bCheckTableData = CheckInstrText (strRelType,strApp2,strTemp2)
			
			strApp3 = tblObj.GetCellData(i+1,3)
			strTemp3 = Split(lstRowVals(i),":")(2)
			bCheckTableData = CheckInstrText(strRelType,strApp3,strTemp3)
			
			If strRelType = "Other" Then
				strApp4 = tblObj.GetCellData(i+1,4)
				strTemp4 = Split(lstRowVals(i),":")(3)
				bCheckTableData = CheckInstrText(strRelType,strApp4,strTemp4)
			End If
		Next
		CheckTableRowData = bCheckTableData
End Function

Public Function CheckInstrText(strRelType,strVal1,strVal2)
	bTableText = False
	If Instr(1,strVal1,strVal2) > 0	Then
		LogMessage "RSLT","Verification",strVal1 & " displayed as Expected for Relationship Type: "& strRelType & " in Entity Details Table",True
		bTableText = True
	Else
		LogMessage "WARN","Verification",strVal1 & " not displayed as Expected for Relationship Type: "& strRelType & " in Entity Details Table, Actual : " & strVal1 & ", Expected:"&strVal2,False
	End If
	CheckInstrText = bTableText
End Function

'[Verify Relationship Type list box value]
Public Function VerifyRelationshipType(strRelationshipType)
	verifyComboList strRelationshipType, HK_OnBoardCustomer_Page.lstRelationshipType()
End Function

'[Verify Branch ID list box value]
Public Function VerifyBranchID(strBranchID)
	verifyComboList Split(strBranchID,"|"), HK_OnBoardCustomer_Page.lstBranchID()
End Function

'[Verify Name in Entity details Table]
Public Function VerifyEntityDetailsName(strName)
	bVerifyEntityDetailsName = False
	If InStr(1,Trim(HK_OnBoardCustomer_Page.eleEntityDetailsName().GetROProperty("innertext")),strName) > 0 Then 
		bVerifyEntityDetailsName  = True
		LogMessage "RSLT","Verification","[Name in English] is displayed as expected in Entity Details Table",True
	Else
		LogMessage "WARN","Verification","Failed - [Name in English] is not displayed as expected in Entity Details Table",False
	End if
	VerifyEntityDetailsName = bVerifyEntityDetailsName
End Function

'[Verify ID Type in Entity details Table]
Public Function VerifyEntityDetailsIDType(strIDType)
	bVerifyEntityDetailsIDType = False
	If InStr(1,Trim(HK_OnBoardCustomer_Page.eleEntityDetailsIDTypeNo().GetROProperty("innertext")),strIDType) > 0  Then 
		bVerifyEntityDetailsIDType  = True
		LogMessage "RSLT","Verification","[ID Type] is displayed as expected in Entity Details Table",True
	Else
		LogMessage "WARN","Verification","Failed - [ID Type] is not displayed as expected in Entity Details Table",False
	End if
	VerifyEntityDetailsIDType = bVerifyEntityDetailsIDType
End Function

'[Verify ID No in Entity details Table]
Public Function VerifyEntityDetailsIDNo(strIDNo)
	bVerifyEntityDetailsID = False
	If InStr(1,Trim(HK_OnBoardCustomer_Page.eleEntityDetailsIDTypeNo().GetROProperty("innertext")),strIDNo) > 0 Then 
		bVerifyEntityDetailsID  = True
		LogMessage "RSLT","Verification","[ID No] is displayed as expected in Entity Details Table",True
	Else
		LogMessage "WARN","Verification","Failed - [ID No] is not displayed as expected in Entity Details Table",False
	End if
	VerifyEntityDetailsIDNo = bVerifyEntityDetailsID
End Function

'[Verify existance of Entity setup drop down in Entity details Table]
Public Function VerifyEntityDetailsEntitySetup()
	bVerifyEntityDetailsEntitySetup = False
	If HK_OnBoardCustomer_Page.eleEntityDetailsEntitySetup().Exist(2) Then 
		bVerifyEntityDetailsEntitySetup  = True
		LogMessage "RSLT","Verification","[Entity Setup Drop down] is displayed as expected in Entity Details Table",True
	Else
		LogMessage "WARN","Verification","Failed - [Entity Setup Drop down] is not displayed as expected in Entity Details Table",False
	End if
	VerifyEntityDetailsEntitySetup = bVerifyEntityDetailsEntitySetup
End Function

'[Verify navigation to Customer OnBoarding page by selecting different Entity Setup]
Public Function VerifyNavigationToCustomerOnBoardingPage(strEntitySetups)
	bVerifyNavigationToCustomerOnBoardingPage = False
	Dim arrEntitySetups : arrEntitySetups = strEntitySetups 'Split(strEntitySetups,"|")
	For i = 0 To UBound(arrEntitySetups) Step 1
		 HK_OnBoardCustomer_Page.lstEntitySetup().Set(arrEntitySetups(i))
		 Wait(2)
		 If Trim(HK_OnBoardCustomer_Page.lstEntitySetup().GetROProperty("value")) = arrEntitySetups(i) Then
		 	bVerifyNavigationToCustomerOnBoardingPage = True
		 	LogMessage "RSLT","Verification","Entity Setup ["&arrEntitySetups(i)& "] is selected as expected in Entity Setup drop down.",True
		 	HK_OnBoardCustomer_Page.btnNext().Click
		 	WaitForICallLoading
		 	If HK_OnBoardCustomer_Page.eleAccountInformationPink().Exist(2) Then
		 		LogMessage "RSLT","Verification","When ["&arrEntitySetups(i)& "] is selected in Entity Set Up drop down user is able to navigate to next page.",True
		 	Else
		 		LogMessage "WARN","Verification","When ["&arrEntitySetups(i)& "] is selected in Entity Set Up drop down user is unable to navigate to next page.",False
		 	End If
		 	HK_OnBoardCustomer_Page.btnPrevious().Click
		 	WaitForICallLoading
		 Else
		 	bVerifyNavigationToCustomerOnBoardingPage = False
		 	LogMessage "WARN","Verification","Failed to select Entity Setup ["&arrEntitySetups(i)& "] in Entity Setup drop down.",False
		 End If
	Next
End Function

'[Enter data in entity relationship page]
Public Function EntityRelationship(strEntitySetup)
	WaitForICallLoading
	VerifyFieldRemovedFromPage HK_OnBoardCustomer_Page.lstCustomerSegement(),"Entity Relationship","Customer Segement"
	VerifyFieldRemovedFromPage HK_OnBoardCustomer_Page.lstPreferredLanguage(),"Entity Relationship","Prefered Language"
	
	If UCase(strEntitySetup) <> "NO ENTITY" Then
		SetValue HK_OnBoardCustomer_Page.lstEntitySetup(),strEntitySetup,"Entity Setup"
	End If
	
	ClickOnObject HK_OnBoardCustomer_Page.btnNext(),"Next Button"
End Function

'[Verify Purpose of Accounts under Account Information]
Public Function VerifyPurposeOfAccounts(lstPurposeOfAccounts)
	bPrpsOfAccnts = False
	ClickOnObject HK_OnBoardCustomer_Page.btnNext(),"Next Button"
	WaitForICallLoading
	For i = 0 To UBound(lstPurposeOfAccounts) Step 1
		Select Case UCase(lstPurposeOfAccounts(i))
			Case "SAVINGS"
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkPurposeOfAccountSavings(),"Account information-Purpose of Account-",lstPurposeOfAccounts(i))
			Case "INVESTMENT"
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkPurposeOfAccountInvestment(),"Account information-Purpose of Account-",lstPurposeOfAccounts(i))
			Case "TRANSACTIONAL"
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkPurposeOfAccountTransactional(),"Account information-Purpose of Account-",lstPurposeOfAccounts(i))
			Case "PAYROLL"
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkPurposeOfAccountPayroll(),"Account information-Purpose of Account-",lstPurposeOfAccounts(i))
			Case "LOAN REPAYMENT"
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkPurposeOfAccountLoanRepayment(),"Account information-Purpose of Account-",lstPurposeOfAccounts(i))
			Case "OTHER"
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkPurposeOfAccountOther(),"Account information-Purpose of Account-",lstPurposeOfAccounts(i))
				bPrpsOfAccnts = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtPurposeOfAccountOther(),"Account information-Purpose of Account-", "Other-Text Box")
		End Select
	Next
	VerifyPurposeOfAccounts = bPrpsOfAccnts
End Function


'[Select Purpose of Account under Account Information]
Public Function SelectPurposeOfAccount(strPurposeOfAccount,strOther)
	WaitForICallLoading
	
	If Not IsArray(strPurposeOfAccount) Then
		strPurposeOfAccount = Array(strPurposeOfAccount)
	End If
	
	For i = 0 To UBound(strPurposeOfAccount) Step 1
		Select Case UCase(strPurposeOfAccount(i))
			Case "SAVINGS"
				ClickOnObject HK_OnBoardCustomer_Page.chkPurposeOfAccountSavings(),"Purpose of Account-"&strPurposeOfAccount(i)
			Case "INVESTMENT"
				ClickOnObject HK_OnBoardCustomer_Page.chkPurposeOfAccountInvestment(),"Purpose of Account-"&strPurposeOfAccount(i)
			Case "TRANSACTIONAL"
				ClickOnObject HK_OnBoardCustomer_Page.chkPurposeOfAccountTransactional(),"Purpose of Account-"&strPurposeOfAccount(i)
			Case "PAYROLL"
				ClickOnObject HK_OnBoardCustomer_Page.chkPurposeOfAccountPayroll(),"Purpose of Account-"&strPurposeOfAccount(i)
			Case "LOAN REPAYMENT"
				ClickOnObject HK_OnBoardCustomer_Page.chkPurposeOfAccountLoanRepayment(),"Purpose of Account-"&strPurposeOfAccount(i)
			Case "OTHER"
				ClickOnObject HK_OnBoardCustomer_Page.chkPurposeOfAccountOther(),"Purpose of Account-"&strPurposeOfAccount(i)
				SetValue HK_OnBoardCustomer_Page.txtPurposeOfAccountOther(),strOther,"Purpose Of Account-Other"
		End Select
	Next
End Function

'[Verify Source of Funds under Account Information]
Public Function VerifySourceOfFunds(strSourceOfFunds)
	bSrcOfFunds = False
	For i = 0 To UBound(strSourceOfFunds) Step 1
		Select Case UCase(strSourceOfFunds(i))
			Case "SALARY"
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfFundsSalary(),"Account information-Source of Funds-",strSourceOfFunds(i))
			Case "SAVINGS"
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfFundsSaving(),"Account information-Source of Funds-",strSourceOfFunds(i))
			Case "SALES OF INVESTMENT"
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfFundsSalesOfInvestment(),"Account information-Source of Funds-",strSourceOfFunds(i))
			Case "SALE OF REAL ESTATE"
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfFundsSaleOfRealEstate(),"Account information-Source of Funds-",strSourceOfFunds(i))
			Case "OWN BUSINESS"
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfFundsOwnBusiness(),"Account information-Source of Funds-",strSourceOfFunds(i))
			Case "OTHER"
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfFundsOther(),"Account information-Source of Funds-",strSourceOfFunds(i))
				bSrcOfFunds = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSourceOfFundsOther(),"Account information-Source of Funds-" ,"Other Text box")
		End Select
	Next
	VerifySourceOfFunds = bSrcOfFunds
End Function

'[Select Source of Funds under Account Information]
Public Function SelectSourceOfFunds(strSourceOfFunds,strOther)
	
	If Not IsArray(strSourceOfFunds) Then
		strSourceOfFunds = Array(strSourceOfFunds)
	End If
	
	For i = 0 To UBound(strSourceOfFunds) Step 1
		Select Case UCase(strSourceOfFunds(i))
			Case "SALARY"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfFundsSalary(),"Source of Funds-"&strSourceOfFunds(i)
			Case "SAVINGS"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfFundsSaving(),"Source of Funds-"&strSourceOfFunds(i)
			Case "SALES OF INVESTMENT"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfFundsSalesOfInvestment(),"Source of Funds-"&strSourceOfFunds(i)
			Case "SALE OF REAL ESTATE"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfFundsSaleOfRealEstate(),"Source of Funds-"&strSourceOfFunds(i)
			Case "OWN BUSINESS"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfFundsOwnBusiness(),"Source of Funds-"&strSourceOfFunds(i)
			Case "OTHER"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfFundsOther(),"Source of Funds-"&strSourceOfFunds(i)
				SetValue HK_OnBoardCustomer_Page.txtSourceOfFundsOther(),strOther,"Source of Funds-Other-"&strOther
		End Select
	Next
End Function

'[Verify Deposit Frequency Dropdown values in Account Information Page]
Public Function VerifyDepositFreqncyDropdown(lstDepositFrqncy,strDefaultDepoFr)
	bDepoFrqncy = False
	bDepoFrqncy = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstDepositFrequency(),strDefaultDepoFr,"Deposit Frequency")
	For i = 0 To UBound(lstDepositFrqncy) Step 1
		bDepoFrqncy = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstDepositFrequency(),lstDepositFrqncy(i),"Deposit Frequency")
	Next
	VerifyDepositFreqncyDropdown = bDepoFrqncy
End Function

'[Verify Deposit Amount Dropdown values in Account Information Page]
Public Function VerifyDepositAmntDropdown(lstDepositAmnts,strDefaultDepoAmnt)
	bDepoAmnt = False
	bDepoAmnt = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstDepositAmount(),strDefaultDepoAmnt,"Deposit Amount")
	For i = 0 To UBound(lstDepositAmnts) Step 1
		bDepoAmnt = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstDepositAmount(),lstDepositAmnts(i),"Deposit Amount")
	Next
	VerifyDepositAmntDropdown = bDepoAmnt
End Function

'[Verify Withdrawal Frequency Dropdown values in Account Information Page]
Public Function VerifyWithdrawalFreqncyDropdown(lstwithdrFrqncy,strDefaultWithFr)
	bwithdrFrqncy = False
	bwithdrFrqncy = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstWithdrawalFrequency(),strDefaultWithFr,"Withdrawal Frequency")
	For i = 0 To UBound(lstwithdrFrqncy) Step 1
		bwithdrFrqncy = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstWithdrawalFrequency(),lstwithdrFrqncy(i),"Withdrawal Frequency")
	Next
	VerifyWithdrawalFreqncyDropdown = bwithdrFrqncy
End Function

'[Verify Withdrawal Amount Dropdown values in Account Information Page]
Public Function VerifyWithdrawalAmntDropdown(lstWithdrawalAmnts,strDefaultWithAmnt)
	bWithdrawalAmnt = False
	bWithdrawalAmnt = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstWithdrawalAmount(),strDefaultWithAmnt,"Withdrawal Amount")
	For i = 0 To UBound(lstWithdrawalAmnts) Step 1
		bWithdrawalAmnt = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstWithdrawalAmount(),lstWithdrawalAmnts(i),"Withdrawal Amount")
	Next
	VerifyWithdrawalAmntDropdown = bWithdrawalAmnt
End Function

'[Enter Anticipated Monthly Volume under Account Information]
Public Function EnterAnticipatedMonthlyVolume(strDepoFrqncy,strDepoAmnt,strWithdrawlFrqncy,strWithdrawlAmnt)
	SetValue HK_OnBoardCustomer_Page.lstDepositFrequency(),strDepoFrqncy,"Deposit Frequency-"
	SetValue HK_OnBoardCustomer_Page.lstDepositAmount(),strDepoAmnt,"Deposit Amount-"
	SetValue HK_OnBoardCustomer_Page.lstWithdrawalFrequency(),strWithdrawlFrqncy,"Withdrawal Frequency-"
	SetValue HK_OnBoardCustomer_Page.lstWithdrawalAmount(),strWithdrawlAmnt,"Withdrawal Amount-"
End Function

'[Verify Customer Interest Options under Account Information]
Public Function VerifyCustomerInterestOptionss(strCustomerInterestOptions)
	bCIO = False
	For i = 0 To UBound(strCustomerInterestOptions) Step 1
		Select Case UCase(strCustomerInterestOptions(i))
			Case "ALL INVESTMENT PRODUCTS"
				bCIO = VerifyFieldExistenceInPage( HK_OnBoardCustomer_Page.chkCustomerInterestOptionsAllInvestment(),"Account information-Customer Interest Options-",strCustomerInterestOptions(i))
			Case "ALL INSURANCE PRODUCTS"
				bCIO = VerifyFieldExistenceInPage( HK_OnBoardCustomer_Page.chkCustomerInterestOptionsAllInsurance(),"Account information-Customer Interest Options-",strCustomerInterestOptions(i))
			Case "MARKET UPDATE"
				bCIO = VerifyFieldExistenceInPage( HK_OnBoardCustomer_Page.chkCustomerInterestOptionsMarketUpdate(),"Account information-Customer Interest Options-",strCustomerInterestOptions(i))
			Case "ALL OF THE ABOVE"
				bCIO = VerifyFieldExistenceInPage( HK_OnBoardCustomer_Page.chkCustomerInterestOptionsAllOfTheAbove(),"Account information-Customer Interest Options-",strCustomerInterestOptions(i))
		End Select
	Next
	VerifyCustomerInterestOptionss = bCIO
End Function

'[Select Customer Interest Options under Account Information]
Public Function SelectCustomerInterestOptions(strCustomerInterestOptions)
	If Not IsArray(strCustomerInterestOptions) Then
		If UCase(strCustomerInterestOptions) = "ALL OF THE ABOVE" Then
			ClickOnObject HK_OnBoardCustomer_Page.chkCustomerInterestOptionsAllOfTheAbove(),"Customer Interest Options-All Investment Products, "&_																					
			"All Insurance Products & Market Update"
		End If
	Else
		For i = 0 To UBound(strCustomerInterestOptions) Step 1
			Select Case UCase(strCustomerInterestOptions(i))
				Case "ALL INVESTMENT PRODUCTS"
					ClickOnObject HK_OnBoardCustomer_Page.chkCustomerInterestOptionsAllInvestment(),"Customer Interest Options-"&strCustomerInterestOptions(i)
				Case "ALL INSURANCE PRODUCTS"
					ClickOnObject HK_OnBoardCustomer_Page.chkCustomerInterestOptionsAllInsurance(),"Customer Interest Options-"&strCustomerInterestOptions(i)
				Case "MARKET UPDATE"
					ClickOnObject HK_OnBoardCustomer_Page.chkCustomerInterestOptionsMarketUpdate(),"Customer Interest Options-"&strCustomerInterestOptions(i)
			End Select
		Next
	End If
End Function

'[Navigate to customer onboarding Account information page]
Public Function NavigateToAccountInformation()
	ClickOnObject HK_OnBoardCustomer_Page.eleNameInApplicantDetails(),"New Customer Name Under Applicant Details"
	For i = 1 To 20 Step 1
		If HK_OnBoardCustomer_Page.dlgEWSSCheckPopup().Exist(1) Then
			ClickOnObject HK_OnBoardCustomer_Page.dlgEWSSCheckPopupCancel(),"Cancel Button in EWSS Pop up"
			Exit For
		End If
	Next
	WaitForICallLoading
End Function

'[Verify Customer Segment Dropdown values in Account Details Page]
Public Function VerifyCustSegmentDropdown(lstCustSegement,strDefaultCustSeg)
	bCustSegement = False
	bCustSegement = VerifyFieldExistenceInPage( HK_OnBoardCustomer_Page.lstCustomerSegement(),"Account Details","Customer Segement")
	If bCustSegement Then
		bCustSegement = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstCustomerSegement(),strDefaultCustSeg,"Customer Segment")
		For i = 0 To UBound(lstCustSegement) Step 1
			bCustSegement = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstCustomerSegement(),lstCustSegement(i),"Customer Segment")
		Next
	Else
		LogMessage "WARN", "Verification","Customer Segment Dropdown does not exist in Account Details Page",False
	End If
	VerifyCustSegmentDropdown = bCustSegement
End Function

'[Verify Preferred Language Dropdown values in Account Details Page]
Public Function VerifyPreferredLanguageDropdown(lstPrfdLang,strDefaultPrfdLang)
	bPrfdLang = False
	bPrfdLang = VerifyFieldExistenceInPage( HK_OnBoardCustomer_Page.lstPreferredLanguage(),"Account Details","Preferred Language")
	If bPrfdLang Then
		bPrfdLang = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstPreferredLanguage(),strDefaultPrfdLang,"Preferred Language")
		For i = 0 To UBound(lstPrfdLang) Step 1
			bPrfdLang = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstPreferredLanguage(),lstPrfdLang(i),"Preferred Language")
		Next
	Else
		LogMessage "WARN", "Verification","Preferred Language Dropdown does not exist in Account Details Page",False
	End If
	VerifyPreferredLanguageDropdown = bPrfdLang
End Function

'[Enter Customer personal details]
Public Function EnterPersonalDetails(strCstmrSegment,strPreferLang,strSalutation,strGender,strStaffIndicator,strStaffID,strDOB,strEducation,strMaritalStatus,strCountryOfBirth,strCityOfBirth,strNationality)
	bEnterPersonalDetails = False
	
	VerifyFieldExistenceInPage HK_OnBoardCustomer_Page.lstCustomerSegement(),"Account Details","Customer Segement"
	SetValue HK_OnBoardCustomer_Page.lstCustomerSegement(),strCstmrSegment,"Customer Segment"
	
	VerifyFieldExistenceInPage HK_OnBoardCustomer_Page.lstPreferredLanguage(),"Account Details","Preferred Language"
	SetValue HK_OnBoardCustomer_Page.lstPreferredLanguage(),strPreferLang,"Preferred Language"
	
	'SetValue HK_OnBoardCustomer_Page.lstSalutation(),strSalutation,"Salutation"
	selectItem_Combobox HK_OnBoardCustomer_Page.lstSalutation(),strSalutation
	
	If UCase(strGender) = "MALE" Then
		SelectRadioButtonGrp "Male", HK_OnBoardCustomer_Page.radioGenderGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioGenderMale(),"Male Radio button"
	Else
		SelectRadioButtonGrp "Female", HK_OnBoardCustomer_Page.radioGenderGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioGenderFemale(),"Female Radio button"
	End If
	
	If UCase(strStaffIndicator) = "YES" Then
		SelectRadioButtonGrp "Yes", HK_OnBoardCustomer_Page.radioStaffIndicatorGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioStaffIndicatorYes(),"Staff Indicator-Yes"
		SetValue HK_OnBoardCustomer_Page.txtStaffID(),strStaffID,"Staff ID"
	Else
		SelectRadioButtonGrp "No", HK_OnBoardCustomer_Page.radioStaffIndicatorGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioStaffIndicatorNo(),"Staff Indicator-No" 
	End If
	
	SetValue HK_OnBoardCustomer_Page.txtDOB(),strDOB,"Date of Birth"
	SetValue HK_OnBoardCustomer_Page.lstEducation(),strEducation,"Education"
	SetValue HK_OnBoardCustomer_Page.lstMaritalStatus(),strMaritalStatus,"Marital Status"
	SetValue HK_OnBoardCustomer_Page.lstCountryOfBirth(),strCountryOfBirth,"Country of Birth"
	SetValue HK_OnBoardCustomer_Page.lstCityOfBirth(),strCityOfBirth,"City of Birth"
	SetValue HK_OnBoardCustomer_Page.lstNationality(),strNationality,"Nationality"
End Function

'[Enter Identification Document details]
Public Function EnterDocumentDetails(strDocType,strDocExprDate,strAltDocIDType,strAltDocIDNo)
	SetValue HK_OnBoardCustomer_Page.txtDocumentExpiryDate(),strDocExprDate,"Document Expiry Date"
	If strAltDocIDType <> "" AND strAltDocIDType <> "BLANK" AND strAltDocIDType <> strDocType Then
		SetValue HK_OnBoardCustomer_Page.lstAlternateIdentificationDocument(),strAltDocIDType,"Alternate Identification Document"
		SetValue HK_OnBoardCustomer_Page.txtAlternateIdentificationDocumentNo(),strAltDocIDNo,"Alternate Identification Document Number"
	End If	
End Function

'[Enter Home Contact details]
Public Function EnterHomeContactDetails(strHomeCntCode,strHomeCity,strHomeNo,strAltHomeCntCode,strAltHomeCity,strAltHomeNo)
	SetValue HK_OnBoardCustomer_Page.txtHomeNumberCode(),strHomeCntCode,"Country Code-Home Number"
	SetValue HK_OnBoardCustomer_Page.txtHomeNumberCity(), strHomeCity,"City-Home Number"
	SetValue HK_OnBoardCustomer_Page.txtHomeNumber(), strHomeNo,"Home Number"
	
	SetValue HK_OnBoardCustomer_Page.txtAltHomeNumberCode(),strAltHomeCntCode,"Country Code-Alt Home Number"
	SetValue HK_OnBoardCustomer_Page.txtAltHomeNumberCity(),strAltHomeCity,"City-Alt Home Number"
	SetValue HK_OnBoardCustomer_Page.txtAltHomeNumber(),strAltHomeNo,"Alt Home Number"
End Function

'[Enter Mobile Contact details]
Public Function EnterMobileContactDetails(strMobileCntCode,strMobileNo,strAltMobileCntCode,strAltMobileNo)
	SetValue HK_OnBoardCustomer_Page.txtMobileNumberCode(),strMobileCntCode,"Country Code-Mobile Number"
	SetValue HK_OnBoardCustomer_Page.txtMobileNumber(),strMobileNo,"Mobile Number"
	
	SetValue HK_OnBoardCustomer_Page.txtAltMobileNumberCode(),strAltMobileCntCode,"Country Code-Alt Mobile Number"
	SetValue HK_OnBoardCustomer_Page.txtAltMobileNumber(),strAltMobileNo,"Alt Mobile Number"
End Function

'[Enter Fax Contact details]
Public Function EnterFaxContactDetails(strFaxCntCode,strFaxCity,strFaxNo)
	SetValue HK_OnBoardCustomer_Page.txtFaxNumberCode(),strFaxCntCode,"Country Code-Fax Number"
	SetValue HK_OnBoardCustomer_Page.txtFaxNumberCity(),strFaxCity,"City-Fax Number"
	SetValue HK_OnBoardCustomer_Page.txtFaxNumber(),strFaxNo,"Fax Number"
End Function

'[Enter Other Contact details]
Public Function EnterOtherContactDetails(strEmail,strYearOfResident,strResidentalStatus,strMonthlyRental)
	SetValue HK_OnBoardCustomer_Page.txtEmail(),strEmail,"Email"
	SetValue HK_OnBoardCustomer_Page.txtYearOfResident(),strYearOfResident,"Years of Resident"
	SetValue HK_OnBoardCustomer_Page.lstResidentStatus(),strResidentalStatus,"Residental Status"
	If UCase(strResidentalStatus) = "RENTED" Then
		SetValue HK_OnBoardCustomer_Page.txtMonthlyRental(),strMonthlyRental,"Monthly Rental"
	End If
End Function

'[Enter Residential Address details]
Public Function EnterResidentialAddress(strCntry,strAddrs1,strAddrs2,strAddrs3)
	If Trim(HK_OnBoardCustomer_Page.lstResidentialAddressCountry().GetROProperty("value")) <> strCntry Then
		SetValue HK_OnBoardCustomer_Page.lstResidentialAddressCountry(),strCntry,"Residential Address Country"
	End If
	SetValue HK_OnBoardCustomer_Page.txtResidentialAddress1(),strAddrs1,"Residental Address Line-1"
	SetValue HK_OnBoardCustomer_Page.txtResidentialAddress2(),strAddrs2,"Residental Address Line-2"
	SetValue HK_OnBoardCustomer_Page.txtResidentialAddress3(),strAddrs3,"Residental Address Line-3"
End Function

'[Enter Permanent Address details]
Public Function EnterPermanentAddress(strSameAsResidentialAddress,strCntry,strAddrs1,strAddrs2,strAddrs3)
	If UCase(strSameAsResidentialAddress) = "YES" Then
		ClickOnObject HK_OnBoardCustomer_Page.chkSameAsResidentialAddress(),"Same As Residential Address-Checkbox"
	Else
		If Trim(HK_OnBoardCustomer_Page.lstPermanentAddressCountry().GetROProperty("value")) <> strCntry Then
			SetValue HK_OnBoardCustomer_Page.lstPermanentAddressCountry(),strCntry,"Permanent Address Country"
		End If
		SetValue HK_OnBoardCustomer_Page.txtPermanentAddress1(),strAddrs1,"Permanent Address Line-1"
		SetValue HK_OnBoardCustomer_Page.txtPermanentAddress2(),strAddrs2,"Permanent Address Line-2"
		SetValue HK_OnBoardCustomer_Page.txtPermanentAddress3(),strAddrs3,"Permanent Address Line-3"
	End If
End Function

'[Enter Current Employment Details]
Public Function EnterCurrentEmploymentDetails(strEmplmentStatus,strOccupation,strEmplrName,strNatureOfBusinessSection,strNatureOfBusiness,strYrsOfService,strMonthsOfService,strPosition,strAnnualIncome)
	SetValue HK_OnBoardCustomer_Page.lstEmploymentStatus(),strEmplmentStatus,"Employement Status"
	SetValue HK_OnBoardCustomer_Page.lstOccupation(),strOccupation,"Occupation"
	SetValue HK_OnBoardCustomer_Page.txtNameOfEmployer(),strEmplrName,"Name of Employer"
	SetValue HK_OnBoardCustomer_Page.lstNatureOfBusinessSection(),strNatureOfBusinessSection,"Nature of Business Section"
	SetValue HK_OnBoardCustomer_Page.lstNatureOfBusiness(),strNatureOfBusiness,"Nature of Business"
	SetValue HK_OnBoardCustomer_Page.txtYearsOfService(),strYrsOfService,"Years of Service"
	SetValue HK_OnBoardCustomer_Page.txtMonthsOfService(),strMonthsOfService,"Months of Service"
	SetValue HK_OnBoardCustomer_Page.lstPosition(),strPosition,"Position"
	SetValue HK_OnBoardCustomer_Page.lstAnnualIncome(),strAnnualIncome,"Annual Income"
End Function

'[Enter Current Employment Details for Staff]
Public Function EnterCurrentEmploymentDetailsStaff(strEmplmentStatus,strOccupation,strEmplrName,strNatureOfBusinessSection,strNatureOfBusiness,strYrsOfService,strMonthsOfService,strPosition,strAnnualIncome)
	SetValue HK_OnBoardCustomer_Page.lstEmploymentStatus(),strEmplmentStatus,"Employement Status"
	SetValue HK_OnBoardCustomer_Page.lstOccupation(),strOccupation,"Occupation"
	SetValue HK_OnBoardCustomer_Page.txtNameOfEmployerStaff(),strEmplrName,"Name of Employer"
	SetValue HK_OnBoardCustomer_Page.lstNatureOfBusinessSection(),strNatureOfBusinessSection,"Nature of Business Section"
	SetValue HK_OnBoardCustomer_Page.lstNatureOfBusiness(),strNatureOfBusiness,"Nature of Business"
	SetValue HK_OnBoardCustomer_Page.txtYearsOfService(),strYrsOfService,"Years of Service"
	SetValue HK_OnBoardCustomer_Page.txtMonthsOfService(),strMonthsOfService,"Months of Service"
	SetValue HK_OnBoardCustomer_Page.lstPosition(),strPosition,"Position"
	SetValue HK_OnBoardCustomer_Page.lstAnnualIncome(),strAnnualIncome,"Annual Income"
End Function

'[Enter Current Employement Contact details]
Public Function EnterCurrentEmployementContactDetails(strOffNoCntry,strOffNoCity,strOffNo,strOffFExtn,strOffFaxCntry,strOffFaxCity,strOffFaxNo)
	SetValue HK_OnBoardCustomer_Page.txtOfficeNumberCode(),strOffNoCntry,"Office Number-Country Code"
	SetValue HK_OnBoardCustomer_Page.txtOfficeNumberCity(),strOffNoCity,"Office Number-City"
	SetValue HK_OnBoardCustomer_Page.txtOfficeNumber(),strOffNo,"Office Number"
	SetValue HK_OnBoardCustomer_Page.txtOfficeExtension(),strOffNo,"Office Extension Number"
	SetValue HK_OnBoardCustomer_Page.txtOfficeFaxNumberCode(),strOffFaxCntry,"Office Fax-Country Code"
	SetValue HK_OnBoardCustomer_Page.txtOfficeFaxNumberCity(),strOffFaxCity,"Office Fax-City"
	SetValue HK_OnBoardCustomer_Page.txtOfficeFaxNumber(),strOffFaxNo,"Office Fax Number"
End Function

'[Enter Current Employement office Address details]
Public Function EnterOfficeAddress(strOffAddrsCntry,strOffAddrs1,strOffAddrs2,strOffAddrs3)
	If Trim(HK_OnBoardCustomer_Page.lstOfficeAddressCountry().GetROProperty("value")) <> strOffAddrsCntry Then
		SetValue HK_OnBoardCustomer_Page.lstOfficeAddressCountry(),strOffAddrsCntry,"Office Address Country"
	End If
	SetValue HK_OnBoardCustomer_Page.txtOfficeAddress1(),strOffAddrs1,"Office Address Line-1"
	SetValue HK_OnBoardCustomer_Page.txtOfficeAddress2(),strOffAddrs2,"Office Address Line-2"
	SetValue HK_OnBoardCustomer_Page.txtOfficeAddress3(),strOffAddrs3,"Office Address Line-3"
End Function

'[Enter Previous Employment Details]
Public Function EnterPreviousEmployment(strPrvsEmplrName,strPrvsNatrOfBusinessSection,strPrvsNatrOfBusiness,strPrvsYearsOfSrvcs,strPrvsMonthsOfSrvs)
	SetValue HK_OnBoardCustomer_Page.txtPreviousNameOfEmployer(),strPrvsEmplrName,"Previous Employer Name"
	SetValue HK_OnBoardCustomer_Page.lstPreviousNatureOfBusinessSection(),strPrvsNatrOfBusinessSection,"Previous Employer Nature Of Business Section"
	SetValue HK_OnBoardCustomer_Page.lstPreviousNatureOfBusiness(),strPrvsNatrOfBusiness,"Previous Employer Nature Of Business"
	SetValue HK_OnBoardCustomer_Page.txtPreviousYearsOfService(),strPrvsYearsOfSrvcs,"Previous Employer Years of Service"
	SetValue HK_OnBoardCustomer_Page.txtPreviousMonthsOfService(),strPrvsMonthsOfSrvs,"Previous Employer Months of Service"
End Function

'[Enter Correspondence Address]
Public Function EnterCorrespondenceAddress(strCorAddrsType,strCorCountry,strCorAddrs1,strCorAddrs2,stCorAddrs3,strStmntCycleDate)
	If UCase(strCorAddrsType) = "OTHER ADDRESS" Then
		SetValue HK_OnBoardCustomer_Page.lstCorrespondenceAddress(),strCorAddrsType,"Correspondence Address Type-"&strCorAddrsType
		If Trim(HK_OnBoardCustomer_Page.lstCorrespondenceAddressCountry().GetROProperty("value")) <> strCorCountry Then
			SetValue HK_OnBoardCustomer_Page.lstCorrespondenceAddressCountry(),strCorCountry,"Correspondence Address Country"
		End If
		SetValue HK_OnBoardCustomer_Page.txtCorrespondenceAddress1(),strCorAddrs1,"Correspondence Address-Line1"
		SetValue HK_OnBoardCustomer_Page.txtCorrespondenceAddress2(),strCorAddrs2,"Correspondence Address-Line2"
		SetValue HK_OnBoardCustomer_Page.txtCorrespondenceAddress3(),stCorAddrs3,"Correspondence Address-Line3"
	Else
		SetValue HK_OnBoardCustomer_Page.lstCorrespondenceAddress(),strCorAddrsType,"Correspondence Address Type-"&strCorAddrsType
	End If
	
	VerifyFieldExistenceInPage HK_OnBoardCustomer_Page.txtStatementCycleDate(),"Accounts Detail","Statement Cycle Date"
	SetValue HK_OnBoardCustomer_Page.txtStatementCycleDate(),strStmntCycleDate,"Statement Cycle Date"
End Function

'[Select staff Relationship]
Public Function SelectStaffRelationship(strStaffRelationship,strStaffName,strStaffRelation)
	If UCase(strStaffRelationship) = "NO" Then
		SelectRadioButtonGrp "No", HK_OnBoardCustomer_Page.radioStaffRelationshipGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioStaffRelationshipNo(),"Staff Relationship-"&strStaffRelationship
	Else
		strText = "Yes, Is an employee or has relatives who are DBS Directors / Employees"
		SelectRadioButtonGrp strText, HK_OnBoardCustomer_Page.radioStaffRelationshipGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioStaffRelationshipYes(),"Staff Relationship-"&strStaffRelationship
		SetValue HK_OnBoardCustomer_Page.txtNameOfRelationshipStaff(),strStaffName,"Name of the relevant director or employee in English/Chinese-"&strStaffName
		SetValue HK_OnBoardCustomer_Page.txtRelationshipWithStaff(),strStaffRelation,"Relationship-"&strStaffRelation
	End If
End Function

'[Select Opt Out Preference]
Public Function SelectOptOutPreference(strChannelTypes,isOptOutFromDirectMarketing)
	If Not IsArray(strChannelTypes) Then
		If UCase(strChannelTypes) = "ALL CHANNELS" Then
			ClickOnObject HK_OnBoardCustomer_Page.chkOptOutChannelAll(),"Opt Out Channels-All channels (including email, mail, SMS, phone)"
		End If
	Else
		For i = 0 To UBound(strChannelTypes) Step 1
			Select Case UCase(strChannelTypes(i))
				Case "SMS"
					ClickOnObject HK_OnBoardCustomer_Page.chkOptOutChannelSMSMMS(),"Opt Out Channels-"&strChannelTypes(i)
				Case "EMAIL"
					ClickOnObject HK_OnBoardCustomer_Page.chkOptOutChannelEmail(),"Opt Out Channels-"&strChannelTypes(i)
			End Select
		Next
	End If
	'[Below code is commented by  - As Direct marketing check box has been removed from application]
'	If UCase(isOptOutFromDirectMarketing) = "YES" Then
'		ClickOnObject HK_OnBoardCustomer_Page.chkOptOutDirectMarketing(),"Opt out from Provision of our Data to Other Persons for Direct Marketing-Transfer to Third Party"
'	End If
End Function

'[Enter FATCA Details]
Public Function EnterFATCA(strFATCACntryCode,strFATCAStatus,strFATCADateOnForm,strFATCAReviewSts,strFATCAReviewDate,strFATCACertType,strFATCATaxIDType,strFATCATaxpayerID)
	SetValue HK_OnBoardCustomer_Page.lstFATCACountryCode(),strFATCACntryCode,"FATCA Country Code"
	SetValue HK_OnBoardCustomer_Page.lstFATCAStatus(),strFATCAStatus,"FATCA Status"
	SetValue HK_OnBoardCustomer_Page.txtFATCADateOnForm(),strFATCADateOnForm,"FATCA Date On Form"
	selectItem_Combobox HK_OnBoardCustomer_Page.lstFATCAReviewStatus(),strFATCAReviewSts
	'SetValue HK_OnBoardCustomer_Page.lstFATCAReviewStatus(),strFATCAReviewSts,"FATCA Review Status"
	SetValue HK_OnBoardCustomer_Page.txtFATCAReviewStatusUpdateDate(),strFATCAReviewDate,"FATCA Review Status Update Date"
	'SetValue HK_OnBoardCustomer_Page.lstFATCACertificationType(),strFATCACertType,"FATCA Certification Type"
	selectItem_Combobox HK_OnBoardCustomer_Page.lstFATCACertificationType(),strFATCACertType
	
	If UCase(strStatus) <> "NO" Then
		SetValue HK_OnBoardCustomer_Page.lstFATCATaxPayerIDType(),strFATCATaxIDType,"FATCA Tax Payer ID Type"
		SetValue HK_OnBoardCustomer_Page.txtFATCATaxPayerID(),strFATCATaxpayerID,"FATCA Tax Payer ID"
	End If
End Function

'[Enter CRS Details]
Public Function EnterCRS(strCRSEntry,strCRSCntryCode,strCRSStatus,strCRSDateOnForm,strCRSReviewSts,strCRSReviewDate,strCRSCertType,strCRSTaxpayerID,strCRSReason,strCRSOthrReason)
	If UCase(strCRSEntry) = "YES" Then
		SelectRadioButtonGrp "Yes", HK_OnBoardCustomer_Page.radiOtherTaxResidencyGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioOtherTaxResidencyYes(),"Other Tax Residency (Other than US) "
		ClickOnObject HK_OnBoardCustomer_Page.btnCRSAdd(),"Other Tax Residency (Other than US)-Add Button"
		For i = 1 To 20 Step 1
			If HK_OnBoardCustomer_Page.lstCRSCountryCode().Exist(1) Then
				SetValue HK_OnBoardCustomer_Page.lstCRSCountryCode(),strCRSCntryCode,"CRS Country Code"
				Exit For
			End If
		Next
		SetValue HK_OnBoardCustomer_Page.lstCRSStatus(),strCRSStatus,"CRS Status"
		SetValue HK_OnBoardCustomer_Page.txtCRSDateOnForm(),strCRSDateOnForm,"CRS Date On Form"
		SetValue HK_OnBoardCustomer_Page.lstCRSReviewStatus(),strCRSReviewSts,"CRS Review Status"
		SetValue HK_OnBoardCustomer_Page.txtCRSReviewStatusUpdateDate(),strCRSReviewDate,"CRS Review Status Update Date"
		SetValue HK_OnBoardCustomer_Page.lstCRSCertificationType(),strCRSCertType,"CRS Certification Type"
		If UCase(strCRSStatus) <> "NO" Then
			If strCRSTaxpayerID <> "" Then
				SetValue HK_OnBoardCustomer_Page.txtCRSTaxPayerID(),strCRSTaxpayerID,"CRS Tax Payer ID"
			Else
				SetValue HK_OnBoardCustomer_Page.lstCRSReason(),strCRSReason,"CRS Reason"
			End If
			If Instr(strReason,"Other") > 0 Then
				SetValue HK_OnBoardCustomer_Page.txtCRSOtherReason(),strCRSOthrReason,"CRS Other Reason"
			End If
		End If
		ClickOnObject HK_OnBoardCustomer_Page.btnCRSSave(),"CRS Save Button"
		Wait(3)
	Else
		SelectRadioButtonGrp "No", HK_OnBoardCustomer_Page.radiOtherTaxResidencyGroup(), ""
		ClickOnObject HK_OnBoardCustomer_Page.radioOtherTaxResidencyNo(),"Other Tax Residency (Other than US) "
	End If
End Function

'[Enter CDD Rating data and RM Data]
Public Function EnterCDDRatingAndRMData(strCDDData,strRMData)
	SetValue HK_OnBoardCustomer_Page.lstCustomerRiskRating(),strCDDData,"Customer Risk Rating"
	SetValue HK_OnBoardCustomer_Page.lstPrimaryRelationshipManagerID(),strRMData,"Primary Relationship Manager ID"
	'selectItem_Combobox HK_OnBoardCustomer_Page.lstCustomerRiskRating(),strCDDData
	'selectItem_Combobox HK_OnBoardCustomer_Page.lstPrimaryRelationshipManagerID(),strRMData
End Function


'[Verify Wealth details Dropdown values in Application Details page]
Public Function VerifyWealthDetails(strSourcesOfWealth)
	bWealthDetls = False
	For i = 0 To UBound(strSourcesOfWealth) Step 1
		Select Case UCase(strSourcesOfWealth(i))
			Case "BUSINESS INCOME"
				bWealthDetls = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfWealthBusinessIncome(),"Application Details-Source of Wealth-",strSourcesOfWealth(i)) 
			Case "SALARY"
				bWealthDetls = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfWealthSalary(),"Application Details-Source of Wealth-",strSourcesOfWealth(i)) 
			Case "RETURN ON INVESTMENTS"
				bWealthDetls = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfWealthReturnOfInvestement(),"Application Details-Source of Wealth-",strSourcesOfWealth(i)) 
			Case "INHERITANCE"
				bWealthDetls = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfWealthInheritance(),"Application Details-Source of Wealth-",strSourcesOfWealth(i)) 
			Case "OTHER"
				bWealthDetls = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSourceOfWealthOther(),"Application Details-Source of Wealth-",strSourcesOfWealth(i)) 
				bWealthDetls = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSourceOfWealthOther(),"Application Details-Source of Wealth-","Other Text Box") 
		End Select
	Next
	VerifyWealthDetails	= bWealthDetls
End Function

'[Verify Estimated Net Worth Dropdown values in Application Details Page]
Public Function VerifyNetWorthDropdown(strNetWorth,strDefaultNet)
	bNetWorth = False
	bNetWorth = VerifyDropdownDefaultValue(HK_OnBoardCustomer_Page.lstEstimatedNetWorth(),strDefaultNet,"Estimated Net Worth")
	For i = 0 To UBound(strNetWorth) Step 1
		bNetWorth = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstEstimatedNetWorth(),strNetWorth(i),"Estimated Net Worth")
	Next
	VerifyNetWorthDropdown = bNetWorth
End Function

'[Select Wealth details]
Public Function EnterWealthDetails(strSourcesOfWealth,strOtherWealthSource,strNetWorth)
	For i = 0 To UBound(strSourcesOfWealth) Step 1
		Select Case UCase(strSourcesOfWealth(i))
			Case "BUSINESS INCOME"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfWealthBusinessIncome(),"Source of Wealth-"&strSourcesOfWealth(i)
			Case "SALARY"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfWealthSalary(),"Source of Wealth-"&strSourcesOfWealth(i)
			Case "RETURN ON INVESTMENTS"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfWealthReturnOfInvestement(),"Source of Wealth-"&strSourcesOfWealth(i)
			Case "INHERITANCE"
				ClickOnObject HK_OnBoardCustomer_Page.chkSourceOfWealthInheritance(),"Source of Wealth-"&strSourcesOfWealth(i)
			Case "OTHER"
				ClickOnObject HK_OnBoabdrdCustomer_Page.chkSourceOfWealthOther(),"Source of Wealth-"&strSourcesOfWealth(i)
				SetValue HK_OnBoardCustomer_Page.txtSourceOfWealthOther(),strOtherWealthSource,"Other Source of Wealth"
		End Select
	Next
	SetValue HK_OnBoardCustomer_Page.lstEstimatedNetWorth(),strNetWorth,"Estimated Net Worth"
End Function

'[Navigate to Account Type Selection Page]
Public Function NavigateToAccountTypeSelectionPage()
	ClickOnObject HK_OnBoardCustomer_Page.eleAccountTypeSelection(),"Account Type Selection"
	WaitForICallLoading
	VerifyFieldRemovedFromPage HK_OnBoardCustomer_Page.txtStatementCycleDate(),"Accounts Type Selection","Statement Cycle Date"
	VerifyFieldRemovedFromPage HK_OnBoardCustomer_Page.eleBranchID(),"Accounts Type Selection","Branch ID"
End Function

'[Saving and Cheque Account Data Setup]
Public Function SavingAccountDataSetup(strSavingAccCurCode,isODPRqd,isCBRqd)
	SetValue HK_OnBoardCustomer_Page.lstSavingChequeAccntCurrencyCode(),strSavingAccCurCode,"Save & Cheque Account Currency Code"
	'[Enter effetive date code]
	
	If UCase(isODPRqd) = "YES" And UCase(strSavingAccCurCode) = "HKD" Then
		If HK_OnBoardCustomer_Page.chkSavingChequeAccntOD().GetROProperty("disabled") <> 1 And HK_OnBoardCustomer_Page.chkSavingChequeAccntOD().GetROProperty("checked") <> 1 Then
			ClickOnObject HK_OnBoardCustomer_Page.chkSavingChequeAccntOD(),"Save & Cheque Account Overdraft Protection"
		End If
	End If
	
	If UCase(isCBRqd) = "YES" Then
		If HK_OnBoardCustomer_Page.chkSavingChequeAccntChequeBook().GetROProperty("disabled") <> 1 And HK_OnBoardCustomer_Page.chkSavingChequeAccntChequeBook().GetROProperty("checked") <> 1 Then
			ClickOnObject HK_OnBoardCustomer_Page.chkSavingChequeAccntChequeBook(),"Save & Cheque Account Cheque Book"
		End If
	End If
End Function

'[Multi Currency Account Data Setup]
Public Function MultiCurrencySavingAccountDataSetup(arrCurrencyCodes)
'	For j = 0 To UBound(arrCurrencyCodes) Step 1
'		SelectCurrencyCode arrCurrencyCodes(j)
'	Next
	SelectCurrencyCodeNew arrCurrencyCodes
End Function

Function SelectCurrencyCodeNewOld(arrCurrencyCodes)
	Set oDesc = Description.Create
	oDesc("micclass").value = "WebCheckBox"
	oDesc("html id").value = "actTypeMC.*"
	Set CheckBoxChildObjects = gObjIServePage.ChildObjects(oDesc)
	strTotalCount = CheckBoxChildObjects.Count-1
	ReDim Preserve arrCurrencyCodes(strTotalCount)
	For i = 0 To strTotalCount Step 1
		If Trim(CheckBoxChildObjects(i).GetROProperty("innertext")) = Trim(arrCurrencyCodes(i)) Then
			If CheckBoxChildObjects(i).GetROProperty("checked") <> 1 Then
				ClickOnObject CheckBoxChildObjects(i), strCurrencyCode & " Check box"
			End If
		Else
			If CheckBoxChildObjects(i).GetROProperty("checked") = 1  Then
				CheckBoxChildObjects(i).Click
			End If
		End If
	Next
End Function

Function SelectCurrencyCodeNew(arrCurrencyCodes)
	Set oDesc = Description.Create
	oDesc("micclass").value = "WebCheckBox"
	oDesc("html id").value = "actTypeMC.*"
	Set CheckBoxChildObjects = gObjIServePage.ChildObjects(oDesc)
	strTotalCount = CheckBoxChildObjects.Count-1
	
	For Iterator = 0 To strTotalCount Step 1
		If CheckBoxChildObjects(Iterator).GetROProperty("checked") = 1  Then
			CheckBoxChildObjects(Iterator).Click
		End If
	Next
	
	For j = 0 To UBound(arrCurrencyCodes) Step 1
		strCurrentVal  = Trim(arrCurrencyCodes(j))
		For i = 0 To strTotalCount Step 1			
			If Trim(CheckBoxChildObjects(i).GetROProperty("innertext")) = strCurrentVal Then
				If CheckBoxChildObjects(i).GetROProperty("checked") <> 1 Then
					ClickOnObject CheckBoxChildObjects(i), "-"&strCurrentVal & " Check box"
				End If
				Exit For
			End If
		Next
	Next
End Function


Function SelectCurrencyCode(strCurrencyCode)
	Set oDesc = Description.Create
	oDesc("micclass").value = "WebCheckBox"
	oDesc("html id").value = "actTypeMC.*"
	Set CheckBoxChildObjects = gObjIServePage.ChildObjects(oDesc)
	
	For i = 0 To CheckBoxChildObjects.Count-1 Step 1
		If Trim(CheckBoxChildObjects(i).GetROProperty("innertext")) = Trim(strCurrencyCode) Then
			If CheckBoxChildObjects(i).GetROProperty("checked") <> 1 Then
				ClickOnObject CheckBoxChildObjects(i), strCurrencyCode & " Check box"
				Exit For
			End If
		End If
	Next
End Function

'[Term Deposit Account Data Setup]
Public Function TermDepositAccountDataSetup(isTDRqd)
	If UCase(isTDRqd) = "NO" And HK_OnBoardCustomer_Page.eleRemoveTimeDepositAccount().Exist(3) Then
		ClickOnObject HK_OnBoardCustomer_Page.eleRemoveTimeDepositAccount(),"Term Deposit Account Removed "
	End If
End Function

'[Credit Card Account Data Setup]
Public Function CreditCardAccountDataSetup(isCCRqd,strRewardScheme,isAppLimit,strCardAddress)
	If UCase(isCCRqd) = "NO" And HK_OnBoardCustomer_Page.eleRemoveCreditCardAccount().Exist(3) Then
		ClickOnObject HK_OnBoardCustomer_Page.eleRemoveCreditCardAccount(),"Credit Card Account Removed"
	Else
		If UCase(strRewardScheme) = "CASH REBATE" Then
			ClickOnObject HK_OnBoardCustomer_Page.radioCashRebate(),"DBS Reward Scheme-Cash Rebate"
		ElseIf UCase(strRewardScheme) = "REDEMPTION REBATE" Then
			ClickOnObject HK_OnBoardCustomer_Page.radioRedemptionRebate(),"DBS Reward Scheme-Redemption Rebate"
		End If
		
		If UCase(isAppLimit) = "YES" Then
			ClickOnObject HK_OnBoardCustomer_Page.radioApproveExcedCreditLimitYES(),"Approve any transaction that would result in credit limit to be exceeded-YES"
		ElseIf UCase(isAppLimit) = "NO" Then
			ClickOnObject HK_OnBoardCustomer_Page.radioApproveExcedCreditLimitNO(),"Approve any transaction that would result in credit limit to be exceeded-NO"
		End If
		
		If UCase(strCardAddress) = "OFFICE" Then
			ClickOnObject HK_OnBoardCustomer_Page.radioCreditCardAddressOffice(),"Credit Card Address-OFFICE"
		ElseIf UCase(strCardAddress) = "RESIDENCE" Then
			ClickOnObject HK_OnBoardCustomer_Page.radioCreditCardAddressResidence(),"Credit Card Address-RESIDENCE"
		End If
	End If
End Function

'[Navigate to ATM Card Page]
Public Function NavigateToATMCardPage()
	ClickOnObject HK_OnBoardCustomer_Page.eleATMCard(),"ATM Card Page"
	WaitForICallLoading
End Function

'[Setup ATM Card details]
Public Function SetupATMCardDetails(strPINSqncNo,strATMCardType,strDBSAtmCardType,strAvssAcc,strAvssAmt,strPrimaryAccNo)
	SetValue HK_OnBoardCustomer_Page.txtATMPINSequenceNo(),strPINSqncNo,"PIN Sequence No"
	SetValue HK_OnBoardCustomer_Page.lstATMCardType(),strATMCardType,"ATM Card Type"
	If UCase(strATMCardType) = "DBS ATM CARD" Then
		SetValue HK_OnBoardCustomer_Page.lstDBSATMCardType(),strDBSAtmCardType,"DBS ATM Card Type"
	ElseIf UCase(strATMCardType) = "DBS OCTOPUS CARD" Then
		SetValue HK_OnBoardCustomer_Page.lstAAVSAccountNumberByOctopusCards(),strAvssAcc,"AAVS Account Number by Octopus Cards"
		SetValue HK_OnBoardCustomer_Page.lstAAVSAmount(),strAvssAmt,"AAVS Amount"
	End If
	SetValue HK_OnBoardCustomer_Page.lstPrimaryAccountNo(),strPrimaryAccNo,"ATM Card-Primary Account No"
	RemoveOtherATMCards
End Function

Public Function RemoveOtherATMCards()
 	Set oRemove = Description.Create
 	oRemove("micclass").Value = "WebButton"
 	oRemove("xpath").Value = "//button[contains(@ng-click,'removeAtmPanel')]"
 	Set oRemoveObj = gObjIServePage.ChildObjects(oRemove)
 	
 	If oRemoveObj.Count > 1 Then
		For i = oRemoveObj.Count-1 To 1 Step -1
 			oRemoveObj(i).Click
 		Next
 	End If
 
 	Set oRemoveObj = Nothing
 	Set oRemove = Nothing
End Function

Public Function RemoveAccounts()
 	Set oRemove = Description.Create
 	oRemove("micclass").Value = "WebButton"
 	oRemove("xpath").Value = "//button[contains(@ng-click,'removeCustomerAccount')]"
 	Set oRemoveObj = gObjIServePage.ChildObjects(oRemove)
 	
 	If oRemoveObj.Count > 1 Then
		For i = oRemoveObj.Count-1 To 2 Step -1
 			oRemoveObj(i).Click
 		Next
 	End If
 	Set oRemoveObj = Nothing
 	Set oRemove = Nothing
End Function

'[Navigate to Phone Banking Page]
Public Function NavigateToPhoneBankingPage()
	ClickOnObject HK_OnBoardCustomer_Page.elePhoneBanking(),"Phone Banking Page"
	WaitForICallLoading
End Function

'[Setup Phone Banking details]
Public Function SetupPhoneBankingDetails(isBranchSetupPB,strPBAcc,strPBCntryCode,strPBCityCode,strPBFaxNo)
	If UCase(isBranchSetupPB) = "YES" Then
		'ClickOnObject HK_OnBoardCustomer_Page.radioBranchSetupPhoneBankingYES(),"Yes Radio Button-Phone Banking Setup By Branch"
		SelectRadioButtonGrp isBranchSetupPB, HK_OnBoardCustomer_Page.radioBranchSetupPhoneBankingGroup(), ""
		SetValue HK_OnBoardCustomer_Page.txtPhoneBankingAccount(),strPBAcc,"Phone Banking Account"
		SetValue HK_OnBoardCustomer_Page.txtPBAccountFaxCountryCode(),strPBCntryCode,"Phone Banking Fax Country Code"
		SetValue HK_OnBoardCustomer_Page.txtPBAccountFaxCity(),strPBCityCode,"Phone Banking Fax City Code"
		SetValue HK_OnBoardCustomer_Page.txtPBAccountFaxNo(),strPBFaxNo,"Phone Banking Fax No"
	Else
		'ClickOnObject HK_OnBoardCustomer_Page.radioBranchSetupPhoneBankingNO(),"No Radio Button-Phone Banking Setup By Branch"
		SelectRadioButtonGrp isBranchSetupPB, HK_OnBoardCustomer_Page.radioBranchSetupPhoneBankingGroup(), ""
	End If
End Function

'[Navigate to iBanking Page]
Public Function NavigateToIBankingPage()
	ClickOnObject HK_OnBoardCustomer_Page.eleIBanking(),"iBanking Page"
	WaitForICallLoading
End Function

'[Setup iBanking details]
Public Function SetupiBankingDetails(isCreatedInGIB,strMailerRefNo,strDeviceSerialNo)
	If UCase(isCreatedInGIB) = "YES" Then
		ClickOnObject HK_OnBoardCustomer_Page.chkCreatedInGIB(),"Created in GIB"
	End If
	SetValue HK_OnBoardCustomer_Page.txtUserMailerRefNo(),strMailerRefNo,"User Mailer Ref No"
	SetValue HK_OnBoardCustomer_Page.txtSecureDeviceSerialNo(),strDeviceSerialNo,"Secure Device Serial No"
End Function

'[Submit Panel]
Public Function SubmitPanel(strCautionList,strEWSS,strCDD,strPEP,strPEPDetails,isDocsVerified,isCRSVerified,isFATCAVerified)
	ClickOnObject HK_OnBoardCustomer_Page.btnSubmitPanel(),"Submit Panel button"
	Wait(2)
	
	'[CDD Check in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnCDDCheck(),"CDD Check"
	SetValue HK_OnBoardCustomer_Page.lstCautionList(),strCautionList,"Caution List"
	SetValue HK_OnBoardCustomer_Page.lstEWSS(),strEWSS,"EWSS"
	SetValue HK_OnBoardCustomer_Page.lstCDD(),strCDD,"CDD"
	SetValue HK_OnBoardCustomer_Page.lstPEP(),strPEP,"PEP"
	'SetValue HK_OnBoardCustomer_Page.lstPEP(),strPEPDetails,"PEP Notes"
	Wait(1)
	
	'[Documents verification in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnDocumentVerification(),"Document Verification"
	SelectRadioButtonGrp isDocsVerified, HK_OnBoardCustomer_Page.radiodGroupDocVerification(), ""
	
	'[CRS verification in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnCRSVerification(),"CRS Verification"
	SelectRadioButtonGrp isCRSVerified, HK_OnBoardCustomer_Page.radiodGroupCRSVerification(), ""
	
	'[FATCA verification in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnFATCAVerification(),"FATCA Verification"
	SelectRadioButtonGrp isFATCAVerified, HK_OnBoardCustomer_Page.radiodGroupFATCAVerification(), ""
	
	ClickOnObject HK_OnBoardCustomer_Page.eleLoggedInUserName(),"Logged in User Name"
	
	ClickOnObject HK_OnBoardCustomer_Page.btnProceedToReview(),"Proceed To Review button"
	WaitForICallLoading
	
End Function

'[Submit Panel for Multiple Customer]
Public Function SubmitPanleMultiCustomer()
	ClickOnObject HK_OnBoardCustomer_Page.btnSubmitPanel(),"Submit Panel button"
	Wait(2)
	
	'[CDD Check in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnCDDCheck(),"CDD Check"
	
	Set oPDesc = Description.Create
	oPDesc("micclass").Value = "WebElement"
	oPDesc("xpath").Value = "//section[@ng-repeat='customer in customerSet']"
	
	Set oPanel = gObjIServePage.ChildObjects(oPDesc)
	
	intSectionCount = oPanel.Count
	Dim arrStatusCheck
	arrStatusCheck = Array("Caution List","EWSS","CDD","PEP")
	For i = 0 To intSectionCount-1 Step 1
		Set oDesc = Description.Create
		oDesc("micclass").Value = "WebEdit"
		oDesc("class").Value = "autocompleteInput.*"
		Set oEdit = oPanel(i).ChildObjects(oDesc)
		intEditCount = oEdit.Count
		ReDim Preserve arrStatusCheck(intEditCount)
		For j = 0 To intEditCount-1 Step 1
			oEdit(j).Click
			If j = 3 Then
				SetValue oEdit(j),"Positive,Approval Attached","Customer:"&i+1 &"-"&arrStatusCheck(j)
				Wait(1)
			Else
				SetValue oEdit(j),"Checked","Customer:"&i+1 &"-"&arrStatusCheck(j)
				Wait(1)
			End If
		Next
	Next
	SelectMultiCustSubmitPanelRadioButtons
	ClickOnObject HK_OnBoardCustomer_Page.btnProceedToReview(),"Proceed To Review button"
	WaitForICallLoading
	
End Function

Public Function SelectMultiCustSubmitPanelRadioButtons()

	'[Documents verification in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnDocumentVerification(),"Document Verification"
	SelectRadioVerification "//*[@ng-model='customer.customerDDDocDetails.documentChecked']"
	
	'[CRS verification in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnCRSVerification(),"CRS Verification"
	SelectRadioVerification "//*[@ng-model='customer.customerDDDocDetails.crsChecked']"
	
	'[FATCA verification in Submit Panel]
	ClickOnObject HK_OnBoardCustomer_Page.btnFATCAVerification(),"FATCA Verification"
	SelectRadioVerification "//*[@ng-model='customer.customerDDDocDetails.selfCertifChecked']"
	
End Function

Public Function SelectRadioVerification(strXpath)
	Wait(2)
	Set oDesc = Description.Create
	oDesc("xpath").Value = strXpath
	oDesc("micclass").Value = "WebElement"
	Set oDocObj = gObjIServePage.ChildObjects(oDesc)
	intRadioCnt = oDocObj.Count
	For i = 0 To intRadioCnt-1 Step 1
		SelectRadioButtonGrp "Yes", oDocObj(i), ""
	Next
End Function

'[Submit Onboarding Application]
Public Function SubmitApplication()
	ClickOnObject HK_OnBoardCustomer_Page.btnSubmit(),"Submit Application button"
	If HK_OnBoardCustomer_Page.elePopupAfterSubmit().Exist(5) Then
		ClickOnObject HK_OnBoardCustomer_Page.btnPopUpConfirmOk(),"Popup Submit Confirmation OK Button"
	End If
	WaitForICallLoading
End Function

'[Verify Application submission confirmation]
Public Function VerifyApplicationSubmission(strEntitySetup)
	For i = 1 To 30 Step 1
		If HK_OnBoardCustomer_Page.elePopupAfterSubmit().Exist(1) Then
			strSubmmsionMsg = Trim(HK_OnBoardCustomer_Page.elePopupAfterSubmit().GetROProperty("innertext"))
			If  strSubmmsionMsg = "Application submitted successfully." Then
				ClickOnObject HK_OnBoardCustomer_Page.btnPopUpConfirmOk(),"Popup-After Submit Confirmation OK Button"
	 			LogMessage "RSLT","Verification","Onboarding Application Submitted Successfully while " &strEntitySetup& " selected",True	
			Else
				LogMessage "WARN","Verification","Unable to submit Onboarding Application while "&strEntitySetup& " selected. ERROR: "&strSubmmsionMsg ,False     
			End If
			Exit For
		Else
			Wait(1)
		End If
	Next
End Function

Public Function SelectRadio(strVerificationType,strRadioText)
	Set oDesc = Description.Create
	'oDesc("micclass").Value = "WebRadioGroup"
	oDesc("micclass").Value = "WebMenu"
	oDesc("html id").Value = "radio.*"
	Set CheckBoxChildObjects = gObjIServePage.ChildObjects(oDesc)
	For i = 0 To CheckBoxChildObjects.Count-1 Step 1
		If Trim(CheckBoxChildObjects(i).GetROProperty("innertext")) = Trim(strRadioText) Then
			CheckBoxChildObjects(i).Select(strRadioText)
			'ClickOnObject CheckBoxChildObjects(i), strVerificationType & strRadioText & " Radio button selected "
			Exit For
		End If 
	Next
End Function

'[Submit Customer onboarding form]
Public Function SubmitOnboardingForm()
	ClickOnObject HK_OnBoardCustomer_Page.btnProceedToReview(),"Review button"
	ClickOnObject HK_OnBoardCustomer_Page.btnSubmit(),"Submit button"
	If HK_OnBoardCustomer_Page.elePopupAfterSubmit().Exist(5) Then
		ClickOnObject HK_OnBoardCustomer_Page.btnPopUpConfirmOk(),"Ok button on application submission pop up"
	End If
End Function

'[Verify Account Name by selecting different entity setup]
Public Function CheckAccountNameForDifferntEntitySetup(arrEntitySetups)
	'arrEntitySetups = Split("Liquidator|Receiver|Administrator|Executor|Guardian|Personal Representatives", "|")
	
	stAppName = Trim(HK_OnBoardCustomer_Page.txtAccountName().GetROProperty("value"))
	IsDisabled = Trim(HK_OnBoardCustomer_Page.txtAccountName().GetROProperty("disabled"))

	If  stAppName = Environment.Value("ApplicantName") And IsDisabled = 1 Then
		LogMessage "RSLT","Verification","Account Name is same as Applicant Name and its read only as expected: No Enitiy Setup value selected.",True
	Else
		LogMessage "RSLT","Verification","Account Name is not same as Applicant Name and its editable: No Enitiy Setup value selected.",False
	End If
	HK_OnBoardCustomer_Page.btnPrevious().Click
	
	For i = 0 To UBound(arrEntitySetups) Step 1
	
		SetValue HK_OnBoardCustomer_Page.lstEntitySetup(),arrEntitySetups(i),"Entity Setup-"
		HK_OnBoardCustomer_Page.btnNext().Click
		WaitForICallLoading
		HK_OnBoardCustomer_Page.eleNameInApplicantDetails().Click
		For j = 1 To 20 Step 1
			If HK_OnBoardCustomer_Page.dlgEWSSCheckPopup().Exist(1) Then
				HK_OnBoardCustomer_Page.dlgEWSSCheckPopupCancel().Click
				Exit For
			End If
		Next
		WaitForICallLoading
		HK_OnBoardCustomer_Page.eleAccountTypeSelection().Click
		WaitForICallLoading
	
		stAppName = Trim(HK_OnBoardCustomer_Page.txtAccountName().GetROProperty("value"))
		IsDisabled = Trim(HK_OnBoardCustomer_Page.txtAccountName().GetROProperty("disabled"))
		
		If  stAppName = Environment.Value("ApplicantName") And IsDisabled = 0 Then
			LogMessage "RSLT","Verification","Account Name is same as Applicant Name and its editable as expected: " &arrEntitySetups(i) &" Entity Setup value selected.",True
		Else
			LogMessage "WARN","Verification","Account Name is not same as Applicant Name and its read only: " &arrEntitySetups(i) &" Entity Setup value selected.",False
		End If
		
		HK_OnBoardCustomer_Page.btnPrevious().Click
		WaitForICallLoading	
	Next
End Function

'[Approve Onboarded Customer]
Public Function ApproveCustomer(strEntityType,strApplicationDate)
	WaitForICallLoading
	strApplicationDate = ConvertTodaysDateFormat
	SetValue HK_CustomerSearch_Page.txtApplicationDate(),strApplicationDate,"Application Date"
	ClickOnObject HK_CustomerSearch_Page.btnFilter(),"Filter button"
	strApplicant = "Applicant Name:"&Environment.Value("ApplicantName")
	lstData = Split(strApplicant,"|")
	blnApplicant = selectTableLink(HK_OnBoardCustomer_Page.tblApplicationsHeader(),HK_OnBoardCustomer_Page.tblApplicationsContent(),lstData,"Application No" ,"Application Ref. No.",false,false,false,false)
	WaitForICallLoading	
	ClickOnObject HK_OnBoardCustomer_Page.btnApprove(),"Approve button"
	If HK_OnBoardCustomer_Page.eleApprovalPopUp().Exist(10) Then
		ClickOnObject HK_OnBoardCustomer_Page.btnApproeOkButtonInPopUp(),"Ok button in Approval pop up"
	End If
End Function

'[Reject Onboarded Customer]
Public Function RejectCustomer(strApplicationDate,strReason)
	WaitForICallLoading
	strApplicationDate = ConvertTodaysDateFormat
	SetValue HK_CustomerSearch_Page.txtApplicationDate(),strApplicationDate,"Application Date"
	ClickOnObject HK_CustomerSearch_Page.btnFilter(),"Filter button"
	strApplicant = "Applicant Name:"&Environment.Value("ApplicantName")
	lstData = Split(strApplicant,"|")
	blnApplicant = selectTableLink(HK_OnBoardCustomer_Page.tblApplicationsHeader(),HK_OnBoardCustomer_Page.tblApplicationsContent(),lstData,"Application No" ,"Application Ref. No.",false,false,false,false)
	WaitForICallLoading	
	ClickOnObject HK_OnBoardCustomer_Page.btnReject(),"Reject button"
	If HK_OnBoardCustomer_Page.eleRejectPopUp().Exist(10) Then
		SetValue HK_CustomerSearch_Page.txtRejectReason(),strReason,"Reject Reason"
		ClickOnObject HK_OnBoardCustomer_Page.btnApproeOkButtonInPopUp(),"Ok button in Reject pop up"
		WaitForICallLoading
		ClickOnObject HK_OnBoardCustomer_Page.btnApproeOkButtonInPopUp(),"Ok button in Rejection Successfull message"
	End If
End Function

'[Navigate to First Customer Applicant Detail page]
Public Function NavigateFirst()
	ClickOnObject HK_OnBoardCustomer_Page.eleNameInApplicantDetails(),"First Customer Name Under Applicant Details"
	SyncCustomerPageNavigation
End Function

'[Navigate to Second Customer Applicant Detail page]
Public Function NavigateSecond()
	ClickOnObject HK_OnBoardCustomer_Page.eleSecondNameInApplicantDetails(),"Second Customer Name Under Applicant Details"
	SyncCustomerPageNavigation
End Function

Public Function SyncCustomerPageNavigation()
	For i = 1 To 20 Step 1
		If HK_OnBoardCustomer_Page.dlgEWSSCheckPopup().Exist(1) Then
			ClickOnObject HK_OnBoardCustomer_Page.dlgEWSSCheckPopupCancel(),"Cancel Button in EWSS Pop up"
			Exit For
		End If
	Next
	WaitForICallLoading
End Function

'[Enter Applicant Personal Info]
Public Function EnterApplicantDetailsSection0(strCstmrSegment,strPreferLang,strSalutation,strGender,strStaffIndicator,strStaffID,strDOB,strEducation,strMaritalStatus,strCountryOfBirth)
	EnterPersonalDetails strCstmrSegment,strPreferLang,strSalutation,strGender,strStaffIndicator,strStaffID,strDOB,strEducation,strMaritalStatus,strCountryOfBirth,strCountryOfBirth,strCountryOfBirth
End Function

'[Enter Applicant Document Info]
Public Function EnterApplicantDetailsSection1(strDocType,strDocExprDate,strAltDocIDType,strAltDocIDNo)
	EnterDocumentDetails strDocType,strDocExprDate,strAltDocIDType,strAltDocIDNo
End Function

'[Enter Applicant Mobile]
Public Function EnterApplicantDetailsSection2(strMobileCntCode,strMobileNo,strAltMobileCntCode,strAltMobileNo)
	EnterMobileContactDetails strMobileCntCode,strMobileNo,strAltMobileCntCode,strAltMobileNo
End Function

'[Enter Applicant Residential Address]
Public Function EnterApplicantDetailsSection3(strRCntry,strRAddrs1,strRAddrs2,strRAddrs3)	
	EnterResidentialAddress strRCntry,strRAddrs1,strRAddrs2,strRAddrs3
	EnterPermanentAddress "Yes",Null,Null,Null,Null
End Function

'[Enter Applicant Current Employment]
Public Function EnterApplicantDetailsSection4(strEmplmentStatus,strOccupation,strEmplrName,strNatureOfBusinessSection,strNatureOfBusiness,strYrsOfService,strMonthsOfService,strPosition,strAnnualIncome)
	EnterCurrentEmploymentDetails strEmplmentStatus,strOccupation,strEmplrName,strNatureOfBusinessSection,strNatureOfBusiness,strYrsOfService,strMonthsOfService,strPosition,strAnnualIncome
End Function

'[Enter Applicant Correspondence Address]
Public Function EnterApplicantDetailsSection5(strCorAddrsType,strStmntCycleDate)
	EnterCorrespondenceAddress strCorAddrsType,Null,Null,Null,Null,strStmntCycleDate
End Function

'[Enter Applicant Staff Relation]
Public Function EnterApplicantDetailsSection6(strStaffRelationship)
	SelectStaffRelationship strStaffRelationship,Null,Null
End Function

'[Enter Applicant FATCA]
Public Function EnterApplicantDetailsSection7(strFATCACntryCode,strFATCAStatus,strFATCADateOnForm,strFATCAReviewSts,strFATCAReviewDate,strFATCACertType,strFATCATaxIDType,strFATCATaxpayerID)
	EnterFATCA strFATCACntryCode,strFATCAStatus,strFATCADateOnForm,strFATCAReviewSts,strFATCAReviewDate,strFATCACertType,strFATCATaxIDType,strFATCATaxpayerID
End Function

'[Enter Applicant CRS]
Public Function EnterApplicantDetailsSection8(strCRSEntry,strCRSCntryCode,strCRSStatus,strCRSDateOnForm,strCRSReviewSts,strCRSReviewDate,strCRSCertType,strCRSTaxpayerID,strCRSReason,strCRSOthrReason)
	EnterCRS strCRSEntry,strCRSCntryCode,strCRSStatus,strCRSDateOnForm,strCRSReviewSts,strCRSReviewDate,strCRSCertType,strCRSTaxpayerID,strCRSReason,strCRSOthrReason
End Function

'[Enter Applicant Wealth Details]
Public Function EnterApplicantDetailsSection9(strCDDData,strRMData,strSourcesOfWealth,strOtherWealthSource,strNetWorth)
	EnterCDDRatingAndRMData strCDDData,strRMData
	EnterWealthDetails strSourcesOfWealth,strOtherWealthSource,strNetWorth
End Function

'[Verify Home Number copied from Primary applicant to Trust Minor]
Public Function HomeNumberCopy(strHomeCntCode,strHomeCity,strHomeNo)
	bHomeNo = False
	bHomeNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtHomeNumberCode(),strHomeCntCode,"Country Code-Home Number-Trust Minor")
	bHomeNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtHomeNumberCity(),strHomeCity,"City-Home Number-Trust Minor")
	bHomeNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtHomeNumber(),strHomeNo,"Home Number-Trust Minor")
	HomeNumberCopy = bHomeNo
End Function

'[Verify Alternate Home Number copied from Primary applicant to Trust Minor]
Public Function AltHomeNumberCopy(strAltHomeCntCode,strAltHomeCity,strAltHomeNo)
	bAltHomeNo = False
	bAltHomeNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtAltHomeNumberCode(),strAltHomeCntCode,"Country Code-Alt Home Number-Trust Minor")
	bAltHomeNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtAltHomeNumberCity(),strAltHomeCity,"City-Alt Home Number-Trust Minor")
	bAltHomeNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtAltHomeNumber(),strAltHomeNo,"Alt Home Number-Trust Minor")
	AltHomeNumberCopy = bAltHomeNo
End Function

'[Verify Mobile Number copied from Primary applicant to Trust Minor]
Public Function MobileNumberCopy(strMobileCntCode,strMobileNo,strAltMobileCntCode,strAltMobileNo)
	bMobileNo = False
	
	bMobileNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtMobileNumberCode(),strMobileCntCode,"Country Code- Mobile Number-Trust Minor")
	
	If HK_OnBoardCustomer_Page.txtMobileNumber().GetROProperty("value") <> strMobileNo Then
		LogMessage "RSLT","Verification","As Expected, Mobile Number is not copied from Primary to Trust Minor Applicant",True
	Else
		LogMessage "WARN","Verification","Failed, Mobile Number is copied from Primary to Trust Minor Applicant",False
	End If
	
	bMobileNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtAltMobileNumberCode(),strAltMobileCntCode,"Country Code-Alt Mobile Number-Trust Minor")
	bMobileNo = verifyFieldValue(HK_OnBoardCustomer_Page.txtAltMobileNumber(),strAltMobileNo,"Alt Mobile Number-Trust Minor")
	
	MobileNumberCopy = bMobileNo
End Function

'[Verify Fax Number copied from Primary applicant to Trust Minor]
Public Function VerifyFaxContactDetailsCopy(strFaxCntCode,strFaxCity,strFaxNo)
	bFax = False
	bFax = verifyFieldValue(HK_OnBoardCustomer_Page.txtFaxNumberCode(),strFaxCntCode,"Country Code-Fax Number-Trust Minor")
	bFax = verifyFieldValue(HK_OnBoardCustomer_Page.txtFaxNumberCity(),strFaxCity,"City-Fax Number-Trust Minor")
	bFax = verifyFieldValue(HK_OnBoardCustomer_Page.txtFaxNumber(),strFaxNo,"Fax Number-Trust Minor")
	VerifyFaxContactDetailsCopy = bFax
End Function

'[Verify Residential Address copied from Primary applicant to Trust Minor]
Public Function VerifyResidentialAddressCopy(strCntry,strAddrs1,strAddrs2,strAddrs3)
	bResAdd = False
	bResAdd = verifyFieldValue(HK_OnBoardCustomer_Page.lstResidentialAddressCountry(),strCntry,"Residential Address Country-Trust Minor")
	bResAdd = verifyFieldValue(HK_OnBoardCustomer_Page.txtResidentialAddress1(),strAddrs1,"Residental Address Line-1-Trust Minor")
	bResAdd = verifyFieldValue(HK_OnBoardCustomer_Page.txtResidentialAddress2(),strAddrs2,"Residental Address Line-2-Trust Minor")
	bResAdd = verifyFieldValue(HK_OnBoardCustomer_Page.txtResidentialAddress3(),strAddrs3,"Residental Address Line-3-Trust Minor")
	VerifyResidentialAddressCopy = bResAdd
End Function

'[Verify Trust Minor DOB greater then 21 years]
Public Function VerifyTrustMinorDOB(strDOB,strMessg)
	bTRDOB = False
	SetValue HK_OnBoardCustomer_Page.txtDOB(),strDOB,"Date of Birth"
	HK_OnBoardCustomer_Page.txtHomeNumberCode().Click
	bTRDOB = verifyInnerText(HK_OnBoardCustomer_Page.eleDOBErrror(),strMessg,"Trust Minor Age > 21 Years")
	VerifyTrustMinorDOB = bTRDOB
End Function

'[Click on Proceed to Preview Button]
Public Function ClickOnProceedToPreviewButton()
	ClickOnObject HK_OnBoardCustomer_Page.btnProceedToReview(),"Proceed To Review button"
	WaitForICallLoading
End Function

'[Edit Currency Details in Account Type Selection Page]
Public Function VerifyValueAfterEdit(arrEditCurrencyCodes)
	bEditVal = False
	strValBeforeEdit = Trim(HK_OnBoardCustomer_Page.eleCurrencyValInPreviewPage().GetROProperty("innertext"))
	
	ClickOnObject HK_OnBoardCustomer_Page.btnEditAccntTypeSelectionInPreviewPage(),"Edit Button-Account Type Selection-Preview Page"
 	WaitForICallLoading
 	If HK_OnBoardCustomer_Page.eleaccntTypeSelectionlbl().Exist(3) Then
 		LogMessage "RSLT","Verification","As Expected, On Clicking Edit Button For Account Type Selection section in Preview Page,Account Type Selection page is displayed.",True
	Else
		LogMessage "WARN","Verification","Failed, On Clicking Edit Button For Account Type Selection section in Preview Page,Account Type Selection page is not displayed.",False
 	End If
 	
 	SelectCurrencyCodeNew arrEditCurrencyCodes
 	
 	NavigateToATMCardPage
 	NavigateToPhoneBankingPage
 	NavigateToIBankingPage
 	
	ClickOnObject HK_OnBoardCustomer_Page.btnProceedToReview(),"Proceed To Review button"
	WaitForICallLoading
	
	strValAfterEdit = Trim(HK_OnBoardCustomer_Page.eleCurrencyValInPreviewPage().GetROProperty("innertext"))
	
	If strValBeforeEdit <> strValAfterEdit Then
		LogMessage "RSLT","Verification","As Expected, After Editing Currency Values in Account Type Selection , Updated Currency values is displayed in Review page as expected",True
		bEditVal = True
	Else
		LogMessage "WARN","Verification","Failed, After Editing Currency Values in Account Type Selection , Updated Currency value is not displayed in Preview page.",False
	End If
 	VerifyValueAfterEdit = bEditVal
End Function 

'[Verify Veiw Profile Tab and Left Menu links for ETB Customer]
Public Function VerifyViewProfileTabForETB(strTabName,lstLeftMenuName)
	bTabVerify = False
	bTabVerify = verifyTabExist(strTabName)
	bTabVerify = selectTab(strTabName)
	
	For i = 0 To UBound(lstLeftMenuName) Step 1
		clickLefmenuLink lstLeftMenuName(i)
	Next
	VerifyViewProfileTabForETB = bTabVerify
End Function

'[Verify Combined Statement and Statement Cycle date in Bank Data Page]
Public Function VerifyCombinedStatement(strLinkName)
	clickLefmenuLink strLinkName
	WaitForICallLoading
	bCombinedState = False
	bCombinedState = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleVPCombineStmntLbl(),"ETB Customer-View Profile-Bank Data","Combined Statement Indicator Label")
	bCombinedState = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleVPStatementCycleDateLbl(),"ETB Customer-View Profile-Bank Data","Statement Cycle Date Indicator Label")
	bCombinedState = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleVPSuspendedLbl(),"ETB Customer-View Profile-Bank Data","Suspended Indicator Label")
	VerifyCombinedStatement = bCombinedState
End Function

'[Verify all fields are read only or not in page]
Public Function VerifyReadOnlyFiledsInPage(strLinkName)
	bVerifyReadPage = False
	clickLefmenuLink strLinkName
	WaitForICallLoading
	bVerifyReadPage = VerifyReadOnly
	If bVerifyReadPage Then
		LogMessage "RSLT","Verification","As Expected, All fields in "&strLinkName&" page are read only.",True
		bVerifyReadPage = True
	Else
		LogMessage "WARN","Verification","Failed, Some fields in "&strLinkName&" page are not read only.",False
	End If
	VerifyReadOnlyFiledsInPage = bVerifyReadPage
End Function

Public Function VerifyReadOnly()
	bVerifyReadOnly = True
	Set oDesc = Description.Create
	oDesc("micclass").Value = "WebElement"
	
	Set EleObj = gObjIServePage.WebElement("xpath:=//*[@ui-view='customerProfile']//md-content").ChildObjects(oDesc)
	
	intEleCnt = EleObj.Count
	For i = 0 To intEleCnt-1 Step 1
		If EleObj(i).GetROProperty("disabled",1) Then
			bVerifyReadOnly = False
			Print bVerifyReadOnly
		End If
	Next
	VerifyReadOnly = bVerifyReadOnly
End Function

'[Verify Default Account Types in Account Selection Page]
Public Function VerifyDefaultAccountTypes()
	bAccountType = False
	bAccountType = AccountTypeVerification(HK_OnBoardCustomer_Page.eleSaveChequeAccountType(), "Save & Cheque")
	bAccountType = AccountTypeVerification(HK_OnBoardCustomer_Page.eleMultiCurrencyAccountType(), "Multi-Currency Savings Account")
	bAccountType = AccountTypeVerification(HK_OnBoardCustomer_Page.eleTimeDepositAccountType(), "Time Deposit")
	bAccountType = AccountTypeVerification(HK_OnBoardCustomer_Page.eleCreditCardAccountType(), "Credit Card")
	VerifyDefaultAccountTypes = bAccountType
End Function

Public Function AccountTypeVerification(ojbAccount, strAccountType)
	bAccMsg = False
	If ojbAccount.Exist(5) Then
		LogMessage "RSLT","Verification","As Expected, " &strAccountType& " is displayed as Default Account Type",True
		bAccMsg = True
	Else
		LogMessage "WARN","Verification","Failed, "&strAccountType& " is not displayed as Default Account Type",False
	End If
	AccountTypeVerification = bAccMsg
End Function

'[Verify No of Account Types User can choose in Account Type Selection Page]
Public Function VerifyNoOfAccountTypes(intNoOfAccounts)
	
	ClickOnObject HK_OnBoardCustomer_Page.btnAddAccount(),"Add Account button"
	
	bNoOfAccntTypes = False
	Set oDesc = Description.Create
	oDesc("xpath").Value = "//*[contains(@id,'accTypeSelAddAccMenu')]/span"
	Set ObjChild = gObjIServePage.ChildObjects(oDesc)
	intAcctypeCnt = ObjChild.Count
	
	strAllAcct = ""
	For i = 0 To intAcctypeCnt-1 Step 1
		strAcct = Trim(ObjChild(i).GetROProperty("innertext"))
		strAllAcct = strAllAcct & strAcct & ", "
	Next
	strAllAcct  = Mid(strAllAcct,1)
	
	If intAcctypeCnt = Cint(intNoOfAccounts) Then
		LogMessage "RSLT","Verification","As Expected, User can choose "&intNoOfAccounts&" Types of account and Those are " &strAllAcct,True
		bNoOfAccntTypes = True
	Else
		LogMessage "WARN","Verification","Failed, User can choose only "&intAcctypeCnt&" Types of account and Those are " &strAllAcct& "Expected: "&intNoOfAccounts ,False
	End If
	VerifyNoOfAccountTypes = bNoOfAccntTypes
End Function

'[Verify Account Types can be Added]
Public Function AddAccountTypes()
	bAdd = False
	ClickOnObject HK_OnBoardCustomer_Page.btnAddAccount(),"Add Account button"
	ClickOnObject HK_OnBoardCustomer_Page.btnSavingAccount(),"Saving Account"
	
	If  HK_OnBoardCustomer_Page.eleSavingAccountType().Exist(4) Then
		LogMessage "RSLT","Verification","As Expected, User can Add more accounts",True
		bAdd = True
	Else
		LogMessage "WARN","Verification","Failed, User cannot add more accounts" ,False
	End If
	AddAccountTypes = bAdd
End Function

'[Verify Account Types can be Removed]
Public Function RemoveAccountTypes()
	bRemove = False
	
	ClickOnObject HK_OnBoardCustomer_Page.eleRemoveSavingAccount(),"Saving Account Remove Button"
	
	If Not HK_OnBoardCustomer_Page.eleSavingAccountType().Exist(4) Then
		LogMessage "RSLT","Verification","As Expected, User can Remove accounts",True
		bRemove = True
	Else
		LogMessage "WARN","Verification","Failed, User cannot remove accounts" ,False
	End If
	RemoveAccountTypes = bRemove
End Function

'[Verify one Current account is mandatory]
Public Function VerifyCurrentAccntMandatory(strMessage)
	bCurrentAccnt = False
	ClickOnObject HK_OnBoardCustomer_Page.eleRemoveSaveChequeAccount(),"Save & Cheque Account Remove Button"

	If HK_OnBoardCustomer_Page.eleDeleteAccountErrorPopUp().Exist(4) Then
		verifyInnerText HK_OnBoardCustomer_Page.eleDeleteAccountErrorPopUp(),strMessage,"Account Type Delete Error Message"
		HK_OnBoardCustomer_Page.btnOkErrorPopup().Click
		LogMessage "RSLT","Verification","As Expected, At least one Current Account is Mandatory for Onboarding a customer",True
		bCurrentAccnt = True
	Else
		LogMessage "WARN","Verification","Failed, User can be onboarded Without any Current Account. Expected: At least one Current Account is mandatory" ,False
	End If
	VerifyCurrentAccntMandatory = bCurrentAccnt
End Function

'[Verify one Multi Currency account is mandatory]
Public Function VerifyMultiCurrAccntMandatory(strMessage)
	bMultiCurrAccnt = False
	ClickOnObject HK_OnBoardCustomer_Page.eleRemoveMultiCurrencyAccount(),"Multi Currency Account Remove Button"

	If HK_OnBoardCustomer_Page.eleDeleteAccountErrorPopUp().Exist(4) Then
		verifyInnerText HK_OnBoardCustomer_Page.eleDeleteAccountErrorPopUp(),strMessage,"Account Type Delete Error Message"
		HK_OnBoardCustomer_Page.btnOkErrorPopup().Click
		LogMessage "RSLT","Verification","As Expected, At least one Multi Currency Account is Mandatory for Onboarding a customer",True
		bMultiCurrAccnt = True
	Else
		LogMessage "WARN","Verification","Failed, User can be onboarded Without any Multi Currency Account. Expected: At least one Multi Currency is mandatory" ,False
	End If
	VerifyMultiCurrAccntMandatory = bMultiCurrAccnt
	
End Function

'[Verify Save and Cheque Account Type Fields and Default Values]
Public Function VerifySaveChequeFieldsAndValue(strSchemeCode,arrCurrencies)
	bSaveChequeVerify = False
	strAccountType = "Save & Cheque"
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSaveChequeAccountType(),strAccountType,"Account Type")
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSavingChequeAccntEffectiveDate(),strAccountType,"Effective Date")
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSavingChequeAccntSchemeCode(),strAccountType,"Scheme Code")
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSavingChequeAccntOD(),strAccountType,"Over Draft Protection")
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSavingChequeAccntChequeBook(),strAccountType,"Cheque Book")
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.lstSavingChequeAccntCurrencyCode(),strAccountType,"Currency Code")
	bSaveChequeVerify = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSavingChequeAccountNumber(),strAccountType,"Account Number")
	
	bSaveChequeVerify = verifyEffectiveDefaultDate(HK_OnBoardCustomer_Page.txtSavingChequeAccntEffectiveDate(),strAccountType&"-Effective Date")
	
	bSaveChequeVerify = verifyInnerText(HK_OnBoardCustomer_Page.eleSaveChequeAccountType(),strAccountType,"-Account Type")
	bSaveChequeVerify = verifyInnerText(HK_OnBoardCustomer_Page.eleSavingChequeAccntSchemeCode(),strSchemeCode,strAccountType&"-Scheme Code")
	bSaveChequeVerify = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstSavingChequeAccntCurrencyCode(),arrCurrencies,strAccountType&"-Currency Code")
	VerifySaveChequeFieldsAndValue = bSaveChequeVerify
End Function

'[Verify Multi Currency Savings Account Type Fields and Default Values]
Public Function VerifyMultiCurrencyFieldsAndValue(strSchemeCode,arrDefultCheckedCurr,arrAllCurrencies)
	bMulti = False
	strAccountType = "Multi-Currency Savings Account"
	bMulti = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleMultiCurrencyAccountType(),strAccountType,"Account Type")
	bMulti = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtMultiCurrencyAccntEffectiveDate(),strAccountType,"Effective Date")
	bMulti = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleMultiCurrSchemeCode(),strAccountType,"Scheme Code")
	
	bMulti = verifyEffectiveDefaultDate(HK_OnBoardCustomer_Page.txtMultiCurrencyAccntEffectiveDate(),strAccountType&"-Effective Date")
	
	bMulti = verifyInnerText(HK_OnBoardCustomer_Page.eleMultiCurrencyAccountType(),strAccountType,strAccountType&"-Account Type")
	bMulti = verifyInnerText(HK_OnBoardCustomer_Page.eleMultiCurrSchemeCode(),strSchemeCode,strAccountType&"-Scheme Code")
	
	Set oDesc = Description.Create
	oDesc("micclass").value = "WebCheckBox"
	oDesc("html id").value = "actTypeMC.*"
	Set CheckBoxChildObjects = gObjIServePage.ChildObjects(oDesc)
	strTotalCount = CheckBoxChildObjects.Count-1
	strAllCurrency = ""
	
	For i = 0 To strTotalCount Step 1
		strAllCurrency = strAllCurrency & Trim(CheckBoxChildObjects(i).GetROProperty("innertext")) & "#"
	Next
	
	strAllCurrency = Left(strAllCurrency,Len(strAllCurrency)-1)
	
	If strAllCurrency = arrAllCurrencies Then
		LogMessage "RSLT","Verification","As Expected, "&arrAllCurrencies&" are the currencies for Multi Currency Saving Account",True
		bMulti = True
	Else
		LogMessage "WARN","Verification","Failed, "&strAllCurrency&" are the currencies for Multi Currency Saving Account, Expected: "&arrAllCurrencies ,False
		bMulti = False
	End If
	
	strTemp = ""
	For j = 0 To UBound(arrDefultCheckedCurr) Step 1
		strCurrentVal  = Trim(arrDefultCheckedCurr(j))
		For i = 0 To strTotalCount Step 1			
			If Trim(CheckBoxChildObjects(i).GetROProperty("innertext")) = strCurrentVal Then
				If CheckBoxChildObjects(i).GetROProperty("checked") = 1 Then
					strTemp = strTemp &  arrDefultCheckedCurr(j) & " Currency Checked by default" & ","
					bCurr = True
				Else
					strTemp = strTemp & arrDefultCheckedCurr(j) & " Curency not Checked by default" & ","
					bCurr = False
				End If
				Exit For
			End If
		Next
	Next
	
	If bCurr Then
		LogMessage "RSLT","Verification","As Expected, "&strTemp&" for Multi Currency Saving Account",True
		bMulti = True
	Else
		LogMessage "WARN","Verification","Failed, "&strTemp&" for Multi Currency Saving Account",False
		bMulti = False
	End If
	
	VerifyMultiCurrencyFieldsAndValue = bMulti
End Function

'[Verify Time Deposit Account Type Fields and Default Values]
Public Function VerifyTDFieldsAndValue(strSchemeCode)
	bTD = False
	strAccountType = "Time Deposit"
	
	bTD = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleTimeDepositAccountType(),strAccountType,"Account Type")
	bTD = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleTDSchemeCode(),strAccountType,"Scheme Code")
	
	bTD = verifyInnerText(HK_OnBoardCustomer_Page.eleTimeDepositAccountType(),strAccountType,strAccountType&"-Account Type")
	bTD = verifyInnerText(HK_OnBoardCustomer_Page.eleTDSchemeCode(),strSchemeCode,strAccountType&"-Scheme Code")
	VerifyTDFieldsAndValue = bTD
End Function

'[Verify Current Account Type Fields and Default Values]
Public Function VerifyCAFieldsAndValue(strSchemeCode,arrCurrencies)
	bCA = False
	strAccountType = "Current Account"
	If Not HK_OnBoardCustomer_Page.eleCurrentAccountType().Exist(3) Then
		HK_OnBoardCustomer_Page.btnAddAccount().Click
		HK_OnBoardCustomer_Page.btnCurrentAccount().Click
		Wait(2)
	End If
	
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleCurrentAccountType(),strAccountType,"Account Type")
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtCurrentAccntEffectiveDate(),strAccountType,"Effective Date")
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleCASchemeCode(),strAccountType,"Scheme Code")
	
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkCAOD(),strAccountType,"Over Draft Protection")
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkCAChequeBook(),strAccountType,"Cheque Book")
	
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.lstCACurrencyCode(),strAccountType,"Currency Code")
	bCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtCAAccountNumber(),strAccountType,"Account Number")
	
	bCA = verifyEffectiveDefaultDate(HK_OnBoardCustomer_Page.txtCurrentAccntEffectiveDate(),strAccountType&"-Effective Date")
	bCA = verifyInnerText(HK_OnBoardCustomer_Page.eleCurrentAccountType(),strAccountType,"-Account Type")
	bCA = verifyInnerText(HK_OnBoardCustomer_Page.eleCASchemeCode(),strSchemeCode,strAccountType&"-Scheme Code")
	
	bCA = verifyDropdownListValues(HK_OnBoardCustomer_Page.lstCACurrencyCode(),arrCurrencies,strAccountType&"-Currency Code")
	
	VerifyCAFieldsAndValue = bCA
	
End Function

'[Verify Staff Account Type Fields and Default Values]
Public Function VerifySCAFieldsAndValue(strSchemeCode,strCurrency)
	bSCA = False
	strAccountType = "Staff Current Account"
	If Not HK_OnBoardCustomer_Page.eleStaffCurrentAccountType().Exist(3) Then
		HK_OnBoardCustomer_Page.btnAddAccount().Click
		HK_OnBoardCustomer_Page.btnStaffCurrentAccount().Click
		Wait(2)
	End If
	
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleStaffCurrentAccountType(),strAccountType,"Account Type")
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSCAEffectiveDate(),strAccountType,"Effective Date")
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSCASchemeCode(),strAccountType,"Scheme Code")
	
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSCAOD(),strAccountType,"Over Draft Protection")
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.chkSCAChequeBook(),strAccountType,"Cheque Book")
	
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSCACurrencyCode(),strAccountType,"Currency Code")
	bSCA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSCAAccountNumber(),strAccountType,"Account Number")
	
	bSCA = verifyEffectiveDefaultDate(HK_OnBoardCustomer_Page.txtSCAEffectiveDate(),strAccountType&"-Effective Date")
	
	bSCA = verifyInnerText(HK_OnBoardCustomer_Page.eleStaffCurrentAccountType(),strAccountType,"-Account Type")
	bSCA = verifyInnerText(HK_OnBoardCustomer_Page.eleSCASchemeCode(),strSchemeCode,strAccountType&"-Scheme Code")
	
	bSCA = verifyInnerText(HK_OnBoardCustomer_Page.eleSCACurrencyCode(),strCurrency,strAccountType&"-Currency Code")
	
	VerifySCAFieldsAndValue = bSCA
End Function

'[Verify Savings Account Type Fields and Default Values]
Public Function VerifySAFieldsAndValue(strSchemeCode,strCurrency)
	bSA = False
	strAccountType = "Savings Account"
	If Not HK_OnBoardCustomer_Page.eleSavingAccountType().Exist(3) Then
		HK_OnBoardCustomer_Page.btnAddAccount().Click
		HK_OnBoardCustomer_Page.btnSavingAccount().Click
		Wait(2)
	End If
	
	bSA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSavingAccountType(),strAccountType,"Account Type")
	bSA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.txtSAEffectiveDate(),strAccountType,"Effective Date")
	bSA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSASchemeCode(),strAccountType,"Scheme Code")
	bSA = VerifyFieldExistenceInPage(HK_OnBoardCustomer_Page.eleSACurrencyCode(),strAccountType,"Currency Code")
	
	bSA = verifyInnerText(HK_OnBoardCustomer_Page.eleSavingAccountType(),strAccountType,"-Account Type")
	bSA = verifyEffectiveDefaultDate(HK_OnBoardCustomer_Page.txtSAEffectiveDate(),strAccountType&"-Effective Date")
	bSA = verifyInnerText(HK_OnBoardCustomer_Page.eleSASchemeCode(),strSchemeCode,strAccountType&"-Scheme Code")
	bSA = verifyInnerText(HK_OnBoardCustomer_Page.eleSACurrencyCode(),strCurrency,strAccountType&"-Currency Code")
	
	VerifySAFieldsAndValue = bSA
End Function

Public Function verifyEffectiveDefaultDate(objDate,strFieldName)
	bEffectiveDate = False
	strEffVal = objDate.GetROProperty("value")
	
'	strEffVal = CDate(strEffVal)
'	strDay1 = Day(strEffVal)
'	If Len(strDay1) = 1 Then strDay1 = "0"&strDay1
'	strMonth1 = Month(strEffVal)
'	If Len(strMonth1) = 1 Then strMonth1 = "0"&strMonth1
'	strYear1 = Year(strEffVal)
'	DateActual = strDay1 & "/" & strMonth1  & "/" & strYear1
		
	DateRequested = ConvertTodaysDateFormat
	
	If strEffVal = DateRequested Then
		LogMessage "RSLT","Verification","As Expected, Effective Date for " &strFieldName&" is Current Date.",True
		bEffectiveDate = True
	Else
		LogMessage "WARN","Verification","Failed, Effective Date for " &strFieldName&" is not Current Date, Actual: "&strEffVal ,False
	End If
	verifyEffectiveDefaultDate = bEffectiveDate
End Function

Public Function ConvertTodaysDateFormat()
	
	strDay = Day(Date)
	If Len(strDay) = 1 Then strDay = "0"&strDay
	strMonth = Month(Date)
	If Len(strMonth) = 1 Then strMonth = "0"&strMonth
	strYear = Year(Date)
	
	ConvertTodaysDateFormat = strDay & "/" & strMonth  & "/" & strYear
End Function
