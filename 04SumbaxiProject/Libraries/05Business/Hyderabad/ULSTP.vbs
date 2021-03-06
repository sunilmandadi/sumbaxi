Dim strError_OnSubmission
strError_OnSubmission = ""
'[Click on Continue button of Login Page]
Public Function clickContinue_ULSTP()
	bclickContinue_ULSTP = true
	bcLoginICALLScreen.btnContinue().Click
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the button Continue in the Home page.",false
		bclickContinue_ULSTP = false
	else
		LogMessage "RSLT","Verification","Button Continue clicked successfully. ",True
	End If
	clickContinue_ULSTP = bclickContinue_ULSTP
End Function

'[Click on ULSTP Money Bag]
Public Function clickULSTP_MoneyBag()
	clickULSTP_MoneyBag = true
	ULSTP.lblULSTPMoneyBag().Click
	WaitForIcallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the ULSTP Money Bag.",false
		clickULSTP_MoneyBag = false
	else
		LogMessage "RSLT","Verification","ULSTP Money Bag clicked successfully. ",True
	End If
End Function

'[Verify the Customer eligibility Dialog Box exists]
Public Function dlgCustomerEligibility_Exists()
	dlgCustomerEligibility_Exists = true
	If ULSTP.dlgCustEligibility().Exist Then
		WaitForIcallLoading
		LogMessage "RSLT","Verification","The dialog box Customer Eligibility exists. ",True
	else
		LogMessage "WARN","Verification","The dialog box Customer Eligibility does not exist.",False
		dlgCustomerEligibility_Exists = false
	End If
End Function

'[Click on Ok button of the Customer Eligibility Dialog Box]
Public Function clickOkBtn_ULSTPEligibility()
	clickOkBtn_ULSTPEligibility = true
	Wait(10)
	For iCount = 1 To 180 Step 1
		If Not ULSTP.btnOkCustEligibility.Exist(0.5) Then
			Wait(0.5)
		else
			ULSTP.btnOkCustEligibility().Click
			Exit for
		End if
	Next	
	WaitForIcallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the Ok Button of the Customer Eligibility Popup.",false
		clickOkBtn_ULSTPEligibility = false
	else
		LogMessage "RSLT","Verification"," Ok Button of the Customer Eligibility Popup clicked successfully. ",True
	End If
End Function

'[Click on DBS Cashline hyperlink of ULSTP Page]
Public Function clickDBSCashline_ULSTP()
	clickDBSCashline_ULSTP = true
	ULSTP.lblDBSCashline().Click
	WaitForIcallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the DBS Cashline.",false
		clickDBSCashline_ULSTP = false
	else
		LogMessage "RSLT","Verification","DBS Cashline clicked successfully. ",True
	End If
End Function

'[Click on POSB Loan Assist hyperlink of ULSTP Page]
Public Function clickPOSBLoanAssist_ULSTP()
	clickPOSBLoanAssist_ULSTP = true
	ULSTP.lblPOSBLoanAssist().Click
	WaitForIcallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the POSB Loan Assist.",false
		clickPOSBLoanAssist_ULSTP = false
	else
		LogMessage "RSLT","Verification","POSB Loan Assist clicked successfully. ",True
	End If
End Function

'[Set the Marital Status value in ULSTP Page as]
Public Function setMaritalStatus_ULSTP(strMaritalStatus)
	setMaritalStatus_ULSTP = true
	ULSTP.txtMaritalStatus().Set strMaritalStatus
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value of Marital Status.",false
		setMaritalStatus_ULSTP = false
	else
		LogMessage "RSLT","Verification","Value set successfully for Marital Status.",True
	End If
End Function

'[Set the Mobile Number Country Code in ULSTP Page as]
Public Function setMobileNo_CountryCode(strMobileNoCountryCode)
	setMobileNo_CountryCode = true
	ULSTP.txtMobileNoCountryCode().Set strMobileNoCountryCode
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value of Mobile Number Country Code.",false
		setMobileNo_CountryCode = false
	else
		LogMessage "RSLT","Verification","Value set successfully for Mobile Number Country Code.",True
	End If
End Function

'[Set the Mobile Number in ULSTP Page as]
Public Function setMobileNo_ULSTP(strMobileNo)
	setMobileNo_ULSTP = true
	ULSTP.txtMobileNumber().Set strMobileNo
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value of Mobile Number.",false
		setMobileNo_ULSTP = false
	else
		LogMessage "RSLT","Verification","Value set successfully for Mobile Number.",True
	End If
End Function

'[Click on Go Button of Mobile Number in ULSTP Page]
Public Function clickGo_MobileNo()
	clickGo_MobileNo = true
	ULSTP.btnGoMobileNo().Click
	WaitForICallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the Go button of Mobile No.",false
		clickGo_MobileNo = false
	else
		LogMessage "RSLT","Verification","Go button of Mobile No clicked successfully.",True
	End If
End Function

'[Set the Income Doc in ULSTP Page as]
Public Function setIncomeDoc_ULSTP(strIncomeDoc)
	setIncomeDoc_ULSTP = true
	ULSTP.txtIncomeDoc().Set strIncomeDoc
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value of Income Doc.",false
		setIncomeDoc_ULSTP = false
	else
		LogMessage "RSLT","Verification","Value set successfully for Income Doc.",True
	End If
End Function

'[Click on Go Button of Income Doc in ULSTP Page]
Public Function clickBtnGo_IncomeDoc()
	clickBtnGo_IncomeDoc = true
	ULSTP.btnGoIncomeDoc().Click
	WaitForICallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the Go button of Income Doc.",false
		clickBtnGo_IncomeDoc = false
	else
		LogMessage "RSLT","Verification","Go button of Income Doc clicked successfully.",True
	End If
End Function

'[Set the others Income Doc as]
Public Function setOthersIncomeDoc(lstOthersIncomeDoc, strAcctNo)
	setOthersIncomeDoc = true
	'Create the child objects
	Set oDesc = Description.Create
	oDesc("xpath").value = "//*[contains(@id,'ulstp_new_eform_others')]"
	Set othersIncomeList = ULSTP.dlgOthersIncomeDoc().ChildObjects(oDesc)
	ctOthers = ubound(lstOthersIncomeDoc)
	For it = 0 To ctOthers Step 1
		'click on checkbox dynamically if the value exists
		If lstOthersIncomeDoc(it) <> "" Then
			'click the checkbox; add 3 on the it as salary crediting is 3rd index
			index = it + 3
			othersIncomeList(index).Click
			If err.number <> 0 Then
				'error occurred; not able to click
				LogMessage "WARN","Verification","Failed to click the check box " &lstOthersIncomeDoc(it),false
				setOthersIncomeDoc = false
			else
				LogMessage "RSLT","Verification","Check box clicked successfully "&lstOthersIncomeDoc(it),True
			End If
			'Fill the account no if the "Salary Crediting" is clicked
			If lstOthersIncomeDoc(it) = "Salary Crediting" Then
				ULSTP.txtAccountNo().Set strAcctNo
			End If
			
		End If		
	Next
End Function

'[Click on Ok Button of Others Income Doc]
Public Function clickOk_OthersIncomeDoc()
	clickOk_OthersIncomeDoc = true
	ULSTP.btnOkOthersIncomeDoc().Click
	WaitForICallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the Ok button of Others Income Popup.",false
		clickOk_OthersIncomeDoc = false
	else
		LogMessage "RSLT","Verification","Ok button of Others Income Popup clicked successfully.",True
	End If
End Function

'[Verify the Account Number in ULSTP Page for Others Income Doc is populated as]
Public Function verifyAcctNoOthersDoc_Populated(lstOthersIncomeDoc, strAcctNo)
	verifyAcctNoOthersDoc_Populated = true
	ctOthers = ubound(lstOthersIncomeDoc)
	For it = 0 To ctOthers Step 1
		'The account no should be auto populated if the "Salary Crediting" was clicked
		If lstOthersIncomeDoc(it) = "Salary Crediting" Then
			'verify the account No is auto populated
			strActualAcctNo = ULSTP.txtAcctNoULSTP().GetRoProperty("value")
			'split the account no
			strAcct = split(strAcctNo," - ")
			strExpectedAct = strAcct(1)
			If strActualAcctNo = strExpectedAct Then
				LogMessage "RSLT","Verification","Account No matched. Actual: " &strActualAcctNo& " Expected: " &strExpectedAct,True
			else
				LogMessage "WARN","Verification","Account No not matched. Actual: " &strActualAcctNo& " Expected: " &strExpectedAct,False
				verifyAcctNoOthersDoc_Populated = false
			End If
		End If		
	Next
End Function

'[Verify the Account Number in ULSTP Page is populated as]
Public Function verifyAcctNo_ULSTP(strAcctNo)
	verifyAcctNo_ULSTP = true
	strActualAcctNo = ULSTP.txtAcctNoULSTP().GetRoProperty("value")
	'split the account no
	strAcct = split(strAcctNo," - ")
	strExpectedAct = strAcct(1)
	If strActualAcctNo = strExpectedAct Then
		LogMessage "RSLT","Verification","Account No matched. Actual: " &strActualAcctNo& " Expected: " &strExpectedAct,True
	else
		LogMessage "WARN","Verification","Account No not matched. Actual: " &strActualAcctNo& " Expected: " &strExpectedAct,False
		verifyAcctNo_ULSTP = false
	End If	
End Function

'[Verify the Salary crediting Account Number in ULSTP Page is populated as]
Public Function verifySalaryCreditAcctNo_ULSTP(strAcctNo)
	verifySalaryCreditAcctNo_ULSTP = true
	strActualAcctNo = ULSTP.txtAcctNoULSTP().GetRoProperty("value")
	'split the account no
	strAcct = split(strAcctNo," - ")
	strExpectedAct = strAcct(0)
	If strActualAcctNo = strExpectedAct Then
		LogMessage "RSLT","Verification","Account No matched. Actual: " &strActualAcctNo& " Expected: " &strExpectedAct,True
	else
		LogMessage "WARN","Verification","Account No not matched. Actual: " &strActualAcctNo& " Expected: " &strExpectedAct,False
		verifySalaryCreditAcctNo_ULSTP = false
	End If
	
End Function

'[Verify the Salary Crediting Dialog Box exists]
Public Function verifySalaryCrediting_Exist()
	verifySalaryCrediting_Exist = true
	If ULSTP.dlgSalaryCreditingIncomeDoc().Exist Then
		LogMessage "RSLT","Verification","The dialog box Salary Crediting exists. ",True
	else
		LogMessage "WARN","Verification","The dialog box Salary Crediting does not exist. ",False
		verifySalaryCrediting_Exist = false
	End If
End Function

'[Select the Account Number for the Salary Crediting as]
Public Function selectAcctNo_SalaryCrediting(strAcctNo)
	selectAcctNo_SalaryCrediting = true
	ULSTP.txtAccountNoSalaryCrediting().Set strAcctNo
	If err.number <> 0 Then
		'error occurred
		selectAcctNo_SalaryCrediting = false
		LogMessage "WARN","Verification","Failed to Set the value : Account Number of Salary Crediting." ,false
	else
		LogMessage "RSLT","Verification","Value set successfully: Account Number of Salary Crediting." ,True
	End If
End Function

'[Click on Get Income button of Salary Crediting Dialog Box]
Public Function clickGetIncome_SalCrediting()
	clickGetIncome_SalCrediting = true
	ULSTP.btnGetIncomeSalaryCrediting().Click
	WaitForICallLoading
	If err.number <> 0 Then
		'error occurred
		clickGetIncome_SalCrediting = false
		LogMessage "WARN","Verification","Failed to click the button : Get Income." ,false
	else
		LogMessage "RSLT","Verification","Button clicked successfully: Get Income." ,True
	End If
End Function

'[Verify the Income is retrieved successfully]
Public Function verifyIncomeRetrival(strMsg)
	verifyIncomeRetrival = true
	strActualMsg = ULSTP.lblSalaryCreditingMsg().GetRoProperty("innertext")
	If strMsg = strActualMsg Then
		LogMessage "RSLT","Verification","Income retrieved successfully. Actual: " &strActualMsg& " Expected: "&strMsg,True
	else
		LogMessage "WARN","Verification","Income not retrieved successfully. Actual: " &strActualMsg& " Expected: "&strMsg,True
	End If
End Function

'[Click on Ok button of Salary Crediting]
Public Function clickOk_SalaryCrediting()
	clickOk_SalaryCrediting = true
	ULSTP.btnOkSalaryCrediting().Click
	WaitForICallLoading
	If Err.Number<>0 Then
       clickOk_SalaryCrediting = false
       LogMessage "WARN","Verification","Failed to Click Button : Ok of Salary Crediting." ,false
   else
   		LogMessage "RSLT","Verification","Button clicked successfully: Ok of Salary Crediting." ,True
   End If
End Function

'***************** Function for Payslip Income Doc ****************
'[Set the Month of Payslip for Payslip Income Doc as]
Public Function setMonthOfPayslip(strMonthPayslip)
	setMonthOfPayslip = true
	strMonthPayslip = checknull(strMonthPayslip)
	If ISNUll(strMonthPayslip)Then	
		strMonthPayslip = monthName(Month(Now),2)&"-"&Year(now)
	End If	
	ULSTP.txtMonthOfPayslip().Set strMonthPayslip
	If Err.Number<>0 Then
       setMonthOfPayslip = false
       LogMessage "WARN","Verification","Failed to set the value: Month of Payslip." ,false
   else
   		LogMessage "RSLT","Verification","Value set successfully: Month of Payslip." ,True
   End If
End Function

'[Set the Basic for Payslip Income Doc as]
Public Function setBasic_Payslip(strAmt)
	setBasic_Payslip = true
	ULSTP.txtBasicPayslip().Set strAmt
	If Err.Number<>0 Then
       setBasic_Payslip = false
       LogMessage "WARN","Verification","Failed to set the value: Basic." ,false
   else
   		LogMessage "RSLT","Verification","Value set successfully: Basic." ,True
   End If
End Function

'[Click on ok button of Payslip Income Doc]
Public Function clickOk_Payslip()
	clickOk_Payslip = true
	ULSTP.btnOkPayslipIncomeDoc().click
	If Err.Number<>0 Then
       clickOk_Payslip = false
       LogMessage "WARN","Verification","Failed to click the ok button of Payslip." ,false
   else
   		LogMessage "RSLT","Verification","ok button of Payslip clicked successfully." ,True
   End If
End Function
'***************** End of Payslip Income Doc ***********************

'********************* Function for CPF Contribution Income Doc *******
'[Upload the CPF file in the Income Doc]
Public Function uploadCPF_IncomeDoc(strFileName)
	uploadCPF_IncomeDoc = true
	'Get the folder path from the OBTAF_Config
	strFolderPath = gstrAttachmentsPath
	filePath = strFolderPath + "\" + strFileName
	ULSTP.wbCPFUploadFile().Set filePath
	WaitForICallLoading
	If Err.Number<>0 Then
       uploadCPF_IncomeDoc = false
       LogMessage "WARN","Verification","Failed to upload the CPF File." ,false
   else
   		LogMessage "RSLT","Verification","CPF file uploaded successfully." ,True
   End If
End Function 

'[Set the employer in the CPF statement as]
Public Function setEmployer_CPF(strEmployer)
	setEmployer_CPF = true
	ULSTP.txtEmployerCPF().Set strEmployer
	If Err.Number<>0 Then
       setEmployer_CPF = false
       LogMessage "WARN","Verification","Failed to set the Employer." ,false
   else
   		LogMessage "RSLT","Verification","The Employer set successfully." ,True
   End If
End Function

'[Select the Month in the CPF Statement as]
Public Function selectMonth_CPF()
	selectMonth_CPF = true
	'click the months displayed
	Set oDesc = Description.Create
	oDesc("xpath").value = "//*[contains(@class,'dt-checkbox')]"
	Set chkMonths = ULSTP.tblCPFPopupContent().childobjects(oDesc)
	ctChkMonths = ubound(chkMonths)
	For it = 0 To ctChkMonths Step 1
		ctChkMonths(it).click
	Next
	
	If Err.Number<>0 Then
       selectMonth_CPF = false
       LogMessage "WARN","Verification","Failed to set the Employer." ,false
   else
   		LogMessage "RSLT","Verification","The Employer set successfully." ,True
   End If
End Function

'[Select the Radio button of STP in CPF Statement as]
Public Function selectSTP_CPF(strSTP)
	selectSTP_CPF = SelectRadioButtonGrp(strSTP, ULSTP.rbgProceedWithSTP, Array("Yes","No"))
	WaitForICallLoading
	If Err.Number<>0 Then
       selectSTP_CPF = false
       LogMessage "WARN","Verification","Failed to Select the Radio button : Proceed with STP." ,false
   else
   		LogMessage "RSLT","Verification"," Radio button clicked successfully:Proceed with STP." ,True
   End If
End Function

'[Click on ok button of CPF Contribution]
Public Function clickOk_CPF()
	clickOk_CPF = true
	ULSTP.btnOk_CPF().Click
	WaitForICallLoading
	If Err.Number<>0 Then
       clickOk_CPF = false
       LogMessage "WARN","Verification","Failed to Select the Ok button : CPF Statement." ,false
   else
   		LogMessage "RSLT","Verification","Ok button clicked successfully:CPF Statement.." ,True
   End If
End Function
'*********************** End of CPF Contribution Income Doc ***********

'[Select the Radio button of Credit Limit Request as]
Public Function selectRB_CreditLimitRqst(strCrdtLmtRqst, strLimit)
   selectRB_CreditLimitRqst = SelectRadioButtonGrp(strCrdtLmtRqst, ULSTP.rbgCreditLimitRqst, Array("(a) Maximum Credit Limit (in SGD)","(b) Preferred Credit Limit (in SGD)"))
   WaitForICallLoading
   If Err.Number<>0 Then
       selectRB_CreditLimitRqst = false
       LogMessage "WARN","Verification","Failed to Click Button : Credit Limit Request" ,false
   else
   		LogMessage "RSLT","Verification","Button clicked successfully: Credit Limit Request" ,True
   End If
   
   'if the Credit Limit Request is "(b) Preferred Credit Limit (in SGD)", fill the limit
   If strCrdtLmtRqst = "(b) Preferred Credit Limit (in SGD)" Then
   		ULSTP.txtCreditLimitRequest().Set strLimit
   		If err.number <> 0 Then
   			selectRB_CreditLimitRqst = false
   			LogMessage "WARN","Verification","Failed to set: Credit Limit Request" ,false
   		else
   			LogMessage "RSLT","Verification","Set successfully: Credit Limit Request" ,true
   		End If
   End If
End Function

'[Select the Radio button of Loan Amount Requested as]
Public Function selectLA_CreditLimitRqst(strCrdtLmtRqst, strLimit)
   selectLA_CreditLimitRqst = True
   SelectRadioButtonGrp strCrdtLmtRqst, ULSTP.rbgCreditLimitRqstLA, Array("(a) Maximum Loan Amount (in SGD)","(b) Loan Amount Requested (in SGD)")
   WaitForICallLoading
   If Err.Number<>0 Then
       selectLA_CreditLimitRqst = false
       LogMessage "WARN","Verification","Failed to Click Button : Loan Amount Request" ,false
   else
   		LogMessage "RSLT","Verification","Button clicked successfully: Loan Amount Request" ,True
   End If
   
   'if the Credit Limit Request is "(b) Loan Amount Requested (in SGD)", fill the Amount
   If strCrdtLmtRqst = "(b) Loan Amount Requested (in SGD)" Then
   		ULSTP.txtCreditLimitRequest().Set strLimit
   		If err.number <> 0 Then
   			selectLA_CreditLimitRqst = false
   			LogMessage "WARN","Verification","Failed to set: Loan Amount Request" ,false
   		else
   			LogMessage "RSLT","Verification","Set successfully: Loan Amount Request" ,true
   		End If
   End If
End Function

'[Set the Loan Tenure Mths in ULSTP Page as]
Public Function setLoanTenureMths_ULSTP(strLoanTenureMths)
	setLoanTenureMths_ULSTP = False
	setLoanTenureMths_ULSTP = selectItem_Combobox (ULSTP.lblLoanTenureMths(), strLoanTenureMths)
	If err.number <> 0 Then
		setLoanTenureMths_ULSTP = false
		LogMessage "WARN","Verification","Failed to set: Loan tenure Mths" ,false
	else
		LogMessage "RSLT","Verification","Set successfully: Loan tenure Mths" ,true
	End If
End Function

'[Set the Loan Servicing Acct in ULSTP Page as]
Public Function setLoanServicingAcct_ULSTP(strLoanServicingAcct)
	setLoanServicingAcct_ULSTP = true
	setLoanServicingAcct_ULSTP = selectItem_Combobox (ULSTP.lblLoanServicingAcct(), strLoanServicingAcct)
	If err.number <> 0 Then
		setLoanServicingAcct_ULSTP = false
		LogMessage "WARN","Verification","Failed to set: Loan Servicing Acct" ,false
	else
		LogMessage "RSLT","Verification","Set successfully: Loan Servicing Acct" ,true
	End If
End Function

'[Set the Campaign Code in ULSTP Page as]
Public Function setCampaignCode_ULSTP(strCampaignCode)
	setCampaignCode_ULSTP = true
	setCampaignCode_ULSTP = selectItem_Combobox (ULSTP.txtCampaignCode(), strCampaignCode)
	If err.number <> 0 Then
		setCampaignCode_ULSTP = false
		LogMessage "WARN","Verification","Failed to set: Campaign Code" ,false
	else
		LogMessage "RSLT","Verification","Set successfully: Campaign Code" ,true
	End If
	wait(5)
End Function

'[set the BTIL Referral Code in ULSTP Page as]
Public Function setBtilReferralCodeTxt(strTxtBtilReferralCode)
	setBtilReferralCodeTxt = true
	ULSTP.txtBtilReferralCode().Set strTxtBtilReferralCode
	If err.number <> 0 Then
		'failed to set the value
		LogMessage "WARN","Verification","Failed to set the value of BTIL Referral Code.",false
		setBtilReferralCodeTxt = false
	else
		LogMessage "RSLT","Verification","Value set successfully for BTIL Referral Code.",True
	End If	
End Function

'[Click on Next Button of the ULSTP Page]
Public Function clickNextBtn()
	clickNextBtn = true
	WaitForICallLoading
	
	 For iCount = 1 To 180 Step 1
		If Not ULSTP.btnNext().Exist(0.5) Then
			Wait(0.5)
		else
			ULSTP.btnNext().Click
			Exit for
		End if
	Next
	wait(5)
	If err.number <> 0 Then
		clickNextBtn = false
		LogMessage "WARN","Verification","Failed to Click: Next Button" ,false
	else
		LogMessage "RSLT","Verification","Button clicked successfully: Next Button" ,true
	End If
End Function

'[Click on Declaration of Cashline]
Public Function clickDeclCL()
	clickDeclCL = true
	ULSTP.lblDecCashline().Click
	wait 3
	If err.number <> 0 Then
		clickDeclCL = false
		LogMessage "WARN","Verification","Failed to Click: Declaration Hyperlink" ,false
	else
		LogMessage "RSLT","Verification","Hyperlink clicked successfully: Declaration" ,true
	End If
End Function

'[Click on I Agree button of the Declaration]
Public Function iAgree_Declaration()
	iAgree_Declaration = true
	ULSTP.btnDeclarationIAgree().Click
	If err.number <> 0 Then
		iAgree_Declaration = false
		LogMessage "WARN","Verification","Failed to Click: Declaration Hyperlink" ,false
	else
		LogMessage "RSLT","Verification","Hyperlink clicked successfully: Declaration" ,true
	End If
End Function

'[Click on Signature Hyperlink of ULSTP Page]
Public Function clickSignature_ULSTP()
	clickSignature_ULSTP = true
	ULSTP.lblSignature().Click
	If err.number <> 0 Then
		clickSignature_ULSTP = false
		LogMessage "WARN","Verification","Failed to Click: Signature Hyperlink" ,false
	else
		LogMessage "RSLT","Verification","Hyperlink clicked successfully: Signature" ,true
	End If
End Function

'[Fill the signature in the pdf and save]
Public Function fillSignature_ULSTP()
	wait 5
	fillSignature_ULSTP = true
	'Click on Display Fields Dialog icon
	ULSTP.lblDisplayFieldDialogs().Click
	wait 2
	ULSTP.lblSigCaptureIcon().Click
	wait 2
	'Fill the signautre; recorded using Analog Recorder
	'Window("Google Chrome").RunAnalog "Track2"
	Window("Google Chrome").RunAnalog "Track3"
	wait 2
	'Save the Signature
	ULSTP.lblSignatureSave().Click
	wait 2
	'Click on Save to DB icon
	ULSTP.lblSavetoDB().Click
	wait 2
	'Confirm the save
	ULSTP.btnSaveSignYes().Click
	wait 4
	ULSTP.btnCloseSignaturePDF().Click
	WaitForICallLoading	
	If err.number <> 0 Then
		fillSignature_ULSTP = false
		LogMessage "WARN","Verification","Failed to fill and save the signature." ,false
	else
		LogMessage "RSLT","Verification","Signature filled and saved successfully." ,true
	End If
End Function

'[Perform the Checker Login]
Public Function performCheckerLogin(strCheckerID, strPassword)
	performCheckerLogin = true
	ULSTP.txtCheckerUserID().Set strCheckerID
	ULSTP.txtCheckerPassword().Set strPassword
	'Click on Login
	ULSTP.btnCheckerLogin().Click
	WaitForICallLoading
	If err.number <> 0 Then
		performCheckerLogin = false
		LogMessage "WARN","Verification","Failed to Login as the checker." ,false
	else
		LogMessage "RSLT","Verification","Successfully logged in as the checker." ,true
	End If
End Function

'[Check the ID Sighted box in verification checklist]
Public Function chkIDSighted_ChkList()
	chkIDSighted_ChkList = true
	ULSTP.chkIDSighted().Click
	If err.number <> 0 Then
		chkIDSighted_ChkList = false
		LogMessage "WARN","Verification","Failed to Check the box: ID Sighted." ,false
	else
		LogMessage "RSLT","Verification","Successfully Checked the box: ID Sighted." ,true
	End If
End Function

'[Check the Mobile Number Changed box in verification checklist]
Public Function chkMobileNoChanged_ChkList()
	chkMobileNoChanged_ChkList = true
	ULSTP.chkMobileNoChanged().Click
	If err.number <> 0 Then
		chkMobileNoChanged_ChkList = false
		LogMessage "WARN","Verification","Failed to Check the box: Mobile No changed." ,false
	else
		LogMessage "RSLT","Verification","Successfully Checked the box: Mobile No changed." ,true
	End If
End Function

'[Check the Salary Crediting in verification checklist]
Public Function chkSalaryCrediting_ChkList()
	chkSalaryCrediting_ChkList = true
	ULSTP.chkSalaryCrediting().click
	If err.number <> 0 Then
		chkSalaryCrediting_ChkList = false
		LogMessage "WARN","Verification","Failed to Check the box: Salary Crediting." ,false
	else
		LogMessage "RSLT","Verification","Successfully Checked the box: Salary Crediting." ,true
	End If
End Function

'[Check the Income Verified in verification checklist]
Public Function chkPayslip_ChkList()
	chkPayslip_ChkList = true
	ULSTP.chkIncomeVerified().click
	If err.number <> 0 Then
		chkPayslip_ChkList = false
		LogMessage "WARN","Verification","Failed to Check the box: Payslip." ,false
	else
		LogMessage "RSLT","Verification","Successfully Checked the box: Payslip." ,true
	End If
End Function

'[Check the Signature verified box in verification checklist]
Public Function chkSignVerified_ChkList()
	chkSignVerified_ChkList = true
	ULSTP.chkSignatureVerified().Click
	If err.number <> 0 Then
		chkSignVerified_ChkList = false
		LogMessage "WARN","Verification","Failed to Check the box: Signature verified." ,false
	else
		LogMessage "RSLT","Verification","Successfully Checked the box: Signature verified." ,true
	End If
End Function

'[Verify the Mobile No in verification checklist displayed as]
Public Function verifyMobileNo_ChkList(strCountryCode, strMobileNo)
	verifyMobileNo_ChkList = true
	strExpecteMobileNo = strCountryCode & " " & strMobileNo
	strActualMobileNo = ULSTP.lblMobileNoCheckList().GetRoProperty("innertext")
	If strExpecteMobileNo = strActualMobileNo Then
		LogMessage "RSLT","Verification","Mobile no in Verification Checklist matching.Actual: " &strActualMobileNo& "Expected: "&strExpecteMobileNo ,true
	else
		LogMessage "WARN","Verification","Mobile no in Verification Checklist not matching.Actual: " &strActualMobileNo& "Expected: "&strExpecteMobileNo ,False
		verifyMobileNo_ChkList = false
	End If  
End Function

'[Verify the Account Number in verification checklist displayed as]
Public Function verifyAcctNo_ChkList(strAcctNo)
	verifyAcctNo_ChkList = true
	strActualAcctNo = ULSTP.txtAcctNoChkList().GetRoProperty("value")
	If strAcctNo = strActualAcctNo Then
		LogMessage "RSLT","Verification","Acct No in Verification Checklist matching.Actual: " &strActualAcctNo& "Expected: "&strAcctNo ,true
	else
		LogMessage "WARN","Verification","Acct no in Verification Checklist matching.Actual: " &strActualAcctNo& "Expected: "&strAcctNo ,False
		verifyAcctNo_ChkList = false
	End If 
End Function

'[Click on submit button of ULSTP Page]
Public Function clickSubmit_ULSTP()
	clickSubmit_ULSTP = true
	'Insert the TimeStamp in the datastore
	strTimeStamp = convertDateTime_WithoutSec(Now)
	gstrRuntimeCommentStep="Click on submit button of ULSTP Page"
	gstrParameterNameStep = "TimeStamp"&replace((replace((replace(now,"/","-"))," ","-")),":","-")
	insertDataStore gstrParameterNameStep, strTimeStamp
	ULSTP.btnSubmitULSTP().Click
	WaitForICallLoading
	
	If err.number <> 0 Then
		clickSubmit_ULSTP = false
		LogMessage "WARN","Verification","Failed to click the button: Submit ULSTP." ,false
	else
		LogMessage "RSLT","Verification","Successfully clicked the button: Submit ULSTP." ,true
	End If	
End Function

'[Verify the fields of the submission popup in ULSTP Page]
Public Function verifySubmissionPopup_ULSTP(strScanningRqd)
	verifySubmissionPopup_ULSTP = true
	'Verify the title
	strStatus = ULSTP.lblSubmissionTitle().GetRoProperty("innertext")
	LogMessage "RSLT","Verification","Title of the Submission Popup is:" &strStatus ,true

	'Depending upon the status verify the fields
	Select Case strStatus
		Case "Error Occurred"
			'Update the global parameter as Error; since no records are displayed in DSA Home page
			strError_OnSubmission = "Error"
			'Then the error occurred; verify the following
			' Message getting displayed
			strActErrMsg = ULSTP.lblSubmissionErrMsg().GetRoProperty("innertext")
			strExpectedMsg = "Sorry, we are facing some issues in submitting the application form for processing. Please print a copy of the application form and scan to process the application."
			If strActErrMsg <> strExpectedMsg Then
				'Mismatch of the error message
				LogMessage "WARN","Verification","Error Message not matching.Actual: " &strActErrMsg& " Expected: "&strExpectedMsg ,false
				verifySubmissionPopup_ULSTP = false
			else
				LogMessage "RSLT","Verification","Error Message matching.Actual: " &strActErrMsg& " Expected: "&strExpectedMsg ,true
			End If
			'Print button exists
			 verifySubmissionPopup_ULSTP = checkPrintBtn_Exists()
			'Application ID,Application Status and CDM App no
			strFaultString = ULSTP.lblFaultStringErrorMsg().GetRoProperty("innertext")
			strSplitVals = split(strFaultString,",")
			'Check if the App ID is not blank
			strAppIDLabel = split(strSplitVals(0),":")
			strAppID = strAppIDLabel(1)
			If strAppID <> "" Then
				LogMessage "RSLT","Verification","Application ID exists. Value is: " &strAppID,true
			else
				LogMessage "WARN","Verification","Application ID is blank " ,false
				verifySubmissionPopup_ULSTP = false
			End If
			'Check the Application status is error
			strAppStatusLabel = split(strSplitVals(1),":")
			strAppStatus = strAppStatusLabel(1)
			If strAppStatus <> "ERROR" Then
				LogMessage "RSLT","Verification","Application Status is matching as expectd.Expected Value is: ERROR",true
			else
				LogMessage "WARN","Verification","Application Status is not matching as expectd.Expected Value is: ERROR. Actual: "&strAppStatus,false
				verifySubmissionPopup_ULSTP = false
			End If
			'Check the CDM Application No
			strCDMLabel = split(strSplitVals(2),":")
			strCDMNo = strCDMLabel(1)
			If strCDMNo <> "" Then
				LogMessage "RSLT","Verification","CDM Application ID exists. Value is: " &strCDMNo,true
			else
				LogMessage "WARN","Verification","Application ID is blank " ,false
				verifySubmissionPopup_ULSTP = false
			End If
			
		Case "Pending"
			'verify the following fields
			'Scanning required
			strActScanningRqd = ULSTP.txtScanningRqd_PendingStatus().GetRoProperty("value")
'			If strScanningRqd =  strActScanningRqd Then
'				LogMessage "RSLT","Verification","Scanning Required matching.Actual: " &strActScanningRqd& " Expected: "&strScanningRqd ,true
'			else
'				LogMessage "WARN","Verification","Scanning Required not matching.Actual: " &strActScanningRqd& " Expected: "&strScanningRqd ,false
'				verifySubmissionPopup_ULSTP = false
'			End If
			'Verify EVO Reference No is populated
			strActEVORef = ULSTP.txtEVORef_PendingStatus().GetRoProperty("value")
			'Store the value in the global variable
			Environment.Value("strStoredEVoRef") = strActEVORef
			
'			If strActEVORef <> "" Then
'				LogMessage "RSLT","Verification","EVO Reference exists. EVO Reference No: " &strActEVORef,true
'			else
'				LogMessage "RSLT","Verification","EVO Reference does not exist.",true
'				verifySubmissionPopup_ULSTP = false
'			End If
			
'			'EVO Reference Number should exist if scanning required is yes
'			If strActEVORef = "" Then
'				'Scanning required should be No
'				If strActScanningRqd = "No" Then
'					LogMessage "RSLT","Verification","EVO Reference is blank and the scanning required is No." ,true
'				else
'					LogMessage "WARN","Verification","EVO Reference is blank and the scanning required is Yes." ,false
'				End If
'				'check the print button exists
'				'verifySubmissionPopup_ULSTP = checkPrintBtn_Exists() [LISA Env this is not required due to its not done yet If has to be uncommented]
'				'select the checkbox also
				ULSTP.chkPrintApp().Click
'				wait 2
'			else
				'EVO ref has value; hence scanning required should be yes
				If strActScanningRqd = "Yes" Then
					LogMessage "RSLT","Verification","EVO Reference is blank and the scanning required is Yes." ,true
				else
					LogMessage "WARN","Verification","EVO Reference is blank and the scanning required is No." ,false
				End If
			'End If
	End Select
End Function

'[Verify the BTIL fields of the submission popup in ULSTP Page]
Public Function verifyBTILSubmissionPopup_ULSTP(strRequestedLoanAmountinSGD,strExpLoanTenure)
	verifyBTILSubmissionPopup_ULSTP = true
	
	'Verify the title
	strStatus = ULSTP.lblSubmissionTitle().GetRoProperty("innertext")
	LogMessage "RSLT","Verification","Title of the Submission Popup is:" &strStatus ,true

	'Depending upon the status verify the fields
	Select Case strStatus
		
		Case "Processed"		
'			'Verify Requested Loan Amount in SGD
'			strActualLoanAmut = ULSTP.lblRequestedLoanAmountinSGD().GetRoProperty("value")			
'			'split the account no
'			strActualLoanAmut = split(strActualLoanAmut,".")			
'			strExpectedAct = strActualLoanAmut(0)
'			If strExpectedAct = strRequestedLoanAmountinSGD Then
'				LogMessage "RSLT","Verification","Requested Loan Amount matched. Actual: " &strRequestedLoanAmountinSGD& " Expected: " &strExpectedAct,True
'			else
'				LogMessage "WARN","Verification","Requested Loan Amount not matched. Actual: " &strRequestedLoanAmountinSGD& " Expected: " &strExpectedAct,False
'			End If
'			
'			'Verify Applied Interest Rate (% p.a.)
'			strBTILCCInterestRate = fetchFromDataStore(gstrRuntimeInterestRateStep,"BLANK","InterestRateInBTPL")(0)
'			strActualBTILCCInterestRate = ULSTP.lblAppliedInterestRatePA().GetRoProperty("value")
'			If strActualBTILCCInterestRate = strBTILCCInterestRate Then
'				LogMessage "RSLT","Verification","Applied Intreset Rate matched. Actual: " &strBTILCCInterestRate& " Expected: " &strActualBTILCCInterestRate,True
'			else
'				LogMessage "WARN","Verification","Applied Intreset Rate not matched. Actual: " &strBTILCCInterestRate& " Expected: " &strActualBTILCCInterestRate,False
'			End If
'			
'			'Verify Loan Tenure (in months)
'			strActualBTILCCLoanTenureInmOnths = ULSTP.lblLoanTenureinmonths().GetRoProperty("value")	
'			If strActualBTILCCLoanTenureInmOnths = strExpLoanTenure Then
'				LogMessage "RSLT","Verification","Requested Loan tenure matched. Actual: " &strExpLoanTenure& " Expected: " &strActualBTILCCLoanTenureInmOnths,True
'			else
'				LogMessage "WARN","Verification","Requested Loan Tenure not matched. Actual: " &strExpLoanTenure& " Expected: " &strActualBTILCCLoanTenureInmOnths,False
'				verifyAcctNo_ULSTP = false
'			End If
			
			'Get the Administrative Fee
			strGetAdministrativeFee = ULSTP.lblAdministrativeFee().GetRoProperty("value")			
			'Insert the Administrative Fee in the datastore
			gstrRuntimeAdministrativeFeeStep="Verify the BTIL fields of the submission popup in ULSTP Page"
			insertDataStore "BTILAdministrativeFee", strGetAdministrativeFee
	
			
			'Get the Effective Interest Rate
			strGetEffectiveInterestRate = ULSTP.lblEffectiveInterestRatePA().GetRoProperty("value")
			'Insert the Effective Interest Rate in the datastore
			gstrRuntimeEffectiveInterestRateStep="Verify the BTIL fields of the submission popup in ULSTP Page"
			insertDataStore "BTILEffectiveInterestRate", strGetEffectiveInterestRate
			
			ULSTP.chkPrintApp().Click
			If Err.Number<>0 Then
				LogMessage "WARN","Verification","Failed to Click: Submitted ULSTP BTIL application Reference in DSA Home Page" ,false
				verifyBTILSubmissionPopup_ULSTP = false
			End If
			
		Case "Error Occurred"
			'Update the global parameter as Error; since no records are displayed in DSA Home page
			strError_OnSubmission = "Error"
			'Then the error occurred; verify the following
			' Message getting displayed
			strActErrMsg = ULSTP.lblSubmissionErrMsg().GetRoProperty("innertext")
			strExpectedMsg = "Sorry, we are facing some issues in submitting the application form for processing. Please print a copy of the application form and scan to process the application."
			If strActErrMsg <> strExpectedMsg Then
				'Mismatch of the error message
				LogMessage "WARN","Verification","Error Message not matching.Actual: " &strActErrMsg& " Expected: "&strExpectedMsg ,false
				verifySubmissionPopup_ULSTP = false
			else
				LogMessage "RSLT","Verification","Error Message matching.Actual: " &strActErrMsg& " Expected: "&strExpectedMsg ,true
			End If
			'Print button exists
			 verifySubmissionPopup_ULSTP = checkPrintBtn_Exists()
			'Application ID,Application Status and CDM App no
			strFaultString = ULSTP.lblFaultStringErrorMsg().GetRoProperty("innertext")
			strSplitVals = split(strFaultString,",")
			'Check if the App ID is not blank
			strAppIDLabel = split(strSplitVals(0),":")
			strAppID = strAppIDLabel(1)
			If strAppID <> "" Then
				LogMessage "RSLT","Verification","Application ID exists. Value is: " &strAppID,true
			else
				LogMessage "WARN","Verification","Application ID is blank " ,false
				verifySubmissionPopup_ULSTP = false
			End If
			'Check the Application status is error
			strAppStatusLabel = split(strSplitVals(1),":")
			strAppStatus = strAppStatusLabel(1)
			If strAppStatus <> "ERROR" Then
				LogMessage "RSLT","Verification","Application Status is matching as expectd.Expected Value is: ERROR",true
			else
				LogMessage "WARN","Verification","Application Status is not matching as expectd.Expected Value is: ERROR. Actual: "&strAppStatus,false
				verifySubmissionPopup_ULSTP = false
			End If
			'Check the CDM Application No
			strCDMLabel = split(strSplitVals(2),":")
			strCDMNo = strCDMLabel(1)
			If strCDMNo <> "" Then
				LogMessage "RSLT","Verification","CDM Application ID exists. Value is: " &strCDMNo,true
			else
				LogMessage "WARN","Verification","Application ID is blank " ,false
				verifySubmissionPopup_ULSTP = false
			End If
			
		Case "Pending"
			'verify the following fields
			'Scanning required
			strActScanningRqd = ULSTP.txtScanningRqd_PendingStatus().GetRoProperty("value")
			strActEVORef = ULSTP.txtEVORef_PendingStatus().GetRoProperty("value")
			'Store the value in the global variable
			Environment.Value("strStoredEVoRef") = strActEVORef
				ULSTP.chkPrintApp().Click
				If strActScanningRqd = "Yes" Then
					LogMessage "RSLT","Verification","EVO Reference is blank and the scanning required is Yes." ,true
				else
					LogMessage "WARN","Verification","EVO Reference is blank and the scanning required is No." ,false
				End If
			'End If
	End Select
End Function

'**** Function to check if the print button exists on submission popup when the error occurs
Public Function checkPrintBtn_Exists()
	checkPrintBtn_Exists = true
	If ULSTP.btnPrintOnError().Exist Then
				LogMessage "RSLT","Verification","Print button in the error submission popup exists as expected.",true
				ULSTP.btnPrintOnError().Click
				WaitForICallLoading
				'Check if the pdf appears on clicking Print button
				If ULSTP.dlgPrintErrorSubmission().Exist Then
					LogMessage "RSLT","Verification","PDF exists on clicking print button.",true
					'Write the function to read the contents of the PDF
					
					'Click on Ok button
					ULSTP.btnOkPrintPDF().Click
					WaitForICallLoading
				else
					LogMessage "WARN","Verification","PDF does not exist on clicking print button.",false
					checkPrintBtn_Exists = false
				End If
			else
				LogMessage "WARN","Verification","Print button in the error submission popup does not exist as expected.",false
				checkPrintBtn_Exists = false
			End If
End Function

'[Click on Close button of the Submission Popup]
Public Function close_SubmissionPopup()
	close_SubmissionPopup = true
	ULSTP.btnCloseSubmissionPopup().Click
	WaitForICallLoading
	If err.number <> 0 Then
		close_SubmissionPopup = false
		LogMessage "WARN","Verification","Failed to click the button of Submission Popup: Close." ,false
	else
		LogMessage "RSLT","Verification","Successfully clicked the button of Submission Popup: Close." ,true
	End If
	wait(5)
End Function

'[Click on the submitted ULSTP application in DSA Home Page]
Public Function clickSubmittedULSTP_HomePage(strCINSuffix)
	clickSubmittedULSTP_HomePage = True
	'fetch the created on date and time from the datastore
	'fetch from dataStore
	strCreatedOn = fetchFromDataStore(gstrRuntimeCommentStep,"BLANK",gstrParameterNameStep)(0)
	If strError_OnSubmission = "Error" Then
		'then exit the function; no need to check as the record is not created
		LogMessage "RSLT","Verification","Error had occurred on submission. No records displayed in DSA Home Page." ,true
		Exit function
	End If

	'Fetch the EVO Ref from the environment value
	'strEVORef = Environment.Value("strStoredEVoRef")
	'create the list of list 
	Dim lstULApplData(1)
	strCreatedOn = replace(strCreatedOn,":","-")
	lstULApplData(0)="Created On:"&strCreatedOn
	lstULApplData(1)="CIN / CIN Suffix:"&strCINSuffix
	'lstULApplData(2)=" EVO Case No.:"&strEVORef
	
	With ULSTP
		clickSubmittedULSTP_HomePage = selectTableLink(.tblCSOHomePage_AppHeader,.tblCSOHomePage_AppContent,lstULApplData,"UL Applications" ,"Status",true,.ULApp_lnkNext ,.ULApp_lnkNext1 ,.ULApp_lnkPrev)
	End With
	
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Click: Submitted ULSTP application Reference in DSA Home Page" ,false
		clickSubmittedULSTP_HomePage = false
	End If	
	'bcverify_Logout.lnkLogout().Click
End Function

'---------------------
'[Click on logout button in ULSTP SGBRDSA1 user]
Public Function clickLogOutULSTPsgbrdsa1User()
	clickLogOutULSTPsgbrdsa1User = True
	bcverify_Logout.lnkLogout().Click
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Click: logout button in ULSTP SGBRDSA1 user" ,false
		clickLogOutULSTPsgbrdsa1User = false
	End If	
End Function

'******************************** For BTIL ****************
'[Click on DBS BTIL hyperlink of ULSTP Page]
Public Function clickDBSBTIL_ULSTP()
	clickDBSBTIL_ULSTP = true
	ULSTP.lblDBSBTIL().Click
	WaitForIcallLoading
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to click the DBS BTIL.",false
		clickDBSBTIL_ULSTP = false
	else
		LogMessage "RSLT","Verification","DBS BTIL clicked successfully. ",True
	End If
End Function

'[Select the Type of Product in BTIL Page as]
Public Function selectTypeOfProduct_BTIL(strProduct)
	selectTypeOfProduct_BTIL = true
	selectTypeOfProduct_BTIL = selectItem_Combobox (ULSTP.txtTypeOfProduct(), strProduct)
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Type of Product.",false
		selectTypeOfProduct_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Type of Product. ",True
	End If
End Function

'[Select the Cashline_CC in BTIL Page as]
Public Function selectCashlineCC_BTIL(strCLCC)
	selectCashlineCC_BTIL = true
	selectCashlineCC_BTIL = selectItem_Combobox (ULSTP.txtCashlineCCNo_BTIL(), strCLCC)
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Cashline/Credit Card.",false
		selectCashlineCC_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Cashline/Credit Card. ",True
	End If
End Function

'[Select the Type of Loan in BTIL Page as]
Public Function selectTypeOfLoan_BTIL(strTypeOfLoan)
	selectTypeOfLoan_BTIL = true
	ULSTP.txtTypeOfLoan_BTIL().Set strTypeOfLoan
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Type of Loan.",false
		selectTypeOfLoan_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Type of Loan. ",True
	End If
End Function

'[Select the Requested Loan Amount in BTIL Page as]
Public Function selectRqstLoanAmt_BTIL(strRqstLoanAmt)
	selectRqstLoanAmt_BTIL = true
	ULSTP.txtRequestedLoanAmt_BTIL().Set strRqstLoanAmt
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Requested Loan Amount.",false
		selectRqstLoanAmt_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Requested Loan Amount. ",True
	End If
End Function

'[Get the Interest Rate in BT PL Page as]
Public Function getInterestRateBTPL()
	getInterestRateBTPL = True
	strInterestRateBTPL = ULSTP.lblBTPLInterestRate.GetROProperty("value")
	txt = InStr(strInterestRateBTPL,"%")
	strInterestRateBTPL = Mid(strInterestRateBTPL,1,txt-1)
	'Insert the Interest Rate in the datastore
	gstrRuntimeInterestRateStep="Get the Interest Rate in BT PL Page as"
	insertDataStore "InterestRateInBTPL", strInterestRateBTPL
End Function


'[Select the Loan Tenure in BTIL Page as]
Public Function selectLoanTenure_BTIL(strLoanTenure)
	selectLoanTenure_BTIL = true
	ULSTP.txtLoanTenure_BTIL().Set strLoanTenure
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Loan Tenure.",false
		selectLoanTenure_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Loan Tenure. ",True
	End If
End Function

'[Select the Transfer to Account in BTIL Page as]
Public Function selectTransferAcct_BTIL(strAcctType)
	selectTransferAcct_BTIL = true
	ULSTP.txtTransferToAcct_BTIL().Set strAcctType
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Transfer to Account.",false
		selectTransferAcct_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Transfer to Account. ",True
	End If
End Function

'[Select the other bank name in BTIL Page as]
Public Function selectBTILOtherBankName(strBTILOtherBankName)
	selectBTILOtherBankName = True
	strBTILOtherBankName = checknull(strBTILOtherBankName)
	If Not IsNull(strBTILOtherBankName)  Then
		selectBTILOtherBankName = selectItem_Combobox (ULSTP.lblBTILCLOthrBank(), strBTILOtherBankName)
	End If
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Other Bank Name.",false
		setOtherBankAccnameBTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Other Bank Name. ",True
	End If	
End Function

'[set the Other Bank Account Name in BTIL Page as]
Public Function setOtherBankAccnameBTIL(strOBAnameBtil)
	setOtherBankAccnameBTIL = True
	strOBAnameBtil = checknull(strOBAnameBtil)
	If Not IsNull(strOBAnameBtil) Then
		ULSTP.lblBTILCLOtherBankAccountName().Set strOBAnameBtil
	End If	
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Other Bank Account Name.",false
		setOtherBankAccnameBTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Other Bank Account Name. ",True
	End If	
End Function

'[set other bank account number in BTIL as]
Public Function setOtherbankAccNumberBTIL(strOtherBanAcNumr)
	setOtherbankAccNumberBTIL = true
	strOtherBanAcNumr = checknull(strOtherBanAcNumr)
	If Not IsNull(strOtherBanAcNumr) Then
		ULSTP.lblBTILCLOtherBankAccountNo().Set strOtherBanAcNumr
	End If	
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Other Bank Account Number.",false
		setOtherbankAccNumberBTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Other Bank Account Number. ",True
	End If	
End Function

'[Select the Account No in BTIL Page as]
Public Function selectAcctNo_BTIL(strAcctNo)
	selectAcctNo_BTIL = true
	strAcctNo = checknull(strAcctNo)
	If Not IsNull(strAcctNo) Then
		ULSTP.txtAcctNo_BTIL().Set strAcctNo
	End If	
	If err.number <> 0 Then
		'failed to click the button
		LogMessage "WARN","Verification","Failed to set the value in Account No.",false
		selectAcctNo_BTIL = false
	else
		LogMessage "RSLT","Verification","Value set successfully in Account No. ",True
	End If
End Function

'[Click on Terms and Conditions of BTIL]
Public Function clickTermsNCond_BTIL()
	clickTermsNCond_BTIL = true
	ULSTP.lblTermsNCond_BTIL().Click
	wait 3
	If err.number <> 0 Then
		clickTermsNCond_BTIL = false
		LogMessage "WARN","Verification","Failed to Click: Terms and Conditions of BTIL" ,false
	else
		LogMessage "RSLT","Verification","Hyperlink clicked successfully: Terms and Conditions of BTIL." ,true
	End If
End Function

'[verify ULSTP Personal Details Section Default values]
Public Function verifyPersonalDetailsSecDefltVals(strSalutation,strCustomerName,strNRICPassportNo,strDateofBirth,strNationality,strPRStatus,strGender,strEducation,strNoofDependents,strEmailAddress)
	verifyPersonalDetailsSecDefltVals = False
	WaitForICallLoading
	If (verifyComboSelectItem (ULSTP.webListSalutation(), strSalutation, "Salutation") and verifyFieldValue (ULSTP.lblCustomerName(),strCustomerName, "Customer Name") and verifyFieldValue (ULSTP.lblNRICPassportNo(),strNRICPassportNo, "NRIC Passport No") and verifyFieldValue (ULSTP.lblDateOfBirth(),strDateofBirth, "Date of Birth") and verifyFieldValue (ULSTP.lblNationality(),strNationality, "Nationality") and verifyFieldValue (ULSTP.lblPRstatus(),strPRStatus, "PR Status") and verifyFieldValue (ULSTP.lblGender(),strGender, "Gender") and verifyComboSelectItem (ULSTP.lblEducation(),strEducation, "Education") and verifyFieldValue (ULSTP.lblNoOfDependents(),strNoofDependents, "No of Dependents") and verifyFieldValue (ULSTP.lblEmailAddress(),strEmailAddress, "Email Address")) Then
		verifyPersonalDetailsSecDefltVals = True
	End If
 	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Check: Personal Details Section Default values" ,false
       verifyPersonalDetailsSecDefltVals = false
    End If    
End Function

'[Verify ULSTP Mailing Address Details Default Values]
Public Function verifyMailingAddDetailsDefValues(strMailingAddressDetails,strPostalCode,strBlockNo,strBlockNo1,strBlockNo2,strStreetName1,strStreetName2,strResidentialStatus,strResidentialType)
verifyMailingAddDetailsDefValues = False
If verifyComboSelectItem (ULSTP.lblMailingAddressDetails(),strMailingAddressDetails, "Mailing Address Details") and verifyFieldValue (ULSTP.lblPostalCode(),strPostalCode, "Postal Code") and verifyFieldValue (ULSTP.lblBlockNo(),strBlockNo, "Block No") and verifyFieldValue (ULSTP.lblBlockNo1(),strBlockNo1, "Block No1") and verifyFieldValue (ULSTP.lblBlockNo2(),strBlockNo2, "Block No2") and verifyFieldValue (ULSTP.lblStreetName1(),strStreetName1, "Street Name1") and	verifyFieldValue (ULSTP.lblStreetName2(),strStreetName2, "Street Name2")  and verifyComboSelectItem (ULSTP.lblResidentialStatus(),strResidentialStatus, "Residential Status") and verifyComboSelectItem (ULSTP.lblResidentialType(),strResidentialType, "Residential Type")Then
	verifyMailingAddDetailsDefValues = True	
End If
If Err.Number<>0 Then
   LogMessage "WARN","Verification","Failed to Check: Mailing Address Details Default Values" ,false
   verifyMailingAddDetailsDefValues = false
End If
End Function

'[Verify ULSTP Current Employment Details Default Values]
Public Function verifyCurrentEmployDetailsDftVal(strCompanyName,strJobStatus,strJobTitle,strIndustryBusinessType,strLengthOfCurrentEmpYears,strLengthOfCurrentEmpMonths,strPreviousCompanyName,strLengthOfPreviousEmpYears,strLengthOfPreviousEmpMonths)
	verifyCurrentEmployDetailsDftVal = False
	If verifyFieldValue(ULSTP.lblCompanyName(),strCompanyName, "Company Name") and verifyComboSelectItem(ULSTP.lblJobStatus(),strJobStatus, "Job Status") and verifyComboSelectItem(ULSTP.lblJobTitle(),strJobTitle, "Job Title") and verifyComboSelectItem(ULSTP.lblIndustryBusinessType(),strIndustryBusinessType, "Industry Business Type") and verifyFieldValue(ULSTP.lblLengthOfCurrentEmpYears(),strLengthOfCurrentEmpYears, "Length Of Current Emp Years") and verifyFieldValue(ULSTP.lblLengthOfCurrentEmpMonths(),strLengthOfCurrentEmpMonths, "Length Of Current Emp Months") and verifyFieldValue(ULSTP.lblPreviousCompanyName(),strPreviousCompanyName, "Previous Company Name") and verifyFieldValue(ULSTP.lblLengthOfPreviousEmpYears(),strLengthOfPreviousEmpYears, "Length Of Previous Emp Years") and verifyFieldValue(ULSTP.lblLengthOfPreviousEmpMonths(),strLengthOfPreviousEmpMonths, "Length Of PreviousEmp Months") Then
		verifyCurrentEmployDetailsDftVal = True
	End If
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Current Employment Details Default Values" ,false
		verifyCurrentEmployDetailsDftVal = false
	End If
End Function

'[Verify ULSTP Loan Assist Post Submit Loan Request Details Section validation]
Public Function verifyULSTPPostSubmitLoanRequestDetailsValidation(strReqLoanAmtCrditLimit,strLoanTenureInMths,strElibleLoaAmountSGD,strLoanSerAccouNoAcNo)
	verifyULSTPPostSubmitLoanRequestDetailsValidation = False
	WaitForICallLoading
	If verifyInnerText(ULSTP.lblRequestedLoanAmountCreditLimit(),strReqLoanAmtCrditLimit, "Requested Loan Amount Credit Limit") and verifyInnerText(ULSTP.lblLoanTenureInMths(),strLoanTenureInMths, "Loan Tenure in mths") and verifyInnerText(ULSTP.lblEligibleLoanAmount(),strElibleLoaAmountSGD, "Eligible Loan Amount SGD") and verifyInnerText(ULSTP.lblLoanServicingAccountNoAccountNo(),strLoanSerAccouNoAcNo, "Loan Servicing Account No Account No") Then
		verifyULSTPPostSubmitLoanRequestDetailsValidation = true
	End If
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Loan Request Details Updated Values" ,false
		verifyULSTPPostSubmitLoanRequestDetailsValidation = false
	End If
End Function

'[Verify ULSTP Loan Assist Post Submit Personal Details validation]
Public Function verifyULSTPPostSumPerDetailVal(strPDSalutation,strPDCustomerName,strPDPersonalDetailsNRICPassportNo,strPDDateOfBirth,strPDNationality,strPDPRStatus,strPDGender,strPDMaritalStatus,strPDEducation,strPDNoOfDependents,strPDMobileNo,strPDHomeNo,strPDOfficeNo,strPDEmailAddress)
	verifyULSTPPostSumPerDetailVal = False
	If verifyInnerText(ULSTP.lblPDSalutation(),strPDSalutation, "Salutation") and verifyInnerText(ULSTP.lblPDCustomerName(),strPDCustomerName, "Customer Name") and verifyInnerText(ULSTP.lblPDPersonalDetailsNRICPassportNo(),strPDPersonalDetailsNRICPassportNo, "NRIC Passport No") and verifyInnerText(ULSTP.lblPDDateOfBirth(),strPDDateOfBirth, "Date Of Birth") and verifyInnerText(ULSTP.lblPDNationality(),strPDNationality,"Nationality") and verifyInnerText(ULSTP.lblPDPRStatus(),strPDPRStatus, "Status") and verifyInnerText(ULSTP.lblPDGender(),strPDGender, "Gender") and verifyInnerText(ULSTP.lblPDMaritalStatus(),strPDMaritalStatus, "Marital Status") and verifyInnerText(ULSTP.lblPDEducation(),strPDEducation, "Education") and verifyInnerText(ULSTP.lblPDNoOfDependents(),strPDNoOfDependents, "No Of Dependents") and verifyInnerText(ULSTP.lblPDMobileNo(),strPDMobileNo, "Mobile No") and verifyInnerText(ULSTP.lblPDHomeNo(),strPDHomeNo, "Home No") and verifyInnerText(ULSTP.lblPDOfficeNo(),strPDOfficeNo, "Office No") and verifyInnerText(ULSTP.lblPDEmailAddress(),strPDEmailAddress, "Email Address") Then
		verifyULSTPPostSumPerDetailVal = True		
	End If
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Post Submit Loan Personal Details Section Updated Values" ,false
		verifyULSTPPostSumPerDetailVal = false
	End If
End Function

'[Verify ULSTP Loan Assist Post Submit Mailing Address Details validation]
Public Function verifyULSTPMailAddDetValid(strMADPostalCode,strMADBlocNoLevelUnitNo,strMADStreetName1,strMADStreetName2,strMADResidentialStatus,strMADResidentialType)
	verifyULSTPMailAddDetValid = False
	If verifyInnerText(ULSTP.lblMADPostalCode(),strMADPostalCode, "Postal Code") and verifyInnerText(ULSTP.lblMADBlocNoLevelUnitNo(),strMADBlocNoLevelUnitNo, "Bloc No Level Unit No") and verifyInnerText(ULSTP.lblMADStreetName1(),strMADStreetName1, "Street Name1") and verifyInnerText(ULSTP.lblMADStreetName2(),strMADStreetName2, "Street Name2") and verifyInnerText(ULSTP.lblMADResidentialStatus(),strMADResidentialStatus, "Residential Status") and verifyInnerText(ULSTP.lblMADResidentialType(),strMADResidentialType, "Residential Type") Then
		verifyULSTPMailAddDetValid = True		
	End If	
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Post Submit Loan Mailing Address Details Section Updated Values" ,false
		verifyULSTPMailAddDetValid = false
	End If
End Function

'[Verify ULSTP Loan Assist Post Submit Current Employment Details validation]
Public Function verifyULSTPCurrntEmpDetVals(strCEDCompanyName,strCEDJobStatus,strCEDJobTitle,strCEDIndustryBusinessType,strCEDLengthOfCurrentEmp,strCEDPreviousCompanyName,strCEDLengthOfPreviousEmp)
	verifyULSTPCurrntEmpDetVals = False
	If verifyInnerText(ULSTP.lblCEDCompanyName(),strCEDCompanyName, "Company Name") and verifyInnerText(ULSTP.lblCEDJobStatus(),strCEDJobStatus, "CED Job Status") and verifyInnerText(ULSTP.lblCEDJobTitle(),strCEDJobTitle, "CED Job Title") and verifyInnerText(ULSTP.lblCEDIndustryBusinessType(),strCEDIndustryBusinessType, "CED Industry Business Type") and verifyInnerText(ULSTP.lblCEDLengthOfCurrentEmp(),strCEDLengthOfCurrentEmp, "CED Length Of Current Emp") and verifyInnerText(ULSTP.lblCEDPreviousCompanyName(),strCEDPreviousCompanyName, "CED Previous Company Name") and verifyInnerText(ULSTP.lblCEDLengthOfPreviousEmp(),strCEDLengthOfPreviousEmp, "CED Length Of Previous Emp") Then
		verifyULSTPCurrntEmpDetVals = True		
	End If
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Post Submit Loan Current Employment Details Section Updated Values" ,false
		verifyULSTPCurrntEmpDetVals = false
	End If
End Function

'[Verify ULSTP Loan Assist Post Submit For Bank Use Only validation]
Public Function verifyULSTPBankuseOnlyVals(strICSIncomeDoc,strFBUOCampaignCode,strBUOBBranchCode,strBUOBAgentCode)
	verifyULSTPBankuseOnlyVals = False
	If verifyInnerText(ULSTP.lblICSIncomeDoc(),strICSIncomeDoc, "Income Doc") and verifyInnerText(ULSTP.lblFBUOCampaignCode(),strFBUOCampaignCode, "FBUO Campaign Code") and verifyInnerText(ULSTP.lblBUOBBranchCode(),strBUOBBranchCode, "BUOB Branch Code") and verifyInnerText(ULSTP.lblBUOBAgentCode(),strBUOBAgentCode, "Company Name") Then
		verifyULSTPBankuseOnlyVals = True
	End If
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Post Submit Loan For Bank Use Only Details Section Updated Values" ,false
		verifyULSTPBankuseOnlyVals = false
	End If
End Function

'[verify ULSTP BT and PL Personal Details Section Default values]
Public Function verifyULSTPBTPLPrlDtlsSecDefltVals(strBTILSalutation,strBTILCustomerName,strBTILNRICPassportNo,strBTILDateOfBirth,strBTILNationality,strBTILGender)
	verifyULSTPBTPLPrlDtlsSecDefltVals = False
	WaitForICallLoading
	If verifyFieldValue (ULSTP.lblBTILSalutation(),strBTILSalutation, "BT IL Salutation") and verifyFieldValue (ULSTP.lblBTILCustomerName(),strBTILCustomerName, "B TIL Customer Name") and verifyFieldValue (ULSTP.lblBTILNRICPassportNo(),strBTILNRICPassportNo, "BTIL NRIC Passport No") and verifyFieldValue (ULSTP.lblBTILDateOfBirth(),strBTILDateOfBirth, "BTIL Date Of Birth") and verifyFieldValue (ULSTP.lblBTILNationality(),strBTILNationality, "BTIL Nationality") and verifyFieldValue (ULSTP.lblBTILGender(),strBTILGender, "BTIL Gender") Then
		 verifyULSTPBTPLPrlDtlsSecDefltVals = True
	End If
 	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Check: ULSTP BT PL Personal Details Section Default values" ,false
       verifyULSTPBTPLPrlDtlsSecDefltVals = false
    End If    
End Function

'[Verify ULSTP Credit Card and BT Post Submit Loan Request Details Section validation]
Public Function verifyULSTPCCBTPstSmtLnReqDetsValidation(strReqLoanAmtCrditLimit,strLoanTenureInMths,strBTILCCTypeOfProduct,strBTILCCCashlineCreditCardNo,strElibleLoaAmountSGD,strBTILCCTypeOfLoan,strBTILCCInterestRate,strBTILCCTransferToAccount,strLoanSerAccouNoAcNo)
	verifyULSTPCCBTPstSmtLnReqDetsValidation = False
		If IsNull(strBTILCCInterestRate) Then
		'fetch the Interest Rate from the datastore
		strBTILCCInterestRate = fetchFromDataStore(gstrRuntimeInterestRateStep,"BLANK","InterestRateInBTPL")(0)
	End If
	
	WaitForICallLoading
	If verifyInnerText(ULSTP.lblRequestedLoanAmountCreditLimit(),strReqLoanAmtCrditLimit, "Requested Loan Amount Credit Limit") and verifyInnerText(ULSTP.lblLoanTenureInMths(),strLoanTenureInMths, "Loan Tenure in mths") and verifyInnerText(ULSTP.lblBTILCCTypeOfProduct(),strBTILCCTypeOfProduct, "BT IL CC Type Of Product") and verifyInnerText(ULSTP.lblBTILCCCashlineCreditCardNo(),strBTILCCCashlineCreditCardNo, "BTIL CC Cash lineCredit Card No") and verifyInnerText(ULSTP.lblEligibleLoanAmount(),strElibleLoaAmountSGD, "Eligible Loan Amount SGD") and verifyInnerText(ULSTP.lblBTILCCTypeOfLoan(),strBTILCCTypeOfLoan, "BT IL CC Type Of Loan") and verifyInnerText(ULSTP.lblBTILCCInterestRate(),strBTILCCInterestRate, "BTIL CC Interest Rate") and verifyInnerText(ULSTP.lblBTILCCTransferToAccount(),strBTILCCTransferToAccount, "BTIL CC Transfer To Account") and verifyInnerText(ULSTP.lblLoanServicingAccountNoAccountNo(),strLoanSerAccouNoAcNo, "Loan Servicing Account No Account No") Then
		verifyULSTPCCBTPstSmtLnReqDetsValidation = true
	End If
	If Err.Number<>0 Then
		LogMessage "WARN","Verification","Failed to Check: Loan Request Details Updated Values" ,false
		verifyULSTPCCBTPstSmtLnReqDetsValidation = false
	End If
End Function
