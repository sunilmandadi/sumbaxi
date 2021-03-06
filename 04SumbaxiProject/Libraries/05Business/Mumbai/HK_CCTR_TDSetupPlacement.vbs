'[Verify Customer by answering Identification and Authentication questions]
Public Function VerifyCustomerQuestionAnswer(intIdentQstns,intAuthQstns,blnSubmitFlag)
	bVerifyCustomerQuestionAnswer = False
	ClickOnObject HK_CCTR_CustomerVerification_Page.btnNotVerified(),"Not Verified Button in Customer Overview Page"
	WaitForICallLoading
	
	SelectManualAuthenticationQuestions intIdentQstns
	
	SelectCustomerPortfolioQuestions intAuthQstns
	Wait(3)
	
	ClickOnObject HK_CCTR_CustomerVerification_Page.btnVerifyCustomer(),"Verify Customer Button in Verify Customer Page"
	WaitForICallLoading
	
	If HK_CCTR_CustomerVerification_Page.eleWarningDialog().Exist(5) Then
		If UCase(blnSubmitFlag) = "YES" Then
			ClickOnObject HK_CCTR_CustomerVerification_Page.btnWarningDialogYes(),"Yes Button - On Verify Customer Warning Dialog"
		Else
			ClickOnObject HK_CCTR_CustomerVerification_Page.btnWarningDialogNo(),"No Button - On Verify Customer Warning Dialog"
		End If
	End If
	
	If HK_CCTR_AmendMaturity_Page.btnVerificationOk().Exist(5) Then
		clickBtnVerificationOk
	End If
	
	VerifyCustomerQuestionAnswer = bVerifyCustomerQuestionAnswer
End Function

'[Perform Customer Verification]
Public Function CustomerVerification(intIdentQstns,intAuthQstns,strButtonText)
	bIncomp = False
	ClickOnObject HK_CCTR_CustomerVerification_Page.btnNotVerified(),"Not Verified Button in Customer Overview Page"
	WaitForICallLoading
	
	bIncomp = SelectManualAuthenticationQuestions(intIdentQstns)
	
	bIncomp = SelectCustomerPortfolioQuestions(intAuthQstns)
	
	ClickOnObject HK_CCTR_CustomerVerification_Page.btnVerifyCustomer(),"Verify Customer Button in Verify Customer Page"
	
	If HK_CCTR_CustomerVerification_Page.eleWarningDialog().Exist(5) Then
		ClickOnObject HK_CCTR_CustomerVerification_Page.btnWarningDialogYes(),"Yes Button - On Verify Customer Warning Dialog"
		Wait(3)
	End If
	
	If HK_CCTR_AmendMaturity_Page.btnVerificationOk().Exist(5) Then
		clickBtnVerificationOk
	End If
	
	verifyInnerText HK_CCTR_CustomerVerification_Page.btnNotVerified(),strButtonText,"Verification Button Text"
	CustomerVerification = bIncomp
End Function

'[Validate Verification Required message details]
Public Function ValidateVerifictionMessageDetails(strExpMsg,strVerType)
	bVerifMsg = False

	If strVerType <> "High Risk Verified" Then
		verifyVerificationDialogBox
		bVerifMsg = verifyInnerText(HK_CCTR_AmendMaturity_Page.eleDlgVerificationMsg(),strExpMsg,"Customer Verification Message")
		clickBtnVerificationOk
	Else
		If Not(HK_CCTR_AmendMaturity_Page.dlgVerification().Exist(10)) Then
			bVerifMsg = True
			LogMessage "RSLT","Verification","As expected,Verification Dialogue Box is not displayed for High Risk Verified Customer" ,True
		Else
			LogMessage "WARN","Verification","Failed, Verification Dialogue Box is displayed for High Risk Customer which is not expected",False
			clickBtnVerificationOk
			bVerifMsg = False
		End If
	End If
	
	ValidateVerifictionMessageDetails = bVerifMsg
End Function


'[Click on Setup Placement Button]
Public Function ClickOnSetupPlacementButton()
	bClick = False
	bClick = ClickOnObject(HK_CCTR_TDSetupPlacement_Page.btnSetupPlacement(),"Setup Placement Button in Placement Info Page")
	WaitForICallLoading
	ClickOnSetupPlacementButton = bClick
End Function

'[Verify Tab Name and default opened page]
Public Function VerifyTDTabNameAndPage(strProduct,strAccntNo,strTabName)
	bVerifyTDTabName = False
	bVerifyTDTabName = verifyInnerText(HK_CCTR_BalanceAndLimits_Page.eleAccountTab(),strProduct&"-"&strAccntNo," TD Tab name")
	
	bDefautlPage = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnSetupPlacement(),"Placement Info Page"," - Setup Placement Button")
	If bDefautlPage Then
		LogMessage "RSLT","Verification","As expected, On Clicking TD Account in Overview Page, Placement Info page opened by default.",True
		bVerifyTDTabName = True
	Else
		LogMessage "WARN","Verification","Failed to display Placement Info page opened by default, On Clicking TD Account in Overview Page.",False
		bVerifyTDTabName = False
	End If
	
	ClickOnSetupPlacementButton
	
	bVerifySetupTabName = verifyInnerText(HK_CCTR_TDSetupPlacement_Page.eleSetupPlacementTab(),strTabName," Setup Placement Tab name")
	If bVerifySetupTabName Then
		LogMessage "RSLT","Verification","As expected, On Clicking Setup Placement Button in Placement Info Page, Setup Placement page opened.",True
		bVerifyTDTabName = True
	Else
		LogMessage "WARN","Verification","Failed to display Setup Placement page, Setup Placement Button in Placement Info Page.",False
		bVerifyTDTabName = False
	End If
	
	VerifyTDTabNameAndPage = bVerifyTDTabName
End Function

'[Verify Left Panel fields in Setup Placement Page]
Public Function VerifyLeftPanelInSetupPlacementPage(sSelAccDetls)
	bVerifyLSec = False
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPLeftPanel(),"Setup Placement Page"," - Left Panel")
	
	bVerifyLSec = verifyTableContentList(HK_CCTR_TDSetupPlacement_Page.tblSelectedAccountHeader(),HK_CCTR_FundFXTransfer_Page.tblSelectedAccountContent(),sSelAccDetls,"Setup Placement - Selected TD Account",False,Null,Null,Null)
	
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountLbl(),"Setup Placement Page"," - Left Panel - Debit Account No Label")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountCurrLbl(),"Setup Placement Page"," - Left Panel - Debit Account Currency Label")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountAvlBalLbl(),"Setup Placement Page"," - Left Panel - Available Balance Label")
	
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountDropDownVal(),"Setup Placement Page"," - Left Panel - Debit Account No Drop Down")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountCurrDropDownVal(),"Setup Placement Page"," - Left Panel - Debit Account Currency Drop Down")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountAvlBalVal(),"Setup Placement Page"," - Left Panel - Available Balance Value")
	
	Environment.Value("DebitAccntCurrentBalance") = Trim(Replace(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountAvlBalVal().GetROProperty("innertext"),",",""))
	
	VerifyLeftPanelInSetupPlacementPage = bVerifyLSec
End Function

'[Verify Left Panel Debit Account and currency Dropdown values in Setup Placement Page]
Public Function VerifyDebitAccountDropdownValues(sDebitAccnts,sDebitCurrency)
	bDebit = False
	bDebit = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountDropDownVal(),sDebitAccnts,"Debit Account No")
	bDebit = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountCurrDropDownVal(),sDebitCurrency,"Debit Account Currency")
	VerifyDebitAccountDropdownValues = bDebit
End Function

'[Select Debit Account Details for Setup Placement]
Public Function SelectDebitAccountDetails(strDebitAcc,strDebitCurrency)
	bDebitAcc = False
	bDebitAcc = SetValue(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountDropDownVal(),strDebitAcc,"Debit Account No Dropdown")
	bDebitAcc = SetValue(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountCurrDropDownVal(),strDebitCurrency,"Debit Account Currency Dropdown")
	SelectDebitAccountDetails = bDebitAcc
End Function

'[Verify Right Panel fields in Setup Placement Page]
Public Function VerifyRightPanelInSetupPlacementPage()
	bVerifyRSec = False
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPRightPanel(),"Setup Placement Page"," - Right Panel")
	
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSchemeCodeLbl(),"Setup Placement Page"," - Right Panel - Scheme Code Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleCurrencyLbl(),"Setup Placement Page"," - Right Panel - Currency Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.elePrincipalAmountLbl(),"Setup Placement Page"," - Right Panel - Principal Amount Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleTenorLbl(),"Setup Placement Page"," - Right Panel - Tenor Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleInterestRateLbl(),"Setup Placement Page"," - Right Panel - Interest Rate (%) Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleMPCodeLbl(),"Setup Placement Page"," - Right Panel - MP Code Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleMaturityInstructionLbl(),"Setup Placement Page"," - Right Panel - Maturity Instruction Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleDisposalAccountNoLbl(),"Setup Placement Pag"," - Right Panel - Disposal Account No. Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleNextTenorLbl(),"Setup Placement Page"," - Right Panel - Next Tenor Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleDescriptionLbl(),"Setup Placement Page"," - Right Panel - Description Label")
	bVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleCommentsLbl(),"Setup Placement Page"," - Right Panel - Comments Label")
	VerifyRightPanelInSetupPlacementPage = bVerifyRSec
End Function

'[Verify Right Panel Scheme Code Dropdown values in Setup Placement Page]
Public Function VerifySchemeCodeDropDown(lstSchemeCode)
	bSchemeCode = False
	bSchemeCode = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtSchemeCodeVal(),lstSchemeCode,"Scheme Code")
	VerifySchemeCodeDropDown = bSchemeCode
End Function

'[Verify Right Panel Currency Dropdown values in Setup Placement Page]
Public Function VerifyCurrencyDropDown(lstCurrencyCodes)
	bCurrency = False
	bCurrency = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtCurrencyVal(),lstCurrencyCodes,"Currency")
	VerifyCurrencyDropDown = bCurrency
End Function

'[Verify Right Panel Tenor Dropdown values in Setup Placement Page]
Public Function VerifyTenorDropDowns(lstTenorTypes)
	bTenor = False
	For i = 0 To UBound(lstTenorTypes) Step 1
		bTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtTenorType(),Split(lstTenorTypes(i),":")(0),"Tenor Type-")
		lstTenorVals = Split(Split(lstTenorTypes(i),":")(1),"#")
		For j = 0 To UBound(lstTenorVals) Step 1
			bTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtTenorTypeVal(),lstTenorVals(j),"Tenor Value-")
		Next
	Next
	VerifyTenorDropDowns = bTenor
End Function

'[Verify Right Panel Maturity Instruction,Roll Over and Withdraw type Dropdown values in Setup Placement Page]
Public Function VerifyMaturityRolloverWithDrawDropDowns(lstMaturity)
	bRollOver = False
	For i = 0 To UBound(lstMaturity) Step 1
		strMaturityType = Split(lstMaturity(i),":")(0)
		
		'Check different values in Maturity Instruction Dropdown
		bRollOver = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtMaturityInstruction(),strMaturityType,"Maturity Instruction")
		
		lstTenorVals = Split(Split(lstMaturity(i),":")(1),"#")
		
		For j = 0 To UBound(lstTenorVals) Step 1
			
			arrtemp = Split(lstTenorVals(j),"*")
			
			For k = 0 To UBound(arrtemp) Step 1
			
				If UCase(arrtemp(k)) <> "NULL" Then
					bRollOver = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtRolloverType(),arrtemp(k),"Rollover Type")
				End If
				
				If UBound(arrtemp) > 0 Then
					If UCase(arrtemp(k+1)) <> "NULL" Then
						bRollOver = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtWithdrawType(),arrtemp(k+1),"Withdraw Type")
					End If
				End If
				
				Exit For
			Next
		Next	
	Next
	VerifyMaturityRolloverWithDrawDropDowns = bRollOver
End Function

'[Verify Right Panel Disposal Account No Dropdown values in Setup Placement Page]
Public Function VerifyDisposalAccountNoDropDown(lstDisposalAccntNo)
	bAccntDrop = False
	bAccntDrop = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtDisposalAccountNo(),lstDisposalAccntNo,"Disposal Account No")
	VerifyDisposalAccountNoDropDown = bAccntDrop
End Function

'[Verify Right Panel Next Tenor Dropdown values in Setup Placement Page]
Public Function VerifyNextTenorType(lstNextTenorTypes)
	bNextTenor = False
	
	For i = 0 To UBound(lstNextTenorTypes) Step 1
		bNextTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtNextTenorType(),Split(lstNextTenorTypes(i),":")(0),"Next Tenor Type-")
		lstNextTenorVals = Split(Split(lstNextTenorTypes(i),":")(1),"#")
		For j = 0 To UBound(lstNextTenorVals) Step 1
			bNextTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtNextTenorTypeVal(),lstNextTenorVals(j),"Next Tenor Value-")
		Next
	Next
	VerifyNextTenorType = bNextTenor
End Function

'[Verify Right Panel Default Description Text in Setup Placement Page]
Public Function VerifyDefaultDescriptionText(strDescText)
	bDesc = False
	bDesc = verifyInnerText(HK_CCTR_TDSetupPlacement_Page.eleDescriptionVal(),strDescText," Right Panel Description Text in Setup Placement Page")
	VerifyDefaultDescriptionText = bDesc
End Function

'[Verify Right Panel Knowledge Base Hyperlink in Setup Placement Page]
Public Function VerifyKnowledgeBaseLink()
	bKBase = False
	bKBase = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.lnkKnowledgeBase(),"Setup Placement Page","-Knowledge Base Hyperlink in Setup Placement Page")
	VerifyKnowledgeBaseLink = bKBase
End Function

'[Select Setup Placement Right Panel data in Setup Placement page]
Public Function SelectDataInSetupPlacementPage(sSchCod,sCurr,sPriAmnt,sTenor,sIntRate,sMatchInst,sWithDrawType,sDisposAccnt,sComment)
	bSelectData = False
	
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtSchemeCodeVal(),sSchCod,"Scheme Code-")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtCurrencyVal(),sCurr,"Currency-")
	
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtPrincipalAmountVal(),sPriAmnt,"Principal Amount-")
	
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtTenorType(),sTenor(0),"Tenor Type-")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtTenorTypeVal(),sTenor(1),"Tenor Type Value-")
	
	If HK_CCTR_TDSetupPlacement_Page.txtInterestRate().GetROProperty("readonly") = 1 Then
		LogMessage "RSLT","Verification","As expected, Interest Rate field is read only.",True
		bSelectData = True
	Else
		LogMessage "WARN","Verification","Failed, Interest Rate field is not read only.",False
	End If
	
	bSelectData = verifyFieldValue(HK_CCTR_TDSetupPlacement_Page.txtInterestRate(),sIntRate," Interest Rate")
	
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtMaturityInstruction(),sMatchInst,"Maturity Instruction")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtWithdrawType(),sWithDrawType,"Withdraw Type")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtDisposalAccountNo(),sDisposAccnt,"Disposal Account No")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtCommentsVal(),sComment,"Comments")
	SelectDataInSetupPlacementPage = bSelectData
End Function

'[Click On Next Button in Setup Placement page]
Public Function ClickOnNextButtonInSetupPlacement()
	bNext = False
	bNext = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnNextButtonSetup(),"Setup Placement Page","-Next Button")
	bNext = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnCancelButtonSetup(),"Setup Placement Page","-Cancel Button")
	bNext = ClickOnObject(HK_CCTR_TDSetupPlacement_Page.btnNextButtonSetup(),"Next Button in Setup Placement Page")
	WaitForICallLoading
	ClickOnNextButtonInSetupPlacement = bNext
End Function

'[Verify Placement Confirmation Window]
Public Function VerifyPlacementConfWindow()
	bConfirm = False
	bConfirm = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnCanelButtonSetup(),"Placement Confirmation Window","-Cancel Button")
	bConfirm = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnProceedButtonSetup(),"Placement Confirmation Window","-Proceed Button")
	If bConfirm Then
		ClickOnObject HK_CCTR_TDSetupPlacement_Page.btnProceedButtonSetup(),"Proceed Button in Placement Confirmation Window"
		WaitForICallLoading
	Else
		If HK_CCTR_TDSetupPlacement_Page.eleErrorOccuredMesg().Exist(2) Then
			strErrorText = Trim(HK_CCTR_TDSetupPlacement_Page.eleErrorOccuredMesg().GetROProperty("innertext"))
			LogMessage "WARN","Verification","Failed,Unexpected Error occured: Error Message: " &strErrorText,False
			HK_CCTR_TDSetupPlacement_Page.btnOKInErrorOccuredMesg().Click
			Wait(2)
		End If 
	End If
	VerifyPlacementConfWindow = bConfirm
End Function

'[Verify Confirmation Message Window]
Public Function VerifyPlacementConfMesg(strMsg)
	bConfirm = False
	bConfirm = verifyInnerText(HK_CCTR_TDSetupPlacement_Page.eleConfMesgText(),strMsg,"Confirmation Message Text")
	bConfirm = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnConfMesgSetupNo(),"Confirmation Message Window","-No Button")
	bConfirm = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnConfMesgSetupYes(),"Confirmation Message Window","-Yes Button")
	If bConfirm Then
		ClickOnObject HK_CCTR_TDSetupPlacement_Page.btnConfMesgSetupYes(),"Yes Button in Confirmation Message Window"
		WaitForICallLoading
	End If
	VerifyPlacementConfMesg = bConfirm
End Function

'[Verify Setup Placement Request Submission Window fields]
Public Function VerifySetupRequestSubmssionWindow()
	bRqstSub = False
	bRqstSub = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionDialog(),"Request Submission","-Popup Window")
	
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionSRNoLbl(),"SR Number" , "Request Submission-SR Number"
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionSRStatusLbl(),"SR Status" , "Request Submission-SR Status"
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionTDAccntNoLbl(),"TD Account Number" , "Request Submission-TD Account Number"
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionTDPlacementNumberLbl(),"Placement Number" , "Request Submission-Placement Number"
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionTDIntrRateLbl(),"Interest Rate" , "Request Submission-Interest Rate"
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionTDIntrAmntLbl(),"Interest Amount" , "Request Submission-Interest Amount"
	verifyInnerText_Pattern HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionTDMatrDetlsLbl(),"Maturity Details" , "Request Submission-Maturity Details"
	 
	If HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionSRNo().Exist(3) Then
		Environment.Value("SRNumber") = HK_CCTR_TDSetupPlacement_Page.eleRequestSubmissionSRNo().GetROProperty("innertext")
	Else
		Environment.Value("SRNumber") = ""
	End If

	ClickOnObject HK_CCTR_TDSetupPlacement_Page.btnOKRqstSubmission(),"Ok Button Request Submission Popup Window"
	WaitForICallLoading
	VerifySetupRequestSubmssionWindow = bRqstSub
End Function

'[Validate invalid inputs in Setup Placement Principal Amount field]
Public Function ValidatePrincipalAmountField(lstPrinAmntsNeg)
	bPrinAmnts = False
	For i = 0 To UBound(lstPrinAmntsNeg) Step 1
		strAmount = Split(lstPrinAmntsNeg(i),":")(0)
		strMessage = Split(lstPrinAmntsNeg(i),":")(1)
		bPrinAmnts = SetValue(HK_CCTR_TDSetupPlacement_Page.txtPrincipalAmountVal(),strAmount,"Principal Amount-")
		HK_CCTR_TDSetupPlacement_Page.txtCommentsVal().Click
		Wait(2)
		If HK_CCTR_TDSetupPlacement_Page.elePrincipalAmountValidationMsg().Exist(3) Then
			strActMsg = Trim(HK_CCTR_TDSetupPlacement_Page.elePrincipalAmountValidationMsg().GetROProperty("innertext"))
			If strActMsg = strMessage Then
				LogMessage "RSLT","Verification","As expected,Error message ("&strMessage& ") displayed for Invalid Principal amount:"&strAmount,True
				bPrinAmnts = True
			Else
				LogMessage "WARN","Verification","Failed, Unexpected Error Message displayed for Invalid Principal amount:"&strAmount & ", Expected : (" &strMessage& ") ,Actual: ("&strActMsg &")",False
			End If
		End If
	Next
	ValidatePrincipalAmountField = bPrinAmnts
End Function

'[Verify Insufficient balance message]
Public Function VerifyInsufficentFund(strMsg)
	bInsuffi = False
	If HK_CCTR_TDSetupPlacement_Page.btnCanelButtonSetup().Exist(3) Then
		HK_CCTR_TDSetupPlacement_Page.btnCanelButtonSetup().Click
		Wait(2)
	End If
	If HK_CCTR_TDSetupPlacement_Page.elePrincipalAmountValidationMsg().Exist(3) Then
			strActMsg = Trim(HK_CCTR_TDSetupPlacement_Page.elePrincipalAmountValidationMsg().GetROProperty("innertext"))
		strActMsg = Trim(Split(strActMsg,":")(0))
		If strActMsg = strMsg Then
			LogMessage "RSLT","Verification","As expected,Insufficient balance message ("&strMsg& ") displayed When Principle Amount greater then Available Balance",True
			bInsuffi = True
		Else
			LogMessage "WARN","Verification","Failed, Unexpected Error Message displayed When Principle Amount greater then Available Balance , Expected : (" &strMsg& ") ,Actual: ("&strActMsg &")",False
		End If
	End If
	VerifyInsufficentFund = bInsuffi
End Function

'[Verify Comma Functionality in Principal Amount field]
Public Function VeifyCommaFunctioninPA(lstCommaInputs)
	bComma = False
	For i = 0 To UBound(lstCommaInputs) Step 1
		strInput = Split(lstCommaInputs(i),":")(0)
		strExptd = Split(lstCommaInputs(i),":")(1)
		HK_CCTR_TDSetupPlacement_Page.txtPrincipalAmountVal().Set(strInput)
		HK_CCTR_TDSetupPlacement_Page.txtCommentsVal().Click
		Wait(2)
		strActVal = Trim(HK_CCTR_TDSetupPlacement_Page.txtPrincipalAmountVal().GetROProperty("value"))
		If strActVal = strExptd Then
			LogMessage "RSLT","Verification","As expected,Comma Validation displayed correctly in Principal Amount for input : "&strInput,True
			bComma = True
		Else
			LogMessage "WARN","Verification","Failed, Comma Validation incorrect fo input: "&strInput & ", Expected : (" &strExptd& ") ,Actual: ("&strActVal &")",False
		End If
		
	Next
	VeifyCommaFunctioninPA = bComma
End Function

'[Validate max allowable characters in Comments field]
Public Function ValidateMaxCharSetupComments(strLongComment,strCommentMaxLength)
	bValidateMaxCommentChars = False
	'SetValue HK_CCTR_TDSetupPlacement_Page.txtCommentsVal(),strLongComment,"Max Chars in Comments Field"
	
	HK_CCTR_TDSetupPlacement_Page.txtCommentsVal().Set strLongComment
	
	Wait(1)
	strActualStringLen = Len(HK_CCTR_TDSetupPlacement_Page.txtCommentsVal().GetROProperty("innertext"))
	If strActualStringLen = strCommentMaxLength Then
		LogMessage "RSLT","Verification","As expected Comments section not allowing more then "& strCommentMaxLength & " characters.",True
		bValidateMaxCommentChars = True
	ElseIf strActualStringLen > strCommentMaxLength Then
		LogMessage "WARN","Verification","Failed , Comments section is allowing more then "& strCommentMaxLength & "characters. Expected: " &strCommentMaxLength& ", Actual: "&strActualStringLen ,False
	ElseIf strActualStringLen < strCommentMaxLength Then
		LogMessage "WARN","Verification","Failed , Comments section is allowing less then "& strCommentMaxLength & " characters. Expected: " &strCommentMaxLength& ", Actual: "&strActualStringLen ,False
	End If
	
	ValidateMaxCharSetupComments = bValidateMaxCommentChars
End Function
