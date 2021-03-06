'[Click on Create TD Button in Overview page]
Public Function ClickOnCreateTDButton()
	bClickCreateTD = False
	bClickCreateTD = ClickOnObject(HK_CCTR_CreateTD_Page.btnCreateTD(),"Create TD Button in Customer Overview Page")
	WaitForICallLoading
	ClickOnCreateTDButton = bClickCreateTD
End Function

'[Verify Create TD Tab Name and default opened page]
Public Function VerifyCreateTDTabNameAndPage(strTabName)
	bVerifyTDTabName = False
	bVerifyTDTabName = verifyInnerText(HK_CCTR_CreateTD_Page.eleCreateTDTab(),strTabName,"Create TD Tab name")
	
	bDefautlPage = VerifyFieldExistenceInPage(HK_CCTR_CreateTD_Page.eleCreateTDTab(),"Create TD Page","Create TD Tab")
	
	If bDefautlPage Then
		LogMessage "RSLT","Verification","As expected, On Clicking Create TD Button in Overview Page, Create TD Page opened.",True
		bVerifyTDTabName = True
	Else
		LogMessage "WARN","Verification","Failed to display Create TD Page, On Clicking Create TD Button in Overview Page.",False
	End If

	VerifyCreateTDTabNameAndPage = bVerifyTDTabName
End Function

'[Verify Left Panel fields in Create TD Page]
Public Function VerifyLeftPanelInCreateTDPage()
	bVerifyLSec = False
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPLeftPanel(),"Create TD Page","Left Panel")
	
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_CreateTD_Page.elePrimaryOwnerLbl(),"Create TD Page"," - Left Panel - Primary Owner Label")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_CreateTD_Page.eleTDAccountSignTypeLbl(),"Create TD Page"," - Left Panel - TD Account Sign Type Label")
	
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountLbl(),"Create TD Page"," - Left Panel - Debit Account No Label")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountCurrLbl(),"Create TD Page"," - Left Panel - Debit Account Currency Label")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountAvlBalLbl(),"Create TD Page"," - Left Panel - Available Balance Label")
	
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_CreateTD_Page.txtTDAccountSignTypeVal(),"Create TD Page"," - Left Panel - TD Account Sign Type Drop Down")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_CreateTD_Page.elePrimaryOwnerVal(),"Create TD Page"," - Left Panel - Primary Owner Value")
	
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountDropDownVal(),"Create TD Page"," - Left Panel - Debit Account No Drop Down")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountCurrDropDownVal(),"Create TD Page"," - Left Panel - Debit Account Currency Drop Down")
	bVerifyLSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountAvlBalVal(),"Create TD Page"," - Left Panel - Available Balance Value")
	
	Environment.Value("DebitAccntCurrentBalance") = Trim(Replace(HK_CCTR_TDSetupPlacement_Page.eleSPDebitAccountAvlBalVal().GetROProperty("innertext"),",",""))
	
	VerifyLeftPanelInCreateTDPage = bVerifyLSec
End Function

'[Verify Left Panel Dropdown values in Create TD Page]
Public Function VerifyTDDebitAccountDropdownValues(sAccntSgnType,sDebitAccnts,sDebitCurrency)
	bTDDebit = False
	bTDDebit = verifyDropdownListValues(HK_CCTR_CreateTD_Page.txtTDAccountSignTypeVal(),sAccntSgnType,"TD Account Sign Type")
	bTDDebit = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountDropDownVal(),sDebitAccnts,"Debit Account No")
	bTDDebit = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountCurrDropDownVal(),sDebitCurrency,"Debit Account Currency")
	VerifyTDDebitAccountDropdownValues = bTDDebit
End Function

'[Select TD Sign type and Debit Account Details for Create TD]
Public Function SelectTDDebitAccountDetails(sAccntSgnType,strDebitAcc,strDebitCurrency)
	bTDDebitAcc = False
	bTDDebitAcc = SetValue(HK_CCTR_CreateTD_Page.txtTDAccountSignTypeVal(),sAccntSgnType,"TD Account Sign Type Dropdown")
	bTDDebitAcc = SetValue(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountDropDownVal(),strDebitAcc,"Debit Account No Dropdown")
	bTDDebitAcc = SetValue(HK_CCTR_TDSetupPlacement_Page.txtSPDebitAccountCurrDropDownVal(),strDebitCurrency,"Debit Account Currency Dropdown")
	SelectTDDebitAccountDetails = bTDDebitAcc
End Function

'[Verify Right Panel fields in Create TD Page]
Public Function VerifyTDRightPanelInSetupPlacementPage()
	bTDVerifyRSec = False
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSPRightPanel(),"Create TD Page"," - Right Panel")
	
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleSchemeCodeLbl(),"Create TD Page"," - Right Panel - Scheme Code Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleCurrencyLbl(),"Create TD Page"," - Right Panel - Currency Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.elePrincipalAmountLbl(),"Create TD Page"," - Right Panel - Principal Amount Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleTenorLbl(),"Create TD Page - Right Panel"," - Tenor Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleInterestRateLbl(),"Create TD Page"," - Right Panel - Interest Rate (%) Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleMPCodeLbl(),"Create TD Page"," - Right Panel - MP Code Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleMaturityInstructionLbl(),"Create TD Page"," - Right Panel - Maturity Instruction Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleDisposalAccountNoLbl(),"Create TD Page"," - Right Panel - Disposal Account No. Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleNextTenorLbl(),"Create TD Page"," - Right Panel - Next Tenor Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleDescriptionLbl(),"Create TD Page"," - Right Panel - Description Label")
	bTDVerifyRSec = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.eleCommentsLbl(),"Create TD Page"," - Right Panel - Comments Label")
	VerifyTDRightPanelInSetupPlacementPage = bTDVerifyRSec
End Function

'[Verify Right Panel Scheme Code Dropdown values in Create TD Page]
Public Function VerifyTDSchemeCodeDropDown(lstSchemeCode)
	bSchemeCode = False
	bSchemeCode = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtSchemeCodeVal(),lstSchemeCode,"Scheme Code")
	VerifyTDSchemeCodeDropDown = bSchemeCode
End Function

'[Verify Right Panel Currency Dropdown values in Create TD Page]
Public Function VerifyTDCurrencyDropDown(lstCurrencyCodes)
	bCurrency = False
	bCurrency = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtCurrencyVal(),lstCurrencyCodes,"Currency")
	VerifyTDCurrencyDropDown = bCurrency
End Function

'[Verify Right Panel Tenor Dropdown values in Create TD Page]
Public Function VerifyTDTenorDropDowns(lstTenorTypes)
	bTenor = False
	For i = 0 To UBound(lstTenorTypes) Step 1
		bTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtTenorType(),Split(lstTenorTypes(i),":")(0),"Tenor Type-")
		lstTenorVals = Split(Split(lstTenorTypes(i),":")(1),"#")
		For j = 0 To UBound(lstTenorVals) Step 1
			bTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtTenorTypeVal(),lstTenorVals(j),"Tenor Value-")
		Next
	Next
	VerifyTDTenorDropDowns = bTenor
End Function

'[Verify Right Panel Maturity Instruction,Roll Over and Withdraw type Dropdown values in Create TD Page]
Public Function VerifyTDMaturityRolloverWithDrawDropDowns(lstMaturity)
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
	VerifyTDMaturityRolloverWithDrawDropDowns = bRollOver
End Function

'[Verify Right Panel Disposal Account No Dropdown values in Create TD Page]
Public Function VerifyTDDisposalAccountNoDropDown(lstDisposalAccntNo)
	bAccntDrop = False
	bAccntDrop = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtDisposalAccountNo(),lstDisposalAccntNo,"Disposal Account No")
	VerifyTDDisposalAccountNoDropDown = bAccntDrop
End Function

'[Verify Right Panel Next Tenor Dropdown values in Create TD Page]
Public Function VerifyTDNextTenorType(lstNextTenorTypes)
	bNextTenor = False
	For i = 0 To UBound(lstNextTenorTypes) Step 1
		bNextTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtNextTenorType(),Split(lstNextTenorTypes(i),":")(0),"Next Tenor Type-")
		lstNextTenorVals = Split(Split(lstNextTenorTypes(i),":")(1),"#")
		For j = 0 To UBound(lstNextTenorVals) Step 1
			bNextTenor = verifyDropdownListValues(HK_CCTR_TDSetupPlacement_Page.txtNextTenorTypeVal(),lstNextTenorVals(j),"Next Tenor Value-")
		Next
	Next
	VerifyTDNextTenorType = bNextTenor
End Function

'[Verify Right Panel Default Description Text in Create TD Page]
Public Function VerifyTDDefaultDescriptionText(strDescText)
	bDesc = False
	bDesc = verifyInnerText(HK_CCTR_TDSetupPlacement_Page.eleDescriptionVal(),strDescText," Right Panel Description Text in Create TD Page")
	VerifyTDDefaultDescriptionText = bDesc
End Function

'[Verify Right Panel Knowledge Base Hyperlink in Create TD Page]
Public Function VerifyTDKnowledgeBaseLink()
	bKBase = False
	bKBase = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.lnkKnowledgeBase(),"Create TD Page"," Knowledge Base Hyperlink")
	VerifyTDKnowledgeBaseLink = bKBase
End Function

'[Select Create TD Right Panel data]
Public Function SelectTDDataRightPanel(sSchCod,sCurr,sPriAmnt,sTenor,sIntRate,sMatchInst,sWithDrawType,sDisposAccnt,sComment)
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
	
	bSelectData = verifyFieldValue(HK_CCTR_TDSetupPlacement_Page.txtInterestRate(),sIntRate,"-Interest Rate")
	
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtMaturityInstruction(),sMatchInst,"Maturity Instruction")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtWithdrawType(),sWithDrawType,"Withdraw Type")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtDisposalAccountNo(),sDisposAccnt,"Disposal Account No")
	bSelectData = SetValue(HK_CCTR_TDSetupPlacement_Page.txtCommentsVal(),sComment,"Comments")
	SelectTDDataRightPanel = bSelectData
End Function

'[Click On Next Button in Create TD page]
Public Function ClickTDOnNextButtonInSetupPlacement()
	bNext = False
	bNext = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnNextButtonSetup(),"Create TD Page","-Next Button")
	bNext = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnCancelButtonSetup(),"Create TD Page","-Cancel Button")
	bNext = ClickOnObject(HK_CCTR_TDSetupPlacement_Page.btnNextButtonSetup(),"Next Button in Create TD Page")
	WaitForICallLoading
	ClickTDOnNextButtonInSetupPlacement = bNext
End Function

'[Verify TD Confirmation Window]
Public Function VerifyTDPlacementConfWindow()
	bConfirm = False
	bConfirm = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnCanelButtonSetup(),"TD Confirmation Window","-Cancel Button")
	bConfirm = VerifyFieldExistenceInPage(HK_CCTR_TDSetupPlacement_Page.btnProceedButtonSetup(),"TD Confirmation Window","-Proceed Button")
	If bConfirm Then
		ClickOnObject HK_CCTR_TDSetupPlacement_Page.btnProceedButtonSetup(),"Proceed Button in TD Confirmation Window"
		WaitForICallLoading
	Else
		If HK_CCTR_TDSetupPlacement_Page.eleErrorOccuredMesg().Exist(2) Then
			strErrorText = Trim(HK_CCTR_TDSetupPlacement_Page.eleErrorOccuredMesg().GetROProperty("innertext"))
			LogMessage "WARN","Verification","Failed,Unexpected Error occured: Error Message: " &strErrorText,False
			HK_CCTR_TDSetupPlacement_Page.btnOKInErrorOccuredMesg().Click
			Wait(2)
		End If 
	End If
	VerifyTDPlacementConfWindow = bConfirm
End Function

'[Verify Create TD Request Submission Window fields]
Public Function VerifyTDRequestSubmssionWindow()
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
	VerifyTDRequestSubmssionWindow = bRqstSub
End Function
