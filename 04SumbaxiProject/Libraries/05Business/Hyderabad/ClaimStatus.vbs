'[Click on Claim Status link under Banking Facilities in Overview Page]
Public Function ClickClaimStatus()
    ClickClaimStatus=true
	bcCustomerOverview.lnkClaimStatus.Click
	If Err.Number<>0 Then
       ClickClaimStatus=false
       LogMessage "WARN","Verification","Failed to Click Button : Claim Status" ,false
       Exit Function
   	End If
	WaitForICallLoading
End Function

'[Verify the Claim Status Tab]
Public Function CheckClaimStatusTab(strCheckClaimStatusTab)
	CheckClaimStatusTab = true	
	If Not IsNull (strCheckClaimStatusTab) Then
		CheckClaimStatusTab = verifyInnerText(ClaimStatus.lblClaimStatusTab(),strCheckClaimStatusTab, "ClaimStatus")
	else
	  CheckClaimStatusTab = false
	  Exit function
	End If	
End Function

'[Verify Pink Panel displayed in Claim Status Page]
Public Function verifyClaimStatusPinkPanel(strName,strCIN,strSegment)
	verifyClaimStatusPinkPanel = false	
	If Not IsNull (strName) and Not IsNull (strCIN) and Not IsNull (strSegment) Then
	   If verifyInnerText(ClaimStatus.lblName(),strName, "Name") and verifyInnerText(ClaimStatus.lblCIN(),strCIN, "CIN") and verifyInnerText(ClaimStatus.lblSegment(),strSegment, "Segment") Then
		verifyClaimStatusPinkPanel = true
	   End If
	End If
End Function

'[Select Account No radio Button]
Public Function SelectAccountNoRadioButton()
    SelectAccountNoRadioButton=true
    ClaimStatus.rbtAccountNo.Click
	If Err.Number<>0 Then
       SelectAccountNoRadioButton=false
       LogMessage "WARN","Verification","Failed to Click Radio Button : Account No" ,false
       Exit Function
   	End If
	WaitForICallLoading	
End Function

'[Select the Account NO from the dropdown as]
Public Function SelectAccountNoDropdown(strAccNo)
   If Not IsNull(strAccNo) Then
	 SelectAccountNoDropdown = selectItem_Combobox(ClaimStatus.lstAccountNo(), strAccNo)
   else
      SelectAccountNoDropdown = false
      Exit function
   End If
   WaitForICallLoading
End Function

'[Click on Go Button]
Public Function ClickOnGoButton()    
	ClickOnGoButton=true
	ClaimStatus.btnGO.Click
	If Err.Number<>0 Then
       ClickOnGoButton=false
       LogMessage "WARN","Verification","Failed to Click Button : GO" ,false
       Exit Function
   	End If
	WaitForICallLoading
End Function

'[Verify the Case Information table details as]
Public Function VerifyCaseStatusInformationTable(lstlstCaseTransactionDetails)
	WaitForICallLoading
	WaitForICallLoading
    VerifyCaseStatusInformationTable=True
    If Not IsNull(lstlstCaseTransactionDetails)Then
       VerifyCaseStatusInformationTable = verifyTableContentList(ClaimStatus.tblCaseStatusInformationHeader,ClaimStatus.tblCaseStatusInformationContent,lstlstCaseTransactionDetails,"CaseTransactionDetails",false,null,null,null)
    else
       VerifyCaseStatusInformationTable = false
       Exit function
    End If
    WaitForICallLoading
 End Function

'[Click on Case No]
Public Function ClickOnCaseNoLink(strColumnName,lstlstCaseTransactionDetails)
  ClickOnCaseNoLink = true
  WaitForICallLoading
  If Not IsNull(lstlstCaseTransactionDetails) and Not IsNull(strColumnName) Then
      ClickOnCaseNoLink = selectTableLink(ClaimStatus.tblCaseStatusInformationHeader,ClaimStatus.tblCaseStatusInformationContent,lstlstCaseTransactionDetails,"CaseTransactionDetails",strColumnName,false,null,null,null)
   else
   ClickOnCaseNoLink = false
   Exit function
  End If
End Function

'[Verify the Case Status displayed in the case details ToolBar]
Public Function VerifyCaseStatusInToolBar(strCaseStatusInToolBar)
   WaitForICallLoading
   If Not IsNull (strCaseStatusInToolBar) Then	
       VerifyCaseStatusInToolBar = verifyInnerText(ClaimStatus.lblCaseDetailsTooBar(),strCaseStatusInToolBar, "CaseToolBarStatus")
   else
     VerifyCaseStatusInToolBar = false
     Exit function
   End If
End Function

'[Verify the Case Information Section details]
Public Function verifyCaseInformationSectionDetails(strDateAndTimeReceived,strCreatedBy, strProcessor, strAuthoriser, strMainCaseNumber)
	verifyCaseInformationSectionDetails = false
	WaitForICallLoading
	If Not IsNull(strDateAndTimeReceived) and Not IsNull(strCreatedBy) and Not IsNull(strProcessor) and Not IsNull(strAuthoriser) and Not IsNull(strMainCaseNumber) Then
	    If verifyInnerText (ClaimStatus.lblCIDateandTimeReceived(),strDateAndTimeReceived, "Date and Time Received") and verifyInnerText (ClaimStatus.lblCICreatedBy(),strCreatedBy, "Created By") and verifyInnerText (ClaimStatus.lblCIProcessor(),strProcessor, "Processor") and verifyInnerText (ClaimStatus.lblCIAuthoriser(),strAuthoriser, "Authoriser") and verifyInnerText (ClaimStatus.lblCIMainCaseNumber(),strMainCaseNumber, "MainCaseNumber") Then
		   verifyCaseInformationSectionDetails = true
	   End If
	End If
	
    ClaimStatus.lblCaseInformation.Click
    If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Case information Header" ,false
       verifyCaseInformationSectionDetails = false
    End If
    WaitForICallLoading    
End Function

'[Verify the Comments History Section]
Public Function verifyCommentsHistorySection(lstlstCommentsHistoryList)
  verifyCommentsHistorySection = true
  ClaimStatus.lblCommentsHistory.Click
  WaitForICallLoading
  If Err.Number<>0 Then
     LogMessage "WARN","Verification","Failed to Click Button : Comments History Section" ,false
     verifyCommentsHistorySection = false
     Exit Function
  End If
  
  If Not IsNull(lstlstCommentsHistoryList)Then
    verifyCommentsHistorySection = verifyTableContentList(ClaimStatus.lblCommentsHistoryHeader,ClaimStatus.lblCommentsHistoryContent,lstlstCommentsHistoryList,"CommentsHistorySection",false,null,null,null)
  else
   verifyCommentsHistorySection = false
  End If  
  
 ClaimStatus.lblCommentsHistory.Click
  If Err.Number<>0 Then
     LogMessage "WARN","Verification","Failed to Click Button : Comments History Section" ,false
     verifyCommentsHistorySection = false
     Exit Function
  End If
  WaitForICallLoading
End Function

'[Verify the Case History Section]
Public Function verifyCaseHistorySection(lstlstCaseHistoryList)
     verifyCaseHistorySection = true
     ClaimStatus.lblCaseHistory.Click
     WaitForICallLoading
     If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Case History Section" ,false
       verifyCaseHistorySection = false
       Exit Function
     End If
     
     If Not IsNull(lstlstCommentsHistoryList)Then
        verifyCaseHistorySection = verifyTableContentList(ClaimStatus.lblCaseHistoryHeader,ClaimStatus.lblCaseHistorycontent,lstlstCaseHistoryList,"CaseHistorySection",false,null,null,null)
     else
        verifyCaseHistorySection = false
     End If
     
     ClaimStatus.lblCaseHistory.Click
     If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Case History Section" ,false
       verifyCaseHistorySection = false
       Exit Function
     End If
     WaitForICallLoading
End Function

'[Verify the Customer Information]
Public Function verifyCustomerInformationSection(strCustomerCIN, strCallerCIN, strCustomerName, strCallerName, strCustomerStatus, strMobileNumber, strCommunicationMode)

    verifyCustomerInformationSection = false
	ClaimStatus.lblCustomerInformation.Click
	WaitForICallLoading
	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Customer Information Section" ,false
       Exit Function
     End If
     
     If Not IsNull(strCustomerCIN) and Not IsNull(strCallerCIN) and Not IsNull(strCustomerName) and Not IsNull(strCallerName) and Not IsNull(strMobileNumber) and Not IsNull(strCommunicationMode) and Not IsNull(strCustomerStatus) Then
       If verifyInnerText (ClaimStatus.lblCuICustomerCIN(),strCustomerCIN, "CustomerCIN") and verifyInnerText (ClaimStatus.lblCuICallerCIN(),strCallerCIN, "CallerCIN") and verifyInnerText (ClaimStatus.lblCuICustomerName(),strCustomerName, "CustomerName") and verifyInnerText (ClaimStatus.lblCuICallerName(),strCallerName, "CallerName") and verifyInnerText (ClaimStatus.lblCuIMobileNumber(),strMobileNumber, "MobileNumber") and verifyInnerText (ClaimStatus.lblCuICommunicationMode(),strCommunicationMode, "CommunicationMode") and verifyInnerText (ClaimStatus.lblCuICustomerStatus(),strCustomerStatus, "CustomerStatus") Then
     	  verifyCustomerInformationSection = true
       End If
     End If
	ClaimStatus.lblCustomerInformation.Click	
	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Customer Information Section" ,false
       verifyCustomerInformationSection = false
       Exit Function
     End If
     WaitForICallLoading
End Function

'[Verify the Claims Information]
Public Function verifyClaimsInformationSection(strRequestType, strCardNumber, strRequestedAmount, strReceivedAmount, strDepositedAmount, strClaimAmount, strMachineIDLocation, strTornNoteSN, strCustomerRemarks, strAccountNumber, strCICSFindings, strTransactionExistAndNoReversalInCPS)
	
	verifyClaimsInformationSection = false
	ClaimStatus.lblClaimsInformation.Click
	WaitForICallLoading
	If Err.Number<>0 Then	
       LogMessage "WARN","Verification","Failed to Click Button : Claims Information Section" ,false
       Exit Function
    End If
    
    If Not IsNull(strRequestType) and Not IsNull(strCardNumber) and Not IsNull(strRequestedAmount) and Not IsNull(strReceivedAmount) and Not IsNull(strDepositedAmount) and Not IsNull(strClaimAmount) and Not IsNull(strMachineIDLocation) and Not IsNull(strTornNoteSN) and Not IsNull(strCustomerRemarks) and Not IsNull(strAccountNumber) and Not IsNull(strCICSFindings) and Not IsNull(strTransactionExistAndNoReversalInCPS) Then
        If verifyInnerText (ClaimStatus.lblClaInRequestType(),strRequestType, "RequestType") and verifyInnerText (ClaimStatus.lblClaInCardNumber(),strCardNumber, "CardNumber") and verifyInnerText (ClaimStatus.lblClaInRequestedAmount(),strRequestedAmount, "RequestedAmount") and verifyInnerText (ClaimStatus.lblClaInDepositedAmount(),strDepositedAmount, "DepositedAmount") and verifyInnerText (ClaimStatus.lblClaInReceivedAmount(),strReceivedAmount, "ReceivedAmount") and verifyInnerText (ClaimStatus.lblClaInClaimAmount(),strClaimAmount, "ClaimAmount") and verifyInnerText (ClaimStatus.lblClaInMachineIDLocation(),strMachineIDLocation, "MachineIDLocation") and verifyInnerText (ClaimStatus.lblClaInTornNoteSN(),strTornNoteSN, "TornNoteSN") and verifyInnerText (ClaimStatus.lblClaInCustomerRemarks(),strCustomerRemarks, "CustomerRemarks") and verifyInnerText (ClaimStatus.lblClaInAccountNumber(),strAccountNumber, "AccountNumber") and verifyInnerText (ClaimStatus.lblClaInCICSFindings(),strCICSFindings, "CICSFindings") and verifyInnerText (ClaimStatus.lblClaInTransactionexistandNoReversalinCPS(),strTransactionExistAndNoReversalInCPS, "TransactionExistAndNoReversalInCPS") Then
        	verifyClaimsInformationSection = true
        End If       
    End If
    
    ClaimStatus.lblClaimsInformation.Click
	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Claims Information Section" ,false
       verifyClaimsInformationSection = false
       Exit Function
    End If
    WaitForICallLoading
End Function

'[Verify the Processor Information]
Public Function verifyProcessorInformationSection(strClaimCategory, strCategoryType, strUserBranchCode, strUserProfitCode, strTransactionDateTime, strValueDate, strTransfereeCIN, strTransfereeName, strTransfereeAccountNumber, strSWCHorCCorGWRef, strAcquirerName, strAcquirerBIN, strApprovedAmount, strTransactionComments)
   verifyProcessorInformationSection = false
   ClaimStatus.lblProcessorInformation.Click
   WaitForICallLoading
   If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Processor Information Section" ,false
       Exit Function
    End If    
    If Not IsNull(strClaimCategory) Then 
		If verifyInnerText (ClaimStatus.lblPIClaimCategory(),strClaimCategory, "ClaimCategory") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF 

    If Not IsNull(strCategoryType) Then 
		If verifyInnerText (ClaimStatus.lblPICategoryType(),strCategoryType, "CategoryType") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF 
	
    If Not IsNull(strUserBranchCode) Then 
		If  verifyInnerText (ClaimStatus.lblPIUserBranchCode(),strUserBranchCode, "UserBranchCode") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF 
	
	If Not IsNull(strUserProfitCode) Then 
		If  verifyInnerText (ClaimStatus.lblPIUserProfitCode(),strUserProfitCode, "UserProfitCode") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF
	
	If Not IsNull(strTransactionDateTime) Then 
		If  verifyInnerText (ClaimStatus.lblPITransactionDateTime(),strTransactionDateTime, "TransactionDateTime") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF	
	
	If Not IsNull(strValueDate) Then 
		If verifyInnerText (ClaimStatus.lblPIValueDate(),strValueDate, "ValueDate") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF	
	
	If Not IsNull(strTransfereeCIN) Then 
		If verifyInnerText (ClaimStatus.lblPITransfereeCIN(),strTransfereeCIN, "TransfereeCIN") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF	
	
	If Not IsNull(strTransfereeName) Then 
		If verifyInnerText (ClaimStatus.lblPITransfereeName(),strTransfereeName, "TransfereeName")  Then
    		verifyProcessorInformationSection = true
    	End If
	End IF

	If Not IsNull(strTransfereeAccountNumber) Then 
		If verifyInnerText (ClaimStatus.lblPITransfereeAccountNumber(),strTransfereeAccountNumber, "TransfereeAccountNumber") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF

	If Not IsNull(strSWCHorCCorGWRef) Then 
		If verifyInnerText (ClaimStatus.lblPISWCHCCGWRef(),strSWCHorCCorGWRef, "SWCHorCCorGWRef") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF

	If Not IsNull(strAcquirerName) Then 
		If verifyInnerText (ClaimStatus.lblPIAcquirerName(),strAcquirerName, "AcquirerName") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF
	
	If Not IsNull(strAcquirerBIN) Then 
		If verifyInnerText (ClaimStatus.lblPIAcquirerBIN(),strAcquirerBIN, "AcquirerBIN") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF
	
	If Not IsNull(strApprovedAmount) Then 
		If verifyInnerText (ClaimStatus.lblPIApprovedAmount(),strApprovedAmount, "ApprovedAmount") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF
	
	If Not IsNull(strTransactionComments) Then 
		If verifyInnerText (ClaimStatus.lblPITransactionComments(),strTransactionComments, "TransactionComments") Then
    		verifyProcessorInformationSection = true
    	End If
	End IF
	
   ClaimStatus.lblProcessorInformation.Click
   If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Processor Information Section" ,false
       verifyProcessorInformationSection = false
       Exit Function
   End If
    WaitForICallLoading
End Function

'[Verify the Payment Instructions]
Public Function VerifyPaymentInstructionsSection(lstlstPaymentInstructionsList)
	
	VerifyPaymentInstructionsSection = true
	ClaimStatus.lblPaymentInstructions.Click
	WaitForICallLoading
	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Payment Instructions Section" ,false
       VerifyPaymentInstructionsSection = false
       Exit Function
    End If     
     
    If Not IsNull(lstlstPaymentInstructionsList)Then
        VerifyPaymentInstructionsSection = verifyTableContentList(ClaimStatus.lblPaymentInstructionsHeader,ClaimStatus.lblPaymentInstructionsContent,lstlstPaymentInstructionsList,"PaymentInstructionsList",false,null,null,null)
    else
       VerifyPaymentInstructionsSection = false
       Exit Function
    End If     
	
	ClaimStatus.lblPaymentInstructions.Click
	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Payment Instructions Section" ,false
       VerifyPaymentInstructionsSection = false
       Exit Function
    End If
     WaitForICallLoading
     ClickOnCaseDetailOKButton()
     WaitForICallLoading
End Function

'[Click on Case Detail Ok Button]
Public Function ClickOnCaseDetailOKButton()
	ClickOnCaseDetailOKButton = true
	ClaimStatus.lblCaseDetailOKButton.Click
	If Err.Number<>0 Then
       ClickOnCaseDetailOKButton = false
       LogMessage "WARN","Verification","Failed to Click Case Detail Button : OK" ,false
   	End If
	WaitForICallLoading
End Function

'[Verify the claim status details displayed for the CIN Number selected by default]
Public Function VerifyCINCaseStatusInformationTable(lstlstCaseCINTransactionDetails)
    VerifyCINCaseStatusInformationTable=True
    If Not IsNull(lstlstCaseCINTransactionDetails)Then
       VerifyCINCaseStatusInformationTable = verifyTableContentList(ClaimStatus.tblCaseStatusInformationHeader,ClaimStatus.tblCaseStatusInformationContent,lstlstCaseCINTransactionDetails,"DefaultCaseTransactionDetails",false,null,null,null)
    else
       VerifyCINCaseStatusInformationTable = false
       Exit function
    End If
    WaitForICallLoading
 End Function

'[Click on Refresh Button]
Public Function ClickOnRefreshButton()
	ClickOnRefreshButton = true
	ClaimStatus.lblRefreshButton.Click
	If Err.Number<>0 Then
       ClickOnRefreshButton=false
       LogMessage "WARN","Verification","Failed to Click Button : Refresh" ,false
       Exit Function
   	End If
   	WaitForICallLoading
End Function

'[Verify the Account No Error Message]
Public Function VerifyAccountNoErrMessage(strAccErrMsg)
	VerifyAccountNoErrMessage = true	
	If Not IsNull(strAccErrMsg)Then
	  VerifyAccountNoErrMessage = verifyInnerText (ClaimStatus.lblAccountNoErrMessage(),strAccErrMsg, "AccountNoErrorMsg")
	Else
	  VerifyAccountNoErrMessage = false
	  Exit function
	End If
	WaitForICallLoading
End Function

'[Select the Caller CIN Radio Button]
Public Function SelectCallerCINRadioButton()	
	SelectCallerCINRadioButton = true
    ClaimStatus.lblCallerCInRadioButton.Click    
	If Err.Number<>0 Then
       SelectCallerCINRadioButton=false
       LogMessage "WARN","Verification","Failed to Click Radio Button : Caller CIN" ,false
       Exit Function
   	End If
	WaitForICallLoading	
End Function

'[Verify the Caller CIN Error Message]
Public Function VerifyCallerCINErrorMesg(strCallerCINErrMsg)
	VerifyCallerCINErrorMesg = true
		If Not IsNull(strCallerCINErrMsg)Then
		   VerifyCallerCINErrorMesg = verifyInnerText (ClaimStatus.lblCallerCINErrMsg(),strCallerCINErrMsg, "CallerCINErrMsg")
		else
		  VerifyCallerCINErrorMesg = false
		  Exit Function
		End If
		WaitForICallLoading
End Function

'[Verify Info warn Icon should be enabled]
Public Function VerifyInfoWarnIcon()
	If ClaimStatus.lblInfoWarnIcon.Exist(1) Then	   
       LogMessage "RSLT", "Verification","Info Warn Icon is Present and Enabled State", True
       VerifyInfoWarnIcon = True
    else
       LogMessage "WARN", "Verification","Info Warn Icon is visible at Claim Status Tab and not Enabled State", False
       VerifyInfoWarnIcon = False	
	End If
	WaitForICallLoading
End Function

'[Click on Info Warn Icon]
Public Function ClickOninfoWarnIcon()
  ClickOninfoWarnIcon = true
  ClaimStatus.lblInfoWarnIcon.Click
  If Err.Number<>0 Then
       ClickOninfoWarnIcon = false
       LogMessage "WARN","Verification","Failed to Click Info Warn Icon : Claim Status" ,false
       Exit Function
  End If
  WaitForICallLoading	
End Function

'[Verify Info warn Message when no claims records available for a customer]
Public Function verifyInfoWarnMeaage(strinfoWarnCaption, strinfoWarnMsg)
	verifyInfoWarnMeaage = false
	If Not IsNull(strinfoWarnCaption) and Not IsNull(strinfoWarnMsg) Then
	   If verifyInnerText (ClaimStatus.lblinfoWarnCaptionSpan(),strinfoWarnCaption, "ClaimDeatils") and verifyInnerText (ClaimStatus.lblinfoWarnMsgSpan(),strinfoWarnMsg, "Info warn Message") Then
	   	  verifyInfoWarnMeaage = true
	   End If
	End If
    ClaimStatus.lblInfoWarnOKButton().Click
     If Err.Number<>0 Then
       verifyInfoWarnMeaage = false
       LogMessage "WARN","Verification","Failed to Click Info Warn Ok Button : Claim Status" ,false
       Exit Function
  End If
    WaitForICallLoading
End Function
