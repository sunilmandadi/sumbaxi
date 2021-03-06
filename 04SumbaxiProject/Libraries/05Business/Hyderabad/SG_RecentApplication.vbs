
'[Select Status link in Table Recent Application Details]
Public Function clickRecentApplicationTable(lstRecentApplicationDetails)	
	WaitForICallLoading
	clickRecentApplicationTable=selectTableLink(RA.tblRecentApplicationHeader,RA.tblRecentApplicationContents,lstRecentApplicationDetails,"Recent Application Details" ,"Status",True,RA.lnkNext ,RA.lnkNext1 ,RA.lnkPrevious)
End Function

'[Verify row Data in Table Memo for Recent Application]
Public Function verifytblMemo_RA(arrRowDataList) 
	WaitForICallLoading
   verifytblMemo_RA=verifyTableContentList(RA.tblMemoHeader,RA.tblMemoContents,arrRowDataList,"Memo",false,null ,null,null)
End Function

'[Verify Pagination in Memo Table for Recent Application]
Public Function verifyPagination_MemoDetails()
	bverifyPagination=true	
	Dim intRow,intRowtemp
	intRow=0
	bNextPageExist=true
	If bPagination Then
		intRowTemp=0
		'Set newTable=tblContent
		intRowTemp=getRecordsCountForColumn (RA.tblRecentApplicationHeader,RA.tblRecentApplicationContents,"Memo Details")
		If intRowTemp<=10 Then
			LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to 10", true
		
		Else
			LogMessage "WARN","Verification","Number of records displayed per page is more than 10 record. Expected Count is less than or equal to 10, Actual "&intRowTemp, false
			bverifyPagination=false
		End If

		'********* To validate next link should not enable if record is less than specified number
		If intRowTemp < 10 Then
			bNextPageExist =matchStr(lnkNext1().GetROProperty("outerhtml"),"disabled")
			If Not bNextPageExist Then
				LogMessage "WARN","Verification","Next link expected to be disable if record is less than 10. Currently it is enable.",false
				bverifyPagination=false
			Else
				LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
				bverifyPagination=true
			End If
		End If
	End If
	verifyPagination_MemoDetails=bverifyPagination
End Function

'[Select Radio button in Audit Log for Recent Application]
Public Function selectRadioButton_AL(strAuditLog)
	selectRadioButton_AL=true
	selectRadioButton_AL=SelectRadioButtonGrp(strAuditLog,RA.rbtnGroup_AuditLog, Array("Application Log","Customer Event Log"))   
	If Err.Number<>0 Then
       selectRadioButton_AL=false
       LogMessage "WARN","Verification","Failed to Click Radio Button : Audit Log Radio Button" ,false
       Exit Function
   End If   
End Function

'[Verify row Data in Table Audit Log for Recent Application]
Public Function verifytblAuditLog_RA(arrRowDataList)  
   verifytblAuditLog_RA=verifyTableContentList(RA.tblAuditLogHeader,RA.tblAuditLogContents,arrRowDataList,"Audit Log" , false,null ,null,null)
End Function

'[Verify Applicant Details in Key Info for Recent Application]
Public Function verifyApplicantDetails_RA(lstApplicantDetails)
	bverifyApplicantDetails_RA=true
	
	If Not IsNull (lstApplicantDetails) then
		strCIN = lstApplicantDetails(0)
		strLogo = lstApplicantDetails(1)
		strCardIndicator = lstApplicantDetails(2)
		strCardNumber = lstApplicantDetails(3)
		strCardHolderFlag = lstApplicantDetails(4)
		strApprovalDeclineCode = lstApplicantDetails(5)
		strMotherMaidenName = lstApplicantDetails(6)
		strSourceCode = lstApplicantDetails(7)
		strCampaignCode = lstApplicantDetails(8)
		strAgentCode = lstApplicantDetails(9)
		'Blank section below Application details
		strOrganization = lstApplicantDetails(10)
		strType = lstApplicantDetails(11)
		strLastUpdateDate = lstApplicantDetails(12)
		strLastUpdateBy = lstApplicantDetails(13)
		
		If Not IsNull(strCIN) Then
			If Not VerifyInnerText(RA.lblCIN_KeyInfo(),strCIN,"CIN") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strLogo) Then
			If Not VerifyInnerText(RA.lblLogo_KeyInfo(),strLogo,"Logo") Then
				bverifyApplicantDetails_RA = False
			End If
        End If       
        If Not IsNull(strCardIndicator) Then
			If Not VerifyInnerText(RA.lblCardIndicator_KeyInfo(),strCardIndicator,"Card Face Indicator") Then
				bverifyApplicantDetails_RA = False
			End If
        End If        
        If Not IsNull(strCardNumber) Then
			If Not VerifyInnerText(RA.lblCardNumber_KeyInfo(),strCardNumber,"Card Number") Then
				bverifyApplicantDetails_RA = False
			End If
        End If        
        If Not IsNull(strCardHolderFlag) Then
			If Not VerifyInnerText(RA.lblCardHolderFlag_KeyInfo(),strCardHolderFlag,"Card Holder Flag") Then
				bverifyApplicantDetails_RA = False
			End If
        End If        
        If Not IsNull(strApprovalDeclineCode) Then
			If Not VerifyInnerText(RA.lblApprovalDeclineCode_KeyInfo(),strApprovalDeclineCode,"Approval\Decline\Cancel Code") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strMotherMaidenName) Then
			If Not VerifyInnerText(RA.lblMotherMaidenName_KeyInfo(),strMotherMaidenName,"Mother's Maiden Name") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strSourceCode) Then
			If Not VerifyInnerText(RA.lblSourceCode_KeyInfo(),strSourceCode,"Source Code") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strCampaignCode) Then
			If Not VerifyInnerText(RA.lblCampaignCode_KeyInfo(),strCampaignCode,"Campaign") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strAgentCode) Then
			If Not VerifyInnerText(RA.lblAgentCode_KeyInfo(),strAgentCode,"Agent Code") Then
				bverifyApplicantDetails_RA = False
			End If
        End If   
        If Not IsNull(strOrganization) Then
			If Not VerifyInnerText(RA.lblOrganization_KeyInfo(),strOrganization,"Organization") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strType) Then
			If Not VerifyInnerText(RA.lblType_KeyInfo(),strType,"Type") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strLastUpdateDate) Then
			If Not VerifyInnerText(RA.lblLastUpdateDate_KeyInfo(),strLastUpdateDate,"Last Update Date") Then
				bverifyApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strLastUpdateBy) Then
			If Not VerifyInnerText(RA.lblLastUpdateBy_KeyInfo(),strLastUpdateBy,"Last Update By") Then
				bverifyApplicantDetails_RA = False
			End If
        End If        
    End If
    verifyApplicantDetails_RA=bverifyApplicantDetails_RA
End Function

'[Verify Co-Applicant Details in Key Info for Recent Application]
Public Function verifyCoApplicantDetails_RA(lstCoApplicantDetails)
	bverifyCoApplicantDetails_RA=true
	
	If Not IsNull (lstCoApplicantDetails) then
		strCIN = lstCoApplicantDetails(0)
		strLogo = lstCoApplicantDetails(1)
		strCardIndicator = lstCoApplicantDetails(2)
		strCardNumber = lstCoApplicantDetails(3)
		strCardHolderFlag = lstCoApplicantDetails(4)
		strApprovalDeclineCode = lstCoApplicantDetails(5)
		strMotherMaidenName = lstCoApplicantDetails(6)
		strSourceCode = lstCoApplicantDetails(7)
		strCampaignCode = lstCoApplicantDetails(8)
		strAgentCode = lstCoApplicantDetails(9)
		
		If Not IsNull(strCIN) Then
			If Not VerifyInnerText(RA.lblCIN_CoApplicant(),strCIN,"CIN") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strLogo) Then
			If Not VerifyInnerText(RA.lblLogo_CoApplicant(),strLogo,"Logo") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If       
        If Not IsNull(strCardIndicator) Then
			If Not VerifyInnerText(RA.lblCardIndicator_CoApplicant(),strCardIndicator,"Card Face Indicator") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If        
        If Not IsNull(strCardNumber) Then
			If Not VerifyInnerText(RA.lblCardNumberSUP1_CoApplicant(),strCardNumber,"Card Number") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If        
        If Not IsNull(strCardHolderFlag) Then
			If Not VerifyInnerText(RA.lblCardHolderFlag_CoApplicant(),strCardHolderFlag,"Card Holder Flag") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If        
        If Not IsNull(strApprovalDeclineCode) Then
			If Not VerifyInnerText(RA.lblApprovalDeclineCode_CoApplicant(),strApprovalDeclineCode,"Approval\Decline\Cancel Code") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strMotherMaidenName) Then
			If Not VerifyInnerText(RA.lblMotherMaidenName_CoApplicant(),strMotherMaidenName,"Mother's Maiden Name") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strSourceCode) Then
			If Not VerifyInnerText(RA.lblSourceCode_CoApplicant(),strSourceCode,"Source Code") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strCampaignCode) Then
			If Not VerifyInnerText(RA.lblCampaignCode_CoApplicant(),strCampaignCode,"Campaign") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If
        If Not IsNull(strAgentCode) Then
			If Not VerifyInnerText(RA.lblAgentCode_CoApplicant(),strAgentCode,"Agent Code") Then
				bverifyCoApplicantDetails_RA = False
			End If
        End If         
    End If
    verifyCoApplicantDetails_RA=bverifyCoApplicantDetails_RA
End Function

'[Verify Guarantor Details in Key Info for Recent Application]
Public Function verifyGuarantorDetails_RA(lstGuarantorDetails)
	bverifyGuarantorDetails_RA =true
	
	If Not IsNull (lstGuarantorDetails) then
		strCIN = lstGuarantorDetails(0)
		strLogo = lstGuarantorDetails(1)
		strCardIndicator = lstGuarantorDetails(2)
		strCardNumber = lstGuarantorDetails(3)
		strCardHolderFlag = lstGuarantorDetails(4)
		strApprovalDeclineCode = lstGuarantorDetails(5)
		strMotherMaidenName = lstGuarantorDetails(6)
		strSourceCode = lstGuarantorDetails(7)
		strCampaignCode = lstGuarantorDetails(8)
		strAgentCode = lstGuarantorDetails(9)
		
		If Not IsNull(strCIN) Then
			If Not VerifyInnerText(RA.lblCIN_Guarantor(),strCIN,"CIN") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If
        If Not IsNull(strLogo) Then
			If Not VerifyInnerText(RA.lblLogo_Guarantor(),strLogo,"Logo") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If       
        If Not IsNull(strCardIndicator) Then
			If Not VerifyInnerText(RA.lblCardIndicator_Guarantor(),strCardIndicator,"Card Face Indicator") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If        
        If Not IsNull(strCardNumber) Then
			If Not VerifyInnerText(RA.lblCardNumberSUP2_Guarantor(),strCardNumber,"Card Number") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If        
        If Not IsNull(strCardHolderFlag) Then
			If Not VerifyInnerText(RA.lblCardHolderFlag_Guarantor(),strCardHolderFlag,"Card Holder Flag") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If        
        If Not IsNull(strApprovalDeclineCode) Then
			If Not VerifyInnerText(RA.lblApprovalDeclineCode_Guarantor(),strApprovalDeclineCode,"Approval\Decline\Cancel Code") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If
        If Not IsNull(strMotherMaidenName) Then
			If Not VerifyInnerText(RA.lblMotherMaidenName_Guarantor(),strMotherMaidenName,"Mother's Maiden Name") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If
        If Not IsNull(strSourceCode) Then
			If Not VerifyInnerText(RA.lblSourceCode_Guarantor(),strSourceCode,"Source Code") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If
        If Not IsNull(strCampaignCode) Then
			If Not VerifyInnerText(RA.lblCampaignCode_Guarantor(),strCampaignCode,"Campaign") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If
        If Not IsNull(strAgentCode) Then
			If Not VerifyInnerText(RA.lblAgentCode_Guarantor(),strAgentCode,"Agent Code") Then
				bverifyGuarantorDetails_RA = False
			End If
        End If         
    End If
	
    verifyGuarantorDetails_RA=bverifyGuarantorDetails_RA
End Function

'[Verify Personal Details in Application Details for Recent Application]
Public Function verifyPersonaldetails_RA(lstPersonalDetails)
	bverifyPersonaldetails = True
	
	If Not IsNull (lstPersonalDetails) then
		strSalutation = lstPersonalDetails(0)
		strCustomerName = lstPersonalDetails(1)
		strNRICNo = lstPersonalDetails(2)
		strDOB = lstPersonalDetails(3)
		strNationality = lstPersonalDetails(4)
		strPRStatus = lstPersonalDetails(5)
		strGender = lstPersonalDetails(6)
		strMaritalStatus = lstPersonalDetails(7)
		strEducation = lstPersonalDetails(8)
		strNoDependents = lstPersonalDetails(9)
		strMobileNo = lstPersonalDetails(10)
		strHomeNo = lstPersonalDetails(11)
		strOfficeNo = lstPersonalDetails(12)
		strEmailAddress = lstPersonalDetails(13)
		
		If Not IsNull(strSalutation) Then
			If Not VerifyInnerText(RA.lblSalutation(),strSalutation,"Salutation") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strCustomerName) Then
			If Not VerifyInnerText(RA.lblCustomerName(),strCustomerName,"Customer Name") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strNRICNo) Then
			If Not VerifyInnerText(RA.lblNRICNo(),strNRICNo,"NRIC Number") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strDOB) Then
			If Not VerifyInnerText(RA.lblDOB(),strDOB,"Date of Birth") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strNationality) Then
			If Not VerifyInnerText(RA.lblNationality(),strNationality,"Nationality") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strPRStatus) Then
			If Not VerifyInnerText(RA.lblPRStatus(),strPRStatus,"PR Status") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strEPExpiry) Then
			If Not VerifyInnerText(RA.lblEPExpiry(),strEPExpiry,"EP Expriy Date") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strGender) Then
			If Not VerifyInnerText(RA.lblGender(),strGender,"Gender") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strMaritalStatus) Then
			If Not VerifyInnerText(RA.lblMaritalStatus(),strMaritalStatus,"Marital Status") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strEducation) Then
			If Not VerifyInnerText(RA.lblEducation(),strEducation,"Education") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strNoDependents) Then
			If Not VerifyInnerText(RA.lblNoOfdependents(),strNoDependents,"No Of Dependents") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strMobileNo) Then
			If Not VerifyInnerText(RA.lblMobileNo(),strMobileNo,"Mobile Number") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strHomeNo) Then
			If Not VerifyInnerText(RA.lblHomeNo(),strHomeNo,"Home Number") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strOfficeNo) Then
			If Not VerifyInnerText(RA.lblOfficeNo(),strOfficeNo,"Office Number") Then
				bverifyPersonaldetails = False
			End If
        End If
        If Not IsNull(strEmailAddress) Then
			If Not VerifyInnerText(RA.lblEmailAddress(),strEmailAddress,"Email Address") Then
				bverifyPersonaldetails = False
			End If
        End If
    End If
    verifyPersonaldetails_RA = bverifyPersonaldetails
End Function

'[Verify Mailing Address Details in Application Details for Recent Application]
Public Function verifyMailingdetails_RA(lstMailingDetails)
	bverifyMailingdetails = True
	
	If Not IsNull (lstMailingDetails) then
		strCountry = lstMailingDetails(0)
		strPostalCode = lstMailingDetails(1)
		strBlockNo = lstMailingDetails(2)
		strStreetName1 = lstMailingDetails(3)
		strStreetName2 = lstMailingDetails(4)
		strResidentialStatus = lstMailingDetails(5)
		strResidentialType = lstMailingDetails(6)	
		
		If Not IsNull(strCountry) Then
			If Not VerifyInnerText(RA.lblCountry(),strCountry,"Country") Then
				bverifyMailingdetails = False
			End If
        End If
        If Not IsNull(strPostalCode) Then
			If Not VerifyInnerText(RA.lblPostalCode(),strPostalCode,"Postal Code") Then
				bverifyMailingdetails = False
			End If
        End If
        If Not IsNull(strBlockNo) Then
			If Not VerifyInnerText(RA.lblBlockNo(),strBlockNo,"Block Number") Then
				bverifyMailingdetails = False
			End If
        End If
        If Not IsNull(strStreetName1) Then
			If Not VerifyInnerText(RA.lblStreetName1(),strStreetName1,"Street Name 1") Then
				bverifyMailingdetails = False
			End If
        End If
        If Not IsNull(strStreetName2) Then
			If Not VerifyInnerText(RA.lblStreetName2(),strStreetName2,"Street Name 2") Then
				bverifyMailingdetails = False
			End If
        End If
        If Not IsNull(strResidentialStatus) Then
			If Not VerifyInnerText(RA.lblResidentialStatus(),strResidentialStatus,"Residential Status") Then
				bverifyMailingdetails = False
			End If
        End If
        If Not IsNull(strResidentialType) Then
			If Not VerifyInnerText(RA.lblResidentialType(),strResidentialType,"Residential Type") Then
				bverifyMailingdetails = False
			End If
        End If
    End If
    verifyMailingdetails_RA = bverifyMailingdetails
End Function

'[Verify Current Employment Details in Application Details for Recent Application]
Public Function verifyEmploymentdetails_RA(lstEmploymentDetails)
	bverifyEmploymentdetails = True

	If Not IsNull (lstEmploymentDetails) then
		strCompanyName = lstEmploymentDetails(0)
		strJobStatus = lstEmploymentDetails(1)
		strJobTitle = lstEmploymentDetails(2)
		strBusinessType = lstEmploymentDetails(3)
		strLengthCurEmp = lstEmploymentDetails(4)
		strPrevCompanyName = lstEmploymentDetails(5)
		strLengthPrevEmp = lstEmploymentDetails(6)	
		
		If Not IsNull(strCompanyName) Then
			If Not VerifyInnerText(RA.lblCompanyName(),strCompanyName,"Company Name") Then
				bverifyEmploymentdetails = False
			End If
        End If
        If Not IsNull(strJobStatus) Then
			If Not VerifyInnerText(RA.lblJobStatus(),strJobStatus,"Job Status") Then
				bverifyEmploymentdetails = False
			End If
        End If
        If Not IsNull(strJobTitle) Then
			If Not VerifyInnerText(RA.lblJobTitle(),strJobTitle,"JobT itle") Then
				bverifyEmploymentdetails = False
			End If
        End If
        If Not IsNull(strBusinessType) Then
			If Not VerifyInnerText(RA.lblBusinessType(),strBusinessType,"Industry / Business Type") Then
				bverifyEmploymentdetails = False
			End If
        End If
        If Not IsNull(strLengthCurEmp) Then
			If Not VerifyInnerText(RA.lblLengthCurEmp(),strLengthCurEmp,"Length of Current Emp") Then
				bverifyEmploymentdetails = False
			End If
        End If
        If Not IsNull(strPrevCompanyName) Then
			If Not VerifyInnerText(RA.lblPrevCompanyName(),strPrevCompanyName,"Previous Company Name") Then
				bverifyEmploymentdetails = False
			End If
        End If
        If Not IsNull(strLengthPrevEmp) Then
			If Not VerifyInnerText(RA.lblLengthPrevEmp(),strLengthPrevEmp,"Length of Previous Emp") Then
				bverifyEmploymentdetails = False
			End If
        End If
    End If
    verifyEmploymentdetails_RA = bverifyEmploymentdetails
End Function

'[Verify Income Submissions in Application Details for Recent Application]
Public Function verifyIncomeSubmission_RA(strIncomeSubmission)
	bverifyIncomeSubmission = True			
	If Not IsNull(strIncomeSubmission) Then
		If Not VerifyInnerText(RA.lblIncDoc(),strIncomeSubmission,"Income Doc") Then
			bverifyIncomeSubmission = False
		End If
    End If 
   verifyIncomeSubmission_RA = bverifyIncomeSubmission
  End Function

'[Verify For Bank Use Only in Application Details for Recent Application]
Public Function verifyBankUseOnly_RA(lstLoanDetails)
	bverifyBankUseOnly = True

	If Not IsNull (lstLoanDetails) then
		strCampaignCode = lstLoanDetails(0)
		strBranchCode = lstLoanDetails(1)
		strAgentCode = lstLoanDetails(2)
		strReferralCode = lstLoanDetails(3)
		strLoanApprovedAmt = lstLoanDetails(4)
		strApplicationStatus = lstLoanDetails(5)
		strReasons = lstLoanDetails(6)	
				
		If Not IsNull(strCampaignCode) Then
			If Not VerifyInnerText(RA.lblCampaignCode(),strCampaignCode,"Campaign Code") Then
				bverifyBankUseOnly = False
			End If
        End If
        If Not IsNull(strBranchCode) Then
			If Not VerifyInnerText(RA.lblBranchCode(),strBranchCode,"Branch Code") Then
				bverifyBankUseOnly = False
			End If
        End If
        If Not IsNull(strAgentCode) Then
			If Not VerifyInnerText(RA.lblAgentCode(),strAgentCode,"Agent Code") Then
				bverifyBankUseOnly = False
			End If
        End If
        If Not IsNull(strReferralCode) Then
			If Not VerifyInnerText(RA.lblReferralCode(),strReferralCode,"Referral Code") Then
				bverifyBankUseOnly = False
			End If
        End If
        If Not IsNull(strLoanApprovedAmt) Then
			If Not VerifyInnerText(RA.lblLoanApprovedAmt(),strLoanApprovedAmt,"Loan Approved Amount/ Credit Limit") Then
				bverifyBankUseOnly = False
			End If
        End If
        If Not IsNull(strApplicationStatus) Then
			If Not VerifyInnerText(RA.lblApplicationStatus(),strApplicationStatus,"Application Status") Then
				bverifyBankUseOnly = False
			End If
        End If
        If Not IsNull(strReasons) Then
			If Not VerifyInnerText(RA.lblReason(),strReasons,"Reason(s)") Then
				bverifyBankUseOnly = False
			End If
        End If
    End If
    verifyBankUseOnly_RA = bverifyBankUseOnly
End Function

'[Verify Loan Request Details in Application Details for Recent Application]
Public Function verifyLoanRequestDetails_RA(lstLoanDetails)
	bverifyLoanRequestDetails = True

	If Not IsNull (lstLoanDetails) then
		strRequestedLoanAmt = lstLoanDetails(0)
		strLoanTenure = lstLoanDetails(1)
		strProductType = lstLoanDetails(2)
		strCardNumber = lstLoanDetails(3)
		strEligibleLoanAmt = lstLoanDetails(4)
		strLoanType = lstLoanDetails(5)
		strInterestRate = lstLoanDetails(6)	
		strTransferAccount = lstLoanDetails(7)	
		strLoanServicingAccount = lstLoanDetails(8)	
				
		If Not IsNull(strRequestedLoanAmt) Then
			If Not VerifyInnerText(RA.lblReqLoanAmt(),strRequestedLoanAmt,"Requested Loan Amount/ Credit Limit") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strLoanTenure) Then
			If Not VerifyInnerText(RA.lblLoanTenure(),strLoanTenure,"Loan Tenure (in mths)") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strProductType) Then
			If Not VerifyInnerText(RA.lblPdtType(),strProductType,"Type of Product") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strCardNumber) Then
			If Not VerifyInnerText(RA.lblCardNo(),strCardNumber,"Cashline/ Credit Card No") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strEligibleLoanAmt) Then
			If Not VerifyInnerText(RA.lblLoanAmt(),strEligibleLoanAmt,"Eligible Loan Amount(SGD)") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strLoanType) Then
			If Not VerifyInnerText(RA.lblLoanType(),strLoanType,"Type of Loan") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strInterestRate) Then
			If Not VerifyInnerText(RA.lblInterestRate(),strInterestRate,"Interest Rate") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strTransferAccount) Then
			If Not VerifyInnerText(RA.lblTransferAcc(),strTransferAccount,"Transfer to Account") Then
				bverifyLoanRequestDetails = False
			End If
        End If
        If Not IsNull(strLoanServicingAccount) Then
			If Not VerifyInnerText(RA.lblServicingAccNo(),strLoanServicingAccount,"Loan Servicing Account No/ Account No") Then
				bverifyLoanRequestDetails = False
			End If
        End If
    End If
    verifyLoanRequestDetails_RA = bverifyLoanRequestDetails
End Function

'[Click Button Close in Recent Application Page displayed] 
Public Function ClickCloseButton_RA()
	RA.btnClose.click
	If Err.Number<>0 Then
       LogMessage "WARN","Verification","Failed to Click Button : Close",false
       ClickCloseButton_RA = False
       Exit Function
    End If
	   ClickCloseButton_RA = true
	   WaitForICallLoading
End Function

