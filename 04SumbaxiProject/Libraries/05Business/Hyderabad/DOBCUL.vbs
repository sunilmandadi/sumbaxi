'*******************Added by Kalyan DOBCUL Home Page 21032017 ****************************

'[Select DOBCUL Product Type Value as]
Public Function selectDOBCUL_ProductTyp(strProdTyp)
	bselectDOBCUL_ProductTyp=true
	If Not IsNull(strProdTyp) Then
       If Not (selectItem_Combobox (DOBCUL.searchTabProdType_comboBox(), strProdTyp))Then
           LogMessage "WARN","Verification","Failed to select :"&strProdTyp&" From Product Type dropdown list" ,false
           bselectDOBCUL_ProductTyp = false
       End If
   End If
   WaitForICallLoading
   selectDOBCUL_ProductTyp=bselectDOBCUL_ProductTyp
End Function 
'
'[Select DOBCULNRIC Value as]
Public Function selectDOBCUL_NRIC(strNRIC)
	bsetDOBCUL_NRIC=true
	If Not IsNull(strNRIC) Then
      DOBCUL.searchTabNRIC_textBox.set strNRIC
      bselectDOBCUL_NRIC=true
   End If
   WaitForICallLoading
   selectDOBCUL_NRIC=bselectDOBCUL_NRIC
End Function

'[Verify row Data in Table DOBCUL]
Public Function verifytblDOBCUL_Records_RowData(arrRowDataList)
   bDevPending=false
   verifytblDOBCUL_Records_RowData=verifyTableContentList(DOBCUL.donData_table_header,DOBCUL.donData_table_content,arrRowDataList,"DOBCUL_Records" ,True,DOBCUL.lnkNext ,DOBCUL.lnkNext1,DOBCUL.lnkPrevious)
End Function

'[Click Status link in Table DOBCUL Records form Home Page]
Public Function clickSts_DOBCUL_Records(lstDOBCLData)
   bDevPending=false
   With CSO_TM_Home
	   clickSts_DOBCUL_Records=selectTableLink(DOBCUL.donData_table_header,DOBCUL.donData_table_content,lstDOBCLData,"DOBCUL_Records" ,"Status",True,DOBCUL.lnkNext ,DOBCUL.lnkNext1,DOBCUL.lnkPrevious)
   End With
End Function

'[Verify DOBCUL Product Details Section]
Public Function verifyDOBCUL_ProductDtlsSec(lstProductDetls)
	bverifyDOBCUL_ProductDtlsSec=true
	intSize=Ubound(lstProductDetls)
		'For i = 0 To intSize - 1
			If Not IsNull(lstProductDetls(0)) Then
				If not VerifyInnerText(DOBCUL.don_AppReqLoanAmt_span, lstProductDetls(0),"Requested Loan Amount") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(1)) Then
				If not VerifyInnerText(DOBCUL.don_AppLoanTenure_span, lstProductDetls(1),"Loan Tenure") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(2)) Then
				If not VerifyInnerText(DOBCUL.don_AppIntRate_span, lstProductDetls(2),"Interest Rate") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(3)) Then
				If not VerifyInnerText(DOBCUL.don_AppInstPerMonth_span, lstProductDetls(3),"Instalment Per Month") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(4)) Then
				If not VerifyInnerText(DOBCUL.don_AppDisbursedAcc_span, lstProductDetls(4),"Disbursed Account") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(5)) Then
				If not VerifyInnerText(DOBCUL.don_AppPrCreditLimit_span, lstProductDetls(5),"Preferred Credit Limit") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(6)) Then
				If not VerifyInnerText(DOBCUL.don_AppPrCashLimit_span, lstProductDetls(6),"Preferred Cashline Limit") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(7)) Then
				If not VerifyInnerText(DOBCUL.don_AppChBkRequired_span, lstProductDetls(7),"Cheque Book Required") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(8)) Then
				If not VerifyInnerText(DOBCUL.don_AppSignAcc_span, lstProductDetls(8),"Signature Account") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(9)) Then
				If not VerifyInnerText(DOBCUL.don_AppCardToRepl_span, lstProductDetls(9),"Card Number to be Replaced") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(10)) Then
				If not VerifyInnerText(DOBCUL.don_AppCardToReplInd_span, lstProductDetls(10),"Card Replacement Indicator") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(11)) Then
				If not VerifyInnerText(DOBCUL.don_AppTmpLimIncrAmt_span, lstProductDetls(11),"Temp Limit Increase Amount") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(12)) Then
				If not VerifyInnerText(DOBCUL.don_AppPurpOfLimIncr_span, lstProductDetls(12),"Purpose of Limit Increase") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
			If Not IsNull(lstProductDetls(13)) Then
				If not VerifyInnerText(DOBCUL.don_AppInstlmntPeriod_span, lstProductDetls(13),"Instalment Period/Tenure") Then
					bverifyDOBCUL_ProductDtlsSec=false
				End If
			End If
			
		'Next
	verifyDOBCUL_ProductDtlsSec=bverifyDOBCUL_ProductDtlsSec
End Function

'[Verify DOBCUL Employment Details Section]
Public Function verifyDOBCUL_EmpDtlsSec(lstEmpDetls)
	bverifyDOBCUL_EmpDtlsSec=true
	intSize=Ubound(lstEmpDetls)
	'For i = 0 To intSize - 1
		If Not IsNull(lstEmpDetls(0)) Then
			If not VerifyInnerText(DOBCUL.don_AppJobStatus_span, lstEmpDetls(0),"don_AppJobStatus_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If 
		End IF
		If Not IsNull(lstEmpDetls(1)) Then
			If not VerifyInnerText(DOBCUL.don_AppJobTitle_span, lstEmpDetls(1),"don_AppJobTitle_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(2)) Then
			If not VerifyInnerText(DOBCUL.don_AppJobIndustry_span, lstEmpDetls(2),"don_AppJobIndustry_span") Then
			bverifyDOBCUL_EmpDtlsSec=false
		    End If
		End If
		If Not IsNull(lstEmpDetls(3)) Then
			If not VerifyInnerText(DOBCUL.don_AppJobPosition_span, lstEmpDetls(3),"don_AppJobPosition_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(4)) Then
			If not VerifyInnerText(DOBCUL.don_AppMonthlyIncome_span, lstEmpDetls(4),"don_AppMonthlyIncome_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(5)) Then
			If not VerifyInnerText(DOBCUL.don_AppVariableIncome_span, lstEmpDetls(5),"don_AppVariableIncome_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(6)) Then
			If not VerifyInnerText(DOBCUL.don_AppCompanyName_span, lstEmpDetls(6),"don_AppCompanyName_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(7)) Then
			If not VerifyInnerText(DOBCUL.don_AppLenOfCurrentEmp_span, lstEmpDetls(7),"don_AppLenOfCurrentEmp_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(8)) Then
			If not VerifyInnerText(DOBCUL.don_AppPreviousCompany_span, lstEmpDetls(8),"don_AppPreviousCompany_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(9)) Then
			If not VerifyInnerText(DOBCUL.don_AppLenOfPrevEmp_span, lstEmpDetls(9),"don_AppLenOfPrevEmp_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
		If Not IsNull(lstEmpDetls(10)) Then
			If not VerifyInnerText(DOBCUL.don_AppAnnualIncome_span, lstEmpDetls(10),"don_AppAnnualIncome_span") Then
				bverifyDOBCUL_EmpDtlsSec=false
			End If
		End IF	
	'Next
	verifyDOBCUL_EmpDtlsSec=bverifyDOBCUL_EmpDtlsSec
End Function

'[Verify DOBCUL Mailing Address Details Section]
Public Function verifyDOBCUL_MailAddrssDtlsSec(lstMailAdrssDetls)
	bverifyDOBCUL_MailAddrssDtlsSec=true
	intSize=Ubound(lstMailAdrssDetls)
	'For i = 0 To intSize - 1
	If Not IsNull(lstMailAdrssDetls(0)) Then
		If not VerifyInnerText(DOBCUL.don_AppBillingAddress_span, lstMailAdrssDetls(0),"don_AppBillingAddress_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
	
	If Not IsNull(lstMailAdrssDetls(1)) Then
		If not VerifyInnerText(DOBCUL.don_AppPostCode_span, lstMailAdrssDetls(1),"don_AppPostCode_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(2)) Then
		If not VerifyInnerText(DOBCUL.don_AppBlock_span, lstMailAdrssDetls(2),"don_AppBlock_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(3)) Then
		If not VerifyInnerText(DOBCUL.don_AppUnit_span, lstMailAdrssDetls(3),"don_AppUnit_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(4)) Then
		If not VerifyInnerText(DOBCUL.don_AppLevel_span, lstMailAdrssDetls(4),"don_AppLevel_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(5)) Then
		If not VerifyInnerText(DOBCUL.don_AppAddrLineList0_span, lstMailAdrssDetls(5),"don_AppAddrLineList0_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(6)) Then
		If not VerifyInnerText(DOBCUL.don_AppCountry_span, lstMailAdrssDetls(6),"don_AppCountry_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(7)) Then
		If not VerifyInnerText(DOBCUL.don_AppResidentialType_span, lstMailAdrssDetls(7),"don_AppResidentialType_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(8)) Then
'		If not VerifyInnerText(DOBCUL.don_AppResidentialStatus_span, lstMailAdrssDetls(8),"don_AppResidentialStatus_span") Then
'			bverifyDOBCUL_MailAddrssDtlsSec=false
'		End If
		If not VerifyInnerText(DOBCUL.don_AppResidentialStatus_span, lstMailAdrssDetls(9),"don_AppResidentialStatus_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstMailAdrssDetls(9)) Then
		If not VerifyInnerText(DOBCUL.don_AppLengthOfStay_span, lstMailAdrssDetls(10),"don_AppLengthOfStay_span") Then
			bverifyDOBCUL_MailAddrssDtlsSec=false
		End If
	End If
	'Next
	verifyDOBCUL_MailAddrssDtlsSec=bverifyDOBCUL_MailAddrssDtlsSec
End Function

'[Verify DOBCUL Address Details Section]
Public Function verifyDOBCUL_AddrssDtlsSec(lstAdrssDetls)
	bverifyDOBCUL_AddrssDtlsSec=true
	intSize=Ubound(lstAdrssDetls)
		If Not IsNull(lstAdrssDetls(0)) Then
			If not VerifyInnerText(DOBCUL.don_AppCompAddrDet_PostCode_span, lstAdrssDetls(0),"don_AppCompAddrDet_PostCode_span") Then
				bverifyDOBCUL_AddrssDtlsSec=false
			End If 
		End If 
		If Not IsNull(lstAdrssDetls(1)) Then
			If not VerifyInnerText(DOBCUL.don_AppCompAddrDet_Block_span, lstAdrssDetls(1),"don_AppCompAddrDet_Block_span") Then
				bverifyDOBCUL_AddrssDtlsSec=false
			End If
		End If 
		If Not IsNull(lstAdrssDetls(2)) Then
			If not VerifyInnerText(DOBCUL.don_AppCompAddrDet_Unit_span, lstAdrssDetls(2),"don_AppCompAddrDet_Unit_span") Then
				bverifyDOBCUL_AddrssDtlsSec=false
			End If
		End If 
		If Not IsNull(lstAdrssDetls(3)) Then
			If not VerifyInnerText(DOBCUL.don_AppCompAddrDet_Level_span, lstAdrssDetls(3),"don_AppCompAddrDet_Level_span") Then
				bverifyDOBCUL_AddrssDtlsSec=false
			End If
		End If
		If Not IsNull(lstAdrssDetls(4)) Then		
			If not VerifyInnerText(DOBCUL.don_compAppAddrLineList0_span, lstAdrssDetls(4),"don_compAppAddrLineList0_span") Then
				bverifyDOBCUL_AddrssDtlsSec=false
			End If
		End If 
	verifyDOBCUL_AddrssDtlsSec=bverifyDOBCUL_AddrssDtlsSec
End Function

'[Verify DOBCUL Personal Details Section]
Public Function verifyDOBCUL_PersonalDtlsSec(lstPrsnlDetls)
	bverifyDOBCUL_PersonalDtlsSec=true
	intSize=Ubound(lstPrsnlDetls)
	'For i = 0 To intSize - 1
	
	If Not IsNull(lstPrsnlDetls(0)) Then
		If not VerifyInnerText(DOBCUL.don_AppSalutation_span, lstPrsnlDetls(0),"don_AppSalutation_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End if
	
	If Not IsNull(lstPrsnlDetls(1)) Then
		If not VerifyInnerText(DOBCUL.don_AppName_span, lstPrsnlDetls(1),"don_AppName_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(2)) Then
		If not VerifyInnerText(DOBCUL.don_AppNricPassport_span, lstPrsnlDetls(2),"don_AppNricPassport_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(3)) Then
		If not VerifyInnerText(DOBCUL.don_AppDob_span, lstPrsnlDetls(3),"don_AppDob_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(4)) Then
		If not VerifyInnerText(DOBCUL.don_AppNationality_span, lstPrsnlDetls(4),"don_AppNationality_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(5)) Then
		If not VerifyInnerText(DOBCUL.don_AppEthnicity_span, lstPrsnlDetls(5),"don_AppEthnicity_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(6)) Then
		If not VerifyInnerText(DOBCUL.don_AppPRStatus_span, lstPrsnlDetls(6),"don_AppPRStatus_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(7)) Then
		If not VerifyInnerText(DOBCUL.don_AppGender_span, lstPrsnlDetls(7),"don_AppGender_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(8)) Then
		If not VerifyInnerText(DOBCUL.don_AppMaritalStatus_span, lstPrsnlDetls(8),"don_AppMaritalStatus_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(9)) Then
		If not VerifyInnerText(DOBCUL.don_AppHighestEducationLevel_span, lstPrsnlDetls(9),"don_AppHighestEducationLevel_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(10)) Then
		If not VerifyInnerText(DOBCUL.don_AppNoOfDependents_span, lstPrsnlDetls(10),"don_AppNoOfDependents_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(11)) Then
		If not VerifyInnerText(DOBCUL.don_AppMobileNumber_span, lstPrsnlDetls(11),"don_AppMobileNumber_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(12)) Then
		If not VerifyInnerText(DOBCUL.don_AppHomeNumber_span, lstPrsnlDetls(12),"don_AppHomeNumber_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(13)) Then
		If not VerifyInnerText(DOBCUL.don_AppOfficeNumber_span, lstPrsnlDetls(13),"don_AppOfficeNumber_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(14)) Then
		If not VerifyInnerText(DOBCUL.don_AppEmail_span, lstPrsnlDetls(14),"don_AppEmail_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(15)) Then
		If not VerifyInnerText(DOBCUL.don_AppNameOfTerInst__span, lstPrsnlDetls(15),"don_AppNameOfTerInst__span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(16)) Then
		If not VerifyInnerText(DOBCUL.don_AppExpctdYrOfGrad_span, lstPrsnlDetls(16),"don_AppExpctdYrOfGrad_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(17)) Then
		If not VerifyInnerText(DOBCUL.don_AppNUSSMemNum_span, lstPrsnlDetls(17),"don_AppNUSSMemNum_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstPrsnlDetls(18)) Then
		If not VerifyInnerText(DOBCUL.don_AppStudOrStaffNum_span, lstPrsnlDetls(18),"don_AppStudOrStaffNum_span") Then
			bverifyDOBCUL_PersonalDtlsSec=false
		End If
	End If
	'Next
	verifyDOBCUL_PersonalDtlsSec=bverifyDOBCUL_PersonalDtlsSec
End Function

'[Verify DOBCUL Bank Use Details Section]
Public Function verifyDOBCUL_BnkUseDtlsSec(lstBnkDetls)
	bverifyDOBCUL_BnkUseDtlsSec=true
	intSize=Ubound(lstBnkDetls)
	'For i = 0 To intSize - 1
	If Not IsNull(lstBnkDetls(0)) Then
		If not VerifyInnerText(DOBCUL.don_AppCampaignCode_span, lstBnkDetls(0),"don_AppCampaignCode_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If 
	End If 
	If Not IsNull(lstBnkDetls(1)) Then
		If not VerifyInnerText(DOBCUL.don_AppPromoCode_span, lstBnkDetls(1),"don_AppPromoCode_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(2)) Then
		If not VerifyInnerText(DOBCUL.don_AppBranchCode_span, lstBnkDetls(2),"don_AppBranchCode_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(3)) Then
		If not VerifyInnerText(DOBCUL.don_AppAgentCode_span, lstBnkDetls(3),"don_AppAgentCode_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(4)) Then
		If not VerifyInnerText(DOBCUL.don_AppReferralCode_span, lstBnkDetls(4),"don_AppReferralCode_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(5)) Then
		If not VerifyInnerText(DOBCUL.don_AppAppLoanAmount_span, lstBnkDetls(5),"don_AppAppLoanAmount_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(6)) Then
		If not VerifyInnerText(DOBCUL.don_AppApplicationStatus_span, lstBnkDetls(6),"don_AppApplicationStatus_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(7)) Then
		If not VerifyInnerText(DOBCUL.don_AppApplicationNumber_span, lstBnkDetls(7),"don_AppApplicationNumber_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(8)) Then
		If not VerifyInnerText(DOBCUL.don_AppTxnReference_span, lstBnkDetls(8),"don_AppTxnReference_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(9)) Then
		If not VerifyInnerText(DOBCUL.don_AppChqBkIndicator_span, lstBnkDetls(9),"don_AppChqBkIndicator_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If 
	If Not IsNull(lstBnkDetls(10)) Then
		If not VerifyInnerText(DOBCUL.don_AppSmsOption_span, lstBnkDetls(10),"don_AppSmsOption_span") Then
			bverifyDOBCUL_BnkUseDtlsSec=false
		End If
	End If
	'Next
	verifyDOBCUL_BnkUseDtlsSec=bverifyDOBCUL_BnkUseDtlsSec
End Function

'[Verify DOBCUL Income Details Section]
Public Function verifyDOBCUL_IncomeDtlsSec(lstIncmeDetls)
	bverifyDOBCUL_IncomeDtlsSec=true
	intSize=Ubound(lstIncmeDetls)
	'For i = 0 To intSize - 1
	If Not IsNull(lstIncmeDetls(0)) Then
		If not VerifyInnerText(DOBCUL.don_AppIncomeProof_span, lstIncmeDetls(0),"don_AppIncomeProof_span") Then
			bverifyDOBCUL_IncomeDtlsSec=false
		End If 
	End If
	'Next
	verifyDOBCUL_IncomeDtlsSec=bverifyDOBCUL_IncomeDtlsSec
End Function

'[Verify DOBCUL Supp Card Details Section]
Public Function verifyDOBCUL_SuppCardDtlsSec(lstSuppDetls)
	bverifyDOBCUL_SuppCardDtlsSec=true
	intSize=Ubound(lstSuppDetls)
	'For i = 0 To intSize - 1
	
	If Not IsNull(lstSuppDetls(0)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppSal_span, lstSuppDetls(0),"don_AppSuppSal_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If 
	End If
	If Not IsNull(lstSuppDetls(1)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppNameInNRICPp_span, lstSuppDetls(1),"don_AppSuppNameInNRICPp_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(2)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppEmbossName_span, lstSuppDetls(2),"don_AppSuppEmbossName_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(3)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppNationality_span, lstSuppDetls(3),"don_AppSuppNationality_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(4)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppNRICPpNum_span, lstSuppDetls(4),"don_AppSuppNRICPpNum_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(5)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppDOB_span, lstSuppDetls(5),"don_AppSuppDOB_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(6)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppGender_span, lstSuppDetls(6),"don_AppSuppGender_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(7)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppMobNum_span, lstSuppDetls(7),"don_AppSuppMobNum_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	If Not IsNull(lstSuppDetls(8)) Then
		If not VerifyInnerText(DOBCUL.don_AppSuppHomeNum_span, lstSuppDetls(8),"don_AppSuppHomeNum_span") Then
			bverifyDOBCUL_SuppCardDtlsSec=false
		End If
	End If
	'Next
	verifyDOBCUL_SuppCardDtlsSec=bverifyDOBCUL_SuppCardDtlsSec
End Function

'[Verify DOBCUL Guardian Details Section]
Public Function verifyDOBCUL_GuardianDtlsSec(lstGurdnDetls)
	bverifyDOBCUL_GuardianDtlsSec=true
	intSize=Ubound(lstGurdnDetls)
	'For i = 0 To intSize - 1
	If Not IsNull(lstGurdnDetls(0)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGSal_span, lstGurdnDetls(0),"don_AppPGSal_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If 
	End If
		
	If Not IsNull(lstGurdnDetls(1)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGNameInNRICPp_span, lstGurdnDetls(1),"don_AppPGNameInNRICPp_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(2)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGNRICPpNum_span, lstGurdnDetls(2),"don_AppPGNRICPpNum_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(3)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGRelnToApplicant_span, lstGurdnDetls(3),"don_AppPGRelnToApplicant_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(4)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGMobNum_span, lstGurdnDetls(4),"don_AppPGMobNum_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(5)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGHomeNum_span, lstGurdnDetls(5),"don_AppPGHomeNum_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(6)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGOfficeNum_span, lstGurdnDetls(6),"don_AppPGOfficeNum_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(7)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGCompName_span, lstGurdnDetls(7),"don_AppPGCompName_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End If
		
	If Not IsNull(lstGurdnDetls(8)) Then
		If not VerifyInnerText(DOBCUL.don_AppPGAnnualIncome_span, lstGurdnDetls(8),"don_AppPGAnnualIncome_span") Then
			bverifyDOBCUL_GuardianDtlsSec=false
		End If
	End if
	'Next
	verifyDOBCUL_GuardianDtlsSec=bverifyDOBCUL_GuardianDtlsSec
End Function
