'[Click Link Suspension Details from leftMenu]
Public Function ClickLink_SuspensionDetails()
bClickLink_SuspensionDetails=true
	'bcAccountOverview_LeftMenu.btnSuspensionDetails.Click
	SuspensionDetails.linkSuspensionDetails.Click
	WaitForIcallLoading
	If Err.Number<>0 Then
       bClickLink_SuspensionDetails=false
       LogMessage "WARN","Verification","Failed to Click Link  : Suspension Details" ,false
       Exit Function
	End If
	Wait 1
	waitForIcallLoading	
ClickLink_SuspensionDetails = bClickLink_SuspensionDetails
End Function 

'[Verify 60DPD in Suspension Details screen displayed as]
 Public Function Verify60DaysPastDueDetails(strCardNumber,StrSuspensionFlag,strSuspendedOn,strUpliftedOn,strOverride,strOvStartDate,strOvEndDate,strOutsideDBS,strOutsideDBSUpdatedOn)
 	Print "Inside function Verify60DaysPastDueDetails"
 	bVerify60DaysPastDueDetails = True
 	If Not IsNull(StrSuspensionFlag) Then
		If StrSuspensionFlag = "RUNTIME" Then
			Call getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPSuspensionFlag=Environment.Value("StrExpPDSuspensionFlag")
			strIserveSuspensionFlag=SuspensionDetails.PastDueSuspensionFlag.GetROProperty("innertext")
			If  Ucase(Trim(strVPSuspensionFlag)) = UCase(Trim(strIserveSuspensionFlag)) Then
				LogMessage "RSLT", "Verification","For Suspension Flag successfully matched with the expected value. Expected: "+ strVPSuspensionFlag &" , Actual: "& strIserveSuspensionFlag, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Suspension Flag not matching with the expected value. Expected: "+ strVPSuspensionFlag &" , Actual: "& strIserveSuspensionFlag, False
				bVerify60DaysPastDueDetails = False
			End If
		Else
			strIserveSuspensionFlag=SuspensionDetails.PastDueSuspensionFlag.GetROProperty("innertext")
			If  Ucase(Trim(StrSuspensionFlag)) = UCase(Trim(strIserveSuspensionFlag)) Then
				LogMessage "RSLT", "Verification","For Suspension Flag successfully matched with the expected value. Expected: "+ StrSuspensionFlag &" , Actual: "& strIserveSuspensionFlag, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Suspension Flag not matching with the expected value. Expected: "+ StrSuspensionFlag &" , Actual: "& strIserveSuspensionFlag, False
				bVerify60DaysPastDueDetails = False
			End If
		End If 
	End If 
   
    If Not IsNull(strSuspendedOn) Then
		If strSuspendedOn = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPSuspendedOn=Environment.Value("StrExpPDSuspensionDate")
			strIserveSuspendedOn=SuspensionDetails.PastDueSuspensedOn.GetROProperty("innertext")
			If  Trim(strVPSuspendedOn) = Trim(strIserveSuspendedOn) Then
				LogMessage "RSLT", "Verification","For SuspendedOn Date successfully matched with the expected value. Expected: "+ strVPSuspendedOn &" , Actual: "& strIserveSuspendedOn, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For SuspendedOn Date doesnt match with the expected value. Expected: "+ strVPSuspendedOn &" , Actual: "& strIserveSuspendedOn, False
				bVerify60DaysPastDueDetails = False
			End If
		Else 
			strIserveSuspendedOn=SuspensionDetails.PastDueSuspensedOn.GetROProperty("innertext")
			If  Trim(strSuspendedOn) = Trim(strIserveSuspendedOn) Then
				LogMessage "RSLT", "Verification","For SuspendedOn Date successfully matched with the expected value. Expected: "+ strSuspendedOn &" , Actual: "& strIserveSuspendedOn, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For SuspendedOn Date doesnt match with the expected value. Expected: "+ strSuspendedOn &" , Actual: "& strIserveSuspendedOn, False
				bVerify60DaysPastDueDetails = False
			End If
		End If
    End If	
    
      If Not IsNull(strUpliftedOn) Then
		If strUpliftedOn = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPUpliftedOn=Environment.Value("StrExpPDUpliftDate")
			strIserveUpliftedOn=SuspensionDetails.PastDueUpliftedOn.GetROProperty("innertext")
			If  Trim(strVPUpliftedOn) = Trim(strIserveUpliftedOn) Then
				LogMessage "RSLT", "Verification","For UpliftedOn Date successfully matched with the expected value. Expected: "+ strVPUpliftedOn &" , Actual: "& strIserveUpliftedOn, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For UpliftedOn Date doesnt match with the expected value. Expected: "+ strVPUpliftedOn &" , Actual: "& strIserveUpliftedOn, False
				bVerify60DaysPastDueDetails = False
			End If
		Else 
			strIserveUpliftedOn=SuspensionDetails.PastDueUpliftedOn.GetROProperty("innertext")
			If  Trim(strUpliftedOn) = Trim(strIserveUpliftedOn) Then
				LogMessage "RSLT", "Verification","For UpliftedOn Date successfully matched with the expected value. Expected: "+ strUpliftedOn &" , Actual: "& strIserveUpliftedOn, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For UpliftedOn Date doesnt match with the expected value. Expected: "+ strUpliftedOn &" , Actual: "& strIserveUpliftedOn, False
				bVerify60DaysPastDueDetails = False
			End If
		End If
    End If	
    	
    	If Not IsNull(strOverride) Then
		If strOverride = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPOverrideFlag=Environment.Value("StrExpPDOverrideFlag")
			strIserveOverrideFlag=SuspensionDetails.PastDueOverride.GetROProperty("innertext")
			If  Ucase(Trim(strVPOverrideFlag)) = UCase(Trim(strIserveOverrideFlag)) Then
				LogMessage "RSLT", "Verification","For Override Flag successfully matched with the expected value. Expected: "+ strVPOverrideFlag &" , Actual: "& strIserveOverrideFlag, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Override Flag not matching with the expected value. Expected: "+ strVPOverrideFlag &" , Actual: "& strIserveOverrideFlag, False
				bVerify60DaysPastDueDetails = False
			End If
		Else 
			strIserveOverrideFlag=SuspensionDetails.PastDueOverride.GetROProperty("innertext")
			If  Ucase(Trim(strOverride)) = UCase(Trim(strIserveOverrideFlag)) Then
				LogMessage "RSLT", "Verification","For Override Flag successfully matched with the expected value. Expected: "+ strOverride &" , Actual: "& strIserveOverrideFlag, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Override Flag not matching with the expected value. Expected: "+ strOverride &" , Actual: "& strIserveOverrideFlag, False
				bVerify60DaysPastDueDetails = False
			End If
		End If
    End If
    
    If Not IsNull(strOvStartDate) Then
		If strOvStartDate = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPOverrideStartDate=Environment.Value("StrExpPDStartDate")
			strIserveOverrideStartDate=SuspensionDetails.PastDueStartDate.GetROProperty("innertext")
			If  Trim(strVPOverrideStartDate) = Trim(strIserveOverrideStartDate) Then
				LogMessage "RSLT", "Verification","For Override StartDate successfully matched with the expected value. Expected: "+ strVPOverrideStartDate &" , Actual: "& strIserveOverrideStartDate, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Override StartDate doesnt match with the expected value. Expected: "+ strVPOverrideStartDate &" , Actual: "& strIserveOverrideStartDate, False
				bVerify60DaysPastDueDetails = False
			End If
		Else 
			strIserveOverrideStartDate=SuspensionDetails.PastDueStartDate.GetROProperty("innertext")
			If  Trim(strOvStartDate) = Trim(strIserveOverrideStartDate) Then
				LogMessage "RSLT", "Verification","For Override StartDate successfully matched with the expected value. Expected: "+ strOvStartDate &" , Actual: "& strIserveOverrideStartDate, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Override StartDate doesnt match with the expected value. Expected: "+ strOvStartDate &" , Actual: "& strIserveOverrideStartDate, False
				bVerify60DaysPastDueDetails = False
			End If
		End If
    End If	 
      
    If Not IsNull(strOvEndDate) Then
		If strOvEndDate = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPOverrideEndDate=Environment.Value("StrExpPDEndDate")
			strIserveOverrideEndDate=SuspensionDetails.PastDueEndDate.GetROProperty("innertext")
			If  Trim(strVPOverrideEndDate) = Trim(strIserveOverrideEndDate) Then
				LogMessage "RSLT", "Verification","For Override EndDate successfully matched with the expected value. Expected: "+ strVPOverrideEndDate &" , Actual: "& strIserveOverrideEndDate, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Override EndDate  doesnt match with the expected value. Expected: "+ strVPOverrideEndDate &" , Actual: "& strIserveOverrideEndDate, False
				bVerify60DaysPastDueDetails = False
			End If
		 Else 
		 	strIserveOverrideEndDate=SuspensionDetails.PastDueEndDate.GetROProperty("innertext")
			If  Trim(strOvEndDate) = Trim(strIserveOverrideEndDate) Then
				LogMessage "RSLT", "Verification","For Override EndDate successfully matched with the expected value. Expected: "+ strOvEndDate &" , Actual: "& strIserveOverrideEndDate, True
				bVerify60DaysPastDueDetails = True
			else
				LogMessage "WARN", "Verification","For Override EndDate  doesnt match with the expected value. Expected: "+ strOvEndDate &" , Actual: "& strIserveOverrideEndDate, False
				bVerify60DaysPastDueDetails = False
			End If
		End If
    End If	     

    If Not IsNull(strOutsideDBS) Then
	 	strIserveOutsideDBS=SuspensionDetails.lblOutsideDBS.GetROProperty("innertext")
		If  Trim(strOutsideDBS) = Trim(strIserveOutsideDBS) Then
			LogMessage "RSLT", "Verification","For Outside DBS successfully matched with the expected value. Expected: "+ strOutsideDBS &" , Actual: "& strIserveOutsideDBS, True
			bVerify60DaysPastDueDetails = True
		else
			LogMessage "WARN", "Verification","For Outside DBS doesnt match with the expected value. Expected: "+ strOutsideDBS &" , Actual: "& strIserveOutsideDBS, False
			bVerify60DaysPastDueDetails = False
		End If
	End If
	
	If Not IsNull(strOutsideDBSUpdatedOn) Then
	 	strIserveOutDBS_UpdatedOn=SuspensionDetails.lblOutsideDBS_UpdatedOn.GetROProperty("innertext")
		If  Trim(strOutsideDBSUpdatedOn) = Trim(strIserveOutDBS_UpdatedOn) Then
			LogMessage "RSLT", "Verification","For Outside DBS updated On successfully matched with the expected value. Expected: "+ strOutsideDBSUpdatedOn &" , Actual: "& strIserveOutDBS_UpdatedOn, True
			bVerify60DaysPastDueDetails = True
		else
			LogMessage "WARN", "Verification","For Outside DBS updated On doesnt match with the expected value. Expected: "+ strOutsideDBSUpdatedOn &" , Actual: "& strIserveOutDBS_UpdatedOn, False
			bVerify60DaysPastDueDetails = False
		End If
	End If
  Verify60DaysPastDueDetails = bVerify60DaysPastDueDetails
  Print "Function Verify60DaysPastDueDetails executed successfully"
 End Function

'[Verify BalancetoIncome in Suspension Details screen displayed as]
Public Function VerifyBIDetails(strCardNumber,StrAI,StrSuspensionFlag,strSuspendedOn,strSuspendedDays,strUpliftedOn,strOverride,strOvStartDate,strOvEndDate)
Print "Inside function VerifyBIDetails"
 	bVerifyBalanceIncomeDetails = True
 	If Not IsNull(StrAI) Then
		If StrAI = "RUNTIME" Then
			getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBIAccreditorIndicator=Environment.Value("StrExpBIAccreditorIndicator")
			strIserveBIAccreditorIndicator=SuspensionDetails.AccreditedInvestor.GetROProperty("innertext")
			If  Ucase(Trim(strVPBIAccreditorIndicator)) = UCase(Trim(strIserveBIAccreditorIndicator)) Then
				LogMessage "RSLT", "Verification","For AccreditorIndicator successfully matched with the expected value. Expected: "+ strVPBIAccreditorIndicator &" , Actual: "& strIserveBIAccreditorIndicator, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For AccreditorIndicator not matching with the expected value. Expected: "+ strVPBIAccreditorIndicator &" , Actual: "& strIserveBIAccreditorIndicator, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBIAccreditorIndicator=SuspensionDetails.AccreditedInvestor.GetROProperty("innertext")
			If  Ucase(Trim(StrAI)) = UCase(Trim(strIserveBIAccreditorIndicator)) Then
				LogMessage "RSLT", "Verification","For AccreditorIndicator successfully matched with the expected value. Expected: "+ StrAI &" , Actual: "& strIserveBIAccreditorIndicator, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For AccreditorIndicator not matching with the expected value. Expected: "+ StrAI &" , Actual: "& strIserveBIAccreditorIndicator, False
				bVerifyBalanceIncomeDetails = False
			End If
		End If
    End If
 	If Not IsNull(StrSuspensionFlag) Then
		If StrSuspensionFlag = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBISuspensionFlag=Environment.Value("StrExpBISuspensionFlag")
			strIserveBISuspensionFlag=SuspensionDetails.BalancetoIncomeSuspensionFlag.GetROProperty("innertext")
			If  Ucase(Trim(strVPBISuspensionFlag)) = UCase(Trim(strIserveBISuspensionFlag)) Then
				LogMessage "RSLT", "Verification","For Balance to Income Suspension Flag successfully matched with the expected value. Expected: "+ strIserveBISuspensionFlag &" , Actual: "& strVPBISuspensionFlag, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income Suspension Flag not matching with the expected value. Expected: "+ strIserveBISuspensionFlag &" , Actual: "& strVPBISuspensionFlag, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBISuspensionFlag=SuspensionDetails.BalancetoIncomeSuspensionFlag.GetROProperty("innertext")
			If  Ucase(Trim(StrSuspensionFlag)) = UCase(Trim(strIserveBISuspensionFlag)) Then
				LogMessage "RSLT", "Verification","For Balance to Income Suspension Flag successfully matched with the expected value. Expected: "+ StrSuspensionFlag &" , Actual: "& strVPBISuspensionFlag, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income Suspension Flag not matching with the expected value. Expected: "+ StrSuspensionFlag &" , Actual: "& strVPBISuspensionFlag, False
				bVerifyBalanceIncomeDetails = False
			End If		
		End If
    End If
    
    If Not IsNull(strSuspendedOn) Then
		If strSuspendedOn = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBISuspendedOn=Environment.Value("StrExpBISuspendedOn")
			strIserveBISuspendedOn=SuspensionDetails.BalancetoIncomeSuspensedOn.GetROProperty("innertext")
			If  Trim(strVPBISuspendedOn) = Trim(strIserveBISuspendedOn) Then
				LogMessage "RSLT", "Verification","For Balance to Income SuspendedOn Date successfully matched with the expected value. Expected: "+ strVPBISuspendedOn &" , Actual: "& strIserveBISuspendedOn, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income SuspendedOn Date doesnt match with the expected value. Expected: "+ strVPBISuspendedOn &" , Actual: "& strIserveBISuspendedOn, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBISuspendedOn=SuspensionDetails.BalancetoIncomeSuspensedOn.GetROProperty("innertext")
			If  Trim(strSuspendedOn) = Trim(strIserveBISuspendedOn) Then
				LogMessage "RSLT", "Verification","For Balance to Income SuspendedOn Date successfully matched with the expected value. Expected: "+ strSuspendedOn &" , Actual: "& strIserveBISuspendedOn, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income SuspendedOn Date doesnt match with the expected value. Expected: "+ strSuspendedOn &" , Actual: "& strIserveBISuspendedOn, False
				bVerifyBalanceIncomeDetails = False
			End If
		End If
    End If	
    
    If Not IsNull(strSuspendedDays) Then
		If strSuspendedDays = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBISuspendedDays=Environment.Value("StrExpBISuspensionDays")
			strIserveBISuspendedDays=SuspensionDetails.BalancetoIncomeSuspensedDays.GetROProperty("innertext")
			If  Trim(strVPBISuspendedDays) = Trim(strIserveBISuspendedDays) Then
				LogMessage "RSLT", "Verification","For Balance to Income Suspended Days successfully matched with the expected value. Expected: "+ strVPBISuspendedDays &" , Actual: "& strIserveBISuspendedDays, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income Suspended Days doesnt match with the expected value. Expected: "+ strVPBISuspendedDays &" , Actual: "& strIserveBISuspendedDays, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBISuspendedDays=SuspensionDetails.BalancetoIncomeSuspensedDays.GetROProperty("innertext")
			If  Trim(strSuspendedDays) = Trim(strIserveBISuspendedDays) Then
				LogMessage "RSLT", "Verification","For Balance to Income Suspended Days successfully matched with the expected value. Expected: "+ strSuspendedDays &" , Actual: "& strIserveBISuspendedDays, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income Suspended Days doesnt match with the expected value. Expected: "+ strSuspendedDays &" , Actual: "& strIserveBISuspendedDays, False
				bVerifyBalanceIncomeDetails = False
			End If
		End If
    End If	

    If Not IsNull(strUpliftedOn) Then
		If strUpliftedOn = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBIUpliftedOn=Environment.Value("StrExpBIUpliftedOn")
			strIserveBIUpliftedOn=SuspensionDetails.BalancetoIncomeUpliftedOn.GetROProperty("innertext")
			If  Trim(strVPBIUpliftedOn) = Trim(strIserveBIUpliftedOn) Then
				LogMessage "RSLT", "Verification","For Balance to Income UpliftedOn Date successfully matched with the expected value. Expected: "+ strVPBIUpliftedOn &" , Actual: "& strIserveBIUpliftedOn, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income UpliftedOn Date doesnt match with the expected value. Expected: "+ strVPBIUpliftedOn &" , Actual: "& strIserveBIUpliftedOn, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBIUpliftedOn=SuspensionDetails.BalancetoIncomeUpliftedOn.GetROProperty("innertext")
			If  Trim(strUpliftedOn) = Trim(strIserveBIUpliftedOn) Then
				LogMessage "RSLT", "Verification","For Balance to Income UpliftedOn Date successfully matched with the expected value. Expected: "+ strUpliftedOn &" , Actual: "& strIserveBIUpliftedOn, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income UpliftedOn Date doesnt match with the expected value. Expected: "+ strUpliftedOn &" , Actual: "& strIserveBIUpliftedOn, False
				bVerifyBalanceIncomeDetails = False
			End If
		End If
    End If	
    	
    If Not IsNull(strOverride) Then
		If strOverride = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBIOverrideFlag=Environment.Value("StrExpBIOverrideFlag")
			strIserveBIOverrideFlag=SuspensionDetails.BalancetoIncomeOverride.GetROProperty("innertext")
			If  Ucase(Trim(strVPBIOverrideFlag)) = UCase(Trim(strIserveBIOverrideFlag)) Then
				LogMessage "RSLT", "Verification","For Balance to Income Override Flag successfully matched with the expected value. Expected: "+ strVPBIOverrideFlag &" , Actual: "& strIserveBIOverrideFlag, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income Override Flag not matching with the expected value. Expected: "+ strVPBIOverrideFlag &" , Actual: "& strIserveBIOverrideFlag, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBIOverrideFlag=SuspensionDetails.BalancetoIncomeOverride.GetROProperty("innertext")
			If  Ucase(Trim(strOverride)) = UCase(Trim(strIserveBIOverrideFlag)) Then
				LogMessage "RSLT", "Verification","For Balance to Income Override Flag successfully matched with the expected value. Expected: "+ strOverride &" , Actual: "& strIserveBIOverrideFlag, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Balance to Income Override Flag not matching with the expected value. Expected: "+ strOverride &" , Actual: "& strIserveBIOverrideFlag, False
				bVerifyBalanceIncomeDetails = False
			End If
		End If
    End If
    
    If Not IsNull(strOvStartDate) Then
		If strOvStartDate = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBIOverrideStartDate=Environment.Value("StrExpBIStartDate")
			strIserveBIOverrideStartDate=SuspensionDetails.BalancetoIncomeStartDate.GetROProperty("innertext")
			If  Trim(strVPBIOverrideStartDate) = Trim(strIserveBIOverrideStartDate) Then
				LogMessage "RSLT", "Verification","For Override StartDate successfully matched with the expected value. Expected: "+ strVPBIOverrideStartDate &" , Actual: "& strIserveBIOverrideStartDate, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Override StartDate doesnt match with the expected value. Expected: "+ strVPBIOverrideStartDate &" , Actual: "& strIserveBIOverrideStartDate, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBIOverrideStartDate=SuspensionDetails.BalancetoIncomeStartDate.GetROProperty("innertext")
			If  Trim(strOvStartDate) = Trim(strIserveBIOverrideStartDate) Then
				LogMessage "RSLT", "Verification","For Override StartDate successfully matched with the expected value. Expected: "+ strOvStartDate &" , Actual: "& strIserveBIOverrideStartDate, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","For Override StartDate doesnt match with the expected value. Expected: "+ strOvStartDate &" , Actual: "& strIserveBIOverrideStartDate, False
				bVerifyBalanceIncomeDetails = False
			End If
		
		End If
    End If	 
      
    If Not IsNull(strOvEndDate) Then
		If strOvStartDate = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPBIOverrideEndDate=Environment.Value("StrExpBIEndDate")
			strIserveBIOverrideEndDate=SuspensionDetails.BalancetoIncomeEndDate.GetROProperty("innertext")
			If  Trim(strVPBIOverrideEndDate) = Trim(strIserveBIOverrideEndDate) Then
				LogMessage "RSLT", "Verification","Balance to Income Override EndDate successfully matched with the expected value. Expected: "+ strVPBIOverrideEndDate &" , Actual: "& strIserveBIOverrideEndDate, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","Balance to Income Override EndDate  doesnt match with the expected value. Expected: "+ strVPBIOverrideEndDate &" , Actual: "& strIserveBIOverrideEndDate, False
				bVerifyBalanceIncomeDetails = False
			End If
		Else 
			strIserveBIOverrideEndDate=SuspensionDetails.BalancetoIncomeEndDate.GetROProperty("innertext")
			If  Trim(strOvEndDate) = Trim(strIserveBIOverrideEndDate) Then
				LogMessage "RSLT", "Verification","Balance to Income Override EndDate successfully matched with the expected value. Expected: "+ strOvEndDate &" , Actual: "& strIserveBIOverrideEndDate, True
				bVerifyBalanceIncomeDetails = True
			else
				LogMessage "WARN", "Verification","Balance to Income Override EndDate  doesnt match with the expected value. Expected: "+ strOvEndDate &" , Actual: "& strIserveBIOverrideEndDate, False
				bVerifyBalanceIncomeDetails = False
			End If
		End If
    End If	      
  VerifyBIDetails = bVerifyBalanceIncomeDetails
  Print "Function VerifyBIDetails executed successfully"
 End Function

'[Verify Reinstatement in Suspension Details screen displayed as]
Public Function VerifyReinstatementDetails(strCardNumber,StrQualifiedFlag,StrQualifiedAttempts,strLastQualifiedDate,strNonQualifiedAttempts,strLastNonQualifiedDate)
Print "Inside function VerifyReinstatementDetails"
 	bReinstatementDetails = True
 	If Not IsNull(StrQualifiedFlag) Then
		If StrQualifiedFlag = "RUNTIME" Then
			getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPQualifiedFlag=Environment.Value("StrExpQualifiedFlag")
			strIserveQualifiedFlag=SuspensionDetails.ReinstatementQualified.GetROProperty("innertext")
			If  Ucase(Trim(strVPQualifiedFlag)) = UCase(Trim(strIserveQualifiedFlag)) Then
				LogMessage "RSLT", "Verification","Reinstatement Qualified Flag successfully matched with the expected value. Expected: "+ strVPQualifiedFlag &" , Actual: "& strIserveQualifiedFlag, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Qualified Flag not matching with the expected value. Expected: "+ strVPQualifiedFlag &" , Actual: "& strIserveQualifiedFlag, False
				bReinstatementDetails = False
			End If
		Else 
			strIserveQualifiedFlag=SuspensionDetails.ReinstatementQualified.GetROProperty("innertext")
			If  Ucase(Trim(StrQualifiedFlag)) = UCase(Trim(strIserveQualifiedFlag)) Then
				LogMessage "RSLT", "Verification","Reinstatement Qualified Flag successfully matched with the expected value. Expected: "+ StrQualifiedFlag &" , Actual: "& strIserveQualifiedFlag, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Qualified Flag not matching with the expected value. Expected: "+ StrQualifiedFlag &" , Actual: "& strIserveQualifiedFlag, False
				bReinstatementDetails = False
			End If
		End If
    End If
 	If Not IsNull(StrQualifiedAttempts) Then
		If StrQualifiedAttempts = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPQualifiedAttemps=Environment.Value("StrExpQualifiedAttempts")
			strIserveQualifiedAttemps=SuspensionDetails.ReinstatementQualifiedAttempts.GetROProperty("innertext")
			'strIserveQualifiedAttemps =Int(strIserveQualifiedAttemps)
			If  Trim(strVPQualifiedAttemps) = Trim(strIserveQualifiedAttemps) Then
				LogMessage "RSLT", "Verification","Reinstatement Qualified Attempts successfully matched with the expected value. Expected: "+ strVPQualifiedAttemps &" , Actual: "& strIserveQualifiedAttemps, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Qualified Attempts not matching with the expected value. Expected: "+ strVPQualifiedAttemps &" , Actual: "& strIserveQualifiedAttemps, False
				bReinstatementDetails = False
			End If
		Else 
			strIserveQualifiedAttemps=SuspensionDetails.ReinstatementQualifiedAttempts.GetROProperty("innertext")
			If  Trim(StrQualifiedAttempts) = Trim(strIserveQualifiedAttemps) Then
				LogMessage "RSLT", "Verification","Reinstatement Qualified Attempts successfully matched with the expected value. Expected: "+ StrQualifiedAttempts &" , Actual: "& strIserveQualifiedAttemps, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Qualified Attempts not matching with the expected value. Expected: "+ StrQualifiedAttempts &" , Actual: "& strIserveQualifiedAttemps, False
				bReinstatementDetails = False
			End If
		End If
    End If
    
    If Not IsNull(strLastQualifiedDate) Then
		If strLastQualifiedDate = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPLastQualifiedDate=Environment.Value("StrExpLastQualifiedDate")
			strIserveLastQualifiedDate=SuspensionDetails.ReinstatementLastQualified.GetROProperty("innertext")
			If  Trim(strVPLastQualifiedDate) = Trim(strIserveLastQualifiedDate) Then
				LogMessage "RSLT", "Verification","Reinstatement Last Qualified Date successfully matched with the expected value. Expected: "+ strVPLastQualifiedDate &" , Actual: "& strIserveLastQualifiedDate, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Last Qualified Date doesnt match with the expected value. Expected: "+ strVPLastQualifiedDate &" , Actual: "& strIserveLastQualifiedDate, False
				bReinstatementDetails = False
			End If
		Else
		 	strIserveLastQualifiedDate=SuspensionDetails.ReinstatementLastQualified.GetROProperty("innertext")
			If  Trim(strLastQualifiedDate) = Trim(strIserveLastQualifiedDate) Then
				LogMessage "RSLT", "Verification","Reinstatement Last Qualified Date successfully matched with the expected value. Expected: "+ strLastQualifiedDate &" , Actual: "& strIserveLastQualifiedDate, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Last Qualified Date doesnt match with the expected value. Expected: "+ strLastQualifiedDate &" , Actual: "& strIserveLastQualifiedDate, False
				bReinstatementDetails = False
			End If
		End If
    End If	
    
    If Not IsNull(strNonQualifiedAttempts) Then
		If strNonQualifiedAttempts = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPNonQualAttempts=Environment.Value("StrExpNonQualifiedAttempts")
			strIserveNonQualAttempts=SuspensionDetails.ReinstatementNonQualifiedAttempts.GetROProperty("innertext")
			'strIserveNonQualAttempts=Int(strIserveNonQualAttempts)
			If  Trim(strVPNonQualAttempts) = Trim(strIserveNonQualAttempts) Then
				LogMessage "RSLT", "Verification","Reinstatement Last Non Qualified Attempts successfully matched with the expected value. Expected: "+ strVPNonQualAttempts &" , Actual: "& strIserveNonQualAttempts, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Last Non Qualified Attempts doesnt match with the expected value. Expected: "+ strVPNonQualAttempts &" , Actual: "& strIserveNonQualAttempts, False
				bReinstatementDetails = False
			End If
		Else 
			strIserveNonQualAttempts=SuspensionDetails.ReinstatementNonQualifiedAttempts.GetROProperty("innertext")
			If  Trim(strNonQualifiedAttempts) = Trim(strIserveNonQualAttempts) Then
				LogMessage "RSLT", "Verification","Reinstatement Last Non Qualified Attempts successfully matched with the expected value. Expected: "+ strNonQualifiedAttempts &" , Actual: "& strIserveNonQualAttempts, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Last Non Qualified Attempts doesnt match with the expected value. Expected: "+ strNonQualifiedAttempts &" , Actual: "& strIserveNonQualAttempts, False
				bReinstatementDetails = False
			End If		
		
		End If
    End If	

    If Not IsNull(strLastNonQualifiedDate) Then
		If strLastNonQualifiedDate = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPLastNonQualDate=Environment.Value("StrExpLastNonQualifiedDate")
			strIserveLastNonQualDate=SuspensionDetails.ReinstatementLastNonQualified.GetROProperty("innertext")
			If  Trim(strVPLastNonQualDate) = Trim(strIserveLastNonQualDate) Then
				LogMessage "RSLT", "Verification","Reinstatement Last Non Qualified Date successfully matched with the expected value. Expected: "+ strVPLastNonQualDate &" , Actual: "& strIserveLastNonQualAttempts, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Last Non Qualified Date doesnt match with the expected value. Expected: "+ strVPLastNonQualDate &" , Actual: "& strIserveLastNonQualAttempts, False
				bReinstatementDetails = False
			End If
		Else 
			strIserveLastNonQualDate=SuspensionDetails.ReinstatementLastNonQualified.GetROProperty("innertext")
			If  Trim(strLastNonQualifiedDate) = Trim(strIserveLastNonQualDate) Then
				LogMessage "RSLT", "Verification","Reinstatement Last Non Qualified Date successfully matched with the expected value. Expected: "+ strLastNonQualifiedDate &" , Actual: "& strIserveLastNonQualDate, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Reinstatement Last Non Qualified Date doesnt match with the expected value. Expected: "+ strLastNonQualifiedDate &" , Actual: "& strIserveLastNonQualDate, False
				bReinstatementDetails = False
			End If
		End If
		
    End If	   
 VerifyReinstatementDetails = bReinstatementDetails
 Print "Function VerifyReinstatementDetails executed successfully"
 End Function

'[Verify Aggregated Balance Details in Suspension screen displayed as]
Public Function VerifyAggBalance(strCardNumber,StrAggregatedBalance,StrAggBalUpdatedOn,StrStaffIndicator)
Print "Inside function VerifyAggBalance"
 	bAggregateBalance = True
 	If Not IsNull(StrAggregatedBalance) Then
		If StrAggregatedBalance = "RUNTIME" Then
			getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPAggrBalance=Environment.Value("StrExpAggregateBalance")
			strIserveAggrBalance=SuspensionDetails.AggregatedBalance.GetROProperty("innertext")
			strIserveAggrBalance = Replace(Replace(strIserveAggrBalance,",",""),".","")
			strIserveAggrBalance = fTrimZero(strIserveAggrBalance)
			'strIserveAggrBalance = cdbl(strIserveAggrBalance)
			If  Trim(strVPAggrBalance) = Trim(strIserveAggrBalance) Then
				LogMessage "RSLT", "Verification","Aggregated Balance successfully matched with the expected value. Expected: "+ strVPAggrBalance &" , Actual: "& strIserveAggrBalance, True
				bAggregateBalance = True
			else
				LogMessage "WARN", "Verification","Aggregated Balance not matching with the expected value. Expected: "+ strVPAggrBalance &" , Actual: "& strIserveAggrBalance, False
				bAggregateBalance = False
			End If
		Else 
			strIserveAggrBalance=SuspensionDetails.AggregatedBalance.GetROProperty("innertext")
			If  Trim(StrAggregatedBalance) = Trim(strIserveAggrBalance) Then
				LogMessage "RSLT", "Verification","Aggregated Balance successfully matched with the expected value. Expected: "+ StrAggregatedBalance &" , Actual: "& strIserveAggrBalance, True
				bAggregateBalance = True
			else
				LogMessage "WARN", "Verification","Aggregated Balance not matching with the expected value. Expected: "+ StrAggregatedBalance &" , Actual: "& strIserveAggrBalance, False
				bAggregateBalance = False
			End If

		End If
    End If
    
 	If Not IsNull(StrAggBalUpdatedOn) Then
		If StrAggBalUpdatedOn = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPUpadatedOn=Environment.Value("StrExpAGUpdatedOn")
			strIserveUpadatedOn=SuspensionDetails.AggregatedBalanceUpdatedOn.GetROProperty("innertext")
			If  Trim(StrAggBalUpdatedOn) = Trim(strIserveUpadatedOn) Then
				LogMessage "RSLT", "Verification","Aggregate Balance UpdatedOn Date successfully matched with the expected value. Expected: "+ strVPUpadatedOn &" , Actual: "& strIserveUpadatedOn, True
				bAggregateBalance = True
			else
				LogMessage "WARN", "Verification","Aggregate Balance UpdatedOn Date doesnt match with the expected value. Expected: "+ strVPUpadatedOn &" , Actual: "& strIserveUpadatedOn, False
				bAggregateBalance = False
			End If
		Else 
			strIserveUpadatedOn=SuspensionDetails.AggregatedBalanceUpdatedOn.GetROProperty("innertext")
			If  Trim(StrAggBalUpdatedOn) = Trim(strIserveUpadatedOn) Then
				LogMessage "RSLT", "Verification","Aggregate Balance UpdatedOn Date successfully matched with the expected value. Expected: "+ StrAggBalUpdatedOn &" , Actual: "& strIserveUpadatedOn, True
				bAggregateBalance = True
			else
				LogMessage "WARN", "Verification","Aggregate Balance UpdatedOn Date doesnt match with the expected value. Expected: "+ StrAggBalUpdatedOn &" , Actual: "& strIserveUpadatedOn, False
				bAggregateBalance = False
			End If
		End If
    End If
    
    If Not IsNull(StrStaffIndicator) Then
		If StrStaffIndicator = "RUNTIME" Then
			'getSuspensionDetails_GFQC_Vplus(strCardNumber)
			strVPStaffIndicator=Environment.Value("StrExpStaffIndicator")
			strIserveStaffIndicator=SuspensionDetails.StaffIndicator.GetROProperty("innertext")
			If  Ucase(Trim(strVPStaffIndicator))= Ucase(Trim(strIserveStaffIndicator)) Then
				LogMessage "RSLT", "Verification","Staff Indicator flag successfully matched with the expected value. Expected: "+ strVPStaffIndicator &" , Actual: "& strIserveStaffIndicator, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Staff Indicator flag doesnt match with the expected value. Expected: "+ strVPStaffIndicator &" , Actual: "& strIserveStaffIndicator, False
				bReinstatementDetails = False
			End If
		Else
			strIserveStaffIndicator=SuspensionDetails.StaffIndicator.GetROProperty("innertext")
			If  Ucase(Trim(StrStaffIndicator))= Ucase(Trim(strIserveStaffIndicator)) Then
				LogMessage "RSLT", "Verification","Staff Indicator flag successfully matched with the expected value. Expected: "+ StrStaffIndicator &" , Actual: "& strIserveStaffIndicator, True
				bReinstatementDetails = True
			else
				LogMessage "WARN", "Verification","Staff Indicator flag doesnt match with the expected value. Expected: "+ StrStaffIndicator &" , Actual: "& strIserveStaffIndicator, False
				bReinstatementDetails = False
			End If
		End If
    End If	
    VerifyAggBalance= bAggregateBalance
    Print "Function VerifyAggBalance executed successfully"
End Function 

'[Verify Income details for NonStaff in Suspension Details screen displayed as]
Public Function VerifyIncomeForNonStaff(StrCardNumber,StrStaffIndicator,StrAppAnnualIncome,StrAppUpdatedOn,StrAUMIndicator,StrSCAnnualIncome,StrSCUpdatedOn,StrIncomeIndicator)
Print "Inside function VerifyIncomeForNonStaff"
bIncomeforNonStaff = True
If StrStaffIndicator = "N" Then
	If Not IsNull(StrAppAnnualIncome) Then
		If StrAppAnnualIncome = "RUNTIME" Then
		getSuspensionDetails_GFQC_Vplus(strCardNumber)
		strVPAppAI=Environment.Value("StrExpApplicationAI")
		StrIServeAppAI = SuspensionDetails.ApplicationAnnualIncome.GetROProperty("innertext")
		StrIServeAppAI = Replace(Replace(StrIServeAppAI,",",""),".","")
		StrIServeAppAI = fTrimZero(StrIServeAppAI)
		If strVPAppAI = StrIServeAppAI Then
			LogMessage "RSLT", "Verification","Application Annual Income details successfully matched with the expected value. Expected: "+ strVPStaffIndicator &" , Actual: "& strIserveStaffIndicator, True
			bIncomeforNonStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income details doesnt match with the expected value. Expected: "+ strVPStaffIndicator &" , Actual: "& strIserveStaffIndicator, False
			bIncomeforNonStaff = False
		End If
	Else 
		StrIServeAppAI = SuspensionDetails.ApplicationAnnualIncome.GetROProperty("innertext")
		If StrAppAnnualIncome = StrIServeAppAI Then
			LogMessage "RSLT", "Verification","Application Annual Income details successfully matched with the expected value. Expected: "+ StrAppAnnualIncome &" , Actual: "& StrIServeAppAI, True
			bIncomeforNonStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income details doesnt match with the expected value. Expected: "+ StrAppAnnualIncome &" , Actual: "& StrIServeAppAI, False
			bIncomeforNonStaff = False
		End If
	End If
	
	If Not IsNull(StrAppUpdatedOn) Then
		If StrAppUpdatedOn = "RUNTIME" Then
			strVPAppUpadatedOn=Environment.Value("StrExpApplicationAIUpdateOn")
			StrIServeAppUpdatedOn = SuspensionDetails.AnnualIncomeUpdatedOn.GetROProperty("innertext")
			If Trim(strVPAppUpadatedOn) = Trim(StrIServeAppUpdatedOn) Then
				LogMessage "RSLT", "Verification","Application Annual Income UpdatedON date successfully matched with the expected value. Expected: "+ strVPAppUpadatedOn &" , Actual: "& StrIServeAppUpdatedOn, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Application Annual Income UpdatedON date doesnt match with the expected value. Expected: "+ strVPAppUpadatedOn &" , Actual: "& StrIServeAppUpdatedOn, False
				bIncomeforNonStaff = False
			End If
		Else
		StrIServeAppUpdatedOn = SuspensionDetails.AnnualIncomeUpdatedOn.GetROProperty("innertext")
		If Trim(StrAppUpdatedOn) = Trim(StrIServeAppUpdatedOn) Then
			LogMessage "RSLT", "Verification","Application Annual Income UpdatedON date successfully matched with the expected value. Expected: "+ StrAppUpdatedOn &" , Actual: "& StrIServeAppUpdatedOn, True
			bIncomeforNonStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income UpdatedON date doesnt match with the expected value. Expected: "+ StrAppUpdatedOn &" , Actual: "& StrIServeAppUpdatedOn, False
			bIncomeforNonStaff = False
		End If
		End If
	End If
	
	If Not IsNull(StrAUMIndicator) Then
		If StrAUMIndicator = "RUNTIME" Then
		strVPAppAUMInd = Environment.Value("StrExpAUMIndicator")
		StrIserveAUMInd = SuspensionDetails.AUMIndicator.GetROProperty("innertext")
		If Trim(StrIserveAUMInd) = Trim(strVPAppAUMInd) Then
			LogMessage "RSLT", "Verification","Application Annual Income AUM Indicator successfully matched with the expected value. Expected: "+ strVPAppAUMInd &" , Actual: "& StrIserveAUMInd, True
			bIncomeforNonStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income AUM Indicator doesnt match with the expected value. Expected: "+ strVPAppAUMInd &" , Actual: "& StrIserveAUMInd, False
			bIncomeforNonStaff = False
		End If
	Else
		StrIserveAUMInd = SuspensionDetails.AUMIndicator.GetROProperty("innertext")
		If Trim(StrIserveAUMInd) = Trim(StrAUMIndicator) Then
			LogMessage "RSLT", "Verification","Application Annual Income AUM Indicator successfully matched with the expected value. Expected: "+ StrAUMIndicator &" , Actual: "& StrIserveAUMInd, True
			bIncomeforNonStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income AUM Indicator doesnt match with the expected value. Expected: "+ StrAUMIndicator &" , Actual: "& StrIserveAUMInd, False
			bIncomeforNonStaff = False
		End If
	End If
	End IF
	
	If Not IsNull(StrSCAnnualIncome) Then
		If StrSCAnnualIncome = "RUNTIME" Then
		strVPSCAI= Environment.Value("StrExpSCAnnualIncome")
		strIserveSCAI =SuspensionDetails.SalaryCreditingAnnualIncome.GetROProperty("innertext")
		strIserveSCAI = Replace(Replace(strIserveSCAI,",",""),".","")
		strIserveSCAI = fTrimZero(strIserveSCAI)
			If strIserveSCAI = strVPSCAI Then
				LogMessage "RSLT", "Verification","Salary Crediting Annual Income successfully matched with the expected value. Expected: "+ strVPSCAI &" , Actual: "& strIserveSCAI, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Salary Crediting Annual Income doesnt match with the expected value. Expected: "+ strVPSCAI &" , Actual: "& strIserveSCAI, False
				bIncomeforNonStaff = False
			End If
		Else 
			strIserveSCAI =SuspensionDetails.SalaryCreditingAnnualIncome.GetROProperty("innertext")
			If strIserveSCAI = StrSCAnnualIncome Then
				LogMessage "RSLT", "Verification","Salary Crediting Annual Income successfully matched with the expected value. Expected: "+ StrSCAnnualIncome &" , Actual: "& strIserveSCAI, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Salary Crediting Annual Income doesnt match with the expected value. Expected: "+ StrSCAnnualIncome &" , Actual: "& strIserveSCAI, False
				bIncomeforNonStaff = False
			End If
		End If
	End If
	
	If Not IsNull(StrSCUpdatedOn) Then
		If StrSCUpdatedOn = "RUNTIME" Then
		strVPSCUpdatedON = Environment.Value("StrExpSCAnnualIncomeUpdatedOn")
		strIserveSCUpdatedON = SuspensionDetails.SalaryCreditingUpdatedOn.GetROProperty("innertext")
			If Trim(strIserveSCUpdatedON) = Trim(strVPSCUpdatedON) Then
				LogMessage "RSLT", "Verification","Salary Crediting Annual Income UpdatedON successfully matched with the expected value. Expected: "+ strVPSCUpdatedON &" , Actual: "& strIserveSCUpdatedON, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Salary Crediting Annual Income UpdatedON doesnt match with the expected value. Expected: "+ strVPSCUpdatedON &" , Actual: "& strIserveSCUpdatedON, False
				bIncomeforNonStaff = False
			End If
		Else
		 strIserveSCUpdatedON = SuspensionDetails.SalaryCreditingUpdatedOn.GetROProperty("innertext")
			If Trim(strIserveSCUpdatedON) = Trim(StrSCUpdatedOn) Then
				LogMessage "RSLT", "Verification","Salary Crediting Annual Income UpdatedON successfully matched with the expected value. Expected: "+ StrSCUpdatedOn &" , Actual: "& strIserveSCUpdatedON, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Salary Crediting Annual Income UpdatedON doesnt match with the expected value. Expected: "+ StrSCUpdatedOn &" , Actual: "& strIserveSCUpdatedON, False
				bIncomeforNonStaff = False
			End IF
		End IF
	End If
	
	If Not IsNull(StrIncomeIndicator) Then
		If StrIncomeIndicator = "RUNTIME" Then
		strVPSCIncomeInd = Environment.Value("StrExpSCIncomeIndicator")
		strIserveSCIncomeInd = SuspensionDetails.IncomeIndicator.GetROProperty("innertext")
			If Trim(strIserveSCIncomeInd) = Trim(strVPSCIncomeInd) Then
				LogMessage "RSLT", "Verification","Salary Crediting Income Indicator successfully matched with the expected value. Expected: "+ strVPSCIncomeInd &" , Actual: "& strIserveSCIncomeInd, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Salary Crediting Income Indicator doesnt match with the expected value. Expected: "+ strVPSCIncomeInd &" , Actual: "& strIserveSCIncomeInd, False
				bIncomeforNonStaff = False
			End If
		Else
			 strIserveSCIncomeInd = SuspensionDetails.IncomeIndicator.GetROProperty("innertext")
			If Trim(StrIncomeIndicator) = Trim(strIserveSCIncomeInd) Then
				LogMessage "RSLT", "Verification","Salary Crediting Annual Income UpdatedON successfully matched with the expected value. Expected: "+ StrIncomeIndicator &" , Actual: "& strIserveSCIncomeInd, True
				bIncomeforNonStaff = True
			Else
				LogMessage "WARN", "Verification","Salary Crediting Annual Income UpdatedON doesnt match with the expected value. Expected: "+ StrIncomeIndicator &" , Actual: "& strIserveSCIncomeInd, False
				bIncomeforNonStaff = False
			End IF
		End IF
	End If
 End If
End If 
VerifyIncomeForNonStaff = bIncomeforNonStaff
Print "Function VerifyIncomeForNonStaff executed successfully"
End Function	

'[Verify Balance to Income History table details displayed as]
Public Function VerifyBIHistorytable(lstlstBalanceIncomeHistory)
	VerifyBIHistorytable=True
	VerifyBIHistorytable=verifyTableContentList(SuspensionDetails.BIHistoryHeader,SuspensionDetails.BIHistorycontent,lstlstBalanceIncomeHistory,"BalanceToIncomeHistory",false,null,null,null)
	'intRow=getRowForColumns(SuspensionDetails.BIHistoryHeader,SuspensionDetails.BIHistorycontent,arrColumns, arrValues)
 End Function
'
'[Verify Income details for staff and nonstaff account in Suspension Details screen displayed as]
Public Function VerifyIncomeForStaff(StrCardNumber,StrStaffIndicator,StrAppAnnualIncome,StrAppUpdatedOn,StrAUMIndicator,StrSCAnnualIncome,StrSCUpdatedOn,StrIncomeIndicator,StrBalanceIncomeHistory)
Print "Inside function VerifyIncomeForStaff"
bIncomeforStaff = True
If StrStaffIndicator = "Y" Then
	If Not IsNull(StrAppAnnualIncome) Then
		StrActAnnualIncome = SuspensionDetails.ApplicationAnnualIncome.GetROProperty("innertext")
		If Trim(StrActAnnualIncome) = Trim(StrAppAnnualIncome) Then
			LogMessage "RSLT", "Verification","Application Annual Income details matched with the expected value. Expected: "+ StrAppAnnualIncome &" , Actual: "& StrActAnnualIncome, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income details doesnt match with the expected value. Expected: "+ StrAppAnnualIncome &" , Actual: "& StrActAnnualIncome, False
			bIncomeforStaff = False
		End If
	End If
	
	If Not IsNull(StrAppUpdatedOn) Then
		StrActUpdatedOn = SuspensionDetails.AnnualIncomeUpdatedOn.GetROProperty("innertext")
		If Trim(StrActUpdatedOn) = Trim(StrAppUpdatedOn) Then
			LogMessage "RSLT", "Verification","Application Annual Income UpdatedON successfully matched with the expected value. Expected: "+ StrAppUpdatedOn &" , Actual: "& StrActUpdatedOn, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income UpdatedON doesnt match with the expected value. Expected: "+ StrAppUpdatedOn &" , Actual: "& StrActUpdatedOn, False
			bIncomeforStaff = False
		End If
	End If
	
	If Not IsNull(StrAUMIndicator) Then
		StrActAUMIndicator = SuspensionDetails.AUMIndicator.GetROProperty("innertext")
		If Trim(StrActAUMIndicator) = Trim(StrAUMIndicator) Then
			LogMessage "RSLT", "Verification","Application Annual Income AUM Indicator successfully matched with the expected value. Expected: "+ StrAUMIndicator &" , Actual: "& StrActAUMIndicator, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Application Annual Income AUM Indicator doesnt match with the expected value. Expected: "+ StrAUMIndicator &" , Actual: "& StrActAUMIndicator, False
			bIncomeforStaff = False
		End If
	End If
	
	If Not IsNull(StrSCAnnualIncome) Then
		StrActSCAnnualIncome =SuspensionDetails.SalaryCreditingAnnualIncome.GetROProperty("innertext")
		If Trim(StrActSCAnnualIncome) = Trim(StrSCAnnualIncome) Then
			LogMessage "RSLT", "Verification","Salary Crediting Annual Income value successfully matched with the expected value. Expected: "+ StrSCAnnualIncome &" , Actual: "& StrActSCAnnualIncome, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Salary Crediting Annual Income value doesnt match with the expected value. Expected: "+ StrSCAnnualIncome &" , Actual: "& StrActSCAnnualIncome, False
			bIncomeforStaff = False
		End If
	End If
	
	If Not IsNull(StrSCUpdatedOn) Then
	StrActSCUpdatedOn = SuspensionDetails.SalaryCreditingUpdatedOn.GetROProperty("innertext")
		If Trim(StrActSCUpdatedOn) = Trim(StrSCUpdatedOn) Then
			LogMessage "RSLT", "Verification","Salary Crediting Annual Income UpdatedON successfully matched with the expected value. Expected: "+ StrSCUpdatedOn &" , Actual: "& StrActSCUpdatedOn, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Salary Crediting Annual Income UpdatedON doesnt match with the expected value. Expected: "+ StrSCUpdatedOn &" , Actual: "& StrActSCUpdatedOn, False
			bIncomeforStaff = False
		End If
	End If
	
	If Not IsNull(StrIncomeIndicator) Then
	StrActIncomeIndicator = SuspensionDetails.IncomeIndicator.GetROProperty("innertext")
		If Trim(StrActIncomeIndicator) = Trim(StrIncomeIndicator) Then
			LogMessage "RSLT", "Verification","Salary Crediting Income Indicator successfully matched with the expected value. Expected: "+ StrIncomeIndicator &" , Actual: "& StrActIncomeIndicator, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Salary Crediting Income Indicator doesnt match with the expected value. Expected: "+ StrIncomeIndicator &" , Actual: "& StrActIncomeIndicator, False
			bIncomeforStaff = False
		End If
	End If
	
	If Not IsNull(StrBalanceIncomeHistory) Then
		StrActBIHistory =SuspensionDetails.BIHistoryRestrict.GetROProperty("innertext")
		If Trim(StrActBIHistory) = Trim(StrBalanceIncomeHistory) Then
			LogMessage "RSLT", "Verification","Balance to Income History value successfully matched with the expected value. Expected: "+ StrBalanceIncomeHistory &" , Actual: "& StrActBIHistory, True
			bIncomeforStaff = True
		Else
			LogMessage "WARN", "Verification","Balance to Income History value doesnt match with the expected value. Expected: "+ StrBalanceIncomeHistory &" , Actual: "& StrActBIHistory, False
			bIncomeforStaff = False
		End If
	End If
End If
VerifyIncomeForStaff = bIncomeforStaff
Print "Function VerifyIncomeForStaff executed successfully"
End Function

