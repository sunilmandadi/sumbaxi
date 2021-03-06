'[Verify field name in Account section displayed as]
Public Function VerifyPendingPayment_Acc(Strlabel)
	blabelExist = True 
	'StrIserveLabelname = BalancesAndLimits.lblBalanceLimits_PendingPayments.GetROProperty("innertext")
	StrIserveLabelname = "Pending Payments"
	If Strlabel <> "" Then
		StrIserveActualdisplayname = Instr(Strlabel, StrIserveLabelname)
		If StrIserveActualdisplayname = 1  Then
			LogMessage "RSLT","Verification","In Account section, new field 'Pending Payments' displayed as expected",True
			blabelExist = True 
		Else 
			LogMessage "RSLT","Verification","In Account section, new field 'Pending Payments' is not displayed",False
			blabelExist = False 
		End If
	End If
	VerifyPendingPayment_Acc = blabelExist
End Function

'[Verify field name in Relationship section displayed as]
Public Function VerifyPendingPayment_Relation(Strlabel)
	blabelExist = True 
	'StrIserveLabelname = BalancesAndLimits.lblRelationship_lblPendingPayments.GetROProperty("innertext")
	StrIserveLabelname = "Pending Payments"
	If Strlabel <> "" Then
		StrIserveActualdisplayname = Instr(Strlabel, StrIserveLabelname)
		If StrIserveActualdisplayname = 1  Then
			LogMessage "RSLT","Verification","In Relationship section, new field 'Pending Payments' displayed as expected",True
			blabelExist = True 
		Else 
			LogMessage "RSLT","Verification","In Relationship section, new field 'Pending Payments' is not displayed",False
			blabelExist = False 
		End If
	End If
	VerifyPendingPayment_Relation = blabelExist
End Function

'[User Click Link View Relationship Statement Summary in Statement page]
Public Function VerifyRelationshipStatementSummary(Strlink)
	bVerifyRelationshipStatementSummary = True 
	StrIserveLinkname = bcStatements.lnkRelationshipSummary.GetROProperty("innertext")
	If Strlink <> "" Then
		StrIserveActual = Instr(Strlink, StrIserveLinkname)
		If StrIserveActual = 1  Then
			bcStatements.lnkRelationshipSummary().Click
			LogMessage "RSLT","Verification","View Relationship Statement Summary link is displayed in statement page as expected.",True
			bVerifyRelationshipStatementSummary=True
		Else 
			LogMessage "RSLT","Verification","View Relationship Statement Summary link is not displayed",False
			bVerifyRelationshipStatementSummary=False
		End If
	End If
	VerifyRelationshipStatementSummary = bVerifyRelationshipStatementSummary
End Function

'[Verify Account Level Pending Payments in Balance and Limits screen displayed as]
Public Function VerifyPendingPayment_Accountsection(strCardNumber,strPendingpayment)
 	bVerifyPendingPayment_Accountsection=True
 	If Not IsNull(strPendingpayment) Then
		If strPendingpayment = "RUNTIME" Then
			Call getAccountPendingPayments_ARIQ_Vplus(strCardNumber)
			strVPBalNLimits_AccountPP=Environment.Value("StrExpAccMemoPayment")
			strIserveBalNLimits_AccountPP= BalancesAndLimits.lblAccountPendingPayments.GetROProperty("innertext")
			If  Ucase(Trim(strVPBalNLimits_AccountPP)) = Ucase(Trim(strIserveBalNLimits_AccountPP)) Then
				LogMessage "RSLT", "Verification","AccountLevel Pending payments value successfully matched with the expected value. Expected: "+ strVPBalNLimits_AccountPP &" , Actual: "& strIserveBalNLimits_AccountPP, True
				bVerifyPendingPayment_Accountsection=True
			else
				LogMessage "WARN", "Verification","AccountLevel Pending payments value doesnt match with the expected value. Expected: "+ strVPBalNLimits_AccountPP &" , Actual: "& strIserveBalNLimits_AccountPP, False
				bVerifyPendingPayment_Accountsection=False
			End If
		End If
    End If
    VerifyPendingPayment_Accountsection=bVerifyPendingPayment_Accountsection
  End Function
    
'[Verify Relationship Level Pending Payments in Balance and Limits screen displayed as]
  Public Function VerifyPendingPayment_RelationshipSection(strCardNumber,strPendingpayment)
 	bVerifyPendingPayment_RelationshipSection = True
 	If Not IsNull(strPendingpayment) Then
		If strPendingpayment = "RUNTIME" Then
			Call getAccountPendingPayments_ARIQ_Vplus(strCardNumber)
			     StrOrgNumber=Environment.Value("StrOrgnizationNumber")
				 StrRelNumber=Environment.Value("StrRelationshipNumber")
			Call getRelationshipPendingPayments_ARIG_Vplus(StrOrgNumber,StrRelNumber)
				strVPBalNLimits_RelationshipPP=Environment.Value("StrExpRelMemoPayments")
				strIserveBalNLimits_RelationshipPP=BalancesAndLimits.lblRelationship_PendingPayments.GetROProperty("innertext")
			If  Ucase(Trim(strVPBalNLimits_RelationshipPP)) = Ucase(Trim(strIserveBalNLimits_RelationshipPP)) Then
				LogMessage "RSLT", "Verification","Relationship section Pending payments value successfully matched with the expected value. Expected: "+ strVPBalNLimits_RelationshipPP &" , Actual: "& strIserveBalNLimits_RelationshipPP, True
				bVerifyPendingPayment_RelationshipSection=True
			else
				LogMessage "WARN", "Verification","Relationship section Pending payments not matched with the expected value. Expected: "+ strVPBalNLimits_RelationshipPP &" , Actual: "& strIserveBalNLimits_RelationshipPP, False
				bVerifyPendingPayment_RelationshipSection=False
			End If
		End If
    End If
    VerifyPendingPayment_RelationshipSection=bVerifyPendingPayment_RelationshipSection
  End Function
  
'[Select Radio Button from Transaction History screen displayed as]
 Public Function SelectTransactionType(strProductType, strType)
 	bSelectTransactionType = True
 	Select Case Ucase(strProductType)
	Case "CREDIT CARD"
		Select Case strType
			Case "Pending"
			TransactionHistory.rbtnPending.Click
			Case "Declined"
			TransactionHistory.rbtnDeclined.Click
		End Select
		LogMessage "RSLT","Verification","Radio button "&strType&"Selected Successfully for Credit Card",true
		SelectTransactionType = bSelectTransactionType
	Case "CASHLINE"
			Case "Pending"
			TransactionHistory.rbtnPending.Click
			Case "Declined"
			TransactionHistory.rbtnDeclined.Click
		End Select
		LogMessage "RSLT","Verification","Radio button "&strType&"Selected Successfully for CashLine",true
		SelectTransactionType = bSelectTransactionType
 End Function
  
'[Verify Pending Transaction Description and Payment Indicator values displayed as]
Public Function VerifyPendingTransactionHistorytable(lstlstTransactionHistory)
	bVerifyPendingTransactionHistorytable=True
	VerifyPendingTransactionHistorytable=verifyTableContentList(TransactionHistory.tblTransactionsHeader_Pend,TransactionHistory.tblTransactions_Pend,lstlstTransactionHistory,"Transaction History for Pending",false,null,null,null)
 End Function
 
'[Verify Declined Transaction Description and Payment Indicator values displayed as]
Public Function VerifyDeclinedTransactionHistorytable(lstlstTransactionHistory)
	bVerifyPendingTransactionHistorytable=True
	VerifyDeclinedTransactionHistorytable=verifyTableContentList(TransactionHistory.tblTransactionsHeader_Decl,TransactionHistory.tblTransactions_Decl,lstlstTransactionHistory,"Transaction History for Declined",false,null,null,null)
 End Function
  
'[Verify Statement Balance value in Statement screen displayed for]
Public Function VerifyCurStmtBal(strCardNumber)
	bVerifyCurStmtBal = True 
    Dim lstlstStatementSummaryData
    'Call Visionplus ARSD system to validate the data required
	Call getStatementHistory_ARSD_Vplus(strCardNumber)
	 strVPStatementBalance=FormatNumber((Environment.Value("StrEndBalance")),2)
	 strVPStatementDue=FormatNumber((Environment.Value("StrStatementDue")),2)
	 strVPStatementDueDate=Trim(Environment.Value("StrStatementDueDate")) 
	'Concatente to Array format for the values derived from host system)	
	lstlstStatementSummaryData = (checknull("(Current Statement Balance:"&strVPStatementBalance&"|Statement Due:"&strVPStatementDue&"|Statement Due Date:"&strVPStatementDueDate&")|"))
	'Call function to compare the table value with the values retrieved from Host system	
	bVerifyCurStmtBal=VerifyStatementSummarytable(lstlstStatementSummaryData)
    VerifyCurStmtBal=bVerifyCurStmtBal
End Function

'[Verify Online Statement Balance value in Statement screen displayed as]
Public Function VerifyOutstandingStmtBal(strCardNumber)
	bVerifyOutstandingStmtBal = True 
	If Not IsNull(strCardNumber) Then
			Call getAccountPendingPayments_ARIQ_Vplus(strCardNumber)
			'Call getStatementHistory_ARSD_Vplus(strCardNumber)
			Call getPaymentdetails_ARQA_Vplus(strCardNumber)
			strVPOutstandingStatementBalance = fcaluclateOutstandingStatementBalance
			If strVPOutstandingStatementBalance < 0 or strVPOutstandingStatementBalance = 0 Then
				strVPOutstandingStatementBalance = "0.00"
			End If
			strVPOutstandingStatementDue = fcaluclateOutstandingStatementDue
			If strVPOutstandingStatementDue < 0 or strVPOutstandingStatementDue= 0 Then
				strVPOutstandingStatementDue = "0.00"
			End If
			If strVPOutstandingStatementBalance = 0 Or strVPOutstandingStatementDue = 0 Then
				StrVPOutstandingCurrentDueDate = "Null"
			Else 
				StrVPOutstandingCurrentDueDate = Environment.Value("StrCurrentDueDate")
			End If	
	'Concatente to Array format for the values derived from host system)	
	lstlstStatementSummaryData = (checknull("(Outstanding Statement Balance:"&strVPOutstandingStatementBalance&"|Outstanding Statement Due:"&strVPOutstandingStatementDue&"|Current Due Date:"&StrVPOutstandingCurrentDueDate&")|"))
	'Call function to compare the table value with the values retrieved from Host system	
	bVerifyOutstandingStmtBal=VerifyStatementSummarytable(lstlstStatementSummaryData)    
    End If 	
    VerifyOutstandingStmtBal =bVerifyOutstandingStmtBal
End Function

'[Verify Statement Balance value in Relationship Statement Summary screen displayed for]
Public Function VerifyRelationshipStmtBal(strCardNumber)
	bVerifyRelationshipStmtBal = True 
	Call getAccountPendingPayments_ARIQ_Vplus(strCardNumber)
	StrOrgNumber = Environment.Value("StrOrgnizationNumber")
	StrRelNumber = Environment.Value("StrRelationshipNumber")
	Call getRelationshipPendingPayments_ARIG_Vplus(StrOrgNumber,StrRelNumber)
	'Vision Plus values 
		 strVPStatementBalance=Environment.Value("StrEndingStmtBalance")
		 strVPStatementDue= Environment.Value("StrTotalPaymentDue")
    'IServe Values 
    	strIserveStatementBalance=bcStatements.lblRelationshipSummary_StatementBalance.GetROProperty("innertext")
		strIserveStatementDue= bcStatements.lblRelationshipSummary_TotalDueStatement.GetROProperty("innertext")
	'Comparing VisionPlus values with the Iserve Values
		If Ucase(Trim(strVPStatementBalance)) = 0 And  Ucase(Trim(strIserveStatementBalance)) = "0.00" Then
			LogMessage "RSLT", "Verification","RelationshipLevel Balance statement value matched as expected. Expected: "+ strVPStatementBalance &" , Actual: "& strIserveStatementBalance, True
			bVerifyRelationshipStmtBal = True
		Else  If Ucase(Trim(strVPStatementBalance)) = Ucase(Trim(strIserveStatementBalance)) Then
			LogMessage "RSLT", "Verification","RelationshipLevel Balance statement value matched as expected. Expected: "+ strVPStatementBalance &" , Actual: "& strIserveStatementBalance, True
			bVerifyRelationshipStmtBal = True
		Else
			LogMessage "WARN", "Verification","RelationshipLevel Balance statement value doesnt match. Expected: "+ strVPStatementBalance &" , Actual: "& strIserveStatementBalance, False
			bVerifyRelationshipStmtBal = False
		End If	
		End If 
		If Ucase(Trim(strVPStatementDue)) = 0  And  Ucase(Trim(strIserveStatementDue)) = "0.00" Then
			LogMessage "RSLT", "Verification","RelationshipLevel statement TotalDue value matched as expected. Expected: "+ strVPStatementDue &" , Actual: "& strIserveStatementDue, True
			bVerifyRelationshipStmtBal = True
		Else If Ucase(Trim(strVPStatementDue)) = Ucase(Trim(strIserveStatementDue)) Then
			LogMessage "RSLT", "Verification","RelationshipLevel statement TotalDue value matched as expected. Expected: "+ strVPStatementDue &" , Actual: "& strIserveStatementDue, True
			bVerifyRelationshipStmtBal = True
		Else
			LogMessage "WARN", "Verification","RelationshipLevel statement TotalDue value doesnt match. Expected: "+ strVPStatementDue &" , Actual: "& strIserveStatementDue, False
			bVerifyRelationshipStmtBal = False
		End If	
		End If
		VerifyRelationshipStmtBal =bVerifyRelationshipStmtBal
End Function

'[Verify Online Statement Balance value in Relationship Statement Summary screen displayed as]
Public Function VerifyRelOutstandingStmtBal(strCardNumber)
	bVerifyRelOutstandingStmtBal = True 
	If Not IsNull(strCardNumber) Then
			Call getAccountPendingPayments_ARIQ_Vplus(strCardNumber)
			StrOrgNumber = Environment.Value("StrOrgnizationNumber")
			StrRelNumber = Environment.Value("StrRelationshipNumber")
			Call getRelationshipPendingPayments_ARIG_Vplus(StrOrgNumber,StrRelNumber)
			Call getPaymentdetails_ARQA_Vplus(strCardNumber)
			strVPOutstandingStatementBalance = fcaluclateRelationshipOutstandingBalance
			strVPOutstandingStatementDue = fcaluclateRelationshipOutstandingTotalDue
			If strVPOutstandingStatementBalance = 0 Or strVPOutstandingStatementDue = 0 Then
				StrVPCurrentDueDate = ""
			Else If strVPOutstandingStatementBalance < 0 Or strVPOutstandingStatementDue < 0 Then
				StrVPCurrentDueDate = ""
			Else
				StrVPCurrentDueDate = Environment.Value("StrCurrentDueDate")
			End If	
			End If
		'IServe Values 
    	strIserveOutstandingStatementBalance=bcStatements.lblRelationshipSummary_OustandingBalance.GetROProperty("innertext")
		strIserveOutstandingStatementDue=bcStatements.lblRelationshipSummary_TotalDueOustanding.GetROProperty("innertext")
		strIserveCurrentDueDate=bcStatements.lblRelationshipSummary_CurrentDueDate.GetROProperty("innertext")
		
	'Comparing VisionPlus values with the Iserve Values 
		' For Outstanding Statement Balance
		If Ucase(Trim(strVPOutstandingStatementBalance)) < 0  Then
			strIserveOutstandingStatementBalance = "0.00"
			LogMessage "RSLT", "Verification","RelationshipLevel Outstanding statement balance matched as expected. Expected: "+ strVPOutstandingStatementBalance &" , Actual: "& strIserveOutstandingStatementBalance, True
			bVerifyRelationshipStmtBal = True
		Else If Ucase(Trim(strVPOutstandingStatementBalance)) = 0 And  Ucase(Trim(strIserveOutstandingStatementBalance)) = "0.00" Then
			LogMessage "RSLT", "Verification","RelationshipLevel Outstanding statement balance matched as expected. Expected: "+ strVPOutstandingStatementBalance &" , Actual: "& strIserveOutstandingStatementBalance, True
			bVerifyRelOutstandingStmtBal = True
		Else  If Ucase(Trim(strVPOutstandingStatementBalance)) = Ucase(Trim(strIserveOutstandingStatementBalance)) Then
			LogMessage "RSLT", "Verification","RelationshipLevel Outstanding statement balance matched as expected. Expected: "+ strVPOutstandingStatementBalance &" , Actual: "& strIserveOutstandingStatementBalance, True
			bVerifyRelOutstandingStmtBal = True
		Else
			LogMessage "WARN", "Verification","RelationshipLevel Outstanding statement balance doesnt match. Expected: "+ strVPOutstandingStatementBalance &" , Actual: "& strIserveOutstandingStatementBalance, False
			bVerifyRelOutstandingStmtBal = False
		End If	
		End If 
		End If
		' For Outstanding Statement Due 
		If Ucase(Trim(strVPOutstandingStatementDue)) < 0  Then
  	        strIserveOutstandingStatementDue = "0.00" 
			LogMessage "RSLT", "Verification","RelationshipLevel Outstanding statement TotalDue matched as expected. Expected: "+ strVPOutstandingStatementDue &" , Actual: "& strIserveOutstandingStatementDue, True
			bVerifyRelOutstandingStmtBal = True
		If Ucase(Trim(strVPOutstandingStatementDue)) = 0 And  Ucase(Trim(strIserveOutstandingStatementDue)) = "0.00" Then
			LogMessage "RSLT", "Verification","RelationshipLevel Outstanding statement TotalDue matched as expected. Expected: "+ strVPOutstandingStatementDue &" , Actual: "& strIserveOutstandingStatementDue, True
			bVerifyRelOutstandingStmtBal = True
		Else  If Ucase(Trim(strVPOutstandingStatementDue)) = Ucase(Trim(strVPOutstandingStatementDue)) Then
			LogMessage "RSLT", "Verification","RelationshipLevel Outstanding statement TotalDue matched as expected. Expected: "+ strVPOutstandingStatementDue &" , Actual: "& strIserveOutstandingStatementDue, True
			bVerifyRelOutstandingStmtBal = True
		Else
			LogMessage "WARN", "Verification","RelationshipLevel Outstanding statement TotalDue doesnt match. Expected: "+ strVPOutstandingStatementDue &" , Actual: "& strIserveOutstandingStatementDue, False
			bVerifyRelOutstandingStmtBal = False
		End If	
		End If
		End If
		' For Current Due Date 
		If StrVPCurrentDueDate = strIserveCurrentDueDate  Then
			LogMessage "RSLT", "Verification","RelationshipLevel Current Total Due Date value matched as expected. Expected: "+ StrVPCurrentDueDate &" , Actual: "& strIserveCurrentDueDate, True
			bVerifyRelOutstandingStmtBal = True
		Else
			LogMessage "WARN", "Verification","RelationshipLevel Current Total Due Date value doesnt match. Expected: "+ StrVPCurrentDueDate &" , Actual: "& strIserveCurrentDueDate, False
			bVerifyRelOutstandingStmtBal = False
		End If	
    End If 	
    VerifyRelOutstandingStmtBal = bVerifyRelOutstandingStmtBal
End Function


'/******************** Verify list of List values for statement and Transaction History table *****************************/
Public Function VerifyStatementSummarytable(lstlstStatementSummaryData)
	bVerifyStatementSummarytable=True
	VerifyStatementSummarytable=verifyTableContentList(bcStatements.tblStatementSummaryHeader,bcStatements.tblStatementSummaryContent,lstlstStatementSummaryData,"StatementSummary",false,null,null,null)
 End Function


 '/********* BEG BAL (ARIQ01 ) - MEMO PAYMENT (ARIQ40) – CREDIT (ARQA03 for all plans) + CTD PMT RVRSL (ARQA06 for all plans)*****************/
Public Function fcaluclateOutstandingStatementBalance()
	VPBeginingStatementBalance = Environment.Value("StrBeginningBalance")
	VPMemoPayment = Environment.Value("StrExpAccMemoPayment")
	VPCreditforallPlans = Environment.Value("StrCreditBalance")
	VPPaymentReversal= Environment.Value("StrPaymentReversal")
	ExpAccountstatementBalance = CCur(VPBeginingStatementBalance) - CCur(VPMemoPayment) - CCur(VPCreditforallPlans) + CCur(VPPaymentReversal)
	fcaluclateOutstandingStatementBalance = FormatNumber(ExpAccountstatementBalance,2)
End Function


'/***********************TOT AMT DUE (ARIQ01) - MEMO PAYMENT (ARIQ40)*********************/
Public Function fcaluclateOutstandingStatementDue()
	VPTotalAmountDue = Environment.Value("StrStmtTotalDue")
	VPMemoPayment = Environment.Value("StrExpAccMemoPayment")
	ExpTotalDue= CCur(VPTotalAmountDue) - CCur(VPMemoPayment)
	fcaluclateOutstandingStatementDue = FormatNumber(ExpTotalDue,2)
End Function


'/********** ENDING STMT BAL (ARIG01 ) - MEMO PAYMENT (ARIG42) – CREDIT (ARQA03 for all plans) + CTD PMT RVRSL (ARQA06 for all plans)*********************/

Function fcaluclateRelationshipOutstandingBalance()
	VPEndingStatementBalance = Environment.Value("StrEndingStmtBalance")
	VPMemoPayment = Environment.Value("StrExpRelMemoPayments")
	VPCreditforallPlans = Environment.Value("StrCreditBalance")
	VPPaymentReversal= Environment.Value("StrPaymentReversal")
	ExpRelOutstandingBalance = CCur(VPEndingStatementBalance)- Ccur(CVPMemoPayment)-Ccur(VPCreditforallPlans)+Ccur(VPPaymentReversal) 
	fcaluclateRelationshipOutstandingBalance = FormatNumber(ExpRelOutstandingBalance,2)
End Function

'/********************* TOT PMT DUE (ARIG01) - MEMO PAYMENT (ARIG42)*********************/

Function fcaluclateRelationshipOutstandingTotalDue()
	VPTotalAmountDue = Environment.Value("StrTotalPaymentDue")
	VPMemoPayment = Environment.Value("StrExpRelMemoPayments")
	ExpTotalDue = CCur(VPTotalAmountDue) - CCur(VPMemoPayment)
	fcaluclateRelationshipOutstandingTotalDue = FormatNumber(ExpTotalDue,2)
End Function

'[LISA Verify Account Level Pending Payments in Balance and Limits screen displayed as]
Public Function verifyPendingPayments_BalancesLimits(strPendingPayments)
	'Iserve field
	strIservePendingPayments_BalancesLimits = BalancesAndLimits.lblBalanceLimits_PendingPayments.GetROProperty("innertext")
	bDevPending=false
   bverifyPendingPayments_BalancesLimits=true
   If strPendingPayments = strIservePendingPayments_BalancesLimits Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Pending Payments in accounts section is as expected: "&strPendingPayments&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Pending Payments in accounts section is not as expected: "&strPendingPayments&"",False
		End if 
	verifyPendingPayments_BalancesLimits = bverifyPendingPayments_BalancesLimits
End Function

'[LISA Verify Relationship Level Pending Payments in Balance and Limits screen displayed as]
Public Function verifyPendingPayments_RelationshipDetails(strPendingPayments)
	'Iserve field
	strIservePendingPayments_RelationshipDetails=BalancesAndLimits.lblRelationship_PendingPayments.GetROProperty("innertext")
	bDevPending=false
    bverifyPendingPayments_RelationshipDetails=true
   If strPendingPayments = strIservePendingPayments_RelationshipDetails Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Pending Payments in relationship section is as expected:"&strPendingPayments&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Pending Payments in relationship section is not as expected:"&strPendingPayments&"",False
		End if 
	verifyPendingPayments_RelationshipDetails = bverifyPendingPayments_RelationshipDetails
End Function

'[LISA Verify Statement Balance and Online Statement Balance value in Statement screen displayed for]
Public Function VerifyStatementBalanceDue(lstlstStatementSummary)
	bVerifyStatementBalanceDue=True
	VerifyStatementBalanceDue=verifyTableContentList(bcStatements.tblStatementSummaryHeader,bcStatements.tblStatementSummaryContent,lstlstStatementSummary,"Statement Summary",false,null,null,null)
	VerifyStatementBalanceDue =bVerifyStatementBalanceDue
 End Function
 
'[LISA Verify Statement Balance and Online Statement Balance in Relationship Summary displayed as]
Public Function verifyStatementOnlineBal_RelationshipSummary(strStatBalance,strStatDue,strOutStatBalance,strOutStatDue,strCurrDue)
	'IServe Fields present in Relationship Summary pop up
	
	strIserveStatBalance = bcStatements.lblRelationshipSummary_StatementBalance.GetROProperty("innertext")
	strIserveStatDue = bcStatements.lblRelationshipSummary_TotalDueStatement.GetROProperty("innertext")
	strIserveOutStatBalance = bcStatements.lblRelationshipSummary_OustandingBalance.GetROProperty("innertext")
	strIserveOutStatDue = bcStatements.lblRelationshipSummary_TotalDueOustanding.GetROProperty("innertext")
	strIserveCurrDue = bcStatements.lblRelationshipSummary_CurrentDueDate.GetROProperty("innertext")
	bDevPending=false
    bverifyStatementOnlineBal_RelationshipSummary=true
    
    If strStatBalance = strIserveStatBalance Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Statement Balance in relationship summary is as expected: "&strStatBalance&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Statement Balance in relationship summary is not as expected: "&strStatBalance&"",False
		End if 
    
    If strStatDue = strIserveStatDue Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Statement Due in relationship summary is as expected: "&strStatDue&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Statement Due in relationship summary is not as expected: "&strStatDue&"",False
		End if 
		
	If strOutStatBalance = strIserveOutStatBalance Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Outstanding Balance in relationship summary is as expected: "&strOutStatBalance&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Outstanding Balance in relationship summary is not as expected: "&strOutStatBalance&"",False
		End if 	
		
	If strOutStatDue = strIserveOutStatDue Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Outstanding Statement Due in relationship summary is as expected: "&strOutStatDue&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Outstanding Statement Due in relationship summary is not as expected: "&strOutStatDue&"",False
		End if 	
		
	If strCurrDue = strIserveCurrDue Then
   	
	  	LogMessage "RSLT","Verification","The Iserve Current Due in relationship summary is as expected: "&strCurrDue&"",True
		Else
	  	LogMessage "WARN","Verification","The Iserve Current Due in relationship summary is not as expected: "&strCurrDue&"",False
		End if 		
		verifyStatementOnlineBal_RelationshipSummary = bverifyStatementOnlineBal_RelationshipSummary
End Function

'*************************************Balances and Limits in BDT framework by Aniket********************************

'[Verify the fields under balances and limits section under Account category]
Public Function verifyfields_BalandLimitsAccount(lstBalLimits)
	bverifyfields_BalandLimitsAccount = true
	intSize = Ubound(lstBalLimits)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstBalLimits(Iterator),":")(0))
		arrValue = trim(Split(lstBalLimits(Iterator),":")(1))
		
	Select Case (arrLabel)		
		Case "Current Balance"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblAccountBalance_AvailableBalance(), arrValue, "Current Balance")Then
				LogMessage "RSLT","Verification","Balances and Limits - Current Balance:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Pending Debits"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblPendingDebits(), arrValue, "Pending Debits")Then
				LogMessage "RSLT","Verification","Balances and Limits - Pending Debits:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Pending Credits"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblPendingCredits(), arrValue, "Pending Credits")Then
				LogMessage "RSLT","Verification","Balances and Limits - Pending Credits:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Pending Payments"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblBalanceLimits_PendingPayments(), arrValue, "Pending Payments")Then
				LogMessage "RSLT","Verification","Balances and Limits - Pending Payments:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Outstanding Balance"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblOutstandingBalance(), arrValue, "Outstanding Balance")Then
				LogMessage "RSLT","Verification","Balances and Limits - Outstanding Balance:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Total Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblTotalCreditLimit(), arrValue, "Total Credit Limit")Then
				LogMessage "RSLT","Verification","Balances and Limits - Total Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		
		Case "Total Available Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblAvailableLimit(), arrValue, "Total Available Limit")Then
				LogMessage "RSLT","Verification","Balances and Limits - Total Available Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Overlimit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblBalancesLimits_OverLimit(), arrValue, "Overlimit")Then
				LogMessage "RSLT","Verification","Balances and Limits - Overlimit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
	End select
	next 
	verifyfields_BalandLimitsAccount = bverifyfields_BalandLimitsAccount
End Function

'[Verify the fields under Cash Advance section under Account category]
Public Function verifyfields_CashAdvanceAccount(lstCashAdvance)
	bverifyfields_CashAdvanceAccount = true
	intSize = Ubound(lstCashAdvance)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstCashAdvance(Iterator),":")(0))
		arrValue = trim(Split(lstCashAdvance(Iterator),":")(1))
		
	Select Case (arrLabel)		
		Case "Current"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCashAdvance_Current(), arrValue, "Current")Then
				LogMessage "RSLT","Verification","Cash Advance - Current:"&arrValue&" is not displayed as expected",false
				bverifyfields_CashAdvanceAccount=false
			End If
		End If
		
		Case "Outstanding"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCashAdvance_Outstanding(), arrValue, "Outstanding")Then
				LogMessage "RSLT","Verification","Cash Advance - Outstanding:"&arrValue&" is not displayed as expected",false
				bverifyfields_CashAdvanceAccount=false
			End If
		End If
		
		Case "Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCashAdvance_CreditLimit(), arrValue, "Credit Limit")Then
				LogMessage "RSLT","Verification","Cash Advance - Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_CashAdvanceAccount=false
			End If
		End If
		
		Case "Available Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCashAdvance_AvailableLimit(), arrValue, "Available Limit")Then
				LogMessage "RSLT","Verification","Cash Advance - Available Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_CashAdvanceAccount=false
			End If
		End If
		
	End select
	next 
	verifyfields_CashAdvanceAccount = bverifyfields_CashAdvanceAccount
End Function

'[Verify the fields under Retail section under Account category]
Public Function verifyfields_RetailAccount(lstRetail)
	bverifyfields_RetailAccount = true
	intSize = Ubound(lstRetail)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstRetail(Iterator),":")(0))
		arrValue = trim(Split(lstRetail(Iterator),":")(1))
		
	Select Case (arrLabel)		
		Case "Current"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_Current(), arrValue, "Current")Then
				LogMessage "RSLT","Verification","Retail - Current:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		
		Case "Outstanding"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_Outstanding(), arrValue, "Outstanding")Then
				LogMessage "RSLT","Verification","Retail - Outstanding:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		
		Case "Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_CreditLimit(), arrValue, "Credit Limit")Then
				LogMessage "RSLT","Verification","Retail - Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		
		Case "Available Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_AvailableLimit(), arrValue, "Available Limit")Then
				LogMessage "RSLT","Verification","Retail - Available Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		Case "Temporary Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_TempCreditLimit(), arrValue, "Temporary Credit Limit")Then
				LogMessage "RSLT","Verification","Retail - Temporary Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		Case "Effective Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_EffectiveDate(), arrValue, "Effective Date")Then
				LogMessage "RSLT","Verification","Retail - Effective Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		Case "Expiry Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_ExpiryDate(), arrValue, "Expiry Date")Then
				LogMessage "RSLT","Verification","Retail - Expiry Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
		Case "Change Reason"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRetail_ChangeReason(), arrValue, "Change Reason")Then
				LogMessage "RSLT","Verification","Retail - Change Reason:"&arrValue&" is not displayed as expected",false
				bverifyfields_RetailAccount=false
			End If
		End If
	End select
	next 
	verifyfields_RetailAccount = bverifyfields_RetailAccount
End Function

'[Verify the fields under CardLimits section under Account category]
Public Function verifyfields_CardLimitsAccount(lstCardLimit)
	bverifyfields_CardLimitsAccount = true
	intSize = Ubound(lstCardLimit)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstCardLimit(Iterator),":")(0))
		arrValue = trim(Split(lstCardLimit(Iterator),":")(1))
		
	Select Case (arrLabel)	
		Case "Per Day"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_WithdrawalPerDay(), arrValue, "Per Day")Then
				LogMessage "RSLT","Verification","Card Limit - Per Day:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		Case "Per Transaction"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_WithdrawalPerTransaction(), arrValue, "Per Transaction")Then
				LogMessage "RSLT","Verification","Card Limit - Per Transaction:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		
		Case "Eligible Transactions"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_EligibleTransactions(), arrValue, "Eligible Transactions")Then
				LogMessage "RSLT","Verification","Card Limit - Eligible Transactions:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		Case "Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_CreditLimit(), arrValue, "Credit Limit")Then
				LogMessage "RSLT","Verification","Card Limit - Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		
		Case "Available Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_AvailableLimit(), arrValue, "Available Limit")Then
				LogMessage "RSLT","Verification","Card Limit - Available Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		
		Case "Temporary Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_TempCreditLimit(), arrValue, "Temporary Credit Limit")Then
				LogMessage "RSLT","Verification","Card Limit - Temporary Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		
		Case "Effective Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_EffectiveDate(), arrValue, "Effective Date")Then
				LogMessage "RSLT","Verification","Card Limit - Effective Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
			End If
		End If
		
		Case "Expiry Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblCardLimit_ExpiryDate(), arrValue, "Expiry Date")Then
				LogMessage "RSLT","Verification","Card Limit - Expiry Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_CardLimitsAccount=false
				End If
			End If		
		End select
		Next
		verifyfields_CardLimitsAccount = bverifyfields_CardLimitsAccount	
End Function 

'[Verify the fields under BalanceLimits section under Relationship category]
Public Function verifyfields_BalLimitsRelationship(lstCardLimit)
	bverifyfields_BalLimitsRelationship = true
	intSize = Ubound(lstCardLimit)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstCardLimit(Iterator),":")(0))
		arrValue = trim(Split(lstCardLimit(Iterator),":")(1))
		
	Select Case (arrLabel)		
		Case "Current Balance"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_CurrentBalance(), arrValue, "Current Balance")Then
				LogMessage "RSLT","Verification","Balances - Current Balance:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			  End If
	       End If
	    Case "Pending Debits"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_PendingDebits(), arrValue, "Pending Debits")Then
				LogMessage "RSLT","Verification","Balances and Limits - Pending Debits:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If
		Case "Pending Credits"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_PendingCredits(), arrValue, "Pending Credits")Then
				LogMessage "RSLT","Verification","Balances and Limits - Pending Credits:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalandLimitsAccount=false
			End If
		End If   
	    Case "Pending Payments"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_PendingPayments(), arrValue, "Pending Payments")Then
				LogMessage "RSLT","Verification","Balances - Pending Payments:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Outstanding Balance"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_OutstandingBalance(), arrValue, "Outstanding Balance")Then
				LogMessage "RSLT","Verification","Balances - Outstanding Balance:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_CreditLimit(), arrValue, "Credit Limit")Then
				LogMessage "RSLT","Verification","Limits - Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Available Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_AvailableCreditLimit(), arrValue, "Available Limit")Then
				LogMessage "RSLT","Verification","Limits - Available Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Overlimit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_OverLimit(), arrValue, "Overlimit")Then
				LogMessage "RSLT","Verification","Limits - Overlimit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Temporary Credit Limit"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_TempCreditLimit(), arrValue, "Temporary Credit Limit")Then
				LogMessage "RSLT","Verification","Limits - Temporary Credit Limit:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Effective Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_EffectiveDate(), arrValue, "Effective Date")Then
				LogMessage "RSLT","Verification","Limits - Effective Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
			End If
		End If
		
		Case "Expiry Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText(BalancesAndLimits.lblRelationship_ExpiryDate(), arrValue, "Expiry Date")Then
				LogMessage "RSLT","Verification","Limits - Expiry Date:"&arrValue&" is not displayed as expected",false
				bverifyfields_BalLimitsRelationship=false
				End If
			End If		
		End select
		Next
		verifyfields_BalLimitsRelationship =bverifyfields_BalLimitsRelationship
End Function 


        '*****************************Function on the Screen **********************************************************

'Public Function verifyBalanceAndLimits( strCurrentBalance, strPendingDebits, strPendingCredits, strOutstandingBalance, strTotalCreditLimit, strAvailableLimit,  _
'strCashAdvance_Current, strCashAdvance_Outstanding, strCashAdvance_CreditLimit, strCashAdvance_AvailableLimit,  _
'strRetail_Current, strRetail_Outstanding, strRetail_CreditLimit, strRetail_AvailableLimit, strRetail_TempCreditLimit, strRetail_EffectiveDate, strRetail_ExpiryDate, strRetail_ChangeReason, _
'strCardLimit_WithdrawalPerDay, strCardLimit_WithdrawalPerTransaction, strCardLimit_EligibleTransactions, strCardLimit_CreditLimit, strCardLimit_AvailableLimit, strCardLimit_TempCreditLimit, strCardLimit_EffectiveDate, strCardLimit_ExpiryDate, strCardLimits_RTLPerDay, strCardLimits_RTLPerMonth, strCardLimits_RTLPerYear, _
' strRelationship_CurrentBalance, strRelationship_PendingDebits, strRelationship_PendingCredits, strRelationship_OutstandingBalance, strRelationship_CreditLimit, strRelationship_TempCreditLimit, strRelationship_AvailableCreditLimit, strRelationship_EffectiveDate, strRelationship_ExpiryDate)
'		bVerifyBalanceAndLimits=true
'	bcAccountOverview_LeftMenu.clickBalanceLimits()
'	WaitForICallLoading
'	If Not pageExists() Then
'		LogMessage "WARN","Verification","Statement Details page does not displayed",false
'		bVerifyBalanceAndLimits=false
'	Else
'		LogMessage "RSLT","Verification","Statement Details page displayed Successfully",true
'	End If
'   If Not IsNull(strCurrentBalance) Then
'		If Not verifyInnerText(lblCurrentBalance() , strCurrentBalance, "Current Balance")Then
'					bVerifyBalanceAndLimits = False
'			End If
'    End If
'
'    If Not IsNull(strPendingDebits) Then
'		If Not verifyInnerText(lblPendingDebits() , strPendingDebits, "Pending Debits")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strPendingCredits) Then
'		If Not verifyInnerText( lblPendingCredits(), strPendingCredits, "Pending Credit")Then
'					bVerifyBalanceAndLimits = False
'		End If		              
'    End If
'
'    If Not IsNull(strOutstandingBalance) Then
'		If Not verifyInnerText(  lblOutstandingBalance(), strOutstandingBalance, "Outstanding Balance")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strTotalCreditLimit) Then
'		If Not verifyInnerText(  lblTotalCreditLimit(), strTotalCreditLimit, "Total Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strAvailableLimit) Then
'		If Not verifyInnerText(  lblAvailableLimit(), strAvailableLimit, "Available Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCashAdvance_Current) Then
'		If Not verifyInnerText( lblCashAdvance_Current(), strCashAdvance_Current, "Cash Advance Current")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCashAdvance_Outstanding) Then
'      If Not verifyInnerText( lblCashAdvance_Outstanding(), strCashAdvance_Outstanding, "Cash Advance Outstanding")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCashAdvance_CreditLimit) Then
'		If Not verifyInnerText(lblCashAdvance_CreditLimit(), strCashAdvance_CreditLimit, "Cash Advance Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCashAdvance_AvailableLimit) Then
'      If Not verifyInnerText(lblCashAdvance_AvailableLimit(), strCashAdvance_AvailableLimit, "Cash Advance Available Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'		
'    End If
'
'    If Not IsNull(strRetail_Current) Then
'		If Not verifyInnerText( lblRetail_Current(), strRetail_Current, "Retail Current")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_Outstanding) Then
'      If Not verifyInnerText(lblRetail_Outstanding(), strRetail_Outstanding, "Retail Outstanding")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_CreditLimit) Then
'		 If Not verifyInnerText(lblRetail_CreditLimit(), strRetail_CreditLimit, "Retail Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_AvailableLimit) Then
'		If Not verifyInnerText(lblRetail_AvailableLimit(), strRetail_AvailableLimit,"Retail Available Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_TempCreditLimit) Then
'      If Not verifyInnerText(lblRetail_TempCreditLimit(), strRetail_TempCreditLimit,"Retail Temporary Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_EffectiveDate) Then
'      If Not verifyInnerText(lblRetail_EffectiveDate(), strRetail_EffectiveDate,"Retail Effective Date")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_ExpiryDate) Then
'      If Not verifyInnerText( lblRetail_ExpiryDate(), strRetail_ExpiryDate,"Retail Expiry Date")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRetail_ChangeReason) Then
'      If Not verifyInnerText(lblRetail_ChangeReason(), strRetail_ChangeReason,"Retail Change Reason")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_WithdrawalPerDay) Then
'      If Not verifyInnerText(lblCardLimit_WithdrawalPerDay(), strCardLimit_WithdrawalPerDay,"Card Limit Withdrawal Limit Per Day")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_WithdrawalPerTransaction) Then
'      If Not verifyInnerText(lblCardLimit_WithdrawalPerTransaction(), strCardLimit_WithdrawalPerTransaction,"Card Limit Withdrawal Limit Per Transaction")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_EligibleTransactions) Then
'      If Not verifyInnerText(lblCardLimit_EligibleTransactions(), strCardLimit_EligibleTransactions,"Card Limit Eligible Transaction")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_CreditLimit) Then
'       If Not verifyInnerText(lblCardLimit_CreditLimit(), strCardLimit_CreditLimit,"Card Limit Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_AvailableLimit) Then
'		 If Not verifyInnerText(lblCardLimit_AvailableLimit(), strCardLimit_AvailableLimit,"Card Limit Available Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_TempCreditLimit) Then
'      If Not verifyInnerText(lblCardLimit_TempCreditLimit(), strCardLimit_TempCreditLimit,"Card Limit Temporary Credit  Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_EffectiveDate) Then
'      If Not verifyInnerText(lblCardLimit_EffectiveDate(), strCardLimit_EffectiveDate,"Card Limit Effective Date")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strCardLimit_ExpiryDate) Then
'		If Not verifyInnerText(lblCardLimit_ExpiryDate(), strCardLimit_ExpiryDate,"Card Limit Expiry Date")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_CurrentBalance) Then
'       If Not verifyInnerText( lblRelationship_CurrentBalance(), strRelationship_CurrentBalance,"Relationship Current Balance")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_PendingDebits) Then
'       If Not verifyInnerText(lblRelationship_PendingDebits(), strRelationship_PendingDebits,"Relationship Pending Debits")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_PendingCredits) Then
'        If Not verifyInnerText(lblRelationship_PendingCredits(), strRelationship_PendingCredits,"Relationship Pending Credits")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'		lblRelationship_OutstandingBalance().click
'    If Not IsNull(strRelationship_OutstandingBalance) Then
'       If Not verifyInnerText(lblRelationship_OutstandingBalance(), strRelationship_OutstandingBalance,"Relationship Outstanding Balance")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_CreditLimit) Then
'       If Not verifyInnerText(lblRelationship_CreditLimit(), strRelationship_CreditLimit,"Relationship Card Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_TempCreditLimit) Then
'		Print (strRelationship_TempCreditLimit)
'       If Not verifyInnerText(lblRelationship_TempCreditLimit(), strRelationship_TempCreditLimit,"Relationship Temporary Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_AvailableCreditLimit) Then
'       If Not verifyInnerText(lblRelationship_AvailableCreditLimit(), strRelationship_AvailableCreditLimit,"Relationship Available Credit Limit")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_EffectiveDate) Then
'        If Not verifyInnerText(lblRelationship_EffectiveDate(), strRelationship_EffectiveDate,"Relationship Effective Date")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'
'    If Not IsNull(strRelationship_ExpiryDate) Then
'       If Not verifyInnerText(lblRelationship_ExpiryDate(), strRelationship_ExpiryDate,"Relationship Expiry Date")Then
'					bVerifyBalanceAndLimits = False
'		End If
'    End If
'    
'    '***************1602 changes*********************
'	If Not IsNull(strCardLimits_RTLPerDay) Then
'	        If Not verifyInnerText(lblCardLimits_RTLPerDay(), strCardLimits_RTLPerDay,"Retail Txn Limit: Per Day")Then
'			bVerifyBalanceAndLimits = False
'			End If
'	End If
'	
'	If Not IsNull(strCardLimits_RTLPerMonth) Then
'	        If Not verifyInnerText(lblCardLimits_RTLPerMonth(), strCardLimits_RTLPerMonth,"Retail Txn Limit: Per Month")Then
'			bVerifyBalanceAndLimits = False
'			End If
'	End If
'	
'	If Not IsNull(strCardLimits_RTLPerYear) Then
'	        If Not verifyInnerText(lblCardLimits_RTLPerYear(), strCardLimits_RTLPerYear,"Retail Txn Limit: Per Year")Then
'			bVerifyBalanceAndLimits = False
'			End If
'	End If
'    verifyBalanceAndLimits = bVerifyBalanceAndLimits
'
'End Function
	
