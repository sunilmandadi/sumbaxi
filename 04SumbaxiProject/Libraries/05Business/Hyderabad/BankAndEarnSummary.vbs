'[Verify and Click Bank & Earn Summary Link from leftMenu]
Public Function ClickLink_BankAndEarnSummary()
bClickLink_BankAndEarnSummary=true
	bcAccountOverview_LeftMenu.btnBankAndEarnSummary.Click
	WaitForIcallLoading
	If Err.Number<>0 Then
       bClickLink_BankAndEarnSummary=false
       LogMessage "WARN","Verification","Failed to Click Link  : Bank and Earn Summary" ,false
       Exit Function
	End If
	Wait 1
	waitForIcallLoading	
ClickLink_BankAndEarnSummary = bClickLink_BankAndEarnSummary
End Function 

'[Select value for rewards period from the combolist]
Public Function SelectRewardPeriod(strPeriod)
   bSelectRewardPeriod = true
   strPeriod = monthname(month(date),true)&" "&year(date)
   BankAndEarnSummary.txtRewardPeriod.Set strPeriod
   WaitForIcallLoading
   If Err.Number<>0 Then
       bSelectRewardPeriod = false
       Exit Function
       else
       LogMessage "RSLT","Verification","The value for Reward Period is selected as expected:"&strPeriod&"",True
    End if 
   SelectRewardPeriod = bSelectRewardPeriod
End Function

'[Verify field Credited amount displayed as]
Public Function verifyCreditedAmount(strCreditedAmount)
   bverifyCreditedAmount=true
   If Not IsNull(strCreditedAmount) Then
       If Not VerifyInnerText(BankAndEarnSummary.lblCreditedAmount(), strCreditedAmount, "Credited Amount")Then
           bverifyCreditedAmount=false
       End If
   End If
   verifyCreditedAmount=bverifyCreditedAmount
End Function

'[Verify field Credited date displayed as]
Public Function verifyCreditedDate(strCreditedDate)
	bverifyCreditedDate=true
	If Not IsNull(strCreditedAmount) Then
		If Not VerifyInnerText(BankAndEarnSummary.lblCreditedDate(), strCreditedDate, "Credited Date")Then
			bverifyCreditedDate=false
		End If
	End If
	verifyCreditedDate=bverifyCreditedDate
End Function

'[Verify Table Bank and Earn Summary Table has following Columns]
Public Function verifyBankAndEarnSummaryTableColumns(arrColumnNameList)
   	verifyBankAndEarnSummaryTableColumns=verifyTableColumns(BankAndEarnSummary.tblBankAndEarnSummaryHeader,arrColumnNameList)
End Function

'[Verify row Data for Bank and Earn Summary Table]
Public Function verifyBankAndEarnSummary_RowData(arrRowDataList)
   verifyBankAndEarnSummary_RowData=verifyTableContentList(BankAndEarnSummary.tblBankAndEarnSummaryHeader,BankAndEarnSummary.tblBankAndEarnSummaryContent,arrRowDataList,"Bank and Earn Summary" ,  false,null ,null,null)
End Function
