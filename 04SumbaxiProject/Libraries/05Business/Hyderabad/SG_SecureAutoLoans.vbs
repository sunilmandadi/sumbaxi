'*******************Added by Kalyan Auto Loans Enquiry 1702 14042017 ****************************

'[Verify row Data in Table Auto Loan Transaction History]
Public Function verifytblAutoLoan_TxnHistory(arrRowDataList)
   bDevPending=false
   verifytblAutoLoan_TxnHistory=verifyTableContentList(SecureAutoLoans.tblSecLoanTxnHistHeader,SecureAutoLoans.tblSecLoanTxnHistContent,arrRowDataList,"TxnHist_Records" ,True,SecureAutoLoans.lnkNext ,SecureAutoLoans.lnkNext1,SecureAutoLoans.lnkPrevious)
End Function

'[Verify Secured Loan Key Info Account Loan Details]
Public Function verifySLKeyInfoAccntLoanDtls(strVechNo,strLnServcngAccnt,strACcntOpnDt,strACcntMtrtyDt,strLnAmt,strLnPrd,strFrmBSts,strFormBStsDt,strScheme,strSplRmrk1,strSplRmrk2,strRmrk)
	bverifySLKeyInfoAccntLoanDtls=true
	
   If Not IsNull(strVechNo) Then
       If Not verifyInnerText(SecureAutoLoans.lblVechNo(), strVechNo, "SLKeyInfoVechNo")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strLnServcngAccnt) Then
       If Not verifyInnerText(SecureAutoLoans.lblLoanSrvngAccnt(), strLnServcngAccnt, "SLKeyInfoServcngAccnt")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strACcntOpnDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblAccntOpnDate(), strACcntOpnDt, "SLKeyInfoAccntOpenDate")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strACcntMtrtyDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblAccntMaturtyDate(), strACcntMtrtyDt, "SLKeyInfoAccntMtrtyDate")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strLnAmt) Then
       If Not verifyInnerText(SecureAutoLoans.lblLoanAmt(), strLnAmt, "SLKeyInfoLoanAmount")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strLnPrd) Then
       If Not verifyInnerText(SecureAutoLoans.lblLoanPrd(), strLnPrd, "SLKeyInfoLoanPeriod")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strFrmBSts) Then
       If Not verifyInnerText(SecureAutoLoans.lblFormBSts(), strFrmBSts, "SLKeyInfoFormBSts")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strFormBStsDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblFormBStsDt(), strFormBStsDt, "SLKeyInfoFormBStsDT")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strScheme) Then
       If Not verifyInnerText(SecureAutoLoans.lblScheme(), strScheme, "SLKeyInfoScheme")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strSplRmrk1) Then
       If Not verifyInnerText(SecureAutoLoans.lblSplRmrks1(), strSplRmrk1, "SLKeyInfoSpecialRemarks1")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strSplRmrk2) Then
       If Not verifyInnerText(SecureAutoLoans.lblSplRmrks2(), strSplRmrk2, "SLKeyInfoSpecialRemarks2")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   If Not IsNull(strRmrk) Then
       If Not verifyInnerText(SecureAutoLoans.lblSplRmrks(), strRmrk, "SLKeyInfoRemarks")Then
           bverifySLKeyInfoAccntLoanDtls=false
       End If
   End If
   
   verifySLKeyInfoAccntLoanDtls=bverifySLKeyInfoAccntLoanDtls
End Function

'[Verify Secured Loan Key Info Intslment Details]
Public Function verifySLKeyInfoInstalmntDtls(strIntrstRate,strPenlIntrstRate,strAccntCurr,strAccntTyp,strGiroAppvlDt,strGiroCmncmntDt,strNoDueInstlmnt,strMnthInstlmntDtls,strFnlInstlmntDtls,strMdOfPaymnt,strDaysPstDue,strPaymntDueDt,strNxtInstlmntDueDt)
	bverifySLKeyInfoInstalmntDtls=true
	
   If Not IsNull(strIntrstRate) Then
       If Not verifyInnerText(SecureAutoLoans.lblIntrstRate(), strIntrstRate, "SLKeyInfoInterestRate")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strPenlIntrstRate) Then
       If Not verifyInnerText(SecureAutoLoans.lblPenalIntrstRate(), strPenlIntrstRate, "SLKeyInfoPenalInterestRate")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strAccntCurr) Then
       If Not verifyInnerText(SecureAutoLoans.lblAccntCurr(), strAccntCurr, "SLKeyInfoAccntCurrency")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strAccntTyp) Then
       If Not verifyInnerText(SecureAutoLoans.lblAccntTyp(), strAccntTyp, "SLKeyInfoAccntType")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strGiroAppvlDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblGiroApprvDate(), strGiroAppvlDt, "SLKeyInfoGiroApprovalDate")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strGiroCmncmntDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblGiroCommncmntDate(), strGiroCmncmntDt, "SLKeyInfoGiroCommencementDate")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strNoDueInstlmnt) Then
       If Not verifyInnerText(SecureAutoLoans.lblNoDueInstlmnts(), strNoDueInstlmnt, "SLKeyInfoNoDueIntsalment")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strMnthInstlmntDtls) Then
       If Not verifyInnerText(SecureAutoLoans.lblMnthlyInstlmnAmt(), strMnthInstlmntDtls, "SLKeyInfoMonthlyIntslmntDtls")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strFnlInstlmntDtls) Then
       If Not verifyInnerText(SecureAutoLoans.lblFnlInstlmnAmt(), strFnlInstlmntDtls, "SLKeyInfoFinalIntsalmentDtls")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strMdOfPaymnt) Then
       If Not verifyInnerText(SecureAutoLoans.lblModeOfPaymnt(), strMdOfPaymnt, "SLKeyInfoModeOfPaymnt")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strDaysPstDue) Then
       If Not verifyInnerText(SecureAutoLoans.lblDaysPstDue(), strDaysPstDue, "SLKeyInfoDaysPstDue")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strPaymntDueDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblPaymntDueDate(), strPaymntDueDt, "SLKeyInfoPaymentDueDate")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   
   If Not IsNull(strNxtInstlmntDueDt) Then
       If Not verifyInnerText(SecureAutoLoans.lblNxtInstlmntDueDate(), strNxtInstlmntDueDt, "SLKeyInfoNxtIntsalmntDueDt")Then
           bverifySLKeyInfoInstalmntDtls=false
       End If
   End If
   verifySLKeyInfoInstalmntDtls=bverifySLKeyInfoInstalmntDtls
End Function

'[Verify row Data in Table Auto Loan Key Info Gurantor]
Public Function verifytblAutoLoan_KeyInfoGurantor(arrRowDataList)
   bDevPending=false
   verifytblAutoLoan_KeyInfoGurantor=verifyTableContentList(SecureAutoLoans.tblGurantorHeader,SecureAutoLoans.tblGurantorContent,arrRowDataList,"KeyInfoGurantor_Records",False,NULL,NULL,NULL)
End Function

'[Set TextBox From Date for Autoloan Statements]
Public Function setStmtFromDt(strFromDt)
bDevPending=true
'SecureAutoLoans.txtStmntFromDate.Set(strFromDt)
'SecureAutoLoans.btnStmntFromDate.click
selectDateFromCalendar SecureAutoLoans.btnStmntFromDate,strFromDt
If Err.Number<>0 Then
		setStmtFromDt=false
		LogMessage "WARN","Verification","Failed to Set Text Box :From" ,false
		Exit Function
End If
setStmtFromDt=true
End Function

'[Set TextBox To Date for Autoloan Statements]
Public Function setStmtToDt(strToDt)
bDevPending=true
'SecureAutoLoans.txtStmntToDate.Set(strToDt)
selectDateFromCalendar SecureAutoLoans.btnStmntToDate,strToDt
If Err.Number<>0 Then
		setStmtToDt=false
		LogMessage "WARN","Verification","Failed to Set Text Box :ToDt" ,false
		Exit Function
End If
setStmtToDt=true
End Function

'[Click button Go for Autoloan Statements]
Public Function clkStmtBtnGo()
bDevPending=true
SecureAutoLoans.lblStmntBtnGo.Click
If Err.Number<>0 Then
		clkStmtBtnGo=false
		LogMessage "WARN","Verification","Failed to click button :Go" ,false
		Exit Function
End If
clkStmtBtnGo=true
End Function

'[Click link Download for Autoloan Statements]
Public Function clkStmtLnkDownload()
bDevPending=true
SecureAutoLoans.lnkStmntDownload.Click
If Err.Number<>0 Then
		clkStmtLnkDownload=false
		LogMessage "WARN","Verification","Failed to click Link :Download" ,false
		Exit Function
End If
clkStmtLnkDownload=true
End Function

'[Verify AutoLoan Stmnt From Date Error Message displayed as]
Public Function verifyStmntFromDateError(strFromDateError)
	bverifyStmntFromDateError=true
	If Not IsNull(strFromDateError) Then
		If not VerifyInnerText(SecureAutoLoans.lblStmntDtErrMsg(), strFromDateError, "From Date Error") Then
        	bverifyStmntFromDateError=false
		End If
	End If
	verifyStmntFromDateError=bverifyStmntFromDateError
End Function

'[Verify AutoLoan Stmnt Created Date Message displayed as]
Public Function verifyStmntCrtdDateMsg(strCrtdDateMsg)
	bverifyStmntCrtdDateMsg=true
	If Not IsNull(strCrtdDateMsg) Then
		If not VerifyInnerText(SecureAutoLoans.lblStmntCrtnDt(), strCrtdDateMsg, "From Created Date Msg") Then
        	bverifyStmntCrtdDateMsg=false
		End If
	End If
	verifyStmntCrtdDateMsg=bverifyStmntCrtdDateMsg
End Function

'[Verify AutoLoan Stmnt Generated Date Message displayed as]
Public Function verifyStmntGnrtdDateMsg(strGnrtdDateMsg)
	bverifyStmntGnrtdDateMsg=true
	If Not IsNull(strGnrtdDateMsg) Then
		If not VerifyInnerText(SecureAutoLoans.lblStmntGenrtdDt(), strGnrtdDateMsg, "From Generated Date Msg") Then
        	bverifyStmntGnrtdDateMsg=false
		End If
	End If
	verifyStmntGnrtdDateMsg=bverifyStmntGnrtdDateMsg
End Function

'[Verify AutoLoan Stmnt Report Generated Message displayed as]
Public Function verifyStmntRprtGnrtdMsg(strRprtGnrtdMsg)
	bverifyStmntRprtGnrtdMsg=true
	If Not IsNull(strRprtGnrtdMsg) Then
		If not VerifyInnerText(SecureAutoLoans.lblStmntRprtGenrtdMsg(), strRprtGnrtdMsg, "Rprt Generated Msg") Then
        	bverifyStmntRprtGnrtdMsg=false
		End If
	End If
	verifyStmntRprtGnrtdMsg=bverifyStmntRprtGnrtdMsg
End Function

'[Verify AutoLoan Stmnt Download Info Message displayed as]
Public Function verifyStmntdwnldInfoMsg(strDwnldInfoMsg)
	bverifyStmntdwnldInfoMsg=true
	If Not IsNull(strDwnldInfoMsg) Then
		If not VerifyInnerText(SecureAutoLoans.lblDwnldStmntDtInfoMsg(), strDwnldInfoMsg, "stmnt Info Dwnld Msg") Then
        	bverifyStmntdwnldInfoMsg=false
		End If
	End If
	verifyStmntdwnldInfoMsg=bverifyStmntdwnldInfoMsg
End Function

'[Set TextBox proposed Payoff Date for Autoloan Redemption]
Public Function setRedmptnPrpsdPayofDt(strPrpsdPayofDt)
bDevPending=true
selectDateFromCalendar SecureAutoLoans.btnProposedPayofDt,strPrpsdPayofDt
If Err.Number<>0 Then
		setRedmptnPrpsdPayofDt=false
		LogMessage "WARN","Verification","Failed to Set Text Box :Proposed Payof Date" ,false
		Exit Function
End If
setRedmptnPrpsdPayofDt=true
End Function

'[Select Radio Button Type for Auto Loan Redemption Report]
Public Function selectRedmptnReportRadio(strRedmptnRprt)
	bstrRedmptnRprt=true
	bstrRedmptnRprt=SelectRadioButtonGrp(strRedmptnRprt, SecureAutoLoans.btnGenrtRedpmtnRprt, Array("Yes","No"))
   WaitForICallLoading
	If Err.Number<>0 Then
       bstrRedmptnRprt=false
          LogMessage "WARN","Verification","Failed to Click Button : Redemption Report" ,false
       Exit Function
   End If
   strRedmptnRprt=bstrRedmptnRprt
End Function

'[Set TextBox Payoff Fee for Autoloan Redemption]
Public Function setRedmptnPayofFee(strPayofFee)
bDevPending=true
SecureAutoLoans.txtProposedPayofFee.Set(strPayofFee)
If Err.Number<>0 Then
		setRedmptnPayofFee=false
		LogMessage "WARN","Verification","Failed to Set Text Box :Payof Fee" ,false
		Exit Function
End If
setRedmptnPayofFee=true
End Function

'[Click Compute button for Autoloan Redemption]
Public Function clickRedmptnComputeBtn()
bDevPending=true
SecureAutoLoans.btnCompute.click
If Err.Number<>0 Then
		clickRedmptnComputeBtn=false
		LogMessage "WARN","Verification","Failed to click :compute button" ,false
		Exit Function
End If
clickRedmptnComputeBtn=true
End Function

'[Verify AutoLoan Redemption orginal Loan Amount displayed as]
Public Function verifyRedmptnOrgLnAmnt(strRedmptnorgLnAmnt)
	bverifyRedmptnOrgLnAmnt=true
	If Not IsNull(strRedmptnorgLnAmnt) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnOrgAmnt(), strRedmptnorgLnAmnt, "Redmptn Org Loan Amount") Then
        	bverifyRedmptnOrgLnAmnt=false
		End If
	End If
	verifyRedmptnOrgLnAmnt=bverifyRedmptnOrgLnAmnt
End Function

'[Verify AutoLoan Redemption Outstanding Loan Balance displayed as]
Public Function verifyRedmptnOutstandLnBal(strRedmptnOutstandlnBal)
	bverifyRedmptnOutstandLnBal=true
	If Not IsNull(strRedmptnOutstandlnBal) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnOutstandAmnt(), strRedmptnOutstandlnBal, "Redmptn Outstanding Loan Balance") Then
        	bverifyRedmptnOutstandLnBal=false
		End If
	End If
	verifyRedmptnOutstandLnBal=bverifyRedmptnOutstandLnBal
End Function

'[Verify AutoLoan Redemption Interest Rebate displayed as]
Public Function verifyRedmptnIntrstRebt(strRedmptnIntrstRebt)
	bverifyRedmptnIntrstRebt=true
	If Not IsNull(strRedmptnIntrstRebt) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnIntrstRebt(), strRedmptnIntrstRebt, "Redmptn Interest Rebate") Then
        	bverifyRedmptnIntrstRebt=false
		End If
	End If
	verifyRedmptnIntrstRebt=bverifyRedmptnIntrstRebt
End Function

'[Verify AutoLoan Redemption Proposed Payoff Fee displayed as]
Public Function verifyRedmptnPropsdPayofFee(strRedmptnPayofFee)
	bverifyRedmptnPropsdPayofFee=true
	If Not IsNull(strRedmptnPayofFee) Then
		If not VerifyInnerText(SecureAutoLoans.txtProposedPayofFee(), strRedmptnPayofFee, "Redmptn Payof Fee") Then
        	bverifyRedmptnPropsdPayofFee=false
		End If
	End If
	verifyRedmptnPropsdPayofFee=bverifyRedmptnPropsdPayofFee
End Function

'[Verify AutoLoan Redemption Remarks displayed as]
Public Function verifyRedmptnRemarks(strRedmptnRemarks)
	bverifyRedmptnRemarks=true
	If Not IsNull(strRedmptnRemarks) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnRmrks(), strRedmptnRemarks, "Redmptn Remarks") Then
        	bverifyRedmptnRemarks=false
		End If
	End If
	verifyRedmptnRemarks=bverifyRedmptnRemarks
End Function

'[Verify AutoLoan Redemption Additional Prepayment displayed as]
Public Function verifyRedmptnAddnlPrePaymnt(strRedmptnPaymnt)
	bverifyRedmptnAddnlPrePaymnt=true
	If Not IsNull(strRedmptnPaymnt) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnAdtnlPaymntFees(), strRedmptnPaymnt, "Redmptn PrePayment") Then
        	bverifyRedmptnAddnlPrePaymnt=false
		End If
	End If
	verifyRedmptnAddnlPrePaymnt=bverifyRedmptnAddnlPrePaymnt
End Function

'[Verify AutoLoan Redemption Total Amount displayed as]
Public Function verifyRedmptnTtlAmnt(strRedmptnTtlAmnt)
	bverifyRedmptnTtlAmnt=true
	If Not IsNull(strRedmptnTtlAmnt) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnTtlamtFees(), strRedmptnTtlAmnt, "Redmptn TotalAmount") Then
        	bverifyRedmptnTtlAmnt=false
		End If
	End If
	verifyRedmptnTtlAmnt=bverifyRedmptnTtlAmnt
End Function

'[Verify AutoLoan Redemption Report Message displayed as]
Public Function verifyRedmptnReprtMsg(strRedmptnReprtMsg)
	bverifyRedmptnReprtMsg=true
	If Not IsNull(strRedmptnReprtMsg) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnStsMsg(), strRedmptnReprtMsg, "Redmptn Report Msg") Then
        	bverifyRedmptnReprtMsg=false
		End If
	End If
	verifyRedmptnReprtMsg=bverifyRedmptnReprtMsg
End Function

'[Verify AutoLoan Redemption PopUp Error Message displayed as]
Public Function verifyRedmptnPopUpErrMsg(strRedmptnPopUpErrMsg)
	bverifyRedmptnPopUpErrMsg=true
	If Not IsNull(strRedmptnPopUpErrMsg) Then
		If not VerifyInnerText(SecureAutoLoans.lblRedmptnPopupErrMsg(), strRedmptnPopUpErrMsg, "Redmptn PopupErr Msg") Then
        	bverifyRedmptnPopUpErrMsg=false
		End If
	End If
	verifyRedmptnPopUpErrMsg=bverifyRedmptnPopUpErrMsg
End Function

'[Click link Download for Autoloan Redemption]
Public Function clkRedmptnLnkDownload()
bDevPending=true
SecureAutoLoans.lnkRedmptnDownload.Click
If Err.Number<>0 Then
		clkRedmptnLnkDownload=false
		LogMessage "WARN","Verification","Failed to click Link :Download" ,false
		Exit Function
End If
clkRedmptnLnkDownload=true
End Function

'[Click button OK for Autoloan Redemption Popup]
Public Function clkRedmptnPopUpOkBtn()
bDevPending=true
SecureAutoLoans.btnRedmptnPopupOk.Click
If Err.Number<>0 Then
		clkRedmptnPopUpOkBtn=false
		LogMessage "WARN","Verification","Failed to click button :Ok" ,false
		Exit Function
End If
clkRedmptnPopUpOkBtn=true
End Function

'[Verify row Data in Table Auto Loan Payment Transaction]
Public Function verifytblAutoLoan_PaymntTran(arrRowDataList)
   	bDevPending=false
	WaitForICallLoading
   	'SecureAutoLoans.slObjIServe().WebElement("html tag:=MD-TAB-ITEM","innertext:=Payment Transactions").Click 
   	verifytblAutoLoan_PaymntTran=verifyTableContentList(SecureAutoLoans.tblPayFeeHistPayTranHeader,SecureAutoLoans.tblPayFeeHistPayTranContent,arrRowDataList,"PaymntTransaction_Records",True,SecureAutoLoans.lnkNext,SecureAutoLoans.lnkNext1,SecureAutoLoans.lnkPrevious)
End Function

'[Click Tab Auto Loan Payment Transaction]
Public Function clickTabAutoLoan_PaymntTran()
   	bDevPending=false
   	SecureAutoLoans.slObjIServe().WebElement("html tag:=MD-TAB-ITEM","innertext:=Payment Transactions").Click
	WaitForICallLoading
End Function

'[Click Tab Auto Loan Overdew Interest]
Public Function clickTabAutoLoan_OvrDewIntrst()
   	bDevPending=false
   	SecureAutoLoans.slObjIServe().WebElement("html tag:=MD-TAB-ITEM","innertext:=Overdue Interest").Click
	WaitForICallLoading
End Function

'[Verify row Data in Table Auto Loan Overdew Interest]
Public Function verifytblAutoLoan_OvrDewIntrst(arrRowDataList)
   bDevPending=false
   WaitForICallLoading
   'SecureAutoLoans.slObjIServe().WebElement("html tag:=MD-TAB-ITEM","innertext:=Overdue Interest").Click
   verifytblAutoLoan_OvrDewIntrst=verifyTableContentList(SecureAutoLoans.tblPayFeeHistOverDewIntrstHeader,SecureAutoLoans.tblPayFeeHistOverDewIntrstContent,arrRowDataList,"OverDewInterest_Records",True,SecureAutoLoans.lnkNext,SecureAutoLoans.lnkNext1,SecureAutoLoans.lnkPrevious)
End Function

'[Set Payment Date for Autoloan Payment History]
Public Function setPaymntHistDt(strPaymntDt)
bDevPending=true
selectDateFromCalendar SecureAutoLoans.btnPaymntDate,strPaymntDt
If Err.Number<>0 Then
		setPaymntHistDt=false
		LogMessage "WARN","Verification","Failed to Set Date :Payment History" ,false
		Exit Function
End If
setPaymntHistDt=true
End Function

'[Verify Secured Loan Payment Fee History Overdue Instalments Details]
Public Function verifySLPayFeeHistOvrDueDtls(strCurntInstlmntDue,strInstlmntToBeDue,strOutstandPnlIntrstDue,strPnlIntrstToBeDue,strOutstandFee,strFreeInstlmntNo,strFreeInstlmtAmt,strInstlmntOvrPaymnt,strTtlAmntDue)
	bverifySLPayFeeHistOvrDueDtls=true
	
   If Not IsNull(strCurntInstlmntDue) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistCrntInstlmntAmntDue(), strCurntInstlmntDue, "CurntInstlmntDue")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strInstlmntToBeDue) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistInstlmntToBeDue(), strInstlmntToBeDue, "InstlmntToBeDue")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strOutstandPnlIntrstDue) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistOutstandPenlDue(), strOutstandPnlIntrstDue, "OutstandPnlIntrstDue")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strPnlIntrstToBeDue) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistPenlIntrstDue(), strPnlIntrstToBeDue, "PnlIntrstToBeDue")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strOutstandFee) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistOutstandFee(), strOutstandFee, "OutstandFee")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strFreeInstlmntNo) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistFreeInstlmntApp(), strFreeInstlmntNo, "FreeInstlmntNo")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strFreeInstlmtAmt) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistFreeInstlmntAmnt(), strFreeInstlmtAmt, "FreeInstlmtAmt")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strInstlmntOvrPaymnt) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistInstlmntOvrPaymnt(), strInstlmntOvrPaymnt, "InstlmntOvrPaymnt")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   If Not IsNull(strTtlAmntDue) Then
       If Not verifyInnerText(SecureAutoLoans.lblPayFeeHistTtlAmntDue(), strTtlAmntDue, "TtlAmntDue")Then
           bverifySLPayFeeHistOvrDueDtls=false
       End If
   End If
   
   verifySLPayFeeHistOvrDueDtls=bverifySLPayFeeHistOvrDueDtls
End Function
