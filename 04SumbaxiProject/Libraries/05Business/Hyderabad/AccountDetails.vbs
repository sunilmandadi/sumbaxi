'*****This is auto generated code using code generator please Re-validate ****************
Dim strRunTimeAccruedOverdraftInterest:strRunTimeAccruedOverdraftInterest=""

'[Verify Account Balance section in Balances and Limits screen for CA]
Public Function verifyAccountBalance_CA(strCardNumber,strAvailableBalance,strLedgerBalance,strEarmarkAmount,strSignals)
	bverifyAccountBalance_BL=true
	Dim strProduct
	If Not IsNull(strAvailableBalance) Then
		If strAvailableBalance = "RUNTIME" Then
			getAvailableBal_CA(strCardNumber)
			'strCISAvalBal=getAvailableBal_CA(strCardNumber)
			'strCISAvalBal=strRunTimeAvailableBalance
			strCISAvalBal=strAvailableBalance
			strIServeAvalBal=BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext")			
			If  Ucase(Trim(strCISAvalBal)) = UCase(Trim(strIServeAvalBal)) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: "+ strCISAvalBal &" , Actual: "& strIServeAvalBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ strCISAvalBal &" , Actual: "& strIServeAvalBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_AvailableBalance(),strAvailableBalance,"Available Balance")Then
				bverifyAccountBalance_BL = False
			End If
	End If
    End If
    
    If Not IsNull(strLedgerBalance) Then
    	If strLedgerBalance = "RUNTIME" Then
    		'strCISLedgerBal=getLedgerBal_CA(strCardNumber)
    		'strCISLedgerBal=strRunTimeLedgerBalance
    		strCISLedgerBal=strLedgerBalance
			strIServeLedgerBal=BalancesAndLimits.lblAccountBalance_LedgerBalance.getroproperty("innertext")
			If  Ucase(Trim(strCISLedgerBal)) = UCase(Trim(strIServeLedgerBal)) Then
				LogMessage "RSLT", "Verification","For Ledger Balance successfully matched with the expected value. Expected: "+ strCISLedgerBal &" , Actual: "& strIServeLedgerBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Ledger Balance not matching with the expected value. Expected: "+ strCISLedgerBal &" , Actual: "& strIServeLedgerBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_LedgerBalance(),strLedgerBalance,"Ledger Balance")Then
				bverifyAccountBalance_BL = False
			End If
		End If	
    End If
    
    If Not IsNull(strEarmarkAmount) Then
    	If strEarmarkAmount = "RUNTIME" Then
    		'strCISEarmarkBal=getEarmarkBal_CA(strCardNumber)
    		'strCISEarmarkBal=strRunTimeEarmarkBalance
    		strCISEarmarkBal=strEarmarkAmount
    		strIServeEarmarkBal=BalancesAndLimits.lblAccountBalance_EarmarkAmount.getroproperty("innertext")
    		If  Ucase(Trim(strCISEarmarkBal)) = UCase(Trim(strIServeEarmarkBal)) Then
				LogMessage "RSLT", "Verification","For Earmark Balance successfully matched with the expected value. Expected: "+ strCISEarmarkBal &" , Actual: "& strIServeEarmarkBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Earmark Balance not matching with the expected value. Expected: "+ strCISEarmarkBal &" , Actual: "& strIServeEarmarkBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_EarmarkAmount(),strEarmarkAmount,"Earmark Amount")Then
				bverifyAccountBalance_BL = False
			End If
		End If		
    End If
    
    If Not IsNull(strSignals) Then
    	If strSignals = "RUNTIME" Then
    		'strCISSignalsBal=getSignals_CA(strCardNumber)
    		'strCISSignalsBal=strRunTimeSignalsBalance
    		strCISSignalsBal = strSignals
    		strIServeSignalsBal=BalancesAndLimits.lblAccountBalance_Signals.getroproperty("innertext")
    		If  Ucase(Trim(strCISSignalsBal)) = UCase(Trim(strIServeSignalsBal)) Then
				LogMessage "RSLT", "Verification","For Signals Balance successfully matched with the expected value. Expected: "+ strCISSignalsBal &" , Actual: "& strIServeSignalsBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Signals Balance not matching with the expected value. Expected: "+ strCISSignalsBal &" , Actual: "& strIServeSignalsBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_Signals(),strSignals,"Signals")Then
				bverifyAccountBalance_BL = False
			End If
		End If	
    End If
    verifyAccountBalance_CA=bverifyAccountBalance_BL
End Function

'[Verify Account Balance section in Balances and Limits screen for SA]
Public Function verifyAccountBalance_SA(strCardNumber,strAvailableBalance,strLedgerBalance,strEarmarkAmount,strSignals)
	bverifyAccountBalance_BL=true
	Dim strProduct
	If Not IsNull(strAvailableBalance) Then
		'If strAvailableBalance = "RUNTIME" Then
		'getAvailableBal_CA(strCardNumber)
		'strCISAvalBal=getAvailableBal_CA(strCardNumber)
		'strCISAvalBal=strRunTimeAvailableBalance
		strCISAvalBal = strAvailableBalance
			strIServeAvalBal=BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext")			
			If  Ucase(Trim(strCISAvalBal)) = UCase(Trim(strIServeAvalBal)) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: "+ strCISAvalBal &" , Actual: "& strIServeAvalBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ strCISAvalBal &" , Actual: "& strIServeAvalBal, False
				bverifyAccountBalance_BL = False
			End If
		'Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_AvailableBalance(),strAvailableBalance,"Available Balance")Then
				bverifyAccountBalance_BL = False
			End If
		'End If
    End If
    
    If Not IsNull(strLedgerBalance) Then
'    	If strLedgerBalance = "RUNTIME" Then
'    		getLedgerBalance_SA(strCardNumber)
'    		'strCISLedgerBal=getLedgerBal_CA(strCardNumber)
'    		strCISLedgerBal=strRunTimeLedgerBalance
			strCISLedgerBal = strLedgerBalance
			strIServeLedgerBal=BalancesAndLimits.lblAccountBalance_LedgerBalance.getroproperty("innertext")
			If  Ucase(Trim(strCISLedgerBal)) = UCase(Trim(strIServeLedgerBal)) Then
				LogMessage "RSLT", "Verification","For Ledger Balance successfully matched with the expected value. Expected: "+ strCISLedgerBal &" , Actual: "& strIServeLedgerBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Ledger Balance not matching with the expected value. Expected: "+ strCISLedgerBal &" , Actual: "& strIServeLedgerBal, False
				bverifyAccountBalance_BL = False
			End If
		'Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_LedgerBalance(),strLedgerBalance,"Ledger Balance")Then
				bverifyAccountBalance_BL = False
			End If
		'End If	
    End If
    
    If Not IsNull(strEarmarkAmount) Then
'    	If strEarmarkAmount = "RUNTIME" Then
'    		'strCISEarmarkBal=getEarmarkBal_CA(strCardNumber)
'    		strCISEarmarkBal=strRunTimeEarmarkBalance
			strCISEarmarkBal = strEarmarkAmount
    		strIServeEarmarkBal=BalancesAndLimits.lblAccountBalance_EarmarkAmount.getroproperty("innertext")
    		If  Ucase(Trim(strCISEarmarkBal)) = UCase(Trim(strIServeEarmarkBal)) Then
				LogMessage "RSLT", "Verification","For Earmark Balance successfully matched with the expected value. Expected: "+ strCISEarmarkBal &" , Actual: "& strIServeEarmarkBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Earmark Balance not matching with the expected value. Expected: "+ strCISEarmarkBal &" , Actual: "& strIServeEarmarkBal, False
				bverifyAccountBalance_BL = False
			End If
		'Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_EarmarkAmount(),strEarmarkAmount,"Earmark Amount")Then
				bverifyAccountBalance_BL = False
			End If
		'End If		
    End If
    
    If Not IsNull(strSignals) Then
    	'If strSignals = "RUNTIME" Then
    		'strCISSignalsBal=getSignals_CA(strCardNumber)
    		'strCISSignalsBal=strRunTimeSignalsBalance
    		strCISSignalsBal = strSignals
    		strIServeSignalsBal=BalancesAndLimits.lblAccountBalance_Signals.getroproperty("innertext")
    		If  Ucase(Trim(strCISSignalsBal)) = UCase(Trim(strIServeSignalsBal)) Then
				LogMessage "RSLT", "Verification","For Signals Balance successfully matched with the expected value. Expected: "+ strCISSignalsBal &" , Actual: "& strIServeSignalsBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Signals Balance not matching with the expected value. Expected: "+ strCISSignalsBal &" , Actual: "& strIServeSignalsBal, False
				bverifyAccountBalance_BL = False
			End If
		'Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_Signals(),strSignals,"Signals")Then
				bverifyAccountBalance_BL = False
			End If
		'End If	
    End If
    verifyAccountBalance_SA=bverifyAccountBalance_BL
End Function

'[Verify Account Balance section in Balances and Limits screen for FCCA]
Public Function verifyAccountBalance_FCCA(strCardNumber,strAvailableBalance,strLedgerBalance,strEarmarkAmount,strSignals)
	bverifyAccountBalance_BL=true
	
	If Not IsNull(strAvailableBalance) Then
		If strAvailableBalance = "RUNTIME" Then
			AvailableBalance_ANQ1(strCardNumber)
			'strCISAvalBal=AvailableBalance_ANQ1(strCardNumber)
			strCISAvalBal=strRunTimeAvailableBalance
			strIServeAvalBal=BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext")			
			If  Ucase(Trim(strCISAvalBal)) = UCase(Trim(strIServeAvalBal)) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: "+ strCISAvalBal &" , Actual: "& strIServeAvalBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ strCISAvalBal &" , Actual: "& strIServeAvalBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_AvailableBalance(),strAvailableBalance,"Available Balance")Then
				bverifyAccountBalance_BL = False
			End If
		End If
    End If
    
    If Not IsNull(strLedgerBalance) Then
    	If strLedgerBalance = "RUNTIME" Then
    		'strCISLedgerBal=strRunTimeLedgerBalance_FCCA(strCardNumber)
    		strCISLedgerBal=strRunTimeLedgerBalance
			strIServeLedgerBal=BalancesAndLimits.lblAccountBalance_LedgerBalance.getroproperty("innertext")
			If  Ucase(Trim(strCISLedgerBal)) = UCase(Trim(strIServeLedgerBal)) Then
				LogMessage "RSLT", "Verification","For Ledger Balance successfully matched with the expected value. Expected: "+ strCISLedgerBal &" , Actual: "& strIServeLedgerBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Ledger Balance not matching with the expected value. Expected: "+ strCISLedgerBal &" , Actual: "& strIServeLedgerBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_LedgerBalance(),strLedgerBalance,"Ledger Balance")Then
				bverifyAccountBalance_BL = False
			End If
		End If	
    End If
    
    If Not IsNull(strEarmarkAmount) Then
    	If strEarmarkAmount = "RUNTIME" Then
    		'strCISEarmarkBal=strRunTimeEarmarkBalance_FCCA(strCardNumber)
    		strCISEarmarkBal=strRunTimeEarmarkBalance
    		strIServeEarmarkBal=BalancesAndLimits.lblAccountBalance_EarmarkAmount.getroproperty("innertext")
    		If  Ucase(Trim(strCISEarmarkBal)) = UCase(Trim(strIServeEarmarkBal)) Then
				LogMessage "RSLT", "Verification","For Earmark Balance successfully matched with the expected value. Expected: "+ strCISEarmarkBal &" , Actual: "& strIServeEarmarkBal, True
				bverifyAccountBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Earmark Balance not matching with the expected value. Expected: "+ strCISEarmarkBal &" , Actual: "& strIServeEarmarkBal, False
				bverifyAccountBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_EarmarkAmount(),strEarmarkAmount,"Earmark Amount")Then
				bverifyAccountBalance_BL = False
			End If
		End If		
    End If
    
    If Not IsNull(strSignals) Then
		If Not verifyInnerText(BalancesAndLimits.lblAccountBalance_Signals(),strSignals,"Signals")Then
			bverifyAccountBalance_BL = False
		End If
    End If
    verifyAccountBalance_FCCA=bverifyAccountBalance_BL
End Function

'[Verify hyperlink for Earmark Amount]
Public Function verifyLinkEarmarkAmount()
	bverifyLinkEarmarkAmount=true
	EarmarkAmount=BalancesAndLimits.lblAccountBalance_EarmarkAmount.GetROProperty("innertext")
	intEarmarkAmount=Instr(BalancesAndLimits.lblAccountBalance_EarmarkAmount.GetROproperty("outerhtml"),("v-button-link"))
	If (EarmarkAmount <> 0.00)  Then
		If intEarmarkAmount = 0 Then
		   LogMessage "WARN","Verifiation","Hyperlink not available for Earmark amount as expected.",false
		   bverifyLinkEarmarkAmount=false
		End If
	Else
		If not intEarmarkAmount = 0 Then
		   LogMessage "WARN","Verifiation","Hyperlink available for Earmark amount. Expected to be unavailable.",false
		   bverifyLinkEarmarkAmount=false
		End If	
	End If
	verifyLinkEarmarkAmount=bverifyLinkEarmarkAmount
End Function

'[Verify Earmark popup]
Public Function verifyEarmarkPopup()
	bverifyEarmarkPopup=true
	BalancesAndLimits.lblEarmarkLink.click
	WaitForICallLoading
	If not BalancesAndLimits.popupEarmarkDetails.Exist Then
		bverifyEarmarkPopup=false
	End If
	verifyEarmarkPopup=bverifyEarmarkPopup
End Function

'[Verify hyperlink for Signals Amount]
Public Function verifyLinkSignalsAmount()
	bverifyLinkSignalsAmount=true
	SignalsAmount=BalancesAndLimits.lblAccountBalance_Signals.GetROProperty("innertext")
	intSignalsAmount=Instr(BalancesAndLimits.lblAccountBalance_Signals.GetROproperty("outerhtml"),("v-button-link"))
	If Not (SignalsAmount="")  Then
		If intSignalsAmount = 0 Then
		   LogMessage "WARN","Verifiation","Hyperlink not available for Signals amount as expected.",false
		   bverifyLinkSignalsAmount=false
		End If
	Else
		If not intSignalsAmount = 0 Then
		   LogMessage "WARN","Verifiation","Hyperlink available for Signals amount. Expected to be unavailable.",false
		   bverifyLinkSignalsAmount=false
		End If	
	End If
	verifyLinkSignalsAmount=bverifyLinkSignalsAmount
End Function

'[Verify Signals popup]
Public Function verifySignalsPopup()
	bverifySignalsPopup=true
	BalancesAndLimits.lblSignalsLink.click
	WaitForICallLoading
	If not BalancesAndLimits.popupSignalDetails.Exist Then
		bverifySignalsPopup=false
	End If
	verifySignalsPopup=bverifySignalsPopup
End Function

'[Verify Signal Details in popup window]
Public Function verifySingalDetails(strCardNumber,arrRowDataList)
	bverifySingalDetails=true
	'Radio button exists only for Saving button; in that case click on History radio button
	If BalancesAndLimits.rbHistory.Exist Then
		BalancesAndLimits.rbHistory.Click
		WaitForICallLoading
	End If
	If arrRowDataList = "RUNTIME" Then
		getSignalDetails_SA(strCardNumber)
		
		If strRunTimeStatus = "R" Then
			strRunTimeStatus = "RELEASED"
		End If
		If strRunTimeStatus = "PLCMT" Then
			strRunTimeStatus = "PLACED"
		End If
				
		'If strRunTimeDBSSignal="1" Then
		'	strRunTimeDBSSignal="0"&strRunTimeDBSSignal
		'End If
		
		'Change date format to bring it as per UI format
		If Not strRunTimeCISDate="" Then			
			If len(Day(CDate(strRunTimeCISDate)))=1 Then
    			strDay="0"&Day(CDate(strRunTimeCISDate))
			else
    			strDay=""&Day(CDate(strRunTimeCISDate))
			End If
			strRunTimeCISDate=""&strDay & " "&monthName(Month(CDate(strRunTimeCISDate)),true) &" " &Year(CDate(strRunTimeCISDate))
    	Else
    		strRunTimeCISDate="NULL"
    	End If    	
		
		If (strRunTimeDespatchDate="00/00/00") or (strRunTimeDespatchDate="") Then
			strRunTimeDespatchDate="NULL"
		Else 
			If len(Day(CDate(strRunTimeDespatchDate)))=1 Then
    			strDay="0"&Day(CDate(strRunTimeDespatchDate))
			else
    			strDay=""&Day(CDate(strRunTimeDespatchDate))
			End If
			strRunTimeDespatchDate=""&strDay & " "&monthName(Month(CDate(strRunTimeDespatchDate)),true) &" " &Year(CDate(strRunTimeDespatchDate))
    	End If
    	
    	If (strRunTimeFileOffDate="00/00/00") or (strRunTimeFileOffDate="") Then
			strRunTimeFileOffDate="NULL"
		Else	
			If len(Day(CDate(strRunTimeFileOffDate)))=1 Then
    			strDay="0"&Day(CDate(strRunTimeFileOffDate))
			else
    			strDay=""&Day(CDate(strRunTimeFileOffDate))
			End If
			strRunTimeFileOffDate=""&strDay & " "&monthName(Month(CDate(strRunTimeFileOffDate)),true) &" " &Year(CDate(strRunTimeFileOffDate))
    	End If
    	
    	'DBS Signal format is different in UI
    	strvalue = left(strRunTimeDBSSignal,1)
    	If strvalue = "0" Then
    		strRunTimeDBSSignal = right(strRunTimeDBSSignal,1)
    	End If
    	
    	arrRowDataList = (checknull("(Officer ID:"&strRunTimeOfficerID&"|Status:"&strRunTimeStatus&"|Date:"&strRunTimeCISDate&"|DBS Signal:"&strRunTimeDBSSignal&"|POSB Signal:"&strRunTimePOSBSignal&"|POSB RC:"&strRunTimePOSBRC&"|Despatch Date/Deleted Date:"&strRunTimeDespatchDate&"|File Off Date:"&strRunTimeFileOffDate&")|"))
		verifySingalDetails=verifyTableContentList(BalancesAndLimits.tblProductsListHeader,BalancesAndLimits.tblProductsListContent,arrRowDataList,"Earmark Details",true,BalancesAndLimits.lnkNextSignal,BalancesAndLimits.lnkNext1Signal,BalancesAndLimits.lnkPreviousSignal)
    Else	
		verifySingalDetails=verifyTableContentList(BalancesAndLimits.tblProductsListHeader,BalancesAndLimits.tblProductsListContent,arrRowDataList,"Signal Details",bPagination,BalancesAndLimits.lnkNextSignal,BalancesAndLimits.lnkNext1Signal,BalancesAndLimits.lnkPreviousSignal)
	End If
	BalancesAndLimits.btnOK_popupSignalDetails.Click
	WaitForICallLoading
End Function

'[Verify Earmark Details in popup window]
Public Function verifyEarmarkDetails(strCardNumber,arrRowDataList)
	bverifyEarmarkDetails=true
	If arrRowDataList = "RUNTIME" Then
		getEarmarkDetails_SA(strCardNumber)
		If strRunTimeStatus = "R" Then
			strRunTimeStatus = "RELEASED"
		End If
		If strRunTimeStatus = "P" Then
			strRunTimeStatus = "PLACED"
		End If
		If Not strRunTimeCISDate="          " Then			
			If len(Day(CDate(strRunTimeCISDate)))=1 Then
    			strDay="0"&Day(CDate(strRunTimeCISDate))
			else
    			strDay=""&Day(CDate(strRunTimeCISDate))
			End If
			strRunTimeCISDate=""&strDay & " "&monthName(Month(CDate(strRunTimeCISDate)),true) &" " &Year(CDate(strRunTimeCISDate))
    	End If
    	If Not strRunTimeExpiryDate="          " Then			
			strRunTimeExpiryDate="NULL"
    	End If
    	
		arrRowDataList = (checknull("(Officer ID:"&strRunTimeOfficerID&"|Status:"&strRunTimeStatus&"|Date & Time:"&strRunTimeCISDate&"|Amount:"&strRunTimeAmount&"|Branch Of Transaction:"&strRunTimeBOT&"|Expiry Date:"&strRunTimeExpiryDate&"|Reference:"&strRunTimeReferenceNumber&")|"))
		verifyEarmarkDetails=verifyTableContentList(BalancesAndLimits.tblEarmarkHeader,BalancesAndLimits.tblEarmarkContent,arrRowDataList,"Earmark Details",true,BalancesAndLimits.lnkNextEarMark,BalancesAndLimits.lnkNext1EarMark,BalancesAndLimits.lnkPreviousEarMark)
	Else
		verifyEarmarkDetails=verifyTableContentList(BalancesAndLimits.tblEarmarkHeader,BalancesAndLimits.tblEarmarkContent,arrRowDataList,"Earmark Details",true,BalancesAndLimits.lnkNextEarMark,BalancesAndLimits.lnkNext1EarMark,BalancesAndLimits.lnkPreviousEarMark)
	End If
	BalancesAndLimits.btnOK_popupEarMarkDetails.Click
	WaitForICallLoading
End Function

'[Verify Hold Balance section in Balances and Limits screen for CA]
Public Function VerifyHoldBalance_CA(strCardNumber,strHalfDay,strOneDay,strTwoDay,strLessTwoDay)
	bVerifyHoldBalance_BL=true		
	 If Not IsNull(strHalfDay) Then
    	If strHalfDay = "RUNTIME" Then
    		'getHalfDayAmount_CA(strCardNumber)
    		'strCISHalfDayAmount=getHalfDayAmount_CA(strCardNumber)
    		strCISHalfDayAmount=strRunTimeHalfDayAmount
    		strIServeHalfDayAmount=BalancesAndLimits.lblHoldBalance_HalfDay.getroproperty("innertext")
    		If  Ucase(Trim(strCISHalfDayAmount)) = UCase(Trim(strIServeHalfDayAmount)) Then
				LogMessage "RSLT", "Verification","For Half Day Amount successfully matched with the expected value. Expected: "+ strCISHalfDayAmount &" , Actual: "& strIServeHalfDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Half Day Amount not matching with the expected value. Expected: "+ strCISHalfDayAmount &" , Actual: "& strIServeHalfDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_HalfDay(),strHalfDay,"Half Day")Then
				bVerifyHoldBalance_BL = False
			End If
		End If		
    End If
    
    If Not IsNull(strOneDay) Then
    	If strOneDay = "RUNTIME" Then
    		'strCISOneDayAmount=getOneDayAmount_CA(strCardNumber)
    		strCISOneDayAmount=strRunTimeOneDayAmount
    		strIServeOneDayAmount=BalancesAndLimits.lblHoldBalance_OneDay.getroproperty("innertext")
    		If  Ucase(Trim(strCISOneDayAmount)) = UCase(Trim(strIServeOneDayAmount)) Then
    			LogMessage "RSLT", "Verification","For One Day Amount successfully matched with the expected value. Expected: "+ strCISOneDayAmount &" , Actual: "& strIServeOneDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For One Day Amount not matching with the expected value. Expected: "+ strCISOneDayAmount &" , Actual: "& strIServeOneDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_OneDay(),strOneDay,"One Day")Then
				bVerifyHoldBalance_BL = False
			End If
		End If	
    End If
	
	If Not IsNull(strTwoDay) Then
		If strTwoDay = "RUNTIME" Then
			'strCISTwoDayAmount=getTwoDayAmount_CA(strCardNumber)
			strCISTwoDayAmount=strRunTimeTwoDayAmount
    		strIServeTwoDayAmount=BalancesAndLimits.lblHoldBalance_TwoDays.getroproperty("innertext")
    		If  Ucase(Trim(strCISTwoDayAmount)) = UCase(Trim(strIServeTwoDayAmount)) Then
    			LogMessage "RSLT", "Verification","For Two Day Amount successfully matched with the expected value. Expected: "+ strCISTwoDayAmount &" , Actual: "& strIServeTwoDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Two Day Amount not matching with the expected value. Expected: "+ strCISTwoDayAmount &" , Actual: "& strIServeTwoDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else					
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_TwoDays(),strTwoDay,"Two Days")Then
				bVerifyHoldBalance_BL = False
			End If
		End If
    End If
    
    If Not IsNull(strLessTwoDay) Then
    	If strLessTwoDay = "RUNTIME" Then
    		'strCISLessTwoDayAmount=LessTwoDayAmount_ANQ1(strCardNumber)
    		strCISLessTwoDayAmount=strRunTimeLessTwoDayAmount
    		strIServeLessTwoDayAmount=BalancesAndLimits.lblHoldBalance_LessThanTwoDays.getroproperty("innertext")
    		If  Ucase(Trim(strCISLessTwoDayAmount)) = UCase(Trim(strIServeLessTwoDayAmount)) Then
    			LogMessage "RSLT", "Verification","For Two Day Amount successfully matched with the expected value. Expected: "+ strCISLessTwoDayAmount &" , Actual: "& strIServeLessTwoDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Two Day Amount not matching with the expected value. Expected: "+ strCISLessTwoDayAmount &" , Actual: "& strIServeLessTwoDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_LessThanTwoDays(),strLessTwoDay,"Less Two Days")Then
				bVerifyHoldBalance_BL = False
			End If
		End If	
    End If	
	VerifyHoldBalance_CA=bVerifyHoldBalance_BL
End Function

'[Verify Hold Balance section in Balances and Limits screen for FCCA]
Public Function VerifyHoldBalance_BL(strCardNumber,strHalfDay,strOneDay,strTwoDay,strLessTwoDay)
	bVerifyHoldBalance_BL=true
		
	 If Not IsNull(strHalfDay) Then
    	If strHalfDay = "RUNTIME" Then
    		'HalfDayAmount_ANQ1(strCardNumber)
    		'strCISHalfDayAmount=HalfDayAmount_ANQ1(strCardNumber)
    		strCISHalfDayAmount=strRunTimeHalfDayAmount
    		strIServeHalfDayAmount=BalancesAndLimits.lblHoldBalance_HalfDay.getroproperty("innertext")
    		If  Ucase(Trim(strCISHalfDayAmount)) = UCase(Trim(strIServeHalfDayAmount)) Then
				LogMessage "RSLT", "Verification","For Half Day Amount successfully matched with the expected value. Expected: "+ strCISHalfDayAmount &" , Actual: "& strIServeHalfDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Half Day Amount not matching with the expected value. Expected: "+ strCISHalfDayAmount &" , Actual: "& strIServeHalfDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_HalfDay(),strHalfDay,"Half Day")Then
				bVerifyHoldBalance_BL = False
			End If
		End If		
    End If
    
    If Not IsNull(strOneDay) Then
    	If strOneDay = "RUNTIME" Then
    		'strCISOneDayAmount=OneDayAmount_ANQ1(strCardNumber)
    		strCISOneDayAmount=strRunTimeOneDayAmount
    		strIServeOneDayAmount=BalancesAndLimits.lblHoldBalance_OneDay.getroproperty("innertext")
    		If  Ucase(Trim(strCISOneDayAmount)) = UCase(Trim(strIServeOneDayAmount)) Then
    			LogMessage "RSLT", "Verification","For One Day Amount successfully matched with the expected value. Expected: "+ strCISOneDayAmount &" , Actual: "& strIServeOneDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For One Day Amount not matching with the expected value. Expected: "+ strCISOneDayAmount &" , Actual: "& strIServeOneDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_OneDay(),strOneDay,"One Day")Then
				bVerifyHoldBalance_BL = False
			End If
		End If	
    End If
	
	If Not IsNull(strTwoDay) Then
		If strTwoDay = "RUNTIME" Then
			'strCISTwoDayAmount=TwoDayAmount_ANQ1(strCardNumber)
			strCISTwoDayAmount=strRunTimeTwoDayAmount
    		strIServeTwoDayAmount=BalancesAndLimits.lblHoldBalance_TwoDays.getroproperty("innertext")
    		If  Ucase(Trim(strCISTwoDayAmount)) = UCase(Trim(strIServeTwoDayAmount)) Then
    			LogMessage "RSLT", "Verification","For Two Day Amount successfully matched with the expected value. Expected: "+ strCISTwoDayAmount &" , Actual: "& strIServeTwoDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Two Day Amount not matching with the expected value. Expected: "+ strCISTwoDayAmount &" , Actual: "& strIServeTwoDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else					
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_TwoDays(),strTwoDay,"Two Days")Then
				bVerifyHoldBalance_BL = False
			End If
		End If
    End If
    
    If Not IsNull(strLessTwoDay) Then
    	If strLessTwoDay = "RUNTIME" Then
    		'strCISLessTwoDayAmount=LessTwoDayAmount_ANQ1(strCardNumber)
    		strCISLessTwoDayAmount=strRunTimeLessTwoDayAmount
    		strIServeLessTwoDayAmount=BalancesAndLimits.lblHoldBalance_LessThanTwoDays.getroproperty("innertext")
    		If  Ucase(Trim(strCISLessTwoDayAmount)) = UCase(Trim(strIServeLessTwoDayAmount)) Then
    			LogMessage "RSLT", "Verification","For Two Day Amount successfully matched with the expected value. Expected: "+ strCISLessTwoDayAmount &" , Actual: "& strIServeLessTwoDayAmount, True
				bVerifyHoldBalance_BL = True
			else
				LogMessage "WARN", "Verification","For Two Day Amount not matching with the expected value. Expected: "+ strCISLessTwoDayAmount &" , Actual: "& strIServeLessTwoDayAmount, False
				bVerifyHoldBalance_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblHoldBalance_LessThanTwoDays(),strLessTwoDay,"Less Two Days")Then
				bVerifyHoldBalance_BL = False
			End If
		End If	
    End If	
	VerifyHoldBalance_BL=bVerifyHoldBalance_BL
End Function

'[Verify Returned Cheque section is Balances and Limits screen]
Public Function verifyReturnedCheque_BL(strCardNumber,strCurrentMonth,strLastMonth,strLast2Months)
	bverifyReturnedCheque_BL=true	
	If Not IsNull(strCurrentMonth) Then
		If strCurrentMonth = "RUNTIME" Then
			'CurrentMonth_ANQ1(strCardNumber)
    		'strCISCurrentMonth=CurrentMonth_ANQ1(strCardNumber)
    		strCISCurrentMonth=strRunTimeCurrentMonth
    		strIServeCurrentMonth=BalancesAndLimits.lblReturnedCheque_CurrentMonth.getroproperty("innertext")
    		If  Ucase(Trim(strCISCurrentMonth)) = UCase(Trim(strIServeCurrentMonth)) Then
    			LogMessage "RSLT", "Verification","For Current Month value successfully matched with the expected value. Expected: "+ strCISCurrentMonth &" , Actual: "& strIServeCurrentMonth, True
				bverifyReturnedCheque_BL = True
			else
				LogMessage "WARN", "Verification","For Current Month value not matching with the expected value. Expected: "+ strCISCurrentMonth &" , Actual: "& strIServeCurrentMonth, False
				bverifyReturnedCheque_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblReturnedCheque_CurrentMonth(),strCurrentMonth,"Current Month")Then
				bverifyReturnedCheque_BL = False
			End If
    	End If
    End If	
    
    If Not IsNull(strLastMonth) Then
    	If strLastMonth = "RUNTIME" Then
    		'strCISLastMonth=LastMonth_ANQ1(strCardNumber)
    		strCISLastMonth=strRunTimeLastMonth
    		strIServeLastMonth=BalancesAndLimits.lblReturnedCheque_LastMonth.getroproperty("innertext")
    		If  Ucase(Trim(strCISLastMonth)) = UCase(Trim(strIServeLastMonth)) Then
    			LogMessage "RSLT", "Verification","For Last Month value successfully matched with the expected value. Expected: "+ strCISLastMonth &" , Actual: "& strIServeLastMonth, True
				bverifyReturnedCheque_BL = True
			else
				LogMessage "WARN", "Verification","For Last Month value not matching with the expected value. Expected: "+ strCISLastMonth &" , Actual: "& strIServeLastMonth, False
				bverifyReturnedCheque_BL = False
			End If
		Else
			If Not verifyInnerText(BalancesAndLimits.lblReturnedCheque_LastMonth(),strLastMonth,"Last Month")Then
				bverifyReturnedCheque_BL = False
			End If
    	End If
    End If	
    
    If Not IsNull(strLast2Months) Then
    	If strLast2Months = "RUNTIME" Then
    		'strCISLast2Month=Last2Month_ANQ1(strCardNumber)
    		strCISLast2Month=strRunTimeLastTwoMonth
    		strIServeLast2Month=BalancesAndLimits.lblReturnedCheque_Last2Months.getroproperty("innertext")
    		If  Ucase(Trim(strCISLast2Month)) = UCase(Trim(strIServeLast2Month)) Then
    			LogMessage "RSLT", "Verification","For Last Two Month value successfully matched with the expected value. Expected: "+ strCISLast2Month &" , Actual: "& strIServeLast2Month, True
				bverifyReturnedCheque_BL = True
			else
				LogMessage "WARN", "Verification","For Last Two Month value not matching with the expected value. Expected: "+ strCISLast2Month &" , Actual: "& strIServeLast2Month, False
				bverifyReturnedCheque_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblReturnedCheque_Last2Months(),strLast2Months,"Last 2 Month")Then
				bverifyReturnedCheque_BL = False
			End If
		End If	
    End If    
    verifyReturnedCheque_BL=bverifyReturnedCheque_BL
End Function

'[Verify Limits section in Balances and Limits screen for CA]
Public Function verifyLimits_CA(strOverdraftLimit,strAccruedOverdraft)
	bverifyLimits_BL=true
	If Not IsNull(strOverdraftLimit) Then
		If strOverdraftLimit = "RUNTIME" Then
			getOverdraftLimit_CA(strCardNumber)
			'strCISOverDraftLimit=getOverdraftLimit_CA(strCardNumber)
			strCISOverDraftLimit=strRunTimeOverdraftLimit
    		strIServeOverDraftLimit=BalancesAndLimits.lblLimits_OverdraftLimit.getroproperty("innertext")
    		If  Ucase(Trim(strCISOverDraftLimit)) = UCase(Trim(strIServeOverDraftLimit)) Then
    			LogMessage "RSLT", "Verification","For Overdraft Limit value successfully matched with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, True
				bverifyLimits_BL = True
			else
				LogMessage "WARN", "Verification","For Overdraft Limit value not matching with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, False
				bverifyLimits_BL = False
			End If
		Else
			If Not verifyInnerText(BalancesAndLimits.lblLimits_OverdraftLimit(),strOverdraftLimit,"Overdraft Limit")Then
				bverifyLimits_BL = False
			End If
		End If
    End If
    
    If Not IsNull(strAccruedOverdraft) Then
    	If strAccruedOverdraft = "RUNTIME" Then
    		strCISOverDraftLimit=strRunTimeAccruedOverdraft
    		strIServeOverDraftLimit=BalancesAndLimits.lblLimits_AccruedOverdraft.getroproperty("innertext")
    		If  Ucase(Trim(strCISOverDraftLimit)) = UCase(Trim(strIServeOverDraftLimit)) Then
    			LogMessage "RSLT", "Verification","For Accrued Overdraft Limit value successfully matched with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, True
				bverifyLimits_BL = True
			else
				LogMessage "WARN", "Verification","For Accrued Overdraft Limit value not matching with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, False
				bverifyLimits_BL = False
			End If
		Else
			If Not verifyInnerText(BalancesAndLimits.lblLimits_AccruedOverdraft(),strAccruedOverdraft,"Accrued Overdraft")Then
				bverifyLimits_BL = False
			End If
		End If	
    End If
    verifyLimits_CA=bverifyLimits_BL
End Function

'[Verify Limits section in Balances and Limits screen for FCCA]
Public Function verifyLimits_BL(strOverdraftLimit,strAccruedOverdraft)
	bverifyLimits_BL=true
	If Not IsNull(strOverdraftLimit) Then
		If strOverdraftLimit = "RUNTIME" Then
			strCISOverDraftLimit=strRunTimeOverdraftLimit
    		strIServeOverDraftLimit=BalancesAndLimits.lblLimits_OverdraftLimit.getroproperty("innertext")
    		If  Ucase(Trim(strCISOverDraftLimit)) = UCase(Trim(strIServeOverDraftLimit)) Then
    			LogMessage "RSLT", "Verification","For Overdraft Limit value successfully matched with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, True
				bverifyLimits_BL = True
			else
				LogMessage "WARN", "Verification","For Overdraft Limit value not matching with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, False
				bverifyLimits_BL = False
			End If
		Else	
			If Not verifyInnerText(BalancesAndLimits.lblLimits_OverdraftLimit(),strOverdraftLimit,"Overdraft Limit")Then
				bverifyLimits_BL = False
			End If
		End If	
    End If
    
    If Not IsNull(strAccruedOverdraft) Then
    	If strAccruedOverdraft = "RUNTIME" Then
    		strCISOverDraftLimit=strRunTimeAccruedOverdraft
    		strIServeOverDraftLimit=BalancesAndLimits.lblLimits_OverdraftLimit.getroproperty("innertext")
    		If  Ucase(Trim(strCISOverDraftLimit)) = UCase(Trim(strIServeOverDraftLimit)) Then
    			LogMessage "RSLT", "Verification","For Accrued Overdraft Limit value successfully matched with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, True
				bverifyLimits_BL = True
			else
				LogMessage "WARN", "Verification","For Accrued Overdraft Limit value not matching with the expected value. Expected: "+ strCISOverDraftLimit &" , Actual: "& strIServeOverDraftLimit, False
				bverifyLimits_BL = False
			End If
		Else
			If Not verifyInnerText(BalancesAndLimits.lblLimits_AccruedOverdraft(),strAccruedOverdraft,"Accrued Overdraft")Then
				bverifyLimits_BL = False
			End If
		End If	
    End If
    verifyLimits_BL=bverifyLimits_BL
End Function

'[Verify hyperlink for Accrued Overdraft Amount]
Public Function verifyLinkAccruedOverdraftAmount()
	bverifyLinkAccruedOverdraftAmount=true
	AccruedOverdraft=BalancesAndLimits.lblLimits_AccruedOverdraft.GetROProperty("innertext")
	intAccruedOverdraftAmount=Instr(BalancesAndLimits.lblLimits_AccruedOverdraft.GetROproperty("outerhtml"),("v-slot-link"))
	If (AccruedOverdraft <> 0.00)  Then
		If intAccruedOverdraftAmount = 0 Then
		   LogMessage "WARN","Verifiation","Hyperlink not available for Accrued Overdraft amount as expected.",false
		   bverifyLinkAccruedOverdraftAmount=false
		End If
	Else
		If not intAccruedOverdraftAmount = 0 Then
		   LogMessage "WARN","Verifiation","Hyperlink available for Accrued Overdraft amount. Expected to be unavailable.",false
		   bverifyLinkAccruedOverdraftAmount=false
		End If	
	End If
	verifyLinkAccruedOverdraftAmount=bverifyLinkAccruedOverdraftAmount
End Function

'[Verify if Action Icon exist]
Public Function verifyActionIconExist()
	bverifyActionIconExist=true
	If not (BalancesAndLimits.lblLimits_ActionIcon().Exist) Then
		LogMessage "WARN","Verification","Action Icon is not available in Balance And Limit Page", False
		bverifyActionIconExist=False		
	End If
	verifyActionIconExist=bverifyActionIconExist
End Function

'[Select SubMenu from Action Icon]
Public Function selectSubMenu_BL(strItem)
	bselectSubMenu_BL=true
	WaitForICallLoading
	strRunTimeAccruedOverdraftInterest=BalancesAndLimits.lblLimits_AccruedOverdraft.getroproperty("innertext")
	Set oDesc=Description.Create
	oDesc("micclass").Value = "WebElement"	
	'oDesc("class").Value = "md-no-padding"
	oDesc("class").Value = "popupForm"
	BalancesAndLimits.lblLimits_ActionIcon.click
	set lstSubMenu=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDesc)
	intItems=lstSubMenu.Count	
	'intItems=BalancesAndLimits.lblLimits_ActionIcon.Count
	
	For iCount=0 to intItems
		Dim strTemp:strTemp=""
		strTemp=lstSubMenu(iCount).GetRoProperty("text")
		If strTemp=strItem Then
			'Check the link enabled or disabled
			bDisabled =InStr(lstSubMenu(iCount).GetROProperty("class"),"disabled-area")
    		If bDisabled <>0Then
				LogMessage "WARN","Verification","Sub Menu Icon is disabled in Row Number ",True
				bMenuDisabled=True
				selectSubMenu_BL=true
			Exit Function
			Else
				bMenuDisabled=False
				LogMessage "INFO","Verification","Sub Menu  is enabled in Row Number"&intRow,True
				Set oMenuItem=Description.Create	
				oMenuItem("micclass").Value = "WebElement"
	  			'oMenuItem("class").Value = ".*flex"
	  			oMenuItem("class").Value = "md-button.*"
	  			set lstSubMenu=lstSubMenu(iCount).ChildObjects(oMenuItem)
				lstSubMenu(iCount).click
				strTemp=lstSubMenu(iCount).GetRoProperty("text")
				If Not IsNUll(strTemp) Then
				   lstSubMenu(iCount).click
				End If
				WaitForICallLoading
				LogMessage "RSLT","Verification","Item "&strItem&" selected from Submenu sucessfully. Item Index is "& intItemIndex,true
				selectSubMenu_BL=true
			Exit Function
			End If
		End If
		intItemIndex=intItemIndex+1
	Next
	selectSubMenu_BL=bselectSubMenu_BL	
End Function

'[Verify data in Account Holder table in Key Info Page]
Public Function verifyAccountHolderTable(strProductCode,strCardNumber,lstlstAccountHolder)
	verifyAccountHolderTable=true
	If Not IsNull (lstlstAccountHolder) Then
'		'If lstlstAccountHolder="RUNTIME" Then
'			'getAccountInfo_CA strProductCode,strCardNumber
'			'strRunTimeCIN=trim(strRunTimeCIN)
'			'strCIN=strRunTimeCIN &" 00"
'			'strRunTimeAccountHolderName=trim(strRunTimeAccountHolderName)
'			lstlstAccountHolder=CheckNull("(Account Holder (s):"&strRunTimeAccountHolderName&"|CIN/CIN Suffix:"&strCIN&")|")
'		End If	
		verifyAccountHolderTable=verifyTableContentList(bcKeyInfo.tblAccountHolderHeader,bcKeyInfo.tblAccountHolderContent,lstlstAccountHolder,"Account Holder",false,null,null,null)		
	End if 
End Function

'[Click on Relationship Column to Validate Relationship Details]
Public Function clickRelationshipColumn_KeyInfo(arrRowDataList)
	bclickRelationshipColumn_KeyInfo=true
	If Not IsNull (arrRowDataList) Then	
		clickRelationshipColumn_KeyInfo=selectTableLink(bcKeyInfo.tblAccountHolderHeader,bcKeyInfo.tblAccountHolderContent,arrRowDataList,"KeyInfo table" ,"Relationship",false,null,null,null)
		WaitForICallLoading
		If not bcKeyInfo.popupRelatedCustomerDetails.Exist Then
			LogMessage "WARN","Verification","Related customer details are not displayed successfully." ,False
			bClickStatementOption=false
		End If
	End If
	clickRelationshipColumn_KeyInfo=bclickRelationshipColumn_KeyInfo
End Function

'[Click on CIN Suffix Number in Account Details table]
Public Function clickColumnCINSuffix_KeyInfo(arrRowDataList)
WaitForICallLoading
If Not IsNull (arrRowDataList) Then	
	bclickRelationshipColumn_KeyInfo=selectTableLink(bcKeyInfo.tblAccountHolderHeader,bcKeyInfo.tblAccountHolderContent,arrRowDataList,"KeyInfo table" ,"CIN/CIN Suffix",false,null,null,null)
End IF 
clickColumnCINSuffix_KeyInfo=bclickRelationshipColumn_KeyInfo
End Function

'[Verify Relationship Details in Key Info Page]
Public Function verifyrelationshipDetails_KeyInfo(arrRowDataList)
	verifyrelationshipDetails_KeyInfo=true
	WaitForICallLoading
	If Not IsNull (arrRowDataList) Then
		verifyrelationshipDetails_KeyInfo=verifyTableContentList(bcKeyInfo.tblRelationshipDetailHeader,bcKeyInfo.tblRelationshipDetailContent,arrRowDataList,"Relationship Table",false,null,null,null)
	End If
	bcKeyInfo.btnOK_popupRelatedCustomerDetails.Click
	WaitForICallLoading
End Function

'[Verify Account Information on Key Info Page for CA displayed as]
Public Function verifyAccountInformation_CA(strAccountShortName,strAccountSignatoryType,strAccountType,strBrandIndicator,strPrimaryCIN,strStatus,strFeeIndicator,strOpeningDate,strClosingDate)
	bverifyAccountInformation=true
	If Not IsNull(strAccountShortName) Then
		If strAccountShortName = "RUNTIME" Then
			strCISShortName=strRunTimeShortName
			strIServeShortName=bcKeyInfo.lblAccountShortName.getroproperty("innertext")			
			If  Ucase(Trim(strCISShortName)) = UCase(Trim(strIServeShortName)) Then
				LogMessage "RSLT", "Verification","For Account Short Name successfully matched with the expected value. Expected: "+ strCISShortName &" , Actual: "& strIServeShortName, True
				bverifyAccountInformation = True
			else
				LogMessage "WARN", "Verification","For Account Short Name not matching with the expected value. Expected: "+ strCISShortName &" , Actual: "& strIServeShortName, False
				bverifyAccountInformation = False
			End If
		Else	
			If Not verifyInnerText(bcKeyInfo.lblAccountShortName(),strAccountShortName,"Account Short Name")Then
				bverifyAccountInformation = False
			End If
		End If	
    End If
    
    If Not IsNull(strAccountSignatoryType) Then
		If Not verifyInnerText(bcKeyInfo.lblAccountSignatoryType(),strAccountSignatoryType,"Account Signatory Type")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strAccountType) Then    	
		If Not verifyInnerText(bcKeyInfo.lblAccountType(),strAccountType,"Account Type")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strBrandIndicator) Then
		If Not verifyInnerText(bcKeyInfo.lblBrandIndicator(),strBrandIndicator,"Brand Indicator")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strPrimaryCIN) Then
		If Not verifyInnerText(bcKeyInfo.lblPrimaryCIN(),strPrimaryCIN,"Primary CIN")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strStatus) Then
		If Not verifyInnerText(bcKeyInfo.lblStatus(),strStatus,"Status")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strFeeIndicator) Then
		If Not verifyInnerText(bcKeyInfo.lblFeeIndicator(),strFeeIndicator,"Fee Indicator")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strOpeningDate) Then
    	If strOpeningDate = "RUNTIME" Then
			strDate=strRunTimeOpeningDate
			If len(Day(CDate(strDate)))=1 Then
        		strDay="0"&Day(CDate(strDate))
    		else
        		strDay=""&Day(CDate(strDate))
    		End If
    		strCISOpeningDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
			strIServeOpeningDate=bcKeyInfo.lblOpeningDate.getroproperty("innertext")			
			If  Ucase(Trim(strCISOpeningDate)) = UCase(Trim(strIServeOpeningDate)) Then
				LogMessage "RSLT", "Verification","For Opening Date successfully matched with the expected value. Expected: "+ strCISOpeningDate &" , Actual: "& strIServeOpeningDate, True
				bverifyAccountInformation = True
			else
				LogMessage "WARN", "Verification","For Opening Date not matching with the expected value. Expected: "+ strCISOpeningDate &" , Actual: "& strIServeOpeningDate, False
				bverifyAccountInformation = False
			End If
		Else
			If Not verifyInnerText(bcKeyInfo.lblOpeningDate(),strOpeningDate,"Opening Date")Then
				bverifyAccountInformation = False
			End If
		End If	
    End If
    If Not IsNull(strClosingDate) Then
    	If strClosingDate = "RUNTIME" Then
			strDate=strRunTimeClosingDate
			If Not strDate="          " Then			
				If len(Day(CDate(strDate)))=1 Then
        			strDay="0"&Day(CDate(strDate))
    			else
        			strDay=""&Day(CDate(strDate))
    			End If
    			strCISClosingDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
				strIServeClosingDate=bcKeyInfo.lblClosingDate.getroproperty("innertext")			
				If  Ucase(Trim(strCISClosingDate)) = UCase(Trim(strIServeClosingDate)) Then
					LogMessage "RSLT", "Verification","For Closing Date successfully matched with the expected value. Expected: "+ strCISClosingDate &" , Actual: "& strIServeClosingDate, True
					bverifyAccountInformation = True			
				else
					LogMessage "WARN", "Verification","For Closing Date not matching with the expected value. Expected: "+ strCISClosingDate &" , Actual: "& strIServeClosingDate, False
					bverifyAccountInformation = False
				End If
			ElseIf Not verifyInnerText(bcKeyInfo.lblClosingDate(),strDate,"Closing Date")Then
				bverifyAccountInformation = False
			End If
		Else	
			If Not verifyInnerText(bcKeyInfo.lblClosingDate(),strClosingDate,"Closing Date")Then
				bverifyAccountInformation = False
			End If			
		End If	
    End If
    verifyAccountInformation_CA=bverifyAccountInformation
End Function

'[Verify Account Information on Key Info Page for SA displayed as]
Public Function verifyAccountInformation_SA(strCardNumber,strAccountShortName,strAccountSignatoryType,strAccountType,strBrandIndicator,strPrimaryCIN,strStatus,strFeeIndicator,strOpeningDate,strClosingDate)
	bverifyAccountInformation=true
	If Not IsNull(strAccountShortName) Then
		If strAccountShortName = "RUNTIME" Then
			strCISShortName=strRunTimeShortName
			strIServeShortName=bcKeyInfo.lblAccountShortName.getroproperty("innertext")			
			If  Ucase(Trim(strCISShortName)) = UCase(Trim(strIServeShortName)) Then
				LogMessage "RSLT", "Verification","For Account Short Name successfully matched with the expected value. Expected: "+ strCISShortName &" , Actual: "& strIServeShortName, True
				bverifyAccountInformation = True
			else
				LogMessage "WARN", "Verification","For Account Short Name not matching with the expected value. Expected: "+ strCISShortName &" , Actual: "& strIServeShortName, False
				bverifyAccountInformation = False
			End If
		Else	
			If Not verifyInnerText(bcKeyInfo.lblAccountShortName(),strAccountShortName,"Account Short Name")Then
				bverifyAccountInformation = False
			End If
		End If	
    End If
    
    If Not IsNull(strAccountSignatoryType) Then
		If Not verifyInnerText(bcKeyInfo.lblAccountSignatoryType(),strAccountSignatoryType,"Account Signatory Type")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strAccountType) Then
    	If strAccountType="RUNTIME" Then
    		getAccountType_SA(strCardNumber)
    		If strRunTimeAccountType = "1" Then
    			strCISAccountType="01 - Other Private Individual & Households"
    		End If
    		
    		strIServeAccountType=bcKeyInfo.lblAccountType.getroproperty("innertext")
    		If  Ucase(Trim(strCISAccountType)) = UCase(Trim(strIServeAccountType)) Then
    			LogMessage "RSLT", "Verification","Account Type successfully matched with the expected value. Expected: "+ strCISAccountType &" , Actual: "& strIServeAccountType, True
				bverifyAccountInformation = True
			else
				LogMessage "WARN", "Verification","Account Type not matching with the expected value. Expected: "+ strCISAccountType &" , Actual: "& strIServeAccountType, False
				bverifyAccountInformation = False
			End If
    	Else
			If Not verifyInnerText(bcKeyInfo.lblAccountType(),strAccountType,"Account Type")Then
				bverifyAccountInformation = False
			End If    	
    	End If
		
    End If
    If Not IsNull(strBrandIndicator) Then
		If Not verifyInnerText(bcKeyInfo.lblBrandIndicator(),strBrandIndicator,"Brand Indicator")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strPrimaryCIN) Then
		If Not verifyInnerText(bcKeyInfo.lblPrimaryCIN(),strPrimaryCIN,"Primary CIN")Then
			bverifyAccountInformation = False
		End If
    End If
    
    If Not IsNull(strStatus) Then
    	If strStatus = "RUNTIME" Then
    		If strRunTimeAccountStatus = "1" Then
    			strCISAccountStatus="01 - Active"
    		ElseIf strRunTimeAccountStatus = "2" Then
    			strCISAccountStatus="02 - Inactive"
    		End If
    		strIServeAccountStatus=bcKeyInfo.lblStatus.getroproperty("innertext")
    		If  Ucase(Trim(strCISAccountStatus)) = UCase(Trim(strIServeAccountStatus)) Then
    			LogMessage "RSLT", "Verification","Account Status successfully matched with the expected value. Expected: "+ strCISAccountStatus &" , Actual: "& strIServeAccountStatus, True
				bverifyAccountInformation = True
			else
				LogMessage "WARN", "Verification","Account Status not matching with the expected value. Expected: "+ strCISAccountStatus &" , Actual: "& strIServeAccountStatus, False
				bverifyAccountInformation = False
			End If
		Else
			If Not verifyInnerText(bcKeyInfo.lblStatus(),strStatus,"Status")Then
				bverifyAccountInformation = False
			End If
		End If	
    End If
    
    If Not IsNull(strFeeIndicator) Then
		If Not verifyInnerText(bcKeyInfo.lblFeeIndicator(),strFeeIndicator,"Fee Indicator")Then
			bverifyAccountInformation = False
		End If
    End If
    If Not IsNull(strOpeningDate) Then
    	If strOpeningDate = "RUNTIME" Then
			strDate=strRunTimeOpeningDate
			If len(Day(CDate(strDate)))=1 Then
        		strDay="0"&Day(CDate(strDate))
    		else
        		strDay=""&Day(CDate(strDate))
    		End If
    		strCISOpeningDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
			strIServeOpeningDate=bcKeyInfo.lblOpeningDate.getroproperty("innertext")			
			If  Ucase(Trim(strCISOpeningDate)) = UCase(Trim(strIServeOpeningDate)) Then
				LogMessage "RSLT", "Verification","For Opening Date successfully matched with the expected value. Expected: "+ strCISOpeningDate &" , Actual: "& strIServeOpeningDate, True
				bverifyAccountInformation = True
			else
				LogMessage "WARN", "Verification","For Opening Date not matching with the expected value. Expected: "+ strCISOpeningDate &" , Actual: "& strIServeOpeningDate, False
				bverifyAccountInformation = False
			End If
		Else
			If Not verifyInnerText(bcKeyInfo.lblOpeningDate(),strOpeningDate,"Opening Date")Then
				bverifyAccountInformation = False
			End If
		End If	
    End If
    If Not IsNull(strClosingDate) Then
    	If strClosingDate = "RUNTIME" Then
			strDate=strRunTimeClosingDate
			If Not strDate="          " Then			
				If len(Day(CDate(strDate)))=1 Then
        			strDay="0"&Day(CDate(strDate))
    			else
        			strDay=""&Day(CDate(strDate))
    			End If
    			strCISClosingDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
				strIServeClosingDate=bcKeyInfo.lblClosingDate.getroproperty("innertext")			
				If  Ucase(Trim(strCISClosingDate)) = UCase(Trim(strIServeClosingDate)) Then
					LogMessage "RSLT", "Verification","For Closing Date successfully matched with the expected value. Expected: "+ strCISClosingDate &" , Actual: "& strIServeClosingDate, True
					bverifyAccountInformation = True			
				else
					LogMessage "WARN", "Verification","For Closing Date not matching with the expected value. Expected: "+ strCISClosingDate &" , Actual: "& strIServeClosingDate, False
					bverifyAccountInformation = False
				End If
			ElseIf Not verifyInnerText(bcKeyInfo.lblClosingDate(),strDate,"Closing Date")Then
				bverifyAccountInformation = False
			End If
		Else	
			If Not verifyInnerText(bcKeyInfo.lblClosingDate(),strClosingDate,"Closing Date")Then
				bverifyAccountInformation = False
			End If			
		End If	
    End If
    verifyAccountInformation_SA=bverifyAccountInformation
End Function

'[Verify Monthly Savings Details in Key Info]
Public Function verifyMonthlySavings(strAccountScheme,strMonthlySavingsAmount,strStaffIndicator,strDebitingAccount,strDeductionDay)
	bverifyMonthlySavings=true
	If Not IsNull(strAccountScheme) Then
		If Not verifyInnerText(bcKeyInfo.lblAccountScheme(),strAccountScheme,"Account Scheme")Then
			bverifyMonthlySavings = False
		End If
    End If
    If Not IsNull(strMonthlySavingsAmount) Then
		If Not verifyInnerText(bcKeyInfo.lblMonthlySavingsAmount(),strMonthlySavingsAmount,"Monthly Savings Amount")Then
			bverifyMonthlySavings = False
		End If
    End If
    If Not IsNull(strStaffIndicator) Then
		If Not verifyInnerText(bcKeyInfo.lblStaffIndicator(),strStaffIndicator,"Staff Indicator")Then
			bverifyMonthlySavings = False
		End If
    End If
    If Not IsNull(strDebitingAccount) Then
		If Not verifyInnerText(bcKeyInfo.lblDebitingAccount(),strDebitingAccount,"Debiting Account")Then
			bverifyMonthlySavings = False
		End If
    End If
    If Not IsNull(strDeductionDay) Then
		If Not verifyInnerText(bcKeyInfo.lblDeductionDay(),strDeductionDay,"Deduction Day")Then
			bverifyMonthlySavings = False
		End If
    End If
    verifyMonthlySavings=bverifyMonthlySavings
End Function

'[Verify record in Address and Account Linkage Screen]
Public Function verifyAddressAcctLinkage(strProductCode,strCardNumber,strCIN,strName,strAddressType,strAddressCIN,strBlock,strLevelUnit,strUnit,strAddressLine1,strAddressLine2,strAddressLine3,strAddressLine4,strPostalCode,strLastUpdatedDate,strLastUpdatedBy,strSavingAccountNo)
	bverifyAddressAcctLinkage=true
	If Not IsNull(strName) Then
		If strName = "RUNTIME" Then
			getAccountInfo_CA strProductCode,strCardNumber				
			strCISName=strRunTimeShortName
    		strIServeName=bcVerify_AccountAndAddress.lblName.getroproperty("innertext")
    		If  Ucase(Trim(strCISName)) = UCase(Trim(strIServeName)) Then
    			LogMessage "RSLT", "Verification","For Name value successfully matched with the expected value. Expected: "+ strCISName &" , Actual: "& strIServeName, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Name value not matching with the expected value. Expected: "+ strCISName &" , Actual: "& strIServeName, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblName(),strName,"Name")Then
				bverifyAddressAcctLinkage = False
			End If
    	End If
    End If	
    If Not IsNull(strAddressType) Then
    	If strAddressType = "RUNTIME" Then
			strCISAddressType=strRunTimeAddressType
    		strIServeAddressType=bcVerify_AccountAndAddress.lblAddressType.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressType)) = UCase(Trim(strIServeAddressType)) Then
    			LogMessage "RSLT", "Verification","For Address Type value successfully matched with the expected value. Expected: "+ strCISAddressType &" , Actual: "& strIServeAddressType, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Type value not matching with the expected value. Expected: "+ strCISAddressType &" , Actual: "& strIServeAddressType, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressType(),strAddressType,"Address Type")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressCIN) Then
    	If strAddressCIN = "RUNTIME" Then
    		getProductInfo_CA (strCIN)
			strCISAddressCIN=strRunTimeAddressCIN
    		strIServeAddressCIN=bcVerify_AccountAndAddress.lblAddressCIN.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressCIN)) = UCase(Trim(strCISAddressCIN)) Then
    			LogMessage "RSLT", "Verification","For Address CIN value successfully matched with the expected value. Expected: "+ strCISAddressCIN &" , Actual: "& strIServeAddressCIN, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address CIN value not matching with the expected value. Expected: "+ strCISAddressCIN &" , Actual: "& strIServeAddressCIN, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressCIN(),strAddressCIN,"Address CIN")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strBlock) Then
    	If strBlock = "RUNTIME" Then
			strCISBlock=strRunTimeBlock
    		strIServeBlock=bcVerify_AccountAndAddress.lblBlock.getroproperty("innertext")
    		If  Ucase(Trim(strCISBlock)) = UCase(Trim(strIServeBlock)) Then
    			LogMessage "RSLT", "Verification","For Block value successfully matched with the expected value. Expected: "+ strCISBlock &" , Actual: "& strIServeBlock, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Block value not matching with the expected value. Expected: "+ strCISBlock &" , Actual: "& strIServeBlock, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblBlock(),strBlock,"Block")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strLevelUnit) Then
    	If strLevelUnit = "RUNTIME" Then
			strCISLevel=strRunTimeLevel
    		strIServeLevel=bcVerify_AccountAndAddress.lblLevelUnit.getroproperty("innertext")
    		If  Ucase(Trim(strCISLevel)) = UCase(Trim(strIServeLevel)) Then
    			LogMessage "RSLT", "Verification","For Level value successfully matched with the expected value. Expected: "+ strCISLevel &" , Actual: "& strIServeLevel, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Level value not matching with the expected value. Expected: "+ strCISLevel &" , Actual: "& strIServeLevel, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblLevelUnit(),strLevelUnit,"Level")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strUnit) Then
    	If strUnit = "RUNTIME" Then
			strCISUnit=strRunTimeUnit
    		strIServeUnit=bcVerify_AccountAndAddress.lblUnit.getroproperty("innertext")
    		If  Ucase(Trim(strCISUnit)) = UCase(Trim(strIServeUnit)) Then
    			LogMessage "RSLT", "Verification","For Unit value successfully matched with the expected value. Expected: "+ strCISUnit &" , Actual: "& strIServeUnit, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Unit value not matching with the expected value. Expected: "+ strCISUnit &" , Actual: "& strIServeUnit, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblUnit(),strUnit,"Unit")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine1) Then
    	If strAddressLine1 = "RUNTIME" Then
			strCISAddressLine1=strRunTimeAddressLine1
    		strIServeAddressLine1=bcVerify_AccountAndAddress.lblAddressLine1.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine1)) = UCase(Trim(strIServeAddressLine1)) Then
    			LogMessage "RSLT", "Verification","For Address Line1 value successfully matched with the expected value. Expected: "+ strCISAddressLine1 &" , Actual: "& strIServeAddressLine1, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line1 value not matching with the expected value. Expected: "+ strCISAddressLine1 &" , Actual: "& strIServeAddressLine1, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine1(),strAddressLine1,"Address Line1")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine2) Then
    	If strAddressLine2 = "RUNTIME" Then
			strCISAddressLine2=strRunTimeAddressLine2
    		strIServeAddressLine2=bcVerify_AccountAndAddress.lblAddressLine2.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine2)) = UCase(Trim(strIServeAddressLine2)) Then
    			LogMessage "RSLT", "Verification","For Address Line2 value successfully matched with the expected value. Expected: "+ strCISAddressLine2 &" , Actual: "& strIServeAddressLine2, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line2 value not matching with the expected value. Expected: "+ strCISAddressLine2 &" , Actual: "& strIServeAddressLine2, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine2(),strAddressLine2,"Address Line2")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine3) Then
    	If strAddressLine3 = "RUNTIME" Then
			strCISAddressLine3=strRunTimeAddressLine3
    		strIServeAddressLine3=bcVerify_AccountAndAddress.lblAddressLine3.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine3)) = UCase(Trim(strIServeAddressLine3)) Then
    			LogMessage "RSLT", "Verification","For Address Line3 value successfully matched with the expected value. Expected: "+ strCISAddressLine3 &" , Actual: "& strIServeAddressLine3, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line3 value not matching with the expected value. Expected: "+ strCISAddressLine3 &" , Actual: "& strIServeAddressLine3, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine3(),strAddressLine3,"Address Line3")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine4) Then
    	If strAddressLine4 = "RUNTIME" Then
			strCISAddressLine4=strRunTimeAddressLine4
    		strIServeAddressLine4=bcVerify_AccountAndAddress.lblAddressLine4.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine4)) = UCase(Trim(strIServeAddressLine4)) Then
    			LogMessage "RSLT", "Verification","For Address Line4 value successfully matched with the expected value. Expected: "+ strCISAddressLine4 &" , Actual: "& strIServeAddressLine4, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line4 value not matching with the expected value. Expected: "+ strCISAddressLine4 &" , Actual: "& strIServeAddressLine4, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine4(),strAddressLine4,"Address Line4")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strPostalCode) Then
    	If strPostalCode = "RUNTIME" Then
			strCISPostalCode=strRunTimePostalCode
    		strIServePostalCode=bcVerify_AccountAndAddress.lblPostalCode.getroproperty("innertext")
    		If  Ucase(Trim(strCISPostalCode)) = UCase(Trim(strIServePostalCode)) Then
    			LogMessage "RSLT", "Verification","For Postal Code value successfully matched with the expected value. Expected: "+ strCISPostalCode &" , Actual: "& strIServePostalCode, True
				bverifyAddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Postal Code value not matching with the expected value. Expected: "+ strCISPostalCode &" , Actual: "& strIServePostalCode, False
				bverifyAddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblPostalCode(),strPostalCode,"Postal Code")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strLastUpdatedDate) Then
    	If strLastUpdatedDate = "RUNTIME" Then
			strCISLastUpdatedDate=strRunTimeLastUpdatedInfo
    		strIServeLastUpdatedDate=bcVerify_AccountAndAddress.lblLastUpdatedDate.getroproperty("innertext")
    		strIServeLastUpdatedDate=FormatDateTime(strIServeLastUpdatedDate, 2)
    		If len(Month(CDate(strIServeLastUpdatedDate)))=1 Then
			
				strMonth="0"&Month(CDate(strIServeLastUpdatedDate))
			else
				strMonth=""&Month(CDate(strIServeLastUpdatedDate))
			End If
			strIServeLastUpdatedDate=Day(CDate(strIServeLastUpdatedDate))& "/"&strMonth& "/"&Year(CDate(strIServeLastUpdatedDate))&""
			
    		If Not matchStr(strCISLastUpdatedDate,strIServeLastUpdatedDate) Then    			
				bverifyAddressAcctLinkage = False			
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblLastUpdatedDate(),strLastUpdatedDate,"Last Updated Date")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strLastUpdatedBy) Then
    	If strLastUpdatedBy = "RUNTIME" Then
			strCISLastUpdatedBy=strRunTimeLastUpdatedInfo
    		strIServeLastUpdatedBy=bcVerify_AccountAndAddress.lblLastUpdatedBy.getroproperty("innertext")    		  		
    		If Not matchStr(strCISLastUpdatedBy,strIServeLastUpdatedBy) Then    			
				bverifyAddressAcctLinkage = False			
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblLastUpdatedBy(),strLastUpdatedBy,"Last Updated By")Then
				bverifyAddressAcctLinkage = False
			End If
		End If	
    End If
'    If Not IsNull(strSavingAccountNo) Then
'    	If strSavingAccountNo = "RUNTIME" Then
'			strCISLinkedSA=strRunTimeLinkedSA
'    		strIServeLinkedSA=bcVerify_AccountAndAddress.lblSavingAccountNo.getroproperty("innertext")
'    		If  Ucase(Trim(strCISLinkedSA)) = UCase(Trim(strIServeLinkedSA)) Then
'    			LogMessage "RSLT", "Verification","For Linked SA value successfully matched with the expected value. Expected: "+ strCISLinkedSA &" , Actual: "& strIServeLinkedSA, True
'				bverifyAddressAcctLinkage = True
'			else
'				LogMessage "WARN", "Verification","For Linked SA value not matching with the expected value. Expected: "+ strCISLinkedSA &" , Actual: "& strIServeLinkedSA, False
'				bverifyAddressAcctLinkage = False
'			End If
'		Else	
'			If Not verifyInnerText(bcVerify_AccountAndAddress.lblSavingAccountNo(),strSavingAccountNo,"Saving Account No")Then
'				bverifyAddressAcctLinkage = False
'			End If
'		End If	
'    End If
    verifyAddressAcctLinkage=bverifyAddressAcctLinkage
End Function

'[Verify Opening Balance in Transaction History Page]
Public Function verifyOpeningBalance(strOpeningBalance)
	bverifyOpeningBalance=true
	If Not IsNull (strOpeningBalance) Then
		If Not verifyInnerText(TransactionHistory.lblOpeningBalance(),strOpeningBalance,"Opening Balance")Then
			bverifyOpeningBalance = False
		End If
	End If
	verifyOpeningBalance=bverifyOpeningBalance
End Function

'[Verify Combobox Transaction Period in Transaction History has items]
Public Function verifyTransactionPeriod_ItemList(lstItems)
   bverifyTransactionPeriod_ItemList=true
   If Not IsNull(lstItems) Then	
       If Not verifyComboboxItems (TransactionHistory.lstTransactionPeriod(), lstItems, "Transaction Period")Then
           bverifyTransactionPeriod_ItemList=false
       End If
   End If
   verifyTransactionPeriod_ItemList=bverifyTransactionPeriod_ItemList
End Function

'[Select Combobox Transaction Period as]
Public Function selectTransactionPeriodComboBox(strTransactionPeriod)
   bDevPending=False
   bselectTransactionPeriodComboBox=true
   If Not IsNull(strTransactionPeriod) Then
	   TransactionHistory.lstTransactionPeriod.RefreshObject
       If Not (selectItem_Combobox (TransactionHistory.lstTransactionPeriod(), strTransactionPeriod))Then
            LogMessage "WARN","Verification","Failed to select :"&strTransactionPeriod&" From Transaction Period drop down list" ,false
           bselectTransactionPeriodComboBox=false
		Else
			  LogMessage "RSLT","Verification","Selected :"&strTransactionPeriod&" From Transaction Period drop down list" ,true
       End If
   End If
   WaitForICallLoading
   TransactionHistory.btnGo.Click
   WaitForICallLoading
   selectTransactionPeriodComboBox=bselectTransactionPeriodComboBox
End Function

'[Verify Transaction Table in Transaction History screen]
Public Function verifyTransactionHistoryTable(arrRowDataList)
	bverifyTransactionHistoryTable=true
	verifyTransactionHistoryTable=verifyTableContentList(TransactionHistory.tblTransactionsHeader,TransactionHistory.tblTransactionsContent,arrRowDataList,"Transaction History Table",false,null,null,null)
End Function

'[Verify Limits section should not available for Saving Account]
Public Function verifyLimitsExist()
	bverifyLimitsExist=true
	If BalancesAndLimits.lblLimits_OverdraftLimit.Exist Then
		LogMessage "WARN","Verification","Limits section available for SA. Expected to be unavailable." ,false
		bverifyLimitsExist=false
	End If
	verifyLimitsExist=bverifyLimitsExist
End Function

'[Verify the Current/History Earmarks link and details]
Public Function verifyEarmarksDetails()
	bverifyEarmarksDetails = true
End Function


''''''''''''''''''''''''''''''''' For Fee Reveresal (Verifying Account balance) added by  16 March 2016 ''''''''''''''''''''''''''''''

'[Verify Available Balance and Ledger Balance in Account Balance section displayed as]
Public Function verifyAccountBal(strCardNumber,strAvailableBal,strLedgerBal)
	WaitForICallLoading
	bverifyAccountBalance = True 
	If Not IsNull(strAvailableBal) Then
		If strAvailableBal = "RUNTIME" Then
			strIServeAvalBal=BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext")	
			Environment.Value("strIServeAvalBal") = strIServeAvalBal	
		Else  	
			strIServeAvalBal=BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext")
			If  Trim(strAvailableBal) = Trim(strIServeAvalBal) Then
				LogMessage "RSLT", "Verification","Available Balance displayed as expected. Expected:"&strAvailableBal&" , Actual:"&strIServeAvalBal&"", True
				bverifyAccountBalance = True
			else
				LogMessage "WARN", "Verification","Available Balance is not displayed as Expected:"&strAvailableBal&" , Actual:"&strIServeAvalBal&"", False
				bverifyAccountBalance = False
			End If
		End IF
    End If
  If Not IsNull(strLedgerBal) Then
    	If strLedgerBal = "RUNTIME" Then
           strIServeLedgerBal=BalancesAndLimits.lblAccountBalance_LedgerBalance.getroproperty("innertext")
           Environment.Value("strIServeLedgerBal") = strIServeLedgerBal
		End If 
  End If 	
    verifyAccountBal = bverifyAccountBalance
 End Function
 
'[Verify fee reversal amount added to AvailableBalance and LedgerBalance in AccountBalance section displayed as]
 Public Function verifyAccountBal_FRAdded(strCardNumber,strAvailableBal,strLedgerBal, StrRequestedAmount)
	bverifyAccountBal_FRAdded=true
	If Not IsNull(strAvailableBal) Then
		If strAvailableBal = "RUNTIME" Then
			strIServePrevAvalBal= Environment.Value("strIServeAvalBal")
			strIServePrevAvalBal = FormatNumber((strIServePrevAvalBal),2)
			strIserveRequestedFeeReversal = FormatNumber((StrRequestedAmount),2)
			StrIServeExpAvalBal = Ccur(strIServePrevAvalBal) +Ccur(strIserveRequestedFeeReversal)
			StrIServeExpAvalBal = FormatNumber((StrIServeExpAvalBal),2)
			strIServeCurrAvalBal = BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext")
			If  Trim(StrIServeExpAvalBal) = Trim(strIServeCurrAvalBal) Then
				LogMessage "RSLT", "Verification","Available Balance successfully matched with the expected value. Expected: "+ StrIServeExpAvalBal &" , Actual: "& strIServeCurrAvalBal, True
				bverifyAccountBal_FRAdded = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ StrIServeExpAvalBal &" , Actual: "& strIServeCurrAvalBal, False
				bverifyAccountBal_FRAdded = False
			End If
		End If
    End If
    
    If Not IsNull(strLedgerBal) Then
		If strLedgerBal = "RUNTIME" Then
			strIServePrevLedgerBal=Environment.Value("strIServeLedgerBal")
			strIServePrevLedgerBal = FormatNumber((strIServePrevLedgerBal),2)
			strIserveRequestedFeeReversal = FormatNumber((StrRequestedAmount),2)
			StrIServeExpLedgerBal = Ccur(strIServePrevLedgerBal) + Ccur(strIserveRequestedFeeReversal)
			StrIServeExpLedgerBal = FormatNumber((StrIServeExpLedgerBal),2)
			strIServeCurrLedgerBal = BalancesAndLimits.lblAccountBalance_LedgerBalance.getroproperty("innertext")	
			If  Trim(StrIServeExpLedgerBal) = Trim(strIServeCurrLedgerBal) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: "+ StrIServeExpLedgerBal &" , Actual: "& strIServeCurrLedgerBal, True
				bverifyAccountBal_FRAdded = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ StrIServeExpLedgerBal &" , Actual: "& strIServeCurrLedgerBal, False
				bverifyAccountBal_FRAdded = False
			End If
		End If
    End If
    verifyAccountBal_FRAdded = bverifyAccountBal_FRAdded
End Function 


'[Verify fee reversal amount is not added to AvailableBalance and LedgerBalance in AccountBalance section displayed as]
 Public Function verifyAccountBal_NoFRAdded(strCardNumber,strAvailableBal,strLedgerBal)
	bverifyAccountBal_NoFRAdded=true
	If Not IsNull(strAvailableBal) Then
		If strAvailableBal = "RUNTIME" Then
			strIServePrevAvalBal= FormatNumber(Environment.Value("strIServeAvalBal"),2)
			strIServeCurrAvalBal = FormatNumber(BalancesAndLimits.lblAccountBalance_AvailableBalance.getroproperty("innertext"),2)
			If  Ucase(Trim(strIServeCurrAvalBal)) = Ucase(Trim(strIServePrevAvalBal)) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: "+ strIServePrevAvalBal &" , Actual: "& strIServeCurrAvalBal, True
				bverifyAccountBal_NoFRAdded = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ strIServePrevAvalBal &" , Actual: "& strIServeCurrAvalBal, False
				bverifyAccountBal_NoFRAdded = False
			End If
		End If
    End If
    
    If Not IsNull(strLedgerBal) Then
		If strLedgerBal = "RUNTIME" Then
			strIServePrevLedgerBal=Environment.Value("strIServeLedgerBal")
			strIServePrevLedgerBal = Replace(strIServePrevLedgerBal,",","")
			strIServeCurrLedgerBal = BalancesAndLimits.lblAccountBalance_LedgerBalance.getroproperty("innertext")
			strIServeCurrLedgerBal = Replace(strIServeCurrLedgerBal,",","")
			If  Ucase(Trim(strIServeCurrLedgerBal)) = Ucase(Trim(strIServePrevLedgerBal)) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: "+ strIServePrevLedgerBal &" , Actual: "& strIServeCurrLedgerBal, True
				bverifyAccountBal_NoFRAdded = True
			else
				LogMessage "WARN", "Verification","For Available Balance not matching with the expected value. Expected: "+ strIServePrevLedgerBal &" , Actual: "& strIServeCurrLedgerBal, False
				bverifyAccountBal_NoFRAdded = False
			End If
		End If
    End If
    verifyAccountBal_NoFRAdded = bverifyAccountBal_NoFRAdded
End Function 

'[LISA Verify Earmark Details in popup window]
Public Function verifyEarmarkDetails_RowData(arrRowDataList)
   bDevPending=false
   bverifyEarmarkDetails_RowData= true
   verifyEarmarkDetails_RowData=verifyTableContentList(BalancesAndLimits.tblEarmarkHeader,BalancesAndLimits.tblEarmarkContent,arrRowDataList,"Earmark Details",false,null,null,null)
   verifyEarmarkDetails_RowData = bverifyEarmarkDetails_RowData
End Function

'[LISA Verify Signal Details in popup window]
Public Function verifySignalDetails_RowData(arrRowDataList)
   bDevPending=false
   bverifySignalDetails_RowData= true
   verifySignalDetails_RowData=verifyTableContentList(BalancesAndLimits.tblProductsListHeader,BalancesAndLimits.tblProductsListContent,arrRowDataList,"Signal Details",false,null,null,null)
   verifySignalDetails_RowData = bverifySignalDetails_RowData
End Function

'[LISA Verify data in Account Holder table in Key Info Page]
Public Function verifyAccountHolder_RowData(lstlstAccountHolder)
   bDevPending=false
   bverifyAccountHolder_RowData= true
   verifyAccountHolder_RowData=verifyTableContentList(bcKeyInfo.tblAccountHolderHeader,bcKeyInfo.tblAccountHolderContent,lstlstAccountHolder,"Account Holder",false,null,null,null)
   verifyAccountHolder_RowData = bverifyAccountHolder_RowData
End Function

'[Select the closed account from overview for account details enquiry]
Public Function clickclosedAcc_AccDetails()
	bclickclosedAcc_AccDetails = true
	bcCustomerOverview.tblclosedAccheader_AccountDetails.click
	clickclosedAcc_AccDetails = bclickclosedAcc_AccDetails
End Function
