Dim strSession:strSession="A"

'[Login to KRSP]
Public Function KRLogin(strUserId,strPassword)
   OpenPerComSession()

	Dim autECLSession 'As Object
	Set autECLSession = CreateObject("PCOMM.autECLSession")
	' Initialize the session
	strSession="A"
	autECLSession.SetConnectionByName (strSession)
	
	
	fWaitForAppAvailable(strSession)
	fWaitForInputReady(strSession)
	
	'Region
	fSetCursorPosition 20,027 'Region Location
	wait 1
	'fSendKeys "E020"
	'fSendKeys "G011"
	fSendKeys gstrRegion_KRSP
	fSendKeys "[enter]"
	fWaitForAppAvailable (strSession)
	fWaitForInputReady (strSession) 
	wait 1
	fSendKeys "[clear]"
	fWaitForInputReady (strSession) 
	wait 1
	'*End of Region
	
	'*CIS Login
	fSendKeys "cesn" 
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 10,026 'UserId Location
	fSendKeys strUserId 'Enter UserID
	fSendKeys "[tab]"
	fSendKeys "[tab]"
	fSendKeys strPassword 'Enter Password
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	'Confirm Sign On successfull
	strActualStatus = fGetText(strSession, "1", "011", "19")
	If instr(strActualStatus,"Sign-on is complete")=0 Then
		LogMessage "WARN","Verification","Login to KRSP failed",false
		KRLogin=false
	 else
	LogMessage "RSLT","Verification","Login to KRSP Successfull",True
		KRLogin=true
	End If

End Function
'*End of Login

'[Unblock Card with status other than 9 from KRSP]
Public Function untagCardNumber(strCardNumber)
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")

    	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 04,16    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "15", "017", "13")
	
	If InStr(strActualCardStatus,"9")=0 Then
	
		fSendKeys "[clear]" 
		fWaitForInputReady (strSession) 
		
		fSendKeys "abut" 'Enter the screen code for unblock card
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		
		'Verify the the screen      code - AB317M2                      or screen name - UNTAGGING HOT CARD
		fSetCursorPosition 09,41    ' Position for card number
		fSendKeys strCardNumber 'Enter Account Number
		fSendKeys "[pf7]"
		fWaitForInputReady (strSession)
		wait 1
		strActualStatus = fGetText(strSession, "22", "002", "99")
		If instr(Trim(strActualStatus),"1010 CARD SUCCESSFULLY UPDATED")=0 AND instr(Trim(strActualStatus),"4662 CARD IS ACTIVE, UNTAG NOT ALLOWED")=0 Then
			LogMessage "WARN","Verification","Untag Card Failed On KRSP failed",false
			untagCardNumber=false
			Exit Function
		 else
		LogMessage "RSLT","Verification","CARD SUCCESSFULLY UPDATED",True
			untagCardNumber=true
			Exit Function
		End If
	else	
			LogMessage "WARN","Verification","Card Status is "&strActualCardStatus& " For Card Number:"&strCardNumber,true		
	End If
	untagCardNumber=true
End Function

'[Verify Card Status and Reason Code for multiple cards in KRSP Screen ABIQ]
Public Function verifyCardStatus_reason_MultipleCards_KR(lstCardNumber,strCardStatus,strReason)
   Dim bStatus:bStatus=true
	For iCount=0 to Ubound(lstCardNumber)
		strCardNumber=lstCardNumber(iCount)
		bVerify=verifyCardStatusandReason_KR(strCardNumber,strCardStatus,strReason)
		If not  bVerify Then
			bStatus=false
		End If
	Next
	verifyCardStatus_reason_MultipleCards_KR=bStatus
End Function
'[Verify Card Status and Reason Code in KRSP Screen ABIQ]
Public Function verifyCardStatusandReason_KR(strCardNumber,strCardStatus,strReason)
'Verify in ABIQ screen for updated Card Status and Reason Code
	strCardNumber=Replace(strCardNumber,"-","")
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 04,16    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "15", "017", "13")
	strActualReason = fGetText(strSession, "15", "040", "1")
	strActualDateTime = fGetText(strSession, "16", "043", "16")
	If strActualDateTime <> "00/00/0000 00:00" Then
		If len(Day(CDate(strActualDateTime)))=1 Then
					strDay="0"&Day(CDate(strActualDateTime))
		else
				strDay=""&Day(CDate(strActualDateTime))
		End If

		If len(Minute(CDate(strActualDateTime)))=1 Then
			strMinutes="0"&Minute(CDate(strActualDateTime))
		else
			strMinutes=""&Minute(CDate(strActualDateTime))
		End If
				strLastUpdatedDate=""&strDay & " "&monthName(Month(CDate(strActualDateTime)),true) &" " &Year(CDate(strActualDateTime)) &" "& Hour(CDate(strActualDateTime)) &":"&strMinutes
		 insertDataStore "DateAndTime_KR", strLastUpdatedDate	
	End If
	If Ucase(Trim(strActualCardStatus))=Ucase(Trim(strCardStatus)) Then
        LogMessage "RSLT","Verification","KRSP: Card Status For Card Number "&strCardNumber&" matched with expected Status"&strCardStatus,True		
	 else
		LogMessage "WARN","Verification","KRSP - Actual Card Status: "&strActualCardStatus&" does not matched with expected Status :"&strCardStatus&" For Card Number "&strCardNumber,false
		bKRResult=false
	End If	
	If Ucase(Trim(strActualReason))=Ucase(Trim(strReason))Then
		LogMessage "RSLT","Verification","KRSP: Card Reason matched with expected Reason"&strReason & " For Card Number "&strCardNumber,True
		
	 else
		LogMessage "WARN","Verification","KRSP - Actual Card Reason: "&strActualReason&" does not matched with expected Reason :"&strReason & " For Card Number "&strCardNumber,false
		bKRResult=false
	End If	
	verifyCardStatusandReason_KR=bKRResult
End Function

'[Verify Card Status and Reason Code in KRSP Screen ABRS]
Public Function verifyCardStatusandReason_KR_ABRS(strCardNumber,strCardStatus,strReason)
	'Verify in ABRS screen for updated Card Status and Reason Code for Cards with status code 9
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abrs"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 08,002    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "08", "031", "3")
	strActualReason = fGetText(strSession, "08", "038", "30")
	If Ucase(Trim(strActualCardStatus))=Ucase(Trim(strCardStatus)) Then
        LogMessage "RSLT","Verification","KRSP: Card Status matched with expected Status"&strCardStatus,True		
	 else
		LogMessage "WARN","Verification","KRSP - Actual Card Status: "&strActualCardStatus&" does not matched with expected Status :"&strCardStatus,false
		bKRResult=false
	End If	
	If Ucase(Trim(strActualReason))=Ucase(Trim(strReason))Then
		LogMessage "RSLT","Verification","KRSP: Card Status matched with expected Status"&strReason,True
		
	 else
		LogMessage "WARN","Verification","KRSP - Actual Card Status: "&strActualReason&" does not matched with expected Status :"&strReason,false
		bKRResult=false
	End If	
	verifyCardStatusandReason_KR_ABRS=bKRResult
End Function

'[Unblock Card with status 9 from KRSP]
Public Function activateCard_Status9_KR(strCardNumber)
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")

	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 04,16    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "15", "017", "13")
	If InStr(strActualCardStatus,"9")<>0 Then
		fSendKeys "[clear]" 
		fWaitForInputReady (strSession)
		
		
		fSendKeys "abrs"
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		
		
		fSetCursorPosition 08,002    ' Position for card number
		fSendKeys strCardNumber 'Enter Account Number
		fSendKeys "[enter]"
		fSendKeys "[pf7]"
		fWaitForInputReady (strSession)
		strActualStatus = fGetText(strSession, "24", "002", "99")
		If instr(Trim(strActualStatus),"1010 CARD SUCCESSFULLY UPDATED")=0 AND instr(Trim(strAccountStatus),"CARD IS ACTIVE")=0 Then
			LogMessage "WARN","Verification","Untag Card Failed On KRSP failed",false
			activateCard_Status9_KR=false
		 else
		LogMessage "RSLT","Verification","CARD SUCCESSFULLY UPDATED",True
			activateCard_Status9_KR=true
		End If
	else
		LogMessage "WARN","Verification","Card Status is "&strActualCardStatus& " For Card Number"&strCardNumber,true	
		activateCard_Status9_KR=true	
	End If
End Function
'[Change Card Status as Deny in KRSP Screen ABTU]
Public Function DenyCard_KR_ABTU(strCardNumber)
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abtu"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	fSetCursorPosition 03,009    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 08,015    ' Position for card status
	fSendKeys "1" 'Enter Account Number
	fSendKeys "[pf7]"
	wait 1
	strStatus = fGetText(strSession, "22", "002", "40")

	If Trim(strStatus)="1010 CARD SUCCESSFULLY UPDATED" Then
		LogMessage "RSLT","Verification","Card status successfull updated as DENY For Card Number "&strCardNumber,true
		bKRResult=true
	else
		LogMessage "WARN","Verification","Failed to update Card status as DENY For Card Number "&strCardNumber&" Error : "&strStatus,false
		bKRResult=false
	End If
		DenyCard_KR_ABTU=bKRResult
End Function

'[Change Card Status as RETAINED in KRSP Screen ABIQ]
Public Function RetainCard_KR_ABIQ(strCardNumber)
	'Verify in ABRS screen for updated Card Status and Reason Code for Cards with status code 9
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 04,016    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "15", "017", "10")
	strActualReason = fGetText(strSession, "15", "040", "30")
	If strActualCardStatus = "9 (CLOSED)" Then
		'unblock card
		activateCard_Status9_KR(strCardNumber)
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
	End If

	If strActualCardStatus="2 (RETAIN)" Then
		RetainCard_KR_ABIQ=true
		Exit Function 
	End If
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)

	fSetCursorPosition 04,016    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)

	fSendKeys "[pf6]"
	fWaitForInputReady (strSession)
	wait 1
	fSendKeys "[pf7]"
	fWaitForInputReady (strSession)
	wait 1
	strActualStatus = fGetText(strSession, "15", "017", "10")
	
	If Ucase(Trim(strActualStatus))=Ucase(Trim("2 (RETAIN)")) Then
        LogMessage "RSLT","Verification","KRSP: Card Successfully Retained",True		
	 else
		LogMessage "WARN","Verification","KRSP - Failed to Retain Card",false
		bKRResult=false
	End If	
	RetainCard_KR_ABIQ=bKRResult
End Function

'[Change Card Status as CLOSED in KRSP Screen ABIQ]
Public Function CloseCard_KR_ABIQ(strCardNumber,strReasonCode)
	'Verify in ABRS screen for updated Card Status and Reason Code for Cards with status code 9
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	wait 1
	fSetCursorPosition 04,016    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "15", "017", "10")
	strActualReason = fGetText(strSession, "15", "040", "30")
	If strActualCardStatus <> "9 (CLOSED)" Then
		
		fSendKeys "[pf9]"
		wait 1
		fWaitForInputReady (strSession)
		If not isNull(strReasonCode) Then
			fSetCursorPosition 12,017    ' Position for card number
			fSendKeys strReasonCode
		End If
		fSendKeys "[pf7]"
		fWaitForInputReady (strSession)
		wait 1
		strCardUpdateStatus = fGetText(strSession, "22", "002", "30")
'		If strActualCardStatus<>"2 (RETAIN)" And strActualCardStatus<>"0 (ACTIVE)"Then
'			wait 1
'			fSendKeys "[pf1]"
'		End If
'		fSendKeys "[clear]"
'		
'		fWaitForInputReady (strSession)
'		wait 1
		strActualStatus = fGetText(strSession, "15", "017", "10")
	
		If Ucase(Trim(strCardUpdateStatus))=Ucase(Trim("1010 CARD SUCCESSFULLY UPDATED")) OR Trim(strActualStatus)="9 (CLOSED)"Then
			LogMessage "RSLT","Verification","KRSP: Card Successfully Closed",True		
		 else
			LogMessage "WARN","Verification","KRSP - Failed to Close Card",false
			bKRResult=false
		End If	
	End If 
	CloseCard_KR_ABIQ=bKRResult
End Function
'[Change Card Status in KRSP Screen as]
Public Function changeCardStatus_KR(strCardNumber, strCardStatus,strReasonCode)
   If not isNull(strCardStatus) Then
		Select Case Ucase(strCardStatus)
        Case "CLOSED"
			changeCardStatus_KR=CloseCard_KR_ABIQ(strCardNumber,strReasonCode)
		 Case "RETAIN"
			changeCardStatus_KR=RetainCard_KR_ABIQ(strCardNumber)
		 Case "DENY"
			changeCardStatus_KR=DenyCard_KR_ABTU(strCardNumber)
		End Select
	else
		changeCardStatus_KR=true
   End If
End Function
'[Unblock Closed Cards blocked from KRSP]
Public Function ActivateMultipleBlockedCards()
'	strBlockedCards=fetchFromDataStore(strRunTimeDataStore_BlockedCardStep,"BLANK",strRunTimeDataStore_VarName)(0)
		strBlockedCards=fetchFromDataStore("Verify row Data in Table Status Block Cards","BLANK","SuccessfullBlockedCards")(0)
	If  (isNull(strBlockedCards) or isEmpty(strBlockedCards)) Then
			strBlockedCards=fetchFromDataStore("Verify Link SRNumber available on Request Submitted popup","BLANK","SingleBlockedCard")(0) 'SingleBlocked Cards
	End If
	If  (isNull(strBlockedCards) or isEmpty(strBlockedCards)) Then
			strBlockedCards=fetchFromDataStore("Verify row Data in Table CardSummary on NewSR Screen","BLANK","CanceledCards")(0) 'BlockedCancelledCards
	End If

	If Not (isNull(strBlockedCards) or isEmpty(strBlockedCards)) Then
		lstCards=Split(strBlockedCards,"|")
		For iCount=0 to Ubound(lstCards)
			If lstCards(iCount)<>"" Then
				strCardToUnblock=Split(lstCards(iCount),":")(0)
				'strCardDesc=Split(lstCards(iCount),":")(1)
	
				activateCard_Status9_KR strCardToUnblock
	
				 untagCardNumber strCardToUnblock
			End If
		Next
	End If

End Function

'[Activate Closed Cards from KRSP and Vplus]
Public Function ActivateCards(lstCardNumber)
	'VPlus
	UnblockCards_VPlus lstCardNumber
	'KRSP
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	
	For iCount=0 to Ubound(lstCardNumber)
	
		strCardNumber=lstCardNumber(iCount)
		strCardNumber=Replace(strCardNumber,"-","")
    	fSendKeys "[clear]" 
		fWaitForInputReady (strSession)
		
		
		fSendKeys "abiq"
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		
		
		fSetCursorPosition 04,16    ' Position for card number
		fSendKeys strCardNumber 'Enter Account Number
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		
		'Verify the Card Status in location 15, 017
		strActualCardStatus = fGetText(strSession, "15", "017", "13")
		
		If Trim(strActualCardStatus)="0 (ACTIVE)" Then
			LogMessage "RSLT","Verification","Card "&strCardNumber&" is already Active",true
		else
			If InStr(strActualCardStatus,"9")<>0 Then
				fSendKeys "[clear]" 
				fWaitForInputReady (strSession)
				
				wait 1
				fSendKeys "abrs"
				fSendKeys "[enter]"
				fWaitForInputReady (strSession)
				
				
				fSetCursorPosition 08,002    ' Position for card number
				fSendKeys strCardNumber 'Enter Account Number
				fSendKeys "[enter]"
				'wait 1
				fWaitForInputReady (strSession)
				fSendKeys "[pf7]"
				fWaitForInputReady (strSession)
				strActualStatus = fGetText(strSession, "24", "002", "99")
				If instr(Trim(strActualStatus),"1010 CARD SUCCESSFULLY UPDATED")=0 AND instr(Trim(strAccountStatus),"CARD IS ACTIVE")=0 Then
					LogMessage "WARN","Verification","Untag Card "&strCardNumber&"  Failed On KRSP failed",True
                else
					LogMessage "RSLT","Verification","CARD "&strCardNumber&" SUCCESSFULLY UPDATED",True
                End If
			End If 
			
			If InStr(strActualCardStatus,"9")=0 AND Trim(strActualCardStatus)<>"0 (ACTIVE)" Then
		
				fSendKeys "[clear]" 
				fWaitForInputReady (strSession) 
				
				fSendKeys "abut" 'Enter the screen code for unblock card
				fSendKeys "[enter]"
				fWaitForInputReady (strSession)
				
				'Verify the the screen      code - AB317M2                      or screen name - UNTAGGING HOT CARD
				fSetCursorPosition 09,41    ' Position for card number
				fSendKeys strCardNumber 'Enter Account Number
				fSendKeys "[pf7]"
				fWaitForInputReady (strSession)
				strActualStatus = fGetText(strSession, "22", "002", "99")
				If instr(Trim(strActualStatus),"1010 CARD SUCCESSFULLY UPDATED")=0 AND instr(Trim(strAccountStatus),"CARD IS ACTIVE")=0 Then
					LogMessage "WARN","Verification","Untag Card  "&strCardNumber&" Failed On KRSP failed",True
				 else
					LogMessage "RSLT","Verification","CARD  "&strCardNumber&" SUCCESSFULLY UPDATED",True
				End If
			End If
		End If
		fSendKeys "[clear]" 
	Next


	ActivateCards=true
End Function
'*******************Data Retrival Function : New PIN******************
'[Recover New PIN Card data]
Public Function RecoverDATA_NewPIN_KRSP(strCardNumber)
   If isNull(strCardNumber) Then
	   strCardNumber=fetchFromDataStore("Verify Field CardNumber on Request Submitted Popup for New Pin displayed as","BLANK","NewPINUsedCard")(0)
   End If

 If isEmpty(strCardNumber) Then
	RecoverDATA_NewPIN_KRSP=true
	Exit Function
 End If
	Dim bKRResult:bKRResult=true
	If not isNull(strCardNumber) Then
		KRLogin gstrUser_KRSP,gstrPassword_KRSP
		strCardNumber=Replace(strCardNumber,"-","")
		wait 1
		fWaitForInputReady (strSession)
		fSendKeys "[clear]" 
	
		fSetCursorPosition 01,01    ' Position for card number
		fSendKeys "xped 5.1.2" '
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		Wait 1
		fSetCursorPosition 06,12    
		fSendKeys "abrpin"
		fSendKeys "[enter]"
		fWaitForInputReady (strSession) '
		fSetCursorPosition 09,013    
		fSendKeys strCardNumber
		fSendKeys "[enter]"
		fWaitForInputReady (strSession) 
		wait 1
		strCard= fGetText(strSession, "16", "013", "16")
		If len(strCardNumber)=15 Then 'For handling Amex Card
			strCardNumber=strCardNumber&"0"
		End If
		'Check 1st 8 cards only. This file data is very big to loop all the cards
		For iCount =0 to 8
			strRow= 16 + iCount
			strCard= fGetText(strSession, strRow, "013", "16")
		
			If Trim(strCard)=strCardNumber Then
				fSetCursorPosition strRow,003 
				fSendKeys "D"
				fSendKeys "[enter]"	
				fWaitForInputReady (strSession) 
				Exit For
			Else
				LogMessage "WARN","Verification","Card Number "&strCardNumber&" not found in New PIN recover List.",True
			End If
		
		Next
   End If
   RecoverDATA_NewPIN_KRSP=bKRResult
End Function
'--------------------------------
'[Verify Card Status in KRSP Screen KRCI]
Public Function verifyCardStatus_KR_KRCI(strCardNumber,strCardStatus)
	'
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "krci"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 14,009    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	wait 1
	fSendKeys "[pf2]"
	fWaitForInputReady (strSession)
	wait 1
	'Verify the Card Status in location 15, 017
	strActualCardStatus = fGetText(strSession, "15", "017", "10")
	If Ucase(Trim(strActualCardStatus))=Ucase(Trim(strCardStatus)) Then
        LogMessage "RSLT","Verification","KRSP: Card Status on KRCI matched with expected Status"&strCardStatus,True		
	 else
		LogMessage "WARN","Verification","KRSP - Actual Card Status: "&strActualCardStatus&" does not matched with expected Status :"&strCardStatus & " from KRCI",false
		bKRResult=false
	End If	
	verifyCardStatus_KR_KRCI=bKRResult
End Function

'[Verify ST NAME BLKNO LVL/UNIT POST CODE not blank in KRSP Screen KRCI]
Public Function verifySTNAME_BLKNO_LVL_POSTCODE_KR_KRCI(strCardNumber)
	'
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "krci"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition 14,009    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	wait 1
	fSendKeys "[pf2]"
	fWaitForInputReady (strSession)
	wait 1
	'Verify the Card Status in location 15, 017
	strSTName1 = fGetText(strSession, "09", "040", "10")
	strBlkNo = fGetText(strSession, "10", "016", "10")
	strLVLUnit = fGetText(strSession, "11", "016", "10")
	strPostalCode = fGetText(strSession, "12", "016", "6")
	
	If Trim(strSTName1)<>"" OR Trim(strBlkNo)<>"" OR Trim(strLVLUnit)<>"" OR Trim(strPostalCode)<>""  Then
        LogMessage "RSLT","Verification","KRSP: ST NAME, BLKNO, LVL/UNIT, POST CODE not blank on KRCI as expected",True		
	 else
		LogMessage "WARN","Verification","KRSP: Either of ST NAME, BLKNO, LVL/UNIT, POST CODE is blank on KRCI. All should not be Blank",false
		bKRResult=false
	End If	
	verifySTNAME_BLKNO_LVL_POSTCODE_KR_KRCI=bKRResult
End Function

'[Verify PIN GEN date in KRSP Screen KRCI]
Public Function verifyPIN_GEN_KRCI(strCardNumber,strPinGenDate)
	If Ucase(strPinGenDate)="TODAY" Then
		If len(Day(CDate(Now)))=1 Then
			strDay="0"&Day(CDate(Now))
		else
			strDay=""&Day(CDate(Now))
		End If
		If len(Month(CDate(Now)))=1 Then
			strMonth="0"&Month(CDate(Now))
		else
			strMonth=""&Month(CDate(Now))
		End If
		strPinGenDate=""&strDay & "/"&strMonth&"/" &Year(CDate(Now))
    End If
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	wait 1
	fSendKeys "krci"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	wait 1
	fSetCursorPosition 14,009    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fSendKeys "[pf2]"
	fWaitForInputReady (strSession)
	wait 1
	'Verify the Card Status in location 15, 017
	strPinGen = fGetText(strSession, "17", "043", "10")

	
	If Trim(strPinGen)=strPinGenDate  Then
        LogMessage "RSLT","Verification","KRSP: PIN Gen matcehd on KRCI as expected",True		
	 else
		LogMessage "RSLT","Verification","KRSP:PIN Gen "&strPinGen&" does not matched with Expected "&strPinGenDate&" on KRCI",false
		bKRResult=false
	End If	
	verifyPIN_GEN_KRCI=bKRResult
End Function

'[Verify PIN Issue and Last Pin Issue date in KRSP Screen ABIQ]
Public Function verifyPINIssue_LastPin_KRCI(strCardNumber,strPinIssue,strLastPinIssueDate)
	'
	'TODO Date handling for TODAY
	strCardNumber=Replace(strCardNumber,"-","")
	If Ucase(strLastPinIssueDate)="TODAY" Then
		If len(Day(CDate(Now)))=1 Then
			strDay="0"&Day(CDate(Now))
		else
			strDay=""&Day(CDate(Now))
		End If
		If len(Month(CDate(Now)))=1 Then
			strMonth="0"&Month(CDate(Now))
		else
			strMonth=""&Month(CDate(Now))
		End If
		strLastPinIssueDate=""&strDay & "/"&strMonth&"/" &Year(CDate(Now))
    End If
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	
	fSetCursorPosition 04,016    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSendKeys "[pf12]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strActualPinIssue = fGetText(strSession, "15", "053", "02")
	strActualPinIssueDt = fGetText(strSession, "15", "068", "10")
	
	If Trim(strActualPinIssue)=strPinIssue  Then
        LogMessage "RSLT","Verification","KRSP: Actual PINs "&strActualPinIssue&" Issued matched on ABIQ as expected",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual PINs Issued "&strActualPinIssue&" does not matched with Expected "&strPinIssue&" on ABIQ",false
		bKRResult=false
	End If	
	If Trim(strActualPinIssueDt)=strLastPinIssueDate  Then
        LogMessage "RSLT","Verification","KRSP: Actual Last Issued Date "&strActualPinIssueDt&" matched on ABIQ as expected",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual Last Issued Date "&strActualPinIssueDt&" does not matched with Expected "&strLastPinIssueDate&" on ABIQ",false
		bKRResult=false
	End If	
	verifyPINIssue_LastPin_KRCI=bKRResult
End Function

'[Verify ACC TYPE BRD ACCOUNT NO and PS are not blank and not tagged as ERR-CLOSED in KRSP Screen ABIQ]
Public Function verifyACC_BRD_PS_KRCI(strCardNumber)
	'
	'TODO Date handling for TODAY
	strCardNumber=Replace(strCardNumber,"-","")
	
	Dim bKRResult:bKRResult=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	
	
	fSendKeys "abiq"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	
	fSetCursorPosition 04,016    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSendKeys "[pf12]"
	fWaitForInputReady (strSession)
	
	'Verify the Card Status in location 15, 017
	strAccount_Actual = fGetText(strSession, "19", "002", "02")
    strTypeActual= fGetText(strSession, "19", "007", "1")
	strBRDActual= fGetText(strSession, "19", "012", "1")
	strAccNumActual= fGetText(strSession, "19", "016", "12")
	strPSActual= fGetText(strSession, "19", "030", "1")
	strERR= fGetText(strSession, "19", "033", "7")
	If Trim(strAccount_Actual)<>""  Then
        LogMessage "RSLT","Verification","KRSP: Actual ACC "&strAccount_Actual&" is not blank as expected on ABIQ-AB05M2 Screen",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual ACC "&strAccount_Actual&" does not matched with Expected is not BLANK on ABIQ-AB05M2 Screen",false
		bKRResult=false
	End If	
	If Trim(strTypeActual)<>""  Then
        LogMessage "RSLT","Verification","KRSP: Actual TYPE "&strTypeActual&" is not blank as expected on ABIQ-AB05M2 Screen",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual TYPE "&strTypeActual&" does not matched with Expected is not BLANK on ABIQ-AB05M2 Screen",false
		bKRResult=false
	End If	
	If Trim(strBRDActual)<>""  Then
        LogMessage "RSLT","Verification","KRSP: Actual BRD "&strBRDActual&" is not blank as expected on ABIQ-AB05M2 Screen",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual BRD "&strBRDActual&" does not matched with Expected is not BLANK on ABIQ-AB05M2 Screen",false
		bKRResult=false
	End If	
	If Trim(strAccNumActual)<>""  Then
        LogMessage "RSLT","Verification","KRSP: Actual Account Number "&strAccNumActual&" is not blank as expected on ABIQ-AB05M2 Screen",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual Account Number "&strAccNumActual&" does not matched with Expected is not BLANK on ABIQ-AB05M2 Screen",false
		bKRResult=false
	End If	
	If Trim(strPSActual)<>""  Then
        LogMessage "RSLT","Verification","KRSP: Actual P/S "&strPSActual&" is not blank as expected on ABIQ-AB05M2 Screen",True		
	 else
		LogMessage "RSLT","Verification","KRSP:Actual P/S "&strPSActual&" does not matched with Expected "&strPS&" on ABIQ-AB05M2 Screen",false
		bKRResult=false
	End If

	If Instr(1,Trim(strERR),"ERR")=0 OR InStr(1,Trim(strERR),"CLOSED")=0 Then
        LogMessage "RSLT","Verification","KRSP: ACC-TYPE  BRD  ACCOUNT NO and PS are tagged as ERR-CLOSED on ABIQ-AB05M2 Screen",True		
	 else
		LogMessage "RSLT","Verification","KRSP:ACC-TYPE  BRD  ACCOUNT NO and PS are tagged as ERR-CLOSED on ABIQ-AB05M2 Screen",false
		bKRResult=false
	End If				
	verifyACC_BRD_PS_KRCI=bKRResult
End Function

'[Reset Amount from KRSL KRSP Screen]
Public Function resetAmount_KRSL(strCardNumber,strAmount)
	bresetAmount_KRSL=true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	fSendKeys "[clear]" 
	'fWaitForInputReady (strSession)
	fSendKeys "krsl"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	strCardNumber=Replace(strCardNumber,"-","")
	fSetCursorPosition 04,016    ' Position for card number
	fSendKeys strCardNumber 'Enter Account Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 11,052
	fSendKeys strAmount
	strDate=fGetText(strSession, "01", "071", "10")
	strDate=Replace(strDate,"/","")
	fSetCursorPosition 13,021
	fSendKeys strDate
	fSetCursorPosition 13,052
	fSendKeys strDate
	fSendKeys "[pf7]"
	fWaitForInputReady (strSession)
	strActualStatus=fGetText(strSession, "21", "002", "30")
	If strActualStatus = "1010 CARD SUCCESSFULLY UPDATED" Then
		LogMessage "RSLT", "Verification", "New Spending Limit amount reset successfully.", true
		bresetAmount_KRSL=true
	else
		LogMessage "WARN", "Verification", "Failed to reset Amount", false
		bresetAmount_KRSL=false		
	End If
	resetAmount_KRSL=bresetAmount_KRSL
End Function

'[Navigate to A419 in KRSP]
Public Function navigateToA419(strCardNo, strDay)
	bDevPending=false
	bverifyATMCAM = true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	fSendKeys "[clear]" 
	'fWaitForInputReady (strSession)
	fSendKeys "A419"
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 05,015     'Position for the Day
	'***************** Check if the day is current or Previous
	Select Case (strDay)
		Case ("Current")
			fSendKeys "C"
		Case ("Previous")
			fSendKeys "P"
	End Select
	'fSendKeys "[enter]"
	wait 2
	fSetCursorPosition 13,029     'Position for the Card Number
	fSendKeys strCardNo 'Enter Account Number
	fSendKeys "[enter]"
	verifyATMCAM = bverifyATMCAM
End Function

'[Verify the values displayed in the Terminal Transaction Details as]
Public Function verifyTerminalTranDetails(lstTerminalDetails)
	bverifyTerminalTranDetails = true
	'read the record number from iServe
	strRecordNoFE = Claims.lblRecordNo().GetRoProperty("innertext")
	'Find the same Record no in Main Frame and start comparing
	For counter = 1 To 20 Step 1
		strRecordNoMF = fGetText(strSession, "03","041","5")
		If strRecordNoMF = strRecordNoFE Then
			'now start comparing the main frame and the FE
			If not isnull (lstTerminalDetails) Then
			For iCount = 0 To ubound(lstTerminalDetails) Step 1
				strLabel = Split(lstTerminalDetails(iCount),":")(0)
				strDetail = Split(lstTerminalDetails(iCount),":")(1)
				Select Case strLabel
					Case "Card Owner"
						Call strCardOwnerComparison(strDetail)
					Case "Business Date"
						Call strBusinessDateComparison(strDetail)
					Case "Sequence No"
						Call strSequenceNoComparison(strDetail)
					Case "Sub Class"
						Call strSubClassComparison(strDetail)
					Case "Transaction Type"
						Call strTransactionTypeComparison(strDetail)
					Case "Ledger Balance"
						Call strLedgerBalanceComparison(strDetail)
					Case "Available Balance"
						Call strAvailableBalanceComparison(strDetail)
					Case "Terminal Owner"
						Call strTerminalOwnerComparison(strDetail)
					Case "Location"
						Call strLocationComparison(strDetail)
					Case "ARQC"
						Call strARQCComparison(strDetail)
					Case "Completion Status"
						Call strCompletionStatusComparison(strDetail)
					Case "Host Indicator"
						Call strHostIndicatorComparison(strDetail)
					Case "PIN Tries"
						Call strPinTriesComparison(strDetail)
					Case "Chip Sequences"
						Call strChipSequencesComparison(strDetail)
					Case "ATB"
						Call strATBComparison(strDetail)
				End Select
			Next
		End If
		
	Exit For
		else
			'Go to the previous screen
			fSendKeys "[pf7]"
			'strRecordNoMF = fGetText(strSession, "03","041","5")
		End If
	Next
	verifyTerminalTranDetails = bverifyTerminalTranDetails
End Function

'********** Function to compare ATB
Public Function strATBComparison(strATB)
	bstrATBComparison = true
	If Not IsNull(strATB) Then
		If strATB = "RUNTIME" Then
			strATBMF = trim(fGetText(strSession, "20","032","4"))
			strATBFE = Claims.lblATB().GetRoProperty("innertext")
			
			If  Ucase(Trim(strChipSequencesFE)) = UCase(Trim(strChipSequencesMF)) Then
				LogMessage "RSLT", "Verification","For ATB successfully matched with the expected value. Expected: " & strATBMF &" , Actual: "& strATBFE, True
				bstrATBComparison = True
			else
				LogMessage "WARN", "Verification","For ATB successfully matched with the expected value. Expected: " & strATBMF &" , Actual: "& strATBFE, False
				bstrATBComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblATB(),strATB,"ATB")Then
				bstrATBComparison = False
			End If
		End If
    End If
	strATBComparison = bstrATBComparison
End Function

'********* Function to compare Chip Sequences
Public Function strChipSequencesComparison(strChipSequences)
	bstrChipSequencesComparison = true
	If Not IsNull(strChipSequences) Then
		If strChipSequences = "RUNTIME" Then
			strChipSequencesMF = trim(fGetText(strSession, "20","068","4"))
			strChipSequencesFE = Claims.lblChipSequences().GetRoProperty("innertext")
			
			If  Ucase(Trim(strChipSequencesFE)) = UCase(Trim(strChipSequencesMF)) Then
				LogMessage "RSLT", "Verification","For Pin Tries successfully matched with the expected value. Expected: " & strChipSequencesMF &" , Actual: "& strChipSequencesFE, True
				bstrChipSequencesComparison = True
			else
				LogMessage "WARN", "Verification","For Pin Tries successfully matched with the expected value. Expected: " & strChipSequencesMF &" , Actual: "& strChipSequencesFE, False
				bstrChipSequencesComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblChipSequences(),strChipSequences,"Pin Tries")Then
				bstrChipSequencesComparison = False
			End If
		End If
    End If
	strChipSequencesComparison = bstrChipSequencesComparison
End Function

'******** Function to compare Pin Tries
Public Function strPinTriesComparison(strPinTries)
	bstrPinTriesComparison = true
	If Not IsNull(strPinTries) Then
		If strPinTries = "RUNTIME" Then
			strPinTriesMF = trim(fGetText(strSession, "18","068","2"))
			strPinTriesFE = Claims.lblPinTries().GetRoProperty("innertext")
			
			If  Ucase(Trim(strPinTriesFE)) = UCase(Trim(strPinTriesMF)) Then
				LogMessage "RSLT", "Verification","For Pin Tries successfully matched with the expected value. Expected: " & strPinTriesMF &" , Actual: "& strPinTriesFE, True
				bstrPinTriesComparison = True
			else
				LogMessage "WARN", "Verification","For Pin Tries successfully matched with the expected value. Expected: " & strPinTriesMF &" , Actual: "& strPinTriesFE, False
				bstrPinTriesComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblPinTries(),strPinTries,"Pin Tries")Then
				bstrPinTriesComparison = False
			End If
		End If
    End If
	strPinTriesComparison = bstrPinTriesComparison
End Function

'******* Function to compare Host Indicator
Public Function strHostIndicatorComparison(strHostIndicator)
	bstrHostIndicatorComparison = true
	If Not IsNull(strHostIndicator) Then
		If strHostIndicator = "RUNTIME" Then
			strHostIndicatorMF = trim(fGetText(strSession, "17","068","2"))
			strHostIndicatorFE = Claims.lblHostIndicator().GetRoProperty("innertext")
			
			If  Ucase(Trim(strHostIndicatorFE)) = UCase(Trim(strHostIndicatorMF)) Then
				LogMessage "RSLT", "Verification","For Host Indicator successfully matched with the expected value. Expected: " & strHostIndicatorMF &" , Actual: "& strHostIndicatorFE, True
				bstrHostIndicatorComparison = True
			else
				LogMessage "WARN", "Verification","For Host Indicator successfully matched with the expected value. Expected: " & strHostIndicatorMF &" , Actual: "& strHostIndicatorFE, False
				bstrHostIndicatorComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblHostIndicator(),strHostIndicator,"Host Indicator")Then
				bstrHostIndicatorComparison = False
			End If
		End If
    End If
    strHostIndicatorComparison = bstrHostIndicatorComparison
End Function

'******* Function to compare Completion Status
Public Function strCompletionStatusComparison(strCompletionStatus)
	bstrCompletionStatusComparison = true
	If Not IsNull(strCompletionStatus) Then
		If strCompletionStatus = "RUNTIME" Then
			strCompletionStatusMF = trim(fGetText(strSession, "15","068","10"))
			strCompletionStatusFE = Claims.lblCompletionStatus().GetRoProperty("innertext")
			
			If  Ucase(Trim(strCompletionStatusFE)) = UCase(Trim(strCompletionStatusMF)) Then
				LogMessage "RSLT", "Verification","For Completion Status successfully matched with the expected value. Expected: " & strCompletionStatusMF &" , Actual: "& strCompletionStatusFE, True
				bstrCompletionStatusComparison = True
			else
				LogMessage "WARN", "Verification","For Completion Status successfully matched with the expected value. Expected: " & strCompletionStatusMF &" , Actual: "& strCompletionStatusFE, False
				bstrCompletionStatusComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblCompletionStatus(),strCompletionStatus,"Completion Status")Then
				bstrCompletionStatusComparison = False
			End If
		End If
    End If
    strCompletionStatusComparison = bstrCompletionStatusComparison
End Function

'******* Function to compare ARQC
Public Function strARQCComparison(strARQC)
	bstrARQCComparison = true
	If Not IsNull(strLocation) Then
		If strLocation = "RUNTIME" Then
			strLocationMF = trim(fGetText(strSession, "17","009","12"))
			strLocationFE = Claims.lblLocation().GetRoProperty("innertext")
			
			If  Ucase(Trim(strLocationFE)) = UCase(Trim(strLocationMF)) Then
				LogMessage "RSLT", "Verification","For Location successfully matched with the expected value. Expected: " & strLocationMF &" , Actual: "& strLocationFE, True
				bstrLocationComparison = True
			else
				LogMessage "WARN", "Verification","For Location successfully matched with the expected value. Expected: " & strLocationMF &" , Actual: "& strLocationFE, False
				bstrLocationComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblLocation(),strLocation,"Location")Then
				bstrLocationComparison = False
			End If
		End If
    End If
    strLocationComparison = bstrLocationComparison
End Function

'******* Function to compare the Location
Public Function strLocationComparison(strLocation)
	bstrLocationComparison = true
	If Not IsNull(strLocation) Then
		If strLocation = "RUNTIME" Then
			strLocationMF = trim(fGetText(strSession, "17","009","12"))
			strLocationFE = Claims.lblLocation().GetRoProperty("innertext")
			
			If  Ucase(Trim(strLocationFE)) = UCase(Trim(strLocationMF)) Then
				LogMessage "RSLT", "Verification","For Location successfully matched with the expected value. Expected: " & strLocationMF &" , Actual: "& strLocationFE, True
				bstrLocationComparison = True
			else
				LogMessage "WARN", "Verification","For Location successfully matched with the expected value. Expected: " & strLocationMF &" , Actual: "& strLocationFE, False
				bstrLocationComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblLocation(),strLocation,"Location")Then
				bstrLocationComparison = False
			End If
		End If
    End If
    strLocationComparison = bstrLocationComparison
End Function

'******** Function to compare the Terminal Owner 
Public Function strTerminalOwnerComparison()
	bstrTerminalOwnerComparison = true
	If Not IsNull(strTerminalOwner) Then
		If strTerminalOwner = "RUNTIME" Then
			strTerminalOwnerMF = trim(fGetText(strSession, "15","009","10"))
			strTerminalOwnerFE = Claims.lblTerminalOwner().GetRoProperty("innertext")
			
			If  Ucase(Trim(strTerminalOwnerFE)) = UCase(Trim(strTerminalOwnerMF)) Then
				LogMessage "RSLT", "Verification","For Terminal Owner successfully matched with the expected value. Expected: " & strTerminalOwnerMF &" , Actual: "& strTerminalOwnerFE, True
				bstrTerminalOwnerComparison = True
			else
				LogMessage "WARN", "Verification","For Terminal Owner successfully matched with the expected value. Expected: " & strTerminalOwnerMF &" , Actual: "& strTerminalOwnerFE, False
				bstrTerminalOwnerComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lbLedgerBalance(),strTerminalOwner,"Terminal Owner")Then
				bstrTerminalOwnerComparison = False
			End If
		End If
    End If
    strTerminalOwnerComparison = bstrTerminalOwnerComparison
End Function

'******** Function to compare the Available Balance
Public Function strAvailableBalanceComparison(strAvailableBalance)
	bstrAvailableBalanceComparison = true
	If Not IsNull(strAvailableBalance) Then
		If strAvailableBalance = "RUNTIME" Then
			strAvailableBalanceMF = trim(fGetText(strSession, "12","060","18"))
			strAvailableBalanceMF = trim(Replace(strAvailableBalanceMF ,"$",""))
			strAvailableBalanceFE = Claims.lbAvailableBalance().GetRoProperty("innertext")
			
			If  Ucase(Trim(strAvailableBalanceFE)) = UCase(Trim(strAvailableBalanceMF)) Then
				LogMessage "RSLT", "Verification","For Available Balance successfully matched with the expected value. Expected: " & strAvailableBalanceMF &" , Actual: "& strAvailableBalanceFE, True
				bstrAvailableBalanceComparison = True
			else
				LogMessage "WARN", "Verification","For Available Balance successfully matched with the expected value. Expected: " & strAvailableBalanceMF &" , Actual: "& strAvailableBalanceFE, False
				bstrAvailableBalanceComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lbLedgerBalance(),strAvailableBalance,"Available Balance")Then
				bstrAvailableBalanceComparison = False
			End If
		End If
    End If
    strAvailableBalanceComparison = bstrAvailableBalanceComparison
End Function

'******* Function to compare the Ledger Balance
Public Function strLedgerBalanceComparison(strLedgerBalance)
	bstrLedgerBalanceComparison = true
	If Not IsNull(strLedgerBalance) Then
		If strLedgerBalance = "RUNTIME" Then
			strLedgerBalanceMF = trim(fGetText(strSession, "11","060","18"))
			strLedgerBalanceMF = trim(Replace(strLedgerBalanceMF,"$",""))
			strLedgerBalanceFE = Claims.lbLedgerBalance().GetRoProperty("innertext")
			
			If  Ucase(Trim(strLedgerBalanceFE)) = UCase(Trim(strTransactionTypeMF)) Then
				LogMessage "RSLT", "Verification","For Ledger Balance successfully matched with the expected value. Expected: " & strLedgerBalanceMF &" , Actual: "& strLedgerBalanceFE, True
				bstrLedgerBalanceComparison = True
			else
				LogMessage "WARN", "Verification","For Ledger Balance successfully matched with the expected value. Expected: " & strLedgerBalanceMF &" , Actual: "& strLedgerBalanceFE, False
				bstrLedgerBalanceComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lbLedgerBalance(),strLedgerBalance,"Transaction Type")Then
				bstrLedgerBalanceComparison = False
			End If
		End If
    End If
    strLedgerBalanceComparison = bstrLedgerBalanceComparison
End Function

'********* Function to compare the Transaction Type
Public Function strTransactionTypeComparison()
	bstrTransactionTypeComparison = true
	If Not IsNull(strTransactionType) Then
		If strTransactionType = "RUNTIME" Then
			strTransactionTypeMF = trim(fGetText(strSession, "18","028","8"))
			strTransactionTypeFE = Claims.lblTransactionType().GetRoProperty("innertext")
			
			If  Ucase(Trim(strTransactionTypeFE)) = UCase(Trim(strTransactionTypeMF)) Then
				LogMessage "RSLT", "Verification","For Transaction Type successfully matched with the expected value. Expected: "+ strTransactionTypeMF &" , Actual: "& strTransactionTypeFE, True
				bstrTransactionTypeComparison = True
			else
				LogMessage "WARN", "Verification","For Transaction Type successfully matched with the expected value. Expected: "+ strTransactionTypeMF &" , Actual: "& strTransactionTypeFE, False
				bstrTransactionTypeComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblTransactionType(),strTransactionType,"Transaction Type")Then
				bstrTransactionTypeComparison = False
			End If
		End If
    End If
End Function

'********* Function to compare the sub class
Public Function strSubClassComparison(strSubClass)
	bstrSubClassComparison = true
	If Not IsNull(strSubClass) Then
		If strSubClass = "RUNTIME" Then
			strSubClassMF = trim(fGetText(strSession, "09","035","6"))
			strSubClassFE = Claims.lblSubClass().GetRoProperty("innertext")
			
			If  Ucase(Trim(strSubClassFE)) = UCase(Trim(strSubClassMF)) Then
				LogMessage "RSLT", "Verification","For Sub Class successfully matched with the expected value. Expected: "+ strSubClassMF &" , Actual: "& strSubClassFE, True
				bstrSubClassComparison = True
			else
				LogMessage "WARN", "Verification","For Sub Class successfully matched with the expected value. Expected: "+ strSubClassMF &" , Actual: "& strSubClassFE, False
				bstrSubClassComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblSubClass(),strSubClass,"Sub Class")Then
				bcompareATMCAM = False
			End If
		End If
    End If
    strSubClassComparison = bstrSubClassComparison
End Function

'******* Function to compare the sequence no
Public Function strSequenceNoComparison()
	bstrSequenceNoComparison = true
	If Not IsNull(strSequenceNo) Then
		If strSequenceNo = "RUNTIME" Then
			strSequenceNoMF = trim(fGetText(strSession, "12","011","8"))
			strSequenceNoFE = Claims.lblSequenceNo().GetRoProperty("innertext")
			
			If  Ucase(Trim(strSequenceNoFE)) = UCase(Trim(strSequenceNoMF)) Then
				LogMessage "RSLT", "Verification","For Sequence No successfully matched with the expected value. Expected: "+ strSequenceNoMF &" , Actual: "& strSequenceNoFE, True
				bstrSequenceNoComparison = True
			else
				LogMessage "WARN", "Verification","For Sequence No successfully matched with the expected value. Expected: "+ strSequenceNoMF &" , Actual: "& strSequenceNoFE, False
				bstrSequenceNoComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblSequenceNo(),strSequenceNo,"Sequence No")Then
				bstrSequenceNoComparison = False
			End If
		End If
    End If
    strSequenceNoComparison = bstrSequenceNoComparison
End Function

'******* Function to compare the Business Date
Public Function strBusinessDateComparison(strBusinessDate)
	bstrBusinessDateComparison = true
	If Not IsNull(strBusinessDate) Then
		If strBusinessDate = "RUNTIME" Then
			strBusinessDateMF = trim(fGetText(strSession, "11","011","8"))
			'Convert the date from dd/mm/yyyy format to dd mon yyyy
			strBusinessDateMF = fConvertDate(strBusinessDateMF)
			strBusinessDateFE = Claims.lblBusinessDate().GetRoProperty("innertext")
			
			If  Ucase(Trim(strBusinessDateFE)) = UCase(Trim(strCardOwnerMF)) Then
				LogMessage "RSLT", "Verification","For Business Date successfully matched with the expected value. Expected: "+ strBusinessDateMF &" , Actual: "& strBusinessDateFE, True
				bstrBusinessDateComparison = True
			else
				LogMessage "WARN", "Verification","For Business Date successfully matched with the expected value. Expected: "+ strBusinessDateMF &" , Actual: "& strBusinessDateFE, False
				bstrBusinessDateComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblBusinessDate(),strBusinessDate,"Business Date")Then
				bstrBusinessDateComparison = False
			End If
		End If
    End If
    strBusinessDateComparison = bstrBusinessDateComparison
End Function

'******* Function to compare the card owner
Public Function strCardOwnerComparison(strCardOwner)
	bstrCardOwnerComparison = true
	If Not IsNull(strCardOwner) Then
		If strCardOwner = "RUNTIME" Then
			strCardOwnerMF = trim(fGetText(strSession, "05","014","7"))
			strCardOwnerFE = Claims.lblCardOwner().GetRoProperty("innertext")
			
			If  Ucase(Trim(strCardOwnerFE)) = UCase(Trim(strCardOwnerMF)) Then
				LogMessage "RSLT", "Verification","For Card Owner successfully matched with the expected value. Expected: "+ strCardOwnerMF &" , Actual: "& strCardOwnerFE, True
				bstrCardOwnerComparison = True
			else
				LogMessage "WARN", "Verification","For Card Owner successfully matched with the expected value. Expected: "+ strCardOwnerMF &" , Actual: "& strCardOwnerFE, False
				bstrCardOwnerComparison = False
			End If
		Else	
			If Not verifyInnerText(Claims.lblCardOwner(),strCardOwner,"Card Owner")Then
				bstrCardOwnerComparison = False
			End If
		End If
    End If
    strCardOwnerComparison = bstrCardOwnerComparison
End Function

'[Verify the values displayed as]
Public Function compareATMCAM(strCardOwner, strBusinessDate, strSequenceNo, strSubClass, strTransactionType, strLedgerBalance)
	bcompareATMCAM = true
	'read the record number from iServe
	strRecordNoFE = Claims.lblRecordNo().GetRoProperty("innertext")
	'Find the same Record no in Main Frame and start comparing
	For counter = 1 To 20 Step 1
		strRecordNoMF = fGetText(strSession, "03","041","5")
		If strRecordNoMF = strRecordNoFE Then
			'now start comparing the main frame and the FE
			'Compare the card owner with the FE
			If Not IsNull(strCardOwner) Then
				If strCardOwner = "RUNTIME" Then
					strCardOwnerMF = trim(fGetText(strSession, "05","014","7"))
					strCardOwnerFE = Claims.lblCardOwner().GetRoProperty("innertext")
					
					If  Ucase(Trim(strCardOwnerFE)) = UCase(Trim(strCardOwnerMF)) Then
						LogMessage "RSLT", "Verification","For Card Owner successfully matched with the expected value. Expected: "+ strCardOwnerMF &" , Actual: "& strCardOwnerFE, True
						bcompareATMCAM = True
					else
						LogMessage "WARN", "Verification","For Card Owner successfully matched with the expected value. Expected: "+ strCardOwnerMF &" , Actual: "& strCardOwnerFE, False
						bcompareATMCAM = False
					End If
				Else	
					If Not verifyInnerText(Claims.lblCardOwner(),strCardOwner,"Card Owner")Then
						bcompareATMCAM = False
					End If
				End If
		    End If
			'************* End of strCardOwner comparison
			
			'Compare the Business Date
			If Not IsNull(strBusinessDate) Then
				If strBusinessDate = "RUNTIME" Then
					strBusinessDateMF = trim(fGetText(strSession, "11","011","8"))
					'Convert the date from dd/mm/yyyy format to dd mon yyyy
					strBusinessDateMF = fConvertDate(strBusinessDateMF)
					strBusinessDateFE = Claims.lblBusinessDate().GetRoProperty("innertext")
					
					If  Ucase(Trim(strBusinessDateFE)) = UCase(Trim(strCardOwnerMF)) Then
						LogMessage "RSLT", "Verification","For Business Date successfully matched with the expected value. Expected: "+ strBusinessDateMF &" , Actual: "& strBusinessDateFE, True
						bcompareATMCAM = True
					else
						LogMessage "WARN", "Verification","For Business Date successfully matched with the expected value. Expected: "+ strBusinessDateMF &" , Actual: "& strBusinessDateFE, False
						bcompareATMCAM = False
					End If
				Else	
					If Not verifyInnerText(Claims.lblBusinessDate(),strBusinessDate,"Business Date")Then
						bcompareATMCAM = False
					End If
				End If
		    End If
			'************* End of Business Date comparison
			
			'Compare the Sequence No
			If Not IsNull(strSequenceNo) Then
				If strSequenceNo = "RUNTIME" Then
					strSequenceNoMF = trim(fGetText(strSession, "12","011","8"))
					strSequenceNoFE = Claims.lblSequenceNo().GetRoProperty("innertext")
					
					If  Ucase(Trim(strSequenceNoFE)) = UCase(Trim(strSequenceNoMF)) Then
						LogMessage "RSLT", "Verification","For Sequence No successfully matched with the expected value. Expected: "+ strSequenceNoMF &" , Actual: "& strSequenceNoFE, True
						bcompareATMCAM = True
					else
						LogMessage "WARN", "Verification","For Sequence No successfully matched with the expected value. Expected: "+ strSequenceNoMF &" , Actual: "& strSequenceNoFE, False
						bcompareATMCAM = False
					End If
				Else	
					If Not verifyInnerText(Claims.lblSequenceNo(),strSequenceNo,"Sequence No")Then
						bcompareATMCAM = False
					End If
				End If
		    End If
			'************* End of Sequence No comparison
			
			'Compare the Sub Class
			If Not IsNull(strSubClass) Then
				If strSubClass = "RUNTIME" Then
					strSubClassMF = trim(fGetText(strSession, "09","035","6"))
					strSubClassFE = Claims.lblSubClass().GetRoProperty("innertext")
					
					If  Ucase(Trim(strSubClassFE)) = UCase(Trim(strSubClassMF)) Then
						LogMessage "RSLT", "Verification","For Sub Class successfully matched with the expected value. Expected: "+ strSubClassMF &" , Actual: "& strSubClassFE, True
						bcompareATMCAM = True
					else
						LogMessage "WARN", "Verification","For Sub Class successfully matched with the expected value. Expected: "+ strSubClassMF &" , Actual: "& strSubClassFE, False
						bcompareATMCAM = False
					End If
				Else	
					If Not verifyInnerText(Claims.lblSubClass(),strSubClass,"Sub Class")Then
						bcompareATMCAM = False
					End If
				End If
		    End If
			'************* End of Sub Class comparison
			
			'Compare the Transaction Type
			If Not IsNull(strTransactionType) Then
				If strTransactionType = "RUNTIME" Then
					strTransactionTypeMF = trim(fGetText(strSession, "18","028","8"))
					strTransactionTypeFE = Claims.lblTransactionType().GetRoProperty("innertext")
					
					If  Ucase(Trim(strTransactionTypeFE)) = UCase(Trim(strTransactionTypeMF)) Then
						LogMessage "RSLT", "Verification","For Transaction Type successfully matched with the expected value. Expected: "+ strTransactionTypeMF &" , Actual: "& strTransactionTypeFE, True
						bcompareATMCAM = True
					else
						LogMessage "WARN", "Verification","For Transaction Type successfully matched with the expected value. Expected: "+ strTransactionTypeMF &" , Actual: "& strTransactionTypeFE, False
						bcompareATMCAM = False
					End If
				Else	
					If Not verifyInnerText(Claims.lblTransactionType(),strTransactionType,"Transaction Type")Then
						bcompareATMCAM = False
					End If
				End If
		    End If
			'************* End of Transaction Type comparison
			
			'Compare the Ledger Balance
			If Not IsNull(strLedgerBalance) Then
				If strLedgerBalance = "RUNTIME" Then
					strLedgerBalanceMF = trim(fGetText(strSession, "11","060","18"))
					strLedgerBalanceMF = trim(Replace(strLedgerBalanceMF,"$",""))
					strLedgerBalanceFE = Claims.lbLedgerBalance().GetRoProperty("innertext")
					
					If  Ucase(Trim(strLedgerBalanceFE)) = UCase(Trim(strTransactionTypeMF)) Then
						LogMessage "RSLT", "Verification","For Ledger Balance successfully matched with the expected value. Expected: " & strLedgerBalanceMF &" , Actual: "& strLedgerBalanceFE, True
						bcompareATMCAM = True
					else
						LogMessage "WARN", "Verification","For Ledger Balance successfully matched with the expected value. Expected: " & strLedgerBalanceMF &" , Actual: "& strLedgerBalanceFE, False
						bcompareATMCAM = False
					End If
				Else	
					If Not verifyInnerText(Claims.lbLedgerBalance(),strLedgerBalance,"Transaction Type")Then
						bcompareATMCAM = False
					End If
				End If
		    End If
			'************* End of Ledger Balance comparison
			
			bverify = compareATMCAM()
			Exit For
		else
			'Go to the previous screen
			fSendKeys "[pf7]"
			strRecordNoMF = fGetText(strSession, "03","041","5")
		End If
	Next
	
		
	strReferenceFE = Claims.lblReference().GetRoProperty("innertext")
	
	strAvailableBalanceFE = Claims.lbAvailableBalance().GetRoProperty("innertext")
	strTerminalOwner = Claims.lblTerminalOwner().GetRoProperty("innertext")
	
	
	'************ Extract the values from the Main frame
	
	
	'strReferenceMF = trim(fGetText(strSession, "11","011","8"))   'Only for Bill payments
	
	strAvailableBalanceMF = trim(fGetText(strSession, "12","060","18"))
	strAvailableBalanceMF = trim(Replace(strAvailableBalanceMF ,"$",""))	
End Function

'[Navigate to A543 in KRSP]
Public Function navigateToA543(strDay)
	bDevPending=false
	bnavigateToA543 = true
	KRLogin gstrUser_KRSP,gstrPassword_KRSP
	fSendKeys "[clear]" 
	'fWaitForInputReady (strSession)
	fSendKeys "A543"
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 11,036     'Position for the Day
	'***************** Check if the day is current or Previous
	Select Case (strDay)
		Case ("Current")
			fSendKeys "C"
		Case ("Previous")
			fSendKeys "P"
	End Select
	'fSendKeys "[enter]"
	wait 2
	fSetCursorPosition 16,037     'Enter the option
	fSendKeys "2" 'Enter Option No 2
	fSendKeys "[enter]"
	navigateToA543 = bnavigateToA543
End Function


