Dim strRunTimeTimeStamp:strRunTimeTimeStamp = " "
Dim strRunTimeDate:strRunTimeDate = " "
'To verify GIRO Setup
Dim strRunTimeBankAccountType:strRunTimeBankAccountType=""
Dim strRunTimeBankID:strRunTimeBankID=""
Dim strRunTimeAccount:strRunTimeAccount=""
Dim strRunTimeStatus_FTSP:strRunTimeStatus_FTSP=""
Dim strRunTimeRequestDay:strRunTimeRequestDay=""
Dim strRunTimePayment:strRunTimePayment=""
Dim strRunTimeNominalAmount:strRunTimeNominalAmount=""

'[Login to VisionPlus]
Public Function VPlusLogin(strClientID,strUserId,strPassword)
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
	'fSendKeys "T101"   ' CRM - Need to change the region for submit as T101 and for Pre Validation D101 default region G101
	fSendKeys gstrRegion_VPlus  ' CRM - Need to change the region for submit as T101 and for Pre Validation D101 default region G101
	fSendKeys "[enter]" 
	fWaitForAppAvailable (strSession)
	fWaitForInputReady (strSession) 
	wait 1
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 
	'*End of Region
	
	'*CIS Login
	fSendKeys "wssn" 
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 6,025 'UserId Location
	fSendKeys strClientID 'Enter UserID
	fSendKeys "[tab]"
	fSendKeys strUserId 'Enter UserID
	fSendKeys "[tab]"
	fSendKeys strPassword 'Enter Password
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	'Confirm Sign On successfull
	strActualStatus = fGetText(strSession, "02", "037", "7")
	If instr(strActualStatus,"WELCOME")=0 Then
		LogMessage "WARN","Verification","Login to V+ failed",false
		VPlusLogin=false
	 else
	LogMessage "RSLT","Verification","Login to V+ Successfull",True
		VPlusLogin=true
	End If
		wait 1
End Function

'[Unblock Account Level Card From VPlus ARMB Screen]
Public Function unblockCard_VPlus_ARMB(strCardNumber)
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	Wait 2	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "armb" 'Enter the screen code for unblock card
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 04,051 'Card Number location
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	Wait 1
	fWaitForInputReady (strSession)
	strInvalidMessage = fGetText(strSession, "09", "008", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		unblockCard_VPlus_ARME=true
		Exit Function
	End If
	fSetCursorPosition 15,018 ''Block Code 1
	fSendKeys " "	'Clear Block Code 1
	'fSendKeys "[tab]"

	fSetCursorPosition 15,024 ''Block Code 2
	fSendKeys " " 'Clear Block Code 2
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	'Delete  Special Handling file
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "ofsa" 'Enter the screen code to verify  unblock card
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 04,012 'Action
	fSendKeys "D"
	'fSendKeys "[tab]"  
	fSendKeys "702" 'Org Code
    fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 02,017 
	fSendKeys "Y" 'Delete
	'fSendKeys "[tab]"  
	fSendKeys "1"
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	strDeleteMessage = fGetText(strSession, "22", "002", "60") 
	LogMessage "RSLT","Verification","Card deletion status from Special Handling Fime is : "& strDeleteMessage,True
	 
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "arqb" 'Enter the screen code to verify  unblock card
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 04,051
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	' Get Block Code 1 & 2
	Wait 1
	strActualBlkCode1 = fGetText(strSession, "15", "018", "2")
	strActualBlkCode2 = fGetText(strSession, "15", "024", "2")
	If Trim(strActualBlkCode1)<>"" or Trim(strActualBlkCode2)<>""Then
		LogMessage "WARN","Verification","Failed to Unblock Card : "& strCardNumber &" from VPlus Screen ARMB",false
		unblockCard_VPlus_ARMB=false
	 else
	LogMessage "RSLT","Verification","Card : "& strCardNumber &" unblocked successfully from VPlus Screen ARMB",True
		unblockCard_VPlus_ARMB=true
	End If
End Function


'[Unblock Card Level Card From VPlus ARME Screen]
Public Function unblockCard_VPlus_ARME(strCardNumber)
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	Wait 2	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "arme" 'Enter the screen code for unblock card
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 07,033 'Card Number location
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	strInvalidMessage = fGetText(strSession, "13", "007", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		unblockCard_VPlus_ARME=true
		Exit Function
	End If
	fSetCursorPosition 01,063
    fSendKeys "03" 
	fSendKeys "[enter]"

	fWaitForInputReady (strSession)
	fSetCursorPosition 06,011 ''Block Code 1
	fSendKeys " "	'Clear Block Code 1
	'fSendKeys "[tab]"

	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	'Delete  Special Handling file
		'
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "ofsa" 'Enter the screen code to verify  unblock card
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 04,012 'Action
	fSendKeys "D"
	'fSendKeys "[tab]"  
	fSendKeys "702" 'Org Code
    fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 02,017 
	fSendKeys "Y" 'Delete
	'fSendKeys "[tab]"  
	fSendKeys "1"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	strDeleteMessage = fGetText(strSession, "22", "002", "60") 
	LogMessage "RSLT","Verification","Card deletion status from Special Handling Fime is : "& strDeleteMessage,True
	 
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "arqe" 'Enter the screen code to verify  unblock card
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
'	fSetCursorPosition 04,051
'	fSendKeys "[eraseeof]"
'    fSendKeys strCardNumber
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 01,063
    fSendKeys "03" 
	fSendKeys "[enter]"


	wait 1
	' Get Block Code 1 & 2
	strActualCardBlkCode1 = fGetText(strSession, "06", "011", "1")

	If Trim(strActualCardBlkCode1)<>"" Then
		LogMessage "WARN","Verification","Failed to Unblock Card : "& strCardNumber &" from VPlus Screen ARME",false
		unblockCard_VPlus_ARME=false
	 else
	LogMessage "RSLT","Verification","Card : "& strCardNumber &" unblocked successfully from VPlus Screen ARME",True
		unblockCard_VPlus_ARME=true
	End If
End Function

'[Change Card Block Code from VPlus ARME Screen]
Public Function changeBlockCode_VPlus_ARME(strCardNumber,strBlockCode)
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]"
	Wait 2
	fWaitForInputReady (strSession)
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "arme" 'Enter the screen code for unblock card
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 07,033 'Card Number location
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
		strInvalidMessage = fGetText(strSession, "13", "007", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		unblockCard_VPlus_ARME=true
		Exit Function
	End If

	fSetCursorPosition 01,063
    fSendKeys "03" 
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	If not isNull(strBlockCode) Then
		fSetCursorPosition 06,011 ''Block Code 1
	fSendKeys strBlockCode	'Clear Block Code 1
	else
		strBlockCode=""
	End If
	
	'fSendKeys "[tab]"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	'Delete  Special Handling file
		'
		 
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "arqe" 'Enter the screen code to verify  unblock card
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
    wait 1
	fSetCursorPosition 01,063
    fSendKeys "03" 
	fSendKeys "[enter]"

	wait 1
	
	fWaitForInputReady (strSession)
	'get card block code 
	strActualBlkCode1 = fGetText(strSession, "06", "011", "1")
	
	If Trim(strActualBlkCode1)<>Trim(strBlockCode) Then
		LogMessage "WARN","Verification","Failed to Change Card Block Code for Card : "& strCardNumber &" from VPlus Screen ARME. Actual Block Code: "&strActualBlkCode1,false
		changeBlockCode_VPlus_ARME=false
	 else
	LogMessage "RSLT","Verification","Card Block Code for Card : "& strCardNumber &" changed  successfully from VPlus Screen ARME",True
		changeBlockCode_VPlus_ARME=true
	End If
End Function

'[Verify Block Code And Reason Code From VPlus OFSA Screen]
Public Function verifyBlock_Reason_Code_OFSA(strCardNumber,strCardType,strBlockCode,strReasonCode)
   bVPlusVerify=true
   strCardNumber=Replace(strCardNumber,"-","")
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]" 
	wait 1
	fWaitForInputReady (strSession) 
	
	fSendKeys "ofsa" 'Enter the screen code for unblock card
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 04,012 'Action Code
'	fSendKeys strAction 'Enter Action Code A=Add;C=Change; D=Delete; I= Inquiry
	fSendKeys "I"
'	fSetCursorPosition 04,029 'Action Code
'	fSendKeys "702"
	fSendKeys "[tab]"
	'fSetCursorPosition 04,051 'Card Number
	fSendKeys strCardNumber ''Add Card Number
	If Instr(strCardType,"AMERICAN EXPRESS CREDIT CARD")<>0 Then
		fSetCursorPosition 09,051
		fSendKeys "Y"
		'fSendKeys "[tab]"
		fSetCursorPosition 15,051 
		fSendKeys "B"
	End If
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	strActualBlkCode = fGetText(strSession, "07", "044", "1")
	strActualReasonCode= fGetText(strSession, "11", "025", "1")
    strActualDateTime = fGetText(strSession, "09", "041", "16")
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
		 insertDataStore "DateAndTime_VPlus", strLastUpdatedDate
	End If
	
	If Trim(strActualBlkCode)= Trim(strBlockCode) Then
		LogMessage "RSLT","Verification","Actual Block Code : "& strActualBlkCode &" matched with Expected "&strBlockCode&" from VPlus Screen OFSA",True
     else
		 LogMessage "WARN","Verification","Actual Block Code : "& strActualBlkCode &" does not matched with Expected "&strBlockCode&" from VPlus Screen OFSA",false
		bVPlusVerify=false
		
	End If

	If Trim(strActualReasonCode)= Trim(strReasonCode) Then
		LogMessage "RSLT","Verification","Actual Reason Code : "& strActualReasonCode &" matched with Expected "&strReasonCode&" from VPlus Screen OFSA",True
     else
		LogMessage "WARN","Verification","Actual Reason Code : "& strActualReasonCode &" does not matched with Expected "&strReasonCode&" from VPlus Screen OFSA",false
		bVPlusVerify=false
	End If
	verifyBlock_Reason_Code_OFSA=bVPlusVerify

End Function

Public Function  unblockMultipleCards_VPlus()
 
	strBlockedCards=fetchFromDataStore("Verify row Data in Table Status Block Cards","BLANK","SuccessfullBlockedCards")(0)
	If  (isNull(strBlockedCards) or isEmpty(strBlockedCards)) Then
			strBlockedCards=fetchFromDataStore("Verify Link SRNumber available on Request Submitted popup","BLANK","SingleBlockedCard")(0) 'SingleBlocked Cards
	End If
	If  (isNull(strBlockedCards) or isEmpty(strBlockedCards)) Then
			strBlockedCards=fetchFromDataStore("Verify row Data in Table CardSummary on NewSR Screen","BLANK","CanceledCards")(0) 'BlockedCancelledCards
	End If
'strBlockedCards="4556210400230178|4556210400230186|"
'	strBlockedCards=fetchFromDataStore(strRunTimeDataStore_BlockedCardStep,"BLANK",strRunTimeDataStore_VarName)(0)
	If Not (isNull(strBlockedCards) or isEmpty(strBlockedCards)) Then
        lstCards=Split(strBlockedCards,"|")

		For iCount=0 to Ubound(lstCards)
			If  lstCards(iCount)<>"" Then
				strCardToUnblock=Split(lstCards(iCount),":")(0)
			'	strCardDesc=Split(lstCards(iCount),":")(1)
	
				unblockCard_VPlus_ARME(strCardToUnblock)
	
				unblockCard_VPlus_ARMB(strCardToUnblock)
			End If
		Next
	End If
	
End Function

'[Change Account Block Code from ARQB from VPlus]
Public Function ChangeAcc_BLK_CD_Vplus(strCardNumber,strBLKCD1,strBLKCD2)
		Dim bVplusRes:bVplusRes=true
		  VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
		strCardNumber=Replace(strCardNumber,"-","")
		fWaitForInputReady (strSession) 

		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSendKeys "[pf6]" 
		wait 1
		If strBLKCD1="" Then
			strBLKCD1=" "
		End If
		If strBLKCD2="" Then
			strBLKCD2=" "
		End If
		If not isNull(strBLKCD1) Then
			fSetCursorPosition 15,018
			fSendKeys strBLKCD1
		else
			strBLKCD1=""
		End If
		If not isNull(strBLKCD2) Then
			fSetCursorPosition 15,024
			fSendKeys strBLKCD2
		 else
			strBLKCD2=""
		End If
		fSendKeys "[enter]"
		wait 1

		fSetCursorPosition 01,009
		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]"

		fWaitForInputReady (strSession)
		strActualBlkCode1 = fGetText(strSession, "15", "018", "2")
		strActualBlkCode2 = fGetText(strSession, "15", "024", "2")
		If  Trim(strActualBlkCode1)=Trim(strBLKCD1) And Trim(strActualBlkCode2)=Trim(strBLKCD2) Then
			LogMessage "RSLT","Verification","Account Level Block Code for card "&strCardNumber&"changed successfully to "& strActualBlkCode1&","& strActualBlkCode2&"  from ARQB ",True

		else
			LogMessage "RSLT","Verification","Failed to change Account Level Block Code for card "&strCardNumber&" to "& strBLKCD1&","& strBLKCD2&"  from ARQB ",False
			bVplusRes=false
		End If
		ChangeAcc_BLK_CD_Vplus=bVplusRes
End Function
'[Unblock multiple cards from VPlus]
Public Function UnblockCards_VPlus(lstCardNumbers)
   VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
'	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 

	For iCount=0 to Ubound(lstCardNumbers)
		strCardNumber=lstCardNumbers(iCount)
		strCardNumber=Replace(strCardNumber,"-","")
		If strCardNumber<>"" Then
			fSendKeys "[clear]" 
			fWaitForInputReady (strSession)
			fSetCursorPosition 01,009 'UserId Location
				
			fSendKeys "arqe" ''Get Card level Block Status
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			fSetCursorPosition 04,033
			fSendKeys "[eraseeof]"
			fWaitForInputReady (strSession)
			fSetCursorPosition 05,033
			fSendKeys "[eraseeof]"
			fWaitForInputReady (strSession)
			fSetCursorPosition 07,033
			fSendKeys "[eraseeof]"
			fSendKeys strCardNumber
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			' Get Block Code 1 & 2
			fSetCursorPosition 01,063
			
			fSendKeys "03"
			fSendKeys "[enter]"
			fWaitForInputReady (strSession)
			wait 1 
			strActualBlkCode1 = fGetText(strSession, "06", "011", "1")
			LogMessage "RSLT","Verification","Card Level Block Code for Card "&strCardNumber&" from ARQE is  : "& strActualBlkCode1,True
			If Trim(strActualBlkCode1)<>"" Then
				'
				fSetCursorPosition 01,009 'UserId Location
				fSendKeys "arme" 'Enter the screen code for unblock card level
				fSendKeys "[enter]"
				fWaitForInputReady (strSession)
				fSetCursorPosition 07,033 'Card Number location
				fSendKeys "[eraseeof]"
				fSendKeys strCardNumber 'Enter Card Number
				fSendKeys "[enter]"
				fWaitForInputReady (strSession)
				strInvalidMessage = fGetText(strSession, "13", "007", "30") 
				If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")=0 Then
					
					'fSendKeys "[enter]"
					fSetCursorPosition 01,063
					fSendKeys "03"
					fSendKeys "[enter]"
					fWaitForInputReady (strSession)
					fSetCursorPosition 06,011 ''Block Code 1
					fSendKeys " "	'Clear Block Code 1
					'fSendKeys "[tab]"
				
					fSendKeys "[enter]" 
					fWaitForInputReady (strSession)
					wait 1
					'Delete  Special Handling file
						'
					fSetCursorPosition 01,009 'UserId Location
					fSendKeys "ofsa" 'Enter the screen code to verify  unblock card
					fSendKeys "[enter]" 
					fWaitForInputReady (strSession)
					wait 1
					fSetCursorPosition 04,012 'Action
					fSendKeys "D"
					'fSendKeys "[tab]"  
					fSendKeys "702" 'Org Code
					fSendKeys "[eraseeof]"
					fSendKeys strCardNumber 'Enter Card Number
					fSendKeys "[enter]"
					fWaitForInputReady (strSession)
					wait 1
					fSetCursorPosition 02,017 
					fSendKeys "Y" 'Delete
					'fSendKeys "[tab]"  
					fSendKeys "1"
					fSendKeys "[enter]"
					fWaitForInputReady (strSession)
					wait 1
					strDeleteMessage = fGetText(strSession, "22", "002", "60") 
					LogMessage "RSLT","Verification","Card deletion status from Special Handling Fime is : "& strDeleteMessage,True
				End If
			End If
		'**Account level Check
				
			fSendKeys "arqb" ''Get Card level Block Status
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			fSetCursorPosition 04,051
			fSendKeys "[eraseeof]"
			fSendKeys strCardNumber
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			' Get Block Code 1 & 2
			fSetCursorPosition 01,063
			
			fSendKeys "01"
			fSendKeys "[enter]"
			fWaitForInputReady (strSession)
			strActualBlkCode1 = fGetText(strSession, "15", "018", "2")
			strActualBlkCode2 = fGetText(strSession, "15", "024", "2")
			LogMessage "RSLT","Verification","Account Level Block Code for card "&strCardNumber&" from ARQB is  : "& strActualBlkCode1,True
			If Trim(strActualBlkCode1)<>"" or Trim(strActualBlkCode2)<>""Then
	
						'Unblock From ARMB
				fSetCursorPosition 01,009 'UserId Location
				fSendKeys "armb" 'Enter the screen code for unblock card
				fSendKeys "[enter]"
			
				fWaitForInputReady (strSession)
		
'				fSetCursorPosition 04,051 'Card Number location
'				fSendKeys "[eraseeof]"
'				fSendKeys strCardNumber 'Enter Card Number
'				fSendKeys "[enter]"
'				fWaitForInputReady (strSession)
				wait 1
				strInvalidMessage = fGetText(strSession, "09", "008", "30") 
				If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")=0 Then
					fSetCursorPosition 01,063
					fSendKeys "01"
					fSendKeys "[enter]"
					fSetCursorPosition 15,018 ''Block Code 1
					fSendKeys " "	'Clear Block Code 1
					'fSendKeys "[tab]"
				
					fSetCursorPosition 15,024 ''Block Code 2
					fSendKeys " " 'Clear Block Code 2
					fSendKeys "[enter]" 
					fWaitForInputReady (strSession)
					wait 1
				End If	
			End If			 
			'Delete  Special Handling file
						fSetCursorPosition 01,009 'UserId Location
					fSendKeys "ofsa" 'Enter the screen code to verify  unblock card
					fSendKeys "[enter]" 
					fWaitForInputReady (strSession)
					wait 1
'					fSetCursorPosition 04,012 'Action
'					fSendKeys "I"
'					'fSendKeys "[tab]"  
'					fSendKeys "702" 'Org Code
'					fSendKeys "[eraseeof]"
'					fSendKeys strCardNumber 'Enter Card Number
'					fSendKeys "[enter]"
'					fWaitForInputReady (strSession)
'					strBlockCodeOFSA = fGetText(strSession, "07", "044", "1") 
'					If Trim(strBlockCodeOFSA)<>"" Then
'						fSendKeys "[clear]" 
'						fWaitForInputReady (strSession)
'						wait 1
'						fSetCursorPosition 04,012 'Action
						fSendKeys "D"
						'fSendKeys "[tab]"  
						fSendKeys "702" 'Org Code
						fSendKeys "[eraseeof]"
						fSendKeys strCardNumber 'Enter Card Number
						fSendKeys "[enter]"
						fWaitForInputReady (strSession)
						wait 1
						fSetCursorPosition 02,017 
						fSendKeys "Y" 'Delete
						'fSendKeys "[tab]"  
						fSendKeys "1"
						fSendKeys "[enter]"
						wait 1
						fWaitForInputReady (strSession)
						strDeleteMessage = fGetText(strSession, "22", "002", "60") 
						LogMessage "RSLT","Verification","Card deletion status from Special Handling Fime is : "& strDeleteMessage,True
				'	End If
		End If
	Next	
End Function
Public Function getEmbossedDebitCardsDescription()
	arrEmbossedDebitDesc=Array(	"DBS MASTERCARD DEBIT",_ 
    "DBS-NUS MONEYSMART MASTERCARD",_ 
    "DBS CUP DEBIT CARD",_ 
	"DBS BUSINESS ADVANCE CARD",_ 
	"MULTI-ACCOUNT DEBIT CARD",_ 
    "MONEYSMART DEBIT MASTERCARD CONTEMP",_ 
	"DBS NUS DEBIT MASTERCARD",_ 
	"POSB MASTERCARD DEBIT",_ 
	"AFFINITY DEBIT MASTERCARD")

	getEmbossedDebitCardsDescription=arrEmbossedDebitDesc
End Function

'[Verify Main/Supplementary Indicator from ARQE]
Public Function VerifyCardHolderFlag__VPlus(strCardNumber,strExpectedCardHolderFlag)
   bVerifyCardHolderFlag__VPlus=False
   VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
'	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		'fSendKeys "[clear]" 
		'fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location
			
		fSendKeys "arqe" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,033
		fSendKeys "[eraseeof]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 05,033
		fSendKeys "[eraseeof]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 07,033
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1 
		strActualCardHolderFlag = fGetText(strSession, "16", "072", "1")
			LogMessage "RSLT","Verification","Card CardHolderFlag  for Card "&strCardNumber&" from ARQE is  : "& strActualCardHolderFlag,True
		If Trim(strActualCardHolderFlag )=Trim(strExpectedCardHolderFlag) Then
			LogMessage "RSLT","Verification","Card CardHolderFlag  for Card "&strCardNumber&" from ARQE matched with expected value : "& strExpectedCardHolderFlag,True
			bVerifyCardHolderFlag__VPlus=True
		else
			LogMessage "RSLT","Verification","Card CardHolderFlag for Card "&strCardNumber&" from ARQE does not matched. Actual : "&strActualCardHolderFlag&" Expected : "& strExpectedCardHolderFlag,False
			bVerifyCardHolderFlag__VPlus=False
		End If
		VerifyCardHolderFlag__VPlus=bVerifyCardHolderFlag__VPlus
End Function

'[Verify Card block Code from ARQE]
Public Function VerifyCardBlockCode__VPlus(strCardNumber,strExpectedCardBLKCode)
   bVerifyCardBlockCode__VPlus=False
   VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
'	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		fSendKeys "[clear]" 
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location
			
		fSendKeys "arqe" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,033
		fSendKeys "[eraseeof]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 05,033
		fSendKeys "[eraseeof]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 07,033
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		' Get Block Code 1 & 2
		fSetCursorPosition 01,063
		
		fSendKeys "03"
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		wait 1 
		strActualCardBlkCode = fGetText(strSession, "06", "011", "1")
			LogMessage "RSLT","Verification","Card Level Block Code for Card "&strCardNumber&" from ARQE is  : "& strActualBlkCode1,True
		If Trim(strActualCardBlkCode)=Trim(strExpectedCardBLKCode) Then
			LogMessage "RSLT","Verification","Card Level Block Code for Card "&strCardNumber&" from ARQE matched with expected value : "& strExpectedCardBLKCode,True
			bVerifyCardBlockCode__VPlus=True
		else
			LogMessage "RSLT","Verification","Card Level Block Code for Card "&strCardNumber&" from ARQE does not matched. Actual : "&strActualCardBlkCode&" Expected : "& strExpectedCardBLKCode,False
			bVerifyCardBlockCode__VPlus=False
		End If
		VerifyCardBlockCode__VPlus=bVerifyCardBlockCode__VPlus
End Function

'[Verify Account block Code from ARQB]
Public Function VerifyAccountBlockCode_ARME_VPlus(strCardNumber,strAccountBLKCode1,strAccountBLKCode2)
		bVerifyAccountBlockCode_ARME_VPlus=True
		VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
		fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		fSendKeys "[clear]"
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location

		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		' Get Block Code 1 & 2
		fSetCursorPosition 01,063
		
		fSendKeys "01"
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		strActualBlkCode1 = fGetText(strSession, "15", "018", "2")
		strActualBlkCode2 = fGetText(strSession, "15", "024", "2")
		LogMessage "RSLT","Verification","Account Level Block Code for card "&strCardNumber&" from ARQB is  : "& strActualBlkCode1,True
		If Trim(strActualBlkCode1)<> strAccountBLKCode1 Then
			LogMessage "RSLT","Verification","Account Block Code1 for card "&strCardNumber&" does not matched. Actual : "& strActualBlkCode1& " Expected: "&strAccountBLKCode1,False
			bVerifyAccountBlockCode_ARME_VPlus=False
		End If
		If Trim(strActualBlkCode2)<> strAccountBLKCode2 Then
			LogMessage "RSLT","Verification","Account Block Code2 for card "&strCardNumber&" does not matched. Actual : "& strActualBlkCode2& " Expected: "&strAccountBLKCode2,False
			bVerifyAccountBlockCode_ARME_VPlus=False
        End If
		If bVerifyAccountBlockCode_ARME_VPlus Then
			LogMessage "RSLT","Verification","Account Block Code1 & 2 for card "&strCardNumber&" matched. Acc BLK Code1 : "& strActualBlkCode1& ", BLK Code2: "&strActualBlkCode2,true
		End If
		VerifyAccountBlockCode_ARME_VPlus=bVerifyAccountBlockCode_ARME_VPlus
End Function

'[Verify Relationship block Code from ARGQ]
Public Function VerifyRelshioBlockCode_ARGQ_VPlus(strCardNumber,strRelationshipBLKCode)
		bVerifyRelshioBlockCode_ARGQ_VPlus=True
		VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
		fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		fSendKeys "[clear]"
		fSendKeys "[clear]" 		
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location

		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)

		fSetCursorPosition 01,009 'UserId Location
		fSendKeys "argq" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
        strActualRelationshipBlkCode1 = fGetText(strSession, "10", "033", "2")
		
		LogMessage "RSLT","Verification","Relationship Block Code for card "&strCardNumber&" from ARGQ is  : "& strActualRelationshipBlkCode1,True
		If Trim(strActualRelationshipBlkCode1)<> strRelationshipBLKCode Then
			LogMessage "RSLT","Verification","Relationship Block Code for card "&strCardNumber&" does not matched. Actual : "& strActualRelationshipBlkCode1& " Expected: "&strRelationshipBLKCode,False
			bVerifyRelshioBlockCode_ARGQ_VPlus=False
		
		else
			LogMessage "RSLT","Verification","Relationship Block Code for card "&strCardNumber&" matched with expected Relationship BLK Code : "&strActualRelationshipBlkCode1,true
		End If
		VerifyRelshioBlockCode_ARGQ_VPlus=bVerifyRelshioBlockCode_ARGQ_VPlus
End Function

'[Change Relationship block Code from ARGM]
Public Function ChangeRelshioBlockCode_ARGM_VPlus(strCardNumber,strRelationshipBLKCode)
		bVerifyRelshioBlockCode_ARGQ_VPlus=True
		If not isNull(strRelationshipBLKCode) Then
			If strRelationshipBLKCode="" Then
				strRelationshipBLKCode=" "
			End If
			VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
		'	fSendKeys "[clear]" 
			fWaitForInputReady (strSession) 
	
			strCardNumber=Replace(strCardNumber,"-","")
			fSendKeys "[clear]"
			fWaitForInputReady (strSession)
			Wait 2
			fSendKeys "[clear]"
			fSendKeys "[clear]"
			fWaitForInputReady (strSession)
			fSetCursorPosition 01,009 'UserId Location
	
			fSendKeys "arqb" ''Get Card level Block Status
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			fSetCursorPosition 04,051
			fSendKeys "[eraseeof]"
			fSendKeys strCardNumber
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
	
			fSetCursorPosition 01,009 'UserId Location
			fSendKeys "argm" ''Get Card level Block Status
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			fSetCursorPosition 10,033
			fSendKeys strRelationshipBLKCode
			fSendKeys "[enter]" 
			wait 1
			fSetCursorPosition 01,009 'UserId Location
			fSendKeys "argq" ''Get Card level Block Status
			fSendKeys "[enter]" 
			wait 1
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			strActualRelationshipBlkCode1 = fGetText(strSession, "10", "033", "2")
			
			LogMessage "RSLT","Verification","Relationship Block Code for card "&strCardNumber&" from ARGQ is  : "& strActualRelationshipBlkCode1,True
			If Trim(strActualRelationshipBlkCode1)<> Trim(strRelationshipBLKCode) Then
				LogMessage "RSLT","Verification","Failed to change Relationship Block Code for card "&strCardNumber,False
				bVerifyRelshioBlockCode_ARGQ_VPlus=False
			
			else
				LogMessage "RSLT","Verification","Relationship Block Code for card "&strCardNumber&" changed with expected Relationship BLK Code : "&strRelationshipBLKCode,true
			End If
		
		End If
		ChangeRelshioBlockCode_ARGM_VPlus=bVerifyRelshioBlockCode_ARGQ_VPlus
End Function
'[Change SMS Indicator flag from ARGM for CreditCard]
Public Function ChangeSMSIndicator_ARGM_VPlus_CC(strCardNumber,strSMSIndicator, StrCardType)
		bChangeSMSIndicator_ARGM_VPlus=True
		If not isNull(strSMSIndicator) Then
			If strSMSIndicator="" Then
				strSMSIndicator=" "
			End If
			VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
		'	fSendKeys "[clear]" 
			fWaitForInputReady (strSession) 
	
			strCardNumber=Replace(strCardNumber,"-","")
			fSendKeys "[clear]"
			fWaitForInputReady (strSession)
			Wait 2
			fSendKeys "[clear]"
			fSendKeys "[clear]"
			fWaitForInputReady (strSession)
			fSetCursorPosition 01,009 'UserId Location
	
			fSendKeys "arqb" ''Get Card level Block Status
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			fSetCursorPosition 04,051
			fSendKeys "[eraseeof]"
			fSendKeys strCardNumber
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
	
			fSetCursorPosition 01,009 'UserId Location
			fSendKeys "argm" ''Get Card level Block Status
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
			fSetCursorPosition 01,063
			fSendKeys "40"
			fSendKeys "[enter]"
			fWaitForInputReady (strSession)
			wait 1

			If StrCardType = "Primary" Then
				fSetCursorPosition 16,025
				fSendKeys strSMSIndicator
				fSendKeys "[enter]"	
				fWaitForInputReady (strSession)
				fSetCursorPosition 01,063
				fSendKeys "40"
				fSendKeys "[enter]"				
				fWaitForInputReady (strSession)
				strActualSMSIndicator = fGetText(strSession, "16", "025", "1")
			ElseIf StrCardType = "Supplementry" Then
				fSetCursorPosition 16,032
				fSendKeys strSMSIndicator
				fSendKeys "[enter]"		
				fWaitForInputReady (strSession)
				fSetCursorPosition 01,063
				fSendKeys "40"
				fSendKeys "[enter]"	
				fWaitForInputReady (strSession)
				strActualSMSIndicator = fGetText(strSession, "16", "032", "1")
			End If
				
			LogMessage "RSLT","Verification","SMS Indicator flag for card "&strCardNumber&" from ARGQ is: "&strActualSMSIndicator,True
			If Trim(strActualSMSIndicator)<> Trim(strSMSIndicator) Then
				LogMessage "RSLT","Verification","Failed to change SMS Indicator Flag for card "&strCardNumber,False
				bChangeSMSIndicator_ARGM_VPlus=False
			
			else
				LogMessage "RSLT","Verification","SMS Indicator flag for card "&strCardNumber&" changed with expected SMS Indicator value : "&strSMSIndicator,true
			End If
		
		End If
		ChangeSMSIndicator_ARGM_VPlus_CC=bChangeSMSIndicator_ARGM_VPlus
End Function

'[Change SMS Indicator flag from ARGM for DebitCard]
Public Function ChangeSMSIndicator_ARMB_VPlus_DC(strCardNumber,strSMSIndicator)
		Dim bVplusRes:bVplusRes=true
		  VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
		strCardNumber=Replace(strCardNumber,"-","")
		fWaitForInputReady (strSession) 

		fSendKeys "armb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 01,063
		fSendKeys "40"
		fSendKeys "[enter]"	
		
		fSetCursorPosition 12,015
		fSendKeys strSMSIndicator
		fSendKeys "[enter]"	
		
		fSetCursorPosition 01,063	
		fSendKeys "40"
		fSendKeys "[enter]"	
		fWaitForInputReady (strSession)
		strActualSMSIndicator = fGetText(strSession, "12", "015", "1")

		If  Trim(strActualSMSIndicator)=Trim(strSMSIndicator) Then
			LogMessage "RSLT","Verification","SMS Indicator for debit card "&strCardNumber&"changed successfully to "& strActualSMSIndicator&"from ARMB ",True
		else
		LogMessage "RSLT","Verification","SMS Indicator for debit card "&strCardNumber&" not changed successfully to "& strActualSMSIndicator&"from ARMB ",False
			bVplusRes=false
		End If
		ChangeSMSIndicator_ARMB_VPlus_DC=bVplusRes
End Function

'[Verify Address Line123 and Postal Code not Blank in VPlus screen ARQN]
Public Function VerifyAddLine_PostalCode_Vplus(strCardNumber)
		bVerifyAddLine_PostalCode_Vplus=True
		VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
		fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		fSendKeys "[clear]"
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location

		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location
		fSendKeys "arqn" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
        strAddLine1 = fGetText(strSession, "06", "017", "26")
		strAddLine2 = fGetText(strSession, "07", "017", "26")
		strAddLine3 = fGetText(strSession, "10", "017", "26")
		strPostalCode = fGetText(strSession, "15", "017", "6")

		If Trim(strAddLine1)="" OR Trim(strAddLine2)="" OR Trim(strAddLine3)="" Then
			bVerifyAddLine_PostalCode_Vplus=false
			LogMessage "RSLT","Verification","Either of Address Line 1, 2, 3 are not blank  in Vplus ARQN screen ",false
		End If
		If Trim(strPostalCode)=""Then
			bVerifyAddLine_PostalCode_Vplus=false
			LogMessage "RSLT","Verification","Postal Code is not blank in Vplus ARQN screen ",false
		End If
		If bVerifyAddLine_PostalCode_Vplus Then
			LogMessage "RSLT","Verification","Address Line 1, 2, 3 and Postal Code are not blank as expected in Vplus ARQN screen ",true
		End If
		VerifyAddLine_PostalCode_Vplus=bVerifyAddLine_PostalCode_Vplus
End Function

'[Change Need To Activate Flag from ARME03 from VPlus]
Public Function ChangeNeedToActivate_ARME03VPlus(strCardNumber,strActivate)
   bChangeNeedToActivate_ARME03VPlus=False
   VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
'	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 

	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession)
	fSetCursorPosition 01,009 'UserId Location
		
	fSendKeys "arme" ''Get Card level Block Status
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 04,033
	fSendKeys "[eraseeof]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 05,033
	fSendKeys "[eraseeof]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 07,033
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 01,063
	
	fSendKeys "03"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 07,036
	fSendKeys strActivate
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	LogMessage "RSLT","Verification","Need To Activate Flag changed to "&strActivate, true
	ChangeNeedToActivate_ARME03VPlus=true
End Function

'[Change Postal Code in VPlus screen ARMN]
Public Function ChangePostalCode_Vplus(strCardNumber,strPostalCode)
		bChangePostalCode_Vplus=True
		If isNull(strPostalCode) Then
			ChangePostalCode_Vplus=true
			Exit Function
		End If
		VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
			fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		fSendKeys "[clear]" 
		fSendKeys "[clear]"
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location
		
		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		
		fSendKeys "[pf5]" 
		fWaitForInputReady (strSession)
		wait 1

		fSetCursorPosition 15,017
		fSendKeys "[eraseeof]"
     	fSendKeys strPostalCode
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 01,063
		fSendKeys "[eraseeof]"
		fSendKeys "03"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		strNewPostalCode = fGetText(strSession, "15", "017", "10")

		If Trim(strNewPostalCode)<>strPostalCode Then
			bChangePostalCode_Vplus=false
			LogMessage "RSLT","Verification","Failed to Change Postal Code to "&strPostalCode&" From Vplus ARMN screen ",false
		End If
		If bVerifyAddLine_PostalCode_Vplus Then
			LogMessage "RSLT","Verification"," Postal Code changed successfully to "&strPostalCode&" From Vplus ARMN screen",true
		End If

		ChangePostalCode_Vplus=bChangePostalCode_Vplus
End Function
'[Change Address and Postal Code in VPlus screen ARMN]
Public Function ChangeAdd_PostalCode_Vplus(strCardNumber,strAddLine1,strAddLine2,strAddLine3,strAddLine4,strAddLine5,strPostalCode)
		bChangePostalCode_Vplus=True
		VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
		fWaitForInputReady (strSession) 

		strCardNumber=Replace(strCardNumber,"-","")
		fSendKeys "[clear]" 
		fWaitForInputReady (strSession)
		fSetCursorPosition 01,009 'UserId Location
		
		fSendKeys "arqb" ''Get Card level Block Status
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 04,051
		fSendKeys "[eraseeof]"
		fSendKeys strCardNumber
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		
		fSendKeys "[pf5]" 
		fWaitForInputReady (strSession)
		wait 1
        fWaitForInputReady (strSession)
		If not isNull(strAddLine1) Then
			If strAddLine1="" Then
				strAddLine1=" "
			End If
			fSetCursorPosition 06,017
			fSendKeys "[eraseeof]"
            fSendKeys strAddLine1
		End If
		If not isNull(strAddLine2) Then
			If strAddLine2="" Then
				strAddLine2=" "
			End If
			fSetCursorPosition 07,017
			fSendKeys "[eraseeof]"
			fSendKeys strAddLine2
		End If
		If not isNull(strAddLine3) Then
			If strAddLine3="" Then
				strAddLine3=" "
			End If
			fSetCursorPosition 10,017
			fSendKeys "[eraseeof]"
			fSendKeys strAddLine3
		End If
		If not isNull(strAddLine4) Then
			If strAddLine4="" Then
				strAddLine4=" "
			End If
			fSetCursorPosition 11,017
			fSendKeys "[eraseeof]"
			fSendKeys strAddLine4
		End If
		If not isNull(strAddLine5) Then
			If strAddLine5="" Then
				strAddLine5=" "
			End If
			fSetCursorPosition 12,017
			fSendKeys "[eraseeof]"
			fSendKeys strAddLine5
		End If
		If not isNull(strPostalCode) Then
			If strPostalCode="" Then
				strPostalCode=" "
			End If
			fSetCursorPosition 15,017
			fSendKeys "[eraseeof]"
			fSendKeys strPostalCode
		End If
        fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1
		fSetCursorPosition 01,063
		fSendKeys "[eraseeof]"
		fSendKeys "03"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1

		LogMessage "RSLT","Verification","Address and  Postal Code changed successfully  From Vplus ARMN screen",true
		
		ChangeAdd_PostalCode_Vplus=bChangePostalCode_Vplus
End Function

'To get curret Date and Time from Vplus.
'By Surendaran on 23rd April 2015

'[Get current Date and Time from VPlus]
Public Function VPlusLogin_DateTime()

   	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	strDate = fGetText(strSession, "01", "070", "10")
	strTime = fGetText(strSession, "02", "070", "5")
	
	If instr(strDate,"/")=0 Then
		LogMessage "WARN","Verification","Login to V+ and fetching date time failed",false
		VPlusLogin_DateTime=false
		Exit Function
	 else
	LogMessage "RSLT","Verification","Login to V+ and fetching date time Successfull",True
		VPlusLogin_DateTime=true
	End If
		wait 1
		
	 '*************** Capturing time stamp from Vision Plus to open Memo for this SR by Manish
	strRunTimeTimeStamp="Get current Date and Time from VPlus"
	
	'var=strDate
	If len(Day(CDate(strDate)))=1 Then
		strDay="0"&Day(CDate(strDate))
	else
		strDay=""&Day(CDate(strDate))
	End If
	
	strLastUpdatedDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
	
	'var_month=mid(strDate,4,2)
	'var_month_change=MonthName(var_month,True)
	'var_month_format=replace(var,var_month,var_month_change)
	strDate=replace(strLastUpdatedDate,"/"," ")
	
	'strDate= FormatDateTime(Now(),vbLongDate)
	'strTempTime=FormatDateTime(Now(),vbShortTime)
	
	strTempTime_Replace=Replace(strTime,":","-")

	'strTempTime_Replace = "21-26"
 	strTimeStamp=strDate&" "&strTempTime_Replace
	insertDataStore "TimeStamp", strTimeStamp

End Function


'To get curret Date from Vplus.
'By Poornima on 30rd April 2015

'[Get current Date from VPlus]
Public Function VPlusLogin_Date()

   	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	strDate = fGetText(strSession, "01", "070", "10")
	'strTime = fGetText(strSession, "02", "070", "5")
	
	If instr(strDate,"/")=0 Then
		LogMessage "WARN","Verification","Login to V+ and fetching date failed",false
		VPlusLogin_Date=false
		Exit Function
	 else
	LogMessage "RSLT","Verification","Login to V+ and fetching date Successfull",True
		VPlusLogin_Date=true
	End If
		wait 1
		
	 '*************** Capturing time stamp from Vision Plus to open Memo for this SR by Manish
	strRunTimeTimeStamp="Get current Date from VPlus"
	
	var=strDate
	
  	'strDate = "04/04/2017"
 	'Commenting below date conversion. As later we have to validate with current date in temp limit.
 	'By Manish on 18 April 16
' 	If len(Day(CDate(strDate)))=1 Then
'        strDay="0"&Day(CDate(strDate))
'    else
'        strDay=""&Day(CDate(strDate))
'    End If
'    strLastUpdatedDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
    
    'var_month=mid(strDate,4,2)
'	var_month_change=MonthName(var_month,True)
'	var_month_format=replace(var,var_month,var_month_change)
'	strDate=replace(var_month_format,"/"," ")
	
	'strDate= FormatDateTime(Now(),vbLongDate)
	'strTempTime=FormatDateTime(Now(),vbShortTime)
	'commenting by poornima
	'--------------------
	'strTempTime_Replace=Replace(strTime,":","-")
	'---------------------------------
	'strTempTime_Replace = "21-26"
 	'strTimeStamp=strLastUpdatedDate
 	strRunTimeDate=strDate
	'insertDataStore "Date_VPLUS", strTimeStamp	
End Function

'[Reset Amount from VPlus ARGM and ARMB Screen]
Public Function resetAmount_VPlus_ARGM_ARMB(strCardNumber,strReltAmount,strAcctAmount)
	bresetAmount_VPlus_ARGM_ARMB=true
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "armb" 'Enter ARMB as an entry point
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 04,051 'Card Number location
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
		strInvalidMessage = fGetText(strSession, "13", "007", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		bresetAmount_VPlus_ARGM_ARMB=true
		Exit Function
	End If
	
	'Goto ARGM to reset amount at Relationship level
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "argm" 'Enter the screen code to go Relationship level
	fSendKeys "[enter]"

	fWaitForInputReady (strSession)
	
	fSetCursorPosition 01,063
    fSendKeys "40" 'Enter Page number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	
	strPageNumber = fGetText(strSession, "01", "063", "02")
	If strPageNumber = "02" Then
		fSendKeys "[enter]"
		fWaitForInputReady (strSession)
		wait 1
	End If
	
	strPageNumber = fGetText(strSession, "01", "063", "02")
	If not strPageNumber = "40" Then
		LogMessage "WARN","Verification","Failed to redirect to Page Numebr 40. Actual Page Number is "& strPageNumber,false
		bresetAmount_VPlus_ARGM_ARMB=true
		wait 1
		Exit Function
	End If
	
	fSetCursorPosition 11,051 'Position where amount need to change
	fSendKeys strReltAmount	'Enter Amount
	'fSendKeys "[tab]"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	
	'Goto ARMB to reset amount at Account level
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "armb" 'Enter the screen code to go Relationship level
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
		
	fSetCursorPosition 01,063
    fSendKeys "40" 'Enter Page number
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition 05,027 'Position where amount need to change
	fSendKeys strAcctAmount	'Enter Amount
	'fSendKeys "[tab]"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1			
	If Err.Number<>0 Then
       bresetAmount_VPlus_ARGM_ARMB=false
            LogMessage "WARN","Verification","Failed to Reset Amount in ARGM and ARMB" ,false
       Exit Function
   End If
   resetAmount_VPlus_ARGM_ARMB=bresetAmount_VPlus_ARGM_ARMB
End Function

'[Reset Amount from VPlus ARME Screen]
Public Function resetAmount_VPlus_ARME(strCardNumber,strCardAmount)
	bresetAmount_VPlus_ARME=true
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "arme"
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 07,033 'Card Number location
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	wait 1
	fWaitForInputReady (strSession)
		strInvalidMessage = fGetText(strSession, "13", "007", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		bresetAmount_VPlus_ARME=true
		Exit Function
	End If
			
	fSetCursorPosition 01,063
    fSendKeys "41" 'Enter Page number where amount need to change
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
		
	strPageNumber = fGetText(strSession, "01", "063", "02")
	If not strPageNumber = "41" Then
		LogMessage "WARN","Verification","Failed to redirect to Page Numebr 41. Actual Page Number is "& strPageNumber,True
		bresetAmount_VPlus_ARME=true
		wait 1
		Exit Function
	End If
	
	fSetCursorPosition 06,028 'Position where amount need to change
	fSendKeys strCardAmount	'Enter Amount
	'fSendKeys "[tab]"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	wait 1
		
	If Err.Number<>0 Then
       bresetAmount_VPlus_ARME=false
       LogMessage "WARN","Verification","Failed to Reset Amount in ARGM and ARMB" ,false
       Exit Function
   End If
   resetAmount_VPlus_ARME=bresetAmount_VPlus_ARME
End Function

'[Close Percomm Application]
Public Function closePercomm()
	bclosePercomm=true
	Dim autECLOIAObj 'As Object

    Dim autECLPSObj 'As Object
    Dim autECLConnList 'As Object
    Set autECLPSObj = CreateObject("PCOMM.autECLPS")
    Set autECLConnList = CreateObject("PCOMM.autECLConnList")
	stopPerCommConnection()
		wait 1
     ' Initialize the connection
    autECLConnList.Refresh
    If autECLConnList.Count <> 0 Then
      For i = 1 To autECLConnList.Count
		strConName = autECLConnList(i).Name
		strConHandle = autECLConnList(i).Handle
		ClosePerComWindow(strConName)
	   Next
	End If 
End Function

'***********Verify Suspension Details from GFQC- Added by ***************************

Public Function getSuspensionDetails_GFQC_Vplus(strCardNumber)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
	strCardNumber = Replace(strCardNumber,"-","")
	
	fSetCursorPosition "01","009" 
	fSendKeys "gfqc" 'code for Suspension details screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "10","018" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	' List of values in Page 01 for 60Days Past Due
	StrExpPDSuspensionFlag = fGetText(strSession, "06", "021","1")
	Environment.Value("StrExpPDSuspensionFlag") = StrExpPDSuspensionFlag
	
	StrExpPDSuspensionDate = fGetText(strSession, "07", "021","10")
	StrExpPDSuspensionDate = fConvertDate(StrExpPDSuspensionDate)
	Environment.Value("StrExpPDSuspensionDate") = StrExpPDSuspensionDate
	
	StrExpPDUpliftDate= fGetText(strSession, "08", "021","10")
	StrExpPDUpliftDate = fConvertDate(StrExpPDUpliftDate)
	Environment.Value("StrExpPDUpliftDate") = StrExpPDUpliftDate
	
	StrExpPDOverrideFlag = fGetText(strSession, "14", "022","1")
	Environment.Value("StrExpPDOverrideFlag") = StrExpPDOverrideFlag
	
	StrExpPDStartDate= fGetText(strSession, "15", "022","10")
	StrExpPDStartDate  = Mid(StrExpPDStartDate,1,2)+"/"+Mid(StrExpPDStartDate,3,2)+"/"+Mid(StrExpPDStartDate,5,4)
	StrExpPDStartDate = fConvertDate(StrExpPDStartDate)
	Environment.Value("StrExpPDStartDate") = StrExpPDStartDate
	
	StrExpPDEndDate = fGetText(strSession, "16", "022","10")
	StrExpPDEndDate  = Mid(StrExpPDEndDate,1,2)+"/"+Mid(StrExpPDEndDate,3,2)+"/"+Mid(StrExpPDEndDate,5,4)
	StrExpPDEndDate = fConvertDate(StrExpPDEndDate)
	Environment.Value("StrExpPDEndDate") = StrExpPDEndDate
	
	' List of values in Page 01 for Reinstatement
	
	StrExpQualifiedFlag = fGetText(strSession, "10", "033","1")
	Environment.Value("StrExpQualifiedFlag") = StrExpQualifiedFlag
	
	StrExpQualifiedAttempts = fGetText(strSession, "11", "033","3")
	StrExpQualifiedAttempts = fTrimZero(StrExpQualifiedAttempts)
	Environment.Value("StrExpQualifiedAttempts") = StrExpQualifiedAttempts
	
	StrExpLastQualifiedDate = fGetText(strSession, "11", "037","10")
	StrExpLastQualifiedDate = fConvertDate(StrExpLastQualifiedDate)
	Environment.Value("StrExpLastQualifiedDate") = StrExpLastQualifiedDate
	
	StrExpNonQualifiedAttempts = fGetText(strSession, "12", "033","3")
	StrExpNonQualifiedAttempts = fTrimZero(StrExpNonQualifiedAttempts)
	Environment.Value("StrExpNonQualifiedAttempts") = StrExpNonQualifiedAttempts
	
	StrExpLastNonQualifiedDate = fGetText(strSession, "12", "037","10")
	StrExpLastNonQualifiedDate = fConvertDate(StrExpLastNonQualifiedDate)
	Environment.Value("StrExpLastNonQualifiedDate") = StrExpLastNonQualifiedDate
	
	fSetCursorPosition "01","064" 'Set location for Page Change 
	fSendKeys "2"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	'List of values in Page 02 for Balance Income 
	StrExpBIAccreditorIndicator = fGetText(strSession, "07", "071","1")
	Environment.Value("StrExpBIAccreditorIndicator") = StrExpBIAccreditorIndicator
	
	StrExpBISuspensionFlag = fGetText(strSession, "06", "019","1")
	Environment.Value("StrExpBISuspensionFlag") = StrExpBISuspensionFlag
	
	StrExpBISuspendedOn = fGetText(strSession, "06", "041","10")'
	StrExpBISuspendedOn = fConvertDate(StrExpBISuspendedOn)
	Environment.Value("StrExpBISuspendedOn") = StrExpBISuspendedOn
	
	StrExpBISuspensionDays = fGetText(strSession, "12", "047","3")
	'StrExpBISuspensionDays = abs(StrExpBISuspensionDays)
	StrExpBISuspensionDays = fTrimZero(StrExpBISuspensionDays)
	Environment.Value("StrExpBISuspensionDays") = StrExpBISuspensionDays
	
	StrExpBIUpliftedOn = fGetText(strSession, "06", "070","10")'
	StrExpBIUpliftedOn = fConvertDate(StrExpBIUpliftedOn) 
	Environment.Value("StrExpBIUpliftedOn") = StrExpBIUpliftedOn
	
	StrExpBIOverrideFlag= fGetText(strSession, "12", "021","1")
	Environment.Value("StrExpBIOverrideFlag") = StrExpBIOverrideFlag
	
	StrExpBIStartDate = fGetText(strSession, "13", "021","10")
	StrExpBIStartDate  = Mid(StrExpBIStartDate,1,2)+"/"+Mid(StrExpBIStartDate,3,2)+"/"+Mid(StrExpBIStartDate,5,4)
	StrExpBIStartDate = fConvertDate(StrExpBIStartDate)
	Environment.Value("StrExpBIStartDate") = StrExpBIStartDate
	
	StrExpBIEndDate = fGetText(strSession, "14", "021","10")
	StrExpBIEndDate  = Mid(StrExpBIEndDate,1,2)+"/"+Mid(StrExpBIEndDate,3,2)+"/"+Mid(StrExpBIEndDate,5,4)
	StrExpBIEndDate = fConvertDate(StrExpBIEndDate)
	Environment.Value("StrExpBIEndDate") = StrExpBIEndDate
		
	'List of values in Page 02 for Aggregated Balance
	
	StrExpAggregateBalance = fGetText(strSession, "09", "018","17")
	StrExpAggregateBalance = fTrimZero(StrExpAggregateBalance)
	Environment.Value("StrExpAggregateBalance") = StrExpAggregateBalance
	
	StrExpAGUpdatedOn = fGetText(strSession, "09", "047","10")
	StrExpAGUpdatedOn = fConvertDate(StrExpAGUpdatedOn)
	Environment.Value("StrExpAGUpdatedOn") = StrExpAGUpdatedOn
	
	StrExpStaffIndicator = fGetText(strSession, "07", "019","1")
	'StrExpStaffIndicator = cdbl(StrExpStaffIndicator)
	Environment.Value("StrExpStaffIndicator") = StrExpStaffIndicator
	
	'List of values in Page 02 for Income (Application)
	StrExpApplicationAI= fGetText(strSession, "10", "018","17")
	'StrExpApplicationAI = abs(StrExpApplicationAI)
	StrExpApplicationAI = fTrimZero(StrExpApplicationAI)
	Environment.Value("StrExpApplicationAI") = StrExpApplicationAI
	
	StrExpApplicationAIUpdateOn = fGetText(strSession, "10", "047","10")
	StrExpApplicationAIUpdateOn = fConvertDate(StrExpApplicationAIUpdateOn)
	Environment.Value("StrExpApplicationAIUpdateOn") = StrExpApplicationAIUpdateOn
	
	StrExpAUMIndicator = fGetText(strSession, "13", "047","1")
	Environment.Value("StrExpAUMIndicator") = StrExpAUMIndicator
	
	'List of values in Page 02 for Income (Salary Crediting)
	StrExpSCAnnualIncome = fGetText(strSession, "11", "018","17")
	'StrExpSCAnnualIncome = abs(StrExpSCAnnualIncome)
	StrExpSCAnnualIncome = fTrimZero(StrExpSCAnnualIncome)
	Environment.Value("StrExpSCAnnualIncome") = StrExpSCAnnualIncome
	
	StrExpSCAnnualIncomeUpdatedOn = fGetText(strSession, "11", "047","10")
	StrExpSCAnnualIncomeUpdatedOn = fConvertDate(StrExpSCAnnualIncomeUpdatedOn)
	Environment.Value("StrExpSCAnnualIncomeUpdatedOn") = StrExpSCAnnualIncomeUpdatedOn
	
	StrExpSCIncomeIndicator= fGetText(strSession, "07", "060","1")
	Environment.Value("StrExpSCIncomeIndicator") = StrExpSCIncomeIndicator
	
End Function

'***********Verify Account Level Pending Payments from ARIQ (Functionality : Instant Card Payments)- Added by  on 28March 2016 ***************************

Public Function getAccountPendingPayments_ARIQ_Vplus(strCardNumber)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
	strCardNumber = Replace(strCardNumber,"-","")
	
	fSetCursorPosition "01","009" 
	fSendKeys "ARIQ" 'code for Account Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "04","038" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrOrgnizationNumber= fGetText(strSession, "03", "006","3")
	Environment.Value("StrOrgnizationNumber") = Trim(StrOrgnizationNumber)
	
	StrRelationshipNumber= fGetText(strSession, "03", "051","19")
	Environment.Value("StrRelationshipNumber") = Trim(StrRelationshipNumber)
	
	StrBeginningBalance= FormatNumber((fGetText(strSession, "15", "031","10")),2)
	Environment.Value("StrBeginningBalance") = Trim(StrBeginningBalance)
	
	StrStmtBalance= FormatNumber((fGetText(strSession, "05", "031","10")),2)
	Environment.Value("StrStmtBalance") = Trim(StrStmtBalance)
	
	StrStmtTotalDue= FormatNumber((fGetText(strSession, "08", "031","10")),2)
	Environment.Value("StrStmtTotalDue") = Trim(StrStmtTotalDue)
	
	StrCurrentDueDate= fGetText(strSession, "08", "069","10")
	StrCurrentDueDate = fConvertDate(StrStmtDueDate)
	Environment.Value("StrCurrentDueDate") = StrCurrentDueDate

	fSetCursorPosition "01","063" 'Set location for Page Change 
	fSendKeys "40"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrExpAccMemoPayment= FormatNumber((fGetText(strSession, "05", "029","10")),2)
	Environment.Value("StrExpAccMemoPayment") = Trim(StrExpAccMemoPayment)
	
	
 End Function
 
 '***********Verify Relationship Level Pending Payments from ARIG (Functionality : Instant Card Payments)- Added by  on 28March 2016 ***************************
 
 Public Function getRelationshipPendingPayments_ARIG_Vplus(StrOrgNumber,StrRelNumber)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
		
	fSetCursorPosition "01","009" 
	fSendKeys "ARIG" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "04","026" 'Set location of OrganizationNumber
	fSendKeys StrOrgNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "05","026" 'Set location of RelationshipNumber
	fSendKeys StrRelNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1

	StrEndingStmtBalance= FormatNumber((fGetText(strSession, "16", "042","10")),2)
	Environment.Value("StrEndingStmtBalance") = Trim(StrEndingStmtBalance)
	
	StrTotalPaymentDue= FormatNumber((fGetText(strSession, "19", "042","10")),2)
	Environment.Value("StrTotalPaymentDue") = Trim(StrTotalPaymentDue)
	
	fSetCursorPosition "01","063" 'Set location for Page Change 
	fSendKeys "40"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrRelCurrentDueDate= fGetText(strSession, "22", "019","10")
	StrRelCurrentDueDate = fConvertDate(StrRelCurrentDueDate)
	Environment.Value("StrRelCurrentDueDate") = StrRelCurrentDueDate
	
	StrRelstatementBalance= FormatNumber((fGetText(strSession, "18", "030","10")),2)
	Environment.Value("StrRelstatementBalance") = Trim(StrRelstatementBalance)
	
	fSetCursorPosition "01","063" 'Set location for Page Change 
	fSendKeys "42"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrExpRelMemoPayments= FormatNumber((fGetText(strSession, "07", "029","10")),2)
	Environment.Value("StrExpRelMemoPayments") = Trim(StrExpRelMemoPayments)
	
 End Function
 
 '***************Verify the statement values in ARSD System - Added by  on 29thMarch2016**********************************
 
  Public Function getStatementHistory_ARSD_Vplus(strCardNumber)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
	strCardNumber = Replace(strCardNumber,"-","")
	
	fSetCursorPosition "01","009" 
	fSendKeys "ARSD" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "04","051" 'Set location of OrganizationNumber
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "01","063" 'Set location for Page Change 
	fSendKeys "05"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrEndBalance= FormatNumber((fGetText(strSession, "14", "031","10")),2)
	Environment.Value("StrEndBalance") = Trim(StrEndBalance)
	
	StrStatementDue= FormatNumber((fGetText(strSession, "13", "071","10")),2)
	Environment.Value("StrStatementDue") = Trim(StrStatementDue)
	
	StrStatementDueDate = fGetText(strSession, "07", "030","10")
	StrStatementDueDate = fConvertDate(StrStatementDueDate)
	Environment.Value("StrStatementDueDate") = StrStatementDueDate
 End Function
 
  '**********************Verify Payments details from ARQA (Functionality : Instant Card Payments)- Added by  on 29March 2016 ***************************
 
 Public Function getPaymentdetails_ARQA_Vplus(strCardNumber)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
		
	fSetCursorPosition "01","009" 
	fSendKeys "ARQA" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	fSetCursorPosition "04","040" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1	
	fSetCursorPosition "15","004" 
	fSendKeys "x" 'Move to Retail Plan
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	
	fSetCursorPosition "01","063" 'Set location for Page Change 
	fSendKeys "03"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrCreditBalance=FormatNumber((fGetText(strSession, "08", "046","10")),2)
	Environment.Value("StrCreditBalance") =Trim(StrCreditBalance)
	
	fSetCursorPosition "01","063" 'Set location for Page Change 
	fSendKeys "06"
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	StrPaymentReversal=FormatNumber((fGetText(strSession, "20", "071","10")),2)
	Environment.Value("StrPaymentReversal") = Trim(StrPaymentReversal)
	
 End Function
 
  
'***/ This function is used to convert the date format from dd/mm/yyyy to dd mon yyyy/***
Public Function fConvertDate(StrDate)
  StrReceivedDate = IsDate(StrDate)
	If StrReceivedDate = True Then
		If len(Day(CDate(strDate)))=1 Then
			strDay="0"&Day(CDate(strDate))
	    else
			strDay=""&Day(CDate(strDate))
	   	End If
	   strExpDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
	   strDate=replace(strExpDate,"/"," ")
	 Else 
	   strDate ="" 
	 End If
  fConvertDate = strDate
End Function

'***/ This function is used to remove the leading zeros from the numbers and to display only the absolute value/***
Public Function fTrimZero(TrimString)
	While Left(TrimString,1)= "0" and TrimString <> "0"
		TrimString = Right(TrimString,Len(TrimString)-1)
	Wend
	fTrimZero = TrimString
End Function

'[Remove EP Verification from ARGM]
Public Function RemoveEPVerification_ARGM_VPlus(strCardNumber)
	bRemoveEPVerification_ARGM_VPlus=True
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	'	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 
	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	Wait 2
	fSendKeys "[clear]"
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "arqb" ''Get Card level Block Status
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 04,051
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "argm" ''Get Card level Block Status
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 01,063
	fSendKeys "40"
	fSendKeys "[enter]"
	wait 1
	fSetCursorPosition 09,056
	fSendKeys "" 'to remove EP verification
	
	strEPVerify = fGetText(strSession, "09", "056", "1")

	If Trim(strEPVerify)<> "" Then
		LogMessage "RSLT","Verification","Failed to remove EP verify for card "&strCardNumber,False
		bRemoveEPVerification_ARGM_VPlus=False			
	else
		LogMessage "RSLT","Verification","EP Verification for card "&strCardNumber&" removed successfully",true
	End If
	RemoveEPVerification_ARGM_VPlus=bRemoveEPVerification_ARGM_VPlus
End Function

'[Verify EP Verification from ARGM]
Public Function VerifyEPVerification_ARGM_VPlus(strCardNumber,strCINStatus)
	bVerifyEPVerification_ARGM_VPlus=True
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus	
	fWaitForInputReady (strSession) 
	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	Wait 2
	fSendKeys "[clear]"
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "arqb" ''Get Card level Block Status
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 04,051
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	
	fSetCursorPosition 01,009 'UserId Location
	fSendKeys "argm" ''Get Card level Block Status
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	fSetCursorPosition 01,063
	fSendKeys "40"
	fSendKeys "[enter]"
	wait 1
	fSetCursorPosition 09,056
	strEPVerify = fGetText(strSession, "09", "056", "1")
	If strCINStatus = "Singapore" Then
		If Trim(strEPVerify)<> "" Then
			LogMessage "RSLT","Verification","Failed to update EP verify for card "&strCardNumber,False
			bVerifyEPVerification_ARGM_VPlus=False			
		else
			LogMessage "RSLT","Verification","EP Verification for card "&strCardNumber&" not updated successfully",true
		End If
	Else
		If Trim(strEPVerify)<> "Y" Then
			LogMessage "RSLT","Verification","Failed to update EP verify for card "&strCardNumber,False
			bVerifyEPVerification_ARGM_VPlus=False			
		else
			LogMessage "RSLT","Verification","EP Verification for card "&strCardNumber&" not updated successfully",true
		End If	
	End If
	
	VerifyEPVerification_ARGM_VPlus=bVerifyEPVerification_ARGM_VPlus
End Function

'[Reset GIRO Setup From VPlus ARMB Screen]
Public Function resetGIROSetup_ARMB(strCardNumber)
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	Wait 2	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "armb" 'Enter the screen code for unblock card
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 04,051 'Card Number location
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	Wait 1
	fWaitForInputReady (strSession)
	strInvalidMessage = fGetText(strSession, "09", "008", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		resetGIROSetup_ARMB=true
		Exit Function
	End If
	
	fSetCursorPosition 01,063 ''To Change Page Number
	fSendKeys "06" 'Enter Page Number
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	'Reset Bank Account Type
	fSetCursorPosition 10,038
	fSendKeys "000"
	
	'Reset Bank ID
	fSetCursorPosition 16,029
	fSendKeys "0000000000"
	
	'Reset Payment Start Date
	fSetCursorPosition 18,033
	fSendKeys "00000000"
	
	'Reset Account
	fSetCursorPosition 19,022
	fSendKeys "           "
	
	'Reset Status and Payment
	fSetCursorPosition 21,034
	fSendKeys " "
	fSetCursorPosition 21,040
	fSendKeys "0"
	
	'Reset INTERIM PAYMENTS
	fSetCursorPosition 12,077
	fSendKeys "0"
	
	'Reset REQUEST DAY
	fSetCursorPosition 19,076
	fSendKeys "00"
	
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	strPageNumber = fGetText(strSession, "01", "063", "2")	
	If Trim(strPageNumber)<>"07" Then
		LogMessage "WARN","Verification","Failed to Reset GIRO Setup for Card: "& strCardNumber &" from VPlus Screen ARMB",false
		resetGIROSetup_ARMB=false
	 else
	LogMessage "RSLT","Verification","Reset GIRO Setup successful for Card : "& strCardNumber &" from VPlus Screen ARMB",True
		resetGIROSetup_ARMB=true
	End If
End Function

'[Verify GIRO Setup From VPlus ARMB Screen]
Public Function verifyGIROSetup_ARMB(strCardNumber)
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]"
	fWaitForInputReady (strSession)
	Wait 2	
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "armb" 'Enter the screen code for unblock card
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 04,051 'Card Number location
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	Wait 1
	fWaitForInputReady (strSession)
	strInvalidMessage = fGetText(strSession, "09", "008", "30") 
	If InStr(strInvalidMessage,"INVALID ORGANIZATION NUMBER")<>0 Then
		LogMessage "WARN","Verification","Card Not Valid In VPlus : "& strInvalidMessage,True
		verifyGIROSetup_ARMB=true
		Exit Function
	End If
	
	fSetCursorPosition 01,063 ''To Change Page Number
	fSendKeys "06" 'Enter Page Number
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	strRunTimeBankAccountType=fGetText(strSession, "10", "038", "3")
	strRunTimeBankID=fGetText(strSession, "16", "038", "4")
	strRunTimeAccount=fGetText(strSession, "19", "022", "10")
	strRunTimeStatus_FTSP=fGetText(strSession, "21", "034", "1")
	strRunTimeRequestDay=fGetText(strSession, "19", "076", "2")
	strRunTimePayment=fGetText(strSession, "21", "040", "1")
	strRunTimeNominalAmount=fGetText(strSession, "22", "018", "1")
	
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	strPageNumber = fGetText(strSession, "01", "063", "2")	
	If Trim(strPageNumber)<>"07" Then
		LogMessage "WARN","Verification","Failed to Verify GIRO Setup for Card: "& strCardNumber &" from VPlus Screen ARMB",false
		verifyGIROSetup_ARMB=false
	 else
	LogMessage "RSLT","Verification","Verify GIRO Setup successful for Card : "& strCardNumber &" from VPlus Screen ARMB",True
		verifyGIROSetup_ARMB=true
	End If
End Function

'[If Direct Debit status is Not Approved then Change it to Approve]
Public Function changeDirectDebitStatus(strCardNumber)
	bchangeDirectDebitStatus=true
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fSendKeys "[clear]" 
	fWaitForInputReady (strSession) 	
	strCardNumber=Replace(strCardNumber,"-","")
	fSendKeys "armb"
	fSendKeys "[enter]"
	fWaitForInputReady (strSession)
	fSetCursorPosition 04,051 'Card Number location
	fSendKeys "[eraseeof]"
	fSendKeys strCardNumber 'Enter Card Number
	fSendKeys "[enter]"
	Wait 1
	fWaitForInputReady (strSession)
	fSetCursorPosition 01,063 ''To Change Page Number
	fSendKeys "06" 'Enter Page Number
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	strStatus=fGetText(strSession, "21", "034", "1")
	If strStatus <> "Y" Then
		fSetCursorPosition 21,034
		fSendKeys "Y"
		fSendKeys "[pf6]"
	End If
	strStatus=fGetText(strSession, "21", "034", "1")
	If strStatus = "Y" Then
		LogMessage "RSLT","Verification","Direct Debit status changed successfully for Card : "& strCardNumber &" from VPlus Screen ARMB",True
		bchangeDirectDebitStatus=true
	End If
	changeDirectDebitStatus=bchangeDirectDebitStatus
End Function

'**********************Verify OtherPlan details from ARQA (Functionality : Otherplan(R1602))- Added by  on 05May 2016 *************************** 
  
 Public Function getOtherDetails_CC_UL_ARQA_Vplus(strCardNumber,strSeqNumber,i,strProduct,iRow)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
		
	fSetCursorPosition "01","009" 
	fSendKeys "ARQA" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	strCardNumber=Replace(strCardNumber,"-","")	
	
	fSetCursorPosition "04","040" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	strSeqNumber = Right("000" & strSeqNumber, 3)
	
	If i > 2 Then  ' Moving to next page incase the record is greater then 3 
		fSendKeys "[pf6]"
	End If
	
	If iRow = 1 Then
		x1 = "12"
		x2 = "13"
		x3 = "14"
		fSetCursorPosition "12","004"
	ElseIf iRow = 2	Then
		x1 = "15"
		x2 = "16"
		x3 = "17"
		fSetCursorPosition "15","004" 
	Else
		x1 = "18"
		x2 = "19"
		x3 = "20"
		fSetCursorPosition "18","004" 
	End If
	
	strRecord = fGetText(strSession, x1, "009", "03") ' Sequence No of first record displayed 
	strPlan = Trim(fGetText(strSession, x2, "040", "05"))
	Environment.Value("strPlan") = strPlan
	strCurBalance = Trim(fGetText(strSession, x1, "025", "10"))
	Environment.Value("strCurBalance") = strCurBalance	
	strPlanDesc = Ucase(Trim(fGetText(strSession, x3, "040", "35")))
	Environment.Value("strPlanDesc") = strPlanDesc	

	fSendKeys "x" 
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1		
	
	If Ucase(strProduct) = "LOAN" Then
		StrCalcRate= fGetText(strSession, "10", "055","9")
		StrCalcRate =formatNumber(StrCalcRate*100,2)
		Environment.Value("StrCalcRate") =Trim(StrCalcRate)	
	
		StrBeginingDate = fGetText(strSession, "10", "067","10")
		StrBeginingDate = fConvertDate(StrBeginingDate)
		Environment.Value("StrBeginingDate") = StrBeginingDate	
		
		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "03"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1		
				
		StrYTDPaidPrincipal=FormatNumber((fGetText(strSession, "12", "045","12")),2)
		Environment.Value("StrYTDPaidPrincipal") =Trim(StrYTDPaidPrincipal)	

		StrYTDPaidInterest=FormatNumber((fGetText(strSession, "13", "045","12")),2)
		Environment.Value("StrYTDPaidInterest") =Trim(StrYTDPaidInterest)					
	
		StrLTDPaidPrincipal=FormatNumber((fGetText(strSession, "12", "068","12")),2)
		Environment.Value("StrLTDPaidPrincipal") =Trim(StrLTDPaidPrincipal)	

		StrLTDPaidInterest=FormatNumber((fGetText(strSession, "13", "068","12")),2)
		Environment.Value("StrLTDPaidInterest") =Trim(StrLTDPaidInterest)				
		
		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "05"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1		
		
		StrOpenDate = fConvertDate(fGetText(strSession, "04", "025","10"))
		Environment.Value("StrOpenDate") = StrOpenDate	
		
		StrAccuredInterest=FormatNumber((fGetText(strSession, "15", "062","10")),4)
		Environment.Value("StrAccuredInterest") =Trim(StrAccuredInterest)
		If Right(StrAccuredInterest,1) = 0 Then
			Environment.Value("StrAccuredInterest") =FormatNumber((StrAccuredInterest),3)
		End If		
	
		StrPerDiem=FormatNumber((fGetText(strSession, "19", "068","10")),4)
		Environment.Value("StrPerDiem") =Trim(StrPerDiem)		
		If Right(StrPerDiem,1) = 0 Then
			Environment.Value("StrPerDiem") =FormatNumber((StrPerDiem),3)
		End If	
		
		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "07"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1			
	' DEFFERED Information table
		StrDeferredInterestORIG= FormatNumber(fGetText(strSession, "06", "022","05"),0)
		Environment.Value("StrDeferredInterestORIG") = StrDeferredInterestORIG

		StrDeferredInsuranceORIG= FormatNumber(fGetText(strSession, "07", "022","05"),0)
		Environment.Value("StrDeferredInsuranceORIG") = StrDeferredInsuranceORIG

		StrDeferredBillingORIG= FormatNumber(fGetText(strSession, "08", "022","05"),0)
		Environment.Value("StrDeferredBillingORIG") = StrDeferredBillingORIG

		StrDeferredPaymentORIG= FormatNumber(fGetText(strSession, "09", "022","05"),0)
		Environment.Value("StrDeferredPaymentORIG") = StrDeferredPaymentORIG
		
		StrDeferredInterestPeriod= FormatNumber(fGetText(strSession, "06", "028","03"),0)
		Environment.Value("StrDeferredInterestPeriod") = StrDeferredInterestPeriod

		StrDeferredInsurancePeriod= FormatNumber(fGetText(strSession, "07", "028","03"),0)
		Environment.Value("StrDeferredInsurancePeriod") = StrDeferredInsurancePeriod

		StrDeferredBillingPeriod= FormatNumber(fGetText(strSession, "08", "028","03"),0)
		Environment.Value("StrDeferredBillingPeriod") = StrDeferredBillingPeriod

		StrDeferredPaymentPeriod= FormatNumber(fGetText(strSession, "09", "028","03"),0)
		Environment.Value("StrDeferredPaymentPeriod") = StrDeferredPaymentPeriod
		
		StrDeferredInterestREM= FormatNumber(fGetText(strSession, "06", "032","05"),0)
		Environment.Value("StrDeferredInterestREM") = StrDeferredInterestREM

		StrDeferredInsuranceREM= FormatNumber(fGetText(strSession, "07", "032","05"),0)
		Environment.Value("StrDeferredInsuranceREM") = StrDeferredInsuranceREM

		StrDeferredBillingREM= FormatNumber(fGetText(strSession, "08", "032","05"),0)
		Environment.Value("StrDeferredBillingREM") = StrDeferredBillingREM

		StrDeferredPaymentREM= FormatNumber(fGetText(strSession, "09", "032","05"),0)
		Environment.Value("StrDeferredPaymentREM") = StrDeferredPaymentREM
	
		If Instr(strPlanDesc,"CASH") = 1  Then
			Environment.Value("StrInitialTerm") = 0
			Environment.Value("StrCurrentTerm") = 0		
			Environment.Value("StrLoanAmount") = 0
			Environment.Value("StrPrincipalAmount") = 0
			Environment.Value("StrInterestAmount") = 0
			Environment.Value("StrFirstPaymentAmount") = 0
			Environment.Value("StrFinalPaymentAmount") = 0
			Environment.Value("StrFirstPaymentDate") = ""			
			Environment.Value("StrFinalPaymentDate") = ""		
			Environment.Value("StrTotalNoOfInstallments") = 0
			Environment.Value("StrRemainingTerm") = 0	
			Environment.Value("StrAnnualPercentageCode") = 0
			Environment.Value("StrTotalDisbursableAmount") = 0
		Else		
			fSetCursorPosition "01","063" 'Set location for Page Change 
			fSendKeys "11"
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1	
			
			StrInitialTerm= FormatNumber(fGetText(strSession, "09", "071","03"),0)
			Environment.Value("StrInitialTerm") = StrInitialTerm
			
			StrCurrentTerm= FormatNumber(fGetText(strSession, "10", "071","03"),0)
			Environment.Value("StrCurrentTerm") = StrCurrentTerm		
			
			StrRemainingTerm =FormatNumber(fGetText(strSession, "11", "071","03"),0)
			Environment.Value("StrRemainingTerm") = StrRemainingTerm	
			
			StrLoanAmount= FormatNumber((fGetText(strSession, "13", "023","12")),2)
			Environment.Value("StrLoanAmount") = StrLoanAmount
			
			StrFirstPaymentAmount= FormatNumber((fGetText(strSession, "08", "026","10")),2)
			Environment.Value("StrFirstPaymentAmount") = StrFirstPaymentAmount
			
			StrFinalPaymentAmount= FormatNumber((fGetText(strSession, "09", "026","10")),2)
			Environment.Value("StrFinalPaymentAmount") = StrFinalPaymentAmount
			
			StrFirstPaymentDate = fConvertDate(fGetText(strSession, "05", "064","10"))
			Environment.Value("StrFirstPaymentDate") = StrFirstPaymentDate			
		
			StrFinalPaymentDate = fConvertDate(fGetText(strSession, "06", "064","10"))
			Environment.Value("StrFinalPaymentDate") = StrFinalPaymentDate	
			
			StrPostingInd = fGetText(strSession, "04", "015","1")
			
			If StrPostingInd = 1 Then
				Environment.Value("StrTotalNoOfInstallments") = 0	
			Else
				fSetCursorPosition "01","063" 'Set location for Page Change 
				fSendKeys "10"
				fSendKeys "[enter]" 
				fWaitForInputReady (strSession)
				wait 1	
				
				StrTotalNoOfInstallments= FormatNumber(fGetText(strSession, "07", "019","03"),0)
				Environment.Value("StrTotalNoOfInstallments") = StrTotalNoOfInstallments
			End If
	
			fSetCursorPosition "01","063" 'Set location for Page Change 
			fSendKeys "12"
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1	
			StrPrincipalAmount= FormatNumber((fGetText(strSession, "13", "027","10")),2)
			Environment.Value("StrPrincipalAmount") = StrPrincipalAmount
	
			StrInterestAmount= FormatNumber((fGetText(strSession, "14", "027","10")),2)
			Environment.Value("StrInterestAmount") = StrInterestAmount
			
			StrAnnualPercentageCode= fGetText(strSession, "07", "027","10")
			StrAnnualPercentageCode =formatNumber(StrAnnualPercentageCode*100,2)
			Environment.Value("StrAnnualPercentageCode") = StrAnnualPercentageCode
	
			fSetCursorPosition "01","063" 'Set location for Page Change 
			fSendKeys "40"
			fSendKeys "[enter]" 
			fWaitForInputReady (strSession)
			wait 1
	
			StrTotalDisbursableAmount= fGetText(strSession, "05", "028","18")
			StrTotalDisbursableAmount =formatNumber(StrTotalDisbursableAmount/100,2)
			Environment.Value("StrTotalDisbursableAmount") = StrTotalDisbursableAmount
		End If	
	ElseIf Ucase(strProduct) = "CREDITCARD"  OR  Ucase(strProduct) = "CASHLINE"  Then
	
		StrBaseRate= fGetText(strSession, "10", "023","9")
		StrBaseRate =formatNumber(StrBaseRate*100,2)
		Environment.Value("StrBaseRate") =Trim(StrBaseRate)

		StrCalcRate= fGetText(strSession, "10", "055","9")
		StrCalcRate =formatNumber(StrCalcRate*100,2)
		Environment.Value("StrCalcRate") =Trim(StrCalcRate)
			
		StrBeginingDate = fGetText(strSession, "10", "067","10")
		StrBeginingDate = fConvertDate(StrBeginingDate)
		Environment.Value("StrBeginingDate") = StrBeginingDate	
		
		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "03"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1		
				
		StrYTDPaidPrincipal=FormatNumber((fGetText(strSession, "12", "045","10")),2)
		Environment.Value("StrYTDPaidPrincipal") =Trim(StrYTDPaidPrincipal)	

		StrYTDPaidInterest=FormatNumber((fGetText(strSession, "13", "045","10")),2)
		Environment.Value("StrYTDPaidInterest") =Trim(StrYTDPaidInterest)					
	
		StrLTDPaidPrincipal=FormatNumber((fGetText(strSession, "12", "068","10")),2)
		Environment.Value("StrLTDPaidPrincipal") =Trim(StrLTDPaidPrincipal)	

		StrLTDPaidInterest=FormatNumber((fGetText(strSession, "13", "068","10")),2)
		Environment.Value("StrLTDPaidInterest") =Trim(StrLTDPaidInterest)				

		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "04"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1

		StrServiceCharges=FormatNumber((fGetText(strSession, "10", "044","10")),2)
		Environment.Value("StrServiceCharges") =Trim(StrServiceCharges)				
	
		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "05"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1		
		
		StrOpenDate = fConvertDate(fGetText(strSession, "04", "025","10"))
		Environment.Value("StrOpenDate") = StrOpenDate	
		
		StrBalTransferMonthlyRemain=FormatNumber((fGetText(strSession, "19", "021","03")),0)
		Environment.Value("StrBalTransferMonthlyRemain") =Trim(StrBalTransferMonthlyRemain)	

		StrBalTransferExpDate = fGetText(strSession,"17","027","8")
		StrBalTransferExpDate  = Mid(StrBalTransferExpDate,1,2)+"/"+Mid(StrBalTransferExpDate,3,2)+"/"+Mid(StrBalTransferExpDate,5,4)
		StrBalTransferExpDate = fConvertDate(StrBalTransferExpDate)
		Environment.Value("StrBalTransferExpDate") = StrBalTransferExpDate	

		StrAccuredInterest=FormatNumber((fGetText(strSession, "15", "062","10")),4)
		Environment.Value("StrAccuredInterest") =Trim(StrAccuredInterest)
		If Right(StrAccuredInterest,1) = 0 Then
			Environment.Value("StrAccuredInterest") =FormatNumber((StrAccuredInterest),3)
		End If		
	
		StrPerDiem=FormatNumber((fGetText(strSession, "19", "068","10")),4)
		Environment.Value("StrPerDiem") =Trim(StrPerDiem)		
		If Right(StrPerDiem,1) = 0 Then
			Environment.Value("StrPerDiem") =FormatNumber((StrPerDiem),3)
		End If 	
	
		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "07"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1		
		
		StrNormalInterestBeginDate = fGetText(strSession, "06", "039","8")
		StrNormalInterestBeginDate  = Mid(StrNormalInterestBeginDate,1,2)+"/"+Mid(StrNormalInterestBeginDate,3,2)+"/"+Mid(StrNormalInterestBeginDate,5,4)
		StrNormalInterestBeginDate = fConvertDate(StrNormalInterestBeginDate)
		Environment.Value("StrNormalInterestBeginDate") = StrNormalInterestBeginDate				
	
	' DEFFERED Information table
		StrDeferredInterestORIG= FormatNumber(fGetText(strSession, "06", "022","05"),0)
		Environment.Value("StrDeferredInterestORIG") = StrDeferredInterestORIG

		StrDeferredInsuranceORIG= FormatNumber(fGetText(strSession, "07", "022","05"),0)
		Environment.Value("StrDeferredInsuranceORIG") = StrDeferredInsuranceORIG

		StrDeferredBillingORIG= FormatNumber(fGetText(strSession, "08", "022","05"),0)
		Environment.Value("StrDeferredBillingORIG") = StrDeferredBillingORIG

		StrDeferredPaymentORIG= FormatNumber(fGetText(strSession, "09", "022","05"),0)
		Environment.Value("StrDeferredPaymentORIG") = StrDeferredPaymentORIG
		
		StrDeferredInterestPeriod= FormatNumber(fGetText(strSession, "06", "028","03"),0)
		Environment.Value("StrDeferredInterestPeriod") = StrDeferredInterestPeriod

		StrDeferredInsurancePeriod= FormatNumber(fGetText(strSession, "07", "028","03"),0)
		Environment.Value("StrDeferredInsurancePeriod") = StrDeferredInsurancePeriod

		StrDeferredBillingPeriod= FormatNumber(fGetText(strSession, "08", "028","03"),0)
		Environment.Value("StrDeferredBillingPeriod") = StrDeferredBillingPeriod

		StrDeferredPaymentPeriod= FormatNumber(fGetText(strSession, "09", "028","03"),0)
		Environment.Value("StrDeferredPaymentPeriod") = StrDeferredPaymentPeriod
		
		StrDeferredInterestREM= FormatNumber(fGetText(strSession, "06", "032","05"),0)
		Environment.Value("StrDeferredInterestREM") = StrDeferredInterestREM

		StrDeferredInsuranceREM= FormatNumber(fGetText(strSession, "07", "032","05"),0)
		Environment.Value("StrDeferredInsuranceREM") = StrDeferredInsuranceREM

		StrDeferredBillingREM= FormatNumber(fGetText(strSession, "08", "032","05"),0)
		Environment.Value("StrDeferredBillingREM") = StrDeferredBillingREM

		StrDeferredPaymentREM= FormatNumber(fGetText(strSession, "09", "032","05"),0)
		Environment.Value("StrDeferredPaymentREM") = StrDeferredPaymentREM
	End If	
 End Function

'**********************Verify OtherPlan details from ARQA (Functionality : Otherplan(R1602))- Added by  on 06May 2016 *************************** 
'//This function is used to verify pagination 
Public Function GetPagination_ARQA_Vplus (strCardNumber,strSeqNumber,strProduct,iPage,iRow)
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
		
	fSetCursorPosition "01","009" 
	fSendKeys "ARQA" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	strCardNumber=Replace(strCardNumber,"-","")	
	fSetCursorPosition "04","040" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	strSeqNumber = Right("000" & strSeqNumber, 3)
	
	For i = 1 To iPage
		fSendKeys "[pf6]"
		wait 1
	Next
	
	If iRow = 1 Then
		x1 = "12"
		x2 = "13"
		x3 = "14"
	ElseIf iRow = 2	Then
		x1 = "15"
		x2 = "16"
		x3 = "17"
	Else
		x1 = "18"
		x2 = "19"
		x3 = "20"
	End If
	
	strRecord = fGetText(strSession, x1, "009", "03") ' Sequence No of first record displayed 
	strPlan = Trim(fGetText(strSession, x2, "040", "05"))
	Environment.Value("strPlan") = strPlan
	strCurBalance = Trim(fGetText(strSession, x1, "025", "10"))
	Environment.Value("strCurBalance") = strCurBalance	
	strPlanDesc = Ucase(Trim(fGetText(strSession, x3, "040", "35")))
	Environment.Value("strPlanDesc") = strPlanDesc	
End Function

 Public Function getOtherDetails_UL_ARVV_Vplus(strCardNumber,strSeqNumber,i)
	
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
	fSetCursorPosition "01","009" 
	fSendKeys "ARVV" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	strCardNumber=Replace(strCardNumber,"-","")	
	fSetCursorPosition "04","045" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	
	strSeqNumber = Right("000" & strSeqNumber, 3)
	If i > 2 Then  ' Moving to next page in V+ incase the record is greater then 3 
		fSendKeys "[F6]"
	End If
	strRecord1 = fGetText(strSession, "10", "017", "03") ' Sequence No of first record displayed 
	strRecord2 = fGetText(strSession, "13", "017", "03") ' Sequence No of second record displayed 
	strRecord3 = fGetText(strSession, "16", "017", "03") ' Sequence No of third record displayed 
	
	If Not IsNull(strRecord1) Then 
	   If strRecord1 = strSeqNumber Then
		strPlan = fGetText(strSession, "10", "011", "05")
			Environment.Value("strPlan") = strPlan	
		strPlanDesc = Ucase(Trim(fGetText(strSession, "12", "050", "30")))
			Environment.Value("strPlanDesc") = strPlanDesc	
		strCurBalance = FormatNumber(fGetText(strSession, "11", "060", "11"),2)
			Environment.Value("strCurBalance") = strCurBalance	
		strSettlementType = fGetText(strSession, "11", "030", "01")
			Environment.Value("strSettlementType") = strSettlementType		
		fSetCursorPosition "10","004"
	   End If
		
	ElseIf Not IsNull(strRecord2) Then
		If strRecord2 = strSeqNumber Then
			strPlan = Trim(fGetText(strSession, "13", "011", "05"))
				Environment.Value("strPlan") = strPlan	
			strPlanDesc = Ucase(Trim(fGetText(strSession, "15", "050", "30")))
				Environment.Value("strPlanDesc") = strPlanDesc	
			strCurBalance = FormatNumber(fGetText(strSession, "14", "062", "10"),2)
				Environment.Value("strCurBalance") = strCurBalance	
			strSettlementType = fGetText(strSession, "13", "030", "01")
				Environment.Value("strSettlementType") = strSettlementType		
			fSetCursorPosition "13","004" 
		End If 
		
	ElseIf Not IsNull(strRecord3) Then
		If strRecord3 = strSeqNumber Then
			strPlan = Trim(fGetText(strSession, "16", "011", "05"))
				Environment.Value("strPlan") = strPlan	
			strPlanDesc = Ucase(Trim(fGetText(strSession, "18", "050", "04")))
				Environment.Value("strPlanDesc") = strPlanDesc	
			strCurBalance = FormatNumber(fGetText(strSession, "17", "062", "10"),2)
				Environment.Value("strCurBalance") = strCurBalance	
			strSettlementType = fGetText(strSession, "16", "030", "01")
				Environment.Value("strSettlementType") = strSettlementType
			fSetCursorPosition "16","004" 
		End If
	End If
	
	If Not IsNull(strSettlementType) Then ' get the required information from the respective screen only if the settlement is set to "S" for the selected Plan '
		fSendKeys "P" 
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1			
	
		StrInitialInsurance= FormatNumber((fGetText(strSession, "17", "025","10")),2)
		Environment.Value("StrInitialInsurance") = StrInitialInsurance
		
		StrInitialTerm= fGetText(strSession, "08", "045","03")
		Environment.Value("StrInitialTerm") = StrInitialTerm			
		
		StrRemainingTerm= fGetText(strSession, "08", "050","03")
		Environment.Value("StrRemainingTerm") = StrRemainingTerm

		fSetCursorPosition "01","063" 'Set location for Page Change 
		fSendKeys "03"
		fSendKeys "[enter]" 
		fWaitForInputReady (strSession)
		wait 1	
'		
'		StrRequestQuoteType= fGetText(strSession, "05", "024","01")
'		Environment.Value("StrRequestQuoteType") = StrRequestQuoteType

		StrPOByStartDate = fConvertDate(fGetText(strSession, "06", "038","10"))
		Environment.Value("StrPOByStartDate") = StrPOByStartDate			
			
		StrPOByEndDate = fConvertDate(fGetText(strSession, "06", "064","10"))
		Environment.Value("StrPOByEndDate") = StrPOByEndDate	

	' Pay Off By start Date 
		StrCurrentOustandingBal_POStart= FormatNumber((fGetText(strSession, "07", "038","17")),2)
		Environment.Value("StrCurrentOustandingBal_POStart") = StrCurrentOustandingBal_POStart

		StrInterestRebate_POStart= FormatNumber((fGetText(strSession, "08", "038","17")),2)
		Environment.Value("StrInterestRebate_POStart") = StrInterestRebate_POStart
		
		StrInsuranceRebate_POStart= FormatNumber((fGetText(strSession, "09", "038","17")),2)
		Environment.Value("StrInsuranceRebate_POStart") = StrInsuranceRebate_POStart
		
		StrInterestPenalty_POStart= FormatNumber((fGetText(strSession, "10", "038","17")),2)
		Environment.Value("StrInterestPenalty_POStart") = StrInterestPenalty_POStart
		
		StrInsurancePenalty_POStart= FormatNumber((fGetText(strSession, "11", "038","17")),2)
		Environment.Value("StrInsurancePenalty_POStart") = StrInsurancePenalty_POStart

		StrTerminationFee_POStart= FormatNumber((fGetText(strSession, "14", "038","17")),2)
		Environment.Value("StrTerminationFee_POStart") = StrTerminationFee_POStart
		
		StrCashRebate_POStart= FormatNumber((fGetText(strSession, "15", "038","17")),2)
		Environment.Value("StrCashRebate_POStart") = StrCashRebate_POStart		
		
		StrProjectedInterest_POStart= FormatNumber((fGetText(strSession, "16", "038","17")),2)
		Environment.Value("StrProjectedInterest_POStart") = StrProjectedInterest_POStart			

		StrPenaltyInterestMonth_POStart= fGetText(strSession, "17", "051","03")
		Environment.Value("StrPenaltyInterestMonth_POStart") = StrPenaltyInterestMonth_POStart	
		
		StrWithoutPaymentDue_POStart= FormatNumber((fGetText(strSession, "18", "038","17")),2)
		Environment.Value("StrWithoutPaymentDue_POStart") = StrWithoutPaymentDue_POStart
			
		StrNetPaymentDue_POStart = fConvertDate(fGetText(strSession, "20", "038","17"))
		Environment.Value("StrNetPaymentDue_POStart") = StrNetPaymentDue_POStart
		
		' Pay Off By End Date 		
		StrCurrentOustandingBal_POEnd= FormatNumber((fGetText(strSession, "07", "063","17")),2)
		Environment.Value("StrCurrentOustandingBal_POEnd") = StrCurrentOustandingBal_POEnd

		StrInterestRebate_POEnd= FormatNumber((fGetText(strSession, "08", "063","17")),2)
		Environment.Value("StrInterestRebate_POEnd") = StrInterestRebate_POEnd
		
		StrInsuranceRebate_POEnd= FormatNumber((fGetText(strSession, "09", "063","17")),2)
		Environment.Value("StrInsuranceRebate_POEnd") = StrInsuranceRebate_POEnd
		
		StrInterestPenalty_POEnd= FormatNumber((fGetText(strSession, "10", "063","17")),2)
		Environment.Value("StrInterestPenalty_POEnd") = StrInterestPenalty_POEnd
		
		StrInsurancePenalty_POEnd= FormatNumber((fGetText(strSession, "11", "063","17")),2)
		Environment.Value("StrInsurancePenalty_POEnd") = StrInsurancePenalty_POEnd

		StrTerminationFee_POEnd= FormatNumber((fGetText(strSession, "14", "063","17")),2)
		Environment.Value("StrTerminationFee_POEnd") = StrTerminationFee_POEnd
		
		StrCashRebate_POEnd= FormatNumber((fGetText(strSession, "15", "063","17")),2)
		Environment.Value("StrCashRebate_POEnd") = StrCashRebate_POEnd		
		
		StrProjectedInterest_POEnd= FormatNumber((fGetText(strSession, "16", "063","17")),2)
		Environment.Value("StrProjectedInterest_POEnd") = StrProjectedInterest_POEnd			

		StrPenaltyInterestMonth_POEnd= fGetText(strSession, "17", "076","03")
		Environment.Value("StrPenaltyInterestMonth_POEnd") = StrPenaltyInterestMonth_POEnd	
		
		StrWithoutPaymentDue_POEnd= FormatNumber((fGetText(strSession, "18", "063","17")),2)
		Environment.Value("StrWithoutPaymentDue_POEnd") = StrWithoutPaymentDue_POEnd
			
		StrNetPaymentDue_POEnd = fConvertDate(fGetText(strSession, "20", "063","17"))
		Environment.Value("StrNetPaymentDue_POEnd") = StrNetPaymentDue_POEnd	
	End If
End Function 


'[Verify SMS Local Min Amount and SMS Foreign Min Amount for Transactions populated From VPlus ARQE Screen]
Public Function VerifySMSMinAmount_ARQE_Vplus(strCardNumber)
VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus
	fWaitForInputReady (strSession) 
	wait 1
		
	fSetCursorPosition "01","009" 
	fSendKeys "ARQE" 'code for Relationship Level PendingPayments screen
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1
	strCardNumber=Replace(strCardNumber,"-","")	
	fSetCursorPosition "07","033" 'Set location of CardNumber 
	fSendKeys strCardNumber  
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession)
	wait 1	
	
	fSetCursorPosition 01,063
	fSendKeys "42"
	fSendKeys "[enter]"
	wait 1
	
	strVPRetailLocalAmt = (fGetText(strSession, "15", "018","11")/100)
	Environment.Value("strVPRetailLocalAmt") = FormatNumber(strVPRetailLocalAmt,2)
	
	strVPCashLocalAmt = (fGetText(strSession, "16", "018","11")/100)
	Environment.Value("strVPCashLocalAmt") = FormatNumber(strVPCashLocalAmt,2)
	
	strVPECommerceLocalAmt = (fGetText(strSession, "17", "018","11")/100)
	Environment.Value("strVPECommerceLocalAmt") = FormatNumber(strVPECommerceLocalAmt,2)
	
	strVPRecurringLocalAmt = (fGetText(strSession, "18", "018","11")/100)
	Environment.Value("strVPRecurringLocalAmt") = FormatNumber(strVPRecurringLocalAmt,2)	
	
	strVPMailOrderLocalAmt = (fGetText(strSession, "19", "018","11")/100)
	Environment.Value("strVPMailOrderLocalAmt") = FormatNumber(strVPMailOrderLocalAmt,2)		
	
	strVPRetailForeignAmt = (fGetText(strSession, "15", "035","11")/100)
	Environment.Value("strVPRetailForeignAmt") = FormatNumber(strVPRetailForeignAmt,2)
	
	strVPCashForeignAmt = (fGetText(strSession, "16", "035","11")/100)
	Environment.Value("strVPCashForeignAmt") = FormatNumber(strVPCashForeignAmt,2)
	
	strVPECommerceForeignAmt = (fGetText(strSession, "17", "035","11")/100)
	Environment.Value("strVPECommerceForeignAmt") = FormatNumber(strVPECommerceForeignAmt,2)
	
	strVPRecurringForeignAmt = (fGetText(strSession, "18", "035","11")/100)
	Environment.Value("strVPRecurringForeignAmt") = FormatNumber(strVPRecurringForeignAmt,2)	
	
	strVPMailOrderForeignAmt = (fGetText(strSession, "19", "035","11")/100)
	Environment.Value("strVPMailOrderForeignAmt") = FormatNumber(strVPMailOrderForeignAmt,2)
	
End Function

'***/ This function is used to convert the date format from dd/mm/yyyy to dd mon yyyy/***
Public Function fConvertDate(StrDate)
  StrReceivedDate = IsDate(StrDate)
	If StrReceivedDate = True Then
		If len(Day(CDate(strDate)))=1 Then
			strDay="0"&Day(CDate(strDate))
	    else
			strDay=""&Day(CDate(strDate))
	   	End If
	   strExpDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
	   strDate=replace(strExpDate,"/"," ")
	 Else 
	   strDate ="" 
	 End If
  fConvertDate = strDate
End Function

'***/ This function is used to remove the leading zeros from the numbers and to display only the absolute value/ Added by ***
Public Function fTrimZero(TrimString)
	While Left(TrimString,1)= "0" and TrimString <> "0"
		TrimString = Right(TrimString,Len(TrimString)-1)
	Wend
	fTrimZero = TrimString
End Function

'*** Funciton to check in V+; returns true(for validation message to exist) if the rule fails
Public Function verifyPopUpExist_VPlusValidation(strCardNumber)
	'The below rule is applicable only for Act Blk Code2 as N;
	
	'Remove the hyphen in Card Number
	strCardNumber = replace(strCardNumber,"-","")
	VPlusLogin gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus	
	fWaitForInputReady (strSession) 
	wait 1
	fSetCursorPosition 01,009	
	fSendKeys "ARPH" 
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession) 
	wait 1
	fSetCursorPosition 05,034
	fSendKeys strCardNumber
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession) 
	wait 1
	strARPHEffDate = cDate(fGetText(strSession, "07", "008", "10"))
	strARPHAmtPosted = trim(fGetText(strSession, "07", "023", "18"))
	strARPHAmtPosted = Replace(strARPHAmtPosted,",","")
	fSetCursorPosition 01,009	
	fSendKeys "ARSD" 
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession) 
	wait 1
	'If "NO STATEMENTS ON FILE FOR THIS ACCOUNT", then return true
	strChkPt = trim(fGetText(strSession, "06", "006", "46"))
	If instr(1,strChkPt,"NO STATEMENTS ON FILE FOR THIS ACCOUNT") Then
		'Rule passes; pop up should not come
		verifyPopUpExist_VPlusValidation = false
		Exit Function
	End If
	
	fSetCursorPosition 07,005
	strARSDLastStDate = cDate(fGetText(strSession, "07", "005", "10"))
	fSetCursorPosition 01,009
	fSendKeys "ARIQ" 
	fSendKeys "[enter]" 
	fWaitForInputReady (strSession) 
	wait 1
	strARIQTotAmtDue = trim(fGetText(strSession, "08", "027", "13"))
	strARIQTotAmtDue = replace(strARIQTotAmtDue,",","")
	
	'If atleast one of the [ARPH].[EFF DATE] (payment made date) > [ARSD].[DD/MM/YYYY] (last statement date) 
	'and Sum of [ARPH pg.1].[AMOUNT POSTED]  - [ARIQ pg.1].[TOT AMT DUE] != (Nagative OR 0)
	If (strARPHEffDate > strARSDLastStDate) OR (strARPHAmtPosted > strARSDLastStDate) Then
		'rule passes;popup should not come
		verifyPopUpExist_VPlusValidation = false 'The validation pop up should not come
	else
		verifyPopUpExist_VPlusValidation = true
	End If
End Function
