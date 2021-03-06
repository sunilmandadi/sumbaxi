'[Click on Voice Biometrics under Banking Facilities in Overview Page]
Public Function clicklinkVioBioMetric_VB()
clicklinkVioBioMetric_VB=true
VoiceBio.lnkVioBiometric.click	
If Err.Number<>0 Then
   clicklinkVioBioMetric_VB=false
   LogMessage "WARN","Verification","Failed to Click Link  :Voice Biometrics" ,false
   Exit Function
End If
waitForIcallLoading	
End Function

'[Verify field Enrolment status in Voice Biometrics Page displayed as]
Public Function verifyfieldEnrolmentStatus_VB(strEnrolStatus)
bverifyfieldEnrolmentStatus=true
If Not IsNull(strEnrolStatus) Then
	If Not VerifyInnerText(VoiceBio.lblEnrolmentStatus(), strEnrolStatus, "Enrolment status") Then
	   bverifyfieldEnrolmentStatus=false
	End If
End If
verifyfieldEnrolmentStatus_VB = bverifyfieldEnrolmentStatus
End Function

'[Verify field First voice print captured in Voice Biometrics Page displayed as]
Public Function VerifyfieldFirstVoicePrint_VB(strFirstVoicePrintDate)
bVerifyfieldFirstVoicePrint=true
If Not IsNull(strFirstVoicePrintDate) Then
	If Not VerifyInnerText(VoiceBio.lblFirstVoicePrintCaptured(), strFirstVoicePrintDate, "first voice print captured") Then
	   bVerifyfieldFirstVoicePrint=false
	End If
End If
verifyfieldFirstVoicePrint_VB = bVerifyfieldFirstVoicePrint
End Function

'[Verify field second voice print captured in Voice Biometrics Page displayed as]
Public Function VerifyfieldSecondVoicePrint_VB(strSecondVoicePrintDate)
bVerifyfieldSecondVoicePrint=true
If Not IsNull(strSecondVoicePrintDate) Then
	If Not VerifyInnerText(VoiceBio.lblSecondVoicePrintCaptured(), strSecondVoicePrintDate, "second voice print captured") Then
	   bVerifyfieldSecondVoicePrint=false
	End If
End If
VerifyfieldSecondVoicePrint_VB = bVerifyfieldSecondVoicePrint
End Function

'[Verify field Last authentication date in Voice Biometrics Page displayed as]
Public Function VerifyLastAuthenticatedDate_VB(strLastAuthenticateDate)
bVerifyLastAuthenticatedDate=true
If Not IsNull(strLastAuthenticateDate) Then
	If Not VerifyInnerText(VoiceBio.lblLastAuthenticationDate(), strLastAuthenticateDate, "Last authentication date") Then
	   bVerifyLastAuthenticatedDate=false
	End If
End If
VerifyLastAuthenticatedDate_VB = bVerifyLastAuthenticatedDate
End Function

'[Verify field Opt-out Status in Voice Biometrics Page displayed as]
Public Function VerifyfieldOptOutStatus_VB(strOptOutStatus)
bVerifyfieldOptOutStatus=true
If Not IsNull(strOptOutStatus) Then
   If Not VerifyInnerText(VoiceBio.lblOptOutStatus(), strOptOutStatus, "Opt-out Status") Then
	   bVerifyfieldOptOutStatus=false
   End If
End If
VerifyfieldOptOutStatus_VB = bVerifyfieldOptOutStatus
End Function

'[Verify field Last Opt-out date in Voice Biometrics Page displayed as]
Public Function VerifyfieldOptOutDate_VB(strOptOutDate)
bVerifyfieldOptOutdate = true
If Not IsNull(strOptOutDate) Then
   If Not VerifyInnerText(VoiceBio.lblLastOptOutDate(), strOptOutDate, "Last Opt-out date") Then
	  bVerifyfieldOptOutdate = false
   End IF
End If
VerifyfieldOptOutDate_VB = bVerifyfieldOptOutdate
End Function

'[Verify From and To dates range in Voice Biometrics Page displayed]
Public Function verifyDefaultDateRange_VB()
    bverifyDefaultDateRange=True
	strActFromDate = VoiceBio.txtFromDate.GetROProperty("value")
	strActToDate = VoiceBio.txtToDate.GetROProperty("value")
	DaysRange = DateDiff("d", strActFromDate, strActToDate)
	If DaysRange = 180  Then
	   LogMessage "RSLT","Verification","From and To Dates are displayed within 30 days range" ,True
	   bverifyDefaultDateRange = True 
	Else 
	   bverifyDefaultDateRange = False 
	End If
   verifyDefaultDateRange_VB = bverifyDefaultDateRange
End Function

'[Verify GO Button enabled in Voice Biometrics Page displayed]
Public Function VerifyButtonGO_VB()
   	bVerifyButtonGO = true
   	intBtnFilter = Instr(VoiceBio.btnGO.GetROproperty("outerhtml"),("v-disabled"))
	If  intBtnFilter=0 Then
		LogMessage "RSLT","Verification","GO button is enabled as expected.",True
		bVerifyButtonGO = true
	Else
		LogMessage "WARN","Verifiation","GO button is disabled.",false
		bVerifyButtonGO = false
	End If
	VerifyButtonGO_VB = bVerifyButtonGO
End Function

'[Click Button GO in Voice Biometrics Page]
Public Function clickButtonGO_VB()
   VoiceBio.btnGO.click 10,10,0 
   If Err.Number<>0 Then
       clickButtonGO_VB=false
            LogMessage "WARN","Verification","Failed to Click Button : Filter" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonGO_VB= True
End Function

'[Verify history table content in Voice Biometrics Page displayed as]
Public Function VerifyHistorydetails_VB(lstHistorytable)
   bVerifyHistorydetails = verifyTableContentList(VoiceBio.tblHistoryHeader,VoiceBio.tblHistoryContent,lstHistorytable,"history table content",false,null,null,null)
   VerifyHistorydetails_VB = bVerifyHistorydetails
End Function

'[Verify Pagination for the History table in Voice Biometrics Page displayed]
Public Function ValidatePagination_VB()
 bValidatePagination = true
 bNextPageExist = True
 While bNextPageExist = True
	intRecordCount = getRecordsCountForColumn(VoiceBio.tblHistoryHeader,VoiceBio.tblHistoryContent," Agent ID ")	
	iCheck = 5 
	If intRecordCount <=iCheck  Then
     LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to "&iCheck, true   
     bValidatePagination=true
	 If intRecordCount < iCheck Then
	   	bNextPageExist =matchStr(VoiceBio.lnkNext.GetROProperty("class"),"enabled")
		If bNextPageExist Then
		LogMessage "WARN","Verification","Next link expected to be disabled if record is less than "&iCheck&". Currently it is enabled.",false
		bValidatePagination=false
		Else
		LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
		End If
	 ElseIf intRecordCount = iCheck Then
		bNextPageExist = matchStr(VoiceBio.lnkNext.GetROProperty("class"),"enabled")
		If bNextPageExist Then
		VoiceBio.lnkNext.Click
		End If
	 End If
	Else 
		LogMessage "RSLT","Verification","Number of records displayed per page not matched with expected. Expected Count is less than or equal to 5", false   
		bNextPageExist = False
	End If
 Wend
 ValidatePagination_VB = bValidatePagination
End Function

'[Verify Opt-Out Opt-In Button enabled in Voice Biometrics Page displayed]
Public Function VerifyButtonOptOut_OptIn_VB()
   	bVerifyButtonOptOut_OptIn = true
   	intBtnFilter = Instr(VoiceBio.btnOptInOPtOut.GetROproperty("outerhtml"),("v-disabled"))
	If  intBtnFilter=0 Then
		LogMessage "RSLT","Verification","Opt-Out Opt-In button is enabled as expected.",True
		bVerifyButtonGO = true
	Else
		LogMessage "WARN","Verifiation","Opt-Out Opt-In button is disabled.",false
		bVerifyButtonOptOut_OptIn = false
	End If
	VerifyButtonOptOut_OptIn_VB = bVerifyButtonOptOut_OptIn
End Function

'[Verify Callback Button enabled in Voice Biometrics Page displayed]
Public Function VerifyButtonCallback_VB()
   	bVerifyButtonCallback = true
   	intBtnFilter = Instr(VoiceBio.btnCallBack.GetROproperty("outerhtml"),("v-disabled"))
	If intBtnFilter=0 Then
	  LogMessage "RSLT","Verification","Call-back button is enabled as expected.",True
	  bVerifyButtonCallback = true
	Else
	  LogMessage "WARN","Verifiation","Call-back button is disabled.",false
	  bVerifyButtonCallback = false
	End If
	VerifyButtonCallback_VB = bVerifyButtonCallback
End Function

'[Click Button Opt-Out Opt-In in Voice Biometrics Page]
Public Function clickButtonOptOutOptIn_VB()
  VoiceBio.btnOptInOptOut.click 10,10,0 
   If Err.Number<>0 Then
      clickButtonOptOutOptIn_VB=False
      LogMessage "WARN","Verification","Failed to Click Button : Opt-Out Opt-In" ,false
      Exit Function
   End If
   WaitForICallLoading
   clickButtonOptOutOptIn_VB=True
End Function

'[Click Button Callback in Voice Biometrics Page]
Public Function clickButtonCallback_VB()
   VoiceBio.btnCallBack.click 10,10,0 
   If Err.Number<>0 Then
      clickButtonCallback_VB=false
      LogMessage "WARN","Verification","Failed to Click Button : Call-back" ,false
      Exit Function
   End If
   WaitForICallLoading
   clickButtonCallback_VB=True
End Function

'[Verify User navigated to SMS Email Enquiry Page on click of SMSEmail link]
Public Function VerifySMSEmailLink_VB(strlink)
bVerifySMSEmailLink  = True
If Not IsNull(strlink) Then
    VoiceBio.lnkSMSEmailEnquiry().Click
    bVerifytab =  verifyTab(strlink)
    If bVerifytab = True  Then
       bVerifySMSEmailLink=True
    Else 
       bVerifySMSEmailLink=False
    End If
End If
VerifySMSEmailLink_VB = bVerifySMSEmailLink
End Function

'[Select From Date using Calendar Icon in Voice Biometrics Page]
Public Function SelectFromDateCalendar_VB(strFromDate)
bSelectFromDateCalendar	 = True 
Set oDesc = Description.Create    
oDesc("xpath").Value="//*[@id='voiceBiometrics_profile_start_date_input']/div[1]/md-icon"
set objCalendar = Browser("Browser_iCall_Home").Page("IServe_Opportunity").ChildObjects(oDesc)
SelectFromDateCalendar=selectDateFromCalendar(objCalendar(0),strFromDate)
strActFromDate = VoiceBio.txtFromDate.GetROProperty("value")
If strActFromDate =  strFromDate Then
   LogMessage "RSLT","Verification","Date displayed in the From Date field by selecting from the calendar is as expected",True
   bSelectFromDateCalendar	 = True 
Else 
	LogMessage "WARN","Verification","Date displayed in the From Date field by selecting from the calendar is not as expected",False
	bSelectFromDateCalendar	 = False 
End If
SelectFromDateCalendar_VB = bSelectFromDateCalendar
End Function

'[Select TO Date using Calendar Icon in Voice Biometrics Page]
Public Function SelectToDateCalendar_VB(strToDate)
bSelectToDateCalendar = True 
Set oDesc = Description.Create    
oDesc("xpath").Value="//*[@id='voiceBiometrics_profile_end_date_input custom-md-input']/DIV[1]/md-icon[1]"
set objCalendar = Browser("Browser_iCall_Home").Page("IServe_Opportunity").ChildObjects(oDesc)
SelectToDateCalendar=selectDateFromCalendar(objCalendar(0),strToDate)
strActToDate = VoiceBio.txtToDate.GetROProperty("value")
If strActToDate =  strToDate Then
   LogMessage "RSLT","Verification","Date displayed in the TO Date field by selecting from the calendar is as expected",True
   bSelectToDateCalendar = True 
Else 
	LogMessage "WARN","Verification","Date displayed in the TO Date field by selecting from the calendar is not as expected",False
	bSelectToDateCalendar = False 
End If
SelectToDateCalendar_VB = bSelectToDateCalendar
End Function
