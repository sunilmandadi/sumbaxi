'[Click Created date hyperlink on Actvities Tab]
Public Function viewIAActivityTab(strType, strComments, strCreatedDtTime)
	Dim bviewIAActivityTab:	viewIAActivityTab = True
	strCreatedByLcase=LCase(strCreatedBy)
	strMonth=Month(strSubmitTime)
	strYear=Year(strSubmitTime)
	strDay=Day(strSubmitTime)
	If strDay>=0 and strDay<10 Then
		strDay=0&strDay
	End If
	strHour=Hour(strSubmitTime)
	'If strHour=-1 Then
	'	strHour=23
	'End If
	If strHour>=0 and strHour<10 Then
		strHour=0&strHour
	End If
		strMin=Minute(strSubmitTime) + 1
	If strMin>=0 and strMin<10 Then
		strMin=0&strMin
	End If
	strMonthName=MonthName(strMonth, True)   
	'Comments:Created by "&strCreatedByLcase&" Date and Time "&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&" "&strComments&"
	If Not Isnull(strComments) Then
		lstActivityDetails = checknull("Created Date:"&strCreatedDtTime&"|Type:"&strType&"")
	else
		lstActivityDetails = checknull("Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"|Type:"&strType&"")
	End If
	bviewIAActivityTab=selectTableLink(ServiceRequest.tblActivityListHeader,ServiceRequest.tblActivityListContent,lstActivityDetails,"ActivityTab" ,"Created Date",True,ServiceRequest.lnkNext ,ServiceRequest.lnkNext1,ServiceRequest.lnkPrevious)
	WaitForICallLoading
	viewIAActivityTab=bviewIAActivityTab
End Function

'[Verify Tab ViewIA displayed]
Public Function verifyTabViewIAexist(strTabName)
   bDevPending=false
   verifyTabViewIAexist=verifyTabExist(strTabName)
End Function

'[Select Tab ViewIA]
Public Function setectTabViewIA(strTabName)
   bDevPending=false
   setectTabViewIA=selectTab(strTabName)
End Function

'[Close Tab ViewIA]
Public Function closeTabViewIA(strTabName)
   bDevPending=false
   closeTabViewIA=closeTab(strTabName)
End Function

'[Verify Field RelatedTo on IA displayed as]
Public Function verifyRelatedToText_IA(strRelatedTo)
   bDevPending=false
   bVerifyRelatedToText=true
	
   If Not IsNull(strRelatedTo) Then
     If Not VerifyInnerText (ViewIA.lblRelatedTo(), strRelatedTo, "Related To")Then
	   bVerifyRelatedToText=false
	End If
   End If
   verifyRelatedToText_IA=bVerifyRelatedToText
End Function


'[Verify Field CardNumber on IA displayed as]
Public Function verifyCardNumberText_IA(strAccountNumber)
   bDevPending=False
   bVerifyCardNumberText=true
   'strAccountNumber=Replace(strAccountNumber,"-","")
   If Not IsNull(strAccountNumber) Then
       If Not VerifyInnerText (ViewIA.lblCardNumber(), strAccountNumber, "Account Number")Then
           bVerifyCardNumberText=false
       End If
   End If
   verifyCardNumberText_IA=bVerifyCardNumberText
End Function

'[Verify Field Type on IA displayed as]
Public Function verifyTypeText_IA(strType)
   bDevPending=False
   bVerifyTypeText=true
   If Not IsNull(strType) Then
       If Not VerifyInnerText (ViewIA.lblType(), strType, "Type")Then
           bVerifyTypeText=false
       End If
   End If
   verifyTypeText_IA=bVerifyTypeText
End Function

'[Verify Field SubType on IA displayed as]
Public Function verifySubTypeText_IA(strSubType)
   bDevPending=false
   bVerifySubTypeText=true
   If Not IsNull(strSubType) Then
       If Not VerifyInnerText (ViewIA.lblSubType(), strSubType, "SubType")Then
           bVerifySubTypeText=false
       End If
   End If
   verifySubTypeText_IA=bVerifySubTypeText
End Function

'[Verify Field Product on IA displayed as]
Public Function verifyProductText_IA(strProduct)
   bDevPending=false
   bVerifySubTypeText=true
   If Not IsNull(strProduct) Then
       If Not VerifyInnerText (ViewIA.lblProduct(), strProduct, "Product")Then
           bVerifySubTypeText=false
       End If
   End If
   verifyProductText_IA=bVerifySubTypeText
End Function

'[Verify Field AssignedTo on IA displayed as]
Public Function verifyAssignedToText_IA(strAssignedTo)
   bDevPending=False
   bVerifyAssignedToText=true
   If Not IsNull(strAssignedTo) Then
       If Not VerifyInnerText (ViewIA.lblAssignedTo(), strAssignedTo, "AssignedTo")Then
           bVerifyAssignedToText=false
       End If
   End If
   verifyAssignedToText_IA=bVerifyAssignedToText
End Function

'[Verify Field Status on ViewIA Screen displayed as]
Public Function verifyStatus_ViewIAText(strStatus)
   bDevPending=False
   bVerifyStatus_ViewSRText=true
   If Not IsNull(strStatus) Then
       If Not VerifyInnerText (ViewIA.lblStatus_ViewIA(), strStatus, "Status_ViewSR")Then
           bVerifyStatus_ViewSRText=false
       End If
   End If
   verifyStatus_ViewIAText=bVerifyStatus_ViewSRText
End Function

'[Verify Field SubStatus on IA displayed as]
Public Function verifySubStatusText_IA(strSubStatus)
   bDevPending=false
   bVerifySubStatusText=true
   If Not IsNull(strSubStatus) Then
       If Not VerifyInnerText (ViewIA.lblSubStatus(), strSubStatus, "SubStatus")Then
           bVerifySubStatusText=false
       End If
   End If
   verifySubStatusText_IA=bVerifySubStatusText
End Function

'[Verify Field Resolution on IA displayed as]
Public Function verifyResolutionText_IA(strResolution)
   bDevPending=false
   verifyResolutionText_IA = true
   If Not IsNull(strResolution) Then
       If Not VerifyInnerText (ViewIA.lblResolution(), strResolution, "Resolution")Then
           verifyResolutionText_IA = false
       End If
   End If
End Function


'[Verify Field Due Date on IA displayed as]
Public Function verifyDueDateText_IA(strDueDate)
   bDevPending=false
   bVerifyDueDateText=true
   If Not IsNull(strDueDate) Then
		If UCase(strDueDate)="DATASTORE DUEDATE" Then
			strDueDate=fetchFromDataStore("Select Open Activity SR from CSO home Page","BLANK","Due Date")(0)
			strDueDate=Replace(strDueDate,".?",":")
		End If
   
       If Not VerifyInnerText (ViewIA.lblDueDate(), strDueDate, "DueDate")Then
           bVerifyDueDateText=false
       End If
   End If
   verifyDueDateText_IA=bVerifyDueDateText
End Function

'[Verify Field Comments on View IA Tab displayed as]
Public Function verifyComments_ViewIA(strExpectedComment,strAccountNumber,strCreatedDate,strCreatedBy)
   bDevPending=false
   bVerifyServiceRequest=true
 '  strAccountNumber=Replace(strAccountNumber,"-","")
   If Not IsNull(strExpectedComment) Then
	   If Ucase(strExpectedComment)="RUNTIME SR COMMENT" Then
		   strExpectedComment=fetchFromDataStore(gstrRuntimeCommentStep,"BLANK",gstrParameterNameStep)(0)
	   End If
	    If Ucase(strExpectedComment)="RUNTIME TM COMMENT" Then
		   strExpectedComment=fetchFromDataStore(gstrRuntimeTMCommentStep,"BLANK","TMComment")(0)
	   End If

      strActualComment= ViewIA.lblComments.GetROProperty("innertext")
	  LogMessage "RSLT","Verification","Comment displayed is  "&strActualComment,true
		  If Not IsNull(strCreatedDate) Then
			If Ucase(strCreatedDate)="TODAY" Then
				If len(Day(CDate(Now)))=1 Then
					strDay="0"&Day(CDate(Now))
				else
					strDay=""&Day(CDate(Now))
				End If
				strCreatedDatePattern=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&", ([0-2][0-9]:[0-9][0-9])"
			else
				strCreatedDatePattern=strCreatedDate &", ([0-2][0-9]:[0-9][0-9])"
			End If
			'strDueDate=Split(getVadinCombo_SelectedItem(lblDueDate),":")(0)
			If IsNull(checknull(strAccountNumber))Then
				strExpectedComment="Created by: "&strCreatedBy&"; Date and Time: "&strCreatedDatePattern &"; "&strExpectedComment
                
			Else
					strExpectedComment="Created by: "&strCreatedBy&"; Date and Time: "&strCreatedDatePattern &"; Account No: "&strAccountNumber&"; "&strExpectedComment
			End If
			If Matchstr(Ucase(strActualComment),Ucase(strExpectedComment)) Then
				LogMessage "RSLT","Verification","Comment pattern matched with expected pattern "&strExpectedComment,true
			else
				LogMessage "WARN","Verification","Comment pattern does not matched with Expected pattern "&strExpectedComment&" , Actual Comment displayed is "&strActualComment,false
				bVerifyServiceRequest=false
			End If
		 End If
   End If		

   verifyComments_ViewIA=bVerifyServiceRequest
End Function

'[Verify OnceAndDone Checkbox Checked on View IA Screen]
Public Function verifyOnceAndDone_IA(bChecked)
	Dim bverifyOnceAndDone:bverifyOnceAndDone=true
		intCheckox = Instr(ViewIA.chkOnceDone.GetROProperty ("outerhtml"),"checked")
		If bChecked Then
			If intCheckox = 0 Then
				LogMessage "RSLT","Verification","Once and Done check box is not unchecked as expected.",false
				bverifyOnceAndDone=false
            Else
				LogMessage "RSLT","Verification","Once and Done check box is by default unchecked as expected.",true
				bverifyOnceAndDone=true
			End If
		else
				If intCheckox = 0 Then
				LogMessage "RSLT","Verification","Once and Done check box is by default unchecked as expected.",true
				bverifyOnceAndDone=true
			Else
				LogMessage "RSLT","Verification","Once and Done check box is not unchecked as expected.",false
				bverifyOnceAndDone=false	
			End If
		End If
		verifyOnceAndDone_IA=bverifyOnceAndDone
End Function

'[Verify Field SRNumber on IA displayed as]
Public Function verifySRNumber_ViewIA(strSRNumber)
   bDevPending=false
   bVerifySRNumberText=true
   'If SR link clicked from table
   If Ucase(strSRNumber)="SELECTED SR IN TABLE" Then
		strSRNumber=fetchFromDataStore("Click SR Number link in Table Request Submitted","BLANK","SelectedSR")(0)
   End If
  ' If SR link clicked from popup lable link
   If Ucase(strSRNumber)="SELECTED SR LINK" Then
	   If not isNull(gstrRuntimeSRNumStep) Then
		   strSRNumber=fetchFromDataStore(gstrRuntimeSRNumStep,"BLANK","SelectedSRLink")(0)
	   End If
	   If isNull(strSRNumber) Then
			strSRNumber=fetchFromDataStore("Click Link SRNumber on Request Submitted popup","BLANK","SelectedSRLink")(0)
	   End If
   End If

   If Not IsNull(strSRNumber) Then
       If Not VerifyInnerText (ViewIA.lblSRNumber(), strSRNumber, "SRNumber")Then
           bVerifySRNumberText=false
       End If
   End If
   verifySRNumber_ViewIA=bVerifySRNumberText
End Function


'[Verify Field Channel on IA displayed as]
Public Function verifyChannelText_IA(strChannel)
   bDevPending=false
   bVerifyChannelText=true
   If Not IsNull(strChannel) Then
       If Not VerifyInnerText (ViewIA.lblChannel(), strChannel, "Channel")Then
           bVerifyChannelText=false
       End If
   End If
   verifyChannelText_IA=bVerifyChannelText
End Function

'[Verify Field Created Date on IA displayed as]
Public Function VerifyFieldCreatedDate_IA(strCreatedDate)
bDevPending=false
   bVerifyFieldCreatedDate=true
	strActualCreated=ViewIA.lblCreatedDate.GetRoProperty("innertext")
	If Not IsNull(strCreatedDate) Then
		If Ucase(strCreatedDate)="TODAY" Then
			If len(Day(CDate(Now)))=1 Then
				strDay="0"&Day(CDate(Now))
			else
				strDay=""&Day(CDate(Now))
			End If
			strCreatedDatePattern=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&" ([0-2][0-9]:[0-9][0-9])"
		else
			strCreatedDatePattern=strCreatedDate' &" ([0-2][0-9]:[0-9][0-9])"
		End If
		If Matchstr(strActualCreated,strCreatedDatePattern) Then
			LogMessage "RSLT","Verification","Created date pattern matched with expected pattern DD MMM YYYY HH:MM",true
		else
			LogMessage "WARN","Verification","Created date pattern does not matched with Expected pattern DD MMM YYYY, HH:MM Expected: "&strCreatedDatePattern&" , Actual Date displayed is "&strActualCreated,false
			bVerifyFieldCreatedDate=false
		End If
	End If
	VerifyFieldCreatedDate_IA=bVerifyFieldCreatedDate
End Function

'[Verify Field Created By on IA displayed as]
Public Function VerifyFieldCreatedBy_IA(strCreatedBy)
   bDevPending=false
   bVerifyFieldCreatedBy=true
	 If Not IsNull(strCreatedBy) Then  'Updated as per code refractoring for View SR
		'strCreatedBy_LAN_Id=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\SR_Screen_Map.ini",  "LAN ID" , strCreatedBy)
	  If Not VerifyInnerText (ViewIA.lblCreatedBy(), strCreatedBy, "Created By")Then
		   bVerifyFieldCreatedBy=false
		End If
	 End If
	 VerifyFieldCreatedBy_IA=bVerifyFieldCreatedBy
End Function

'[Verify Field Overdue on IA displayed as]
Public Function verifyOverdueText_IA(strOverdue)
   bDevPending=false
   bVerifyOverdueText=true
   If Not IsNull(strOverdue) Then
       If Not VerifyInnerText (ViewIA.lblOverdue(), strOverdue, "Overdue")Then
           bVerifyOverdueText=false
       End If
   End If
   verifyOverdueText_IA=bVerifyOverdueText
End Function

'[Verify Field Duration on IA displayed as]
Public Function verifyDurationText_IA(strDuration)
   bDevPending=false
   bVerifyDurationText=true
   If Not IsNull(strDuration) Then
       If Not VerifyInnerText (ViewIA.lblDuration(), strDuration, "Duration")Then
           bVerifyDurationText=false
       End If
   End If
   verifyDurationText_IA=bVerifyDurationText
End Function

'[Verify Field Closed Date on IA displayed as]
Public Function VerifyFieldClosedDate_IA(strClosedDate)
   bDevPending=false
   bVerifyFieldClosedDate=true
	strActualClosed=ViewIA.lblClosedDate.GetRoProperty("innertext")
	If Not IsNull(strClosedDate) Then
		If Ucase(strClosedDate)="TODAY" Then
			If len(Day(CDate(Now)))=1 Then
				strDay="0"&Day(CDate(Now))
			else
				strDay=""&Day(CDate(Now))
			End If			
			strClosedDate=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&" ([0-2][0-9]:[0-9][0-9])"
		End If

		If Matchstr(strActualClosed,strClosedDate) Then
			LogMessage "RSLT","Verification","Closed date pattern matched with expected pattern DD MMM YYYY HH:MM",true
		else
			LogMessage "WARN","Verification","Closed date pattern does not matched with Expected pattern DD MMM YYYY, HH:MM Expected: "&strClosedDate&" , Actual Date displayed is "&strActualClosed,false
			bVerifyFieldClosedDate=false
		End If
	End If
	VerifyFieldClosedDate_IA=bVerifyFieldClosedDate
End Function

'[Verify Field Last Updated Date on IA displayed as]
Public Function VerifyFieldLastUpdatedDate_IA(strLastUpdatedDate)
   bDevPending=false
   bVerifyFieldLastUpdatedDate=true
	strActualLastUpdate=ViewIA.lblLastUpdatedDate.GetRoProperty("innertext")
	If Not IsNull(strLastUpdatedDate) Then
		If Ucase(strLastUpdatedDate)="TODAY" Then
			If len(Day(CDate(Now)))=1 Then
				strDay="0"&Day(CDate(Now))
			else
				strDay=""&Day(CDate(Now))
			End If
			strLastUpdatedDate=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&" ([0-2][0-9]:[0-9][0-9])"
		End If	
		If Matchstr(strActualLastUpdate,strLastUpdatedDate) Then
			LogMessage "RSLT","Verification","Last Update date pattern matched with expected pattern DD MMM YYYY HH:MM",true
		else
			LogMessage "WARN","Verification","Last Update date pattern does not matched with Expected pattern DD MMM YYYY, HH:MM Expected: "&strLastUpdatedDate&" , Actual Date displayed is "&strActualLastUpdate,false
			bVerifyFieldLastUpdatedDate=false
		End If
	End If
		VerifyFieldLastUpdatedDate_IA=bVerifyFieldLastUpdatedDate
End Function

'[Verify Field Last Updated By on IA displayed as]
Public Function VerifyFieldLastUpdateBy_IA(strLastUpdatedBy)
   bDevPending=false
	bVerifyFieldLastUpdateBy=true
	If Not IsNull(strLastUpdatedBy) Then 'Updated as per code refractoring for View SR
'		strLastUpdatedBy_LAN_Id=""
'		If not strLastUpdatedBy="" Then
'			strLastUpdatedBy_LAN_Id=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\SR_Screen_Map.ini",  "LAN ID" , strLastUpdatedBy)
'		End If
	  If Not VerifyInnerText (ViewIA.lblLastUpdatedBy, strLastUpdatedBy, "Last Updated By")Then
		   bVerifyFieldLastUpdateBy=false
		End If
	End If
		VerifyFieldLastUpdateBy_IA=bVerifyFieldLastUpdateBy
End Function

'Added by Kalyan Prospects 1603 Dated 29/07/2016

'[Verify View IA Field prospectid for Prospects displayed as]
Public Function verifyViewIARefNoProspects(strProspectId)
   bDevPending=false
   bverifyViewIARefNoProspects=true
	
   If Not IsNull(strRefNo) Then
     If Not VerifyInnerText (ViewIA.lblProspectsRef(), strProspectId, "Prospect Id")Then
	   bverifyViewIARefNoProspects=false
	End If
   End If
   verifyViewIARefNoProspects=bverifyViewIARefNoProspects
End Function

'[Verify View IA Field NRIC for Prospects displayed as]
Public Function verifyViewIANRICProspects(strNRIC)
   bDevPending=false
   bverifyViewIANRICProspects=true
	
   If Not IsNull(strNRIC) Then
     If Not VerifyInnerText (ViewIA.lblProspectsNRIC(), strNRIC, "NRIC")Then
	   bverifyViewIANRICProspects=false
	End If
   End If
   verifyViewIANRICProspects=bverifyViewIANRICProspects
End Function

'[Verify View IA Field Mobile for Prospects displayed as]
Public Function verifyViewIAMobileProspects(strMobile)
   bDevPending=false
   bverifyViewIAMobileProspects=true
	
   If Not IsNull(strMobile) Then
     If Not VerifyInnerText (ViewIA.lblProspectsMobile(), strMobile, "Mobile")Then
	   bverifyViewIAMobileProspects=false
	End If
   End If
   verifyViewIAMobileProspects=bverifyViewIAMobileProspects
End Function

'[Verify View IA Field FirstName for Prospects displayed as]
Public Function verifyViewIAFstNameProspects(strFstName)
   bDevPending=false
   bverifyViewIAFstNameProspects=true
	
   If Not IsNull(strFstName) Then
     If Not VerifyInnerText (ViewIA.lblProspectsFirstName(), strFstName, "FirstName")Then
	   bverifyViewIAFstNameProspects=false
	End If
   End If
   verifyViewIAFstNameProspects=bverifyViewIAFstNameProspects
End Function

'[Verify View IA Field DOB for Prospects displayed as]
Public Function verifyViewIADOBProspects(strDOB)
   bDevPending=false
   bverifyViewIADOBProspects=true
	
   If Not IsNull(strDOB) Then
     If Not VerifyInnerText (ViewIA.lblProspectsDateofBirth(), strDOB, "DOB")Then
	   bverifyViewIADOBProspects=false
	End If
   End If
   verifyViewIADOBProspects=bverifyViewIADOBProspects
End Function

'[Verify View IA Field EmailId for Prospects displayed as]
Public Function verifyViewIAEmailIdProspects(strEmailId)
   bDevPending=false
   bverifyViewIAEmailIdProspects=true
	
   If Not IsNull(strEmailId) Then
     If Not VerifyInnerText (ViewIA.lblProspectsEmailID(), strEmailId, "EmailId")Then
	   bverifyViewIAEmailIdProspects=false
	End If
   End If
   verifyViewIAEmailIdProspects=bverifyViewIAEmailIdProspects
End Function

'[Verify View IA Field LastName for Prospects displayed as]
Public Function verifyViewIALstNameProspects(strLstName)
   bDevPending=false
   bverifyViewIALstNameProspects=true
	
   If Not IsNull(strLstName) Then
     If Not VerifyInnerText (ViewIA.lblProspectsLastName(), strLstName, "LastName")Then
	   bverifyViewIALstNameProspects=false
	End If
   End If
   verifyViewIALstNameProspects=bverifyViewIALstNameProspects
End Function

'[Verify Pink Panel displayed in ViewIA Page]
Public Function verifyViewIAPinkPanel(strName,strCIN,strSegment)
	verifyViewIAPinkPanel = false	
	If Not IsNull (strName) and Not IsNull (strCIN) and Not IsNull (strSegment) Then
	   If verifyInnerText(ViewIA.lblName(),strName, "Name") and verifyInnerText(ViewIA.lblCIN(),strCIN, "CIN") and verifyInnerText(ViewIA.lblSegment(),strSegment, "Segment") Then
		verifyViewIAPinkPanel = true
	   End If
	End If
End Function

'[Verify View IA Number Field displayed as]
Public Function verifyIANumber_ViewIA(strIAnumberViewIAField)
verifyIANumber_ViewIA = false
If Not IsNull(strIAnumberViewIAField) Then
	If VerifyInnerText(ViewIA.lblviewIANumber(),strIAnumberViewIAField, "IA Number") Then
		verifyIANumber_ViewIA = true
	End If
End If	
End Function

'[Verify the shortcut buttons on View IA page]
Public Function verifyShortCutButtonViewIAPage()
	verifyShortCutButtonViewIAPage = false
	If (ViewIA.btnClose.Exist(1))Then
		LogMessage "RSLT","Verification","Close button is available for View IA Page." ,True
        verifyShortCutButtonViewIAPage = true
	Else
	    LogMessage "WARN","Verification","Close button is not available for View IA Page. Expected to be Enabled." ,false
	End If

    If (ViewIA.btnEdit.Exist(1)) Then
		LogMessage "RSLT","Verification","Edit Button is available for View IA Page." ,True
        verifyShortCutButtonViewIAPage = true
	Else
	   LogMessage "WARN","Verification","Edit button is not available for View IA Page. Expected to be Enabled." ,false
	End If
	WaitForICallLoading
End Function

''[Select Activity from View SR screen]
'Public Function viewIAActivityTabTMFlow(lstActivityDetails)
'          Dim bviewIAActivityTabTMFlow:	viewIAActivityTabTMFlow = True
'          bviewIAActivityTabTMFlow=selectTableLink(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent,lstActivityDetails,"ActivityTab" ,"Created Date",False,NULL,NULL,NULL)
'          WaitForICallLoading
'		  viewIAActivityTabTMFlow=bviewIAActivityTabTMFlow
'End Function
