'[Click on Send SMS Button]
Public Function clickBtnSendSMS()
 SendEmailSMS.ButtonSendSMS.click
	If Err.Number<>0 Then
       clickBtnSendSMS=false
       LogMessage "WARN","Verification","Failed to Click Button : Send SMS" ,false
       Exit Function
   End If
	clickBtnSendSMS=true
	WaitForICallLoading
End Function

'[Click on link SMS/Email Inquiry in Customer Overview Page]
Public Function clickSMSInquiry()
	bcCustomerOverview.lnkEmailSMSInquiry.Click
	WaitForICallLoading
	If Err.Number<>0 Then
       clickSMSInquiry=false
       LogMessage "WARN","Verification","Failed to Click Button : SMS/Email Inquiry" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickSMSInquiry=true
	WaitForICallLoading
End Function

'[Verify Button SendSMS display on SMS Email Inquiry Page]
Public Function VerifybtnSendSMS_SendSMSInquriy(strAction)
    bVerifybtnSendSMS_SendSMSInquriy=true
    If strAction = "Enabled" Then
    	intBtnSendSMS=Instr(SendEmailSMS.ButtonSendSMS.GetROproperty("outerhtml"),("v-disabled"))
		If  intBtnSendSMS=0 Then
			LogMessage "RSLT","Verification","Send SMS button is enabled as expected.",True
			bVerifybtnSendSMS_SendSMSInquriy=true
		Else
			LogMessage "WARN","Verifiation","Send SMS button is disabled.",false
			bVerifybtnSendSMS_SendSMSInquriy=false
		End If
	Else  If strAction = "Disabled" Then
	    	intBtnSendSMS=Instr(SendEmailSMS.ButtonSendSMS.GetROproperty("outerhtml"),("v-enabled"))
		If  intBtnSendSMS=0 Then
			LogMessage "RSLT","Verification","Send SMS button is disabled as expected.",True
			bVerifybtnSendSMS_SendSMSInquriy=true
		Else
			LogMessage "WARN","Verifiation","Send SMS button is enabled.",false
			bVerifybtnSendSMS_SendSMSInquriy=false
		End If
	End If 
    End If
	VerifybtnSendSMS_SendSMSInquriy=bVerifybtnSendSMS_SendSMSInquriy
End Function

'[Click Button GO in SMS EMail Inquiry Page]
Public Function clickButtonGO_SMSInquiry()
	WaitForIcallLoading
   SendEmailSMS.btnGO.click
   WaitForIcallLoading
   If Err.Number<>0 Then
       clickButtonGO_SMSInquiry=false
            LogMessage "WARN","Verification","Failed to Click Button : GO in SMS EMail INquiry Popup" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonGO_SMSInquiry=true
End Function

'[Verify default show dropdown value in SMS Email Enquiry Page displayed as]
Public Function verifyShowDefault_SMSInquiry(strStatus)
   bverifyShowdropdown=true
   If Not IsNull(strStatus) Then
       If Not verifyComboSelectItem(SendEmailSMS.showdropdown(),strStatus, "Show")Then
       	  LogMessage "WARN","Verification","Default Show value doesnt match with expected value" ,false
          bverifyShowdropdown=false
       End If
   End If
   verifyShowDefault_SMSInquiry = bverifyShowdropdown
End Function

'[Verify list of values in show dropdown displayed in SMS Email Enquiry Page as]
Public Function verifyShowList_SMSInquiry(lstShow)
   bverifyShowDropdown=true
   If Not IsNull(lstShow) Then
       If Not verifyComboboxItems(SendEmailSMS.showdropdown(),lstShow, "Show combobox")Then
           bverifyShowDropdown=false
       End If
   End If
   verifyShowList_SMSInquiry = bverifyShowDropdown
End Function

'[Select Combobox Show on SendSMS Enquiry Page as]
Public Function selectShowComboBox_SMSEnquiry(strShow)
	WaitForICallLoading
   bselectShowComboBox_SMSEnquiry=true
   WaitForICallLoading
   If Not IsNull(strShow) Then
   		WaitForICallLoading
       If Not (selectItem_Combobox (SendEmailSMS.showdropdown(), strShow))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Show drop down list" ,false
           bSelectShowComboBox=false
       End If
   End If
   WaitForICallLoading
   selectShowComboBox_SMSEnquiry=bselectShowComboBox_SMSEnquiry
End Function

'[Verify list of values displayed in Show dropdown as]
Public Function verifyShowdropdown(lstShow)
   bverifyShowdropdown=true
   If Not IsNull(lstShow) Then
       If Not verifyComboboxItems(SendEmailSMS.showdropdown(),lstShow, "Show Dropdown")Then
           bverifyShowdropdown=false
       End If
   End If
   verifyShowdropdown=bverifyShowdropdown
End Function

'[Validate Pagination of SMS Email Inquiry Page]
Public Function validatePagination_SMSInquiry(strshow)
 bvalidatePagination=true
 If strshow = "All" OR strshow = "Email" Then
	iCheck = 5
ElseIf strshow = "SMS" Then
	iCheck = 10
End If
bNextPageExist = True
While bNextPageExist = True
	 intRecordCount = getRecordsCountForColumn(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent, "Type")
	 If intRecordCount <=iCheck  Then
	     LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to "&iCheck, true   
	     bvalidatePagination_SMSInquiry=true
		 If intRecordCount < iCheck Then
		   	bNextPageExist =matchStr(SendEmailSMS.lnkNext.GetROProperty("outerhtml"),"v-enabled")
			If bNextPageExist Then
			LogMessage "WARN","Verification","Next link expected to be disabled if record is less than "&iCheck&". Currently it is enabled.",false
			bvalidatePagination=false
			Else
			LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
			End If
		ElseIf intRecordCount = iCheck Then
			bNextPageExist = matchStr(SendEmailSMS.lnkNext.GetROProperty("outerhtml"),"v-enabled")
			If bNextPageExist Then
				SendEmailSMS.lnkNext.Click
			End If
		End If
	Else 
		LogMessage "RSLT","Verification","Number of records displayed per page not matched with expected. Expected Count is less than or equal to 5", false 
		bvalidatePagination = False
		bNextPageExist = False
	End If
Wend
validatePagination_SMSInquiry = bvalidatePagination
End Function

'[Set TextBox From Date on SMS Email Inquiry to]
Public Function setFromTextbox_SMSInquiry(strFrom)
bsetFromTextbox_SMSInquiry = True 
   If Instr(strFrom, "TODAY")  Then
	  strFromDate = SetDateField_Validation(strFrom)
	  strFromDate = Replace(strFromDate,"-"," ")
	  SendEmailSMS.txtFrom.Set(strFromDate)	
   Else
   	  SendEmailSMS.txtFrom.Set(strFrom)
   End IF
   If Err.Number<>0 Then
       bsetFromTextbox_SMSInquiry=false
       LogMessage "WARN","Verification","Failed to Set Text Box :From" ,false
       Exit Function
   End If
setFromTextbox_SMSInquiry = bsetFromTextbox_SMSInquiry
End Function
   
'[Set TextBox To Date on SMS Email Inquiry to]
Public Function setToDate_SMSInquiry(strTO)
   bsetTOTextbox_SMSInquiry = True 
   If Instr(strTO, "TODAY")  Then
	  strTODate = SetDateField_Validation(strTO)
	  strTODate = Replace(strTODate,"-"," ")
	  SendEmailSMS.txtTo.Set(strTODate)
   Else
   	  SendEmailSMS.txtTo.Set(strTO)
   End IF 
  If Err.Number<>0 Then
   bsetTOTextbox_SMSInquiry = false
   LogMessage "WARN","Verification","Failed to Set Text Box :TO Date" ,false
   Exit Function
  End If
setToDate_SMSInquiry = bsetTOTextbox_SMSInquiry
End Function

'[Verify Inline error Message on SMS Email Inquiry page displayed as]
Public Function verifyInlineErrorMessage_SendSMSInquiry(strInLineMessage)
bverifyInlineErrorMessage=true
If Not IsNull(strInLineMessage) Then
   If Not VerifyInnerText (SendEmailSMS.lblDateErrorMessage(), strInLineMessage, "Inline Info Message")Then
       bverifyInlineErrorMessage=false
   End If
End If
verifyInlineErrorMessage_SendSMSInquiry=bverifyInlineErrorMessage
End Function

'[Verify default From date and To Date displayed as]
Public Function VerifyDefaultDate_SMSInquiry(strToDate)
   bverifydefaultdate = True 
   IserveSMSFromDate = SendEmailSMS.txtFrom.GetROProperty("value") 
   IserveSMSToDate= SendEmailSMS.txtTo.GetROProperty("value") 
   If not isNull(strToDate) Then
	   If Ucase(strToDate)="TODAY" Then
			If len(Day(CDate(Now)))=1 Then
				strDay="0"&Day(CDate(Now))
			else
				strDay=""&Day(CDate(Now))
			End If
			strExpTODate=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&""
			'Set From Date
			strExpFromDate = Cdate(Now)-29
			If len(Day(strExpFromDate))=1 Then
				strDay="0"&Day(strExpFromDate)
			else
				strDay=""&Day(strExpFromDate)
			End If
			strExpFromDate=""&strDay & " "&monthName(Month(strExpFromDate),true) &" " &Year(strExpFromDate)&""			
			If Ucase(Trim(strExpFromDate)) = Ucase(Trim(IserveSMSFromDate)) AND Ucase(Trim(strExpTODate)) = Ucase(Trim(IserveSMSToDate)) Then
			   LogMessage "RSLT", "Verification","From and TO Dates not set as expected.Expected:"+ strExpFromDate &" "+ strExpTODate &", Actual:"+ strExpFromDate &" "+ strExpTODate &"", True
			Else 
				LogMessage "RSLT", "Verification","From and TO Dates are not as expected.Expected:"+ strExpFromDate &" "+ strExpTODate &", Actual:"+ strExpFromDate &" "+ strExpTODate &"", False
				bverifydefaultdate = False 
			End If
'			If Ucase(Trim(strExpTODate)) = Ucase(Trim(IserveSMSToDate)) Then
'			   LogMessage "RSLT", "Verification","To Date is set to 30 days before today's date. Expected: "+ strExpToDate &" , Actual: "& IserveSMSToDate, True
'			End If			
	   End If
   End If
   VerifyDefaultDate_SMSInquiry = bverifydefaultdate
End Function

'[Verify table details displayed in SMS Email Inquiry Page]
Public Function VerifySMSEmailDetails(lstlstSMSEmailDetails,lstSMSEmail,StrColumnname, strMessage)
bVerifySMSEmailDetails = True
WaitForIcallLoading
bverifySMStable = verifyTableContentList(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,lstlstSMSEmailDetails,"SMS/Email Enquiry",False,null,null,null)
	If bverifySMStable = True Then	
	WaitForIcallLoading	
  	   bUserclickMessageLink = selectTableLink(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,lstSMSEmail, "SMS/Email Enquiry", StrColumnname, False, null, null, null)
  	   WaitForIcallLoading
		If bUserclickMessageLink Then
			strExpMessage = Replace(strMessage,"@","=")
			If verifyInnerText(SendEmailSMS.SMSEmailMessagePopup(),strExpMessage, "Message text") Then
			Else
			   bVerifySMSEmailDetails = False
			End If
			SendEmailSMS.ButtonOK.Click
		Else
			LogMessage "WARN","Verification","Link for Message column not clicked from the record table" ,False
		End If		
   End If
   WaitForIcallLoading
VerifySMSEmailDetails = bVerifySMSEmailDetails
End Function

'[Verify details displayed in table SMS Email Inquiry]
Public Function ClickMessageLink_SMSEnquiry(strSMSText,strStatus,strTmpId,strSentTo)
	strToday = CDate(Now)
	If len(Day(strToday))=1 Then
		strDay="0"&Day(strToday)
	else
		strDay=""&Day(strToday)
	End If
	strToday=""&strDay & " "&monthName(Month(strToday),true) &" " &Year(strToday)&""
	intRecordCount = getRecordsCountForColumn(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent, "Type")
	For i = 0 To intRecordCount - 1
		strCellValueType=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Type")
		strCellValueDir=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Direction")
		strCellValueCreated=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Created Date")
		strCellValueCreated = Left(strCellValueCreated,12)
		strCellValueSentDate=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Sent Date")
		strCellValueSentDate = Left(strCellValueSentDate,12)
		strCellValueStatus=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Message Status")
		strCellValueTmpId=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Template Id")
		strCellValueSentTo=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Sent To")
		strCellValueSentBy=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Sent By")
		If UCase(strCellValueType) = "SMS" AND UCase(strCellValueDir) = "OUT" AND strCellValueCreated = strToday AND strCellValueSentDate = strToday AND strCellValueStatus = strStatus AND strCellValueTmpId = strTmpId AND strCellValueSentBy = "ICAL" AND strCellValueSentTo = strSentTo Then
			UserclickMessageLink=selectTableLink(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,Array("Sent By:"&strSentBy),"SMS/Email Inquiry table","Message",False,null,null,null)
			If UserclickMessageLink Then
				LogMessage "RSLT","Verification","Message is clicked successfully from the record table" ,True
				If verifyInnerText(SendEmailSMS.SMSEmailMessagePopup(),strSMSText, "Message text") Then
					LogMessage "RSLT","Verification","Message displayed matched with the SMS text sent" ,True
				Else
					LogMessage "WARN","Verification","Message displayed matched with the SMS text sent" ,False
				End If
				SendEmailSMS.ButtonOK.Click
				WaitForIcallLoading
			Else
				LogMessage "WARN","Verification","Message is not clicked successfully from the record table" ,False
			End If			
		Exit For
		ElseIf i=4 Then
			LogMessage "WARN","Verification","SMS Sent is not found in Page" ,False
		End If
	Next
End Function

'[Verify Message status column value for Email and SMS]
Public Function verifyMessageStatus_SendSMSInquiry()
	intRecordCount = getRecordsCountForColumn(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent, "Type")
	'loop for number of rows to verify hyperlink
	For i = 0 To intRecordCount - 1
		'intCol=getColIndex (objTableHeader,"Message Status")
		Set objAllRows=getAllRows(SendEmailSMS.SMSEmailEnquirytblContent)
		set objCountInCell=getCellObject(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i,"Message Status","md-button")
		iObjCount=objCountInCell.count
		strCellValue=getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i, "Message Status")

		For j = 0 to iObjCount - 1	
			strClassName=objCountInCell(j).getRoProperty("class")
			If instr (1,strClassName,"md-button",0) or instr (1,strClassName,"v-button-text-link",0)Then
				If (objCountInCell(j).getRoProperty("innertext") =  getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,objAllRows(i),i, "Message Status")) Then
					bDisabled =matchStr(objCountInCell(j).GetROProperty("outerhtml"),"disabled")
					Exit For
				End If
			End If
		Next
		'strCellvalType =  getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,i,"Message Status")
		If Ucase(strCellValue) = "SUCCESSFUL" OR Ucase(strCellValue) = "CUSTOMER NOT FOUND" OR Ucase(strCellValue) = "CONTACT INVALID" Then
			If bDisabled Then
				LogMessage "RSLT","Verification","Link is disabled.",True
			ElseIf Not bDisabled Then
				LogMessage "WARN","Verification","Link is enabled.",False
				Exit For
			End If
		ElseIf Ucase(strCellValue) = "FAILURE" OR Ucase(strCellValue) = "WAITING" OR Ucase(strCellValue) = "EXPIRED" Then
			If Not bDisabled Then
				LogMessage "RSLT","Verification","Link is enabled.",True
				objCountInCell(j).click
				WaitForIcallLoading
				SendEmailSMS.btnOK.click
			Else
				LogMessage "WARN","Verification","Link is disabled.",False
				Exit For
			End If
		Else 
			LogMessage "WARN","Verification","Message Status Other Than Expected Status",False
		End If
	Next
End Function

'[Verify Sent TO and From columns values displayed based on for Direction]
Public Function VerifyColumnforDirection_SendSMSInquiry()
	For i = 1 To 10
		intRecordCount = getRecordsCountForColumn(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent, "Type")
		For j = 0 To intRecordCount - 1
			strCellValDir = getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,j, "Direction")
			strCellValFrom = getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,j, "From")
			strCellValSentTo = getCellTextFor(SendEmailSMS.SMSEmailEnquirytblheader,SendEmailSMS.SMSEmailEnquirytblContent,j, "Sent To")
			If UCase(strCellValDir) = "OUT" Then
				If Not IsNull(strCellValSentTo) And trim(strCellValFrom) ="" Then
					LogMessage "RSLT","Verification","Sent To and From column values are displayed as Expected for Out Direction.",True
				Else
					LogMessage "WARN","Verification","Sent To and From column values are not displayed as Expected for Out Direction.",False
					Exit For
					Exit For
				End If
			ElseIf UCase(strCellValDir) = "IN" Then
				If Not IsNull(strCellValFrom) AND Trim(strCellValSentTo) ="" Then
					LogMessage "RSLT","Verification","Sent To and From column values are displayed as Expected for In Direction.",True
				Else
					LogMessage "WARN","Verification","Sent To and From column values are not displayed as Expected for In Direction.",False
					Exit For
					Exit For
				End If
			End If
		Next
		bNextPageExist =matchStr(SendEmailSMS.lnkNext.GetROProperty("outerhtml"),"v-enabled")
		If bNextPageExist Then
			SendEmailSMS.lnkNext.Click
		Else
			Exit For
		End If
	Next
End Function

'[Verify Validation Message displayed on SMS Registration Code as]
Public Function verifyValidationMessage_SMSRegistrCode(strValidationMsgSMSRegCode)
	verifyValidationMessage_SMSRegistrCode = True
	For iCountj = 1 To 180 Step 1
		If Not SendEmailSMS.popupPreValidateSMSRegistCode.Exist(0.5) Then
			Wait(0.5)
			verifyValidationMessage_SMSRegistrCode = False
		else
			If Not IsNull(strValidationMsgSMSRegCode) Then
				If Not VerifyInnerText (SendEmailSMS.lblPopUpPreValidContent(), strValidationMsgSMSRegCode, "Validation Message")Then
					verifyValidationMessage_SMSRegistrCode = false
				End If
			End If
			Exit for
		End if
	Next	
	For icountl = 1 To 180 Step 1
		If Not SendEmailSMS.btnPopupOkbtn.Exist(0.5) Then
			Wait(0.5)
			verifyValidationMessage_SMSRegistrCode = False
		else
			SendEmailSMS.btnPopupOkbtn.Click
			Exit for
		End If
	Next
	
	If Err.Number<>0 Then
		verifyValidationMessage_SMSRegistrCode = false
		LogMessage "WARN","Verification","Failed to Click button Ok" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","The button Ok is clicked successfully",true
	End If
	WaitForIcallLoading
End Function

'[Verify Left Panel Fields displayed in IB Registration Code in SMS Registration Code Screen]
Public Function verifyIBRegistrationCodeSMS(strIBRegistrationCodeSMS)
	verifyIBRegistrationCodeSMS = True
	For iCountq = 1 To 180 Step 1
		If Not SendEmailSMS.lbliBRegistrationCode.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(strIBRegistrationCodeSMS) Then
				If Not VerifyInnerText (SendEmailSMS.lbliBRegistrationCode(), strIBRegistrationCodeSMS, "Validation Message")Then
					verifyIBRegistrationCodeSMS = false
				End If
			End If
			Exit for
		End if
	Next
	If Err.Number<>0 Then
		verifyIBRegistrationCodeSMS = false
		LogMessage "WARN","Verification","Failed to Verify Inner Text" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","Verify Inner Text is Sccessfull",true
	End If
	WaitForIcallLoading
End Function

'[Verify Default Mobile No as]
Public Function verifyDefltMobileNo(strDefMobileNo)
	verifyDefltMobileNo = True
	For iCountw = 1 To 180 Step 1
		If Not SendEmailSMS.txtMobileNoSendSMS.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(strDefMobileNo) Then
				If Not verifyFieldValue (SendEmailSMS.txtMobileNoSendSMS(), strDefMobileNo, "Validation Message")Then
					verifyDefltMobileNo = false
				End If
			End If
			Exit for
		End if
	Next
	If Err.Number<>0 Then
		verifyDefltMobileNo = false
		LogMessage "WARN","Verification","Failed to Verify Inner Text" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","Verify Inner Text is Sccessfull",true
	End If
	WaitForIcallLoading	
End Function

'[Verify the SMS Text Msg of Send SMS Request as]
Public Function verifySMSTxtSendRegstCode(strSMSTxtregistCodea)
	verifySMSTxtSendRegstCode = True
	For iCounte = 1 To 180 Step 1
		If Not SendEmailSMS.txtSMSTxtSendRegistCode.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(strSMSTxtregistCodea) Then
				If Not VerifyInnerText (SendEmailSMS.txtSMSTxtSendRegistCode(), strSMSTxtregistCodea, "Validation Message")Then
					verifySMSTxtSendRegstCode = false
				End If
			End If
			Exit for
		End if
	Next
	If Err.Number<>0 Then
		verifySMSTxtSendRegstCode = false
		LogMessage "WARN","Verification","Failed to Verify Inner Text" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","Verify Inner Text is Sccessfull",true
	End If
	WaitForIcallLoading	
End Function

'[Verify the Character Counter inline message of SMS Text as]
Public Function verifyCharactersCountInMsgSMSTxt(strCharactersCountInMsgSMSTxt)
	verifyCharactersCountInMsgSMSTxt = True
	If Not IsNull(strCharactersCountInMsgSMSTxt) Then
		If Not VerifyInnerText (SendEmailSMS.lblInlineMsgCharLen(), strCharactersCountInMsgSMSTxt, "Validation Message")Then
			verifySMSTxtSendRegstCode = false
		End If
	End If
	If Err.Number<>0 Then
		verifyCharactersCountInMsgSMSTxt = false
		LogMessage "WARN","Verification","Failed to Verify Inner Text" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","Verify Inner Text is Sccessfull",true
	End If
	WaitForIcallLoading	
End Function

'[Perform Add Notes by clicking Add Notes Button on SMS Registration Code Screen as]
Public Function addNoteSMSRegCode(strNote)
	bDevPending=false
	Dim bVerifypopupNotes:bVerifypopupNotes=true	
	If not isNull(strNote) Then
		SendEmailSMS.btnAddNotes.click
		WaitForICallLoading
		If not ServiceRequest.popupVerification.exist(5)Then
			LogMessage "WARN","Verification","New Note dialog did not displayed",false
			bVerifypopupNotes=false
		else
			strMessage=ServiceRequest.lblMaxAllowed.GetROProperty("innerText")
			If not strMessage="Max allowed - 3000" Then
				LogMessage "WARN","Verification","Add New Comment popup dialog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage,false
				bVerifypopupNotes=false
			End If
			ServiceRequest.txtNewComment.set strNote			  
			ServiceRequest.clickSave_Popup
			WaitForIcallLoading
		End If 
	End If 
	addNoteSMSRegCode=bVerifypopupNotes
End Function

'[Verify Button Add Notes is disabled on SMS Registration Code Screen as]
Public Function VerifybtnAddNoteDisabledSMSRegCde()
	bDevPending=False
  	Dim bVerifyAddNote:bVerifyAddNote=true
	intBtnNote = SendEmailSMS.btnAddNote.GetRoproperty("disabled")
	If not intBtnNote=0 Then
		LogMessage "RSLT","Verification","Note button is disabled as per expectation.",True
		bVerifyAddNote=true
	Else
		LogMessage "WARN","Verifiation","Note button is enable. Expected to be disabale.",false
		bVerifyAddNote=false
	End If
	VerifybtnAddNoteDisabledSMSRegCde=bVerifyAddNote
End Function

'[Set TextBox Comment on SMS Registration Code to]
Public Function setCommentTextboxSMSRegCode(strComment)
	bDevPending=False
	strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comment on SMS Registration Code to"
	gstrParameterNameStep = "TimeStamp"&replace((replace((replace(now,"/","-"))," ","-")),":","-")
	insertDataStore gstrParameterNameStep, strComment
	'insertDataStore "SRComment", strComment	
	SendEmailSMS.txtComment.Set(strComment )
	If Err.Number<>0 Then
		setCommentTextboxSMSRegCode = false
		LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
		Exit Function
	End If
	setCommentTextboxSMSRegCode = true
End Function

'[Verify text description on Send SMS SR Screen displayed as]
Public Function verifyDescription_SendSMS(strDescription)
   bverifyDescription_SendSMS=true
   If Not IsNull(strDescription) Then
       If Not VerifyInnerText (SendEmailSMS.lblDescription(), strDescription, "Description")Then
           bverifyDescription_SendSMS=false
       End If
   End If
   verifyDescription_SendSMS=bverifyDescription_SendSMS
End Function

'[Verify Field KnowledgeBase on SMS Registration Code Screen displayed as]
Public Function verifyKnowledgeBaseSendSMS(strExpectedLink)
	bverifyKnowledgeBaseSMS=true
	If Not IsNull(strExpectedLink) Then		
		'Set oDesc_KB = Description.Create()
		'oDesc_KB("micclass").Value = "Link"		
		strKBLink = SendEmailSMS.lnkKnowledgeBase.GetROProperty("href")
		strExpectedLink=Replace(strExpectedLink,"@","=")
		If not MatchStr(strKBLink, strExpectedLink)Then
			LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
			bverifyKnowledgeBaseSMS = false
		else
			LogMessage "RSLT","Verification","Knowledge base link matched with expected link",true
		End If
	End If
	verifyKnowledgeBaseSendSMS = bverifyKnowledgeBaseSMS
End Function

'[Click Button Submit on SMS Registration Code Page]
Public Function clickButtonSubmitSendSMS()
   SendEmailSMS.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmitSendSMS = false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonSubmitSendSMS = true
End Function

'[Verify Popup Request Submitted exist for Send SMS]
Public Function verifyPopupSubmittedReq(bExist)
   bDevPending=false
   WaitForIcallLoading   
   For iCountrr = 1 To 180 Step 1
		If Not SendEmailSMS.popupRequestSubmitted.Exist(0.5) Then
			Wait(0.5)
		else
			bActualExist=SendEmailSMS.popupRequestSubmitted.Exist(4)
			Exit for
		End if
	Next	
	If Err.Number<>0 Then
		verifyPopupSubmittedReq = false
		LogMessage "WARN","Verification","Not triggered Request Submitted Popup Window" ,false
		Exit Function
	End If 
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupSubmittedReq=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupSubmittedReq=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupSubmittedReq=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupSubmittedReq=False
   End If
End Function

'[Verify Field Status_RequestSubmitted For Send SMS displayed as]
Public Function verifyStatusRequestSubmitSMs(strStatusReqSubSMS)
	verifyStatusRequestSubmitSMs = True
	WaitForIcallLoading
	If Not IsNull(strStatusReqSubSMS) Then
		WaitForIcallLoading
       If Not VerifyInnerText (SendEmailSMS.lblSRStatus(), strStatusReqSubSMS, "SR Status")Then
           verifyStatusRequestSubmitSMs=false
       End If
   End If
End Function

'[Click button Close on Request Submitted Popup]
Public Function clickBtnCloseRequestSubmitSendSMS()
   SendEmailSMS.btnClose.click
   If Err.Number<>0 Then
       clickBtnCloseRequestSubmitSendSMS = false
       LogMessage "WARN","Verification","Failed to Click Button : Close_RequestSubmitted" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickBtnCloseRequestSubmitSendSMS = true
End Function

'[Select the SMS Template Comboxbox as]
Public Function selectSMSTemplate_SendSMS(strSMSTemplate)
	bselectSMSTemplate_SendSMS=true
	If Not IsNull(strSMSTemplate) Then
       If Not (selectItem_Combobox (SendEmailSMS.lblSMSTemplate(), strSMSTemplate))Then
            LogMessage "WARN","Verification","Failed to select :"&strSMSTemplate&" From SMS Template Type drop down list" ,false
           bselectSMSTemplate_SendSMS=false
       End If
   End If
   WaitForICallLoading
   selectSMSTemplate_SendSMS=bselectSMSTemplate_SendSMS
End Function

'[Verify the SMS Template Comboxbox displayed as]
Public Function verifySMSTemplateDefault(strSMSTemplateDef)
   bDevPending=false
   verifySMSTemplateDefault = true
   If Not IsNull(strSMSTemplateDef) Then
       If Not verifyComboSelectItem (SendEmailSMS.lblSMSTemplate(),strSMSTemplateDef, "SMS Template")Then
           verifySMSTemplateDefault=false
       End If
   End If
End Function

'[set SMS Template ComboBox value as]
Public Function setSMSTemplateComBoxVal(strsetSMSTemplateComBoxVal)
	setSMSTemplateComBoxVal = True
	For iCountr = 1 To 180 Step 1
		If Not SendEmailSMS.txtSMSTemplate.Exist(0.5) Then
			Wait(0.5)
		else
			SendEmailSMS.txtSMSTemplate.Set strsetSMSTemplateComBoxVal
			Exit for
		End if
	Next
	If Err.Number<>0 Then
		setSMSTemplateComBoxVal = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	Else 
		LogMessage "RSLT","Verification","set value is Sccessfull",true
	End If
	WaitForIcallLoading	
End Function

'[Click the Overview contact details tab]
Public Function clickOverViewContactDetailsTab()
	For iCounter = 1 To 180 Step 1
		If Not bcCustomerOverview.weTabContactDetails.Exist(0.5) Then
			Wait(0.5)
		else
			bcCustomerOverview.weTabContactDetails.Click
			Exit for
		End if
	Next
	If Err.Number<>0 Then
       clickOverViewContactDetailsTab = false
       LogMessage "WARN","Verification","Failed to Click Contact Details Tab" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickOverViewContactDetailsTab = True
End Function

'[set Icomm SMS Email Enq frm Date text Box as]
Public Function IcommSMSEmailEnqfrmDatetextBox(strvaldate)
	IcommSMSEmailEnqfrmDatetextBox = True	
	If Not IsNull(strvaldate) Then
		IcommSMSEmailEnqfrmDatetextBox = SelectDatePicker_FromDate(strvaldate)
		'SendEmailSMS.txticommsmsEmailEnqfrmDatetextBox.Set strvaldate
	End If
	If Err.Number<>0 Then
		IcommSMSEmailEnqfrmDatetextBox = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	End If
	WaitForIcallLoading	
End Function

'[set Secount Icomm SMS Email Enq frm Date text Box as]
Public Function IcommTwoSMSEmailEnqfrmDatetextBox(strvaldateIs)
	IcommTwoSMSEmailEnqfrmDatetextBox = True	
	If Not IsNull(strvaldateIs) Then
		IcommTwoSMSEmailEnqfrmDatetextBox = SelectDatePicker_TODate(strvaldateIs)
		'SendEmailSMS.txticommsmsEmailEnqfrmDatetextBoxTwo.Set strvaldateIs
	End If
	If Err.Number<>0 Then
		IcommTwoSMSEmailEnqfrmDatetextBox = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	End If
	WaitForIcallLoading
End Function


'[Verify the Cancel Confirmation message displayed as]
Public Function verifyConfirmationPop_Cancel(strConfirmMsg)
   bverifyConfirmationPop=true
   If Not IsNull(strConfirmMsg) Then
       If Not verifyInnerText(SendEmailSMS.txtConfirmationMsg() , strConfirmMsg, "ConfirmationPopup")Then
			bverifyConfirmationPop = False
		End If
   End If
   verifyConfirmationPop_Cancel=bverifyConfirmationPop
End Function

'[Click Button Cancel on Send SMS Page]
Public Function clickButtonCancel_SendSMS()
   SendEmailSMS.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_SendSMS=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel_SendSMS=true
End Function

'[Click button Yes on Cancel Confirmation Popup]
Public Function clickBtnYesRequestCancelSendSMS()
   SendEmailSMS.btnYesCanPop.click
   If Err.Number<>0 Then
       clickBtnYesRequestCancelSendSMS = false
       LogMessage "WARN","Verification","Failed to Click Button : Cancel Confirmation" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickBtnYesRequestCancelSendSMS = true
End Function

'[Verify Contact Details Page is Displayed]
Public Function verifyContactDetailPag()
	verifyContactDetailPag = False
	For iCounters = 1 To 180 Step 1
		If Not SendEmailSMS.lblContactDetailsPage.Exist(0.5) Then
			Wait(0.5)
		else
			verifyContactDetailPag = True
			Exit for
		End if
	Next
	If Err.Number<>0 Then
       verifyContactDetailPag = false
       LogMessage "WARN","Verification","Failed to load Contact Details Tab" ,false
       Exit Function
   End If
   WaitForICallLoading
End Function

'[Enter Hours in the Hours Text box as]
Public Function setHours_SendSMS(strHours)
	setHours_SendSMS = True
	If Not IsNull(strHours) Then
		SendEmailSMS.txtHours.Set strHours
	End If
	If Err.Number<>0 Then
		setHours_SendSMS = false
		LogMessage "WARN","Verification","Failed to Set Text Box :Hours" ,false
		Exit Function
	End If
End Function

'[Select the Fee Type Comboxbox from Send SMS as]
Public Function selectSMSFeeTypeSendSMS(strSMSFeeType)
	selectSMSFeeTypeSendSMS=true
	If Not IsNull(strSMSFeeType) Then
       If Not (selectItem_Combobox (SendEmailSMS.lstFeeTypeSendSMS(), strSMSFeeType))Then
            LogMessage "WARN","Verification","Failed to select :"&strSMSFeeType&" From Fee Type drop down list" ,false
           selectSMSFeeTypeSendSMS=false
       End If
   End If
   WaitForICallLoading
End Function

'[set the alphanumeric filed value as]
Public Function txtAlphanumericval(strSetval)
	txtAlphanumericval = True	
	If Not IsNull(strSetval) Then
		SendEmailSMS.txtAlphanumeric.Set strSetval
	End If
	If Err.Number<>0 Then
		txtAlphanumericval = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	End If
	WaitForIcallLoading	
End Function

'[set the alphanumeric Secound filed value as]
Public Function txtAlphanumericvalSec(strSetvalSec)
	txtAlphanumericvalSec = True	
	If Not IsNull(strSetvalSec) Then
		SendEmailSMS.txtAlphanumericSecFld.Set strSetvalSec
	End If
	If Err.Number<>0 Then
		txtAlphanumericvalSec = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	End If
	WaitForIcallLoading	
End Function

'[verify inline validation message at Send SMS page]
Public Function lblInlineMessageSMSSend(strInlineMsg)
	lblInlineMessageSMSSend=true
	If Not IsNull(strInlineMsg) Then
       If Not verifyInnerText(SendEmailSMS.lblInlineMessageTxt() , strInlineMsg, "InlineMessage Validation")Then
			lblInlineMessageSMSSend = False
		End If
   End If
End Function

'[select Title val from the dropdown at send SMS Page]
Public Function lstTitleSendSMS(strTitleSendSMS)
	lstTitleSendSMS=true
	If Not IsNull(strTitleSendSMS) Then
       If Not (selectItem_Combobox (SendEmailSMS.lstTitleSendSMS(), strTitleSendSMS))Then
            LogMessage "WARN","Verification","Failed to select :"&strTitleSendSMS&" From Title drop down list" ,false
           lstTitleSendSMS=false
       End If
   End If
   WaitForICallLoading
End Function

'[Select Account Number at Send SMS Page as]
Public Function lstAccountNoSendSMS(strAccountNo)
	lstAccountNoSendSMS = True
	If Not IsNull(strAccountNo) Then
       If Not (selectItem_Combobox (SendEmailSMS.lstAccountNoSendSMS(), strAccountNo))Then
            LogMessage "WARN","Verification","Failed to select :"&strAccountNo&" From Account No drop down list" ,false
           lstAccountNoSendSMS=false
       End If
   End If
   WaitForICallLoading	
End Function

'[set the mobile number at Send SMS Page as]
Public Function txtMobileNOsendSMS(strMoblNoSMS)
	txtMobileNOsendSMS = True	
	If Not IsNull(strMoblNoSMS) Then
		SendEmailSMS.txtMobileNoSendSMS.Set strMoblNoSMS
	End If
	If Err.Number<>0 Then
		txtMobileNOsendSMS = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	End If
	WaitForIcallLoading	
End Function

'[Verify Field TM Approval Message on Send SMS Screen displayed as]
Public Function verifyTMApprovalMsgSendSMS(strValidationMessage)
   verifyTMApprovalMsgSendSMS=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (SendEmailSMS.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           verifyTMApprovalMsgSendSMS=false
       End If
   End If
   bcSTPStatementCopy.btnOK_ValidationPopup.Click
   If Err.Number<>0 Then
		verifyTMApprovalMsgSendSMS = false
		LogMessage "WARN","Verification","Set value is Failed" ,false
		Exit Function
	End If
   WaitForIcallLoading
End Function
