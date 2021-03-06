'[Verify the fields of events details displayed as]
Public Function verifyFields_EventDetails(lstEventsDetails)
	bverifyFields_EventDetails = true
	intSize = Ubound(lstEventsDetails)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstEventsDetails(Iterator),":")(0))
		arrValue = trim(Split(lstEventsDetails(Iterator),":")(1))
		
	Select Case (arrLabel)
		Case "Event Type"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcEvents.lblEventType(), arrValue, "Event Type")Then
				LogMessage "RSLT","Verification","Event Details - Event Type:"&arrValue&" is not displayed as expected",false
				bverifyFields_EventDetails=false
			End If
		End If
		
		Case "Date"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcEvents.lblDate_Events(), arrValue, "Date")Then
				LogMessage "RSLT","Verification","Event Details - Date:"&arrValue&" is not displayed as expected",false
				bverifyFields_EventDetails=false
			End If
		End If
		
		Case "Time"
		If Not IsNull(arrValue) Then
		arrvalue_new = Replace (arrValue,"@",":")
			If Not VerifyInnerText (bcEvents.lblTime_Events(), arrvalue_new, "Time")Then
				LogMessage "RSLT","Verification","Event Details - Time:"&arrValue&" is not displayed as expected",false
				bverifyFields_EventDetails=false
			End If
		End If
		
		Case "Transaction Code"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcEvents.lblTransactionCode_Events(), arrValue, "Transaction Code")Then
				LogMessage "RSLT","Verification","Event Details - Transaction Code:"&arrValue&" is not displayed as expected",false
				bverifyFields_EventDetails=false
			End If
		End If
		Case "Transaction Description"
		If Not IsNull(arrValue) Then
			If Not VerifyInnerText (bcEvents.lblTransactionDesc_Events(), arrValue, "Transaction Description")Then
				LogMessage "RSLT","Verification","Event Details - Transaction Description:"&arrValue&" is not displayed as expected",false
				bverifyFields_EventDetails=false
			End If
		End If
		End Select
		Next
		verifyFields_EventDetails = bverifyFields_EventDetails
End Function

'[Verify the field SMS messge sent and Other details displayed as]
Public Function verifySMSandOtherDetails_Event(strSMSMessage,strOtherDetails,strSuggestedAction)

	'I.Serve field displayed as
	strIserveSMSMessage = bcEvents.txtSMSMessageSent_Events.GetROProperty("innertext")
	strIserveOtherDetails = bcEvents.txtOtherdetails_Events.GetROProperty("innertext")
	strIserveSuggestedAction = bcEvents.txtSuggestedActions_Events.GetROProperty("innertext")
	bverifySMSandOtherDetails_Event = true
	
		If strSMSMessage = strIserveSMSMessage Then
	  	LogMessage "RSLT","Verification","The Iserve SMS message sent is displayed as: "&strSMSMessage&"",True
			Else
	  	LogMessage "RSLT","Verification","The Iserve SMS message sent is not displayed as:: "&strSMSMessage&"",False
		End if 
		
		If strOtherDetails = strIserveOtherDetails Then
	  	LogMessage "RSLT","Verification","The Iserve Other details is displayed as: "&strOtherDetails&"",True
			Else
	  	LogMessage "RSLT","Verification","The Iserve Other details is not displayed as:: "&strOtherDetails&"",False
		End if
		
		If strSuggestedAction = strIserveSuggestedAction Then
	  	LogMessage "RSLT","Verification","The Iserve Suggested Action is displayed as: "&strSuggestedAction&"",True
			Else
	  	LogMessage "RSLT","Verification","The Iserve Suggested Action is not displayed as:: "&strSuggestedAction&"",False
		End if
		verifySMSandOtherDetails_Event = bverifySMSandOtherDetails_Event
End Function

'[Set the comment text area in events]
Public Function setComments_Events(strComment)
	bDevPending=False
	bcEvents.txtComments_Events.set strComment
	If Err.Number<>0 Then
       setComments_Events=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comments" ,false
       Exit Function
   End If
	setComments_Events = bsetComments_Events
End Function

'[Click on the save button of the events page]
Public Function clickbuttonSave_Events()
	bDevPending=False
   bcEvents.btnSave_Events.click
   If Err.Number<>0 Then
       clickbuttonSave_Events=false
            LogMessage "WARN","Verification","Failed to Click Button : Save" ,false
       Exit Function
   End If
   clickbuttonSave_Events=true
End Function

'[Click on the Cancel button of the events page]
Public Function clickbuttonCancel_Events()
	bDevPending=False
   bcEvents.btnCancel_Events.click
   If Err.Number<>0 Then
       clickbuttonCancel_Events=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickbuttonCancel_Events=true
End Function

'[Verify the button save is disabled once it is saved]
Public Function verifybuttonSaveDisabled_Events()
 bverifybuttonSaveDisabled_Events = true
 'Getting the enabled property of the object
 enabled_Obj = bcEvents.btnSave_Events.GetROProperty("disabled")
 
 If enabled_Obj= 0 Then
 	LogMessage "RSLT","Verification","The button Save is disabled as expected",True
 	else
 	LogMessage "RSLT","Verification","The button Save is not disabled as expected",False
 End If
	verifybuttonSaveDisabled_Events = bverifybuttonSaveDisabled_Events
End Function

'[Select the radiobutton call prediction in the events page]
Public Function selectRbtnCallPred_Events(strCallPrdction)
	bselectRbtnCallPred_Events=true
	intRadio_Events=Instr(bcEvents.radiogrpCallPrediction_Events.GetROproperty("class"),"disabled-area")
	If intRadio_Events = 0 Then
		bselectRadioButton_rewards=SelectRadioButtonGrp(strCallPrdction,bcEvents.radiogrpCallPrediction_Events, Array("Yes","No"))
	Else
		LogMessage "RSLT","Verifiation","Radio button is disabled by default.",True
	End If
	If Err.Number<>0 Then
       bselectRbtnCallPred_Events=false
          LogMessage "WARN","Verification","Failed to select radiobutton or radiobutton is disabled" ,True
       Exit Function
   End If
   selectRbtnCallPred_Events = bselectRbtnCallPred_Events
End Function

'[Verify the row data for events table displayed as]
Public Function verifyrowdata_Events(arrRowDataList)
	bverifyrowdata_Events = true
	verifyrowdata_Events = verifyTableContentList(bcEvents.tblEvents_header,bcEvents.tblEvents_Content,arrRowDataList,"Events",false,Null,Null,Null)
	verifyrowdata_Events = bverifyrowdata_Events
End Function

'[Click the events type link from the events table]
Public Function clickEventType_Events(lstRowData)
	bclickEventType_Events = true
	clickEventType_Events = selectTableLink(bcEvents.tblEvents_header,bcEvents.tblEvents_Content,lstRowData,"Events","Event Type",false,Null,Null,Null)
	clickEventType_Events = bclickEventType_Events
End Function
