'[Verify SMS Email Enquiry Link Enable Disable State]
Public Function verifySMSEmailEnquirylinkState(strstatus)
    bVerifyState=True
	If Not IsNull(strstatus) Then
		If coSmsEmailEnquiry_Page.lnkSmsEmailEnquiry.Exist(0) Then
			stractStatus = coSmsEmailEnquiry_Page.lnkSmsEmailEnquiry.GetRoProperty("disabled")
			If strstatus="Enable" Then
				If stractStatus>0 Then
					LogMessage "RSLT","Verification","Special Offer link is in Enabled Mode", True
				Else
					LogMessage "RSLT","Verification","Special Offer link is in Disabled Mode", False
					bVerifyState=False
				End If
			ElseIf strstatus="Disable" Then
				If stractStatus=0 Then
					LogMessage "RSLT","Verification","Special Offer link is in Disabled Mode", True
				Else
					LogMessage "RSLT","Verification","Special Offer link is in Enabled Mode", False
					bVerifyState=False
				End If
			Else
			bVerifyState=False
			End If
		Else
		   bVerifyState=False
		End If
	Else
	bVerifyState=False
	End If
	verifySMSEmailEnquirylinkState=bVerifyState
End Function

'[Verify SMS Related details label displayed as]
Public Function VerifySmsRelatedDetails(strSmsRelated)
	bVerify=True	
	If Not IsNull(strSmsRelated) Then	
		If Not verifyInnerText(coSmsEmailEnquiry_Page.lblSmsRelated(),strSmsRelated,"SMS Related Details")Then
				bVerify = False
			End If
	End If				
	VerifySmsRelatedDetails=bVerify	
End Function

'[Verify table row Registered Mobile Number displayed as]
Public Function VerifytblRegMobNo(lstRegMobNo)
	VerifytblSelectedCard_CC = VerifyTableSingleRowData(coSmsEmailEnquiry_Page.tblRegMobNoListHeader,coSmsEmailEnquiry_Page.tblRegMobNoListBody,lstRegMobNo,"REGISTERED MOBILE NUMBER")
End Function

'[Verify list of values displayed in SMS Category dropdown as]
Public Function VerifylstSmsCategoryDropDwn(lstSmsCategory)
	bVerifySmsCate=True
	If Not IsNull(lstSmsCategory) Then
		bVerifySmsCate=verifyComboboxItems1(coSmsEmailEnquiry_Page.drpSmsCategory,coSmsEmailEnquiry_Page.lstSmsCategoryObj,lstSmsCategory,"SMS Category")	
	End If
	VerifylstSmsCategoryDropDwn=bVerifySmsCate	
End Function

'[Select SMS Category dropdown as]
Public Function SetSmsCategoryDropDwn(strSmsCategory)
	SetSmsCategoryDropDwn=True
	If Not IsNull(strSmsCategory) Then
		SetSmsCategoryDropDwn=SelectComboBoxItem(coSmsEmailEnquiry_Page.setSmsCategory,strSmsCategory,"SMS Category")
	End If
	If Err.Number <> 0 Then 
		LogMessage "WARN","Verification","Failed to Set SMS Category", False
		SetSmsCategoryDropDwn=False
	End If
End Function

'[Verify list of values displayed in SMS Template dropdown as]
Public Function VerifylstSmsTemplateDropDwn(lstSmsTemplatedrp)
	bVerifySmsTemp=True
	If Not IsNull(lstSmsTemplatedrp) Then
		bVerifySmsTemp=verifyComboboxItems(coSmsEmailEnquiry_Page.lstSmsTemplate,lstSmsTemplatedrp,"SMS Template")	
	End If
	VerifylstSmsTemplateDropDwn=bVerifySmsTemp	
End Function

'[Select SMS Template dropdown as]
Public Function SetSmsTemplateDropDwn(strSmsTemplate)
	scrollPageDown 2
	SetSmsTemplateDropDwn=True
	If Not IsNull(strSmsTemplate) Then
		SetSmsTemplateDropDwn=SelectComboBoxItem(coSmsEmailEnquiry_Page.lstSmsTemplate,strSmsTemplate,"SMS Template")
	End If
	If Err.Number <> 0 Then 
		LogMessage "WARN","Verification","Failed to Set SMS Template", False
		SetSmsTemplateDropDwn=False
	End If
End Function

'[Verify Send SMS Label SMS Template ID as]
Public Function vrfySmsTemplateIDlable(strSmsTemplateIDlable)
	bVerifySmsTemplateIDlable=True	
	If Not IsNull(strSmsTemplateIDlable) Then	
		If Not verifyInnerText(coSmsEmailEnquiry_Page.lblSmsTemplateId(),strSmsTemplateIDlable,"Send SMS Template ID label Name")Then
				bVerifySmsTemplateIDlable = False
			End If
	End If				
	vrfySmsTemplateIDlable=bVerifySmsTemplateIDlable	
End Function

'[Verify Send SMS Template ID Value]
Public Function vrfySmsTemplateIDVal(strSmsTemplateIDval)
	bVerifySmsTemplateIDVal=True
	If Not IsNull(strSmsTemplateIDval) Then
		If Not verifyInnerText(coSmsEmailEnquiry_Page.txtSmsTemplateId(),strSmsTemplateIDval,"Send SMS Template ID value")Then
				bVerifySmsTemplateIDVal=False
			End If
	End If	
	vrfySmsTemplateIDVal=bVerifySmsTemplateIDVal
End Function

'[Verify Send SMS Label SMS Offer Code as]
Public Function vrfySmsOfferCodeLbl(strSmsOfferCodelbl)
	bvrfySmsOfferCodeLbl=True
	If Not IsNull(strSmsOfferCodelbl) Then
		If Not verifyInnerText(coSmsEmailEnquiry_Page.lblSmsOfferCode(),strSmsOfferCodelbl,"Send SMS Offer Code lable Name")Then
				bvrfySmsOfferCodeLbl=False
			End If
	End If	
	vrfySmsOfferCodeLbl=bvrfySmsOfferCodeLbl
End Function

'[Verify Send SMS Offer code Value]
Public Function vrfySmsOfferCodeVal(strSmsOfferCodeVal)
	bvrfySmsOfferCodeVal=True
	If Not IsNull(strSmsOfferCodeVal) Then
		If Not verifyInnerText(coSmsEmailEnquiry_Page.txtSmsOfferCode(),strSmsOfferCodeVal,"Send SMS Offer Code Value")Then
				bvrfySmsOfferCodeVal=False
			End If
	End If	
	vrfySmsOfferCodeVal=bvrfySmsOfferCodeVal
End Function

'[Verify Send SMS Label SMS Text as]
Public Function vrfySmsTxtlbl(strSmsTxtlbl)
	bvrfySmsTxtlbl=True
	If Not IsNull(strSmsTxtlbl) Then
		If Not verifyInnerText(coSmsEmailEnquiry_Page.lblSmsText(),strSmsTxtlbl,"Send SMS Text lable Name")Then
				bvrfySmsTxtlbl=False
			End If
	End If	
	vrfySmsTxtlbl=bvrfySmsTxtlbl
End Function

'[Verify Send SMS Text Value]
Public Function vrfySmsTxtVal(strSmsTxtVal)
	bvrfySmsTxtVal=True
	If Not IsNull(strSmsTxtVal) Then
		If Not verifyInnerText(coSmsEmailEnquiry_Page.txtSmsText(),strSmsTxtVal,"Send SMS Text Value")Then
				bvrfySmsTxtVal=False
			End If
	End If	
	vrfySmsTxtVal=bvrfySmsTxtVal
End Function

'[Enter Comments in Optional textbox field displayed]
Public Function SetCommentEditboxOptional(strComment)
    SetCommentEditboxOptional=True
	If Not IsNull(strComment) Then
		coSmsEmailEnquiry_Page.txtCommentOptional.Set strComment
		WaitForIServeLoading
	End If
	If Err.Number<>0 Then
	   SetCommentEditboxOptional=False
	   LogMessage "WARN","Verification","Failed to Set Edit Box : Comments Optional" , False
	   Exit Function
	End If
End Function

'[Verify list of values displayed in Fee Type dropdown as]
Public Function vrfylstSmsFeeTypeDropDwn(lstSmsFeeType)
	bVerifySmsFeeType=True
	If Not IsNull(lstSmsFeeType) Then
		bVerifySmsFeeType=verifyComboboxItems1(coSmsEmailEnquiry_Page.drpSmsFeeType,coSmsEmailEnquiry_Page.lstSmsCategoryObj,lstSmsFeeType,"SMS Fee Type")	
	End If
	vrfylstSmsFeeTypeDropDwn=bVerifySmsFeeType	
End Function

'[Select SMS Fee Type dropdown as]
Public Function SetSmsFeeTypeDropDwn(strSmsFeeType)
	SetSmsFeeTypeDropDwn=True
	If Not IsNull(strSmsFeeType) Then
		SetSmsCategoryDropDwn=SelectComboBoxItem(coSmsEmailEnquiry_Page.setSmsFeeType,strSmsFeeType,"SMS Fee Type")
	End If
	If Err.Number <> 0 Then 
		LogMessage "WARN","Verification","Failed to Set SMS Fee Type", False
		SetSmsFeeTypeDropDwn=False
	End If
End Function

'[Select Date using Date Picker in Send SMS Page]
Public Function SelectDateSendSMS(strFromDate)
	bverifyDate = True
	If Not IsNull(strFromDate) Then	
		If Trim(strFromDate) = "TODAY" Then
			strFromDate = Day(Now) & " " & MonthName(Month(Now),True) &" "& Year(Now)
		End If		
		bverifyDate = SelectDateFromIDCalendar(coSmsEmailEnquiry_Page.txtSendSMSDate,strFromDate)
		strExpFromDate = Right("0" & Datepart("d",strFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strFromDate),2))&" " & Year(strFromDate)		
		If bverifyDate Then
		
			strActFromDate = coSmsEmailEnquiry_Page.txtSendSMSDate.GetROProperty("value")
			strActFromDate = Right("0" & Datepart("d",strActFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strActFromDate),2))&" " & Year(strActFromDate)
			
			If Trim(strActFromDate) = Trim(strExpFromDate) Then
			   LogMessage "RSLT","Verification","Selected date "&strFromDate&" in date text box is displayed as expected", True
			   bverifyDate = True 
			Else
				LogMessage "WARN","Verification","As expected, Selected date "&strFromDate&" in date text box is not displayed.", False
			   bverifyDate = False 
			End If			
		End If		
	End If	
	SelectDateSendSMS = bverifyDate
End Function
