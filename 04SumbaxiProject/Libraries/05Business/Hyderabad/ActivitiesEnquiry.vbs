'[Verify CSO Home Page default options in Show dropdown displayed as]
Public Function VrfyDefValCSOShowDropdown(strDefValShowDropdownEQ)
   VrfyDefValCSOShowDropdown = true
   If Not IsNull(strDefValShowDropdownEQ) Then
       If Not verifyComboSelectItem (CSO_TM_Home.lstShow(),strDefValShowDropdownEQ, "Show")Then
    	  LogMessage "WARN","Verification","Expected Default Show type:"&strDefValShowDropdownEQ&" not displayed in the Show field" ,false
          VrfyDefValShowDropdownEQ = false
       End If
   End If
End Function

'[Verify CSO Home Page list of values displayed in Show dropdown]
Public Function VerifylistShowDrpDwn(lstShowDrpDwn) 
	VerifylistShowDrpDwn = True 
	If Not IsNull(lstShowDrpDwn) Then
       If Not verifyComboboxItems(CSO_TM_Home.lstShow(),lstShowDrpDwn, "Show")Then
       	   LogMessage "WARN","Verification","List of Show displayed in the combox box is not as expected" ,false
           VerifylistShowDrpDwn = false
       End If
   End If
End Function

'[Verify CSO Home Page row Data in Table Activity Records]
Public Function verifytblCSOHome_Activity_Records_RowData(arrRowDataList)
   bDevPending=false
   verifytblCSOHome_Activity_Records_RowData = verifyTableContentList(CSO_TM_Home.lblActivityTableHeader,CSO_TM_Home.lblActivityTableContent,arrRowDataList,"Activity_Records" ,True,CSO_TM_Home.lnkNext ,CSO_TM_Home.lnkNext1,CSO_TM_Home.lnkPrevious)
End Function

'[Click on Created On hyperlink from Activities Table]
Public Function ClickOnCreatedOnHyperlinkActie(lstCreatedOn)
    ClickOnCreatedOnHyperlinkActie = false
	ClickOnCreatedOnHyperlinkActie = selectTableLink(CSO_TM_Home.lblActivityTableHeader, CSO_TM_Home.lblActivityTableContent, lstCreatedOn, "Activities Details", "Created On", True, CSO_TM_Home.lnkNext ,CSO_TM_Home.lnkNext1, CSO_TM_Home.lnkPrevious)
	WaitForICallLoading
End Function

'[Verify Pink Panel displayed in New IA Page]
Public Function verifyNewIAPinkPanel(strName,strCIN,strSegment)
	verifyNewIAPinkPanel = false	
	If Not IsNull (strName) and Not IsNull (strCIN) and Not IsNull (strSegment) Then
	   If verifyInnerText(NewActivities.lblName(),strName, "Name") and verifyInnerText(NewActivities.lblCIN(),strCIN, "CIN") and verifyInnerText(NewActivities.lblSegment(),strSegment, "Segment") Then
		verifyNewIAPinkPanel = true
	   End If
	End If
End Function

'[Verify NEW IA Page Default Values for Selected Activities]
Public Function verifyNewIAPageDefaultValuesDetails(strRelatedTo, strAccountNumber, strType, strSubType, strProduct, strAssignedTo, strStatus, strResolution, strComments, strOnceDone)
	verifyNewIAPageDefaultValuesDetails = false	
	If Not IsNull(strRelatedTo) and Not IsNull(strAccountNumber) and Not IsNull(strType) and Not IsNull(strSubType) and Not IsNull(strProduct) and Not IsNull(strAssignedTo) and Not IsNull(strStatus) and Not IsNull(strResolution) and Not IsNull(strComments) Then
	    If verifyComboSelectItem (NewActivities.lstRelatedTo(),strRelatedTo, "Related To") and verifyComboSelectItem (NewActivities.lstAccount(),strAccountNumber, "Account Number") and verifyComboSelectItem (NewActivities.lstType(),strType, "Type") and verifyComboSelectItem (NewActivities.lstSubType(),strSubType, "Sub Type") and verifyComboSelectItem (NewActivities.lstProduct(),strProduct, "Product") and verifyComboSelectItem (NewActivities.lstAssignTo(),strAssignedTo, "Assigned To") and verifyComboSelectItem (NewActivities.lstStatus(),strStatus, "Status") and verifyComboSelectItem (NewActivities.lstResolution(),strResolution, "Resolution") and verifyInnerText (NewActivities.txtComment(),strComments, "Comments") and verifyInnerText (NewActivities.chkOnceDone(),strOnceDone, "Once and Done") Then
		   verifyNewIAPageDefaultValuesDetails = true
	   End If
	End If   
End Function

'[Verify New IA Page Default Values for Selected Activities Section]
Public Function verifyNewIAPageDftActvitieSection(strIANumber, strSRNumber, strChannel, strCreatedDate, strCreatedBy, strOverdue, strDuration, strClosedDate, strLastUpdatedDate, strLastUpdatedBy)
	verifyNewIAPageDftActvitieSection = false
	If Not IsNull(strIANumber) and Not IsNull(strSRNumber) and Not IsNull(strChannel) and Not IsNull(strCreatedDate) and Not IsNull(strCreatedBy) and Not IsNull(strOverdue) and Not IsNull(strDuration) and Not IsNull(strClosedDate) and Not IsNull(strLastUpdatedDate) and Not IsNull(strLastUpdatedBy) Then
		If verifyInnerText (NewActivities.lblIANumber(),strIANumber, "IA Number") and verifyInnerText (NewActivities.lblServiceRequestNo(),strSRNumber, "SR Number") and verifyInnerText (NewActivities.lblChannel(),strChannel, "Channel") and verifyInnerText (NewActivities.lblCreatedDate(),strCreatedDate, "Created Date") and verifyInnerText (NewActivities.lblCreatedBy(),strCreatedBy, "Created By") and verifyInnerText (NewActivities.lblOverDue(),strOverdue, "Over Due") and verifyInnerText (NewActivities.lblDuration(),strDuration, "Duration") and verifyInnerText (NewActivities.lblClosedDate(),strClosedDate, "Closed Date") and verifyInnerText (NewActivities.lblLastUpdateDate(),strLastUpdatedDate, "Last Updated date") and verifyInnerText (NewActivities.lstLastUpdatedBy(),strLastUpdatedBy, "Last Updated By") Then
			verifyNewIAPageDftActvitieSection = true
		End If
	End If
End Function

'[Verify the shortcut buttons on New IA page]
Public Function verifyShortCutButtonNewIAPage()
	verifyShortCutButtonNewIAPage = false	
	If NewActivities.btnAttachments.Exist(1) and NewActivities.btnCancel.Exist(1) Then
		LogMessage "RSLT","Verification","Attachments and Cancel button is available for New IA Page." ,True
		verifyShortCutButtonNewIAPage = true
	Else
	    LogMessage "WARN","Verification","Attachments or Cancel button is not available for New IA Page. Expected to be Enabled." ,false
	End If
	WaitForICallLoading
End Function

'[Verify CSO Home Page default options in Group Select and Staff dropdown displayed as]
Public Function VrfyDefValCSOGroupSelStaffDropdown(strDefValCSOGroupDropdown, strDefValCSOSelectDropdown, strDefValCSOStaffDropdown)
   VrfyDefValCSOGroupSelStaffDropdown = true
   If Not IsNull(strDefValCSOGroupDropdown) and Not ISNull(strDefValCSOSelectDropdown) and Not IsNull(strDefValCSOStaffDropdown) Then
       If Not verifyComboSelectItem (CSO_TM_Home.lblGroupDropdown(),strDefValCSOGroupDropdown, "Group") and Not verifyComboSelectItem (CSO_TM_Home.lblSelectDropdown(),strDefValCSOSelectDropdown, "Select") and Not verifyComboSelectItem (CSO_TM_Home.lblStaffDropdown(),strDefValCSOStaffDropdown, "Staff") Then
    	  LogMessage "WARN","Verification","Expected Default Group Dropdown:"&strDefValCSOGroupDropdown& " Select "& strDefValCSOSelectDropdown & " Staff"&strDefValCSOStaffDropdown&" not displayed in the Group field" ,false
          VrfyDefValCSOGroupSelStaffDropdown = false
       End If
   End If
End Function

'[Verify Default Activities From and To Date Transaction Date]
Public Function verifyDefaultActivitiesTransDate()
  verifyDefaultActivitiesTransDate = true
  strActivitiesFromDate = CSO_TM_Home.txtFrom.GetROProperty("value")
  strActivitiesToDate = CSO_TM_Home.txtTo.GetROProperty("value")
  DaysRange = DateDiff("d", strActivitiesFromDate, strActivitiesToDate)
  If DaysRange = 30  Then
     LogMessage "RSLT","Verification","From Date:"+ trim(strActivitiesFromDate) +"and To Dates:"+ trim(strActivitiesToDate) +" are displayed diff by 30 days default range",True
     verifyDefaultTransDate = True 
  Else 
     LogMessage "WARN", "Verification","From Date:"+ trim(strActivitiesFromDate) +"and To Dates:"+ trim(strActivitiesToDate) +" are not displayed diff by 30 days default range", False
     verifyDefaultTransDate = False 
  End If
  WaitForICallLoading
End Function

'[Verify the Date Range exceeds 30 days Error Msg]
Public Function verifyDateRangeExceedsErrMsg(strDateRangeExceedsErrMsg)
   verifyDateRangeExceedsErrMsg = true
   If Not IsNull(strDateRangeExceedsErrMsg) Then
       If Not VerifyInnerText (CSO_TM_Home.lblTODateExceedsError(), strDateRangeExceedsErrMsg, "the Date Range exceeds 30 days")Then
           verifyDateRangeExceedsErrMsg=false
       End If
   End If
End Function

'[Verify To Date Range cannot be future dated Error Msg]
Public Function verifyToDateRangeCantFuturDatedErrMsg(strToDateRangeCantFuturDatedErrMsg)
   verifyToDateRangeCantFuturDatedErrMsg = true
   If Not IsNull(strToDateRangeCantFuturDatedErrMsg) Then
       If Not VerifyInnerText (CSO_TM_Home.lblToFutDateErr(), strToDateRangeCantFuturDatedErrMsg, "To Date Range cannot be future dated")Then
           verifyToDateRangeCantFuturDatedErrMsg = false
       End If
   End If
End Function

'[Verify From Date Range cannot be future dated Error Msg]
Public Function verifyFromDateRangeCantFuturDatedErrMsg(strFromDateRangeCantFuturDatedErrMsg)
   verifyFromDateRangeCantFuturDatedErrMsg = true
   If Not IsNull(strFromDateRangeCantFuturDatedErrMsg) Then
       If Not VerifyInnerText (CSO_TM_Home.lblFromFutDateErr(), strFromDateRangeCantFuturDatedErrMsg, "From Date Range cannot be future dated")Then
           verifyFromDateRangeCantFuturDatedErrMsg = false
       End If
   End If
End Function


'[Verify To Date Range cannot be less than FROM date Error Msg]
Public Function verifyToDateRangeCantlessFromeDateErrMsg(strToDateRangeCantlessFromeDateErrMsg)
   verifyToDateRangeCantlessFromeDateErrMsg = true
   If Not IsNull(strToDateRangeCantlessFromeDateErrMsg) Then
       If Not VerifyInnerText (CSO_TM_Home.lblToLessDateError(), strToDateRangeCantlessFromeDateErrMsg, "To Date Range cannot be less than FROM date")Then
           verifyToDateRangeCantlessFromeDateErrMsg = false
       End If
   End If
End Function

'[Verify View IA Page Default Values for Selected Activities]
Public Function VerifyViewIAFDisplayValues(strRelatedTo, strAccountNumber, strType, strSubType, strProduct, strAssignedTo, strStatus, strResolution, strDueDate )
	VerifyViewIAFDisplayValues = false
	If verifyRelatedToText_IA(strRelatedTo) and verifyCardNumberText_IA(strAccountNumber) and verifyTypeText_IA(strType) and verifySubTypeText_IA(strSubType) and verifyProductText_IA(strProduct) and verifyAssignedToText_IA(strAssignedTo) and verifyStatus_ViewIAText(strStatus) and verifyResolutionText_IA(strResolution) and verifyDueDateText_IA(strDueDate) Then
		VerifyViewIAFDisplayValues = true
	End If
End Function

'[Verify View IA Page Default Values for Selected Activities Section]
Public Function VerifyViewIASectionDisplayValues(strIANumber, strSRNumber, strChannel, strCreatedDate, strCreatedBy, strOverdue, strDuration, strClosedDate, strLastUpdatedDate, strLastUpdatedBy)
    VerifyViewIASectionDisplayValues = false
    If verifyIANumber_ViewIA(strIANumber) and verifySRNumber_ViewIA(strSRNumber) and verifyChannelText_IA(strChannel) and VerifyFieldCreatedDate_IA(strCreatedDate) and VerifyFieldCreatedBy_IA(strCreatedBy) and verifyOverdueText_IA(strOverdue) and verifyDurationText_IA(strDuration) and VerifyFieldClosedDate_IA(strClosedDate) and VerifyFieldLastUpdatedDate_IA(strLastUpdatedDate) and VerifyFieldLastUpdateBy_IA(strLastUpdatedBy) Then
	   VerifyViewIASectionDisplayValues = true
    End If
End Function
