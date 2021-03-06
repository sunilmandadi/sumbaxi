'[Verify existenence of fields displayed in Search panel]
Public Function VerifyFieldExistenece_IA()
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblStatus(),"Dashboard IA","Status")
		If bVerifyfields <> 0  Then
		   LogMessage "WARN","Verification","Field label: Status not displayed as expected", False
		   VerifyFieldExistenece_IA = bVerifyfields
		   Exit Function 
		End If
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblGroup(),"Dashboard IA","Group")
		If bVerifyfields <> 0  Then
		   LogMessage "WARN","Verification","Field label: Group not displayed as expected", False
		   VerifyFieldExistenece_IA = bVerifyfields
		   Exit Function 
		End If
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblSelect(),"Dashboard IA","Select")
		If bVerifyfields <> 0  Then
		   LogMessage "WARN","Verification","Field label: Select not displayed as expected", False
		   VerifyFieldExistenece_IA = bVerifyfields
		   Exit Function 
		End If
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblStaff(),"Dashboard IA","Staff")
		If bVerifyfields <> 0  Then
		   LogMessage "WARN","Verification","Field label: Staff not displayed as expected", False
		   VerifyFieldExistenece_IA = bVerifyfields
		   Exit Function 
		End If
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblSearchChannel(),"Dashboard IA","Channel")
		If bVerifyfields <> 0  Then
		   LogMessage "WARN","Verification","Field label: Channel not displayed as expected", False
		   VerifyFieldExistenece_IA = bVerifyfields
		   Exit Function 
		End If	
VerifyFieldExistenece_IA =  bVerifyfields
End Function

'[Verify list of values in Status dropdown displayed in Search panel]
Public Function VerifyStatusDropdown_IA(lstStatus)
bVerifyValues = False
If Not IsNull(lstStatus) Then
bVerifyValues = verifyComboboxItems(coDashboard_IA_Page.lstComboBoxStatus(),lstStatus,"Status")		
End If
VerifyStatusDropdown_IA = bVerifyValues
End Function

'[Verify list of values in Group dropdown displayed in Search panel]
Public Function VerifyGroupDropdown_IA(lstGroup)
bVerifyValues = False
If Not IsNull(lstGroup) Then
bVerifyValues = verifyComboboxItems(coDashboard_IA_Page.lstComboBoxGroup(),lstGroup,"Group")		
End If
VerifyGroupDropdown_IA = bVerifyValues
End Function

'[Verify list of values in Select dropdown displayed in Search panel]
Public Function VerifySelectDropdown_IA(lstSelect)
bVerifyValues = False
If Not IsNull(lstSelect) Then
bVerifyValues = verifyComboboxItems(coDashboard_IA_Page.lstComboBoxSelect(),lstSelect,"Select")		
End If
VerifySelectDropdown_IA = bVerifyValues
End Function

'[Verify display of channel field Search panel]
Public Function VerifyDropdownDisable_IA(status)
bVerifyValues = False
If Not IsNull(status) Then
bVerifyValues = VerifyObjectEnabledDisabled(coDashboard_IA_Page.lstComboBoxChannel(),status,"Channel Dropdown")		
End If
VerifyDropdownDisable_IA = bVerifyValues
End Function

'[Verify default value displayed in the dropdown]
Public Function VerifyDefaultDropdown_IA(StrStatus,StrGroup,StrSelect,StrStaff,StrChannel)
bVerifyDefaultvalue = False
	If Not IsNull(StrStatus) Then
		bVerifyStatus = VerifyDropdownDefaultValue(coDashboard_IA_Page.lstComboBoxStatus,StrStatus,"Status")
	End IF
	If Not IsNull(StrGroup) Then
		bVerifyGrp = VerifyDropdownDefaultValue(coDashboard_IA_Page.lstComboBoxGroup,StrGroup,"Group")
	End IF
	If Not IsNull(StrSelect) Then
		bVerifySelect = VerifyDropdownDefaultValue(coDashboard_IA_Page.lstComboBoxSelect,StrSelect,"Select")
	End IF 
	If Not IsNull(StrStaff) Then
		bVerifystaff = VerifyDropdownDefaultValue(coDashboard_IA_Page.lstComboBoxStaff,StrStaff,"Staff")
	End IF 
	If Not IsNull(StrChannel) Then
		bVerifyChannel = VerifyDropdownDefaultValue(coDashboard_IA_Page.lstComboBoxChannel,StrChannel,"Channel")
	End IF 
	If bVerifyStatus and bVerifyGrp and bVerifySelect and bVerifystaff and bVerifyChannel Then
		bVerifyDefaultvalue = True
	Else
		bVerifyDefaultvalue = False
	End If
VerifyDefaultDropdown_IA = bVerifyDefaultvalue	
End Function

'[Select Combobox Status in Dashboard IA Search Panel]
Public Function SelectStatuscombobox_DashboardIA(strItem)
	WaitForIServeLoading
	coDashboard_IA_Page.tabInteractionActivities.Click
	bVerify = True
	strExpitem = coDashboard_IA_Page.lstComboBoxStatus.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coDashboard_IA_Page.lstComboBoxStatus,strItem,"Status")
	End If
	SelectStatuscombobox_DashboardIA = bVerify
End Function

'[Select Group Combobox in Search Panel]
Public Function SelectGroupcombobox_IA(strItem)	
	SelectGroupcombobox_IA = SelectComboBoxItem(coDashboard_IA_Page.lstComboBoxGroup,strItem,"Group")
End Function

'[Set Select Combobox in Search Panel]
Public Function SetSelectcombobox_IA(strItem)
	WaitForIServeLoading
	bVerify = True
	strExpitem = coDashboard_IA_Page.lnpSelect.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coDashboard_IA_Page.lnpSelect,strItem,"Select")
	End If
	SetSelectcombobox_IA = bVerify
End Function

'[Select Staff Combobox in Search Panel]
Public Function SelectStaffcombobox_IA(strItem)
	SelectStaffcombobox_IA = SelectComboBoxItem(coDashboard_IA_Page.lstComboBoxStaff,strItem,"Staff")
End Function

'[Set Staff Combobox in Search Panel]
Public Function SetStaffcombobox_IA(strItem)
	WaitForIServeLoading
	bVerify = True
	strExpitem = coDashboard_IA_Page.lnpStaff.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coDashboard_IA_Page.lnpStaff,strItem,"Staff")
	End If	
	SetStaffcombobox_IA = bVerify
End Function

'[Select From Date using Date Picker in Search Panel]
Public Function SelectFromDate_IA(strFromDate)
bverifyDate = True

If Not IsNull(strFromDate) Then

	If Trim(strFromDate) = "TODAY" Then
	  strFromDate = Day(Now) & " " & MonthName(Month(Now),True) &" "& Year(Now)
	End If
	WaitForIServeLoading
	SelectFromDate_IA =  SelectDateFromIDCalendar(coDashboard_IA_Page.txtFromDate,strFromDate)
	strExpFromDate = Right("0" & Datepart("d",strFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strFromDate),2))&" " & Year(strFromDate)
	
	If SelectFromDate_IA Then	
		strActFromDate = coDashboard_IA_Page.txtFromDate.GetROProperty("value")
		strActFromDate = Right("0" & Datepart("d",strActFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strActFromDate),2))&" " & Year(strActFromDate)
		
		If Trim(strActFromDate) = Trim(strExpFromDate) Then
		   LogMessage "RSLT","Verification","Selected date "&strFromDate&" in From date text box is displayed as expected", True
		   bverifyDate = True 
		Else
		   bverifyDate = False 
		End If	
		
	End If

End If

SelectFromDate_IA = bverifyDate
End Function

'[Select TO Date using Date Picker in Search Panel]
Public Function SelectTODate_IA(strTODate)
bverifyDate = True 

If Not IsNull(strTODate) Then

	If Trim(strTODate) = "TODAY" Then
	   strTODate = Day(Now) & " " & MonthName(Month(Now),True) &" "& Year(Now)
	End If
	
	SelectTODate_IA =  SelectDateFromIDCalendar(coDashboard_IA_Page.txtToDate,strTODate)
	StrExpToDate = Right("0" & Datepart("d",strTODate),2) &" "& MonthName(Right("0" & Datepart("m",strTODate),2))&" " & Year(strTODate)
	
	If SelectTODate_IA Then
	
		strActTODate = coDashboard_IA_Page.txtToDate.GetROProperty("value")
		strActTODate = Right("0" & Datepart("d",strActTODate),2) &" "& MonthName(Right("0" & Datepart("m",strActTODate),2))&" " & Year(strActTODate)
		
		If Trim(strActTODate) = Trim(StrExpToDate) Then
		   LogMessage "RSLT","Verification","Selected date "&strTODate&" in TO date text box is displayed as expected", True
		   bverifyDate = True 
		Else
		  bverifyDate = False 
		End If	
		
	End If
	
End IF
SelectTODate_IA = bverifyDate
End Function

'[Verify default From and To date displayed in Dashboard IA Search panel]
Public Function verifyDefaultDateRange_IA(strDateRange,StrToDate)	
verifyDefaultDateRange = VerifyDateRange(coDashboard_IA_Page.txtFromDate,coDashboard_IA_Page.txtToDate,strDateRange,StrToDate)
End Function

'[Verify display of Submit button in Search panel]
Public Function VerifyButtonSubmit_IA(strCheckFlag)
VerifyButtonSubmit_IA = VerifyObjectEnabledDisabled(coDashboard_IA_Page.btnSubmit,strCheckFlag,"Submit Button")
End Function

'[Verify display of New IA button in Search panel]
Public Function VerifyButtonNewIA_IA(strCheckFlag)
VerifyButtonNewIA_IA = VerifyObjectEnabledDisabled(coDashboard_IA_Page.btnNewIA,strCheckFlag,"NewIA Button")
End Function

'[Click on Submit Button in Search Panel]
Public Function clickButtonSubmitSearchPanel()
 coDashboard_IA_Page.btnSubmit.click 
  If Err.Number <> 0 Then
      clickButtonSubmitSearchPanel = False
      LogMessage "WARN","Verification","Failed to Click Button : Submit", False
      Exit Function
  Else
  	clickButtonSubmitSearchPanel = True
  End If
  WaitForIServeLoading 
End Function

'[Click on New IA Button in Search Panel]
Public Function clickButtonNewIA_IA()
  coDashboard_IA_Page.btnNewIA.click 
  If Err.Number <> 0 Then
      clickButtonNEWIA_IA = False
      LogMessage "WARN","Verification","Failed to Click Button : NEW IA", False
      Exit Function
  Else
  	  clickButtonNewIA_IA = True
  End If
  WaitForIServeLoading
End Function

'[Verify Inline error message displayed in Search panel]
Public Function VerifyInlineErrorMsg_IA(strErrorMsg)
bverifyInlineErrorMsg = False
If VerifyInnerText(coDashboard_IA_Page.lblInlineMessage(), strErrorMsg, "Inline Date Error") Then
   bverifyInlineErrorMsg = True
End If
VerifyInlineErrorMsg_IA = bverifyInlineErrorMsg
End Function

'[Verify records displayed in IA Summary table based on Selected From and To Date]
Public Function VerifyRecordDisplayedBasedonDates_IA()
strFromDate = coDashboard_IA_Page.txtFromDate.GetROProperty("value")
strToDate = coDashboard_IA_Page.txtToDate.GetROProperty("value")
VerifyRecordDisplayedBasedonDates_IA = VerifyDateSearchRecordsdisplayed(coDashboard_IA_Page.tblDashboardIAHeader,coDashboard_IA_Page.tblDashboardIABody,strFromDate,strToDate,"CREATED ON")
End Function

'[Verify records displayed in IA LIST table based on Selected From and To Date]
Public Function VerifyRecordDisplayedBasedonDates_IAL()
strFromDate = coDashboard_IA_Page.txtFromDate.GetROProperty("value")
strToDate = coDashboard_IA_Page.txtToDate.GetROProperty("value")
VerifyRecordDisplayedBasedonDates_IAL = VerifyDateSearchRecordsdisplayed(coDashboard_IA_Page.tblIAListHeader,coDashboard_IA_Page.tblIAListBody,strFromDate,strToDate,"CREATED ON")
End Function

'[Enter CustomerName OR CIN Number in textbox]
Public Function SetCustomerNametext_IA(StrCustCIN)
	SetCustomerNametext_IA = SetEditBoxInsideTable(coDashboard_IA_Page.txtCustomerCIN(),StrCustCIN)
End Function

'[Enter Assigned To in textbox]
Public Function SetAssignedTOtext_IA(StrAssignedTo)
	SetAssignedTOtext_IA = SetEditBoxInsideTable(coDashboard_IA_Page.txtAssignedTo(),StrAssignedTo)
End Function

'[Verify Pagination for table displayed in Dashboard IA Page]
Public Function VerifyPaginationDB_IA(NoOfRows)
 bVerifyPagination = False  
 Wait 1
 gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
 Wait 2
 Set objFristPage = SetObjectFirstPage(coDashboard_IA_Page.tblPager)
 Set objPreviousPage = SetObjectPreviousPage(coDashboard_IA_Page.tblPager)
 Set objNextPage = SetObjectNextPage(coDashboard_IA_Page.tblPager)
 Set objLastPage = SetObjectLastPage(coDashboard_IA_Page.tblPager)
 
 bVerifyPagination = VerifytablePagination(coDashboard_IA_Page.tblDashboardIAHeader,coDashboard_IA_Page.tblDashboardIABody,objFristPage,objPreviousPage,objNextPage,objLastPage,"CREATED ON",NoOfRows)
 
 Set objFristPage = Nothing
 Set objPreviousPage = Nothing
 Set objNextPage = Nothing
 Set objLastPage = Nothing
 Wait 1
 gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 0")
 Wait 2
 VerifyPaginationDB_IA = bVerifyPagination 
End Function

'[Verify Pagination for table displayed in IA LIST Page]
Public Function VerifyPagination_IA(NoOfRows)
 bVerifyPagination = False  
 Wait 1
 gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
 Wait 2
 Set objFristPage = SetObjectFirstPage(coDashboard_IA_Page.tblIAListPager)
 Set objPreviousPage = SetObjectPreviousPage(coDashboard_IA_Page.tblIAListPager)
 Set objNextPage = SetObjectNextPage(coDashboard_IA_Page.tblIAListPager)
 Set objLastPage = SetObjectLastPage(coDashboard_IA_Page.tblIAListPager)
 
 bVerifyPagination = VerifytablePagination(coDashboard_IA_Page.tblIAListHeader,coDashboard_IA_Page.tblIAListBody,objFristPage,objPreviousPage,objNextPage,objLastPage,"CREATED ON",NoOfRows)
 
 Set objFristPage = Nothing
 Set objPreviousPage = Nothing
 Set objNextPage = Nothing
 Set objLastPage = Nothing
 Wait 1
 gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 0")
 Wait 2
 VerifyPagination_IA = bVerifyPagination 
End Function

'[Verify records displayed in IA Summary table based on Customer Name or CIN Number search]
Public Function VerifyResultsdisplayed_IA(StrExpValue)
 VerifyResultsdisplayed_IA = VerifySearchRecordsdisplayed(coDashboard_IA_Page.tblDashboardIAHeader,coDashboard_IA_Page.tblDashboardIABody,"Customer Name / CIF",StrExpValue)
End Function

'[Verify records displayed in IA LIST table based on AssignedTo search]
Public Function VerifyResultsdisplayed_IAL(StrExpValue)
 VerifyResultsdisplayed_IAL = VerifySearchRecordsdisplayed(coDashboard_IA_Page.tblIAListHeader,coDashboard_IA_Page.tblIAListBody,"IA List Assigned To",StrExpValue)
End Function

'[Select on row displayed in the results table]
Public Function ClickTableRow_IA(lstRowData)
ClickTableRow_IA = SelectTableRow(coDashboard_IA_Page.tblDashboardIAHeader,coDashboard_IA_Page.tblDashboardIABody,lstRowData,"Dashboard IA Summary","IA NUMBER",False,False)
End Function

'[Select on row displayed in the IA LIST results table]
Public Function ClickTableRow_IAL(lstRowData)
ClickTableRow_IAL = SelectTableRow(coDashboard_IA_Page.tblIAListHeader,coDashboard_IA_Page.tblIAListBody,lstRowData,"Overview IA LIST","IA NUMBER",False,False)
End Function

'[Verify existenence of fields displayed in View IA Page]
Public Function VerifyFieldExistenece_ViewIA()
bVerifyfields = False
	WaitForIServeLoading
	'Field in customer information section 
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblName(),"View IA","Name")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblCIF(),"View IA","CIF")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblSegment(),"View IA","Segment")
	' Field in grey panel of IA Information
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblIANumber(),"View IA","IA Number")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblSRNumber(),"View IA","SR Number")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblOverDue(),"View IA","Overdue")
	
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblCTIReferenceNo(),"View IA","CTI Ref No")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblChatRefereneNo(),"View IA","Chat Ref No")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblVAChatRefereneNo(),"View IA","VA chat Ref No")
	
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblClosedDate(),"View IA","Closed Date")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblLastUpdatedDate(),"View IA","Last Updated Date")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblLastUpdatedBy(),"View IA","Last Updated By")
	
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblParentChatReference(),"View IA","Parent Chat Ref No")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblChannel(),"View IA","Channel")
	bVerifyfields = VerifyFieldExistenceInPage(coDashboard_IA_Page.lblChatIntent(),"View IA","Chat Intent")
	' Field in grey panel of IA Information
	WaitForIServeLoading
VerifyFieldExistenece_ViewIA =  bVerifyfields
End Function

'[Verify grey panel Customer Information fields displayed in IA Page]
Public Function VerifyCustomerInformation_IA(lstCustomerInfo)
	bVerifyCustomerInfo = true
	intSize = Ubound(lstCustomerInfo)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstCustomerInfo(Iterator),":")(0))
		arrValue = trim(Split(lstCustomerInfo(Iterator),":")(1))	
		Select Case (arrLabel)		
			Case "Name"
				If Not IsNull(arrValue) Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblName_Span(),arrValue, "Customer Name") Then
						LogMessage "RSLT","Verification","Name:"&arrValue&" is not displayed as expected",False
						bVerifyCustomerInfo = False
						Exit For
					End If
				End If
			Case "CIF"
				If Not IsNull(arrValue) Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblCIF_Span(),arrValue, "CIF") Then
						LogMessage "RSLT","Verification","CIF:"&arrValue&" is not displayed as expected",False
						bVerifyCustomerInfo = False
						Exit For
					End If
				End If		
			Case "Segment"
				If Not IsNull(arrValue) Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblSegment_Span(),arrValue, "Customer Segment") Then
						LogMessage "RSLT","Verification","Segment:"&arrValue&" is not displayed as expected",False
						bVerifyCustomerInfo = False
						Exit For
					End If
				End If
		  End Select
	  Next
   VerifyCustomerInformation_IA = bVerifyCustomerInfo
End Function

'[Verify Created By and Created On text displayed in IA Page]
Public Function VerifyCreatedInfo_IA(strCreatedBy, StrCreatedDate)
	bVerifyCreatedInfo = False
	Err.Clear
	If Not IsNull(strCreatedBy) Then
	   strActCreatedBy = coDashboard_IA_Page.lblIACreatedBy.GetROProperty("innertext")
	Else 
	   strActCreatedBy = ""
	End If
	If Not IsNull(StrCreatedDate) Then	
		strActCreatedDate = coDashboard_IA_Page.lblIACreatedDate.GetROProperty("innertext")	
	Else 
	   strActCreatedDate = ""
	End If		
	strActCreatedInfo = "Created By "&UCase(Trim(strActCreatedBy))&" on "&strActCreatedDate&""
	
	strExpCreatedInfo = "Created By "&UCase(Trim(strCreatedBy))&" on "&StrCreatedDate&""
	If Trim(strActCreatedInfo) = Trim(strExpCreatedInfo) Then
		LogMessage "RSLT","Verification","Created By :"&strCreatedBy&" and Created Date :"&StrCreatedDate&" not displayed as expected",True
		bVerifyCreatedInfo = True
	End If
	VerifyCreatedInfo_IA = bVerifyCreatedInfo
End Function

'[Verify IA Related Information displayed in grey panel in IA Page]
Public Function VerifyGreyPanelIARelatedInfo(lstIARelatedInfo)
	bVerifyIARelatedInfo = true
	intSize = Ubound(lstIARelatedInfo)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstIARelatedInfo(Iterator),":",2)(0))
		arrValue = trim(Split(lstIARelatedInfo(Iterator),":",2)(1))
		
		Select Case (arrLabel)		
			Case "IA Number"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblIANumber_Span(),arrValue,"IA Number") Then
						LogMessage "RSLT","Verification","IA NUmber:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If
			Case "SR Number"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblSRNumber_Span(),arrValue,"SR Number") Then
						LogMessage "RSLT","Verification","SR Number:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If		
			Case "Overdue"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblOverDue_Span(),arrValue,"Overdue") Then
						LogMessage "RSLT","Verification","Overdue:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If
			Case "Last Updated By"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblLastUpdatedBy_Span(),arrValue,"Last Updated By") Then
						LogMessage "RSLT","Verification","Last Updated By:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If
			Case "Last Updated Date"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblLastUpdatedDate_Span(),arrValue,"Last Updated Date") Then
						LogMessage "RSLT","Verification","Last Updated Date:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If
			Case "Closed Date"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblClosedDate_Span(),arrValue,"Closed Date") Then
						LogMessage "RSLT","Verification","Closed Date:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If
			Case "Channel"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblChannel_Span(), arrValue,"Channel") Then
						LogMessage "RSLT","Verification","Channel:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If	
			Case "Source"
				If Not IsNull(arrValue) OR  arrValue <> "" Then
					If Not VerifyInnerText(coDashboard_IA_Page.lblSource_Span(), arrValue,"Source") Then
						LogMessage "RSLT","Verification","Source:"&arrValue&" is not displayed as expected",false
						bVerifyIARelatedInfo = False
						Exit For
					End If
				End If	
		End select
   Next 
VerifyGreyPanelIARelatedInfo = bVerifyIARelatedInfo
End Function

'[Verify Chat Related Information displayed in grey panel in IA Page]
Public Function VerifyGreyPanelChatRelatedInfo(lstChatRelatedInfo)
bVerifyChatRelatedInfo = true
intSize = Ubound(lstChatRelatedInfo)
For Iterator = 0 To intSize Step 1
	arrLabel = trim(Split(lstChatRelatedInfo(Iterator),":")(0))
	arrValue = trim(Split(lstChatRelatedInfo(Iterator),":")(1))
	Select Case (arrLabel)		
		Case "CTI Ref No"
			If Not IsNull(arrValue) Then
				If Not VerifyInnerText(coDashboard_IA_Page.lblCTIReferenceNo_Span(),arrValue,"CTI Ref No") Then
					LogMessage "RSLT","Verification","CTI Ref No:"&arrValue&" is not displayed as expected",False
					bVerifyChatRelatedInfo = False
					Exit For
				End If
			End If
		Case "Chat Ref No"
			If Not IsNull(arrValue) Then
				If Not VerifyInnerText(coDashboard_IA_Page.lblChatRefereneNo_Span(),arrValue,"Chat Ref No") Then
					LogMessage "RSLT","Verification","Chat Ref No:"&arrValue&" is not displayed as expected",False
					bVerifyChatRelatedInfo = False
					Exit For
				End If
			End If		
		Case "VA Chat Ref No"
			If Not IsNull(arrValue) Then
				If Not VerifyInnerText(coDashboard_IA_Page.lblVAChatRefereneNo_Span(),arrValue,"VA Chat Ref No") Then
					LogMessage "RSLT","Verification","VA Chat Ref No:"&arrValue&" is not displayed as expected",False
					bVerifyChatRelatedInfo = False
					Exit For
				End If
			End If
		Case "Chat Intent"
			If Not IsNull(arrValue) Then
				If Not VerifyInnerText(coDashboard_IA_Page.lblChatIntent_Span(),arrValue,"Chat Intent") Then
					LogMessage "RSLT","Verification","Chat Intent:"&arrValue&" is not displayed as expected",False
					bVerifyChatRelatedInfo = False
					Exit For
				End If
			End If
		Case "Parent Chat Ref"
			If Not IsNull(arrValue) Then
				If Not VerifyInnerText(coDashboard_IA_Page.lblParentChatReference_Span(),arrValue,"Parent Chat Ref") Then
					LogMessage "RSLT","Verification","Parent Chat Ref:"&arrValue&" is not displayed as expected",False
					bVerifyChatRelatedInfo = False
					Exit For
				End If
			End If
 	End select
Next 
VerifyGreyPanelChatRelatedInfo = bVerifyChatRelatedInfo
End Function

'[Verify field Related To in IA Page]
Public Function VerifyFieldRelatedTo_IA(strRelatedTo)
bVerifyFieldRelatedTo = false	
 If Not IsNull(strRelatedTo) Then
	 If VerifyInnerText(coDashboard_IA_Page.lblRelatedTo_Span(),strRelatedTo,"Related To") Then
	   bVerifyFieldRelatedTo = True
	 End If
 End If
VerifyFieldRelatedTo_IA = bVerifyFieldRelatedTo
End Function

'[Set dropdown Related To in IA Page]
Public Function SetRelatedToCombobox_IA(strRelatedTo)
	SetRelatedToCombobox_IA = SetValue(coDashboard_IA_Page.InpRelatedTo,strRelatedTo,"Related To")
End Function

'[Verify field Account Number in IA Page]
Public Function VerifyTypeAccountNumber_IA(strAccNumber)
bVerifyFieldAccNo = True	
 If Not IsNull(strAccNumber) Then
	 If Not VerifyInnerText(coDashboard_IA_Page.lblAccountNumber_Span(),strAccNumber,"Account Number") Then
	   bVerifyFieldAccNo = False
	 End If
 End If
VerifyTypeAccountNumber_IA = bVerifyFieldAccNo
End Function

'[Set dropdown Account Number in IA Page]
Public Function SetAccNumberCombobox_IA(strItem)
	SetAccNumberCombobox_IA = SetValue(coDashboard_IA_Page.lnpAccountNo,strItem,"Account Number")
End Function

'[Verify field Type in IA Page]
Public Function VerifyFieldType_IA(strType)
bVerifyFieldType = False	
If Not IsNull(strType) Then
	If VerifyInnerText(coDashboard_IA_Page.lblType_Span(),strType,"Type") Then
	bVerifyFieldType = True
	End If
End If
VerifyFieldType_IA = bVerifyFieldType
End Function

'[Set dropdown Type in IA Page]
Public Function SetTypeCombobox_IA(strItem)
	SetTypeCombobox_IA = SetValue(coDashboard_IA_Page.InpType,strItem,"Type")
End Function

'[Verify field SubType in IA Page]
Public Function VerifyFieldSubType_IA(strSubType)
bVerifyFieldSubType = False	
If Not IsNull(strSubType) Then
	If VerifyInnerText(coDashboard_IA_Page.lblSubType_Span(),strSubType,"Sub Type") Then
	bVerifyFieldSubType = True
	End If
End If
VerifyFieldSubType_IA = bVerifyFieldSubType
End Function

'[Set dropdown SubType in IA Page]
Public Function SetSubTypeCombobox_IA(strItem)
	WaitForIServeLoading
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 200")
	Wait 2
	SetSubTypeCombobox_IA = SetValue(coDashboard_IA_Page.InpSubType,strItem,"Sub Type")
End Function

'[Verify field Status in IA Page]
Public Function VerifyFieldStatus_IA(strStatus)
bVerifyFieldStatus = False	
If Not IsNull(strStatus) Then
	If VerifyInnerText(coDashboard_IA_Page.lblStatus_Span(),strStatus,"Status") Then
	bVerifyFieldStatus = True
	End If
End If
VerifyFieldStatus_IA = bVerifyFieldStatus
End Function

'[Verify Default value displayed in Status dropdown in IA Page]
Public Function VerifyDefaultStatus_IA(strStatus)
bVerifyFieldStatus = False	
If Not IsNull(strStatus) Then
	If verifyFieldValue(coDashboard_IA_Page.txtStatus(),strStatus,"Status") Then
	   bVerifyFieldStatus = True
	End If
End If
VerifyDefaultStatus_IA = bVerifyFieldStatus
End Function

'[Select dropdown Status in IA Page]
Public Function SelectStatusCombobox_IA(strItem)
	WaitForIServeLoading
	bVerify = True
	strExpitem = coDashboard_IA_Page.lstStatus.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coDashboard_IA_Page.lstStatus,strItem,"Status")
	End If
	SelectStatusCombobox_IA = bVerify
End Function

'[Set dropdown Status in IA Page]
Public Function SetStatusCombobox_IA(strItem)
	WaitForIServeLoading
	SetStatusCombobox_IA = SetValue(coDashboard_IA_Page.lstStatus,strItem,"Status")
End Function

'[Verify field SubStatus in IA Page]
Public Function VerifyFieldSubStatus_IA(strSubStatus)
	bVerifyFieldSubStatus = True
	'gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 200")	
	'Wait 2
	If Not IsNull(strSubStatus) Then
		If Not VerifyInnerText(coDashboard_IA_Page.lblSubStatus_Span(),strSubStatus,"Sub Status") Then
		bVerifyFieldSubStatus = False
		End If
	End If
	VerifyFieldSubStatus_IA = bVerifyFieldSubStatus
End Function

'[Set dropdown SubStatus in IA Page]
Public Function SetSubStatusCombobox_IA(strItem)	
	SetSubStatusCombobox_IA = SetValue(coDashboard_IA_Page.InpSubStatus,strItem,"Sub Status")
End Function

'[Verify field Assigned To in IA Page]
Public Function VerifyFieldAssignedTo_IA(strAssignedTo)
bVerifyFieldAssignedTo = False	
If Not IsNull(strAssignedTo) Then
	If VerifyInnerText(coDashboard_IA_Page.lblAssignedTo_Span(),strAssignedTo,"Assigned To") Then
	bVerifyFieldAssignedTo = True
	End If
End If
VerifyFieldAssignedTo_IA = bVerifyFieldAssignedTo
End Function

'[Verify Default value displayed in Assigned To dropdown in IA Page]
Public Function VerifyDefaultAssignedTo_IA(strAssignedTo)
bVerifyFieldAssignedTO = False	
If Not IsNull(strAssignedTo) Then
	If verifyFieldValue(coDashboard_IA_Page.InpAssignedTo(),strAssignedTo,"Assigned To") Then
	bVerifyFieldAssignedTO = True
	End If
End If
VerifyDefaultAssignedTo_IA = bVerifyFieldAssignedTO
End Function

'[Set dropdown Assigned To in IA Page]
Public Function SetAssignedToCombobox_IA(strItem)
	WaitForIServeLoading
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
	Wait 2
	SetAssignedToCombobox_IA = SetValue(coDashboard_IA_Page.InpAssignedTo,strItem,"Assigned To")
	WaitForIServeLoading
End Function

'[Verify field Duration in IA Page]
Public Function VerifyFieldDuration_IA(strDuration)
bVerifyFieldAssignedTo = False	
If Not IsNull(strDuration) Then
	If VerifyInnerText(coDashboard_IA_Page.lblDuration_Span(),strDuration,"Duration") Then
	bVerifyFieldAssignedTo = True
	End If
End If
VerifyFieldDuration_IA = bVerifyFieldAssignedTo
End Function

'[Verify Default value displayed Duration dropdown in IA Page]
Public Function VerifyDefautDuration_IA(strDuration)
bVerifyFieldDuration = False	
If Not IsNull(strDuration) Then
	If verifyFieldValue(coDashboard_IA_Page.txtDuration(),strDuration,"Duration") Then
	bVerifyFieldDuration = True
	End If
End If
VerifyDefautDuration_IA = bVerifyFieldDuration
End Function

'[Set dropdown Duration in IA Page]
Public Function SelectDurationCombobox_IA(strItem)
	WaitForIServeLoading
	bVerify = True
	strExpitem = coDashboard_IA_Page.lstDuration.GetRoproperty("value")
	If Not (Ucase(Trim(strExpitem)) = Ucase(Trim(strItem))) Then
		bVerify = SelectComboBoxItem(coDashboard_IA_Page.lstDuration,strItem,"Duration")
	End If
	WaitForIServeLoading
	SelectDurationCombobox_IA = bVerify
End Function

'[Verify field Source in IA Page]
Public Function VerifyFieldSource_IA(strSource)
bVerifyFieldSource = True	
If Not IsNull(strSource) Then
	If Not VerifyInnerText(coDashboard_IA_Page.lblSource_Span(),strSource,"Source") Then
	bVerifyFieldSource = False
	End If
End If

VerifyFieldSource_IA = bVerifyFieldSource
End Function

'[Set dropdown Source in IA Page]
Public Function SetSourceCombobox_IA(strItem)	
	SetSourceCombobox_IA = SetValue(coDashboard_IA_Page.lnpSource,strItem,"Source")
End Function

'[Verify field Comments in IA Page]
Public Function VerifyFieldComments_IA(strComments)
bVerifyFieldComments = True	
If Not IsNull(strComments) Then
	If Not VerifyInnerText(coDashboard_IA_Page.lblComments_Span(),strComments,"Comments") Then
	bVerifyFieldComments = False
	End If
End If
VerifyFieldComments_IA = bVerifyFieldComments
End Function

'[Enter Comments textbox in IA Page]
Public Function SetComments_IA(StrComment)
bVerifytext = True
WaitForIServeLoading
gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 600")
Wait 2
If Not (IsNull(StrComment) Or StrComment="BLANK" or Trim(StrComment)="") Then
	coDashboard_IA_Page.txtComments().set StrComment
End If
	If Err.Number <> 0 Then
	  bVerifytext = False
	  LogMessage "WARN","Verification","Failed to Set Comments in text box", False
	  Exit Function
	End If
WaitForIServeLoading
SetComments_IA = bVerifytext
End Function

'[Verify checkbox OnceDone checked or Unchecked in IA Page]
Public Function VerifyCheckbox_IA(strCheckFlag)		
VerifyCheckbox_IA = VerifyObjectCheckedUnchecked(coDashboard_IA_Page.ChkboxOnceDone,strCheckFlag,"OnceDone Checkbox")
End Function

'[Verify checkbox OnceDone Enabled or disabled in IA Page]
Public Function VerifyCheckboxEnabled_IA(strCheckFlag)
	VerifyCheckboxEnabled_IA = VerifyObjectEnabledDisabled(coDashboard_IA_Page.ChkboxOnceDone,strCheckFlag,"OnceDone Checkbox")
End Function
	
'[Select checkbox OnceDone and verify in IA Page]
Public Function SelectCheckBox_IA(strcheck)	
bSelectCheckBox = False
gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 100")
  If strcheck = "Check" Then
	bSelectCheckBox= SelectCheckBoxAndVerify_ID(coDashboard_IA_Page.ChkboxOnceDone,"OnceDone Checkbox")
  End If	
  SelectCheckBox_IA = bSelectCheckBox
  gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 0")
End Function

'[Verify text displayed below Attachment section in IA Page]
Public Function VerifyAttachmentText_IA(StrExpText1,StrExpText2,StrExpText3)
bverifyAttachmenttext = False 
	StrActText1 = coDashboard_IA_Page.txtAttachment1.GetROProperty("innertext")
	StrActText2 = coDashboard_IA_Page.txtAttachment2.GetROProperty("innertext")
	StrActText3 = coDashboard_IA_Page.txtAttachment3.GetROProperty("innertext")

	If (StrExpText1 = StrActText1) AND (StrExpText2 = StrActText2) AND (StrExpText3 = StrActText3) Then
	   LogMessage "RSLT","Verification","Text displayed below the Attachment section is displayed as expected",True
	   bverifyAttachmenttext = True
	Else 
	   LogMessage "WARN","Verification","Text displayed below the Attachment section is not displayed as expected", False
	End If
VerifyAttachmentText_IA  = bverifyAttachmenttext
End Function

'[Add one File by Clicking Attachement Button in IA Page]
Public Function AddAttachments_IA(strFileName)
  bverifyAddedAttachment = False
  gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 600")
  Wait 2
  Setting.WebPackage("ReplayType") = 2
  coDashboard_IA_Page.btnAddAttachments.click   
  If Err.Number <> 0 Then
      LogMessage "WARN","Verification","Failed to Click Button : NEW IA", False
      Exit Function
  End If  
  Wait 3
  WaitForIServeLoading
  'Get the folder path from the OBTAF_Config
  strFolderPath = gstrAttachmentsPath
  filePath = strFolderPath + "\" + strFileName  
  coDashboard_IA_Page.txtFileName.Set filePath
  coDashboard_IA_Page.btnOpen.Click
  Setting.WebPackage("ReplayType") = 1   
  
  StrAddedFileName = coDashboard_IA_Page.lblFileName.GetROProperty("innertext")  
  If Len(strFileName)<=75 Then
  	 If Trim(strFileName) = Trim(StrAddedFileName) Then
  	 LogMessage "RSLT","Verification","File Added is displayed in IA Page as expected",True
  	 bverifyAddedAttachment = True
     Else 
  	 LogMessage "WARN","Verification","File Added is not displayed in IA Page", False
  	 bverifyAddedAttachment = False
     End If
  Else
  	 bverifyAddedAttachment = True
  End If
    
  AddAttachments_IA = bverifyAddedAttachment
  gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 0")
  Wait 2
End Function

'[Verify Inline error message displayed related to Attachments added]
Public Function VerifyAttachmentInlineMsg_IA(strErrorMsg) 
bVerifyInlineMsg = False
If VerifyInnerText(coDashboard_IA_Page.lblInlineErrortxt(), strErrorMsg, "Invalid Attachment Error") Then
	bVerifyInlineMsg = True
End IF 
VerifyAttachmentInlineMsg_IA = bVerifyInlineMsg
End Function

'[Verify CreatedBy and CreatedOn displayed for Added Attachments in IA Page]
Public Function VerifyAttachmentCreatedInfo_IA(strCreatedBy, strCreatedOn)
bVerifyCreatedInfo = False

strActCreatedBy = coDashboard_IA_Page.lblCreatedBy.GetROProperty("innertext")

If Ucase(Trim(strActCreatedBy))  = Ucase(Trim(strCreatedBy)) Then
   LogMessage "WARN","Verification","CreatedBy "&strCreatedBy&" displayed as expected in Attachments section",True
   bVerifyCreatedInfo = True
   
   If Not (IsNull(Trim(strCreatedOn)) OR Trim(strCreatedOn)="" OR Trim(strCreatedOn)="BLANK") Then
   	  strActCreatedDate = coDashboard_IA_Page.lblCreatedOn.GetROProperty("innertext")
   	  If strActCreatedDate = strCreatedOn  Then
   	  	 LogMessage "WARN","Verification","CreatedOn "&strCreatedOn&" displayed as expected in Attachments section",True
         bVerifyCreatedInfo = True
      Else 
   	  	 LogMessage "WARN","Verification","CreatedOn "&strCreatedOn&" displayed as expected in Attachments section",False
         bVerifyCreatedInfo = False			         
   	  End If
   End If
   
Else
   LogMessage "WARN","Verification","CreatedBy "&strCreatedBy&" not displayed in Attachments section as expected",False
End If

VerifyAttachmentCreatedInfo_IA = bVerifyCreatedInfo
End Function

'[Verify display of description field in IA Page]
Public Function VerifyFieldDescription_IA(strCheckFlag)
VerifyFieldDescription_IA = VerifyObjectDisabled(coDashboard_IA_Page.txtAttachmentComment,strCheckFlag,"Description or Comments Field")
End Function

'[Click button Remove Attachment in IA Page]
Public Function clickButtonRemoveAttachment_IA()
gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 600")
Wait 1
coDashboard_IA_Page.btnRemoveAttachment.click 
If Err.Number <> 0 Then
  clickButtonRemoveAttachment_IA = False
  LogMessage "WARN","Verification","Failed to Click Button: Remove Attachment", False
  Exit Function
Else
 clickButtonRemoveAttachment_IA = True
End If

gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 0")
Wait 1
End Function

'[Verify display of Submit Button in IA Page]
Public Function VerifyButtondisplaySubmit_IA(strCheckFlag)
VerifyButtondisplaySubmit_IA = VerifyObjectDisabled(coDashboard_IA_Page.btnSubmitIA,strCheckFlag,"Submit Button")
End Function

'[Click on Submit Button in View Page]
Public Function clickButtonSubmitViewPage_IA()
coDashboard_IA_Page.btnSubmitIA.click 

If Err.Number <> 0 Then
  clickButtonSubmitViewPage_IA = False
  LogMessage "WARN","Verification","Failed to Click Button: Submit", False
  Exit Function
Else
  clickButtonSubmitViewPage_IA = True
End If
WaitForIServeLoading
End Function

'[Click OK for Submission Message popup]
Public Function ClickButtonOK_IA()
coDashboard_IA_Page.btnOK.click 
If Err.Number <> 0 Then
  ClickButtonOK_IA = False
  LogMessage "WARN","Verification","Failed to Click Button: OK", False
  Exit Function
Else
 ClickButtonOK_IA = True
End If
WaitForIServeLoading

End Function

'[Verify display of Add Attachment button in IA Page]
Public Function VerifyButtonAddAttachment_IA(strCheckFlag)
VerifyButtonAddAttachment_IA = VerifyObjectEnabledDisabled(coDashboard_IA_Page.btnAddAttachments,strCheckFlag,"Add Attachment Button")	
End Function

'[Verify Infowarn Message displayed in Dashboard IA Page]
Public Function VerifyInfowan_DashboardIA(strInfoMsgtext)
	VerifyInfowan_DashboardIA = VerifyInfowarntext(coDashboard_IA_Page.lblInfowarn,strInfoMsgtext)
End Function

'[Verify Record count displayed based on Selected Status in IA Summary table]
Public Function VerifyIARecordCount_DashboardIA(strStatus)
bVerifyRecordCount = False

strDisplayedMsgtext = coDashboard_IA_Page.lblRecordCount.GetRoProperty("innertext")
strMsgText = "Interaction Activities "&strStatus

	If Instr(1,strDisplayedMsgtext,strMsgText,1) > 0 Then 
	   LogMessage "WARN","Verification","Record Count text message is displayed as expected", True
	   bVerifyRecordCount = True
	Else 
	   LogMessage "WARN","Verification","Record Count text message is not displayed as expected", False
	   bVerifyRecordCount = False	
	End IF 
	
VerifyIARecordCount_DashboardIA = bVerifyRecordCount
End Function

'[Veriy display of Edit IA link in View IA Page]
Public Function VerifyEditLinkDisplay_IA(strFlag)
VerifyEditLinkDisplay_IA = VerifyObjectEnabledDisabled(coDashboard_IA_Page.lnkEdit,strFlag,"Edit Link")
End Function

'[Verify No records Message displayed in IA]
Public Function VerifyMsg_IA(strErrorMsg)
bverifyMsg = False
If VerifyInnerText(coDashboard_IA_Page.lblWarn(), strErrorMsg, "No Records - IA") Then
   bverifyMsg = True
End If
VerifyMsg_IA = bverifyMsg
End Function

'[Verify text displayed for total no of records in IA LIST]
Public Function VerifyIARecordCount_IAList(strRecCnt)
bVerifyRecordCount = False

strDisplayedMsgtext = coDashboard_IA_Page.lblIARecCount.GetRoProperty("innertext")
strMsgText = strRecCnt&" Results Found"

	If Instr(1,strDisplayedMsgtext,strMsgText,1) > 0 Then 
	   LogMessage "WARN","Verification","Record Count text message is displayed as expected", True
	   bVerifyRecordCount = True
	Else 
	   LogMessage "WARN","Verification","Record Count text message is not displayed as expected", False
	   bVerifyRecordCount = False	
	End IF 
	
VerifyIARecordCount_IAList = bVerifyRecordCount
End Function

'[Click on Cancel Button in NEW IA Page]
Public Function clickCancelButton_IA()
gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 600")
	coDashboard_IA_Page.btnCancel.click
	If Err.Number <> 0 Then
		clickCancelButton_IA = False
		LogMessage "WARN","Verification","Failed to Click Button: Cancel", False
		Exit Function
	Else
		clickCancelButton_IA = True
	End If
	WaitForIServeLoading
	
End Function

'[Verify the Cancel Confirmation message in IA Page]
Public Function VerifyCancelMessage_IA(strCancelMsg)
	bVerifyCancelMessage = False	
	If Not IsNull(strCancelMsg) Then
	   If verifyInnerText(coDashboard_IA_Page.lblCancelMsgIA,strCancelMsg,"Cancel Message") Then
			bVerifyCancelMessage = True
		End If
	End If
	VerifyCancelMessage_IA = bVerifyCancelMessage
End Function

'[Expand View Activity Details displayed in View SR Page]
Public Function clickExpanIcon_VSR()
	bVerify = false
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 1000")
	Wait 2
	Set objAccordionGrp = coServiceRequest_Page.eleViewSRAccordionObject
	bVerify = ExpandSingleAccordion(objAccordionGrp,"View Activity Details")
	clickExpanIcon_VSR=bVerify
End Function

'[Click on Add Activity Link displayed in View Activity Details]
Public Function clickActivityButton_SR()
	coDashboard_IA_Page.btnAddActivity.click
	If Err.Number <> 0 Then
		clickActivityButton_SR = False
		LogMessage "WARN","Verification","Failed to Click Button: Add Activity", False
		Exit Function
	Else
		clickActivityButton_SR = True
	End If
	WaitForIServeLoading
End Function

'[Click on Refresh Button in Add Activity Section]
Public Function clickActivityRfButton_SR()
	coDashboard_IA_Page.btnRefreshAddActivity.click
	If Err.Number <> 0 Then
		clickActivityRfButton_SR = False
		LogMessage "WARN","Verification","Failed to Click Button: Add Activity Refresh", False
		Exit Function
	Else
		clickActivityRfButton_SR = True
	End If
	WaitForIServeLoading
End Function

'[Select on row displayed in View Activity Details table in SR]
Public Function ClickTableRowAA_SR(lstRowData)
ClickTableRowAA_SR = SelectTableRow(coServiceRequest_Page.tblViewActivityDetailsSRHeader,coServiceRequest_Page.tblViewActivityDetailsSRBody,lstRowData,"View Activity Details in SR","CREATED DATE",False,False)
End Function

'[Click on Edit IA Link in View IA]
Public Function clickEditIAButton_IA()
	coDashboard_IA_Page.lnkEdit.click
	If Err.Number <> 0 Then
		clickEditIAButton_IA = False
		LogMessage "WARN","Verification","Failed to Click Button: Edit IA", False
		Exit Function
	Else
		clickEditIAButton_IA = True
	End If
	WaitForIServeLoading
End Function

'[Expand Attachment Details Accordion in Edit IA Tab]
Public Function clickAttachments_AC()
	bVerify = false
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 1000")
	Wait 2
	Set objAccordionGrp = coDashboard_IA_Page.EditAccordian
	bVerify = ExpandSingleAccordion(objAccordionGrp,"Attachment Details")
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 800")
	Wait 2
	clickAttachments_AC=bVerify
End Function

'[Click on IA or SR Triplet displayed in grid panel overview page]
Public Function ClickTriplet_Overview(strTriplets)
 ClickTriplet_Overview = False
 Set TotalRowsPanel = SetObjPanelRow(coOverview_Page.elePanelHeader)	
	 ActualNoOfRows = TotalRowsPanel.Count
If Not IsNull(strTriplets)  Then
	 If IsArray(strTriplets) Then
		strTriplet = Join(strTriplets,"|")
	 Else
	 	strTriplet = Trim(strTriplets)
	 End If
	 If ActualNoOfRows <> 0 Then	 
		 Set ObjTriplets = SetObjTriplets(coOverview_Page.elePanelHeader)	 
		 For i  = 0 To ActualNoOfRows-1 Step 1			 
			 ActualTriplet = ObjTriplets(i).GetRoProperty("innertext")		 
			 If Ucase(Trim(ActualTriplet)) =  Ucase(Trim(strTriplet)) Then
			 	ObjTriplets(i).Click
			 	If Err.Number <> 0 Then
				   ClickTriplet_Overview = False
				   LogMessage "WARN","Verification","Failed to Click IA or SR Triplet in Overview Page." ,False
				Else 
				   LogMessage "RSLT","Verification","Clicked on IA or SR Number: "&strTriplet&" without any errors", True
		    	   ClickTriplet_Overview = True
				End If
				Exit Function
			 End If
	     Next 
	  	 LogMessage "WARN","Verification","Matching Request Number not found in a panel" ,False
	 End If  
 End If
 Set TotalRowsPanel = Nothing
 Set ObjTriplets = Nothing
End Function

'[Click on Update Link in View IA]
Public Function clickUpdateIAButton_IA()
	coDashboard_IA_Page.lnkUpdate.click
	If Err.Number <> 0 Then
		clickUpdateIAButton_IA = False
		LogMessage "WARN","Verification","Failed to Click Button: Update IA", False
		Exit Function
	Else
	clickUpdateIAButton_IA = True
	End If
	WaitForIServeLoading
End Function

'[Click Yes Or No Button in cancel Message displayed in IA]
Public Function SelectButtonCancel_IA(strSelect)
	bClickButton = False
	If Not IsNull(strSelect) Then
	   If strSelect = "YES" Then
	   	 coDashboard_IA_Page.btnYesIA.click
	   ElseIf strSelect = "NO" Then
	     coDashboard_IA_Page.btnNoIA.click
	   End If 	   
	   If Err.Number<>0 Then
		   LogMessage "WARN","Verification","Failed to Click Button YES/NO", False
	   Else
		   bClickButton = True
	   End If
	End If
	WaitForIServeLoading
	SelectButtonCancel_IA = bClickButton
End Function
