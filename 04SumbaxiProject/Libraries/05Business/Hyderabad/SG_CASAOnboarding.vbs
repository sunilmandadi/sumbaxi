'[Verify Combobox Application show displayed as]
Public Function verifyApplication_default(strApplication)
   bDevPending=false
   bverifyApplication_default=true
   If Not IsNull(strFeeReversalType) Then
       If Not verifyComboSelectItem (CASAOnboarding.ApplicationShowDropdown(),strApplication, "ApplicationShow")Then
           bverifyApplication_default=false
       End If
   End If
   verifyApplication_default=bverifyApplication_default
End Function

'[Verify list of values displayed in Application show dropdown]
Public Function verifyApplicationShow(lstShow)
   bDevPending=false
   bverifyApplicationShow=true
   If Not IsNull(lstShow) Then
       If Not verifyComboboxItems(CASAOnboarding.ApplicationShowDropdown(),lstShow, "ApplicationShow")Then
           bverifyApplicationShow=false
       End If
   End If
   verifyApplicationShow=bverifyApplicationShow
End Function

''[Select value from Application show combobox as]
'Public Function selectApplicationShowComboBox(strShow)
'   bDevPending=false
'   bselectApplicationShowComboBox=true
'   If Not IsNull(strShow) Then
'       If Not (selectItem_Combobox(CASAOnboarding.ApplicationShowDropdown(), strShow))Then
'            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Show drop down list" ,false
'           bselectApplicationShowComboBox=false
'       End If
'   End If
'   WaitForICallLoading
'   selectApplicationShowComboBox=bselectApplicationShowComboBox
'End Function

'[Verify list of values displayed in AccountType Combobox as]
Public Function selectAccounttypeComboBox(strGroup)
   bDevPending=true
   bselectAccounttypeComboBox=true
   If Not IsNull(strGroup) Then
       If Not (selectItem_Combobox(CASAOnboarding.AccountTypeDropdown(), strGroup))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Group drop down list" ,false
           bselectAccounttypeComboBox=false
       End If
   End If
   selectAccounttypeComboBox=bselectAccounttypeComboBox
End Function

'[Select Application status Radio Button Type on CASA Home Page]
Public Function selectApplicationStatusRadio(strType)
	bDevPending=False
	bselectApplicationStatusRadio=true
	bselectApplicationStatusRadio=SelectRadioButtonGrp(strType, CASAOnboarding.ApplicationStatus, Array("Open","Pend-Auth","Re-work","Closed","Incomplete","All status"))
   WaitForICallLoading
	If Err.Number<>0 Then
       bselectApplicationStatusRadio=false
       LogMessage "WARN","Verification","Failed to Click Button : Type on Home Page" ,false
       Exit Function
   End If
   selectApplicationStatusRadio=bselectApplicationStatusRadio
End Function

'[Select value from show AccountType dropdown as]
Public Function selectAccountTypeComboBox(strAccType)
   bDevPending=false
   bselectAccountTypeComboBox=true
   If Not IsNull(strAccType) Then
       If Not (selectItem_Combobox(CASAOnboarding.AccountTypeDropdown(), strAccType))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Show drop down list" ,false
           bselectAccountTypeComboBox=false
       End If
   End If
   WaitForICallLoading
   selectAccountTypeComboBox=bselectAccountTypeComboBox
End Function

'[Enter Application ID in text box as]
 Public Function SetCASAApplicationID(strApplicationID)
 	bSetCASAApplicationID = True 
 	If strApplicationID = "RUNTIME" Then
 	   strApplicationID = Environment.Value("ApplicationID")
 	   SetApplicationID = CASAOnboarding.ApplicationID.Set(strApplicationID)
 	   LogMessage "RSLT","Verification","Application ID entered in the text box" ,True
 	   bSetCASAApplicationID = True 
 	Else If Not IsNull(strApplicationID) Then
 	   SetApplicationID = CASAOnboarding.ApplicationID.Set(strApplicationID)
 	   LogMessage "RSLT","Verification","Application ID entered in the text box" ,True
 	   bSetCASAApplicationID = True 
 	End If
 	End If
 	SetCASAApplicationID = bSetCASAApplicationID
 End Function

'[Verify Total record count displayed for partial Search]
Public Function VerifyTotalRecordCount(lstrecordvalue)
   bVerifyTotalRecordCount = True
   If IsNull(lstrecordvalue) Then
       If VerifyInnerText (CASAOnboarding.RecordCount, "0 Records", "Validation Message")Then
           LogMessage "RSLT","Verification","No results are displayed as expected when searched with partial value" ,True
           bVerifyTotalRecordCount = True
       End If
   Else 
   		 bVerifyTotalRecordCount = False
   End If 
   VerifySearchRecordtable = bVerifySearchRecordtable
 End Function

'[Enter NRIC/PassportNumber in text box as]
 Public Function SetPassportNumber(strNRICNo)
 bSetPassportNumber = True
 	If Not IsNull(strNRICNo) Then
 	   SetPassportNumber = CASAOnboarding.NRICPassport.Set(strNRICNo)
 	   LogMessage "RSLT","Verification","Successfully entered NRIC/PassportNumber in the text field" ,True
 	   bSetPassportNumber = True
	End If 
	SetPassportNumber = bSetPassportNumber
 End Function
 
'[Enter FullName in text box as]
 Public Function EnterName(strName)
 bEnterName = True
 	If Not IsNull(strName) Then
 	   SetName = CASAOnboarding.Name.Set(strName)
 	   LogMessage "RSLT","Verification","Successfully entered Name in the name text field" ,True
 	   bEnterName = True
	End If 
	EnterName = bEnterName
 End Function

'[Click Button Filter on CASA Home Page]
Public Function clickCASAButtonFilter()
   bDevPending=true
   Wait 5
   CASAOnboarding.ButtonFilter.click 10,10,0 
   If Err.Number<>0 Then
       clickCASAButtonFilter=false
            LogMessage "WARN","Verification","Failed to Click Button : Filter" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickCASAButtonFilter=true
End Function

'[Verify Record table is displayed based on Search Criteria]
Public Function VerifySearchRecordtable(lstrecordvalue)
   bVerifySearchRecordtable = True
   strQueryTotalCount = Environment.Value("strQueryTotalCount") 
   If IsNull(lstrecordvalue) Then
       If VerifyInnerText (CASAOnboarding.RecordCount, "0 Records", "Validation Message")Then
           LogMessage "RSLT","Verification","No results are displayed as expected when searched with partial value" ,True
           bVerifySearchRecordtable = True
       End If
   Else If strQueryTotalCount = 0 Then
       If VerifyInnerText (CASAOnboarding.RecordCount, "0 Records", "Validation Message")Then
           LogMessage "RSLT","Verification","Total Count of Results in db matched with the Iserve Record Count" ,True
           bVerifySearchRecordtable = True
       End If
   Else 
   	VerifySearchRecordtable=verifyTableContentList(CASAOnboarding.SearchtableHeader,CASAOnboarding.SearchtableContent,lstrecordvalue,"Search Record Table",false,null,null,null)
   End If
  End If 
   VerifySearchRecordtable = bVerifySearchRecordtable
 End Function

'[Click on NRIC number from the Record table]
Public Function ClickPassportNumber(strApplicationID)
 bClickPassportNumber = True 
'	If strApplicationID = "RUNTIME" Then
' 	   strApplicationID = Environment.Value("ApplicationID")
 	   UserclickNRICNoLink=selectTableLink(CASAOnboarding.SearchtableHeader,CASAOnboarding.SearchtableContent,Array("Application ID:"&strApplicationID),"CASA Record table","NRIC / Passport No.",true,null,null,null)	
 	   LogMessage "RSLT","Verification","NRIC Number is clicked successfully from the record table" ,True
 	   bClickPassportNumber = True 
' 	Else 
' 	   bClickPassportNumber = False
' 	End If
'	WaitForIcallLoading	
	ClickPassportNumber = bClickPassportNumber
End Function

'[Set TextBox From on CASA Home Page to]
Public Function setFromTextbox_CASA(strFrom)
   bDevPending=true
   If not isNull(strFrom) Then
	   If Ucase(strFrom)="TODAY" Then
			If len(Day(CDate(Now)))=1 Then
				strDay="0"&Day(CDate(Now))
			else
				strDay=""&Day(CDate(Now))
			End If
			strFrom=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&""
	   End If
	   CASAOnboarding.FromDate.Set(strFrom)
   End If
   If Err.Number<>0 Then
       setFromTextbox_CASA=false
            LogMessage "WARN","Verification","Failed to Set Text Box :From" ,false
       Exit Function
   End If
   setFromTextbox_CASA=true
End Function

'[Set TextBox TO on CASA Home Page to]
Public Function setToTextbox_CASA(strTo)
   bDevPending=true
   If not isNull(strTo) Then
	   If Ucase(strTo)="TODAY" Then
			If len(Day(CDate(Now)))=1 Then
				strDay="0"&Day(CDate(Now))
			else
				strDay=""&Day(CDate(Now))
			End If
			strTo=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&""
	   End If
   CASAOnboarding.ToDate.Set(strTo)
   End If
   If Err.Number<>0 Then
       setToTextbox_CASA=false
            LogMessage "WARN","Verification","Failed to Set Text Box :To" ,false
       Exit Function
   End If
   setToTextbox_CASA=true
End Function

'[Verify Inline error Message on CASA Homepage]
Public Function verifyInlineErrorMessage(strInLineMessage)
	bverifyInlineErrorMessage=true
	If Not IsNull(strInLineMessage) Then
       If Not VerifyInnerText (CASAOnboarding.InlineMessagetext(), strInLineMessage, "Inline Info Message")Then
           bverifyInlineErrorMessage=false
       End If
   End If
   verifyInlineErrorMessage=bverifyInlineErrorMessage
End Function

'[Verify Total Record Count based on Application status from DB]
Public Function FetchTotalrecordCount(Strstatus)
	bFetchTotalrecordCount = True 
	strQuery = "select Count(distinct a1.application_id) from iserve_casa_application a1, iserve_casa_appln_details a2 where a1.casa_app_status='"&Strstatus&"' and a1.application_id  = a2.app_id"
	strQueryTotalCount=getDBValForColumn_NZD(strQuery)(0)
	IServeTotalCount = CASAOnboarding.RecordCount.GetROProperty("innertext")
	IServeTotalCount = Instr(IServeTotalCount,"Records","")
	If Ucase(Trim(IServeTotalCount)) = Ucase(Trim(strQueryTotalCount))  Then
		LogMessage "RSLT","Verification","Total Record Count matched with the expected value" ,True
		bFetchTotalrecordCount = True 
	Else 
		bFetchTotalrecordCount = False
	End If
	FetchTotalrecordCount = bFetchTotalrecordCount
End Function

'[Verify Total Record Count based on AccountType from DB]
Public Function FetchTotalrecordCount(Strstatus, strAccountType)
	bFetchTotalrecordCount = True 
	If Strstatus = "All status" Then
		strQuery = "select Count(distinct a1.application_id) from iserve_casa_application a1, iserve_casa_appln_details a2 where a1.prod_desc = '"&strAccountType&"' and a1.application_id  = a2.app_id"
	Else 
		strQuery = "select Count(distinct a1.application_id) from iserve_casa_application a1, iserve_casa_appln_details a2 where a1.casa_app_status='"&Strstatus&"' and a1.prod_desc = '"&strAccountType&"' and a1.application_id  = a2.app_id"
	End If
	strQueryTotalCount=getDBValForColumn_NZD(strQuery)(0)
	
	Environment.Value("strQueryTotalCount") = strQueryTotalCount
	IServeTotalCount = CASAOnboarding.RecordCount.GetROProperty("innertext")
	IServeTotalCount = Replace(IServeTotalCount,"Records","")
	If Ucase(Trim(IServeTotalCount)) = Ucase(Trim(strQueryTotalCount))  Then
		LogMessage "RSLT","Verification","Total Record Count matched with the expected value" ,True
		bFetchTotalrecordCount = True 
	Else 
		bFetchTotalrecordCount = False
	End If
	FetchTotalrecordCount = bFetchTotalrecordCount
End Function

'[Fetch Application ID from DB]
Public Function FetchAppIDStatus(Strstatus,StrMakerID)
	bFetchAppIDStatus = True
	strQuery = "select Distinct(a1.application_id) from iserve_casa_application a1, iserve_casa_appln_details a2,Iserve_CASA_Audit a3 where a1.casa_app_status='"&Strstatus&"' and a1.application_id  = a2.app_id and a3.Checker_ID is Null and a3.Maker_ID != '"&StrMakerID&"'"
	strQueryAppIDStatus=getDBValForColumn_NZD(strQuery)(0)
	If Not IsNull(strQueryAppIDStatus)Then
		Environment.Value("ApplicationID") = strQueryAppIDStatus
		LogMessage "RSLT","Verification","Application ID found in DB table" ,True
		bFetchAppIDStatus = True
	Else 
		LogMessage "WARN","Verification","No Application ID found in DB table" ,False
		bFetchAppIDStatus = False
	End If
	FetchAppIDStatus = strQueryAppIDStatus
End Function

'[Fetch Application ID from DB with substatus]
Public Function FetchAppIDSubStatus(Strstatus, strsubstatus, StrMakerID)
	bFetchAppIDSubStatus = True
	'strQuery = "select app_id from iserve_casa_audit where STATUS_TO ='"&Strstatus&"' and SUB_STATUS_TO ='"&strsubstatus&"' and Checker_ID is not Null "
	'strQuery = "select a1.application_id from iserve_casa_application a1, Iserve_CASA_Audit a2 where a2.sub_status_to = '"&strsubstatus&"' and a1.casa_app_status='"&Strstatus&"' and A2.Maker_ID != '"&StrMakerID&"'"
	strQuery = "select Distinct(a1.application_id),a3.Maker_ID , a3.Checker_ID from iserve_casa_application a1, iserve_casa_appln_details a2,Iserve_CASA_Audit a3 where a1.casa_app_status='"&Strstatus&"' and a3.sub_status_to = '"&strsubstatus&"' and a1.application_id  = a2.app_id and a3.Checker_ID is Null and a3.Maker_ID != '"&StrMakerID&"'"
	strQueryAppIDSubStatus=getDBValForColumn_NZD(strQuery)(0)
	If Not IsNull(strQueryAppIDSubStatus)Then
		Environment.Value("ApplicationID") = strQueryAppIDSubStatus
		LogMessage "RSLT","Verification","Application ID found in DB table" ,True
		bFetchAppIDSubStatus = True
	Else 
		LogMessage "WARN","Verification","No Application ID found in DB table" ,False
		bFetchAppIDSubStatus = False
	End If
	FetchApplicationID = strQueryAppIDSubStatus
End Function

'[Verify list of values in Staus dropdown]
Public Function VerifyStatusdropdown(lstUpdateStatus)
   bDevPending=true
   bVerifyStatusdropdown=true
   If Not IsNull(lstUpdateStatus) Then
       If Not verifyComboboxItems(CASAOnboarding.UpdateStatus(), lstUpdateStatus,"Status") Then
            LogMessage "WARN","Verification","List of values displayed in Status dropdown doesnt match with expected values" ,false
           bVerifyStatusdropdown=false
       End If
   End If
   selectAccounttypeComboBox=bVerifyStatusdropdown
End Function

'[Verify list of values in SubStaus dropdown]
Public Function VerifySubStatusdropdown(strUpdateSubStatus)
   bDevPending=true
   bVerifySubStatusdropdown=true
   If Not IsNull(strUpdateSubStatus) Then
       If Not verifyComboboxItems (CASAOnboarding.UpdateSubstatus(), strUpdateSubStatus, "SubStatus") Then
            LogMessage "WARN","Verification","List of values displayed in Substatus dropdown doesnt match with expected values" ,false
           bVerifySubStatusdropdown=false
       End If
   End If
   VerifySubStatusdropdown=bVerifySubStatusdropdown
End Function

'[Verify default status value displayed as]
Public Function verifyStatus_Default(strStatus)
   bDevPending=false
   bverifyStatus_Default=true
   If Not IsNull(strStatus) Then
       If Not verifyComboSelectItem(CASAOnboarding.UpdateStatus(),strStatus, "Status")Then
       	  LogMessage "WARN","Verification","Default Status value doesnt match with expected value" ,false
           bverifyStatus_Default=false
       End If
   End If
   verifyStatus_Default=bverifyStatus_Default
End Function

'[Verify default Sub-status value displayed as]
Public Function verifySubStatus_Default(strSubStatus)
   bDevPending=false
   bverifySubStatus_Default=true
   If Not IsNull(strSubStatus) Then
       If Not verifyComboSelectItem(CASAOnboarding.UpdateSubstatus(),strSubStatus, "SubStatus")Then
       	   LogMessage "WARN","Verification","Default Substatus value doesnt match with expected value" ,false
           bverifyStatus_Default=false
       End If
   End If
   verifySubStatus_Default=bverifySubStatus_Default
End Function

'[Verify list of values in Rejected Reason dropdown]
Public Function VerifyRejectReasondropdown(strRejectReason)
   bDevPending=true
   bVerifyRejectReasondropdown=true
   If Not IsNull(strRejectReason) Then
       If Not verifyComboboxItems (CASAOnboarding.UpdateRejectReason(), strRejectReason, "RejectedReason") Then
            LogMessage "WARN","Verification","List of values displayed in Rejected Reason dropdown doesnt match with expected values" ,false
           bVerifyRejectReasondropdown=false
       End If
   End If
   VerifyRejectReasondropdown=bVerifyRejectReasondropdown
End Function

'[Select status value from the dropdown as]
Public Function selectUpdateStatus(strstatus)
	bselectUpdateStatus=true
	If Not IsNull(strstatus) Then
       If Not (selectItem_Combobox (CASAOnboarding.UpdateStatus(), strstatus))Then
            LogMessage "WARN","Verification","Failed to select :"&strstatus&" From status drop down list" ,false
           bselectUpdateStatus=false
       End If
   End If
   WaitForICallLoading
   selectUpdateStatus=bselectUpdateStatus
End Function

'[Select Substatus value from the dropdown as]
Public Function selectSubStatus(strSubstatus)
	bselectSubStatus=true
	If Not IsNull(strSubstatus) Then
       If Not (selectItem_Combobox (CASAOnboarding.UpdateSubstatus(), strSubstatus))Then
            LogMessage "WARN","Verification","Failed to select :"&strSubstatus&" From Substatus drop down list" ,false
           bselectSubStatus=false
       End If
   End If
   WaitForICallLoading
   selectSubStatus=bselectSubStatus
End Function

'[Select Rejected Reason value from the dropdown as]
Public Function selectRejectedReason(strRejectReason)
	bselectRejectedReason=true
	If Not IsNull(strRejectReason) Then
       If Not (selectItem_Combobox (CASAOnboarding.UpdateRejectReason(), strRejectReason))Then
            LogMessage "WARN","Verification","Failed to select :"&strRejectReason&" From Substatus drop down list" ,false
           bselectRejectedReason=false
       End If
   End If
   WaitForICallLoading
   selectRejectedReason=bselectRejectedReason
End Function

'[Enter Comments in the comments text box]
Public Function setCommentsTextbox_AppPage(strcomment)
	bsetCommentsTextbox_AppPage=true
	CASAOnboarding.UpdateComments.Set strcomment
   If Err.Number<>0 Then
       bsetCommentsTextbox_AppPage=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comments" ,false
       Exit Function
   End If
   setCommentsTextbox_AppPage=true
End Function

'[Click Update Button in CASA Application Page]
Public Function clickButtonUpdateApplication()
   bDevPending=true
   Wait 1
   CASAOnboarding.ButtonUpdateApplication.click 10,10,0 
   If Err.Number<>0 Then
       clickButtonUpdateApplication=false
            LogMessage "WARN","Verification","Failed to Click Button : UpdateApplication" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonUpdateApplication=true
End Function

'[Click Submit Button in Update Application dialog box]
Public Function clickUpdateAppButtonSubmit()
   bDevPending=true
   Wait 1
   CASAOnboarding.SubmitButton.click 10,10,0 
   If Err.Number<>0 Then
       clickButtonSubmit=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit " ,false
       Exit Function
   End If
   WaitForICallLoading
   clickUpdateAppButtonSubmit=true
End Function

'[Click Cancel Button in Update Application dialog box]
Public Function clickUpdateAppButtonCancel()
   bDevPending=true
   Wait 1
   CASAOnboarding.CancelButton.click 10,10,0 
   If Err.Number<>0 Then
       clickButtonCancel=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickUpdateAppButtonCancel=true
End Function

'[Click Button Close in Application Page]
Public Function clickButtonClose_App()
   bDevPending=true
   Wait 1
   CASAOnboarding.ButtonClose.click 10,10,0 
   If Err.Number<>0 Then
       clickButtonClose_App=false
            LogMessage "WARN","Verification","Failed to Click Button : Close" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonClose_App=true
End Function

'[Verify Inline message displayed as]
Public Function verifyUpdateErrorMessage(strInLineMessage)
	bverifyUpdateErrorMessage=true
	If Not IsNull(strInLineMessage) Then
       If Not VerifyInnerText (CASAOnboarding.UpdateApplication_Error(), strInLineMessage, "Inline Error Message")Then
           bverifyUpdateErrorMessage=false
       End If
   End If
   verifyUpdateErrorMessage=bverifyUpdateErrorMessage
End Function

'[Verify Application Status and Sub-Status displayed as]
Public Function verifyUpdatedStatus_AppPage(strStatus, StrSubStatus)
	bverifyUpdatedStatus_AppPage=true
	If Not IsNull(strStatus) Then
       If Not VerifyInnerText (CASAOnboarding.AppPanelStatus(), strStatus, "ApplicationStatus")Then
           bverifyUpdatedStatus_AppPage=false
       End If
   End If
   If Not IsNull(StrSubStatus) Then
       If Not VerifyInnerText (CASAOnboarding.AppPanelSubStatus(), StrSubStatus, "ApplicationSubStatus")Then
           bverifyUpdatedStatus_AppPage=false
       End If
   End If
   verifyUpdatedStatus_AppPage=bverifyUpdatedStatus_AppPage
End Function

'[Verify User an existing customer in Application page]
Public Function verifyExistingCustomer_AppPage(strApplicationID)
	bverifyExistingCustomer_AppPage=true
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select Existing_customer from iserve_casa_appln_details where app_id ='"&strApplicationID&"'"
	strQueryExistingCustomer=getDBValForColumn_NZD(strQuery)(0)
	strQueryExistingCustomer= Ucase(Trim(strQueryExistingCustomer))
	If  strQueryExistingCustomer = "TRUE" Then
		If Not VerifyInnerText (CASAOnboarding.AppPanelExistingCustomer(), "Existing Customer", "Existing Customer")Then
           bverifyExistingCustomer_AppPage=false
        End If
    End If
   verifyExistingCustomer_AppPage=bverifyExistingCustomer_AppPage
End Function

'[Verify Channel Information displayed in Application page]
Public Function verifyChannelInfo_AppPage(strApplicationID)
	bverifyChannelInfo_AppPage=true
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select channel from Iserve_Casa_Application where application_id ='"&strApplicationID&"'"
	strQueryChannel=getDBValForColumn_NZD(strQuery)(0)
	strQueryChannel= Ucase(Trim(strQueryChannel))
	If Not VerifyInnerText (CASAOnboarding.AppPanelChannel(), strQueryChannel, "Channel")Then
        bverifyChannelInfo_AppPage=false
    End If
   verifyChannelInfo_AppPage=bverifyChannelInfo_AppPage
End Function

'[Verify Application Comment table displayed as]
Public Function verifytableComment_AppPage(StrRole,strCSO,strcomment,strRejectReason)
	lstlstCommentstableData= (checknull("(Maker/Checker:"&StrRole&"|User ID:"&strCSO&"|Comments:"&strcomment&"|Reject Reason:"&strRejectReason&")|"))
	verifytableComment_AppPage=verifyTableContentList(CASAOnboarding.Commentstableheader,CASAOnboarding.CommentstableContent,lstlstCommentstableData,"Comments table",false,null,null,null)
End Function

'[Verify Update Application Button disabled]
Public Function VerifyButtonUpdate_AppPage()
	bDevPending=false
    bVerifyButtonUpdate_AppPage=true
	intBtnUpdateApp=Instr(CASAOnboarding.ButtonUpdateApplication.GetROproperty("outerhtml"),("v-Enabled"))
	If  intBtnUpdateApp = 1 Then
		LogMessage "RSLT","Verification","Update Application button is disabled as expected.",True
		bVerifyButtonUpdate_AppPage=true
	Else
		LogMessage "WARN","Verifiation","Update Application button is enabled.",false
		bVerifyButtonUpdate_AppPage=false
	End If
	VerifyButtonUpdate_AppPage=bVerifyButtonUpdate_AppPage
End Function

'[Verify Button display in Application Page]
Public Function VerifyButtondisplay_AppPage()
	bDevPending=false
    bVerifyButtondisplay_AppPage=true
	intBtnUpdateApp= Instr(CASAOnboarding.ButtonUpdateApplication.GetROproperty("outerhtml"),("v-Enabled"))
	intBtnPrintAll =Instr(CASAOnboarding.ButtonPrintAll.GetROproperty("outerhtml"),("v-Enabled"))
	intBtnPrintApplication = Instr(CASAOnboarding.ButtonPrintApplication.GetROproperty("outerhtml"),("v-Enabled"))
	If  intBtnSubmit = 1 and intBtnPrintAll = 1 And intBtnPrintApplication = 1 Then
		LogMessage "RSLT","Verification","TM/CSO User are unable to MOdify the Application status and All Buttons are disabled as expected .",True
		bVerifyButtondisplay_AppPage=true
	Else
		LogMessage "WARN","Verifiation","TM/CSO User are allowed to edit and Buttons are ediatable in Application Page.",false
		bVerifyButtondisplay_AppPage=false
	End If
	VerifyButtondisplay_AppPage=bVerifyButtondisplay_AppPage
End Function

'[Verify display of Export to Excel Button in disabled mode]
Public Function VerifyButtonExportExcelDisabled_CASA()
	bDevPending=false
    bVerifyButtonExportExcelDisabled_CASA=true
	intBtnSubmit=Instr(CASAOnboarding.ExportExcelButton.GetROproperty("outerhtml"),("v-Enabled"))
	If  intBtnSubmit = 0 Then
		LogMessage "RSLT","Verification","Export to Excel Button disabled as expected.",True
		bVerifyButtonExportExcelDisabled_CASA=true
	Else
		LogMessage "WARN","Verifiation","Export to Excel Button is enabled.",false
		bVerifyButtonExportExcelDisabled_CASA=false
	End If
	VerifyButtonExportExcelDisabled_CASA=bVerifyButtonExportExcelDisabled_CASA
End Function

'[Verify display of Export to Excel Button in enabled mode]
Public Function VerifyButtonExportExcel_CASA()
	bDevPending=false
    bVerifyButtonExportExcel_CASA=true
	intBtnSubmit=Instr(CASAOnboarding.ExportExcelButton.GetROproperty("outerhtml"),("v-disabled"))
	If  intBtnSubmit = 0 Then
		LogMessage "RSLT","Verification","Export to Excel Button enabled as expected.",True
		bVerifyButtonExportExcel_CASA=true
	Else
		LogMessage "WARN","Verifiation","Export to Excel Button is not enabled.",false
		bVerifyButtonExportExcel_CASA=false
	End If
	bVerifyButtonExportExcel_CASA=bVerifyButtonExportExcel_CASA
End Function

'[Validate Pagination of CASA Home Page]
Public Function validatePagination_CASA()
   bvalidatePagination_CASA=true
		 intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSearchRecordHeader,CSO_TM_Home.tblSearchRecordContent, "Created On")
		 If intRecordCount <=10 Then
			LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to 10", true
			bvalidatePagination_CASA=true
		  Else
			LogMessage "WARN","Verification","Number of records displayed per page is more than 10 record. Expected Count is less than or equal to 10, Actual "&intRecordCount, false
			bvalidatePagination_CASA=false
		  End If

		  If intRecordCount < 10 Then
				bNextPageExist =matchStr(lnkNext1().GetROProperty("outerhtml"),"v-disabled")
				If Not bNextPageExist Then
				LogMessage "WARN","Verification","Next link expected to be disable if record is less than 10. Currently it is enable.",false
				bvalidatePagination_CASA=false
				Else
				LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
				bvalidatePagination_CASA=true
				End If
			End If
			validatePagination_CASA=bvalidatePagination_CASA
End Function

'[Verify Iserve details displayed in Application Page]
Public Function verifyIservedetails_AppPage(strApplicationID)
	bverifyIservedetails_AppPage=true
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select maker_ID, checker_ID, checker_Date,maker_Date from iserve_casa_audit where app_id ='"&strApplicationID&"'"
	strQueryIserveDetailsArray=getDBValForMultipleColumn_NZD(strQuery)
	DBMakerID = strQueryIserveDetailsArray(0,0)
	DBCheckerID = strQueryIserveDetailsArray(0,1)
	DBCheckerDate = Left(strQueryIserveDetailsArray(0,2),10)
	DBCheckerDate = fConvertDate(DBCheckerDate)
    DBMakerDate = Left(strQueryIserveDetailsArray(0,3),10)
    DBMakerDate = fConvertDate(DBMakerDate)
	
	' Iserver details fetched from Iserve Application page 
	IserveMakerID = CASAOnboarding.LastMaker.GetROProperty("innertext")
	IserveMakerDate = CASAOnboarding.LastMakerDate.GetROProperty("innertext")
	IserveCheckerID = CASAOnboarding.LastChecker.GetROProperty("innertext")
	IserveCheckerDate = CASAOnboarding.LastCheckerDate.GetROProperty("innertext")
	
	IF Ucase(Trim(DBMakerID)) =  Ucase(Trim(IserveMakerID)) Then
		LogMessage "WARN","Verification","MakerID matched with the expected Value Expected: "&DBMakerID&" Actual: "&IserveMakerID&"",True
	Else
		LogMessage "WARN","Verification","MakerID doesnt match with the expected Value Expected: "&DBMakerID&" Actual: "&IserveMakerID&"",False	
	End IF 
	
	IF Ucase(Trim(DBMakerDate)) =  Ucase(Trim(IserveMakerDate)) Then
		LogMessage "WARN","Verification","Maker Date matched with the expected Value Expected: "&DBMakerDate&" Actual: "&IserveMakerDate&"",True
	Else
		LogMessage "WARN","Verification","Maker Date doesnt match with the expected Value Expected: "&DBMakerDate&" Actual: "&IserveMakerDate&"",False	
	End IF 
	
	IF Ucase(Trim(DBCheckerID)) =  Ucase(Trim(IserveCheckerID)) Then
		LogMessage "WARN","Verification","CheckerID matched with the expected Value Expected: "&DBCheckerID&" Actual: "&IserveCheckerID&"",True
	Else
		LogMessage "WARN","Verification","CheckerID doesnt match with the expected Value Expected: "&DBCheckerID&" Actual: "&IserveCheckerID&"",False	
	End IF 

	IF Ucase(Trim(DBCheckerDate)) =  Ucase(Trim(IserveCheckerDate)) Then
		LogMessage "WARN","Verification","Checker Date matched with the expected Value Expected: "&DBCheckerDate&" Actual: "&IserveCheckerDate&"",True
	Else
		LogMessage "WARN","Verification","Checker Date doesnt match with the expected Value Expected: "&DBCheckerDate&" Actual: "&IserveCheckerDate&"",False	
	End IF 
	
End Function

'[Verify the personal Details displayed in the Application Page]
Public Function VerifyApplicationPersonalDetails(strApplicationID)
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select salutation,first_name,last_name,ic_no,email,mobile,alt_contact_no,race,name_on_card,gender,dob from iserve_casa_appln_details where app_id ='"&strApplicationID&"'"
	strResultPersonalDetailsArray=getDBValForMultipleColumn_NZD(strQuery)
	DBSalutation= strResultPersonalDetailsArray(0,0)
	DBFirstName = strResultPersonalDetailsArray(0,1)
	DBLastName = strResultPersonalDetailsArray(0,2)
	DBNRIC = strResultPersonalDetailsArray(0,3)
	DBEmail = strResultPersonalDetailsArray(0,4)
	DBMobile = strResultPersonalDetailsArray(0,5)
	DBALTNo = strResultPersonalDetailsArray(0,6)
	DBRace = strResultPersonalDetailsArray(0,7)
	DBNameOnCard = strResultPersonalDetailsArray(0,8)
	DBGender = strResultPersonalDetailsArray(0,9)
	DBDOB = strResultPersonalDetailsArray(0,10)
	DBDOB = fConvertDate(DBDOB)
	
	' Customer information fetched from Iserve Application page 
	IserveSalutation = CASAOnboarding.ApplicationSalutation.GetROProperty("innertext")
	IserveFirstName = CASAOnboarding.ApplicationFirstName.GetROProperty("innertext")
	IserveLastName = CASAOnboarding.ApplicationLastName.GetROProperty("innertext")
	IservePassportNumber = CASAOnboarding.ApplicationNRICPassport.GetROProperty("innertext")
	IserveEmailAddress = CASAOnboarding.ApplicationEmailAddress.GetROProperty("innertext")
	IserveMobileNumber = CASAOnboarding.ApplicationMobileNumber.GetROProperty("innertext")
	IserveAltContactNumber = CASAOnboarding.ApplicationAltContactNo.GetROProperty("innertext")
	IserveRace = CASAOnboarding.ApplicationRace.GetROProperty("innertext")
	IserveCardName = CASAOnboarding.ApplicationNameonCard.GetROProperty("innertext")
	IserveGender = CASAOnboarding.ApplicationGender.GetROProperty("innertext")
	IserveDOB = CASAOnboarding.ApplicationDOB.GetROProperty("innertext")

	IF Ucase(Trim(DBSalutation)) =  Ucase(Trim(IserveSalutation)) Then
		LogMessage "RSLT","Verification","Salutation value matched with the expected Value Expected: "&DBSalutation&" Actual: "&IserveSalutation&"",True
	Else
		LogMessage "WARN","Verification","Salutation value matched with the expected Value Expected: "&DBSalutation&" Actual: "&IserveSalutation&"",False	
	End IF 
	
	IF Ucase(Trim(DBFirstName)) =  Ucase(Trim(IserveFirstName)) Then
		LogMessage "RSLT","Verification","FirstName value matched with the expected Value Expected: "&DBFirstName&" Actual: "&IserveFirstName&"",True
	Else
		LogMessage "WARN","Verification","FirstName value matched with the expected Value Expected: "&DBFirstName&" Actual: "&IserveFirstName&"",False	
	End IF 
	
	IF Ucase(Trim(DBLastName)) =  Ucase(Trim(IserveLastName)) Then
		LogMessage "RSLT","Verification","LastName value matched with the expected Value Expected: "&DBLastName&" Actual: "&IserveLastName&"",True
	Else
		LogMessage "WARN","Verification","LastName value matched with the expected Value Expected: "&DBLastName&" Actual: "&IserveLastName&"",False	
	End IF 
	
	IF Ucase(Trim(DBNRIC)) =  Ucase(Trim(IservePassportNumber)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBNRIC&" Actual: "&IservePassportNumber&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBNRIC&" Actual: "&IservePassportNumber&"",False	
	End IF 
	
	IF Ucase(Trim(DBEmail)) =  Ucase(Trim(IserveEmailAddress)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBEmail&" Actual: "&IserveEmailAddress&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBEmail&" Actual: "&IserveEmailAddress&"",False	
	End IF 
	
	IF Ucase(Trim(DBMobile)) =  Ucase(Trim(IserveMobileNumber)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBMobile&" Actual: "&IserveMobileNumber&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBMobile&" Actual: "&IserveMobileNumber&"",False	
	End IF 
	
	IF Ucase(Trim(DBALTNo)) =  Ucase(Trim(IserveAltContactNumber)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBALTNo&" Actual: "&IserveAltContactNumber&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBALTNo&" Actual: "&IserveAltContactNumber&"",False	
	End IF 
	
	IF Ucase(Trim(DBRace)) =  Ucase(Trim(IserveRace)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBRace&" Actual: "&IserveRace&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBRace&" Actual: "&IserveRace&"",False	
	End IF 
	
	IF Ucase(Trim(DBGender)) =  Ucase(Trim(IserveGender)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBGender&" Actual: "&IserveGender&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBGender&" Actual: "&IserveGender&"",False	
	End IF 
	
	IF Ucase(Trim(DBDOB)) =  Ucase(Trim(IserveDOB)) Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&DBDOB&" Actual: "&IserveDOB&"",True
	Else
		LogMessage "WARN","Verification","Nationality value matched with the expected Value Expected: "&DBDOB&" Actual: "&IserveDOB&"",False	
	End IF 
		
End Function

'[Verify Residential Address displayed in Application Page]
Public Function VerifyResidentialAddress(strApplicationID)
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select country, postal_code, house_no, level_no, unit_no, street1, street2, street3, same_as_residential from iserve_casa_address_detail where address_type= 'Residential' and app_id = '"&strApplicationID&"'"
	strResultResidentialAddArray=getDBValForMultipleColumn_NZD(strQuery)
	DBResidentialCountry = strResultResidentialAddArray(0,0)
	DBResidentialPostcalCode = strResultResidentialAddArray(0,1)
	DBResidentialHouseNo = strResultResidentialAddArray(0,2)
	DBResidentialLevelNo = strResultResidentialAddArray(0,3)
	DBResidentialUnitNo = strResultResidentialAddArray(0,4)
	DBResidentialwholeUnit = ""&DBResidentialHouseNo&" / "&DBResidentialLevelNo& "/ "&DBResidentialUnitNo&""
	DBResidentialStreet1 = strResultResidentialAddArray(0,5)
	DBResidentialStreet2 = strResultResidentialAddArray(0,6)
	DBResidentialStreet3 = strResultResidentialAddArray(0,7)
	DBResidentialAddress = strResultResidentialAddArray(0,8)
	
	' Customer information fetched from Iserve Application page 
	IserveResidentialCountry= CASAOnboarding.ResidentialCountry.GetROProperty("innertext")
	IserveResidentialPostcalCode = CASAOnboarding.ResidentialPostalCode.GetROProperty("innertext")
	IserveResidentialHouseNo = CASAOnboarding.ResidentialHouseNo.GetROProperty("innertext")
	IserveResidentialStreet1 = CASAOnboarding.ResidentialStreet1.GetROProperty("innertext")
	IserveResidentialStreet2 = CASAOnboarding.ResidentialStreet2.GetROProperty("innertext")
	IserveResidentialStreet3 = CASAOnboarding.ResidentialStreet3.GetROProperty("innertext")
	IserveResidentialAddress = CASAOnboarding.ResidentialAddressFlag.GetROProperty("innertext")

	IF Ucase(Trim(DBResidentialCountry)) =  Ucase(Trim(IserveResidentialCountry)) Then
		LogMessage "RSLT","Verification","Residential Country value matched with the expected Value Expected: "&DBResidentialCountry&" Actual: "&IserveResidentialCountry&"",True
	Else
		LogMessage "WARN","Verification","Residential Country value doesnt match with the expected Value Expected: "&DBResidentialCountry&" Actual: "&IserveResidentialCountry&"",False	
	End IF 

	IF Ucase(Trim(DBResidentialPostcalCode)) =  Ucase(Trim(IserveResidentialPostcalCode)) Then
		LogMessage "RSLT","Verification","Residential Postal code value matched with the expected Value Expected: "&DBResidentialPostcalCode&" Actual: "&IserveResidentialPostcalCode&"",True
	Else
		LogMessage "WARN","Verification","Residential Postal code doesnt match with the expected Value Expected: "&DBResidentialPostcalCode&" Actual: "&IserveResidentialPostcalCode&"",False	
	End IF 
	
	IF Ucase(Trim(DBResidentialwholeUnit)) =  Ucase(Trim(IserveResidentialHouseNo)) Then
		LogMessage "RSLT","Verification","Residential Whole Unit Number matched with the expected Value Expected: "&DBResidentialwholeUnit&" Actual: "&IserveResidentialHouseNo&"",True
	Else
		LogMessage "WARN","Verification","Residential Whole Unit Number doesnt match with the expected Value Expected: "&DBResidentialwholeUnit&" Actual: "&IserveResidentialHouseNo&"",False	
	End IF 
	
	IF Ucase(Trim(DBResidentialStreet1)) =  Ucase(Trim(IserveResidentialStreet1)) Then
		LogMessage "RSLT","Verification","Residential Street1 value matched with the expected Value Expected: "&DBResidentialStreet1&" Actual: "&IserveResidentialStreet1&"",True
	Else
		LogMessage "WARN","Verification","Residential Street1 value doesnt match with the expected Value Expected: "&DBResidentialStreet1&" Actual: "&IserveResidentialStreet1&"",False	
	End IF 
	
	IF Ucase(Trim(DBResidentialStreet2)) =  Ucase(Trim(IserveResidentialStreet2)) Then
		LogMessage "RSLT","Verification","Residential Street2 value matched with the expected Value Expected: "&DBResidentialStreet2&" Actual: "&IserveResidentialStreet2&"",True
	Else
		LogMessage "WARN","Verification","Residential Street2 value doesnt match with the expected Value Expected: "&DBResidentialStreet2&" Actual: "&IserveResidentialStreet2&"",False	
	End IF 
	
	IF Ucase(Trim(DBResidentialStreet3)) =  Ucase(Trim(IserveResidentialStreet3)) Then
		LogMessage "RSLT","Verification","Residential Street3 value matched with the expected Value Expected: "&DBResidentialStreet3&" Actual: "&IserveResidentialStreet3&"",True
	Else
		LogMessage "WARN","Verification","Residential Street3 value doesnt match with the expected Value Expected: "&DBResidentialStreet3&" Actual: "&IserveResidentialStreet3&"",False	
	End IF 
	
	IF Ucase(Trim(DBResidentialAddress)) =  Ucase(Trim(IserveResidentialAddress)) Then
		LogMessage "RSLT","Verification","Residential Address value matched with the expected Value Expected: "&DBResidentialAddress&" Actual: "&IserveResidentialAddress&"",True
	Else
		LogMessage "WARN","Verification","Residential Address value doesnt match with the expected Value Expected: "&DBResidentialAddress&" Actual: "&IserveResidentialAddress&"",False	
	End IF 


End Function

'[Verify Mailing Address displayed in Application Page]
Public Function VerifyResidentialMailingAddress(strApplicationID)
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select country, postal_code, house_no, level_no, unit_no, street1, street2, street3, same_as_residential from iserve_casa_address_detail where address_type= 'Mailing' and app_id = '"&strApplicationID&"'"
	strResultMailingAddArray=getDBValForMultipleColumn_NZD(strQuery)
	DBMailingCountry = strResultMailingAddArray(0,0)
	DBMailingPostcalCode = strResultMailingAddArray(0,1)
	DBMailingHouseNo = strResultMailingAddArray(0,2)
	DBMailingLevelNo = strResultMailingAddArray(0,3)
	DBMailingUnitNo = strResultMailingAddArray(0,4)
	DBMailingwholeUnit = ""&DBMailingHouseNo&" / "&DBMailingLevelNo& "/ "&DBMailingUnitNo&""
	DBMailingStreet1 = strResultMailingAddArray(0,5)
	DBMailingStreet2 = strResultMailingAddArray(0,6)
	DBMailingStreet3 = strResultMailingAddArray(0,7)
	DBMailingAddress = strResultMailingAddArray(0,8)
	
	' Customer information fetched from Iserve Application page 
	IserveMailingCountry= CASAOnboarding.MailingCountry.GetROProperty("innertext")
	IserveMailingPostcalCode = CASAOnboarding.MailingPostalCode.GetROProperty("innertext")
	IserveMailingHouseNo = CASAOnboarding.MailingHouseNo.GetROProperty("innertext")
	IserveMailingStreet1 = CASAOnboarding.MailingStreet1.GetROProperty("innertext")
	IserveMailingStreet2 = CASAOnboarding.MailingStreet2.GetROProperty("innertext")
	IserveMailingStreet3 = CASAOnboarding.MailingStreet3.GetROProperty("innertext")
	IserveMailingAddress = CASAOnboarding.MailingAddressFlag.GetROProperty("innertext")
	
	IF Ucase(Trim(DBMailingCountry)) =  Ucase(Trim(IserveMailingCountry)) Then
		LogMessage "RSLT","Verification","Mailing Country value matched with the expected Value Expected: "&DBMailingCountry&" Actual: "&IserveMailingCountry&"",True
	Else
		LogMessage "WARN","Verification","Mailing Country value doesnt match with the expected Value Expected: "&DBMailingCountry&" Actual: "&IserveMailingCountry&"",False	
	End IF 

	IF Ucase(Trim(DBMailingPostcalCode)) =  Ucase(Trim(IserveMailingPostcalCode)) Then
		LogMessage "RSLT","Verification","Mailing Postal code value matched with the expected Value Expected: "&DBMailingPostcalCode&" Actual: "&IserveMailingPostcalCode&"",True
	Else
		LogMessage "WARN","Verification","Mailing Postal code doesnt match with the expected Value Expected: "&DBMailingPostcalCode&" Actual: "&IserveMailingPostcalCode&"",False	
	End IF 
	
	IF Ucase(Trim(DBMailingwholeUnit)) =  Ucase(Trim(IserveMailingHouseNo)) Then
		LogMessage "RSLT","Verification","Mailing Whole Unit Number matched with the expected Value Expected: "&DBMailingwholeUnit&" Actual: "&IserveMailingHouseNo&"",True
	Else
		LogMessage "WARN","Verification","Mailing Whole Unit Number doesnt match with the expected Value Expected: "&DBMailingwholeUnit&" Actual: "&IserveMailingHouseNo&"",False	
	End IF 
	
	IF Ucase(Trim(DBMailingStreet1)) =  Ucase(Trim(IserveMailingStreet1)) Then
		LogMessage "RSLT","Verification","Mailing Street1 value matched with the expected Value Expected: "&DBMailingStreet1&" Actual: "&IserveMailingStreet1&"",True
	Else
		LogMessage "WARN","Verification","Mailing Street1 value doesnt match with the expected Value Expected: "&DBMailingStreet1&" Actual: "&IserveMailingStreet1&"",False	
	End IF 
	
	IF Ucase(Trim(DBMailingStreet2)) =  Ucase(Trim(IserveMailingStreet2)) Then
		LogMessage "RSLT","Verification","Mailing Street2 value matched with the expected Value Expected: "&DBMailingStreet2&" Actual: "&IserveMailingStreet2&"",True
	Else
		LogMessage "WARN","Verification","Mailing Street2 value doesnt match with the expected Value Expected: "&DBMailingStreet2&" Actual: "&IserveMailingStreet2&"",False	
	End IF 
	
	IF Ucase(Trim(DBMailingStreet3)) =  Ucase(Trim(IserveMailingStreet3)) Then
		LogMessage "RSLT","Verification","Mailing Street3 value matched with the expected Value Expected: "&DBMailingStreet3&" Actual: "&IserveMailingStreet3&"",True
	Else
		LogMessage "WARN","Verification","Mailing Street3 value doesnt match with the expected Value Expected: "&DBMailingStreet3&" Actual: "&IserveMailingStreet3&"",False	
	End IF 

	IF Ucase(Trim(DBMailingAddress)) =  Ucase(Trim(IserveMailingAddress)) Then
		LogMessage "RSLT","Verification","Mailing Address value matched with the expected Value Expected: "&DBMailingAddress&" Actual: "&IserveMailingAddress&"",True
	Else
		LogMessage "WARN","Verification","Mailing Address value doesnt match with the expected Value Expected: "&DBMailingAddress&" Actual: "&IserveMailingAddress&"",False	
	End IF 
	
End Function

'[Verify FATCA and income details displayed in Application page]
Public Function VerifyFATCADeclaration(strApplicationID)
	strApplicationID = Environment.Value("ApplicationID")
	strQuery = "Select usa_resident, usa_citizen, usa_green_card, occupation_type,sel_emp_type, sel_emp_desc,income_source, annual_sal, est_net_worth,purpose_desc from iserve_casa_other_details where app_id ='"&strApplicationID&"'"
	strQueryOtherDetailsAddArray=getDBValForMultipleColumn_NZD(strQuery)
	DBFATCAResident = strQueryOtherDetailsAddArray(0,0)
	DBFATCACitizen = strQueryOtherDetailsAddArray(0,1)
	DBFATCAGreenCardHolder = strQueryOtherDetailsAddArray(0,2)
	DBOccupationType = strQueryOtherDetailsAddArray(0,3)
	'DBSelfEmployedType = strQueryOtherDetailsAddArray(0,4)
	DBSelfEmployedDetails = strQueryOtherDetailsAddArray(0,5)
	DBPrimaryIncome = strQueryOtherDetailsAddArray(0,6)
	DBAnnualIncome = strQueryOtherDetailsAddArray(0,7)
	DBNetworth = strQueryOtherDetailsAddArray(0,8)
	DBPurposeOfAccount = strQueryOtherDetailsAddArray(0,9)
	
	' Customer information fetched from Iserve Application page 
	IserveFATCAResident= CASAOnboarding.USResidentFlag.GetROProperty("innertext")
	IserveFATCACitizen = CASAOnboarding.USCitizenFlag.GetROProperty("innertext")
	IserveFATCAGreenCardHolder = CASAOnboarding.GreenCardHolder.GetROProperty("innertext")
	IserveOccupationType= CASAOnboarding.OccuptionType.GetROProperty("innertext")
	'IserveSelfEmployedType= CASAOnboarding.SelfEmployedType.GetROProperty("innertext")
	IserveSelfEmployedDetails= CASAOnboarding.SelfEmployedDetails.GetROProperty("innertext")
	IservePrimaryIncome = CASAOnboarding.PrimarySourceIncome.GetROProperty("innertext")
	IserveAnnualIncome = CASAOnboarding.AnnualIncomeRange.GetROProperty("innertext")
	IserveNetworth = CASAOnboarding.EstimatedNetworth.GetROProperty("innertext")
	IservePurposeOfAccount = CASAOnboarding.IncomeAccountPurpose.GetROProperty("innertext")

	IF Ucase(Trim(DBFATCAResident)) = "FALSE" and Ucase(Trim(IserveFATCAResident)) = "NO" Then
		LogMessage "RSLT","Verification","US Resident value matched with the expected Value Expected: "&DBFATCAResident&" Actual: "&IserveFATCAResident&"",True
	Else IF Ucase(Trim(DBFATCAResident)) = "TRUE" and Ucase(Trim(IserveFATCAResident)) = "YES" Then
		LogMessage "RSLT","Verification","US Resident value matched with the expected Value Expected: "&DBFATCAResident&" Actual: "&IserveFATCAResident&"",True 
	Else
		LogMessage "WARN","Verification","US Resident value doesnt match with the expected Value Expected: "&DBFATCAResident&" Actual: "&IserveFATCAResident&"",False	
	End IF 
	End If 
	
	IF Ucase(Trim(DBFATCACitizen)) = "FALSE"  AND  Ucase(Trim(IserveFATCACitizen)) = "NO" Then
		LogMessage "RSLT","Verification","US Citizen value matched with the expected Value Expected: "&DBFATCACitizen&" Actual: "&IserveFATCACitizen&"",True
	Else IF Ucase(Trim(DBFATCACitizen)) = "TRUE"  AND  Ucase(Trim(IserveFATCACitizen)) = "YES" Then
		LogMessage "RSLT","Verification","US Citizen value matched with the expected Value Expected: "&DBFATCACitizen&" Actual: "&IserveFATCACitizen&"",True
	Else
		LogMessage "WARN","Verification","US Citizen value doesnt match with the expected Value Expected: "&DBFATCACitizen&" Actual: "&IserveFATCACitizen&"",False	
	End IF 
	End If 
	
	IF Ucase(Trim(DBFATCAGreenCardHolder)) = "FALSE" AND Ucase(Trim(IserveFATCAGreenCardHolder)) = "NO" Then
		LogMessage "RSLT","Verification","GreenCard Holder value matched with the expected Value Expected: "&DBFATCAGreenCardHolder&" Actual: "&IserveFATCAGreenCardHolder&"",True
	Else  IF Ucase(Trim(DBFATCAGreenCardHolder)) = "TRUE" AND Ucase(Trim(IserveFATCAGreenCardHolder)) = "YES" Then
		LogMessage "RSLT","Verification","GreenCard Holder value matched with the expected Value Expected: "&DBFATCAGreenCardHolder&" Actual: "&IserveFATCAGreenCardHolder&"",True
	Else
		LogMessage "WARN","Verification","GreenCard Holder value doesnt match with the expected Value Expected: "&DBFATCAGreenCardHolder&" Actual: "&IserveFATCAGreenCardHolder&"",False	
	End IF 
	End If 
	
	IF Ucase(Trim(DBOccupationType)) = Ucase(Trim(IserveOccupationType)) Then
		LogMessage "RSLT","Verification","OccupationType matched with the expected Value Expected: "&DBOccupationType&" Actual: "&IserveOccupationType&"",True
	Else
		LogMessage "WARN","Verification","OccupationType doesnt match with the expected Value Expected: "&DBOccupationType&" Actual: "&IserveOccupationType&"",False	
	End IF 
	'commenting as this value is not captured in Vaadin as well. 
'	IF Ucase(Trim(DBSelfEmployedType)) = Ucase(Trim(IserveSelfEmployedType)) Then
'		LogMessage "WARN","Verification","SelfEmployed Type matched with the expected Value Expected: "&DBSelfEmployedType&" Actual: "&IserveSelfEmployedType&"",True
'	Else
'		LogMessage "WARN","Verification","SelfEmployed Type doesnt match with the expected Value Expected: "&DBSelfEmployedType&" Actual: "&IserveSelfEmployedType&"",False	
'	End IF 
	
	IF Ucase(Trim(DBSelfEmployedDetails)) = Ucase(Trim(IserveSelfEmployedDetails)) Then
		LogMessage "RSLT","Verification","SelfEmployed Details matched with the expected Value Expected: "&DBSelfEmployedDetails&" Actual: "&IserveSelfEmployedDetails&"",True
	Else
		LogMessage "WARN","Verification","SelfEmployed Details doesnt match with the expected Value Expected: "&DBSelfEmployedDetails&" Actual: "&IserveSelfEmployedDetails&"",False	
	End IF 
	
	IF Ucase(Trim(DBPrimaryIncome)) =  Ucase(Trim(IservePrimaryIncome)) Then
		LogMessage "RSLT","Verification","Primary Income matched with the expected Value Expected: "&DBPrimaryIncome&" Actual: "&IservePrimaryIncome&"",True
	Else
		LogMessage "WARN","Verification","Primary Income doesnt match with the expected Value Expected: "&DBPrimaryIncome&" Actual: "&IservePrimaryIncome&"",False	
	End IF 
	
	IF Ucase(Trim(DBAnnualIncome)) =  Ucase(Trim(IserveAnnualIncome))Then
		LogMessage "RSLT","Verification","Annual Income matched with the expected Value Expected: "&DBAnnualIncome&" Actual: "&IserveAnnualIncome&"",True
	Else
		LogMessage "WARN","Verification","Annual Income doesnt match with the expected Value Expected: "&DBAnnualIncome&" Actual: "&IserveAnnualIncome&"",False	
	End IF 
	
	IF Ucase(Trim(DBNetworth)) =  Ucase(Trim(IserveNetworth)) Then
		LogMessage "RSLT","Verification","Estimated Networth matched with the expected Value Expected: "&DBNetworth&" Actual: "&IserveNetworth&"",True
	Else
		LogMessage "WARN","Verification","Estimated Networth doesnt match with the expected Value Expected: "&DBNetworth&" Actual: "&IserveNetworth&"",False	
	End IF 
	
	IF Ucase(Trim(DBPurposeOfAccount)) =  Ucase(Trim(IservePurposeOfAccount)) Then
		LogMessage "RSLT","Verification","Reason for Account Openig with DBS value matched with the expected Value Expected: "&DBPurposeOfAccount&" Actual: "&IservePurposeOfAccount&"",True
	Else
		LogMessage "WARN","Verification","Reason for Account Openig with DBS value doesnt match with the expected Value Expected: "&DBPurposeOfAccount&" Actual: "&IservePurposeOfAccount&"",False	
	End IF 
		
End Function

'[Verify Declarations displayed in Application page]
Public Function VerifyDeclarations(strApplicationID)
	'strApplicationID = Environment.Value("ApplicationID")
	strQuery = "select App_Submit_Status , Marketing from ISERVE_CASA_APPLN_DETAILS where app_id = '"&strApplicationID&"'"
	strQueryDeclarationAddArray=getDBValForMultipleColumn_NZD(strQuery)
	DBTermsConditions = strQueryDeclarationAddArray(0,0)
	DBMarketing = strQueryDeclarationAddArray(0,1)
		
	' Customer information fetched from Iserve Application page 
	IserveTermsConditions= CASAOnboarding.TermsConditionsFlag.GetROProperty("innertext")
	IserveSMSMarketing = CASAOnboarding.SMSMarketingFlag.GetROProperty("innertext")
	IserveEmailMarketing = CASAOnboarding.EmailMarketingFlag.GetROProperty("innertext")

	IF Ucase(Trim(DBTermsConditions)) = "COMPLETE" and Ucase(Trim(IserveTermsConditions)) = "Y" Then
		LogMessage "RSLT","Verification","Declarations Terms and conditions is set to Yes when Application submit status is complete: "&DBTermsConditions&" Actual: "&IserveTermsConditions&"",True
		If instr(DBMarketing, "email")= 1 AND  IserveEmailMarketing = "Y" Then
		   LogMessage "RSLT","Verification","Agreed to Email Marketing is set to Yes when Application submit status is complete: "&DBMarketing&" Actual: "&IserveEmailMarketing&"",True	
		Else If instr(DBMarketing, "email")= 0 AND  IserveEmailMarketing = "N" Then
			LogMessage "RSLT","Verification","Agreed to Email Marketing is set to No as expected: "&DBMarketing&" Actual: "&IserveEmailMarketing&"",True	
			 End If
		End If 
		If instr(DBMarketing, "mobile") = 1 AND  IserveSMSMarketing = "Y" Then
			LogMessage "RSLT","Verification","Agreed to SMS Marketing is set to Yes when Application submit status is complete: "&DBMarketing&" Actual: "&IserveSMSMarketing&"",True	
		Else If instr(DBMarketing, "mobile") = 0 AND  IserveSMSMarketing = "N" Then
			LogMessage "RSLT","Verification","Agreed to SMS Marketing is set to No as expected: "&DBMarketing&" Actual: "&IserveSMSMarketing&"",True	
			End If 
		End If
	Else IF Ucase(Trim(DBTermsConditions)) = "PARTIAL" Then
		If IserveTermsConditions = BLANK AND  IserveSMSMarketing = BLANK AND IserveEmailMarketing = BLANK Then
		LogMessage "RSLT","Verification","Declarations Flag are set to Blank when Application submit status is partial: "&DBTermsConditions&" Actual: "&IserveTermsConditions&"",True 
		End If 
	Else
		LogMessage "WARN","Verification","Appliation Terms and conditions doesnt match with the expected Value Expected: "&DBTermsConditions&" Actual: "&IserveTermsConditions&"",False	
	End IF 
	End If 			
End Function


'[Verify Nationality displayed in Application page]
Public Function VerifyNationality(strApplicationID)
	'strApplicationID = Environment.Value("ApplicationID")
	
	strQuery = "Select IC_Type from ISERVE_CASA_APPLN_DETAILS where App_ID = '"&strApplicationID&"'"
	strQueryICType = getDBValForColumn_NZD(strQuery)
	
	' Customer information fetched from Iserve Application page 
    IserveNationality = CASAOnboarding.ApplicationNationality.GetROProperty("innertext")
	
	If strQueryICType = "01" And IserveNationality = "Singapore NRIC" Then 
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&strQueryICType&" Actual: "&IserveNationality&"",True
	Else If strQueryICType = "02" And IserveNationality = "Singapore PR" Then
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&strQueryICType&" Actual: "&IserveNationality&"",True
	Else If strQueryICType = "03" And IserveNationality = "Malaysian" Then	
		LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&strQueryICType&" Actual: "&IserveNationality&"",True
	Else If strQueryICType = "04" Then
	   strQuery = "select A2.country_Desc from ISERVE_CASA_APPLN_DETAILS A1, ISERVE_COUNTRY_MASTER A2 where A1.App_Id = '"&strApplicationID&"' and A1.Nationality = A2.Two_Char_Code" 
	   strQueryNationality =getDBValForColumn_NZD(strQuery)
	   If Ucase(Trim(strQueryNationality)) =  Ucase(Trim(IserveNationality)) Then
		   LogMessage "RSLT","Verification","Nationality value matched with the expected Value Expected: "&strQueryNationality&" Actual: "&IserveNationality&"",True	
	   End If
	Else
		LogMessage "WARN","Verification","Nationality value doesnt match with the expected Value Expected: "&strQueryNationality&" Actual: "&IserveNationality&"",False	
	End If	
	End If	
	End If	
	End If		
End Function
