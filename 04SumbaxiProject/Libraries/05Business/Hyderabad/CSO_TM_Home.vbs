'[Select Combobox Show on Home Page as]
Public Function selectShowComboBox(strShow)
   bDevPending=false
   bSelectShowComboBox=true
   WaitForICallLoading
   If Not IsNull(strShow) Then
      If Not (selectItemFromComboBox (Browser("Browser_iCall_BlockCancelCard").Page("iCall_CSO_TM_Home"),"showMapKey",strShow))Then
            LogMessage "WARN","Verification","Failed to select :"&strShow&" From Show drop down list" ,false
           bSelectShowComboBox=false
       End If
   End If
   WaitForICallLoading
   selectShowComboBox=bSelectShowComboBox
End Function

'[Select Radio Button Type on Home Page]
Public Function selectTypeRadio(strType)
	bDevPending=False
	bSelectTypeRadio=true
	bSelectTypeRadio=SelectRadioButtonGrp(strType, CSO_TM_Home.rbtnType, Array("Open","Pending","Failed","Pending Approval", "On Hold"))
   WaitForICallLoading
	If Err.Number<>0 Then
       selectTypeRadio=false
          LogMessage "WARN","Verification","Failed to Click Button : Type on Home Page" ,false
       Exit Function
   End If
   selectTypeRadio=bSelectTypeRadio
End Function

'[Select Radio Button Type for IA on Home Page]
Public Function selectTypeRadio_IA(strType)
	bDevPending=False
	bSelectTypeRadio=true
	bSelectTypeRadio=SelectRadioButtonGrp(strType, CSO_TM_Home.rbtnType, Array("Pending","Open"))
   WaitForICallLoading
	If Err.Number<>0 Then
       selectTypeRadio=false
          LogMessage "WARN","Verification","Failed to Click Button : Type on Home Page" ,false
       Exit Function
   End If
   WaitForICallLoading
   selectTypeRadio_IA=bSelectTypeRadio
End Function

'[Click Button Add on Home Page]
Public Function clickButtonAdd()
   bDevPending=true
   CSO_TM_Home.btnAdd.click
   If Err.Number<>0 Then
       clickButtonAdd=false
            LogMessage "WARN","Verification","Failed to Click Button : Add" ,false
       Exit Function
   End If
   clickButtonAdd=true
End Function

'[Select Combobox Group on Home Page as]
Public Function selectGroupComboBox(strGroup)
   WaitForICallLoading
   bDevPending=true
   bSelectGroupComboBox=true
   If Not IsNull(strGroup) Then
       If Not (selectItemFromComboBox (Browser("Browser_iCall_BlockCancelCard").Page("iCall_CSO_TM_Home"),"groupValue",strGroup))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Group drop down list" ,false
           bSelectGroupComboBox=false
       End If
   End If
   selectGroupComboBox=bSelectGroupComboBox
End Function

'[Select Combobox Select on Home Page as]
Public Function selectSelectComboBox(strSelect)
   bDevPending=true
   bSelectSelectComboBox=true
   If Not IsNull(strSelect) Then   
       If Not (selectItemFromComboBox (Browser("Browser_iCall_BlockCancelCard").Page("iCall_CSO_TM_Home"),"selectValue", strSelect))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Select drop down list" ,false
           bSelectSelectComboBox=false
       End If
   End If
   selectSelectComboBox=bSelectSelectComboBox
End Function

'[Select value from Combobox on Home Page as]
Public Function selectAnyComboBox(strComboBoxName,strSelect)
   bDevPending=true
   bselectAnyComboBox=true
   If Not IsNull(strSelect) Then   
       If Not (selectItemFromComboBox (Browser("Browser_iCall_BlockCancelCard").Page("iCall_CSO_TM_Home"),strComboBoxName, strSelect))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Select drop down list" ,false
           bselectAnyComboBox=false
       End If
   End If
   selectAnyComboBox=bselectAnyComboBox
End Function


'[Select Combobox CSO on Home Page as]
Public Function selectCSOComboBox(strCSO)
   bDevPending=true
   bSelectCSOComboBox=true
   If Not IsNull(strCSO) Then
       If Not (selectItemFromComboBox (Browser("Browser_iCall_BlockCancelCard").Page("iCall_CSO_TM_Home"),"csoVal", strCSO))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From CSO drop down list" ,false
           bSelectCSOComboBox=false
       End If
   End If
   WaitForICallLoading
   selectCSOComboBox=bSelectCSOComboBox
End Function

'[Set TextBox From on Home Page to]
Public Function setFromTextbox(strFrom)
	bDevPending=true
	If not isNull(strFrom) Then	
		Select Case (Ucase(strFrom))
			Case "TODAY"
				If len(Day(CDate(Now)))=1 Then
					strDay="0"&Day(CDate(Now))
				else
					strDay=""&Day(CDate(Now))
				End If
				strFrom=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&""
			Case "FUTURE"
				strFutureDate = DateAdd("d",1,now)
				If len(Day(CDate(strFutureDate)))=1 Then
					strDay="0"&Day(CDate(strFutureDate))
				else
					strDay=""&Day(CDate(strFutureDate))
				End If
				strFrom=""&strDay & " "&monthName(Month(CDate(strFutureDate)),true) &" " &Year(CDate(strFutureDate))&""
				
			End Select
		CSO_TM_Home.txtFrom.Set(strFrom)
	End If
	If Err.Number<>0 Then
		setFromTextbox=false
		LogMessage "WARN","Verification","Failed to Set Text Box :From" ,false
		Exit Function
	End If
	setFromTextbox=true
End Function

'[Set TextBox To on Home Page to]
Public Function setToTextbox(strTo)
	bDevPending=true
	If not isNull(strTo) Then		
		Select Case (Ucase(strTo))
			Case "TODAY"	
				If len(Day(CDate(Now)))=1 Then
					strDays="0"&Day(CDate(Now))
				else
					strDays=""&Day(CDate(Now))
				End If
				strTo=""&strDays & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&""
			Case "FUTURE"
				strFutureDates = DateAdd("d",1,now)
				If len(Day(CDate(strFutureDates)))=1 Then
					strDays="0"&Day(CDate(strFutureDates))
				else
					strDays=""&Day(CDate(strFutureDates))
				End If
				strTo=""&strDays & " "&monthName(Month(CDate(strFutureDates)),true) &" " &Year(CDate(strFutureDates))&""		
		End Select
		CSO_TM_Home.txtTo.Set(strTo)
	End If
	If Err.Number<>0 Then
		setToTextbox=false
		LogMessage "WARN","Verification","Failed to Set Text Box :To" ,false
		Exit Function
	End If
	setToTextbox=true
End Function

'[Click Button Filter on Home Page]
Public Function clickButtonFilter()
   bDevPending=true
   Wait 1
   CSO_TM_Home.btnFilter.click 10,10,0 
   WaitForICallLoading
   If Err.Number<>0 Then
       clickButtonFilter=false
            LogMessage "WARN","Verification","Failed to Click Button : Filter" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonFilter=true
End Function

'[Verify row Data in Table SR_Activity_Records]
Public Function verifytblSR_Activity_Records_RowData(arrRowDataList)
   bDevPending=false
   verifytblSR_Activity_Records_RowData=verifyTableContentList(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent,arrRowDataList,"SR_Activity_Records" ,True,CSO_TM_Home.lnkNext ,CSO_TM_Home.lnkNext1,CSO_TM_Home.lnkPrevious)
End Function

'[Select Pending Approval SR from TM home Page]
Public Function selectPendingSR_TMHome(strCSO,strType,strAccount,strCreatedOn)
	WaitForICallLoading
   	Dim bSelectPendingSR:bSelectPendingSR=true
   	'Wait(10)
   	'Commenting as there is object id issue after 1701, and Service request will be selected as default in TM homepage 
	'selectShowComboBox("Service Requests")
	selectTypeRadio("Pending Approval")
	WaitForICallLoading
	' Click on Verify Button
	For iCounterr = 1 To 180 Step 1
		If Not CSO_TM_Home.txtStaff.Exist(0.5) Then
			Wait(0.5)
		else
			WaitForICallLoading
			CSO_TM_Home.txtStaff.Set(strCSO)
			WaitForICallLoading
			selectCSOComboBox(strCSO)
			Print "Staff filed value set done"
			Exit for
		End if
	Next   
	If Err.Number<>0 Then
		selectPendingSR_TMHome=false
		LogMessage "WARN","Verification","Failed to set Staff field value" ,false
		Exit Function
	End If	
	WaitForICallLoading
	SelectHomePageDatePickerFromDate("TODAY")
	SelectHomePageDatePicker_TODate("TODAY")
	clickButtonFilter
	WaitForIcallLoading
	Dim lstSRActivityData
	If not isnull(strCreatedOn) Then
		If strCreatedOn="SR CREATED ON" Then
			strComment=fetchFromDataStore(gstrRuntimeCommentStep,"BLANK",gstrParameterNameStep)(0)
			strQuery_CreatedOn="Select createdOn from cca_tmapproval_sr where SRTYPE ='"&strType&"' and comments LIKE '%"&strComment&"%'"
			strTST=getDBValForColumn_FE(strQuery_CreatedOn)(0)				
			strTempDate=Left(strTST, Len(strTST) - 6)
			strTime=FormatDateTime(strTempDate &" "&Right(strTST,2),4)
			strTime=Replace(strTime,":",".?")
			If len(Day(CDate(strTempDate)))=1 Then
			
				strDay="0"&Day(CDate(strTempDate))
			else
				strDay=""&Day(CDate(strTempDate))
			End If
			strCreatedDatePattern=""&strDay & " "&monthName(Month(CDate(strTempDate)),true) &" " &Year(CDate(strTempDate))&" " &strTime
			insertDataStore "CreatedOn", strCreatedDatePattern
		End If
		
		If not isnull(strAccount) Then
			ReDim lstSRActivityData(2)
		Else
			ReDim lstSRActivityData(1)
		End If
		lstSRActivityData(0)="Type:"&strType
		lstSRActivityData(1)="Created Date:"&strCreatedDatePattern
		If not isnull(strAccount) Then
			lstSRActivityData(2)="Account / Card No.:"&strAccount
		End If
		
	else
		ReDim lstSRActivityData(1)
		lstSRActivityData(0)="Type:"&strType
		lstSRActivityData(1)="Account / Card No.:"&strAccount
		
    End If
    WaitForIcallLoading
	bSelectPendingSR=clickType_SR_Activity_Records(lstSRActivityData)
	selectPendingSR_TMHome=bSelectPendingSR
End Function

'[Click Type link in Table SR_Activity Records form Home Page]
Public Function clickType_SR_Activity_Records(lstSRActivityData)
   bDevPending=false
   With CSO_TM_Home
	   clickType_SR_Activity_Records=selectTableLink(.tblSRActivityHeader,.tblSRActivityContent,lstSRActivityData,"SRActivityRecords" ,"Type",true,.lnkNext ,.lnkNext1 ,.lnkPrevious)
   End With
End Function

'[Select Failed SR from TM home Page]
Public Function selectFailedSR_TMHome(strCSO,strCIN,strType,strAccount,strCreatedOn)
 	bDevPending=false
   Dim bSelectPendingSR:bSelectPendingSR=true
	selectShowComboBox("Service Requests")
	selectTypeRadio("Failed")
	'selectCSOComboBox(strCSO)
	Browser("Browser_iCall_BlockCancelCard").Page("iCall_CSO_TM_Home").WebEdit("txtCSO").Set(strCSO)
	setFromTextbox("TODAY")
	setToTextbox("TODAY")
	clickButtonFilter
	WaitForIcallLoading
	Dim lstSRActivityData
	If not isnull(strCreatedOn) Then
		If strCreatedOn="SR CREATED ON" Then
			'strComment=fetchFromDataStore(gstrRuntimeCommentStep,"BLANK","SRComment")(0)
			strQuery_CreatedOn="select to_char(created_datetime) from orchsvc_sr where contact_cin ='"&strCIN&"' and sr_type='"&strType&"' order by created_datetime desc"
			'strTST=getDBValForColumn_CC(strQuery_CreatedOn)(0)
			strTST=getDBValForColumn_OL(strQuery_CreatedOn)(0)
			'strTST=getDBValForColumn_UAT(strQuery_CreatedOn)(0)					
			strTempDate=Left(strTST,15)
			strTime=FormatDateTime(Left(strTST,15) &" "&Right(strTST,2),4)
			strTime=Replace(strTime,":",".?")
			If len(Day(CDate(strTempDate)))=1 Then
			
				strDay="0"&Day(CDate(strTempDate))
			else
				strDay=""&Day(CDate(strTempDate))
			End If
			strCreatedDatePattern=""&strDay & " "&monthName(Month(CDate(strTempDate)),true) &" " &Year(CDate(strTempDate))&" " &strTime
			insertDataStore "CreatedOn", strCreatedDatePattern
		End If
		ReDim lstSRActivityData(2)
		lstSRActivityData(0)="Type:"&strType
		lstSRActivityData(1)="Account / Card No.:"&strAccount
		lstSRActivityData(2)="Created Date:"&strCreatedDatePattern
	else
		ReDim lstSRActivityData(1)
		lstSRActivityData(0)="Type:"&strType
		lstSRActivityData(1)="Account / Card No.:"&strAccount
		
    End If
	bSelectPendingSR=clickType_SR_Activity_Records(lstSRActivityData)
	selectFailedSR_TMHome=bSelectPendingSR
End Function

'[Select Open Activity SR from CSO home Page]
Public Function selectOpenIA_CSOHome(strCSO,strActivityType)
  bselectOpenIA_CSOHome=true
  Dim lstSRActivityData
    strCreatedBy_LAN_Id=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\SR_Screen_Map.ini",  "LAN ID" , strCSO)
	strQuery_Activity="select to_char(duedatetime) from orchsvc_ia where created_datetime = (select Max(created_datetime) FROM orchsvc_ia where created_by='"&strCreatedBy_LAN_Id&"')"
    strTST=getDBValForColumn_OL(strQuery_Activity)(0)                  
	strTempDate=Left(strTST,15)
			strTime=FormatDateTime(Left(strTST,15) &" "&Right(strTST,2),4)
			strTime=Replace(strTime,":",".?")
			If len(Day(CDate(strTempDate)))=1 Then
				strDay="0"&Day(CDate(strTempDate))
			else
				strDay=""&Day(CDate(strTempDate))
			End If
			strCreatedDatePattern=""&strDay & " "&monthName(Month(CDate(strTempDate)),true) &" " &Year(CDate(strTempDate))&" " &strTime
			insertDataStore "Due Date", strCreatedDatePattern

		ReDim lstSRActivityData(2)
		lstSRActivityData(0)="Type:"&strActivityType
		lstSRActivityData(1)="Assigned To:"&strCreatedBy_LAN_Id
		lstSRActivityData(2)="Due Date:"&strCreatedDatePattern

		bselectOpenIA_CSOHome=clickCreatedOn_Activity_Records(lstSRActivityData)
        selectOpenIA_CSOHome=bselectOpenIA_CSOHome
End Function

'[Click Created On link in Table Open Activity Records form Home Page]
Public Function clickCreatedOn_Activity_Records(lstSRActivityData)
   bDevPending=false
   WaitForIcallLoading
   With CSO_TM_Home
	   clickCreatedOn_Activity_Records=selectTableLink(.tblSRActivityHeader,.tblSRActivityContent,lstSRActivityData,"SRActivityRecords" ,"Created On",true,.lnkNext ,.lnkNext1 ,.lnkPrevious)
   End With
End Function

'[Verify CSO TM Home Page is displayed]
Public Function verifyCSOHomePage(strType, strShow)
   bverifyCSOHomePage=true
	If not VerifyInnerText(CSO_TM_Home.rbtnType(), strType,"CSO Radio Button") Then
		bverifyCSOHomePage=false
	 End If
	 If Not verifyComboSelectItem (CSO_TM_Home.lstShow(), strShow, "CSO Show")Then
           bverifyCSOHomePage=false
      End If
	verifyCSOHomePage=bverifyCSOHomePage
End Function

'[Verify Show Combobox has Items]
Public Function verifyShowComboboxItems(lstItems)
   bverifyShowComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CSO_TM_Home.lstShow(), lstItems, "CSO Show")Then
           bverifyShowComboboxItems=false
       End If
   End If
   verifyShowComboboxItems=bverifyShowComboboxItems
End Function

'[Verify available radio buttons on CSO TM Home Page]
Public Function verifyRadioButtonforServiceRequest(strType)
   bverifyRadioButtonforServiceRequest=true
	If not VerifyInnerText(CSO_TM_Home.rbtnType(), strType,"SR Radio Button") Then
		bverifyRadioButtonforServiceRequest=false
	 End If
	verifyRadioButtonforServiceRequest=bverifyRadioButtonforServiceRequest
End Function

'[Verify Radio Button Type on Home Page Screen Selected as]
Public Function verifyDefaultSRRadioSelection(strSelectedradioButton)
	bDevPending=False
	bSRRadioSelection=true
	bSRRadioSelection=VerifyRadioButtonGrpSelection(strSelectedradioButton, CSO_TM_Home.rbtnType, Array("Open","Pending","Failed","Pending Approval", "On Hold"))
	If bSRRadioSelection Then
		          LogMessage "RSLT","Verification","Radio Button :Type selected as ecxpected. Selected value is "&strSelectedradioButton ,true
		else
				LogMessage "RSLT","Verification","Radio Button :Type is not selected as ecxpected. Selected value is "&strSelectedradioButton ,false
	End If
    If Err.Number<>0 Then
       bSRRadioSelection=false
          LogMessage "WARN","Verification","Failed to Verify Radio Button :Service Requests" ,false
       Exit Function
   End If
   verifyDefaultSRRadioSelection=bSRRadioSelection
End Function

'[Verify Group Combobox has Items]
Public Function verifyGroupComboboxItems(lstItems)
   bverifyGroupComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CSO_TM_Home.lstGroup(), lstItems, "CSO Group")Then
           bverifyGroupComboboxItems=false
       End If
   End If
   verifyGroupComboboxItems=bverifyGroupComboboxItems
End Function

'[Verify Select Combobox has Items]
Public Function verifySelectComboboxItems(lstItems)
   bverifySelectComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CSO_TM_Home.lstSelect(), lstItems, "CSO Select")Then
           bverifySelectComboboxItems=false
       End If
   End If
   verifySelectComboboxItems=bverifySelectComboboxItems
End Function

'[Set TextBox Account No. on Home Page to]
Public Function setAccountNo(strAccountNo)
	bsetAccountNo=true
	Wait 1
	If Not Isnull (strAccountNo) Then
		CSO_TM_Home.txtAccountNo.Set(strAccountNo)
	End If
	If Err.Number<>0 Then
       bsetAccountNo=false
         LogMessage "WARN","Verification","Failed to Set Text Box : Account No" ,false
       Exit Function
   End If
   setAccountNo=bsetAccountNo
End Function

'[Set TextBox NRIC on Home Page to]
Public Function setNRIC(strNRIC)
	bsetNRIC=true
	Wait 1
	If Not Isnull (strNRIC) Then
		CSO_TM_Home.txtNRICPassport.Set(strNRIC)
	End If
	If Err.Number<>0 Then
       bsetNRIC=false
         LogMessage "WARN","Verification","Failed to Set Text Box : NRIC/Passport" ,false
       Exit Function
   End If
   setNRIC=bsetNRIC
End Function

'[Validate User and Date fields are disabled on CSO Home Page]
Public Function validateUserDateFieldDisable()
	bvalidateUserDateFieldDisable=true
	Wait 1
	strUserClassProp=CSO_TM_Home.lblUser_HomePage.GetROProperty("outerhtml")
	intDisable=Instr(strUserClassProp,("v-disabled"))
	If not intDisable=0 Then
		LogMessage "RSLT","Verification","User Field Section is disabled as expected", True
	Else
		LogMessage "RSLT","Verification","User Field Section is enable. Expected to be disable", False
		bvalidateUserDateFieldDisable=false
	End If

	strDateClassProp=CSO_TM_Home.lblDate_HomePage.GetROProperty("outerhtml")
	intDisableDate=Instr(strDateClassProp,("v-disabled"))
	If not intDisableDate=0 Then
		LogMessage "RSLT","Verification","Date Field Section is disabled as expected", True
	Else
		LogMessage "RSLT","Verification","Date Field Section is enable. Expected to be disable", False
		bvalidateUserDateFieldDisable=false
	End If	
	validateUserDateFieldDisable=bvalidateUserDateFieldDisable
End Function

'[Validate User and Date fields are enable on CSO Home Page]
Public Function validateUserDateFieldEnable()
	bvalidateUserDateFieldEnable=true
	Wait 1
	strUserClassProp=CSO_TM_Home.lblUser_HomePage.GetROProperty("outerhtml")
	intdisable=Instr(strUserClassProp,("v-disabled"))
	If intenable=0 Then
		LogMessage "RSLT","Verification","User Field Section is enable as expected", True
	Else
		LogMessage "RSLT","Verification","User Field Section is disable. Expected to be enable", False
		bvalidateUserDateFieldEnable=false
	End If

	strDateClassProp=CSO_TM_Home.lblDate_HomePage.GetROProperty("outerhtml")
	intDisableDate=Instr(strDateClassProp,("v-disabled"))
	If intDisableDate=0 Then
		LogMessage "RSLT","Verification","Date Field Section is enable as expected", True
	Else
		LogMessage "RSLT","Verification","Date Field Section is disable. Expected to be enable", False
		bvalidateUserDateFieldEnable=false
	End If	
	validateUserDateFieldEnable=bvalidateUserDateFieldEnable
End Function

'[Verify From Date Error Message displayed as]
Public Function verifyFromDateError(strFromDateError)
	bverifyFromDateError=true
	If not VerifyInnerText(CSO_TM_Home.lblFromDateError(), strFromDateError, "From Date Error") Then
        bverifyFromDateError=false
	End If
	verifyFromDateError=bverifyFromDateError
End Function

'[Verify To Date Error Message displayed as]
Public Function verifyToDateError(strToDateError)
	bverifyToDateError=true
	If not VerifyInnerText(CSO_TM_Home.lblToDateError(), strToDateError, "To Date Error") Then
        bverifyToDateError=false
	End If
	verifyToDateError=bverifyToDateError
End Function

'[Validate Pagination of CSO TM Home Page]
Public Function validatePagination()
   bvalidatePagination=true
		 intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent, "Created On")
		 If intRecordCount <=10 Then
			LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to 10", true
			bvalidatePagination=true
		  Else
			LogMessage "WARN","Verification","Number of records displayed per page is more than 10 record. Expected Count is less than or equal to 10, Actual "&intRecordCount, false
			bvalidatePagination=false
		  End If

		  If intRecordCount < 10 Then
				bNextPageExist =matchStr(lnkNext1().GetROProperty("outerhtml"),"v-disabled")
				If Not bNextPageExist Then
				LogMessage "WARN","Verification","Next link expected to be disable if record is less than 10. Currently it is enable.",false
				bvalidatePagination=false
				Else
				LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
				bvalidatePagination=true
				End If
			End If
			validatePagination=bvalidatePagination
End Function

'[Verify Pagination for the CSO Home Page Activities Log table displayed]
Public Function ValidatePagination_ActivitiesDetailstable()
 bValidatePagination_ActivitiesDetailstable = true
 bNextPageExist = True
	While bNextPageExist = True
	 intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent,"Card Number")	
	 iCheck = 10 
		If intRecordCount <=iCheck  Then
	     LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to "&iCheck, true   
	     bValidatePagination_ActivitiesDetailstable=true
		 If intRecordCount < iCheck Then
		   	bNextPageExist =matchStr(CSO_TM_Home.lnkNext.GetROProperty("class"),"enabled")
			If bNextPageExist Then
			   LogMessage "WARN","Verification","Next link expected to be disabled if record is less than "&iCheck&". Currently it is enabled.",false
			   bValidatePagination_ActivitiesDetailstable=false
			Else
			   LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
			End If
		  ElseIf intRecordCount = iCheck Then
			bNextPageExist = matchStr(CSO_TM_Home.lnkNext.GetROProperty("class"),"enabled")
			If bNextPageExist Then
			   CSO_TM_Home.lnkNext.Click
			End If
		  End If
		Else 
			LogMessage "RSLT","Verification","Number of records displayed per page not matched with expected. Expected Count is less than or equal to 12", false   
			bNextPageExist = False
		End If
   Wend
   ValidatePagination_ActivitiesDetailstable = bValidatePagination_ActivitiesDetailstable
End Function

'[Validate Account No field is disabled on CSO Home Page]
Public Function validateAccountNoFieldDisable()
	bvalidateAccountNoFieldDisable=true
	Wait 1
	strAccountNoClassProp=CSO_TM_Home.txtAccountNo.GetROProperty("outerhtml")
	intDisable=Instr(strAccountNoClassProp,("v-disabled"))
	If not intDisable=0 Then
		LogMessage "RSLT","Verification","Account No Field is disabled as expected", True
	Else
		LogMessage "RSLT","Verification","Account No Field is enable. Expected to be disable", False
		bvalidateAccountNoFieldDisable=false
	End If
	validateAccountNoFieldDisable=bvalidateAccountNoFieldDisable
End Function

'[Verify Info Warn Message on CSO TM Home Page]
Public Function verifyInfoWarnMessage(strInfoWarnMessage)
	bverifyInfoWarnMessage=true
	If Not VerifyInnerText(CSO_TM_Home.txt_InfoMsg(), strInfoWarnMessage, "Info Warn") Then
		bverifyInfoWarnMessage=false
	End If
	verifyInfoWarnMessage=bverifyInfoWarnMessage
End Function

'[Verify the status of SR in CSO Home Page]
Public Function verifySRStatus(strStatus)
	verifySRStatus = true
	'if the status is Open
	If strStatus = "Open" Then
		'select the radio button as Open
		CSO_TM_Home.rdOpenRadio.Click
		WaitForICallLoading
		intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent, "Created Date")
		If intRecordCount >= 5 Then
			loopCount = 5
		else
			loopCount = intRecordCount
		End If
		For it = 0 To loopCount - 1 Step 1
			'Click on the Type Column
			ReDim lstSRActivityData(1)
			Set objAllRows = getAllRows(CSO_TM_Home.tblSRActivityContent)
			strTableType = getCellTextFor(CSO_TM_Home.tblSRActivityHeader,objAllRows(it),it,"Type")
			strTableCreatedDate = getCellTextFor(CSO_TM_Home.tblSRActivityHeader,objAllRows(it),it,"Created Date")
			lstSRActivityData(0)="Type:"&strTableType
			lstSRActivityData(1)="Created Date:"&strTableCreatedDate
			Call clickType_SR_Activity_RecordsCS(lstSRActivityData)
			strActualSRNo = ServiceRequest.lblServiceRequestNo().GetRoProperty("innertext")
			If strActualSRNo <> "" Then
				'SR number exists
				LogMessage "RSLT","Verification","SR No exists; hence the SR is in Open Status", True
			Else
				LogMessage "WARN","Verification","SR No does not exist; hence the SR is not in Open Status", False
			End If
			
			'Close the tab
			closeTab_SR("View SR*")

		Next
	ElseIf strStatus="Failed" Then
		CSO_TM_Home.rdFailedRadio.Click
		WaitForICallLoading
		intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent, "Created Date")
		If intRecordCount >= 5 Then
			loopCount = 5
		else
			loopCount = intRecordCount
		End If
		For it = 0 To loopCount - 1 Step 1
			'Click on the Type Column
			ReDim lstSRActivityData(1)
			Set objAllRows = getAllRows(CSO_TM_Home.tblSRActivityContent)
			strTableType = getCellTextFor(CSO_TM_Home.tblSRActivityHeader,objAllRows(it),it,"Type")
			strTableCreatedDate = getCellTextFor(CSO_TM_Home.tblSRActivityHeader,objAllRows(it),it,"Created Date")
			lstSRActivityData(0)="Type:"&strTableType
			lstSRActivityData(1)="Created Date:"&strTableCreatedDate
			Call clickType_SR_Activity_RecordsCS(lstSRActivityData)
			strActualSRNo = ServiceRequest.lblServiceRequestNo().GetRoProperty("innertext")
			If strActualSRNo = "" Then
				'SR number exists
				LogMessage "RSLT","Verification","SR No does not exist; hence the SR is in Failed Status", True
			Else
				LogMessage "WARN","Verification","SR No exists; hence the SR is not in Failed Status", False
			End If
			
			'Close the tab
			closeAllTabExceptOverview

		Next
	End If
End Function

'[Verify Account Number or CIN replaced displayed in the search filter of CSO Home Page]
Public Function verifyAcctNoCIN_CSO(strStatus,strAccountNo,strCIN)
	verifyAcctNoCIN_CSO = true
	If strStatus = "Open" Then
		'Click on Open Radio button
		CSO_TM_Home.rdOpenRadio.Click
		WaitForICallLoading
	ElseIf strStatus = "Pending" Then
		'Click on Pending Radio button
		CSO_TM_Home.rdPendingRadio.Click
		WaitForICallLoading
	ElseIf strStatus = "Failed" Then
		'Click on Failed Radio button
		CSO_TM_Home.rdFailedRadio.Click
		WaitForICallLoading
	Else
		'Click on Pending Approval Radio button
		CSO_TM_Home.rdPendingApprovalRadio.Click
		WaitForICallLoading
	End If
	
	If strAccountNo <> "" Then
		'Click on the Account Number
		strTrimmedAccountNo = replace(strAccountNo,"-","")
		CSO_TM_Home.txtAccountNo.Set strTrimmedAccountNo
		Call clickButtonFilter
		intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent, "Created Date")
		If intRecordCount >= 5 Then
			loopCount = 5
		else
			loopCount = intRecordCount
		End If
		For it = 0 To loopCount - 1 Step 1
			'Click on the Type Column
			Set objAllRows = getAllRows(CSO_TM_Home.tblSRActivityContent)
			strActualAcctNo = getCellTextFor(CSO_TM_Home.tblSRActivityHeader,objAllRows(it),it,"Account / Card No.")
			If strActualAcctNo = strAccountNo Then
				'Matching
				LogMessage "RSLT","Verification","Account/Card No matching. Actual: " &strActualAcctNo& " Expected: " &strAccountNo, True
			else
				LogMessage "WARN","Verification","Account/Card No not matching. Actual: " &strActualAcctNo& " Expected: " &strAccountNo, False
				verifyAcctNoCIN_CSO = false
			End If
		Next
	End If
	
	If strCIN <> "" Then
		'Click on the Account Number
		CSO_TM_Home.txtNRICPassport.Set strCIN
		Call clickButtonFilter
		intRecordCount = getRecordsCountForColumn(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent, "Created Date")
		If intRecordCount >= 5 Then
			loopCount = 5
		else
			loopCount = intRecordCount
		End If
		For it = 0 To loopCount - 1 Step 1
			'Click on the Type Column
			Set objAllRows = getAllRows(CSO_TM_Home.tblSRActivityContent)
			strActualCIN = getCellTextFor(CSO_TM_Home.tblSRActivityHeader,objAllRows(it),it,"CIN")
			If strActualCIN = strCIN Then
				'Matching
				LogMessage "RSLT","Verification","CIN matching. Actual: " &strActualCIN& " Expected: " &strCIN, True
			else
				LogMessage "WARN","Verification","CIN not matching. Actual: " &strActualCIN& " Expected: " &strCIN, False
				verifyAcctNoCIN_CSO = false
			End If
		Next
	End If
	
End Function

'Added by Kalyan 12/07/2016

'[Verify Staff Combobox has Items]
Public Function verifyStafComboboxItems(lstItems)
   bverifyStafComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboList (lstItems, CSO_TM_Home.lblStaffDropdown())Then
           bverifyStafComboboxItems=false
       End If
   End If
   verifyStafComboboxItems=bverifyStafComboboxItems
End Function

'[Navigate to Home Page]
Public Function navigateHomePage()
   bDevPending=true
   Wait 1
   CSO_TM_Home.NavigateHmePage.click 
   If Err.Number<>0 Then
       navigateHomePage=false
            LogMessage "WARN","Verification","Failed to Click Navigate to Home Page Link" ,false
       Exit Function
   End If
   WaitForICallLoading
   navigateHomePage=true
End Function

'[Set the staff value on CSO TM home page]
Public Function selectstaff_CSO_TM(strStaff)
	bselectstaff_CSO_TM = true
	CSO_TM_Home.txtCSO.set strStaff
	selectstaff_CSO_TM = bselectstaff_CSO_TM
End Function

'******************Added by Kalyan as part of 1701 changes**************************

'[Click Type hyperlink for service request]
Public Function viewSRDetails(strType, strSubType)
          Dim bviewSRDetails:	viewSRDetails = True
		  getDateMnthYear(strSrSubmitTime)
          lstSRDetails = checknull("Type:"&strType&"|Sub Type:"&strSubType&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
          
          bviewSRDetails=selectTableLink(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent,lstSRDetails,"ViewSRDetails","Type",True,CSO_TM_Home.lnkNext ,CSO_TM_Home.lnkNext1,CSO_TM_Home.lnkPrevious)          				 
          WaitForICallLoading
		  viewSRDetails=bviewSRDetails
End Function

'[Set Pending Approval Channel Combobox as]
Public Function selectPenApprvlChnlComboBox(strChannel)
bselectPenApprvlChnlComboBox=true
If Not IsNull(strChannel) Then
NewSR.txtPendAprvChnl().set strChannel
End If
waitForIcallLoading
selectPenApprvlChnlComboBox=bselectPenApprvlChnlComboBox
End Function

'[Verify Pending Approval Channel Enabled]
Public Function PenApprvlChnlEnabled()
bPenApprvlChnlEnabled=true
intPendAppChnl =InStr(NewSR.txtPendAprvChnl.GetROProperty("visible"),"True")
	If not intPendAppChnl=0 Then
		LogMessage "RSLT","Verification","Channel is displayed as expected.",True
		bPenApprvlChnlEnabled=true
	Else
		LogMessage "WARN","Verifiation","Channel not displayed. Expected to be displayed.",false
		bPenApprvlChnlEnabled=false
	End If
PenApprvlChnlEnabled=bPenApprvlChnlEnabled
End Function

'[Verify Pending Approval Channel Disabled]
Public Function PenApprvlChnlDisabled()
bPenApprvlChnlDisabled=true
strSttaus=NewSR.txtPendAprvChnl.GetROProperty("visible")
intPendAppChnl =InStr(NewSR.txtPendAprvChnl.GetROProperty("visible"),"False")
	If not intPendAppChnl=0 Then
		LogMessage "RSLT","Verification","Channel not displayed as expected.",True
		bPenApprvlChnlDisabled=true
	Else
		LogMessage "WARN","Verifiation","Channel displayed. Expected to be not displayed.",false
		bPenApprvlChnlDisabled=false
	End If
PenApprvlChnlDisabled=bPenApprvlChnlDisabled
End Function


'[Verify Staff Combobox default field value]
Public Function verifyStafComboboxFieldVal(lstItem)
   bverifyStafComboboxFieldVal=true
  If Not IsNull(lstItem) Then
If Not verifyFieldValue (CSO_TM_Home.txtStaff(), lstItem, "DefaultFieldValue")Then
bverifyStafComboboxFieldVal=false
End If
End If
   verifyStafComboboxFieldVal=bverifyStafComboboxFieldVal
End Function

'[Set Staff Combobox field value as]
Public Function selectStafComboboxFieldVal(strItem)
   bselectStafComboboxFieldVal=true
 If Not IsNull(strItem) Then
CSO_TM_Home.txtStaff().set strItem
bselectStafComboboxFieldVal=selectItem_Combobox (CSO_TM_Home.lstCSO(), strItem)
End If
   selectStafComboboxFieldVal=bselectStafComboboxFieldVal
End Function

'[Verify Pending Approval Channel Combobox default field value]
Public Function selectPenApprvlChnlDefltVal(strChannel)
bselectPenApprvlChnlDefltVal=true
If Not IsNull(strChannel) Then
If Not verifyFieldValue (NewSR.txtPendAprvChnl(), strChannel, "DefaultFieldValue")Then
bselectPenApprvlChnlDefltVal=false
End If
End If
waitForIcallLoading
selectPenApprvlChnlDefltVal=bselectPenApprvlChnlDefltVal
End Function

'[Verify Show Combobox default field value]
Public Function selectShowComboboxDefltVal(strShow)
bselectShowComboboxDefltVal=true
If Not IsNull(strShow) Then
If Not verifyFieldValue (CSO_TM_Home.txtShow(), strShow, "DefaultFieldValue")Then
bselectShowComboboxDefltVal=false
End If
End If
waitForIcallLoading
selectShowComboboxDefltVal=bselectShowComboboxDefltVal
End Function

'Added by Kalyan BULK SR dated 13022017

'[Set TextBox Blk SR Assigned To on Home Page]
Public Function setBlkSRAsgndTo(strtxtBlkAsgndTo)
	bsetBlkSRAsgndTo=true
	Wait 1
	If Not Isnull (strtxtBlkAsgndTo) Then
		CSO_TM_Home.txtAssignTo.Set(strtxtBlkAsgndTo)
	End If
	If Err.Number<>0 Then
       bsetBlkSRAsgndTo=false
         LogMessage "WARN","Verification","Failed to Set Text Box : txtAssignTo" ,false
       Exit Function
   End If
   setBlkSRAsgndTo=bsetBlkSRAsgndTo
End Function

'[Verify Max BULK SR Err Message on CSO TM Home Page]
Public Function verifyMaxBlkSRErrMsg(strMaxBlkSrErrMsg)
	bverifyMaxBlkSRErrMsg=true
	If Not VerifyInnerText(CSO_TM_Home.lblErrBulkMaxRecSR(), strMaxBlkSrErrMsg, "MaxBlkSR") Then
		bverifyMaxBlkSRErrMsg=false
	End If
	verifyMaxBlkSRErrMsg=bverifyMaxBlkSRErrMsg
End Function

'[Verify Min BULK SR Err Message on CSO TM Home Page]
Public Function verifyMinBlkSRErrMsg(strMinBlkSrErrMsg)
	bverifyMinBlkSRErrMsg=true
	If Not VerifyInnerText(CSO_TM_Home.lblErrBulMinRecSR(), strMinBlkSrErrMsg, "MinBlkSR") Then
		bverifyMinBlkSRErrMsg=false
	End If
	verifyMinBlkSRErrMsg=bverifyMinBlkSRErrMsg
End Function

'[Verify BULK SR Success Message on CSO TM Home Page]
Public Function verifyBlkSRSuccessMsg(strBlkSrSuccssMsg)
	bverifyBlkSRSuccessMsg=true
	If Not VerifyInnerText(CSO_TM_Home.lblBlkSrSuccssMsg(), strBlkSrSuccssMsg, "BlkSRSuccssMsg") Then
		bverifyBlkSRSuccessMsg=false
	End If
	verifyBlkSRSuccessMsg=bverifyBlkSRSuccessMsg
End Function

'[Click Button Blk SR Assigned To on Home Page]
Public Function clickBtnBlkSRAsgndTo()
   bDevPending=true
   Wait 1
   CSO_TM_Home.btnAssignTo.click
   If Err.Number<>0 Then
       clickBtnBlkSRAsgndTo=false
            LogMessage "WARN","Verification","Failed to Click Button : BlkSRAsgndTo" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickBtnBlkSRAsgndTo=true
End Function

'[Click Button Blk SR Cancel on Home Page]
Public Function clickBtnBlkSRCancel()
   bDevPending=true
   Wait 1
   CSO_TM_Home.btnCancelBlkSr.click 
   If Err.Number<>0 Then
       clickBtnBlkSRCancel=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickBtnBlkSRCancel=true
End Function

'[Click Button Blk SR Submit on Home Page]
Public Function clickBtnBlkSRSubmit()
   bDevPending=true
   Wait 1
   CSO_TM_Home.btnSubmitBlkSR.click
   If Err.Number<>0 Then
       clickBtnBlkSRSubmit=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickBtnBlkSRSubmit=true
End Function

'[Verify row Data in BLK SR Table SR Assigned To Activity Records]
Public Function verifytblBlkSR_Activity(arrRowDataList)
   bDevPending=false
   verifytblBlkSR_Activity=verifyTableContentList(CSO_TM_Home.tblBlkSrHeader,CSO_TM_Home.tblBlkSrContent,arrRowDataList,"BlkSR_Activity_Records" ,False,CSO_TM_Home.lnkBlkNext ,CSO_TM_Home.lnkBlkNext1,CSO_TM_Home.lnkBlkPrevious)
End Function

'[Click Check Box Bulk SR Select All on Home Page]
Public Function clickChkBoxBlkSrSlctAll()
   bDevPending=true
   CSO_TM_Home.lblOpnSRSlctAll.click
   If Err.Number<>0 Then
       clickChkBoxBlkSrSlctAll=false
            LogMessage "WARN","Verification","Failed to Click Blk SR Selct All : CheckBox" ,false
       Exit Function
   End If
   clickChkBoxBlkSrSlctAll=true
End Function


'Bulk SR functionality added by kalyan 07022017
'[Check Bulk SR checkboxes]
Public Function CheckChkBoxBulkSRUpdateAssignTo(lstlstBulkOpnSrData)
	bCheckChkBoxBulkSRUpdateAssignTo = True
	If Not IsNULL (lstlstBulkOpnSrData) Then
		bCheckChkBoxBulkSRUpdateAssignTo = selectTableCheckBox_MultipleRow(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent,lstlstBulkOpnSrData,"BULKSRTable","",TRUE,CSO_TM_Home.lnkNext,CSO_TM_Home.lnkNext1,CSO_TM_Home.lnkPrevious)
	End If
	
End Function

'[Click button close Bulk SR on Home Page]
Public Function clickBtnCloseBlkSr()
   bDevPending=true
   CSO_TM_Home.btnClose.click
   If Err.Number<>0 Then
       clickBtnCloseBlkSr=false
            LogMessage "WARN","Verification","Failed to Click button : Close" ,false
       Exit Function
   End If
   clickBtnCloseBlkSr=true
End Function

'[Select Pending Approval SR from TM home Page LISA]
Public Function selectPendingSR_TMHome_LISA(lstPendingApproval)
 	bDevPending=false
   Dim bselectPendingSR_TMHome_LISA:bselectPendingSR_TMHome_LISA=true
	bselectPendingSR_TMHome_LISA=selectTableLink(CSO_TM_Home.tblSRActivityHeader,CSO_TM_Home.tblSRActivityContent,lstPendingApproval,"ViewSRDetails","Type",True,CSO_TM_Home.lnkNext ,CSO_TM_Home.lnkNext1,CSO_TM_Home.lnkPrevious)          				 
    WaitForICallLoading
	selectPendingSR_TMHome_LISA=bselectPendingSR_TMHome_LISA
End Function
