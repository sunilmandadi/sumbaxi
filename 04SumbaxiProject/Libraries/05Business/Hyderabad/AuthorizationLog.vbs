'[Verify Default Transaction Date]
Public Function verifyDefaultTransDate()
  verifyDefaultTransDate = true
  strAuthLogDateFrom = AuthorizationLog.txtTransactionDateFrom.GetROProperty("value")
  strAuthLogDateTo = AuthorizationLog.txtTransactionLogDateTo.GetROProperty("value")
  DaysRange = DateDiff("d", strAuthLogDateFrom, strAuthLogDateTo)
  If DaysRange = 6  Then
     LogMessage "RSLT","Verification","From Date:"+ trim(strAuthLogDateFrom) +"and To Dates:"+ trim(strAuthLogDateTo) +" are displayed diff by 7 days default range",True
     verifyDefaultTransDate = True 
  Else 
     LogMessage "WARN", "Verification","From Date:"+ trim(strAuthLogDateFrom) +"and To Dates:"+ trim(strAuthLogDateTo) +" are not displayed diff by 7 days default range", False
     verifyDefaultTransDate = False 
  End If
End Function

'[Verify Merchant Name Default By]
Public Function verifyMerchantName(strMerchantName)
   verifyMerchantName = verifyInnerText(AuthorizationLog.lblMerchantName(),strMerchantName, "Merchant Name")
End Function

'[Verify Authorization Log screen Table list data]
Public Function verifyAuthlogTransTableList(lstlstAuthlogTransTableList)
   verifyAuthlogTransTableList = verifyTableContentList(AuthorizationLog.lblauthLogDatatableheader, AuthorizationLog.lblauthLogDataTableContent,lstlstAuthlogTransTableList,"Authorization log Transaction Details",false,NULL,NULL,NULL)
   WaitForICallLoading
End Function

'[Verify the shortcut buttons on AuthorizationLog page]
Public Function verifyShortCutButtonAuthLogPage()
	verifyShortCutButtonAuthLogPage = true
	If (AuthorizationLog.lblchargeDisputeButton.Exist(1)) Then
		LogMessage "RSLT","Verification","Charge Dispute button is available for AuthorizationLog Page." ,True
	Else
	    LogMessage "WARN","Verification","Charge Dispute button is not available for AuthorizationLog Page. Expected to be Enabled." ,false
	    verifyShortCutButtonAuthLogPage = False
	End If

    If (AuthorizationLog.lblMoreServicesButton.Exist(1)) Then
		LogMessage "RSLT","Verification","More Services Button is available for AuthorizationLog Page." ,True		
	Else
	   LogMessage "WARN","Verification","More Services button is not available for AuthorizationLog Page. Expected to be Enabled." ,false
	   verifyShortCutButtonAuthLogPage = False
	End If
	WaitForICallLoading
End Function

'[Set Transaction From Date]
Public Function SetAugTransFromDate(strFromDate)
   SetAugTransFromDate = True
   AuthorizationLog.txtTransactionDateFrom.Set strFromDate
   If Err.Number<>0 Then
       SetAugTransFromDate=false
       LogMessage "WARN","Verifscation","Failed to Set From Date : AuthTransPage" ,false
       Exit Function
   Else
   	LogMessage "RSLT","Verification","Successfuly Set the From Date : AuthTransPage", true  
   End If
   WaitForICallLoading
End Function

'[Set Transaction To Date]
Public Function SetAugTransToDate(strToDate)
   SetAugTransToDate = True
   AuthorizationLog.txtTransactionLogDateTo.Set strToDate
   If Err.Number<>0 Then
       SetAugTransToDate=false
       LogMessage "WARN","Verification","Failed to Set To Date : AuthTransPage" ,false
       Exit Function
   Else
   	LogMessage "RSLT","Verification","Successfuly Set the To Date : AuthTransPage", true   
   End If
   'WaitForICallLoading
End Function

'[Click on the Go Button from AuthPage]
Public Function ClickOnGoButtonFromAuthPage()    
	ClickOnGoButtonFromAuthPage = true
	'WaitForICallLoading
	AuthorizationLog.btnGoButtonAuthPage.Click
	If Err.Number<>0 Then
       ClickOnGoButtonFromAuthPage = false
       LogMessage "WARN","Verification","Failed to Click Button : GO" ,false
    else
     LogMessage "RSLT","Verification","Successfuly Clicked on Go Button", true   
   	End If
	WaitForICallLoading
End Function

'[Verify No Records Found for selected Range]
Public Function verifyNoRecordsFoundMsg(strNoRecordsFoundMsg)
   verifyNoRecordsFoundMsg = true
   If Not IsNull(strNoRecordsFoundMsg) Then
       If Not VerifyInnerText (AuthorizationLog.lblauthLogNoAuthLogerror(), strNoRecordsFoundMsg, "No Records Found for selected Range")Then
           verifyNoRecordsFoundMsg=false
       End If
   End If
End Function

'[Verify Pagination for the Authorizational Log table displayed]
Public Function ValidatePagination_AuthDetailstable()
 bValidatePagination_AuthDetailstable=true
 bNextPageExist = True
	While bNextPageExist = True
	 intRecordCount = getRecordsCountForColumn(AuthorizationLog.lblauthLogDatatableheader,AuthorizationLog.lblauthLogDataTableContent,"Card Number")	
	 iCheck = 12 
		If intRecordCount <=iCheck  Then
	     LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to "&iCheck, true   
	     bValidatePagination_AuthDetailstable=true
		 If intRecordCount < iCheck Then
		   	bNextPageExist =matchStr(AuthorizationLog.lnkAuthLogDataTableNext.GetROProperty("class"),"enabled")
			If bNextPageExist Then
			   LogMessage "WARN","Verification","Next link expected to be disabled if record is less than "&iCheck&". Currently it is enabled.",false
			   bValidatePagination_AuthDetailstable=false
			Else
			   LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
			End If
		  ElseIf intRecordCount = iCheck Then
			bNextPageExist = matchStr(AuthorizationLog.lnkAuthLogDataTableNext.GetROProperty("class"),"enabled")
			If bNextPageExist Then
			   AuthorizationLog.lnkAuthLogDataTableNext.Click
			End If
		  End If
		Else 
			LogMessage "RSLT","Verification","Number of records displayed per page not matched with expected. Expected Count is less than or equal to 12", false   
			bNextPageExist = False
		End If
   Wend
   ValidatePagination_AuthDetailstable = bValidatePagination_AuthDetailstable
End Function
