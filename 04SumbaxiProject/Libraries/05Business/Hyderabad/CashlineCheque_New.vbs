'[Verify the row data for Cashline Cheque Info table]
Public Function verifyrowdata_CashlineCheque(arrRowDataList)
	bverifyrowdata_CashlineCheque = true
	verifyrowdata_CashlineCheque = verifyTableContentList(bcCashline_Cheque.tblChequeInfoHeader,bcCashline_Cheque.tblChequeInfoContent,arrRowDataList,"Cashline Cheque Info",false,null ,null,null)
	verifyrowdata_CashlineCheque = bverifyrowdata_CashlineCheque
End Function

'[Click on View Cheque info from Cheque Info Table]
Public Function clickView_CashlineCheque(lstRowData)
	bclickView_CashlineCheque= true
	clickView_CashlineCheque = selectTableLink(bcCashline_Cheque.tblChequeInfoHeader,bcCashline_Cheque.tblChequeInfoContent,lstRowData,"Cashline Cheque","Cheque Status",false,null,null,null)	
	clickView_CashlineCheque = bclickView_CashlineCheque	
End Function

'[Verify total records and pagination for Cashline Cheque]
Public Function ValidatePagination_CashlineChequeInfo()
 bValidatePagination_CashlineChequeInfo=true
 bNextPageExist = True
	While bNextPageExist = True
	 intRecordCount = getRecordsCountForColumn(bcCashline_Cheque.tblChequeInfoHeader,bcCashline_Cheque.tblChequeInfoContent,"Tape No.")	
	 iCheck = 10 
		If intRecordCount <=iCheck  Then
		     LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to "&iCheck, true   
		     ValidatePagination_CashlineChequeInfo=bValidatePagination_CashlineChequeInfo
			 If intRecordCount < iCheck Then
			   	bNextPageExist =matchStr(bcCashline_Cheque.lnkNext.GetROProperty("class"),"enabled")
				If bNextPageExist Then
				LogMessage "RSLT","Verification","Next link expected to be disabled if record is less than "&iCheck&". Currently it is enabled.",false
				bvalidatePagination=false
				Else
				LogMessage "RSLT","Verification","Next link is disabled as per expectation.",true
				End If
			ElseIf intRecordCount = iCheck Then
				bNextPageExist = matchStr(bcCashline_Cheque.lnkNext.GetROProperty("class"),"enabled")
				If bNextPageExist Then
					bcCashline_Cheque.lnkNext.Click
				End If
			End If
		Else 
			LogMessage "RSLT","Verification","Number of records displayed per page not matched with expected. Expected Count is less than or equal to 10", false   
			bNextPageExist = False
		End If
   Wend
End Function

'[Verify the Cheque Status popup exist on click of View]
Public Function verifyChequeStatusExist(bExist)
   bDevPending=false
   bActualExist=bcCashline_Cheque.PopupChequeStatus.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Cheque Info Exists As Expected" ,true
       verifyChequeStatusExist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Cheque Info does not Exists As Expected" ,true
       verifyChequeStatusExist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Cheque Info does not Exists As Expected" ,False
       verifyChequeStatusExist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Cheque Info Still Exists" ,False
       verifyChequeStatusExist=False
   End If
End Function

'[Verify the fields of Cheque Status popup]
Public Function verifyField_ChequeStatusPopup(strIssuedDate,strPaidCheque,strTotalCheque)
	
	'Getting the values for the fields from the I.Serve
	strIserveIssuedDate = bcCashline_Cheque.lblIssuedDate.GetROProperty("innertext")
	strIservePaidCheque = bcCashline_Cheque.lblPaidCheque.GetROProperty("innertext")
	strIserveTotalCheque =bcCashline_Cheque.lblTotalNoofCheques.GetROProperty("innertext")
	bDevPending=false
   bverifyField_ChequeStatusPopup=true
   
   If strIssuedDate = strIserveIssuedDate Then
   	    LogMessage "RSLT","Verification","The Iserve Issued Date is as expected: "&strIssuedDate&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Issued Date is not as expected: "&strIssuedDate&"",False
   End If
	
	If strPaidCheque = strIservePaidCheque Then
   	    LogMessage "RSLT","Verification","The Iserve Paid Cheque is as expected: "&strPaidCheque&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Paid Cheque is not as expected: "&strPaidCheque&"",False
   End If
   
	If strTotalCheque = strIserveTotalCheque Then
   	    LogMessage "RSLT","Verification","The Iserve Total No. of Cheque is as expected: "&strTotalCheque&"",True
		Else
	  	LogMessage "RSLT","Verification","The Iserve Total No. of Cheque is not as expected: "&strTotalCheque&"",False
   End If
	verifyField_ChequeStatusPopup = bverifyField_ChequeStatusPopup
End Function

'[Verify the row data for cheque status table]
Public Function verifyrowdata_ChequeStatus(arrRowDataList)
	bverifyrowdata_ChequeStatus = true
	verifyrowdata_ChequeStatus = verifyTableContentList(bcCashline_Cheque.tblChequeStatusHeader,bcCashline_Cheque.tblChequeStatusTable,arrRowDataList,"Cheque Status Info",false,null ,null,null)
	verifyrowdata_ChequeStatus = bverifyrowdata_ChequeStatus
End Function

'[Click the Cheque no from the cheque status table]
Public Function clickCheque_CheqStatus(arrRowDataList)
	bclickCheque_CheqStatus= true
	clickCheque_CheqStatus = selectTableLink(bcCashline_Cheque.tblChequeStatusHeader,bcCashline_Cheque.tblChequeStatusTable,arrRowDataList,"Cheque Status","Cheque No.",false,null,null,null)	
	clickCheque_CheqStatus = bclickCheque_CheqStatus
End Function

'[Verify the row data for Cheque Transaction Details]
Public Function verifyrowdata_ChequeTransDetails(arrRowDataList)
	bverifyrowdata_ChequeTransDetails = true
	verifyrowdata_ChequeTransDetails = verifyTableContentList(bcCashline_Cheque.tblChequeTransactionHeader,bcCashline_Cheque.tblChequeTransactionContent,arrRowDataList,"Cheque Trans Details",false,null ,null,null)
	verifyrowdata_ChequeTransDetails = bverifyrowdata_ChequeTransDetails	
End Function

'[Click on Ok button for Cheque Status pop up]
Public Function clickButtonOK_ChequeStatus()
   bDevPending=true
   bcCashline_Cheque.btnOK.click
   If Err.Number<>0 Then
       clickButtonOK_ChequeStatus=false
            LogMessage "RSLT","Verification","Failed to Click Button : OK" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonOK_ChequeStatus=true
End Function
