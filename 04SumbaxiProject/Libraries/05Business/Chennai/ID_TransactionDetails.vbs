'[Click Expand icon for Accordion displayed in CASA Page]
Public Function ClickExpandIcon_CASA(strPage,strAccordion)
	bVerify = false
	if Not IsNull(strPage) Then
		if Not IsNull(strAccordion) Then
			Select Case strPage
				Case "Savings Account"
				Set objAccordionGrp = coTransactionDetails_Page.objSAAccordionGrp
				Case "Current Account"
				Set objAccordionGrp = coTransactionDetails_Page.objCAAccordionGrp
			End Select
			bVerify = ExpandSingleAccordion(objAccordionGrp,strAccordion)
		End If
	End If
	ClickExpandIcon_CASA=bVerify
End Function

'[Click Collapse icon for Accordion displayed in CASA Page]
Public Function ClickCollapseIcon_CASA(strPage,strAccordion)
	bVerify = false
	if Not IsNull(strPage) Then
		if Not IsNull(strAccordion) Then
			Select Case strPage
				Case "Savings Account"
				Set objAccordionGrp = coTransactionDetails_Page.objSAAccordionGrp
				Case "Current Account"
				Set objAccordionGrp = coTransactionDetails_Page.objCAAccordionGrp
			End Select
			bVerify = CollapseSingleAccordion(objAccordionGrp,strAccordion)
		End If
	End If
	ClickCollapseIcon_CASA = bVerify
End Function

'[Click Expand icon for Accordion displayed in CCCL Page]
Public Function ClickExpandIcon_CCCL(strAccordion)
	if Trim(Ucase(strAccordion)) = "MEMO" Then
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 200")
	End if
	ClickExpandIcon_CCCL=ExpandSingleAccordion(coTransactionDetails_Page.objCCCLAccordionGrp,strAccordion)
	WaitForIServeLoading
End Function

'[Click Collapse icon for Accordion displayed in CCCL Page]
Public Function ClickCollapseIcon_CCCL(strAccordion)
	ClickCollapseIcon_CCCL = CollapseSingleAccordion(coTransactionDetails_Page.objCCCLAccordionGrp,strAccordion)
	WaitForIServeLoading
End Function

'[Verify default From and To date displayed in CASA Page]
Public Function verifyDefaultDateRange_CASA(strDateRange,StrToDate)	
	verifyDefaultDateRange_CASA = VerifyDateRange(coTransactionDetails_Page.txtTranDtlsFrom,coTransactionDetails_Page.txtTranDtlsTo,strDateRange,StrToDate)
End Function

'[Select From Date using Date Picker in CASA]
Public Function SelectFromDate_CASA(strFromDate)
bverifyDate = True
gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 200")
WaitForIServeLoading
If Not IsNull(strFromDate) Then

	If Trim(strFromDate) = "TODAY" Then
	   strFromDate = Date()
	End If
	
	SelectFromDate_CASA =  SelectDateFromIDCalendar(coTransactionDetails_Page.txtTranDtlsFrom,strFromDate)
	strExpFromDate = Right("0" & Datepart("d",strFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strFromDate),2))&" " & Year(strFromDate)
	
	If SelectFromDate_CASA Then
	
		strActFromDate = coTransactionDetails_Page.txtTranDtlsFrom.GetROProperty("value")
		strActFromDate = Right("0" & Datepart("d",strActFromDate),2) &" "& MonthName(Right("0" & Datepart("m",strActFromDate),2))&" " & Year(strActFromDate)
		
		If Trim(strActFromDate) = Trim(strExpFromDate) Then
		   LogMessage "RSLT","Verification","Selected date "&strFromDate&" in From date text box is displayed as expected", True
		   bverifyDate = True 
		Else
		   bverifyDate = False 
		End If	
		
	End If

End If

SelectFromDate_CASA = bverifyDate
End Function

'[Select TO Date using Date Picker in CASA]
Public Function SelectTODate_CASA(strTODate)
bverifyDate = True 
WaitForIServeLoading
If Not IsNull(strTODate) Then

	If Trim(strTODate) = "TODAY" Then
	   strTODate = Date()
	End If
	
	SelectTODate_CASA =  SelectDateFromIDCalendar(coTransactionDetails_Page.txtTranDtlsTo,strTODate)
	StrExpToDate = Right("0" & Datepart("d",strTODate),2) &" "& MonthName(Right("0" & Datepart("m",strTODate),2))&" " & Year(strTODate)
	
	If SelectTODate_CASA Then
	
		strActTODate = coTransactionDetails_Page.txtTranDtlsTo.GetROProperty("value")
		strActTODate = Right("0" & Datepart("d",strActTODate),2) &" "& MonthName(Right("0" & Datepart("m",strActTODate),2))&" " & Year(strActTODate)
		
		If Trim(strActTODate) = Trim(StrExpToDate) Then
		   LogMessage "RSLT","Verification","Selected date "&strTODate&" in TO date text box is displayed as expected", True
		   bverifyDate = True 
		Else
		  bverifyDate = False 
		End If	
		
	End If

End IF

SelectTODate_CASA = bverifyDate
End Function

'[Expand Accordion using Refresh Icon in CASA]
Public Function ClickRefreshIcon_CASA(strPage,strAccordion)
	bVerify = false
	if Not IsNull(strPage) Then
		if Not IsNull(strAccordion) Then
			Select Case strPage
				Case "Savings Account"
				Set objAccordionGrp = coTransactionDetails_Page.objSAAccordionGrp
				Case "Current Account"
				Set objAccordionGrp = coTransactionDetails_Page.objCAAccordionGrp
			End Select
			bVerify = VerifyAccordionRefresh(objAccordionGrp,strAccordion)
			WaitForIServeLoading
		End If
	End If
	ClickRefreshIcon_CASA = bVerify
End Function

'[Expand Accordion using Refresh Icon in CCCL]
Public Function ClickRefreshIcon_CCCL(strAccordion)
	ClickRefreshIcon_CCCL = VerifyAccordionRefresh(coTransactionDetails_Page.objCCCLAccordionGrp,strAccordion)
	WaitForIServeLoading
End Function

'[Verify Table Row Data in CASA]
Public Function VerifyTableRowData_CASA(lstlstRARowData)
	VerifyTableRowData_CASA = VerifyTableSingleRowData(coTransactionDetails_Page.tblTranDetailsHdr(),coTransactionDetails_Page.tblTranDetailsBody(),lstlstRARowData,"CASA Page")	
End Function

'[Verify Table Row Data in CCCL]
Public Function VerifyTableRowData_CCCL(lstlstRARowData)
	VerifyTableRowData_CCCL = VerifyTableSingleRowData(coTransactionDetails_Page.tblCCTranDetailsHdr(),coTransactionDetails_Page.tblCCTranDetailsBody(),lstlstRARowData,"CCCL Page")	
End Function

'[Click on Submit Button in Transaction Details]
Public Function clickButtonSubmit_TranDtl()
  coTransactionDetails_Page.btnSubmit.click 
  If Err.Number <> 0 Then
      clickButtonSubmit_TranDtl = False
      LogMessage "WARN","Verification","Failed to Click Button : Submit", False
      Exit Function
  End If
  WaitForIServeLoading
  clickButtonSubmit_TranDtl = True
End Function

'[Click on Cycle to Date Button in Transaction Details]
Public Function clickButtonCycletoDt_TranDtl()
'gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
  WaitForIServeLoading
  coTransactionDetails_Page.btnCycleToDate.click 
  If Err.Number <> 0 Then
      clickButtonCycletoDt_TranDtl = False
      LogMessage "WARN","Verification","Failed to Click Button : Submit", False
      Exit Function
  End If
  WaitForIServeLoading
  clickButtonCycletoDt_TranDtl = True
End Function

'[Click on Pending Button in Transaction Details]
Public Function clickButtonPending_TranDtl()
  coTransactionDetails_Page.btnPending.click 
  If Err.Number <> 0 Then
      clickButtonPending_TranDtl = False
      LogMessage "WARN","Verification","Failed to Click Button : Pending", False
      Exit Function
  End If
  WaitForIServeLoading
  clickButtonPending_TranDtl = True
End Function

'[Click on Warehoused Button in Transaction Details]
Public Function clickButtonWare_TranDtl()
  coTransactionDetails_Page.btnWarehoused.click 
  If Err.Number <> 0 Then
      clickButtonWare_TranDtl = False
      LogMessage "WARN","Verification","Failed to Click Button : Warehoused", False
      Exit Function
  End If
  WaitForIServeLoading
  clickButtonWare_TranDtl = True
End Function

'[Click on Disputed Button in Transaction Details]
Public Function clickButtonDisputed_TranDtl()
  coTransactionDetails_Page.btnDisputed.click 
  If Err.Number <> 0 Then
      clickButtonDisputed_TranDtl = False
      LogMessage "WARN","Verification","Failed to Click Button : Disputed", False
      Exit Function
  End If
  WaitForIServeLoading
  clickButtonDisputed_TranDtl = True
End Function

'[Verify Inline error message displayed in CASA]
Public Function VerifyInlineErrorMsg_CASA(strErrorMsg)
bverifyInlineErrorMsg = True
If not VerifyInnerText(coTransactionDetails_Page.lblinlineMsg(), strErrorMsg, "Inline Date Error") Then
   bverifyInlineErrorMsg = False
End If
VerifyInlineErrorMsg_CASA = bverifyInlineErrorMsg
End Function

'[Verify No Data available msg when enter non available Transaction Details value in CASA]
Public Function VerifyNoDatainTable_CASA(strTranDtls,strMsg)
	bVerify=True
	bVerify1=True
	If Not IsNull(strTranDtls) Then
		If Not SetValue(coTransactionDetails_Page.txtTranDetails(),strTranDtls,"Transaction Details")Then
				bVerify = False
		End If
	End If
	WaitForIServeLoading
	If Not IsNull(strMsg) Then
		If Not verifyInnerText(coTransactionDetails_Page.lblNoData(),strMsg,"Transaction Details Table") Then
			bVerify1=False
		End If
	End If
	
	If bVerify And bVerify1 Then
        VerifyNoDatainTable_CASA = True
    Else
        VerifyNoDatainTable_CASA = False
    End If
	
End Function

'[Verify No Data available msg when enter non available Transaction Date value in CCCL]
Public Function VerifyNoDatainTable_CCCL(strTranDate,strMsg)
	bVerify=True
	bVerify1=True
	If Not IsNull(strTranDate) Then
		If Not SetValue(coTransactionDetails_Page.txtTransactionDate(),strTranDate,"Transaction Date")Then
				bVerify = False
		End If
	End If
	WaitForIServeLoading
	If Not IsNull(strMsg) Then
		If Not verifyInnerText(coTransactionDetails_Page.lblNoData(),strMsg,"Transaction Details Table") Then
			bVerify1=False
		End If
	End If
	
	If bVerify And bVerify1 Then
        VerifyNoDatainTable_CCCL = True
    Else
        VerifyNoDatainTable_CCCL = False
    End If
	
End Function

'[Verify Transaction Details value has been entered in Transaction Details Accordion]
Public Function VerifyTableRowDataAppType_TranDetails(strTranDtls)
	bVerify=True
	If Not IsNull(strTranDtls) Then
		If Not SetValue(coTransactionDetails_Page.txtTranDetails(),strTranDtls,"Transaction Details")Then
				bVerify = False
		End If
	End If
        VerifyTableRowDataAppType_TranDetails = bVerify
End Function

'[Verify Transaction Date value has been entered in Transaction Details Accordion]
Public Function VerifyTableRowDataAppType_TranDate(strTranDate)
	bVerify=True
	If Not IsNull(strTranDate) Then
		If Not SetValue(coTransactionDetails_Page.txtTransactionDate(),strTranDate,"Transaction Date")Then
				bVerify = False
		End If
	End If
        VerifyTableRowDataAppType_TranDate = bVerify
End Function

'[Verify Infowarn message displayed in CASA OR CCCL]
Public Function VerifyInfowan_CASA(strInfoMsgtext)
	VerifyInfowan_CASA = VerifyInfowarntext(coTransactionDetails_Page.lblInfowarn,strInfoMsgtext)
End Function

'[Verify Record count displayed in Transaction Details]
Public Function VerifyIARecordCount_TranDetails(strNum)
bVerifyRecordCount = False

strDisplayedMsgtext = coTransactionDetails_Page.lblNumResults.GetRoProperty("innertext")
strMsgText = strNum+" Results Found"

	If Instr(1,strDisplayedMsgtext,strMsgText,1) > 0 Then 
	   LogMessage "WARN","Verification","Record Count text message is displayed as expected", True
	   bVerifyRecordCount = True
	Else 
	   LogMessage "WARN","Verification","Record Count text message is not displayed as expected", False
	   bVerifyRecordCount = False	
	End IF 
	
VerifyIARecordCount_TranDetails = bVerifyRecordCount
End Function

'[Verify Pagination for Transactions Details Accordian]
Public Function Verifypagination_TranDetails(strPage)
bVerify = False
If Not IsNull(strPage) Then
			Select Case strPage
				Case "CASA"
					Wait(2)
					Set tblObj = coTransactionDetails_Page.tblTranDetails()
					strcolumnname = "TRANSACTION DATE"
					strRows = 10
					gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
					Wait(2)
				Case "CCCL"
					Wait (2)
					Set tblObj = coTransactionDetails_Page.tblCCTranDetails()
					strcolumnname = "TRANSACTION DATE"
					strRows = 5
					gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 400")
					Wait (2)
				End Select
		End If
		WaitForIServeLoading
	bVerify = VerifytblPagination(tblObj,strcolumnname,strRows,strPage+" - Transaction Details")		
    Verifypagination_TranDetails = bVerify
End Function

'[Click Column Number from Product list table in Overview Page in TD]
Public Function SelectColumnNumberProductList_Overview(strProduct,lstProduct)
	bClickColNumber = False
	gObjIServePage.RunScript("document.getElementsByTagName('isrv-routing-proxy')[0].scrollTop = 100")
	WaitForIServeLoading
	If Not IsNull (strProduct) Then			
		Select Case strProduct
			Case "Deposits"
				Set ObjtableHeader = coOverview_Page.tblProductListDepositsHeader
				Set ObjtableContent = coOverview_Page.tblProductListDepositsContent
			Case "Credit Cards"	
				Set ObjtableHeader = coOverview_Page.tblProductListCCHeader
				Set ObjtableContent = coOverview_Page.tblProductListCCContent
			Case "Cashline"
				Set ObjtableHeader = coOverview_Page.tblProductListCLHeader
				Set ObjtableContent = coOverview_Page.tblProductListCLContent
			Case "Loans"
				Set ObjtableHeader = coOverview_Page.tblProductListLoanHeader
				Set ObjtableContent = coOverview_Page.tblProductListLoanContent
			Case "Debit/ATM Cards"
				Set ObjtableHeader = coOverview_Page.tblProductListDebitHeader
				Set ObjtableContent = coOverview_Page.tblProductListDebitContent
		End Select
	End If	
	bClickColNumber = SelectTableRow(ObjtableHeader,ObjtableContent,lstProduct ,strProduct & " Overview Product List Table","NUMBER",False,False)
	WaitForIServeLoading
	SelectColumnNumberProductList_Overview = bClickColNumber
	Set ObjtableHeader = Nothing
	Set ObjtableContent = Nothing
End Function