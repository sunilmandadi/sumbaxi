'*****This is auto generated code using code generator please Re-validate ****************

'[Verify Cheque Info Page opened successfully]
Public Function verifyChequeInfo()
	bverifyChequeInfo=true
	WaitForICallLoading
	If (bcCheque_Info.tblChequeInfoHeader().exist(0)) Then
	  bverifyChequeInfo = true
	else
	  LogMessage "RSLT","Verification","Unable to open Cheque Info page successfully." ,False
	  bverifyChequeInfo = false	  
	End If
	verifyChequeInfo=bverifyChequeInfo
End Function

'[Verify row data in Cheque Info Screen]
Public Function verifyChequeInfo_Details(lstlstChequeInfo)
	bverifyChequeInfo_Details=true
	If Not IsNull (lstlstChequeInfo) Then			
		bverifyChequeInfo_Details=verifyTableContentList(bcCheque_Info.tblChequeInfoHeader,bcCheque_Info.tblChequeInfoContent,lstlstChequeInfo,"Cheque Info",false,null,null,null)
	End If
	verifyChequeInfo_Details = bverifyChequeInfo_Details
End Function

'[Verify the row data for Cheque Transaction Details]
Public Function verifyChequeTransaction_Details(lstlstChequeTrans)
	bClickChequeTransactiondetails=true
	If Not IsNull (lstlstChequeTrans) Then			
		bClickChequeTransactiondetails = verifyTableContentList(bcCheque_Info.tblChequeInfoHeader,bcCheque_Info.tblChequeInfoContent,lstlstChequeTrans,"Cheque Transaction",false,null,null,null)
	End If
	verifyChequeTransaction_Details = bClickChequeTransactiondetails
End Function

'[Click on Cheque Status in Cheque Info Table]
Public Function ClickChequeStatus_ChequeInfo(lstChequeInfo)
	bClickChequeStatus_ChequeInfo=true
	If Not IsNull (lstChequeInfo) Then		
		bClickChequeStatus_ChequeInfo=selectTableLink(bcCheque_Info.tblChequeInfoHeader,bcCheque_Info.tblChequeInfoContent,lstChequeInfo,"Cheque Info" ,"Cheque Status",false,null,null,null)
	End If
	WaitForICallLoading
	If not bcCheque_Info.lblIssuedDate.Exist Then
		LogMessage "RSLT","Verification","Cheque Detail section not open successfully." ,False
		bClickChequeStatus_ChequeInfo=false
	End If
	ClickChequeStatus_ChequeInfo=bClickChequeStatus_ChequeInfo
End Function

'[Click on Mailing Address in Cheque Info Table]
Public Function ClickMailingAddress_ChequeInfo(lstChequeInfo)
	bClickChequeStatus_ChequeInfo=true
	If Not IsNull (lstChequeInfo) Then		
		bClickChequeStatus_ChequeInfo=selectTableLink(bcCheque_Info.tblChequeInfoHeader,bcCheque_Info.tblChequeInfoContent,lstChequeInfo,"Cheque Info" ,"Mailing Address",true,bcCheque_Info.lnkNext,bcCheque_Info.lnkNext1,bcCheque_Info.lnkPrevious)
	End If
	WaitForICallLoading
	If not bcCheque_Info.PopupMailingAddress.Exist Then
		LogMessage "RSLT","Verification","Mailing Address popup not open successfully." ,False
		bClickChequeStatus_ChequeInfo=false
	End If
	ClickMailingAddress_ChequeInfo=bClickChequeStatus_ChequeInfo
End Function

'[Verify Mailing Address Detail popup dispalyed as]
Public Function verifyMailingAddressDetails(strAddressLine1,strAddressLine2,strAddressLine3)
	bverifyMailingAddressDetails=true
	'If Not IsNull(strName) Then
     '  If Not VerifyField(bcCheque_Info.lblName_ChequeInfo(), strName, "Name")Then
     '      bverifyMailingAddressDetails=false
     '  End If
   'End If
   If Not IsNull(strAddressLine1) Then
       If Not verifyInnerText(bcCheque_Info.lblAddressLine1_ChequeInfo(), strAddressLine1, "Address Line1")Then
           bverifyMailingAddressDetails=false
       End If
   End If
   If Not IsNull(strAddressLine2) Then
       If Not verifyInnerText(bcCheque_Info.lblAddressLin2_ChequeInfo(), strAddressLine2, "Address Line2")Then
           bverifyMailingAddressDetails=false
       End If
   End If
   If Not IsNull(strAddressLine3) Then
       If Not verifyInnerText(bcCheque_Info.lblAddressLine3_ChequeInfo(), strAddressLine3, "Address Line3")Then
           bverifyMailingAddressDetails=false
       End If
   End If
   bcCheque_Info.btnOK_ChequeInfo.Click
   verifyMailingAddressDetails=bverifyMailingAddressDetails
End Function

'[Verify Number of Records per Page for Cheque Info]
Public Function verifyNoRecordsChequeInfo()
	bverifyNoRecordsChequeInfo=true
	Dim intRow,intRowtemp
	intRow = 0
    bNextPageExist = true	
	Do
	intRowtemp = 0
    intRowTemp = getRecordsCountForColumn (bcCheque_Info.tblChequeInfoHeader,bcCheque_Info.tblChequeInfoContent,"Cheque Range")
	'Added to check pagination functionality		
		If intRowTemp <= 5 Then
			LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to 5, Actual Row count :"& intRowTemp, true
			bverifyNoRecordsChequeInfo=true
		Else
			LogMessage "RSLT","Verification","Number of records displayed per page is more than 5 record. Expected Count is less than or equal to 5,  Actual Row count :"& intRowTemp, false
			bverifyNoRecordsChequeInfo=false
		End If

	Loop while Not  bNextPageExist
	verifyNoRecordsChequeInfo=bverifyNoRecordsChequeInfo
End Function

'[Verify Cheque Details Label in Cheque Info Screen]
Public Function verifyChequeDetails_Label(strCardNumber,strChequeNo,strIssueDate,strPaidCheque,strTotalCheque)
	bverifyChequeDetails_Label=true
	If Not Isnull (strIssueDate) Then
		If strIssueDate="RUNTIME" Then
			getChequeDetails_DCNS strCardNumber,strChequeNo
			strDate=trim(strRunTimeIssueDate)
			If len(Day(CDate(strDate)))=1 Then
			
				strDay="0"&Day(CDate(strDate))
			else
				strDay=""&Day(CDate(strDate))
			End If
			strIssueDate=""&strDay & " "&monthName(Month(CDate(strDate)),true) &" " &Year(CDate(strDate))
			strPaidCheque=trim(strRunTimePaidCheque)
			strTotalCheque=trim(strRunTimeTotalCheque)
		End If
		If Not verifyInnerText(bcCheque_Info.lblIssuedDate(),strIssueDate,"Issue Date")Then
			bverifyChequeDetails_Label = False
		End If
		If Not IsNull (strPaidCheque) Then		
			If Not verifyInnerText(bcCheque_Info.lblPaidCheque(),strPaidCheque,"Paid Cheque")Then
				bverifyChequeDetails_Label = False
			End If
		End If
		If Not verifyInnerText(bcCheque_Info.lblTotalNoofCheques(),strTotalCheque,"Total Cheque")Then
			bverifyChequeDetails_Label = False
		End If
	End If
	verifyChequeDetails_Label=bverifyChequeDetails_Label
End Function

'[Verify row data in Cheque Details Table]
Public Function verifyChequeDetails(lstlstChequeDetails)
	bverifyChequeDetails=true
	If Not IsNull (lstlstChequeDetails) Then			
		verifyChequeDetails=verifyTableContentList(bcCheque_Info.tblChequeStatusHeader,bcCheque_Info.tblChequeStatusTable,lstlstChequeDetails,"Cheque Details",false,null,null,null)
	End If
End Function

'[Click on Amount in Cheque Detail Table]
Public Function ClickChequeStatus_ChequeDetail(lstChequeDetail)
	bClickChequeStatus_ChequeDetail=true
	If Not IsNull (lstChequeInfo) Then		
		ClickChequeStatus_ChequeDetail=selectTableLink(bcCheque_Info.tblChequeStatusHeader,bcCheque_Info.tblChequeStatusTable,lstChequeDetail,"Cheque Details" ,"Amount",false,null,null,null)
	End If
	WaitForICallLoading	
End Function

'[Expand the account for cheque info verification when selected account is closed]
Public Function clickAccCheqInfo_ClosedAcc()
	bclickAccCheqInfo_ClosedAcc = true
	 bcCheque_Info.tblCusAccClosedAccount_Content.click
	 clickAccCheqInfo_ClosedAcc = bclickAccCheqInfo_ClosedAcc
End Function
