'[Click link SMS Threshold displayed in the left Menu]
Public Function ClickLink_SMSThreshold()
bClickLink_SMSThreshold=true
	bcAccountOverview_LeftMenu.clickSMSThresholds
	WaitForIcallLoading
	If Err.Number<>0 Then
       bClickLink_SMSThreshold=false
       LogMessage "WARN","Verification","Failed to Click Link  : SMS Threshold" ,false
       Exit Function
	End If
	Wait 1
	waitForIcallLoading	
ClickLink_SMSThreshold = bClickLink_SMSThreshold
End Function 

'[Click Button SMS Alert]
 Public Function clickButtonSMSAlert()
   WaitForIcallLoading
   SMS_Thresholds.btnSMSAlert.click
   If Err.Number<>0 Then
       clickButtonSMSAlert=false
            LogMessage "WARN","Verification","Failed to Click Button : SMSAlert" ,false
       Exit Function
   End If
   clickButtonSMSAlert=true
End Function

'[Click Button Proceed]
 Public Function clickButtonProceed()
   SMS_Thresholds.btnProceed.click
   If Err.Number<>0 Then
       clickButtonProceed=false
            LogMessage "WARN","Verification","Failed to Click Button : Proceed" ,false
       Exit Function
   End If
   clickButtonProceed=true
End Function
  
'[Verify label selected Account Card displayed in SMS Alert Page]
Public Function verifylabel_SelectedCard()
   bverifylabel_SelectedCard=true
   If Not VerifyInnerText (SMS_Thresholds.lblSelectedAccount(), "Selected Account/Card", "label:Selected Account/Card")Then
      bverifylabel_SelectedCard=false
   End If
   verifylabel_SelectedCard=bverifylabel_SelectedCard
End Function

'[Verify label Current SMS Threshold displayed in SMS Alert Page]
Public Function verifylabel_CurrentSMSThrehsold()
   bverifylabel_CurrentSMSThrehsold=true
   If Not VerifyInnerText (SMS_Thresholds.lblCurrentSMSThreshold(), "Current SMS Threshold", "label:Current SMS Threshold")Then
      bverifylabel_CurrentSMS=false
   End If
   verifylabel_CurrentSMSThrehsold=bverifylabel_CurrentSMSThrehsold
End Function

'[Verify label Select Card displayed in SMS Alert Page]
Public Function verifylabelSelectcard_SMSAlert()
   bverifylabelSelectcard_SMSAlert=true
   If Not VerifyInnerText (SMS_Thresholds.lblSelectCard(), "Select Card", "label:SMS_Alert Select Card")Then
      bverifylabelSelectcard_SMSAlert=false
   End If
   verifylabelSelectcard_SMSAlert=bverifylabelSelectcard_SMSAlert
End Function

'[Verify Icon Message Select maximum of 3 Cards displayed in Amend SMS Card List Page]
Public Function verifylabel_CurrentSMS()
   bDevPending=true
   bverifylabel_CurrentSMS=true
   If Not VerifyInnerText (SMS_Thresholds.lblSelectMax3Cards(), "Select maximum of 3 cards", "label:Select maximum of 3 cards")Then
      bverifylabel_CurrentSMS=false
   End If
   verifylabel_CurrentSMS=bverifylabel_CurrentSMS
End Function

'[Verify verification message displayed on navigating to SMS Alert page]
Public Function verifyMAMessage_SMSThreshold(strExpectedText)
   WaitForICallLoading
   bverifyMAMessage_SMSThreshold=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (SMS_Thresholds.lblValidationMsg(), strExpectedText, "Verification Message")Then
           bverifyMAMessage_SMSThreshold=false
       End If
   End If
   SMS_Thresholds.btnOK.Click
   verifyMAMessage_SMSThreshold=bverifyMAMessage_SMSThreshold
End Function

'[Verify row data in Table SelectedCards for Amend SMS Threshold]
Public Function verifytblSelectedCards_SMSThreshold(arrRowDataList)
   waitForIcallLoading
   verifytblSelectedCards_SMSThreshold=verifyTableContentList(SMS_Thresholds.tblSelectedCardsHeader,SMS_Thresholds.tblSelectedCardsContent,arrRowDataList,"Selected Cards Content" , false,null ,null,null)
End Function

'[Verify row data in Transaction type table in SMS Threshold Page]
 Public Function VerifyTransactiontable_SMSThreshold(lstSMSThresholdtable)
 wait(5)
  waitForIcallLoading
   	bverifytransactiondetails = True 	
	If Not IsNull(lstSMSThresholdtable) Then
			strVPRetailLocalAmt = trim(lstSMSThresholdtable(0))
			strVPCashLocalAmt = trim(lstSMSThresholdtable(1))
			strVPECommerceLocalAmt = trim(lstSMSThresholdtable(2))
			strVPRecurringLocalAmt = trim(lstSMSThresholdtable(3))
			strVPMailOrderLocalAmt = trim(lstSMSThresholdtable(4))
			strVPRetailForeignAmt = trim(lstSMSThresholdtable(5))
			strVPCashForeignAmt = trim(lstSMSThresholdtable(6))
			strVPECommerceForeignAmt = trim(lstSMSThresholdtable(7))
			strVPRecurringForeignAmt = trim(lstSMSThresholdtable(8))
			strVPMailOrderForeignAmt = trim(lstSMSThresholdtable(9))	
		lstSMSThresholdtable= (checknull("(Transaction Type:Retail|Local:"&strVPRetailLocalAmt&"|Overseas:"&strVPRetailForeignAmt&")|(Transaction Type:Cash|Local:"&strVPCashLocalAmt&"|Overseas:"&strVPCashForeignAmt&")|(Transaction Type:E-Commerce|Local:"&strVPECommerceLocalAmt&"|Overseas:"&strVPECommerceForeignAmt&")|(Transaction Type:Recurring|Local:"&strVPRecurringLocalAmt&"|Overseas:"&strVPRecurringForeignAmt&")|(Transaction Type:Mail Order|Local:"&strVPMailOrderLocalAmt&"|Overseas:"&strVPMailOrderForeignAmt&")|"))
	    VerifyTransactiontable_SMSThreshold=verifyTableContentList(SMS_Thresholds.tblSMSThresholdTransHeader,SMS_Thresholds.tblSMSThreshholdTransContent,lstSMSThresholdtable,"Amend SMS Threshold table",false,null,null,null)
	    If Not VerifyTransactiontable_CurrentSMSThreshold = False Then
	    	bverifytransactiondetails = False 
	    End If
	End If
	VerifyTransactiontable_SMSThreshold = bverifytransactiondetails	
'VerifyTransactiontable_SMSThreshold(strCardNumber,lstSMSThresholdtable)
' Call VerifySMSMinAmount_ARQE_Vplus(strCardNumber)
' 		 strVPRetailLocalAmt = Environment.Value("strVPRetailLocalAmt")
'		 strVPCashLocalAmt = Environment.Value("strVPCashLocalAmt")
'		 strVPECommerceLocalAmt = Environment.Value("strVPECommerceLocalAmt") 
' 		 strVPRecurringLocalAmt = Environment.Value("strVPRecurringLocalAmt") 
'		 strVPMailOrderLocalAmt = Environment.Value("strVPMailOrderLocalAmt") 	 
'		 strVPRetailForeignAmt = Environment.Value("strVPRetailForeignAmt")  
'		 strVPCashForeignAmt = Environment.Value("strVPCashForeignAmt")  
'		 strVPECommerceForeignAmt = Environment.Value("strVPECommerceForeignAmt")
'		 strVPRecurringForeignAmt = Environment.Value("strVPRecurringForeignAmt") 
'		 strVPMailOrderForeignAmt = Environment.Value("strVPMailOrderForeignAmt") 
'   lstSMSThresholdtable= (checknull("(Transaction Type:Retail|Local:"&strVPRetailLocalAmt&"|Overseas:"&strVPRetailForeignAmt&")|(Transaction Type:Cash|Local:"&strVPCashLocalAmt&"|Overseas:"&strVPCashForeignAmt&")|(Transaction Type:E-Commerce|Local:"&strVPECommerceLocalAmt&"|Overseas:"&strVPECommerceForeignAmt&")|(Transaction Type:Recurring|Local:"&strVPRecurringLocalAmt&"|Overseas:"&strVPRecurringForeignAmt&")|(Transaction Type:Mail Order|Local:"&strVPMailOrderLocalAmt&"|Overseas:"&strVPMailOrderForeignAmt&")|"))
' VerifyTransactiontable_SMSThreshold=verifyTableContentList(SMS_Thresholds.tblSMSThresholdTransHeader,SMS_Thresholds.tblSMSThreshholdTransContent,lstlstSMSThresholdtable,"Amend SMS Threshold table",false,null,null,null)
 End Function  

'[Verify row data in Current SMS Threshold table1 in Amend SMS Threshold Page]
Public Function VerifyCurrentSMSThreshold_table1(lstTHresholddetails)
 	wait(3)
 	bverifytransactiondetails = True 	
	If Not IsNull(lstTHresholddetails) Then
	 waitForIcallLoading
			strVPRetailLocalAmt = trim(lstTHresholddetails(0))
			strVPCashLocalAmt = trim(lstTHresholddetails(1))
			strVPECommerceLocalAmt = trim(lstTHresholddetails(2))
			strVPRecurringLocalAmt = trim(lstTHresholddetails(3))
			strVPMailOrderLocalAmt = trim(lstTHresholddetails(4))
			strVPRetailForeignAmt = trim(lstTHresholddetails(5))
			strVPCashForeignAmt = trim(lstTHresholddetails(6))
			strVPECommerceForeignAmt = trim(lstTHresholddetails(7))
			strVPRecurringForeignAmt = trim(lstTHresholddetails(8))
			strVPMailOrderForeignAmt = trim(lstTHresholddetails(9))					
		lstCurrentSMSThresholdtable1= (checknull("(Transaction Type:Retail|Local:"&strVPRetailLocalAmt&"|Foreign:"&strVPRetailForeignAmt&")|(Transaction Type:Cash|Local:"&strVPCashLocalAmt&"|Foreign:"&strVPCashForeignAmt&")|(Transaction Type:E-Commerce|Local:"&strVPECommerceLocalAmt&"|Foreign:"&strVPECommerceForeignAmt&")|(Transaction Type:Recurring|Local:"&strVPRecurringLocalAmt&"|Foreign:"&strVPRecurringForeignAmt&")|(Transaction Type:Mail Order|Local:"&strVPMailOrderLocalAmt&"|Foreign:"&strVPMailOrderForeignAmt&")|"))
	    VerifyTransactiontable_CurrentSMSThreshold=verifyTableContentList(SMS_Thresholds.tblCurrentSMSThresholdHeader_Card1,SMS_Thresholds.tblCurrentSMSThresholdContent_Card1,lstCurrentSMSThresholdtable1,"Current SMS Threshold table1",false,null,null,null)		
	    If Not VerifyTransactiontable_CurrentSMSThreshold Then
	    	bverifytransactiondetails = False 
	    End If
	End If
	VerifyCurrentSMSThreshold_table1 = bverifytransactiondetails
End Function 
	
'[Verify row data in Current SMS Threshold table2 in Amend SMS Threshold Page]
Public Function VerifyCurrentSMSThreshold_table2(lstTHresholddetails)
 	waitForIcallLoading
 	bverifytransactiondetails = True 	
	If Not IsNull(lstTHresholddetails) Then
		'intSize = Ubound(lstCard1)
		'For Iterator = 0 To intSize Step 1
			strVPRetailLocalAmt = trim(lstTHresholddetails(0))
			strVPCashLocalAmt = trim(lstTHresholddetails(1))
			strVPECommerceLocalAmt = trim(lstTHresholddetails(2))
			strVPRecurringLocalAmt = trim(lstTHresholddetails(3))
			strVPMailOrderLocalAmt = trim(lstTHresholddetails(4))
			strVPRetailForeignAmt = trim(lstTHresholddetails(5))
			strVPCashForeignAmt = trim(lstTHresholddetails(6))
			strVPECommerceForeignAmt = trim(lstTHresholddetails(7))
			strVPRecurringForeignAmt = trim(lstTHresholddetails(8))
			strVPMailOrderForeignAmt = trim(lstTHresholddetails(9))					
		lstCurrentSMSThresholdtable2= (checknull("(Transaction Type:Retail|Local:"&strVPRetailLocalAmt&"|Foreign:"&strVPRetailForeignAmt&")|(Transaction Type:Cash|Local:"&strVPCashLocalAmt&"|Foreign:"&strVPCashForeignAmt&")|(Transaction Type:E-Commerce|Local:"&strVPECommerceLocalAmt&"|Foreign:"&strVPECommerceForeignAmt&")|(Transaction Type:Recurring|Local:"&strVPRecurringLocalAmt&"|Foreign:"&strVPRecurringForeignAmt&")|(Transaction Type:Mail Order|Local:"&strVPMailOrderLocalAmt&"|Foreign:"&strVPMailOrderForeignAmt&")|"))
	    VerifyTransactiontable_CurrentSMSThreshold=verifyTableContentList(SMS_Thresholds.tblCurrentSMSThresholdHeader_Card2,SMS_Thresholds.tblCurrentSMSThresholdContent_Card2,lstCurrentSMSThresholdtable2,"Current SMS Threshold table2",false,null,null,null)
	    If Not VerifyTransactiontable_CurrentSMSThreshold Then
	    	bverifytransactiondetails = False 
	    End If
	End If
	VerifyCurrentSMSThreshold_table2 = bverifytransactiondetails
End Function 

'[Verify row data in Current SMS Threshold table3 in Amend SMS Threshold Page]
Public Function VerifyCurrentSMSThreshold_table3(lstTHresholddetails)
 	waitForIcallLoading
 	bverifytransactiondetails = True 	
' VerifyTransactiontable_CurrentSMSThreshold(strCardNumber, strtable)
' Call VerifySMSMinAmount_ARQE_Vplus(strCardNumber)
' 		 strVPRetailLocalAmt = Environment.Value("strVPRetailLocalAmt")
'		 strVPCashLocalAmt = Environment.Value("strVPCashLocalAmt")
'		 strVPECommerceLocalAmt = Environment.Value("strVPECommerceLocalAmt")
' 		 strVPRecurringLocalAmt = Environment.Value("strVPRecurringLocalAmt")
'		 strVPMailOrderLocalAmt = Environment.Value("strVPMailOrderLocalAmt") 
'		 strVPRetailForeignAmt = Environment.Value("strVPRetailForeignAmt")  
'		 strVPCashForeignAmt = Environment.Value("strVPCashForeignAmt")
'		 strVPECommerceForeignAmt = Environment.Value("strVPECommerceForeignAmt") 
'		 strVPRecurringForeignAmt = Environment.Value("strVPRecurringForeignAmt")
'		 strVPMailOrderForeignAmt = Environment.Value("strVPMailOrderForeignAmt")
	If Not IsNull(lstTHresholddetails) Then
		'intSize = Ubound(lstCard1)
		'For Iterator = 0 To intSize Step 1
			strVPRetailLocalAmt = trim(lstTHresholddetails(0))
			strVPCashLocalAmt = trim(lstTHresholddetails(1))
			strVPECommerceLocalAmt = trim(lstTHresholddetails(2))
			strVPRecurringLocalAmt = trim(lstTHresholddetails(3))
			strVPMailOrderLocalAmt = trim(lstTHresholddetails(4))
			strVPRetailForeignAmt = trim(lstTHresholddetails(5))
			strVPCashForeignAmt = trim(lstTHresholddetails(6))
			strVPECommerceForeignAmt = trim(lstTHresholddetails(7))
			strVPRecurringForeignAmt = trim(lstTHresholddetails(8))
			strVPMailOrderForeignAmt = trim(lstTHresholddetails(9))					
		lstCurrentSMSThresholdtable3= (checknull("(Transaction Type:Retail|Local:"&strVPRetailLocalAmt&"|Foreign:"&strVPRetailForeignAmt&")|(Transaction Type:Cash|Local:"&strVPCashLocalAmt&"|Foreign:"&strVPCashForeignAmt&")|(Transaction Type:E-Commerce|Local:"&strVPECommerceLocalAmt&"|Foreign:"&strVPECommerceForeignAmt&")|(Transaction Type:Recurring|Local:"&strVPRecurringLocalAmt&"|Foreign:"&strVPRecurringForeignAmt&")|(Transaction Type:Mail Order|Local:"&strVPMailOrderLocalAmt&"|Foreign:"&strVPMailOrderForeignAmt&")|"))
	    VerifyTransactiontable_CurrentSMSThreshold=verifyTableContentList(SMS_Thresholds.tblCurrentSMSThresholdHeader_Card3,SMS_Thresholds.tblCurrentSMSThresholdContent_Card3,lstCurrentSMSThresholdtable3,"Current SMS Threshold table3",false,null,null,null)
	    If Not VerifyTransactiontable_CurrentSMSThreshold  Then
	    	bverifytransactiondetails = False 
	    End If
	End If
	VerifyCurrentSMSThreshold_table3 = bverifytransactiondetails
End Function

'[Verify defaulted Local and Foreign SMS Threshold Amount displayed for different Transaction Types]
Public Function verifySMSThreshold_defaultAmount(lstTHresholddetails)
'verifySMSThreshold_defaultAmount(strCardNumber)
   bDevPending=false
   bverifySMSThreshold_defaultAmount=true
   
'   Call VerifySMSMinAmount_ARQE_Vplus(strCardNumber)
'	 strVPRetailLocalAmt = Environment.Value("strVPRetailLocalAmt")
'	 strVPCashLocalAmt = Environment.Value("strVPCashLocalAmt") 
'	 strVPECommerceLocalAmt = Environment.Value("strVPECommerceLocalAmt")
'	 strVPRecurringLocalAmt = Environment.Value("strVPRecurringLocalAmt") 
'	 strVPMailOrderLocalAmt = Environment.Value("strVPMailOrderLocalAmt")  
'	 strVPRetailForeignAmt = Environment.Value("strVPRetailForeignAmt") 
'	 strVPCashForeignAmt = Environment.Value("strVPCashForeignAmt")
'	 strVPECommerceForeignAmt = Environment.Value("strVPECommerceForeignAmt")  
'	 strVPRecurringForeignAmt = Environment.Value("strVPRecurringForeignAmt") 
'	 strVPMailOrderForeignAmt = Environment.Value("strVPMailOrderForeignAmt")
	strVPRetailLocalAmt = trim(lstTHresholddetails(0))
	strVPCashLocalAmt = trim(lstTHresholddetails(1))
	strVPECommerceLocalAmt = trim(lstTHresholddetails(2))
	strVPRecurringLocalAmt = trim(lstTHresholddetails(3))
	strVPMailOrderLocalAmt = trim(lstTHresholddetails(4))
	strVPRetailForeignAmt = trim(lstTHresholddetails(5))
	strVPCashForeignAmt = trim(lstTHresholddetails(6))
	strVPECommerceForeignAmt = trim(lstTHresholddetails(7))
	strVPRecurringForeignAmt = trim(lstTHresholddetails(8))
	strVPMailOrderForeignAmt = trim(lstTHresholddetails(9))	

	strIserveRetailLocalAmt = SMS_Thresholds.RetailLocalAmt.GetROProperty("value")
	strIserveCashLocalAmt = SMS_Thresholds.CashLocalAmt.GetROProperty("value")
	strIserveECommerceLocalAmt = SMS_Thresholds.EcommerceLocalAmt.GetROProperty("value")
	strIserveRecurringLocalAmt = SMS_Thresholds.RecurringLocalAmt.GetROProperty("value")
	strIserveMailOrderLocalAmt = SMS_Thresholds.MailOrderLocalAmt.GetROProperty("value")
	strIserveRetailForeignAmt = SMS_Thresholds.RetailForeignAmt.GetROProperty("value")
	strIserveCashForeignAmt = SMS_Thresholds.CashForeignAmt.GetROProperty("value")
	strIserveECommerceForeignAmt = SMS_Thresholds.EcommerceForeignAmt.GetROProperty("value")
	strIserveRecurringForeignAmt = SMS_Thresholds.RecurringForeignAmt.GetROProperty("value")
	strIserveMailOrderForeignAmt = SMS_Thresholds.MailOrderForeignAmt.GetROProperty("value")		
	
   If Trim(strIserveRetailLocalAmt) =  Trim(strVPRetailLocalAmt) Then
			LogMessage "RSLT","Verification","Retail Local Amount value matched with the expected Value Expected: "&strIserveRetailLocalAmt&" Actual: "&strVPRetailLocalAmt&"",True
		Else
			LogMessage "WARN","Verification","Retail Local Amount value doesnt match with the expected Value Expected: "&strIserveRetailLocalAmt&" Actual: "&strVPRetailLocalAmt&"",False	
   End IF 
   
   If Trim(strIserveCashLocalAmt) =  Trim(strVPCashLocalAmt) Then
			LogMessage "RSLT","Verification","Cash Local Amount value matched with the expected Value Expected: "&strIserveCashLocalAmt&" Actual: "&strVPCashLocalAmt&"",True
		Else
			LogMessage "WARN","Verification","Cash Local Amount value doesnt match with the expected Value Expected: "&strIserveCashLocalAmt&" Actual: "&strVPCashLocalAmt&"",False	
   End IF   
   
   If Trim(strIserveECommerceLocalAmt) =  Trim(strVPECommerceLocalAmt) Then
			LogMessage "RSLT","Verification","E-commerce Local Amount value matched with the expected Value Expected: "&strIserveECommerceLocalAmt&" Actual: "&strVPECommerceLocalAmt&"",True
		Else
			LogMessage "WARN","Verification","E-commerce Local Amount value doesnt match with the expected Value Expected: "&strIserveECommerceLocalAmt&" Actual: "&strVPECommerceLocalAmt&"",False	
   End IF 
   
   If Trim(strIserveRecurringLocalAmt) =  Trim(strVPRecurringLocalAmt) Then
			LogMessage "RSLT","Verification","Recurring Local Amount value matched with the expected Value Expected: "&strIserveRecurringLocalAmt&" Actual: "&strVPRecurringLocalAmt&"",True
		Else
			LogMessage "WARN","Verification","Recurring Local Amount value doesnt match with the expected Value Expected: "&strIserveRecurringLocalAmt&" Actual: "&strVPRecurringLocalAmt&"",False	
   End IF    
   
   If Trim(strIserveMailOrderLocalAmt) =  Trim(strVPMailOrderLocalAmt) Then
			LogMessage "RSLT","Verification","MailOrder Local Amount value matched with the expected Value Expected: "&strIserveMailOrderLocalAmt&" Actual: "&strVPMailOrderLocalAmt&"",True
		Else
			LogMessage "WARN","Verification","MailOrder Local Amount value doesnt match with the expected Value Expected: "&strIserveMailOrderLocalAmt&" Actual: "&strVPMailOrderLocalAmt&"",False	
   End IF 
   
   If Trim(strIserveRetailForeignAmt) =  Trim(strVPRetailForeignAmt) Then
			LogMessage "RSLT","Verification","Retail Foreign Amount value matched with the expected Value Expected: "&strIserveRetailForeignAmt&" Actual: "&strVPRetailForeignAmt&"",True
		Else
			LogMessage "WARN","Verification","Retail Foreign Amount value doesnt match with the expected Value Expected: "&strIserveRetailForeignAmt&" Actual: "&strVPRetailForeignAmt&"",False	
   End IF    
   
   If Trim(strIserveCashForeignAmt) =  Trim(strVPCashForeignAmt) Then
			LogMessage "RSLT","Verification","Cash Foreign Amount value matched with the expected Value Expected: "&strIserveCashForeignAmt&" Actual: "&strVPCashForeignAmt&"",True
		Else
			LogMessage "WARN","Verification","Cash Foreign Amount value doesnt match with the expected Value Expected: "&strIserveCashForeignAmt&" Actual: "&strVPCashForeignAmt&"",False	
   End IF 
   
   If Trim(strIserveECommerceForeignAmt) =  Trim(strVPECommerceForeignAmt) Then
			LogMessage "RSLT","Verification","E-commerce Foreign Amount value matched with the expected Value Expected: "&strIserveECommerceForeignAmt&" Actual: "&strVPECommerceForeignAmt&"",True
		Else
			LogMessage "WARN","Verification","E-commerce Foreign Amount value doesnt match with the expected Value Expected: "&strIserveECommerceForeignAmt&" Actual: "&strVPECommerceForeignAmt&"",False	
   End IF  
   
   If Trim(strIserveRecurringForeignAmt) =  Trim(strVPRecurringForeignAmt) Then
			LogMessage "RSLT","Verification","Recurring Foreign Amount value matched with the expected Value Expected: "&strIserveRecurringForeignAmt&" Actual: "&strVPRecurringForeignAmt&"",True
		Else
			LogMessage "WARN","Verification","Recurring Foreign Amount value doesnt match with the expected Value Expected: "&strIserveRecurringForeignAmt&" Actual: "&strVPRecurringForeignAmt&"",False	
   End IF 
   
   If Trim(strIserveMailOrderForeignAmt) =  Trim(strVPMailOrderForeignAmt) Then
			LogMessage "RSLT","Verification","Mail Order foreign Amount value matched with the expected Value Expected: "&strIserveMailOrderForeignAmt&" Actual: "&strVPMailOrderForeignAmt&"",True
		Else
			LogMessage "WARN","Verification","Mail Order foreign Amount value doesnt match with the expected Value Expected: "&strIserveMailOrderForeignAmt&" Actual: "&strVPMailOrderForeignAmt&"",False	
   End IF   
      
End Function

''[Verify Field Description displayed on SMS Alert page as]
'Public Function verifyDescriptionText_SMSThreshold(strExpectedText)
'   bDevPending=true
'   bverifyDescriptionText_SMSThreshold=true
'   If Not IsNull(strExpectedText) Then
'       If Not VerifyInnerText (SMS_Thresholds.lblDescription(), strExpectedText, "Description")Then
'           bverifyDescriptionText_SMSThreshold=false
'       End If
'   End If
'   verifyDescriptionText_SMSThreshold=bverifyDescriptionText_SMSThreshold
'End Function

'[Verify Field KnowledgeBase on Amend SMS Threshold SR Screen displayed as]
Public Function verifyKnowledgeBase_SMSThreshold(strExpectedLink)
   bDevPending=false
   bverifyKnowledgeBase_SMSThreshold=true
   If Not IsNull(strExpectedLink) Then		
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"		
			strKBLink=SMS_Thresholds.lnkKnowledgeBase.GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bverifyKnowledgeBase_SMSThreshold=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_SMSThreshold=bverifyKnowledgeBase_SMSThreshold
End Function

'[Click Button Cancel on Amend SMS Threshold Page]
Public Function clickButtonCancel_SMSThreshold()
   bDevPending=false
   SMS_Thresholds.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_SMSThreshold=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel_SMSThreshold=true
End Function

'[Verify the Cancel Confirmation message in Amend SMS Threshold page displayed as]
Public Function verifyConfirmationCancel_SMSThreshold(strConfirmMsg)
   bDevPending=false
   bverifyConfirmationCancel_SMSThreshold=true
   If Not IsNull(strConfirmMsg) Then
       If Not verifyInnerText(SMS_Thresholds.lblValidationMsg(),strConfirmMsg, "ConfirmationPopup")Then
			bverifyConfirmationCancel_SMSThreshold = False
		End If
   End If
   verifyConfirmationCancel_SMSThreshold=bverifyConfirmationCancel_SMSThreshold
End Function

'[Click Button NO in the Message displayed]
Public Function clickButtonNO_SMSThreshold()
   bDevPending=false
   SMS_Thresholds.btnNO.click
   If Err.Number<>0 Then
       clickButtonNO_SMSThreshold=false
            LogMessage "WARN","Verification","Failed to Click Button : NO in Confirmation Popup" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonNO_SMSThreshold=true
End Function

'[Click Button Yes in the Message displayed]
Public Function clickButtonYES_SMSThreshold()
   bDevPending=false
   SMS_Thresholds.btnYES.click
   If Err.Number<>0 Then
       clickButtonYES_SMSThreshold=False
            LogMessage "WARN","Verification","Failed to Click Button : YES in Confirmation Popup" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonYES_SMSThreshold=true
End Function

'[Verify Proceed Button display on Amend SMS Threshold Page]
Public Function VerifybtnProceed_SMSThreshold(strAction)
	bDevPending=false
   	bVerifybtndisplay_SMSThreshold=true
    If strAction = "Enabled" Then
    	intBtnSubmit=Instr(SMS_Thresholds.btnProceed.GetROproperty("outerhtml"),("v-disabled"))
		If  intBtnSubmit=0 Then
			LogMessage "RSLT","Verification","Proceed button is enabled as expected.",True
			bVerifybtndisplay_SMSThreshold=true
		Else
			LogMessage "WARN","Verifiation","Proceed button is disabled.",false
			bVerifybtndisplay_SMSThreshold=false
		End If
	ElseIf strAction = "Disabled" Then
	    	intBtnSubmit=Instr(SMS_Thresholds.btnProceed.GetROproperty("outerhtml"),("v-enabled"))
		If  intBtnSubmit=0 Then
			LogMessage "RSLT","Verification","Proceed button is disabled as expected.",True
			bVerifybtndisplay_SMSThreshold=true
		Else
			LogMessage "WARN","Verifiation","Proceed button is enabled.",false
			bVerifybtndisplay_SMSThreshold=false
		End If
	End If 
'    End If
	VerifybtnProceed_SMSThreshold=bVerifybtndisplay_SMSThreshold
End Function

'[Verify Field ValidationMessage for SMS Threshold displayed as]
Public Function verifyValidationMessage_SMSThreshold(strExpectedText)
   bverifyValidationMessage=True
   WaitForIcallLoading
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (SMS_Thresholds.lblValidationMsg(), strExpectedText, "ValidationMessage") Then
           bverifyValidationMessage=False
       End If
      SMS_Thresholds.btnOK.click
   End If
   If Err.Number<>0 Then
       bverifyValidationMessage=false
       LogMessage "WARN","Verification","Failed to Click OK : Verification Required" ,false
       Exit Function
   End If
   WaitForICallLoading  
   verifyValidationMessage_SMSThreshold = bverifyValidationMessage
End Function

'[Verify cards table in SMS Alert Page has following Columns]
Public Function verifyCardsTableColumns_SMSAlert(arrColumnNameList)
   verifyCardsTableColumns_SMSAlert=verifyTableColumns(SMS_Thresholds.tblSMSAlertCardlistHeader,arrColumnNameList)
End Function

'[Verify row Data in Cardlist table displayed in SMS Alert Page]
Public Function verifytblCardsToAmendSMSThreshold_RowData(lstClosedCards)
   verifytblCardsToAmendSMSThreshold_RowData=verifyTableContentList(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,lstClosedCards,"Closed Card List" , false,null ,null,null)
End Function

Public Function selectCheckBoxForSingleCard_SMSAlert(lstlstCardlists)
   'This Function Selects Single Card
   intCount = Ubound(lstlstCardlists,1)
   iCount = 0
   Set objAllRows=getAllRows(SMS_Thresholds.tblSMSAlertCardlistContent)
   intRow = getRecordsCountForColumn(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,"Card Number")
   For i = 0 To intRow -1  
	   	strCardNumber=getCellTextForCheckBox(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,i,"Card Number")
	   	strCardStatus=getCellTextForCheckBox(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,i,"Card Status")
	   	
	   	verifytableSMS_SelectCard=verifyTableContentList(SMS_Thresholds.tblSMSAlertCardlistHeader,objAllRows(i),lstCardToAmendSMSThreshold,"SMS Alert Select Card",false,null,null,null)
	   	If verifytableSMS_SelectCard Then
	   		bCheckBoxEnabled = selectCheckbox_tblCell(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,i, "")
	   		iCount = iCount+1
	   	End If
	   	If iCount = intCount Then
	   		Exit For
	   	End If
   Next
   
End Function

'[Select Check box for Multiple Cards to Amend SMS Threshold values from Cardlist table]
Public Function selectCheckBoxForMultipleCard_SMSThreshold(lstlstCardsToAmendSMSThreshold)
   'This Function Selects multiple cards
   bDevPending=False
  	 Set objAllcolumn=getAllRows(SMS_Thresholds.tblSMSAlertCardlistHeader)
  		 objAllcolumn(0,0) = "Checkbox"
   selectCheckBoxForMultipleCard_SMSThreshold=selectTableCheckBox_MultipleRow(objAllcolumn,SMS_Thresholds.tblSMSAlertCardlistContent,lstlstCardsToAmendSMSThreshold,"SMS Alert Card lists" ,"Check", false,null ,null,null)
End Function

'[Verify Check box display for Cards to Amend SMS Threshold values displayed in SMS Alert Page]
Public Function verifyCheckBoxForCardlist_Enabled(lstCLosedCards,strStatus)
   'This Function Selects multiple cards Card
   bDevPending=false 
   strActualStatus=verifyCheckBoxEnabled_tblCell(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,lstCLosedCards,"Closed cards in Cardlist table" , "",false,null ,null,null)
  If UCase(strStatus)=Ucase(strActualStatus) Then
	LogMessage "RSLT","Verification","Expected Status for Check box Matched with Actual Status for Card Number : "&strCardNumber ,true
	bStatus=true
  else
	LogMessage "RSLT","Verification","Expected Status "&strStatus&" for Check box not Matched with Actual Status "&strActualStatus&" for Card Number : "&strCardNumber ,false
	bStatus=false
  End If
  verifyCheckBoxForCardlist_Enabled=bStatus
End Function

 '[Verify defaulted Fore SMS Threshold Amount displayed for different Transactions]
Public Function verifyRequestedAmount_FR(strRequestedAmount)
   bDevPending=false
   bverifyRequestedAmount_FR=true
   strRequestedAmount = FormatNumber(strRequestedAmount,2)
   If Not IsNull(strRequestedAmount) Then
       If Not verifyComboSelectItem (FeeReversal.txtRequestedAmount(),strRequestedAmount, "Requested Amount")Then
       	LogMessage "WARN","Verification","Requested Amount is not defaulted to the amount debited for the trasaction selected" ,false
        bverifyRequestedAmount_FR=false
       End If
   End If
   verifyRequestedAmount_FR=bverifyRequestedAmount_FR
End Function

Public Function VerifyCheckbox_SMSAlert
	intRecordCount = getRecordsCountForColumn(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent, "Card Status")'Change the table name otherwise there will be concerns
	''''''''/////*****OtherPlans.tblPlanSummaryHeader(0,0) = "CBox"
	For i = 0 To intRecordcount - 1
		sCardStatus = getCellTextFor(OtherPlans.tblPlanSummaryHeader,OtherPlans.tblPlanSummaryContent,i, "Plan No")	'this also
		If UCase(sCardStatus) = "CLOSED" Then
			
		End If		
	Next
End Function

'[Set SMS Amend values on SMS Alert Page]
Public Function setSMSAmendAmount_SMSAlert(strThresholdAmount, Strlabelname)
	bveridfysetAmt = True 
	WaitForICallLoading
   If Strlabelname = "RETAIL LOCAL" Then
   	     SMS_Thresholds.RetailLocalAmt.Set(strThresholdAmount)
   ElseIf Strlabelname = "RETAIL FOREIGN" Then
    	 SMS_Thresholds.RetailForeignAmt.Set(strThresholdAmount)  
   ElseIf Strlabelname = "CASH LOCAL" Then
   	     SMS_Thresholds.CashLocalAmt.Set(strThresholdAmount)
   ElseIf Strlabelname = "CASH FOREIGN" Then
    	 SMS_Thresholds.CashForeignAmt.Set(strThresholdAmount)    
   ElseIf Strlabelname = "ECOMMERCE LOCAL" Then
   	     SMS_Thresholds.EcommerceLocalAmt.Set(strThresholdAmount)
   ElseIf Strlabelname = "ECOMMERCE FOREIGN" Then
    	 SMS_Thresholds.EcommerceForeignAmt.Set(strThresholdAmount)  
   ElseIf Strlabelname = "RECURRING LOCAL" Then
   	     SMS_Thresholds.RecurringLocalAmt.Set(strThresholdAmount)
   ElseIf Strlabelname = "RECURRING FOREIGN" Then
    	 SMS_Thresholds.RecurringForeignAmt.Set(strThresholdAmount)  
   ElseIf Strlabelname = "MAILORDER LOCAL" Then
   	     SMS_Thresholds.MailOrderLocalAmt.Set(strThresholdAmount)
   ElseIf Strlabelname = "MAILORDER FOREIGN" Then
    	 SMS_Thresholds.MailOrderForeignAmt.Set(strThresholdAmount)  
   End If
   If Err.Number<>0 Then
       bveridfysetAmt =false
       LogMessage "WARN","Verification","Failed to Set Text Box :SMS Threshold values", false
       Exit Function
   End If  	
   setSMSAmendAmount_SMSAlert = bveridfysetAmt
End Function

'[Verify Inline Info error message displayed in SMS Alert page]
Public Function verifyInlineErrormsg_SMSAlert(strInLineMessage, Strlabelname)
	bverifyInlineerrormsg=true	
	wait 2
   If Strlabelname = "RETAIL LOCAL" Then
  	   If Not VerifyInnerText (SMS_Thresholds.lblLocalInlineErrorMsg1(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "RETAIL FOREIGN" Then
  	   If Not VerifyInnerText (SMS_Thresholds.lblForeignInlineErrorMsg1(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "CASH LOCAL" Then
   	   If Not VerifyInnerText (SMS_Thresholds.lblLocalInlineErrorMsg2(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "CASH FOREIGN" Then
       If Not VerifyInnerText (SMS_Thresholds.lblForeignInlineErrorMsg2(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If    
   ElseIf Strlabelname = "ECOMMERCE LOCAL" Then
   	   If Not VerifyInnerText (SMS_Thresholds.lblLocalInlineErrorMsg3(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "ECOMMERCE FOREIGN" Then
       If Not VerifyInnerText (SMS_Thresholds.lblForeignInlineErrorMsg3(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "RECURRING LOCAL" Then
   	   If Not VerifyInnerText (SMS_Thresholds.lblLocalInlineErrorMsg4(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "RECURRING FOREIGN" Then
    	If Not VerifyInnerText (SMS_Thresholds.lblForeignInlineErrorMsg4(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "MAILORDER LOCAL" Then
   	   If Not VerifyInnerText (SMS_Thresholds.lblLocalInlineErrorMsg5(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   ElseIf Strlabelname = "MAILORDER FOREIGN" Then
    	If Not VerifyInnerText (SMS_Thresholds.lblForeignInlineErrorMsg5(), strInLineMessage, "Inline error Message")Then
           bverifyInlineerrormsg=false
       End If 
   End If
   verifyInlineErrormsg_SMSAlert = bverifyInlineerrormsg
End Function

'[Verify Field Description displayed on SMS Alert Screen as]
Public Function verifyDescriptionText_SMSAlert(strExpectedText)
   bDevPending=true
   bverifyDescriptionText_SMSAlert=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (SMS_Thresholds.lblDescription(), strExpectedText, "Description")Then
           bverifyDescriptionText_SMSAlert=false
       End If
   End If
   verifyDescriptionText_SMSAlert=bverifyDescriptionText_SMSAlert
End Function

'[Perform Add Notes by clicking Add Notes Button on SMS Alert SR Screen]
'Public Function AddNote_SMSThreshold(strNote)
'   bDevPending=false
'   baddNote_SMSThreshold=true	
'	If not isNull(strNote) Then
'		SMS_Thresholds.btnAddNotes.click
'		WaitForICallLoading
'		SMS_Thresholds.txtComment_Notes.set strNote
'		If Err.Number<>0 Then
'       		baddNote_SMSThreshold=false
'            LogMessage "WARN","Verification","Failed to Set Text Box :Notes" ,false
'       	Exit Function
'       	End If
'		SMS_Thresholds.btnSave.Click
'		WaitForIcallLoading
'	End If		
'	AddNote_SMSThreshold=baddNote_SMSThreshold
'End Function
'
'[Set TextBox Comments to SMS Threshold]
Public Function setCommentsTextbox_SMSAlert(strComment)
   bDevPending=false
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comments to Fee Reversal"
	insertDataStore "SRComment", strComment
   SMS_Thresholds.txtComment.Set strComment
   If Err.Number<>0 Then
       setCommentsTextbox_SMSAlert=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comments" ,false
       Exit Function
   End If
   setCommentsTextbox_SMSAlert=true
End Function

'[Verify Button Submit is enabled on SMS Alert Screen]
Public Function VerifybtnSubmit_SMSAlert()
	bDevPending=false
    bVerifybtnSubmit_SMSAlert=true
	intBtnSubmit=Instr(SMS_Thresholds.btnSubmit.GetROproperty("outerhtml"),("v-disabled"))
	If  intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Submit button is enabled as expected.",True
		bVerifybtnSubmit_SMSAlert=true
	Else
		LogMessage "WARN","Verifiation","Submit button is disabled.",false
		bVerifybtnSubmit_SMSAlert=false
	End If
	VerifybtnSubmit_SMSAlert=bVerifybtnSubmit_SMSAlert
End Function

'[Click Button Submit on SMS Alert Screen]
Public Function clickButtonSubmit_SMSAlert()
   bDevPending=false
   SMS_Thresholds.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_SMSAlert=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonSubmit_SMSAlert=true
End Function

'[Verify Field CardNumber on Request Submitted Popup for SMS Alert displayed as]
Public Function verifyCardNumber_RequestSubmitted_SMSALert(strCardNumber)
   bDevPending=false
   bverifyCardNumber_RequestSubmitted=true
   insertDataStore "NewCardNumber", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       If Not VerifyInnerText (SMS_Thresholds.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bverifyCardNumber_RequestSubmitted=false
       End If
   End If
   verifyCardNumber_RequestSubmitted_SMSALert=bverifyCardNumber_RequestSubmitted
End Function

'[Verify Field ProductDescription on Request Submitted Popup for SMS Alert displayed as]
Public Function verifyProductDescription_RequestSubmitted_SMSAlert(strProductDescription)
   bDevPending=false
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (SMS_Thresholds.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_SMSAlert=bVerifyProductDescription_RequestSubmittedText
End Function

'[Verify Field SR Status_RequestSubmitted For SMS Alert displayed as]
Public Function verifySRStatus_RequestSubmitted(strExpectedText)
   bDevPending=false
   bVerifyStatus_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (SMS_Thresholds.lblSRStatus_RequestSubmitted(), strExpectedText, "Status_RequestSubmitted")Then
           bVerifyStatus_RequestSubmittedText=false
       End If
   End If
   verifySRStatus_RequestSubmitted=bVerifyStatus_RequestSubmittedText
End Function

'[Verify CaseNumber on Request Submitted Popup for Fee Reversal]
Public Function VerifyCaseNumber_RequestSubmitted_FR()
   bDevPending=false
   strCaseNumber=FeeReversal.lblCaseNumber_RequestSubmitted.GetRoProperty("innerText")
	If strCaseNumber<>"" Then
		 insertDataStore "CaseNumber", strCaseNumber
	   Environment.Value("CaseNumber") = strCaseNumber
	 else
   		LogMessage "RSLT","Verification","Case Number did not display on Request Submitted pop up",false
	End If
   WaitForIcallLoading
   VerifyCaseNumber_RequestSubmitted_FR=true
End Function

'[Verify SRNumber on Request Submitted Popup for Fee Reversal displayed]
Public Function VerifySRNumber_RequestSubmitted_FR()
   bDevPending=false
   strSRNumber=FeeReversal.lblCardNumber_RequestSubmitted.GetRoProperty("innerText")
	If strSRNumber<>"" Then
		 insertDataStore "SRNumber", strSRNumber
	   Environment.Value("SRNumber") = strSRNumber
	 else
   		LogMessage "RSLT","Verification","ServiceRequest Number did not display on Request Submitted pop up",false
	End If
   WaitForIcallLoading
   VerifySRNumber_RequestSubmitted_FR=true
End Function

'[Click button Close on Request Submitted Popup for SMS Alert]
Public Function clickBtnClose_ReqSubmit_SMSAlert()
	bDevPending=false
   SMS_Thresholds.btnClose.click
   If Err.Number<>0 Then
       clickBtnClose_ReqSubmit_SMSAlert=false
            LogMessage "WARN","Verification","Failed to Click Button : Close Button" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickBtnClose_ReqSubmit_SMSAlert=true
End Function

Public Function selectCheckBox_SelectedCardList(lstlstCardList)
	'split the columns and values
	intSize = Ubound(lstlstCardList)
	ReDim arrCol(intSize)
	ReDim arrVal(intSize)
	'Fetch the total no of rows
	Set objAllRows=getAllRows(SMS_Thresholds.tblSMSAlertCardlistContent)
	intRow=objAllRows.Count  
	For colLoop = 0 To intSize-1 Step 1
		'Select the checkbox only if arrval() is True
		strColName = arrCol(colLoop)
		If arrval(colLoop) = "True" Then
			'write the function to select the checkbox
			selectChk = selectCheckBox(FeeWaiver.tblSelectedWaiverDetailsHeader,strColName)
			'write the function to read the value
			strCellVal=getCellTextFor(FeeWaiver.tblSelectedWaiverDetailsHeader,objAllRows(0),rowLoop,strColName)
			If strCellVal="" Then
				strCellVal = 0
			End If
			totalAmt = totalAmt + strCellVal
		End If		
	Next
End Function 

'[Select Check box for Single Card to Amend SMS Threshold values from Cardlist table]
Public Function selectCheckBoxForSingleCard_SMSAlert(lstlstCardlists)
   'This Function Selects Single Card
   intCount = Ubound(lstlstCardlists,1)
	ReDim arrCol(intCount)
	ReDim arrVal(intCount)
   iCount = 0
   Set objAllRows=getAllRows(SMS_Thresholds.tblSMSAlertCardlistContent)
   intRow = objAllRows.Count
   	verifytableSMS_SelectCard=verifyTableContentList(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,lstlstCardlists,"SelectCardList-SMSAlert",false,null,null,null)	
   	If verifytableSMS_SelectCard Then
   		For i = 0 To intCount	  
	   		strExpCardNum = lstlstCardlists(i,0)
	   		strExpCardNum = split(strExpCardNum,":")
	  		arrCol(i) = strExpCardNum(0)
			arrVal(i) = strExpCardNum(1)  			
			For j = 0 To intRow -1   
			  strActualCardNumber=getCellTextFor(SMS_Thresholds.tblSMSAlertCardlistHeader,objAllRows(j),j,"Account/Card No.")
				If StrComp(arrVal(i),strActualCardNumber) = 0 Then 
		   		  	selectChk = selectCheckBox_selectcardlist(objAllRows(j),j ,"")
		   		  	Exit For
	   			 End If 
	   		Next
   		Next
   	End If
End Function

Public Function selectCheckBox_selectcardlist(objTableContent,introw,strColName)
	selectCheckBox_selectcardlist = true
	Set oDesc = Description.Create
	oDesc("xpath").value ="//div[contains(@class,'dt-cell ng-scope')]"
	Set tableColObj = objTableContent.childobjects(oDesc)
	'check which childobject contains the class "csat-icon-checkbox ng-binding"
	Dim strColHeader
	strColHeader=tableColObj(0).GetROProperty("innertext")
	Set chkBox = Description.Create
	chkBox("xpath").value = "//div[contains(@class,'md-container md-ink-ripple')]"
	'Set chkBoxChildObj = tableColObj(0).childobjects(chkBox)
	Set chkBoxChildObj = tableColObj(introw).childobjects(chkBox)
	countChk = chkBoxChildObj.Count
	print countChk
	'chkBoxChildObj(0).click	
	chkBoxChildObj(introw).click	
End Function

Public Function selectCheckboxdisabled_ClosedStatus(lstlstCardlists)
   'This Function Selects Single Card
   bDevPending=False
   intCount = Ubound(lstlstCardlists,1)
	ReDim arrCol(intCount)
	ReDim arrVal(intCount)
   iCount = 0
   Set objAllRows=getAllRows(SMS_Thresholds.tblSMSAlertCardlistContent)
   intRow = objAllRows.Count
   	verifytableSMS_SelectCard=verifyTableContentList(SMS_Thresholds.tblSMSAlertCardlistHeader,SMS_Thresholds.tblSMSAlertCardlistContent,lstlstCardlists,"SelectCardList-SMSAlert",false,null,null,null)	
   	If verifytableSMS_SelectCard Then
   		For i = 0 To intCount	  
	   		strExpCardNum = lstlstCardlists(i,0)
	   		strExpCardNum = split(strExpCardNum,":")
	  		arrCol(i) = strExpCardNum(0)
			arrVal(i) = strExpCardNum(1)  		
		
			For j = 0 To intRow -1   
			  strActualCardNumber=getCellTextFor(SMS_Thresholds.tblSMSAlertCardlistHeader,objAllRows(j),j,"Card Number")
				If StrComp(arrVal(i),strActualCardNumber) = 0 Then 
		   		  	selectChk = selectCheckBox_selectcardlist(objAllRows(j),"")
		   		  	Exit For
	   			 End If 
	   		Next
   		Next
   	End If
End Function

'[Verify Field SMS Status displayed in the SMS Threshold Page]
Public Function VerifySMSStatus(StrStatus)
   bSMSstatus = True 
   WaitForIcallLoading
   If Not IsNull(StrStatus) Then
   	   Wait(20)
       If Not VerifyInnerText (SMS_Thresholds.lblSMSStatus(), StrStatus, "SMS Status") Then
           bSMSstatus=false
       End If
   End If
   VerifySMSStatus = bSMSstatus
End Function
