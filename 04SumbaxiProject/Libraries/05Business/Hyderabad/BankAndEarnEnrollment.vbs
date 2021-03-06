'[Verify and Click Bank & Earn Enrollment Link from leftMenu]
Public Function ClickLink_BankAndEarnEnrollment()
bClickLink_BankAndEarnEnrollment=true
	bcAccountOverview_LeftMenu.btnBankAndEarnEnrollment.Click
	WaitForIcallLoading
	If Err.Number<>0 Then
       bClickLink_BankAndEarnEnrollment=false
       LogMessage "WARN","Verification","Failed to Click Link  : Bank and Earn Enrollment" ,false
       Exit Function
	End If
	Wait 1
	waitForIcallLoading	
ClickLink_BankAndEarnEnrollment = bClickLink_BankAndEarnEnrollment
End Function 

''[Click Rewards link on Customer Overview Screen]
'Public Function clicklink_Rewards()
'	bDevPending=false
'   bcCustomerOverview.lnkRewards.click
'   If Err.Number<>0 Then
'       clicklink_Rewards=false
'            LogMessage "WARN","Verification","Failed to Click Link : Rewards" ,false
'       Exit Function
'   End If
'   WaitForICallLoading
'   clicklink_Rewards=true
'End Function

'[Verify Pink Panel displayed in Rewards Page]
Public Function verifyPinkPanel_Rewards(strName,strCIN,strSegment)
bverifyPinkPanel_Rewards=true
	
	If Not IsNull(strName) Then
		If Not VerifyInnerText(ServiceRequest.lblName(),strName,"Name") Then
	bverifyPinkPanel_Rewards = False
	End If
End If

	If Not IsNull(StrCIN) Then
		If Not VerifyInnerText(ServiceRequest.lblCIN(),strCIN,"CIN") Then
bverifyPinkPanel_Rewards = False
	End If
End If
	
	If Not IsNull(StrSegment) Then
		If Not VerifyInnerText(ServiceRequest.lblSegment(),strSegment,"Segment") Then
bverifyPinkPanel_Rewards = False
	End If
End If
verifyPinkPanel_Rewards=bverifyPinkPanel_Rewards
End Function

'[Select value for Program from the combolist]
Public Function SelectProgramList(strProgram)
   bDevPending=false
   bSelectProgramList=true
   BankAndEarnEnrollment.txtProgram.Set strProgram
   WaitForIcallLoading
   SelectProgramList=bSelectProgramList
End Function

'[Verify the field under Bank and Earn Details Section]
Public Function verifyLoyaltyProgramFields(strProgram,strProgReferenceNo,strRewardType,strYTDEarnings)
	
	'Get all Iserve fields under Program section
	strIserveProgram = BankAndEarnEnrollment.lblProgram.GetROProperty("innertext")
	strIserveProgReferenceNo = BankAndEarnEnrollment.lblProgramReferenceNo.GetROProperty("innertext")
	strIserveRewardType = BankAndEarnEnrollment.lblRewardType.GetROProperty("innertext")
	strIserveYTDEarnings = BankAndEarnEnrollment.lblYTDEarnings.GetROProperty("innertext")
	
   bDevPending=true
   bverifyLoyaltyProgramFields=true
   
  	 If strProgram = strIserveProgram Then
   	
		  	LogMessage "RSLT","Verification","The value for Program is as expected: "&strProgram&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Program is not as expected: "&strProgram&"",False
		End if 
		
	If strProgReferenceNo = strIserveProgReferenceNo Then
   	
		  	LogMessage "RSLT","Verification","The value for Program Reference No. is as expected: "&strProgReferenceNo&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Program Reference No is not as expected: "&strProgReferenceNo&"",False
		End if 
			
	If strRewardType = strIserveRewardType Then
   	
		  	LogMessage "RSLT","Verification","The value for Reward Type is as expected: "&strRewardType&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Reward Type  is not as expected: "&strRewardType&"",False
		End if		
  	
  	If strYTDEarnings = strIserveYTDEarnings Then
   	
		  	LogMessage "RSLT","Verification","The value for YTD Earnings is as expected: "&strYTDEarnings&"",True
				Else
		  	LogMessage "WARN","Verification","The value for YTD Earnings is not as expected: "&strYTDEarnings&"",False
		End if
  	verifyLoyaltyProgramFields = bverifyLoyaltyProgramFields
End Function

'[Verify the field under Enrolment Details section]
Public Function verifyEnrolmentDetailsFields(strEnrolmentDate,strChannel,strBranch,strCreatedBy,strStatementOption)
	
	'Get all Iserve fields under Program section
	strIserveEnrolmentDate = BankAndEarnEnrollment.lblEnrolmentDate.GetROProperty("innertext")
	strIserveChannel = BankAndEarnEnrollment.lblChannel.GetROProperty("innertext")
	strIserveBranch = BankAndEarnEnrollment.lblBranch.GetROProperty("innertext")
	strIserveCreatedBy = BankAndEarnEnrollment.lblCreatedBy.GetROProperty("innertext")
	strIserveStatementOption = BankAndEarnEnrollment.lblStatementOption.GetROProperty("innertext")
	
	bDevPending=true
   bverifyEnrolmentDetailsFields=true
   
  	 If strEnrolmentDate = strIserveEnrolmentDate Then
   	
		  	LogMessage "RSLT","Verification","The value for Enrolment Date is as expected: "&strEnrolmentDate&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Enrolment Date is not as expected: "&strEnrolmentDate&"",False
		End if 
		
	If strChannel = strIserveChannel Then
   	
		  	LogMessage "RSLT","Verification","The value for Channel is as expected: "&strChannel&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Channel is not as expected: "&strChannel&"",False
		End if 
			
	If strCreatedBy = strIserveCreatedBy Then
   	
		  	LogMessage "RSLT","Verification","The value for Created By is as expected: "&strCreatedBy&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Created By is not as expected: "&strCreatedBy&"",False
		End if		
  	
  	If strStatementOption = strIserveStatementOption Then
   	
		  	LogMessage "RSLT","Verification","The value for statement option is as expected: "&strStatementOption&"",True
				Else
		  	LogMessage "WARN","Verification","The value for statement option is not as expected: "&strStatementOption&"",False
		End if
  	verifyEnrolmentDetailsFields = bverifyEnrolmentDetailsFields
End Function

'[Verify the field under Crediting Account Details Section]
Public Function verifyCreditingAccDetailFields(strCredAccNo,strProduct,strAccInd,strMailingAddress)
	
	'Get all Iserve fields under Program section
	strIserveCredAccNo = BankAndEarnEnrollment.lblCreditingAccNo.GetROProperty("innertext")
	strIserveProduct = BankAndEarnEnrollment.lblProduct.GetROProperty("innertext")
	strIserveAccInd = BankAndEarnEnrollment.lblAccCardInd.GetROProperty("innertext")
	strIserveMailingAddress = BankAndEarnEnrollment.lblMailingAddress.GetROProperty("innertext")
	
	bDevPending=true
    bverifyCreditingAccDetailFields=true
   
  	 If strCredAccNo = strIserveCredAccNo Then
   	
		  	LogMessage "RSLT","Verification","The value for Crediting Acc No. is as expected: "&strCredAccNo&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Crediting Acc No. is not as expected: "&strCredAccNo&"",False
		End if 
		
	If strProduct = strIserveProduct Then
   	
		  	LogMessage "RSLT","Verification","The value for Product is as expected: "&strProduct&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Product is not as expected: "&strProduct&"",False
		End if 
			
	If strAccInd = strIserveAccInd Then
   	
		  	LogMessage "RSLT","Verification","The value for Acc/Card Indicator is as expected: "&strAccInd&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Acc/Card Indicator is not as expected: "&strAccInd&"",False
		End if		
  	
  	If not isnull(strMailingAddress) Then
  		If strMailingAddress = strIserveMailingAddress Then
		  	LogMessage "RSLT","Verification","The value for Mailing Address is as expected: "&strMailingAddress&"",True
				Else
		  	LogMessage "WARN","Verification","The value for Mailing Address is not as expected: "&strMailingAddress&"",False
		End if
  		else
  		LogMessage "RSLT","Verification","The value for Mailing Address is as expected: "&strMailingAddress&"",True
  	End If 	
  	verifyCreditingAccDetailFields = bverifyCreditingAccDetailFields
End Function

'[Verify the field under DeEnrolment Details Section]
Public Function verifyDeEnrolmentDetailFields(strDeEnrolmentBranch,strDeEnrolmentChannel,strDeEnrolmentCreatedBy,strDeEnrolmentDate)
	
	'Get all Iserve fields under Program section
	strIserveDeEnrolmentBranch = BankAndEarnEnrollment.lblDeEnrolmentBranch.GetROProperty("innertext")
	strIserveDeEnrolmentChannel = BankAndEarnEnrollment.lblDeEnrolmentChannel.GetROProperty("innertext")
	strIserveDeEnrolmentCreatedBy = BankAndEarnEnrollment.lblDeEnrolmentCreatedBy.GetROProperty("innertext")
	strIserveDeEnrolmentDate = BankAndEarnEnrollment.lblDeEnrolmentDate.GetROProperty("innertext")
	'bDevPending=true
    bverifyDeEnrolmentDetailFields=true
    
     If Not IsNull(strDeEnrolmentBranch) Then
     If strDeEnrolmentBranch = strIserveDeEnrolmentBranch Then
		  	LogMessage "RSLT","Verification","The value for De-enrolment Branch is as expected: "&strDeEnrolmentBranch&"",True
				Else
		  	LogMessage "WARN","Verification","The value for De-enrolment Branch is not as expected: "&strDeEnrolmentBranch&"",False
		      End if 
     	    Else
  	 			LogMessage "RSLT","Verification","The value for De-enrolment Branch is as expected: "&strDeEnrolmentBranch&"",True
		 End If
		 
	If strDeEnrolmentChannel = strIserveDeEnrolmentChannel Then
   	
		  	LogMessage "RSLT","Verification","The value for De-enrolment Channel is as expected: "&strDeEnrolmentChannel&"",True
				Else
		  	LogMessage "WARN","Verification","The value for De-enrolment Channel is not as expected: "&strDeEnrolmentChannel&"",False
		End if 
			
	If strDeEnrolmentCreatedBy = strIserveDeEnrolmentCreatedBy Then
   	
		  	LogMessage "RSLT","Verification","The value for De-enrolment Created By is as expected: "&strDeEnrolmentCreatedBy&"",True
				Else
		  	LogMessage "WARN","Verification","The value for De-enrolment Created By is not as expected: "&strDeEnrolmentCreatedBy&"",False
		End if		
  	
  	If strDeEnrolmentDate = strIserveDeEnrolmentDate Then
   	
		  	LogMessage "RSLT","Verification","The value for De-enrolment Date is as expected: "&strDeEnrolmentDate&"",True
				Else
		  	LogMessage "WARN","Verification","The value for De-enrolment Date is not as expected: "&strDeEnrolmentDate&"",False
		End if
  	verifyDeEnrolmentDetailFields = bverifyDeEnrolmentDetailFields
End Function

'[Verify Shortcut Button on Bank and Earn Enrollment Page]
Public Function VerifyShortcutButton_BankAndEarnEnrollment()
	 bDevPending=False
	 bVerifyShortcutButton_BankAndEarnEnrollment = true
	 
	 If (BankAndEarnEnrollment.btnAmendment().exist) Then
	 	LogMessage "RSLT","Verification","The Shortcut button Amendment exist",True
	 	else
	 	LogMessage "RSLT","Verification","The Shortcut button Amendment does not exist",False
	 End If
	 
	 If (BankAndEarnEnrollment.btnDeEnrolment().exist) Then
	 	LogMessage "RSLT","Verification","The Shortcut button De-Enrolment exist",True
	 	else
	 	LogMessage "RSLT","Verification","The Shortcut button De-Enrolment does not exist",False
	 End If
	 
	 If (BankAndEarnEnrollment.btnEnrolment().exist) Then
	 	LogMessage "RSLT","Verification","The Shortcut button Enrolment exist",True
	 	else
	 	LogMessage "RSLT","Verification","The Shortcut button Enrolment does not exist",False
	 End If
	  VerifyShortcutButton_BankAndEarnEnrollment= bVerifyShortcutButton_BankAndEarnEnrollment
End Function
