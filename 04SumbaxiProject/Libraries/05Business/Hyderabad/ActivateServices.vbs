'*****This is auto generated code using code generator please Re-validate ****************

'[Select Combobox Type of Activation as]
Public Function selectTypeofActivation(strTypeofActivation)
   bDevPending=false
   bSelectTypeofActivationComboBox=true
   If Not IsNull(strTypeofActivation) Then
       If Not (selectItem_Combobox (ServiceActivation.lstTypeofActivation(), strTypeofActivation))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Type of Activation drop down list" ,false
           bSelectTypeofActivationComboBox=false
       End If
   End If
   WaitForIcallLoading
   selectTypeofActivation=bSelectTypeofActivationComboBox
End Function

'[Verify Service Activation screen displayed]
Public Function verifyServiceActivationScreen()
	bverifyServiceActivationScreen=true
	If (ServiceActivation.tblSelectedCardsHeader().exist) Then
    	bverifyServiceActivationScreen=true
	Else
		LogMessage "RSLT","Verification","Failed to open service activation screen successfully.",False 
		bverifyServiceActivationScreen=false
	End If
	verifyServiceActivationScreen=bverifyServiceActivationScreen
End Function


'[Get selected item from combo box TypeofActivation]
Public Function getTypeofActivationSelectedItem()
   bDevPending=true
   getTypeofActivationSelectedItem=getVadinCombo_SelectedItem(ServiceActivation.lstTypeofActivation)
End Function

'[Verify Combobox Type of Activation displayed as]
Public Function verifyTypeofActivationText(strDefaultTypeofActivation)
   bDevPending=false
   bVerifyTypeofActivationText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (ServiceActivation.lstTypeofActivation(), strDefaultTypeofActivation, "Type of Activation")Then
           bVerifyTypeofActivationText=false
       End If
   End If
   verifyTypeofActivationText=bVerifyTypeofActivationText
End Function

'[Verify Table SelectedCards displayed]
Public Function verifySelectedCardsTabledisplayed()
   bDevPending=true
   verifySelectedCardsdisplayed= ServiceActivation.tblSelectedCards.Exist(1)
End Function
'[Verify Table SelectedCards has following Columns]
Public Function verifySelectedCardsTableColumns(arrColumnNameList)
   bDevPending=true
   verifySelectedCardsTableColumns=verifyTableColumns(ServiceActivation.tblSelectedCards,arrColumnNameList)
End Function
'[Verify row Data in Table SelectedCards for Service Activation]
Public Function verifytblSelectedCards_RowData_SA(arrRowDataList)
   bDevPending=false
   verifytblSelectedCards_RowData_SA=verifyTableContentList(ServiceActivation.tblSelectedCardsHeader,ServiceActivation.tblSelectedCardsContent,arrRowDataList,"SelectedCards" ,  false,null ,null,null)
End Function

'[Verify Field VerificationRequirement displayed as]
Public Function verifyVerificationRequirementText(strExpectedText)
   bDevPending=true
   bVerifyVerificationRequirementText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (ServiceActivation.lblVerificationRequirement(), strExpectedText, "VerificationRequirement")Then
           bVerifyVerificationRequirementText=false
       End If
   End If
   verifyVerificationRequirementText=bVerifyVerificationRequirementText
End Function

'[Set Check Box Verification Requirement as]
Public Function setCheckBoxVerificationRequirement(strONOFF)
   bDevPending=true
   If Ucase(strONOFF)="ON" Then
      ServiceActivation.cboxVerificationRequirement.Set("ON")
   Else
      ServiceActivation.cboxVerificationRequirement.Set("OFF")
   End If
   WaitForIcallLoading
   If Err.Number<>0 Then
       setCheckBoxVerificationRequirement=false
            LogMessage "WARN","Verification","Failed to Set Check Box : Verification Requirement" ,false
       Exit Function
   End If
   setCheckBoxVerificationRequirement=true
End Function


'[Verify Field CurrentStatus displayed as]
Public Function verifyCurrentStatusText(strExpectedText)
   bDevPending=false
   bVerifyCurrentStatusText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (ServiceActivation.lblCurrentStatus(), strExpectedText, "CurrentStatus")Then
           bVerifyCurrentStatusText=false
       End If
   End If
   verifyCurrentStatusText=bVerifyCurrentStatusText
End Function

'[Select Combobox NewStatus as]
Public Function selectNewStatusComboBox(strNewStatus)
   bDevPending=true
   bSelectNewStatusComboBox=true
   If Not IsNull(strNewStatus) Then
       If Not (selectItem_Combobox (ServiceActivation.lstNewStatus(), strNewStatus))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From NewStatus drop down list" ,false
           bSelectNewStatusComboBox=false
       End If
   End If
   WaitForIcallLoading
   selectNewStatusComboBox=bSelectNewStatusComboBox
End Function

'[Get selected item from combo box NewStatus]
Public Function getNewStatusSelectedItem()
   bDevPending=true
   getNewStatusSelectedItem=getVadinCombo_SelectedItem(ServiceActivation.lstNewStatus)
End Function

'[Verify Combobox NewStatus displayed as]
Public Function verifyNewStatusText(strExpectedText)
   bDevPending=true
   bVerifyNewStatusText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (ServiceActivation.lstNewStatus(), strExpectedText, "NewStatus")Then
           bVerifyNewStatusText=false
       End If
   End If
   verifyNewStatusText=bVerifyNewStatusText
End Function

'[Get StartDate Label Text]
Public Function getStartDateText()
   bDevPending=true
   getStartDateText=ServiceActivation.txtStartDate.GetRoProperty("innertext")
End Function

'[Verify Field StartDate displayed as]
Public Function verifyStartDateText(strStartDate)
   bDevPending=false
   bVerifyStartDateText=true
   strActualStartDate=ServiceActivation.txtStartDate.GetRoProperty("value")
   If Not IsNull(strStartDate) Then
	   If Ucase(strStartDate)="TODAY" Then
				If len(Day(CDate(Now)))=1 Then
					strDay="0"&Day(CDate(Now))
				else
					strDay=""&Day(CDate(Now))
				End If
				strStartDatePattern=""&strDay & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))&""
			else
				strStartDatePattern=strStartDate &""
			End If
		 If Matchstr(strActualStartDate,strStartDatePattern) Then
			LogMessage "RSLT","Verification","Created date pattern matched with expected pattern DD MMM YYYY",true
		else
			LogMessage "WARN","Verification","Created date pattern does not matched with Expected pattern DD MMM YYYY Expected: "&strStartDatePattern&" , Actual Date displayed is "&strActualStartDate,false
			bVerifyStartDateText=false
		End If
   End If
   verifyStartDateText=bVerifyStartDateText
End Function

'[Set TextBox StartDate to]
Public Function setStartDateTextbox(strStartDate)
   bDevPending=false
   setStartDateTextbox=true
   If Not IsNull (strStartDate) Then
     ServiceActivation.txtStartDate.Set(strStartDate)
   Else
     strActualDateTime = fGetText(strSession, "01", "071", "10")
     If len(Day(CDate(strActualDateTime)))=1 Then
	  strDay="0"&Day(CDate(strActualDateTime))
     else
	  strDay=""&Day(CDate(strActualDateTime))
	 End If	
      strStartDate=""&strDay & " "&monthName(Month(CDate(strActualDateTime)),true) &" " &Year(CDate(strActualDateTime))
      ServiceActivation.txtStartDate.Set(strStartDate)
      gstrRuntimeEffectiveDateStep="Set TextBox StartDate to"
      insertDataStore "SRStartDate", strStartDate
   End If
   
   If Err.Number<>0 Then
       setStartDateTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Start Date" ,false
       Exit Function
   End If
   setStartDateTextbox=true
End Function

'[Get EndDate Label Text]
Public Function getEndDateText()
   bDevPending=true
   getEndDateText=ServiceActivation.txtEndDate.GetRoProperty("innertext")
End Function

'[Verify Field EndDate displayed as]
Public Function verifyEndDateText(strExpectedText)
   bDevPending=false
   bVerifyEndDateText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( ServiceActivation.txtEndDate(), strExpectedText, "EndDate")Then
           bVerifyEndDateText=false
       End If
   End If
   verifyEndDateText=bVerifyEndDateText
End Function

'[Set TextBox EndDate to]
Public Function setEndDateTextbox(strEndDate)
   bDevPending=false
   setEndDateTextbox=true
   If Not IsNull (strEndDate) Then
     ServiceActivation.txtEndDate.Set(strEndDate)
   Else
     strActualDateTime = fGetText(strSession, "01", "071", "10")
     If len(Day(CDate(strActualDateTime)))=1 Then
	  strDay="0"&Day(CDate(strActualDateTime))
     else
	  strDay=""&Day(CDate(strActualDateTime))
	 End If	
      strEndDate=""&strDay & " "&monthName(Month(CDate(strActualDateTime)),true) &" " &Year(CDate(strActualDateTime))
      ServiceActivation.txtEndDate.Set(strEndDate)
    End If
   
   If Err.Number<>0 Then
       setEndDateTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :End Date" ,false
       Exit Function
   End If
   setEndDateTextbox=true
End Function

'[Verify Field Description displayed on Activate Services as]
Public Function verifyDescriptionText_SA(strExpectedText)
   bDevPending=false
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (ServiceActivation.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_SA=bVerifyDescriptionText
End Function

'[Get Comment Label Text]
Public Function getCommentText()
   bDevPending=true
   getCommentText=ServiceActivation.txtComment.GetRoProperty("innertext")
End Function

'[Verify Field Comment displayed as]
Public Function verifyCommentText(strExpectedText)
   bDevPending=true
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( ServiceActivation.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText=bVerifyCommentText
End Function


'[Set TextBox Comment on Service Activation to]
Public Function setCommentTextbox_SA(strComment)
   bDevPending=False
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comment on Service Activation to"
	insertDataStore "SRComment", strComment
	
   ServiceActivation.txtComment.Set(strComment )
   If Err.Number<>0 Then
       setCommentTextbox_SA=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_SA=true
End Function

'[Click Button AddNotes]
Public Function clickButtonAddNotes()
   bDevPending=true
   ServiceActivation.btnAddNotes.click
   If Err.Number<>0 Then
       clickButtonAddNotes=false
            LogMessage "WARN","Verification","Failed to Click Button : AddNotes" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonAddNotes=true
End Function

'[Perform Add Notes by clicking Add Notes Button on Service Activation Screen]
Public Function addNote_SA(strNote)
   bDevPending=false
   bVerifypopupNotes=true
	Dim bVerifypopupNotes:VerifypopupNotes=true
	
	If not isNull(strNote) Then
		ServiceActivation.btnAddNotes.click
		WaitForICallLoading
            If not   ServiceRequest.popupVerification.exist(5)Then
				LogMessage "WARN","Verification","New Note dialog did not displayed",false
				bVerifypopupNotes=false
			 else
			 strMessage=ServiceActivation.lblMaxAllowed.GetROProperty("innerText")
				If not strMessage="Max allowed - 3000" Then
					LogMessage "WARN","Verification","Add New Comment popup dislog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage,false
					bVerifypopupNotes=false
				End If
			   ServiceRequest.txtNewComment.set strNote
			  
				   ServiceRequest.clickSave_Popup
				  WaitForIcallLoading
		   End If 
		End If 
	addNote_SA=bVerifypopupNotes
End Function

'[Verify Popup ValidationMessage exist]
Public Function verifyPopupValidationMessageexist(bExist)
   bDevPending=true
   bActualExist=strTestAppFrameClass.Exist()
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage Exists As Expected" ,true
       verifyPopupValidationMessageexist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage does not Exists As Expected" ,true
       verifyPopupValidationMessageexist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage does not Exists As Expected" ,False
       verifyPopupValidationMessageexist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage Still Exists" ,False
       verifyPopupValidationMessageexist=False
   End If
End Function

'[Click Button OK_ValidationPopup]
Public Function clickButtonOK_ValidationPopup()
   bDevPending=true
   ServiceActivation.btnOK_ValidationPopup.click
   If Err.Number<>0 Then
       clickButtonOK_ValidationPopup=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationPopup" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonOK_ValidationPopup=true
End Function

'[Get ValidationMessage Label Text]
Public Function getValidationMessageText()
   bDevPending=true
   getValidationMessageText=ServiceActivation.lblValidationMessage.GetRoProperty("innertext")
End Function

'[Verify Field ValidationMessage displayed as]
Public Function verifyValidationMessageText(strExpectedText)
   bDevPending=true
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (ServiceActivation.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
           bVerifyValidationMessageText=false
       End If
   End If
   verifyValidationMessageText=bVerifyValidationMessageText
End Function

'[Verify Validation Message displayed as]
Public Function verifyValidationMessage_SA(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (ServiceActivation.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
   End If
   ServiceActivation.btnOK_ValidationPopup.Click
   WaitForIcallLoading
   verifyValidationMessage_SA=bVerifyValidationMessageText
End Function

'[Verify Button Submit is enabled on Service Activation Screen]
Public Function VerifybtnSubmit_SA()
	bDevPending=False
   Dim bVerifybtnSubmit_CC:bVerifybtnSubmit_CC=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	intBtnSubmit=Instr(ServiceActivation.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
	If  intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Submit button is enable as per expectation.",True
		bVerifybtnSubmit_CC=true
	Else
		LogMessage "WARN","Verifiation","Submit button is disable. Expected to be enable.",false
		bVerifybtnSubmit_CC=false
	End If
	VerifybtnSubmit_SA=bVerifybtnSubmit_CC
End Function

'[Click Button Submit for service activation]
Public Function clickButtonSubmit_ServActivation()
   bDevPending=true
   ServiceActivation.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_ServActivation=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
    WaitForIcallLoading
   clickButtonSubmit_ServActivation=true
End Function

'[Click Button Cancel on Service Activation]
Public Function clickButtonCancel_SA()
   bDevPending=true
   ServiceActivation.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel=true
End Function

'[Verify Field TM Approval Message displayed as]
Public Function verifyTMApprovalMessage_SA(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (ServiceActivation.lblConfirmationMsg(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
   End If
   ServiceActivation.btnOK_Confirmation.Click
   WaitForIcallLoading
   verifyTMApprovalMessage_SA=bVerifyValidationMessageText
End Function

'[Verify Tab Services Activation is displayed]
Public Function verifyTabActivationExist_SA()
   bDevPending=false
   verifyTabActivationExist_SA=verifyTabExist("Services Activation")
End Function

'[Select Radio Button Validity on Service Activation]
Public Function selectValidityRadio(strValidity)
	bDevPending=False
	bselectValidityRadio=true
	bselectValidityRadio=SelectRadioButtonGrp(strValidity, ServiceActivation.rbtnValidity, Array("Temporary","Permanent"))   
	If Err.Number<>0 Then
       bselectValidityRadio=false
          LogMessage "WARN","Verification","Failed to Click Button : Validity on Service Activation" ,false
       Exit Function
   End If
   selectValidityRadio=bselectValidityRadio
End Function

'[Click Button Submit on Service Activation Screen]
Public Function clickButtonSubmit_SA()
   bDevPending=false
   ServiceActivation.btnSubmit.click
   If Err.Number<>0 Then
       clickButtonSubmit_SA=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonSubmit_SA=true
End Function

'[Click Button Cancel on Service Activation Screen]
Public Function clickButtonCancel_SA()
   bDevPending=true
   ServiceActivation.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_SA=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel_SA=true
End Function

'[Verify Popup Request Submitted exist for Service Activation]
Public Function verifyPopupRequestSubmitted_SA(bExist)
   bDevPending=false
   bActualExist=ServiceActivation.popupRequestSubmitted.Exist(4)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupRequestSubmitted_SA=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupRequestSubmitted_SA=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupRequestSubmitted_SA=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupRequestSubmitted_SA=False
   End If
End Function

'[Verify Field CardNumber on Request Submitted Popup for Service Activation displayed as]
Public Function verifyCardNumber_RequestSubmitted_SA(strCardNumber)
   bDevPending=false
   bVerifyCardNumber_RequestSubmittedText=true
   insertDataStore "NewSAUsedCard", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       If Not VerifyInnerText (ServiceActivation.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber_RequestSubmittedText=false
       End If
   End If
   verifyCardNumber_RequestSubmitted_SA=bVerifyCardNumber_RequestSubmittedText
End Function

'[Verify Field ProductDescription on Request Submitted Popup for Service Activation displayed as]
Public Function verifyProductDescription_RequestSubmitted_SA(strProductDescription)
   bDevPending=false
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (ServiceActivation.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_SA=bVerifyProductDescription_RequestSubmittedText
End Function

'[Click Link SRNumber on Request Submitted Popup for Service Activation]
Public Function clickLinkSRNumber_RequestSubmitted_SA()
   bDevPending=false
   gstrRuntimeSRNumStep="Click Link SRNumber on Request Submitted Popup for Service Activation"
   strSelectedSR=ServiceActivation.lnkSRNumber_RequestSubmitted.GetRoProperty("innerText")
	If strSelectedSR<>"" Then
		 insertDataStore "SelectedSRLink", strSelectedSR
	
	   ServiceActivation.lnkSRNumber_RequestSubmitted.click
	 else
   		LogMessage "RSLT","Verification","SR Number did not displayed on Request Submitted pop up",false
	End If
   WaitForIcallLoading
   If Err.Number<>0 Then
       clickLinkSRNumber_RequestSubmitted_SA=false
            LogMessage "WARN","Verification","Failed to Click Link : SRNumber_RequestSubmitted" ,false
       Exit Function
   End If
   clickLinkSRNumber_RequestSubmitted_SA=true
End Function

'[Verify Field Status_RequestSubmitted For Service Activation displayed as]
Public Function verifyStatus_RequestSubmitted_SA(strExpectedText)
   bDevPending=false
   bVerifyStatus_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (ServiceActivation.lblStatus_RequestSubmitted(), strExpectedText, "Status_RequestSubmitted")Then
           bVerifyStatus_RequestSubmittedText=false
       End If
   End If
   verifyStatus_RequestSubmitted_SA=bVerifyStatus_RequestSubmittedText
End Function

'[Click Button RefreshStatus For Service Activation]
Public Function clickButtonRefreshStatus_SA()
   bDevPending=false
   ServiceActivation.btnRefreshStatus.click
	WaitForICallLoading
    		'Get Status
		If ServiceActivation.lblStatus_RequestSubmitted.getROProperty("innertext")="In Progress" then 
			bStatus=true
		 else
			bStatus=false
		End If		
	
	While  bStatus AND (iCount<60)
		ServiceActivation.btnRefreshStatus.click
		wait 1
        	'Get Status
			strStatus=ServiceActivation.lblStatus_RequestSubmitted.getROProperty("innertext")
			If Trim(strStatus)="In Progress" then 
				bStatus=true
			 else
				LogMessage "WARN","Verification","Status displayed as  :"&strStatus ,true
				bStatus=false
			End If
		wait 5
		intBtnRefreshStatus=Instr(ServiceActivation.btnRefreshStatus.GetROproperty("outerhtml"),"v-disabled")
		If intBtnRefreshStatus<>0 Then
			LogMessage "WARN","Verification","Button : RefreshStatus is disabled" ,true
			bStatust=true
		End If
		iCount=iCount+1
	  Wend	

   If Err.Number<>0 Then
       
            LogMessage "WARN","Verification","Failed to Click Button : RefreshStatus" ,false
			clickButtonRefreshStatus_SA=false
       Exit Function
   End If
   WaitForICallLoading
  
   clickButtonRefreshStatus_SA=true
End Function

'[Click Button Cancel_Request Submitted for Service Activation]
Public Function clickButtonCancel_RequestSubmittedSA()
   bDevPending=false
   ServiceActivation.btnCancel_RequestSubmitted.click
   waitForIcallLoading
   If Err.Number<>0 Then
       clickButtonCancel_RequestSubmittedSA=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel_RequestSubmitted" ,false
       Exit Function
   End If
   clickButtonCancel_RequestSubmittedSA=true
End Function

'[Verify Link SRNumber available on Request Submitted popup for Service Activation Screen]
Public Function verifyLinkSRNumber_RequestSubmitted_SA()
   bDevPending=False
   bverifyLinkSRNumber_RequestSubmitted=true
	strSelectedSR=ServiceActivation.lnkSRNumber_RequestSubmitted.GetRoProperty("innerText")
	insertDataStore "SelectedSRLink", strSelectedSR
	If instr(ServiceActivation.lnkSRNumber_RequestSubmitted.GetRoProperty("class"),"link")=0 Then
		bverifyLinkSRNumber_RequestSubmitted=false
	else
		bverifyLinkSRNumber_RequestSubmitted=true
	end If
	LogMessage "RSLT","Verification","SR Number link "& strSelectedSR &" displayed on Request Submitted popup",true
	If IsNull(strSRNumber) Then
		LogMessage "WARN","Verification", "SR Number not available with link on Request Submitted popup.",false
		bverifyLinkSRNumber_RequestSubmitted=false
	End If

   verifyLinkSRNumber_RequestSubmitted_SA=bverifyLinkSRNumber_RequestSubmitted
End Function

'[Perform Additional Verification for TPIN call from Service Activation]
Public Function performTPINPlusOne_ServiceActivation(bExpectedVerification, strValidationMessage)
   Dim bPerformTPINPlusOneVerification:bPerformTPINPlusOneVerification=true
   bActualExist=ServiceActivation.popupValidationMessage.Exist(1)
	'bActualExist=true	
   If bExpectedVerification Then
		If not bActualExist Then
			bPerformTPINPlusOneVerification=false
		 else
			strActualValidationMessage=ServiceActivation.lblValidationMessage.GetRoProperty("innertext")
			If Not IsNull(strActualValidationMessage) Then
				If Not VerifyInnerText (ServiceActivation.lblValidationMessage(), strValidationMessage, "ValidationMessage")Then
					bPerformTPINPlusOneVerification=false
				End If
			End If
			ServiceActivation.btnOK_ValidationPopup.Click
			VerifyCustomer.btnVerify.Click
			waitForIcallLoading
			
			'Check Required Verification met Check box
			Set oDesc= Description.Create()
			oDesc("micclass").Value="WebCheckBox"
			'CardActivation.cbTPINVerification.ChildObjects(oDesc)(0).Set "ON" '******** Check box removed
			CardActivation.rbAdditionalAnswer.click
			'Wait 1
			VerifyCustomer.btnVerifyCustomer().click
			waitForIcallLoading
			If err.number<>0 Then
				bPerformTPINPlusOneVerification=false
			End If
		End If	 
   End If
   performTPINPlusOne_ServiceActivation=bPerformTPINPlusOneVerification
End Function

