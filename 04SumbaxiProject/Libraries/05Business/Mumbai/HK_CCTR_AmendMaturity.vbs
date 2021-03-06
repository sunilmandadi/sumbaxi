'[Click on Amend Maturity Action link]
Public Function clickLinkAmendMaturityAction()
	ClickOnObject HK_CCTR_AmendMaturity_Page.lnkAmendMaturityAction(),"Amend Maturity Action link"
	WaitForICallLoading
End Function
'[Click on Update MI link]
Public Function clickLinkUpdateMI()
	ClickOnObject HK_CCTR_AmendMaturity_Page.lnkUpdateMI(),"Update MI link"
	WaitForICallLoading
End Function
'[Verify dialog verification required displayed or not]
Public Function verifyVerificationDialogBox()
	blnverifyVerificationDialogBox=True
	If Not(HK_CCTR_AmendMaturity_Page.dlgVerification().Exist(10)) Then
		blnverifyVerificationDialogBox=false
		LogMessage "WARN","Verification","Verification Dialogue Box is not displayed" ,false
	Else
		LogMessage "RSLT","Verification","Verification Dialogue Box is displayed as expected",True
		blnverifyVerificationDialogBox=true
	End If
	verifyVerificationDialogBox=blnverifyVerificationDialogBox
End Function
'[Click on Verification Ok Button]
Public Function clickBtnVerificationOk()
	ClickOnObject HK_CCTR_AmendMaturity_Page.btnVerificationOk(),"Verification Ok Button"
	WaitForICallLoading
End Function
'[Click on Not Verified Button]
Public Function clickBtnNotVerified()
	ClickOnObject HK_CCTR_AmendMaturity_Page.btnNotVerified(),"Not Verified Button"
	WaitForICallLoading
End Function
'[Verify customer by answering required level of Identification questions]
Public Function SelectManualAuthenticationQuestions(intNoOfIdentificationQuestions)
    blnSelectManualAuthenticationQuestions = True
   
    ' Check if the Manual Identification and Authentican required
    If intNoOfIdentificationQuestions = "0" Then
        LogMessage "RSLT","Verification","No manual Identification required or demanded",True
        blnSelectManualAuthenticationQuestions = True
        Exit Function
    End If
    ' Answer the questions if the required Idendification is greater than zero
    If intNoOfIdentificationQuestions <> "0" Then
        Set oDesc = Description.Create()
        oDesc("html id").Value = "vrfyCustIdenQ_div.*"
        Set questionCollection = gObjIServePage.ChildObjects(oDesc)
        NumberOfQues = questionCollection.Count
        If NumberOfQues <> "0" Then
            For i  = 0 To intNoOfIdentificationQuestions - 1 Step 1
               Set objQuesRow=questionCollection(i)        
               Set oDescRadio = Description.Create()
               oDescRadio("class").Value = "md-radio-button"
               Set RadioCollection = objQuesRow.ChildObjects(oDescRadio)
               selectRadioGroup RadioCollection,"Pass",Array("Pass","Fail")
            Next
         Else
         	blnSelectManualAuthenticationQuestions=False
       End If
    End If
    
SelectManualAuthenticationQuestions=blnSelectManualAuthenticationQuestions

End Function
'[Verify customer by answering required level of Customer Portfolio questions]
Public Function SelectCustomerPortfolioQuestions(intNoOfCustomerPortfolioQuestions)
    blnSelectManualAuthenticationQuestions = True
   	WaitForICallLoading
    ' Check if the Manual Identification and Authentican required
    If intNoOfCustomerPortfolioQuestions = "0" Then
        LogMessage "RSLT","Verification","No manual Identification required or demanded",True
        blnSelectManualAuthenticationQuestions = True
        Exit Function
    End If
    ' Answer the questions if the required Idendification is greater than zero
    If intNoOfCustomerPortfolioQuestions <> "0" Then
        Set oDesc = Description.Create()
        oDesc("html id").Value = "vrfyCustCustPortQ_div.*"
        Set questionCollection = gObjIServePage.ChildObjects(oDesc)
        NumberOfQues = questionCollection.Count
        If NumberOfQues <> "0" Then
            For i  = 0 To intNoOfCustomerPortfolioQuestions - 1 Step 1
                Set objQuesRow=questionCollection(i)        
               Set oDescRadio = Description.Create()
               oDescRadio("class").Value = "md-radio-button"
               Set RadioCollection = objQuesRow.ChildObjects(oDescRadio)
               selectRadioGroup RadioCollection,"Pass",Array("Pass","Fail")
            Next
         Else
         	blnSelectManualAuthenticationQuestions=False
       End If
    End If
    
SelectCustomerPortfolioQuestions=blnSelectManualAuthenticationQuestions

End Function
'[Click on Verify Customer Button]
Public Function clickBtnVerifyCustomer()
	ClickOnObject HK_CCTR_AmendMaturity_Page.btnVerifyCustomer(),"Verify Customer Button"
	WaitForICallLoading
	WaitForICallLoading
End Function
'[Verify row Data in Table for TD Account No in Udate MI Page]
Public Function verifytblContentTDAccountUpdateMI(arrKeyInfoRowDataList)
	wait(10)
   verifytblContentTDAccountUpdateMI=verifyTableContentList(HK_CCTR_AmendMaturity_Page.tblTDaccountNoHeader(),HK_CCTR_AmendMaturity_Page.tblTDaccountNoContent(),arrKeyInfoRowDataList,"Update MI",false,NULL,NULL,NULL)
End Function
'[Verify row Data in Table for Deposit Number in Udate MI Page]
Public Function verifytblContentTDDepositNumberUpdateMI(arrKeyInfoRowDataList)
	wait(10)
   verifytblContentTDDepositNumberUpdateMI=verifyTableContentList(HK_CCTR_AmendMaturity_Page.tblTDDepositNumberHeader(),HK_CCTR_AmendMaturity_Page.tblTDDepositNumberContent(),arrKeyInfoRowDataList,"Update MI",false,NULL,NULL,NULL)
End Function
'[Verify row Data in Table for Maturity Instruction in Udate MI Page]
Public Function verifytblContentTDMaturityInstructionUpdateMI(arrKeyInfoRowDataList)
	wait(10)
   verifytblContentTDMaturityInstructionUpdateMI=verifyTableContentList(HK_CCTR_AmendMaturity_Page.tblTDMaturityInstructionHeader(),HK_CCTR_AmendMaturity_Page.tblTDMaturityInstructionContent(),arrKeyInfoRowDataList,"Update MI",false,NULL,NULL,NULL)
End Function
'[Select Combobox New Maturity Instruction]
Public Function selectNewMaturityInstructionComboBox(strNewMaturityInstruction)
	blnselectNewMaturityInstructionComboBox=true
	If Not IsNull(strNewMaturityInstruction) Then
		If Not (selectItem_Combobox (HK_CCTR_AmendMaturity_Page.lstNewMaturityInstruction(),strNewMaturityInstruction))Then
			LogMessage "WARN","Verification","Failed to select :"&strNewMaturityInstruction&" From Show drop down list" ,false
			blnselectNewMaturityInstructionComboBox=false
		End If
	End If
	WaitForICallLoading
	selectNewMaturityInstructionComboBox=blnselectNewMaturityInstructionComboBox
End Function
'[Select Combobox New Next Tenor Days]
Public Function selectNewNextTenorDaysComboBox(strNewNextTenorDays)
	blnselectNewNextTenorDaysComboBox=true
	If Not IsNull(strNewNextTenorDays) Then
		If Not (selectItem_Combobox (HK_CCTR_AmendMaturity_Page.lstNewNextTenorDays(),strNewNextTenorDays))Then
			LogMessage "WARN","Verification","Failed to select :" & strNewNextTenorDays & " From Show drop down list" ,false
			blnselectNewNextTenorDaysComboBox=false
		End If
	End If
	WaitForICallLoading
	selectNewNextTenorDaysComboBox=blnselectNewNextTenorDaysComboBox
End Function
'[Select Combobox New Next Tenor Days Count]
Public Function selectNewNextTenorDaysCountComboBox(strNewNextTenorDaysCount)
	blnselectNewNextTenorDaysCountComboBox=true
	If Not IsNull(strNewNextTenorDaysCount) Then
		If Not (selectItem_Combobox (HK_CCTR_AmendMaturity_Page.lstNewNextTenorDaysCount(),strNewNextTenorDaysCount))Then
			LogMessage "WARN","Verification","Failed to select :"&strNewNextTenorDaysCount&" From Show drop down list" ,false
			blnselectNewNextTenorDaysCountComboBox=false
		End If
	End If
	WaitForICallLoading
	selectNewNextTenorDaysCountComboBox=blnselectNewNextTenorDaysCountComboBox
End Function
'[Verify Next button is enabled]
Public Function verifyNextBtn_Enable()
	VerifyObjectEnabledDisabled HK_CCTR_AmendMaturity_Page.btnNext(),"Enable","Next link"
End Function
'[Click on Next Button]
Public Function clickNextButton()
	ClickOnObject HK_CCTR_AmendMaturity_Page.btnNext(),"Next Button"
	WaitForICallLoading
End Function
'[Click on Proceed Button]
Public Function clickProceedButton()
	ClickOnObject HK_CCTR_AmendMaturity_Page.btnProceed(),"Proceed Button"
	WaitForICallLoading
End Function
'[Verify Update Confirmation dialog displayed]
Public Function verifyUdateConfirmationDialogue()
	blnverifyUdateConfirmationDialogue=True
	If Not(HK_CCTR_AmendMaturity_Page.dlgUpdateConfirmation().Exist(10)) Then
		blnverifyUdateConfirmationDialogue=false
		LogMessage "WARN","Verification","Update Confirmation Dialogue Box is not displayed" ,false
	Else
		LogMessage "RSLT","Verification","Update Confirmation Dialogue Box is displayed as expected",True
		blnverifyUdateConfirmationDialogue=true
	End If
	verifyUdateConfirmationDialogue=blnverifyUdateConfirmationDialogue
End Function
'[Click on Ok Button in Update MI]
Public Function clickOkButtonUpdateMI()
	ClickOnObject HK_CCTR_AmendMaturity_Page.btnOkUpdateMI(),"Ok Button"
	WaitForICallLoading
End Function
'[Click on Confirmation Yes Button in Update MI]
Public Function clickConfirmationYesButtonUpdateMI()
	blnclickConfirmationYesButtonUpdateMI=true
	
	HK_CCTR_AmendMaturity_Page.btnConfirmationYesUpdateMI().Click
	
	If Err.Number<>0 Then
		blnclickConfirmationYesButtonUpdateMI=false
		LogMessage "WARN","Verification","Failed to Click Button :Confirmation - Yes" ,false
	Else
		LogMessage "RSLT","Verification","Clicked on Confirmation - Yes Button as expected.",True
		blnclickConfirmationYesButtonUpdateMI=true
	End If
	clickConfirmationYesButtonUpdateMI=blnclickConfirmationYesButtonUpdateMI
End Function
