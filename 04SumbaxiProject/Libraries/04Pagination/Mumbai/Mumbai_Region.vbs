Dim HK_CCTR_AmendMaturity_Page

Set HK_CCTR_AmendMaturity_Page = cHKAmendMaturity()

Public Function cHKAmendMaturity()
	'Dim cHKLoginIServe
	Set cHKAmendMaturity = New clsHKAmendMaturity
End Function

Class clsHKAmendMaturity

	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	'******************************** Object Initialization ******************************************************************
	
	Public Function lnkAmendMaturityAction()
		Set lnkAmendMaturityAction=gObjIServePage.WebElement("xpath:=//*[@id='_content']/div/div/div/div/div[7]/div/div/md-menu/span/i")
	End Function
	
	Public Function lnkUpdateMI()
		Set lnkUpdateMI=gObjIServePage.WebElement("xpath:=//A[@id='updateTDSTP_link']")
	End Function
	
	Public Function dlgVerification()
		Set dlgVerification=gObjIServePage.WebElement("xpath:=//MD-DIALOG[@id='maVerify_popup']")
	End Function
	
	Public Function eleDlgVerificationMsg()
		Set eleDlgVerificationMsg = gObjIServePage.WebElement("xpath:=//MD-DIALOG-CONTENT[@id='dialog_maVerify_popup']")
	End Function
	
	Public Function btnVerificationOk()
		Set btnVerificationOk=gObjIServePage.WebElement("xpath:=//BUTTON[@id='maVerifyOk_button']")
	End Function
	
	Public Function btnNotVerified()
		Set btnNotVerified=gObjIServePage.WebButton("xpath:=//BUTTON[@id='navBarVerify_button']")
	End Function
	Public Function btnVerifyCustomer()
		Set btnVerifyCustomer=gObjIServePage.WebButton("xpath:=//BUTTON[@id='vrfyCustSubmit_button']")
	End Function
	Public Function tblTDaccountNoHeader()
		Set tblTDaccountNoHeader=gObjIServePage.WebElement("xpath:=//DIV[@id='selectedAccount_table_header']")
	End Function
	Public Function tblTDaccountNoContent()
		Set tblTDaccountNoContent=gObjIServePage.WebElement("xpath:=//DT-BODY[@id='selectedAccount_table_content']")
	End Function
	Public Function tblTDDepositNumberHeader()
		Set tblTDDepositNumberHeader=gObjIServePage.WebElement("xpath:=//DIV[@id='depositAcct_table_header']")
	End Function
	Public Function tblTDDepositNumberContent()
		Set tblTDDepositNumberContent=gObjIServePage.WebElement("xpath:=//DT-BODY[@id='depositAcct_table_content']")
	End Function
	Public Function tblTDMaturityInstructionHeader()
		Set tblTDMaturityInstructionHeader=gObjIServePage.WebElement("xpath:=//DIV[@id='maturity_table_header']")
	End Function
	Public Function tblTDMaturityInstructionContent()
		Set tblTDMaturityInstructionContent=gObjIServePage.WebElement("xpath:=//DT-BODY[@id='maturity_table_content']")
	End Function
	Public Function lstNewMaturityInstruction()
		Set lstNewMaturityInstruction=gObjIServePage.WebElement("xpath:=(//*[@id='rightPanelLayout']//md-autocomplete)[1]")
	End Function
	Public Function lstNewNextTenorDays()
		Set lstNewNextTenorDays=gObjIServePage.WebElement("xpath:=//md-autocomplete[@md-selected-item='newNextTenorType']")
	End Function
	Public Function lstNewNextTenorDaysCount()
		Set lstNewNextTenorDaysCount=gObjIServePage.WebElement("xpath:=//md-autocomplete[@md-selected-item='newNextTenor']")
	End Function
	Public Function btnNext()
		Set btnNext=gObjIServePage.WebButton("innerhtml:=Next","innervalue:=Next")
	End Function
	Public Function dlgUpdateConfirmation()
		Set dlgUpdateConfirmation=gObjIServePage.WebElement("xpath:=//MD-DIALOG[@id='odDetails_popup']")
	End Function
	Public Function btnProceed()
		Set btnProceed=gObjIServePage.WebElement("xpath:=//*[@id='odDetails_popup']/div/div[2]/button")
	End Function
	Public Function btnOkUpdateMI()
		Set btnOkUpdateMI=gObjIServePage.WebButton("xpath:=//*[@id='submit_popup']/div/button")
	End Function
	Public Function btnConfirmationYesUpdateMI()
		Set btnConfirmationYesUpdateMI=gObjIServePage.WebButton("xpath:=//BUTTON[@id='stpCancelYes_button']")
	End Function
	
	
	'******************************** End of Object Initialization ******************************************************************
End Class

