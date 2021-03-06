'*****This is auto generated code using code generator please Re-validate ****************

'[Click Button Address Change]
Public Function clickButtonAddressChange()
   bDevPending=false
   AddressChange_Page.btnAddressChange.click
   If Err.Number<>0 Then
       clickButtonAddressChange=false
       LogMessage "WARN","Verification","Failed to Click Button : AddressChange" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonAddressChange=true
End Function

'[Verify Button Address Change not exist]
Public Function verifyButtonAddressChange()
   bDevPending=false
   bclickButtonAddressChange=true
    If not (AddressChange_Page.btnAddressChange().exist(2)) Then
      LogMessage "RSLT","Verification","Address Change button is not available as expected.", True
      bclickButtonAddressChange=true
    else
      LogMessage "WARN","Verification","Address Change button is available. Expected to be not present.", False
      bclickButtonAddressChange = false
    End If      
   verifyButtonAddressChange=bclickButtonAddressChange
End Function

'[Verify Table Selected Cards Content displayed]
Public Function verifyADSelectedCardsContentTabledisplayed()
   bDevPending=false
   WaitForIcallLoading
   verifyADSelectedCardsContentTabledisplayed= AddressChange_Page.tblSelectedCardsContent.Exist(1)
   WaitForIcallLoading
End Function

'[Verify Table Selected Cards Content has following Columns]
Public Function verifySelectedCardsContentTableColumns(arrColumnNameList)
   bDevPending=false
   verifySelectedCardsContentTableColumns=verifyTableColumns(AddressChange_Page.tblSelectedCardsContent,arrColumnNameList)
End Function

'[Verify row Data in Table Selected Cards for Address Change]
Public Function verifytblSelectedCardsContent_AddrChange(arrRowDataList)
   bDevPending=false
   WaitForIcallLoading
   verifytblSelectedCardsContent_AddrChange=verifyTableContentList(AddressChange_Page.tblSelectedCardsHeader,AddressChange_Page.tblSelectedCardsContent,arrRowDataList,"SelectedCardsContent",false,null,null,null)
End Function

'[Click <Column Name> link in Table Selected Cards Content]
Public Function clickSelectedCardsContent_link(arrRowDataList)
   bDevPending=false
   clickSelectedCardsContent_link=selectTableLink(AddressChange_Page.tblSelectedCardsHeader,AddressChange_Page.tblSelectedCardsContent,arrRowDataList,"SelectedCardsContent" ,"Column Name",false,null,null,null)
End Function

'[Verify Table Selected Cards Header displayed]
Public Function verifySelectedCardsHeaderTabledisplayed()
   bDevPending=false
   verifySelectedCardsHeaderdisplayed= AddressChange_Page.tblSelectedCardsHeader.Exist(1)
End Function

'[Verify Table Selected Cards Header has following Columns]
Public Function verifySelectedCardsHeaderTableColumns(arrColumnNameList)
   bDevPending=false
   verifySelectedCardsHeaderTableColumns=verifyTableColumns(AddressChange_Page.tblSelectedCardsHeader,arrColumnNameList)
End Function

'[Verify row Data in Table Selected Cards Header]
Public Function verifytblSelectedCardsHeader_RowData(arrRowDataList)
   bDevPending=false
   verifytblSelectedCardsHeader_RowData=verifyTableContentList(AddressChange_Page.tblSelectedCardsHeaderHeader,AddressChange_Page.tblSelectedCardsHeaderContent,arrRowDataList,"SelectedCardsHeader" , bPagination,AddressChange_Page.lnkNext ,AddressChange_Page.lnkNext1,AddressChange_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table Selected Cards Header]
Public Function clickSelectedCardsHeader_link(arrRowDataList)
   bDevPending=false
   clickSelectedCardsHeader_link=selectTableLink(AddressChange_Page.tblSelectedCardsHeaderHeader,AddressChange_Page.tblSelectedCardsHeaderContent,arrRowDataList,"SelectedCardsHeader" ,"Column Name",bPagination,AddressChange_Page.lnkNext ,AddressChange_Page.lnkNext1 ,AddressChange_Page.lnkPrevious)
End Function

'[Get AddressLine1 Label Text on Address Change]
Public Function getAddressLine1Text_AC()
   bDevPending=false
   getAddressLine1Text_AC=AddressChange_Page.lblAddressLine1.GetRoProperty("innertext")
End Function

'[Verify Field Address Line1 on Address Change displayed as]
Public Function verifyAddressLine1Text_AC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine1Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblAddressLine1(), strExpectedText, "AddressLine1")Then
           bVerifyAddressLine1Text=false
       End If
   End If
   verifyAddressLine1Text_AC=bVerifyAddressLine1Text
End Function

'[Get Address Line2 Label Text on Address Change]
Public Function getAddressLine2Text_AC()
   bDevPending=false
   getAddressLine2Text_AC=AddressChange_Page.lblAddressLine2.GetRoProperty("innertext")
End Function

'[Verify Field Address Line2 on Address Change displayed as]
Public Function verifyAddressLine2Text_AC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine2Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblAddressLine2(), strExpectedText, "AddressLine2")Then
           bVerifyAddressLine2Text=false
       End If
   End If
   verifyAddressLine2Text_AC=bVerifyAddressLine2Text
End Function

'[Get AddressLine3 Label Text on Address Change]
Public Function getAddressLine3Text_AC()
   bDevPending=false
   getAddressLine3Text_AC=AddressChange_Page.lblAddressLine3.GetRoProperty("innertext")
End Function

'[Verify Field Address Line3 on Address Change displayed as]
Public Function verifyAddressLine3Text_AC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine3Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblAddressLine3(), strExpectedText, "AddressLine3")Then
           bVerifyAddressLine3Text=false
       End If
   End If
   verifyAddressLine3Text_AC=bVerifyAddressLine3Text
End Function

'[Get AddressLine4 Label Text on Address Change]
Public Function getAddressLine4Text_AC()
   bDevPending=false
   getAddressLine4Text_AC=AddressChange_Page.lblAddressLine4.GetRoProperty("innertext")
End Function

'[Verify Field Address Line4 on Address Change displayed as]
Public Function verifyAddressLine4Text_AC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine4Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblAddressLine4(), strExpectedText, "AddressLine4")Then
           bVerifyAddressLine4Text=false
       End If
   End If
   verifyAddressLine4Text_AC=bVerifyAddressLine4Text
End Function

'[Get AddressLine5 Label Text on Address Change]
Public Function getAddressLine5Text_AC()
   bDevPending=false
   getAddressLine5Text_AC=AddressChange_Page.lblAddressLine5.GetRoProperty("innertext")
End Function

'[Verify Field Address Line5 on Address Change displayed as]
Public Function verifyAddressLine5Text_AC(strExpectedText)
   bDevPending=false
   bVerifyAddressLine5Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblAddressLine5(), strExpectedText, "AddressLine5")Then
           bVerifyAddressLine5Text=false
       End If
   End If
   verifyAddressLine5Text_AC=bVerifyAddressLine5Text
End Function

'[Get Country Label Text on Address Change]
Public Function getCountryText_AC()
   bDevPending=false
   getCountryText_AC=AddressChange_Page.lblCountry.GetRoProperty("innertext")
End Function

'[Verify Field Country on Address Change displayed as]
Public Function verifyCountryText_AC(strExpectedText)
   bDevPending=false
   bVerifyCountryText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblCountry(), strExpectedText, "Country")Then
           bVerifyCountryText=false
       End If
   End If
   verifyCountryText_AC=bVerifyCountryText
End Function

'[Get PostalCode Label Text on Address Change screen]
Public Function getPostalCodeText_AC()
   bDevPending=false
   getPostalCodeText_AC=AddressChange_Page.lblPostalCode.GetRoProperty("innertext")
End Function

'[Verify Field PostalCode on Address Change displayed as]
Public Function verifyPostalCodeText_AC(strExpectedText)
   bDevPending=false
   bVerifyPostalCodeText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblPostalCode(), strExpectedText, "PostalCode")Then
           bVerifyPostalCodeText=false
       End If
   End If
   verifyPostalCodeText_AC=bVerifyPostalCodeText
End Function

'[Select Combobox Country on Address Change as]
Public Function selectCountryComboBox_AC(strCountry)
   bDevPending=false
   bSelectCountryComboBox=true
   If Not IsNull(strCountry) Then
       If Not (selectItem_Combobox (AddressChange_Page.lstCountry(), strCountry))Then
            LogMessage "WARN","Verification","Failed to select :"&strCountry&" From Country drop down list" ,false
           bSelectCountryComboBox=false
       End If
   End If
   selectCountryComboBox_AC=bSelectCountryComboBox
End Function

'[Get selected item from combo box Country]
Public Function getCountrySelectedItem()
   bDevPending=false
   getCountrySelectedItem=getVadinCombo_SelectedItem(AddressChange_Page.lstCountry)
End Function

'[Verify Combobox Country displayed as]
Public Function verifyCountryText(strExpectedText)
   bDevPending=false
   bVerifyCountryText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (AddressChange_Page.lstCountry(), strExpectedText, "Country")Then
           bVerifyCountryText=false
       End If
   End If
   verifyCountryText=bVerifyCountryText
End Function

'[Get PostalCode Label Text]
Public Function getPostalCodeText()
   bDevPending=false
   getPostalCodeText=AddressChange_Page.txtPostalCode.GetRoProperty("innertext")
End Function

'[Verify Field PostalCode displayed as]
Public Function verifyPostalCodeText(strExpectedText)
   bDevPending=false
   bVerifyPostalCodeText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtPostalCode(), strExpectedText, "PostalCode")Then
           bVerifyPostalCodeText=false
       End If
   End If
   verifyPostalCodeText=bVerifyPostalCodeText
End Function


'[Set TextBox Postal Code to]
Public Function setPostalCodeTextbox(strPostalCode)
   bDevPending=false
   AddressChange_Page.txtPostalCode.Set(strPostalCode)
   If Err.Number<>0 Then
       setPostalCodeTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :PostalCode" ,false
       Exit Function
   End If
   setPostalCodeTextbox=true
End Function

'[Click Button Lookup]
Public Function clickButtonLookup()
   bDevPending=false
   AddressChange_Page.btnLookup.click
   If Err.Number<>0 Then
       clickButtonLookup=false
            LogMessage "WARN","Verification","Failed to Click Button : Lookup" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonLookup=true
End Function

'[Get AddressLine1 Label Text]
Public Function getAddressLine1Text()
   bDevPending=false
   getAddressLine1Text=AddressChange_Page.txtAddressLine1.GetRoProperty("innertext")
End Function

'[Verify Field AddressLine1 displayed as]
Public Function verifyAddressLine1Text(strExpectedText)
   bDevPending=false
   bVerifyAddressLine1Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtAddressLine1(), strExpectedText, "AddressLine1")Then
           bVerifyAddressLine1Text=false
       End If
   End If
   verifyAddressLine1Text=bVerifyAddressLine1Text
End Function


'[Set TextBox Address Line1 on Address Change to]
Public Function setAddressLine1Textbox_AC(strAddressLine1)
   bDevPending=false
   AddressChange_Page.txtAddressLine1.Set(strAddressLine1)
   If Err.Number<>0 Then
       setAddressLine1Textbox_AC=false
            LogMessage "WARN","Verification","Failed to Set Text Box :AddressLine1" ,false
       Exit Function
   End If
   setAddressLine1Textbox_AC=true
End Function

'[Get AddressLine2 Label Text]
Public Function getAddressLine2Text()
   bDevPending=false
   getAddressLine2Text=AddressChange_Page.txtAddressLine2.GetRoProperty("innertext")
End Function

'[Verify Field AddressLine2 displayed as]
Public Function verifyAddressLine2Text(strExpectedText)
   bDevPending=false
   bVerifyAddressLine2Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtAddressLine2(), strExpectedText, "AddressLine2")Then
           bVerifyAddressLine2Text=false
       End If
   End If
   verifyAddressLine2Text=bVerifyAddressLine2Text
End Function


'[Set TextBox Address Line2 on Address Change to]
Public Function setAddressLine2Textbox_AC(strAddressLine2)
   bDevPending=false
   AddressChange_Page.txtAddressLine2.Set(strAddressLine2)
   If Err.Number<>0 Then
       setAddressLine2Textbox_AC=false
            LogMessage "WARN","Verification","Failed to Set Text Box :AddressLine2" ,false
       Exit Function
   End If
   setAddressLine2Textbox_AC=true
End Function

'[Get AddressLine3 Label Text]
Public Function getAddressLine3Text()
   bDevPending=false
   getAddressLine3Text=AddressChange_Page.txtAddressLine3.GetRoProperty("innertext")
End Function

'[Verify Field AddressLine3 displayed as]
Public Function verifyAddressLine3Text(strExpectedText)
   bDevPending=false
   bVerifyAddressLine3Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtAddressLine3(), strExpectedText, "AddressLine3")Then
           bVerifyAddressLine3Text=false
       End If
   End If
   verifyAddressLine3Text=bVerifyAddressLine3Text
End Function


'[Set TextBox Address Line3 on Address Change to]
Public Function setAddressLine3Textbox_AC(strAddressLine3)
   bDevPending=false
   AddressChange_Page.txtAddressLine3.Set(strAddressLine3)
   If Err.Number<>0 Then
       setAddressLine3Textbox_AC=false
            LogMessage "WARN","Verification","Failed to Set Text Box :AddressLine3" ,false
       Exit Function
   End If
   setAddressLine3Textbox_AC=true
End Function

'[Get AddressLine4 Label Text]
Public Function getAddressLine4Text()
   bDevPending=false
   getAddressLine4Text=AddressChange_Page.txtAddressLine4.GetRoProperty("innertext")
End Function

'[Verify Field AddressLine4 displayed as]
Public Function verifyAddressLine4Text(strExpectedText)
   bDevPending=false
   bVerifyAddressLine4Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtAddressLine4(), strExpectedText, "AddressLine4")Then
           bVerifyAddressLine4Text=false
       End If
   End If
   verifyAddressLine4Text=bVerifyAddressLine4Text
End Function


'[Set TextBox Address Line4 on Address Change to]
Public Function setAddressLine4Textbox_AC(strAddressLine4)
   bDevPending=false
   AddressChange_Page.txtAddressLine4.Set(strAddressLine4)
   If Err.Number<>0 Then
       setAddressLine4Textbox_AC=false
            LogMessage "WARN","Verification","Failed to Set Text Box :AddressLine4" ,false
       Exit Function
   End If
   setAddressLine4Textbox_AC=true
End Function

'[Get AddressLine5 Label Text]
Public Function getAddressLine5Text()
   bDevPending=false
   getAddressLine5Text=AddressChange_Page.txtAddressLine5.GetRoProperty("innertext")
End Function

'[Verify Field AddressLine5 displayed as]
Public Function verifyAddressLine5Text(strExpectedText)
   bDevPending=false
   bVerifyAddressLine5Text=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtAddressLine5(), strExpectedText, "AddressLine5")Then
           bVerifyAddressLine5Text=false
       End If
   End If
   verifyAddressLine5Text=bVerifyAddressLine5Text
End Function


'[Set TextBox Address Line5 on Address Change to]
Public Function setAddressLine5Textbox_AC(strAddressLine5)
   bDevPending=false
   AddressChange_Page.txtAddressLine5.Set(strAddressLine5)
   If Err.Number<>0 Then
       setAddressLine5Textbox_AC=false
            LogMessage "WARN","Verification","Failed to Set Text Box :AddressLine5" ,false
       Exit Function
   End If
   setAddressLine5Textbox_AC=true
End Function

'[Get Description Label Text]
Public Function getDescriptionText()
   bDevPending=false
   getDescriptionText=AddressChange_Page.lblDescription.GetRoProperty("innertext")
End Function

'[Verify Field Description on Address Change displayed as]
Public Function verifyDescriptionText_AC(strExpectedText)
   bDevPending=false
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_AC=bVerifyDescriptionText
End Function

'[Verify Field KnowledgeBase on Address Change SR Screen displayed as]
Public Function verifyKnowledgeBase_AddrChange(strExpectedLink)
   bDevPending=false
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strExpectedLink) Then
		
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"
			strKBLink = AddressChange_Page.lnkKnowledgeBase.GetROProperty("href") 'strKBLink=AddressChange_Page.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_AddrChange=bVerifyKnowledgeBaseText
End Function

'[Click Button AddNotes]
Public Function clickButtonAddNotes()
   bDevPending=false
   AddressChange_Page.btnAddNotes.click
   If Err.Number<>0 Then
       clickButtonAddNotes=false
            LogMessage "WARN","Verification","Failed to Click Button : AddNotes" ,false
       Exit Function
   End If
   clickButtonAddNotes=true
End Function

'[Get Comment Label Text]
Public Function getCommentText()
   bDevPending=false
   getCommentText=AddressChange_Page.txtComment.GetRoProperty("innertext")
End Function

'[Verify Field Comment displayed as]
Public Function verifyCommentText(strExpectedText)
   bDevPending=false
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText=bVerifyCommentText
End Function

'[Set TextBox Comment on Address Change to]
Public Function setCommentTextbox_AC(strComment)
   bDevPending=False
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comment on Address Change to"
	gstrParameterNameStep = "TimeStamp"&replace((replace((replace(now,"/","-"))," ","-")),":","-")
	insertDataStore gstrParameterNameStep, strComment
	'insertDataStore "SRComment", strComment	
   AddressChange_Page.txtComment.Set(strComment)
   WaitForIcallLoading
   If Err.Number<>0 Then
       setCommentTextbox_AC=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_AC=true
End Function

'[Verify Popup ValidationMessage exist]
Public Function verifyPopupValidationMessageexist(bExist)
   bDevPending=false
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

'[Get Notes Label Text]
Public Function getNotesText()
   bDevPending=false
   getNotesText=AddressChange_Page.txtNotes.GetRoProperty("innertext")
End Function

'[Verify Field Notes displayed as]
Public Function verifyNotesText(strExpectedText)
   bDevPending=false
   bVerifyNotesText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( AddressChange_Page.txtNotes(), strExpectedText, "Notes")Then
           bVerifyNotesText=false
       End If
   End If
   verifyNotesText=bVerifyNotesText
End Function

'[Set TextBox Notes to]
Public Function setNotesTextbox(strNotes)
   bDevPending=false
   AddressChange_Page.txtNotes.Set(strNotes)
   If Err.Number<>0 Then
       setNotesTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Notes" ,false
       Exit Function
   End If
   setNotesTextbox=true
End Function

'[Perform Add Notes by clicking Add Notes Button on Address Change Screen]
Public Function addNote_AC(strNote)
   bDevPending=false
   bVerifypopupNotes=true
	Dim bVerifypopupNotes:VerifypopupNotes=true
	
	If not isNull(strNote) Then
		AddressChange_Page.btnAddNotes.click
		WaitForICallLoading
            If not AddressChange_Page.popupValidationMessage.exist(5)Then
				LogMessage "WARN","Verification","New Note dialog did not displayed",false
				bVerifypopupNotes=false
			 else
			 strMessage=AddressChange_Page.lblMaxAllowed.GetROProperty("innerText")
				If not strMessage="Max allowed - 3000" Then
					LogMessage "WARN","Verification","Add New Comment popup dislog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage,false
					bVerifypopupNotes=false
				End If
			   ServiceRequest.txtNewComment.set strNote
			  
				   ServiceRequest.clickSave_Popup
			  WaitForIcallLoading
		   End If 
		End If 
	addNote_AC=bVerifypopupNotes
End Function

'[Click Button OK_ValidationMsg]
Public Function clickButtonOK_ValidationMsg()
   bDevPending=false
   AddressChange_Page.btnOK_ValidationMsg.click
   If Err.Number<>0 Then
       clickButtonOK_ValidationMsg=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationMsg" ,false
       Exit Function
   End If
   clickButtonOK_ValidationMsg=true
End Function

'[Get ValidationMessage Label Text]
Public Function getValidationMessageText()
   bDevPending=false
   getValidationMessageText=AddressChange_Page.lblValidationMessage.GetRoProperty("innertext")
End Function

'[Verify Field ValidationMessage displayed on Address Change as]
Public Function verifyValidationMessageText_AC(strExpectedText)
   bDevPending=false
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblValidationMessage(), strExpectedText, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
   End If
   AddressChange_Page.btnOK_ValidationMsg.Click
   WaitForIcallLoading
   verifyValidationMessageText_AC=bVerifyValidationMessageText
End Function

'[Click Button Cancel on Address Change Screen]
Public Function clickButtonCancel_AC()
   bDevPending=false
   AddressChange_Page.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_AC=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickButtonCancel_AC=true
End Function

'[Click Button Submit on Address Change]
Public Function clickButtonSubmit_AC()
   bDevPending=false
   AddressChange_Page.btnSubmit.click
   WaitForICallLoading
   If Err.Number<>0 Then
       clickButtonSubmit_AC=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonSubmit_AC=true
End Function

'[Verify Button Submit is disable on Address Change Screen]
Public Function VerifybtnSubmit_AC()
	bDevPending=False
   Dim bVerifybtnSubmit_AC:bVerifybtnSubmit_AC=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	intBtnSubmit=Instr(AddressChange_Page.btnSubmit.GetROproperty("outerhtml"),("disabled"))
	If  not intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Submit button is disable as per expectation.",True
		bVerifybtnSubmit_AC=true
	Else
		LogMessage "WARN","Verifiation","Submit button is enable. Expected to be disable.",false
		bVerifybtnSubmit_AC=false
	End If
	VerifybtnSubmit_AC=bVerifybtnSubmit_AC
End Function

'[Verify Validation Message displayed on Address Change as]
Public Function verifyValidationMessage_AC(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (AddressChange_Page.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
   End If
   ServiceActivation.btnOK_ValidationPopup.Click
   'AddressChange.btnOK_ValidationMsg.Click
   WaitForIcallLoading
   verifyValidationMessage_AC=bVerifyValidationMessageText
End Function

'[Verify Overseas address warning message]
Public Function verifyOverseasAddress_AC(strExpectedText)
   bDevPending=false
   bverifyOverseasAddress_AC=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (AddressChange_Page.lblForeignVerificationMessage(), strExpectedText, "Overseas Warning")Then
           bverifyOverseasAddress_AC=false
       End If
   End If
   verifyOverseasAddress_AC=bverifyOverseasAddress_AC
End Function

'[Verify Field TM Approval Message displayed on Address Change as]
Public Function verifyTMApprovalMessage_AC(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (AddressChange_Page.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
   End If
   AddressChange_Page.btnOK_ValidationMsg.Click
   WaitForIcallLoading
   verifyTMApprovalMessage_AC=bVerifyValidationMessageText
End Function

'[Click button Close on Request Submitted Popup for Address Change]
Public Function clickBtnClose_RequestSubmittedAC()
	bDevPending=false
	clickBtnClose_RequestSubmittedAC = true
	For iCount = 1 To 240 Step 1
		If Not AddressChange_Page.btnCancel_RequestSubmitted.Exist(0.5) Then
			Wait(0.5)
		else
			AddressChange_Page.btnCancel_RequestSubmitted.click
			Exit for
		End if
	Next
   If Err.Number<>0 Then
       clickBtnClose_RequestSubmittedAC=false
            LogMessage "WARN","Verification","Failed to Click Button : Close_RequestSubmitted" ,false
       Exit Function
   End If
   WaitForICallLoading
   'clickBtnClose_RequestSubmittedAC=true
End Function
