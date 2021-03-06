'*****This is auto generated code using code generator please Re-validate ****************
Dim strAddressLine1:strAddressLine1=""
Dim strAddressLine2:strAddressLine2=""
Dim strAddressLine3:strAddressLine3=""
Dim strAddressLine4:strAddressLine4=""
Dim strAddressLine5:strAddressLine5=""
Dim strPostalCode:strPostalCode=""
Dim strCountry:strCountry=""
Dim strPointsExpired:strPointsExpired=""
Dim strSubtotal:strSubtotal=""

'[Get Address details from Relationship Details page]
Public Function getAddress_RR()
strAddressLine1=Relationship_Details.lblAddressLine1.GetROProperty("innertext")
strAddressLine2=Relationship_Details.lblAddressLine2.GetROProperty("innertext")
strAddressLine3=Relationship_Details.lblAddressLine3.GetROProperty("innertext")
strAddressLine4=Relationship_Details.lblAddressLine4.GetROProperty("innertext")
strAddressLine5=Relationship_Details.lblAddressLine5.GetROProperty("innertext")
strPostalCode=Relationship_Details.lblPostalCode.GetROProperty("innertext")
strCountry=Relationship_Details.lblCountry.GetROProperty("innertext")
 'strExpirePoints = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 3)
 'strSubtotal = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 5, "WebElement", 0).getroproperty("innertext")         
End Function

'[Get Rewards Points details from Rewards Redemption page]
Public Function getAmount_RR()
'strPointsExpired = PointsthatExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 1)
strPointsExpired = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 1)
strSubtotal = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 5, "WebElement", 0).getroproperty("innertext")         
End Function


'[Click Button RewardsRedemption]
Public Function clickButtonRewardsRedemption()
   bDevPending=false
   bclickButtonRewardsRedemption=true
   RewardsRedemption_Page.btnRewardsRedemption.click
   If Err.Number<>0 Then
       bclickButtonRewardsRedemption=false
            LogMessage "WARN","Verification","Failed to Click Button : RewardsRedemption" ,false
       Exit Function
   End If
   WaitForICallLoading
   clickButtonRewardsRedemption=bclickButtonRewardsRedemption
End Function

'[Verify Confirmation Popup on Rewards Redemption]
Public Function verifyConfirmationPopup_CL()
	bverifyConfirmationPopup=true
	If Not verifyInnerText(RewardsRedemption_Page.lblValidationMessage(), "Are you sure you want to discard the changes (if any) and leave this page?", "Verify Pop up confirmation") Then
		bverifyConfirmationPopup=false
	End If
	RewardsRedemption_Page.btnYes_ValidationMsg.click
	  If Err.Number<>0 Then
       bverifyConfirmationPopup=false
            LogMessage "WARN","Verification","Failed to Click Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
	verifyConfirmationPopup_CL=bverifyConfirmationPopup
End Function

'[Verify Table SelectedCardsContent displayed]
Public Function verifySelectedCardsContentTabledisplayed()
   bDevPending=true
   verifySelectedCardsContentdisplayed= RewardsRedemption_Page.tblSelectedCardsContent.Exist(1)
End Function
'[Verify Table SelectedCardsContent has following Columns]
Public Function verifySelectedCardsContentTableColumns(arrColumnNameList)
   bDevPending=true
   verifySelectedCardsContentTableColumns=verifyTableColumns(RewardsRedemption_Page.tblSelectedCardsContent,arrColumnNameList)
End Function
'[Verify row Data in Table SelectedCardsContent]
Public Function verifytblSelectedCardsContent_RowData(arrRowDataList)
   bDevPending=true
   verifytblSelectedCardsContent_RowData=verifyTableContentList(RewardsRedemption_Page.tblSelectedCardsContentHeader,RewardsRedemption_Page.tblSelectedCardsContentContent,arrRowDataList,"SelectedCardsContent" , bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1,RewardsRedemption_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table SelectedCardsContent]
Public Function clickSelectedCardsContent_link(arrRowDataList)
   bDevPending=true
   clickSelectedCardsContent_link=selectTableLink(RewardsRedemption_Page.tblSelectedCardsContentHeader,RewardsRedemption_Page.tblSelectedCardsContentContent,arrRowDataList,"SelectedCardsContent" ,"Column Name",bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1 ,RewardsRedemption_Page.lnkPrevious)
End Function

'[Verify Table SelectedCardsHeader displayed]
Public Function verifySelectedCardsHeaderTabledisplayed()
   bDevPending=true
   verifySelectedCardsHeaderdisplayed= RewardsRedemption_Page.tblSelectedCardsHeader.Exist(1)
End Function
'[Verify Table SelectedCardsHeader has following Columns]
Public Function verifySelectedCardsHeaderTableColumns(arrColumnNameList)
   bDevPending=true
   verifySelectedCardsHeaderTableColumns=verifyTableColumns(RewardsRedemption_Page.tblSelectedCardsHeader,arrColumnNameList)
End Function
'[Verify row Data in Table SelectedCardsHeader]
Public Function verifytblSelectedCardsHeader_RowData(arrRowDataList)
   bDevPending=true
   verifytblSelectedCardsHeader_RowData=verifyTableContentList(RewardsRedemption_Page.tblSelectedCardsHeaderHeader,RewardsRedemption_Page.tblSelectedCardsHeaderContent,arrRowDataList,"SelectedCardsHeader" , bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1,RewardsRedemption_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table SelectedCardsHeader]
Public Function clickSelectedCardsHeader_link(arrRowDataList)
   bDevPending=true
   clickSelectedCardsHeader_link=selectTableLink(RewardsRedemption_Page.tblSelectedCardsHeaderHeader,RewardsRedemption_Page.tblSelectedCardsHeaderContent,arrRowDataList,"SelectedCardsHeader" ,"Column Name",bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1 ,RewardsRedemption_Page.lnkPrevious)
End Function

'[Get AddressLine1 Label Text]
Public Function getAddressLine1Text()
   bDevPending=true
   getAddressLine1Text=RewardsRedemption_Page.lblAddressLine1.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption AddressLine1 displayed as]
Public Function verifyAddressLine1Text_RR(strAddressLine1)
   bDevPending=false
   bVerifyAddressLine1Text=true
   wait 5
   If Not IsNull(strAddressLine1) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblAddressLine1(), strAddressLine1, "AddressLine1")Then
           bVerifyAddressLine1Text=false
       End If
   End If
   verifyAddressLine1Text_RR=bVerifyAddressLine1Text
End Function

'[Get AddressLine2 Label Text]
Public Function getAddressLine2Text()
   bDevPending=true
   getAddressLine2Text=RewardsRedemption_Page.lblAddressLine2.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption AddressLine2 displayed as]
Public Function verifyAddressLine2Text_RR(strAddressLine2)
   bDevPending=false
   bVerifyAddressLine2Text=true   
   If Not IsNull(strAddressLine2) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblAddressLine2(), strAddressLine2, "AddressLine2")Then
           bVerifyAddressLine2Text=false
       End If
   End If
   verifyAddressLine2Text_RR=bVerifyAddressLine2Text
End Function

'[Get AddressLine3 Label Text]
Public Function getAddressLine3Text()
   bDevPending=true
   getAddressLine3Text=RewardsRedemption_Page.lblAddressLine3.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption AddressLine3 displayed as]
Public Function verifyAddressLine3Text_RR(strAddressLine3)
   bDevPending=false
   bVerifyAddressLine3Text=true
   If Not IsNull(strAddressLine3) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblAddressLine3(), strAddressLine3, "AddressLine3")Then
           bVerifyAddressLine3Text=false
       End If
   End If
   verifyAddressLine3Text_RR=bVerifyAddressLine3Text
End Function

'[Get AddressLine4 Label Text]
Public Function getAddressLine4Text()
   bDevPending=true
   getAddressLine4Text=RewardsRedemption_Page.lblAddressLine4.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption AddressLine4 displayed as]
Public Function verifyAddressLine4Text_RR(strAddressLine4)
   bDevPending=false
   bVerifyAddressLine4Text=true
   If Not IsNull(strAddressLine4) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblAddressLine4(), strAddressLine4, "AddressLine4")Then
           bVerifyAddressLine4Text=false
       End If
   End If
   verifyAddressLine4Text_RR=bVerifyAddressLine4Text
End Function

'[Get AddressLine5 Label Text]
Public Function getAddressLine5Text()
   bDevPending=true
   getAddressLine5Text=RewardsRedemption_Page.lblAddressLine5.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption AddressLine5 displayed as]
Public Function verifyAddressLine5Text_RR(strAddressLine5)
   bDevPending=false
   bVerifyAddressLine5Text=true
   If Not IsNull(strAddressLine5) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblAddressLine5(), strAddressLine5, "AddressLine5")Then
           bVerifyAddressLine5Text=false
       End If
   End If
   verifyAddressLine5Text_RR=bVerifyAddressLine5Text
End Function

'[Get Country Label Text]
Public Function getCountryText()
   bDevPending=true
   getCountryText=RewardsRedemption_Page.lblCountry.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption Country displayed as]
Public Function verifyCountryText_RR(strCountry)
   bDevPending=false
   bVerifyCountryText=true
   If Not IsNull(strCountry) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblCountry(), strCountry, "Country")Then
           bVerifyCountryText=false
       End If
   End If
   verifyCountryText_RR=bVerifyCountryText
End Function

'[Get PostalCode Label Text]
Public Function getPostalCodeText()
   bDevPending=true
   getPostalCodeText=RewardsRedemption_Page.lblPostalCode.GetRoProperty("innertext")
End Function

'[Verify Field Rewards Redemption PostalCode displayed as]
Public Function verifyPostalCodeText_RR(strPostalCode)
   bDevPending=false
   bVerifyPostalCodeText=true
   If Not IsNull(strPostalCode) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblPostalCode(), strPostalCode, "PostalCode")Then
           bVerifyPostalCodeText=false
       End If
   End If
   verifyPostalCodeText_RR=bVerifyPostalCodeText
End Function

'[Get Description Label Text]
Public Function getDescriptionText()
   bDevPending=true
   getDescriptionText=RewardsRedemption_Page.lblDescription.GetRoProperty("innertext")
End Function

'[Verify Field Description displayed as]
Public Function verifyDescriptionText_RR(strExpectedText)
   bDevPending=false
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_RR=bVerifyDescriptionText
End Function

'[Click Link KnowledgeBase]
Public Function clickLinkKnowledgeBase_RR()
   bDevPending=true
   RewardsRedemption_Page.lnkKnowledgeBase.click
   If Err.Number<>0 Then
       clickLinkKnowledgeBase=false
            LogMessage "WARN","Verification","Failed to Click Link : KnowledgeBase" ,false
       Exit Function
   End If
   clickLinkKnowledgeBase_RR=true
End Function

'[Click Button AddNotes]
Public Function clickButtonAddNotes()
   bDevPending=false
   RewardsRedemption_Page.btnAddNotes.click
   If Err.Number<>0 Then
       clickButtonAddNotes=false
            LogMessage "WARN","Verification","Failed to Click Button : AddNotes" ,false
       Exit Function
   End If
   clickButtonAddNotes=true
End Function

'[Get Comment Label Text]
Public Function getCommentText()
   bDevPending=true
   getCommentText=RewardsRedemption_Page.txtComment.GetRoProperty("innertext")
End Function

'[Verify Field Comment displayed as]
Public Function verifyCommentText(strExpectedText)
   bDevPending=false
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( RewardsRedemption_Page.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText=bVerifyCommentText
End Function


'[Set TextBox Comments in Rewards Redemption as]
Public Function setCommentTextbox_RR(strComment)
   bDevPending=false
   bsetCommentTextbox=true
   strTimeStamp = ""&now
	strComment =strComment &" "&strTimeStamp
	gstrRuntimeCommentStep="Set TextBox Comments in Rewards Redemption as"
	insertDataStore "SRComment", strComment
   RewardsRedemption_Page.txtComment.Set(strComment)
   If Err.Number<>0 Then
       bsetCommentTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_RR=bsetCommentTextbox
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
   bDevPending=true
   getNotesText=RewardsRedemption_Page.txtNotes.GetRoProperty("innertext")
End Function

'[Verify Field Notes displayed as]
Public Function verifyNotesText(strExpectedText)
   bDevPending=false
   bVerifyNotesText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( RewardsRedemption_Page.txtNotes(), strExpectedText, "Notes")Then
           bVerifyNotesText=false
       End If
   End If
   verifyNotesText=bVerifyNotesText
End Function


'[Set TextBox Notes to]
Public Function setNotesTextbox(strNotes)
   bDevPending=false
    bsetNotesTextbox=true
   RewardsRedemption_Page.txtNotes.Set(strNotes)
   If Err.Number<>0 Then
       bsetNotesTextbox=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Notes" ,false
       Exit Function
   End If
   setNotesTextbox=bsetNotesTextbox
End Function

'[Click Button OK_ValidationMsg]
Public Function clickButtonOK_ValidationMsg()
   bDevPending=false
   bclickButtonOK_ValidationMsg=true
   RewardsRedemption_Page.btnOK_ValidationMsg.click
   If Err.Number<>0 Then
       bclickButtonOK_ValidationMsg=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationMsg" ,false
       Exit Function
   End If
   clickButtonOK_ValidationMsg=bclickButtonOK_ValidationMsg
End Function

'[Get ValidationMessage Label Text]
Public Function getValidationMessageText()
   bDevPending=true
   getValidationMessageText=RewardsRedemption_Page.lblValidationMessage.GetRoProperty("innertext")
End Function

'[Verify Field ValidationMessage displayed as]
Public Function verifyValidationMessageText(strExpectedText)
   bDevPending=false
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (RewardsRedemption_Page.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
           bVerifyValidationMessageText=false
       End If
   End If
   verifyValidationMessageText=bVerifyValidationMessageText
End Function

'[Click Button Cancel on Rewards Redemption screen]
Public Function clickButtonCancel()
   bDevPending=false
   bclickButtonCancel=true
   RewardsRedemption_Page.btnCancel.click
   If Err.Number<>0 Then
       bclickButtonCancel=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickButtonCancel=bclickButtonCancel
End Function

'[Click Button Submit on Rewards Redemption screen]
Public Function clickButtonSubmit()
   bDevPending=true
   bclickButtonSubmit=true
   RewardsRedemption_Page.btnSubmit.click
   If Err.Number<>0 Then
       bclickButtonSubmit=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   clickButtonSubmit=bclickButtonSubmit
End Function

'[Verify Table RelationshipPointSummaryContent displayed]
Public Function verifyRelationshipPointSummaryContentTabledisplayed()
   bDevPending=true
   verifyRelationshipPointSummaryContentdisplayed= RewardsRedemption_Page.tblRelationshipPointSummaryContent.Exist(1)
End Function
'[Verify Table RelationshipPointSummaryContent has following Columns]
Public Function verifyRelationshipPointSummaryContentTableColumns(arrColumnNameList)
   bDevPending=true
   verifyRelationshipPointSummaryContentTableColumns=verifyTableColumns(RewardsRedemption_Page.tblRelationshipPointSummaryContent,arrColumnNameList)
End Function
'[Verify row Data in Table RelationshipPointSummaryContent]
Public Function verifytblRelationshipPointSummaryContent_RowData(arrRowDataList)
   bDevPending=true
   verifytblRelationshipPointSummaryContent_RowData=verifyTableContentList(RewardsRedemption_Page.tblRelationshipPointSummaryContentHeader,RewardsRedemption_Page.tblRelationshipPointSummaryContentContent,arrRowDataList,"RelationshipPointSummaryContent" , bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1,RewardsRedemption_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table RelationshipPointSummaryContent]
Public Function clickRelationshipPointSummaryContent_link(arrRowDataList)
   bDevPending=true
   clickRelationshipPointSummaryContent_link=selectTableLink(RewardsRedemption_Page.tblRelationshipPointSummaryContentHeader,RewardsRedemption_Page.tblRelationshipPointSummaryContentContent,arrRowDataList,"RelationshipPointSummaryContent" ,"Column Name",bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1 ,RewardsRedemption_Page.lnkPrevious)
End Function

'[Verify Table RelationshipPointSummaryHeader displayed]
Public Function verifyRelationshipPointSummaryHeaderTabledisplayed()
   bDevPending=true
   verifyRelationshipPointSummaryHeaderdisplayed= RewardsRedemption_Page.tblRelationshipPointSummaryHeader.Exist(1)
End Function
'[Verify Table RelationshipPointSummaryHeader has following Columns]
Public Function verifyRelationshipPointSummaryHeaderTableColumns(arrColumnNameList)
   bDevPending=true
   verifyRelationshipPointSummaryHeaderTableColumns=verifyTableColumns(RewardsRedemption_Page.tblRelationshipPointSummaryHeader,arrColumnNameList)
End Function
'[Verify row Data in Table RelationshipPointSummaryHeader]
Public Function verifytblRelationshipPointSummaryHeader_RowData(arrRowDataList)
   bDevPending=true
   verifytblRelationshipPointSummaryHeader_RowData=verifyTableContentList(RewardsRedemption_Page.tblRelationshipPointSummaryHeaderHeader,RewardsRedemption_Page.tblRelationshipPointSummaryHeaderContent,arrRowDataList,"RelationshipPointSummaryHeader" , bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1,RewardsRedemption_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table RelationshipPointSummaryHeader]
Public Function clickRelationshipPointSummaryHeader_link(arrRowDataList)
   bDevPending=true
   clickRelationshipPointSummaryHeader_link=selectTableLink(RewardsRedemption_Page.tblRelationshipPointSummaryHeaderHeader,RewardsRedemption_Page.tblRelationshipPointSummaryHeaderContent,arrRowDataList,"RelationshipPointSummaryHeader" ,"Column Name",bPagination,RewardsRedemption_Page.lnkNext ,RewardsRedemption_Page.lnkNext1 ,RewardsRedemption_Page.lnkPrevious)
End Function

'[Perform Add Notes by clicking Add Notes Button on Rewards Redemption Screen]
Public Function addNote_RR(strNote)
   bDevPending=false
   baddNote_RR=true
	Dim baddNote_RR:addNote_RR=true
		If not isNull(strNote) Then
		RewardsRedemption_Page.btnAddNotes.click
		WaitForICallLoading
           If Not RewardsRedemption_Page.popupNotes.exist(5)Then
				LogMessage "WARN","Verification","Add New Comment action failed"
					baddNote_RR=false
					else
					LogMessage "RSLT","Verification","Add New Comment performed successfully" ,true
					baddNote_RR=True
				End If
			   RewardsRedemption_Page.txtNotesDescription.set strNote
			   RewardsRedemption_Page.btnNotesSave.Click
			  WaitForIcallLoading
		  	End If 
	addNote_RR=baddNote_RR
End Function

'[Verify Button AddNote is disabled on Rewards Redemption Screen]
Public Function VerifybtnAddNoteDisable_RR()
	bDevPending=False
   Dim bVerifybtnSubmit_RR:bVerifybtnSubmit_RR=true
	intBtnSubmit=Instr(RewardsRedemption_Page.btnAddNotes.GetROproperty("outerhtml"),("v-disabled"))
	If  not intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Add Note button is disable as per expectation.",True
		bVerifybtnSubmit_RR=true
	Else
		LogMessage "WARN","Verifiation","Add Note button is enable. Expected to be disable.",false
		bVerifybtnSubmit_RR=false
	End If
	VerifybtnAddNoteDisable_RR=bVerifybtnSubmit_RR
End Function

'Public Function VerifybtnAddNoteDisable_RR()
'   
'End function 


'[Select Rewards Option Category Combobox as]
Public Function selectRewardsOptionCategoryComboBox(strCategory)
	bDevPending=False
   Dim bselectRewardsOptionCategoryComboBox:bselectRewardsOptionCategoryComboBox=true
   wait 5
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=1 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("All") Then             
             Set oDesc=Description.Create
		      oDesc("micclass").Value = "WebElement"
		      oDesc("class").Value = "v-filterselect-button"
		      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
		      set lstObj=objComboBox.ChildObjects(oDesc)
		      lstObj(0).Click
		     		            
		      Set oDescCombo=Description.Create
		      oDescCombo("micclass").Value = "WebElement"
		      oDescCombo("class").Value = "gwt-MenuItem.*"
		      Wait 2
		      ' oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
		      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
		       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
		  
		      intItems=lstCombo.Count
		      
		      For iCount=0 to intItems-1
		        Dim strTemp:strTemp=""
		     
		      strTemp=lstCombo(iCount).GetRoProperty("text")
		      If strTemp = strCategory Then
		      	 
		          lstCombo(iCount).click
		          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
		          If  result <>  strCategory    Then
		          	LogMessage "WARN","Verifiation","Rewards Option first Category not selected",false
					bselectRewardsOptionCategoryComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option first Category selected successfully",True
					bselectRewardsOptionCategoryComboBox=true
		          End If
		          selectRewardsOptionCategoryComboBox=bselectRewardsOptionCategoryComboBox
		          Exit Function
		         'lstCombo(iCount).click
		          Exit for
		      End If
		      Next             
             
         End If
        Next  
'  If strTemp = strCategory Then
'            Exit for
'       End If        
    Next
	selectRewardsOptionCategoryComboBox=bselectRewardsOptionCategoryComboBox
End Function

'[Select Rewards Option Product Combobox as]
Public Function selectRewardsOptionProductComboBox(strProduct)
	bDevPending=False
   Dim bselectRewardsOptionProductComboBox:bselectRewardsOptionProductComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=1 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("Please Select") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      ''''''''''''''''''''''''
           
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      Wait 2
       'oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
     
     ''''''''''''''''''''''''''' Select the item from 11 - 20 in combobox dropdown
'     
'      Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").Webelement("lblNext").Click
'      wait 2

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
     
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strProduct Then
          lstCombo(iCount).click
          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
          		          If  result <>  strProduct    Then
		          	LogMessage "WARN","Verifiation","Rewards Option first Product not selected",false
					bselectRewardsOptionProductComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option first Product selected successfully",True
					bselectRewardsOptionProductComboBox=true
		          End If
		          selectRewardsOptionProductComboBox=bselectRewardsOptionProductComboBox
          Exit Function
'             lstCombo(iCount).FireEvent "onclick"
      End If
      Next             
             
         End If
        Next        
    Next
	selectRewardsOptionProductComboBox=bselectRewardsOptionProductComboBox
End Function

'[Select Rewards Option Qty Combobox as]
Public Function selectRewardsOptionQtyComboBox(strQty)
	bDevPending=False
   Dim bselectRewardsOptionQtyComboBox:bselectRewardsOptionQtyComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=1 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 4, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("1") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1,4, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      wait 2
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strQty Then
           lstCombo(iCount).click
           wait 2
           result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")  
		          If  result =  strQty    Then
		          	LogMessage "WARN","Verifiation","Rewards Option first Quantity selected successfully",true
					bselectRewardsOptionQtyComboBox=true
					else
					LogMessage "RSLT","Verification","Rewards Option first Quantity not selected ",false
					bselectRewardsOptionQtyComboBox=false
		          End If
		      selectRewardsOptionQtyComboBox=bselectRewardsOptionQtyComboBox
           Exit Function
       End If
      Next             
             
         End If
        Next        
    Next
	selectRewardsOptionQtyComboBox=bselectRewardsOptionQtyComboBox
End Function

'[Select Add new row in Rewards Option table]
Public Function selectRewardsOptionAddButton()
	bDevPending=False
   Dim bsselectRewardsOptionAddButton:bsselectRewardsOptionAddButton=true
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
      Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objAddBtn =Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1,6, "Webelement", 0)
     ' set lstObj=objComboBox.ChildObjects(oDesc)
      objAddBtn.Click
      objAddBtn.Click
	selectRewardsOptionAddButton=bsselectRewardsOptionAddButton
End Function

'''''''''''''''''''''''''''' Poornima stepfunctions''''''''''''''''''''''''''''''''

'[Verify Button RewardsRedemption not exist]
Public Function RewardsRedemption()
   bDevPending=false
   bVerifyButtonRewardsRedemption=true
    If not (RewardsRedemption_page.btnRewardsRedemption().exist(2)) Then
      LogMessage "RSLT","Verification","Rewards Redemption button is not available as expected.", True
      bVerifyButtonRewardsRedemption=true
    else
      LogMessage "WARN","Verification","Rewards Redemption  button is available. Expected to be not present.", False
      bVerifyButtonRewardsRedemption = false
    End If      
   RewardsRedemption=bVerifyButtonRewardsRedemption
End Function

'[Overall Current Balance Points]
Public Function CurrentBalances()
   bDevPending=false
   bVerifyCurrentbalance=true
   wait 4
    CurrentBalancesPointsthatExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 1)
    CurrentBalancesPointsthatDonotExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 2)
    CurrentBalancesExpirePoint = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 3)
    If cint(CurrentBalancesPointsthatExpire) > 0  or cint(CurrentBalancesPointsthatDonotExpire) > 0 or cint(CurrentBalancesExpirePoint) > 0 Then
      LogMessage "RSLT","Verification","In Rewards Redemption Current Balances as expected.", True
      bVerifyCurrentbalance=true
    else
      LogMessage "WARN","Verification","In Rewards Redemption Current Balances is not as Expected", False    
      bVerifyCurrentbalance = false
    End If      
   CurrentBalances=bVerifyCurrentbalance
End Function

'[Verify Validation Message displayed on RewardsRedemption as]
Public Function verifyValidationMessage_RR1(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   wait 5
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (RewardsRedemption_page.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
        Else
         bVerifyValidationMessageText=True
       End If
   End If
   verifyValidationMessage_RR1=bVerifyValidationMessageText
End Function

'[Click Ok button on SR Status popup]
Public Function ClickbtnOk_SR_Popup_RR()
   bDevPending=False
   bClickbtnOk_SR_Popup_RR=true
   wait 1
   RewardsRedemption_page.btnOK_ValidationMsg.Click
	If Err.Number<>0 Then
       bClickbtnOk_SR_Popup_RR=false
            LogMessage "WARN","Verification","Failed to Click Ok Button : Popup" ,false
       Exit Function
   End If
   ClickbtnOk_SR_Popup_RR=bClickbtnOk_SR_Popup_RR
End Function


'[Verify SubTotal for Greaterthan or EqualTo]
Public Function GESubTotal()
   bDevPending=false
   bVerifySubTotal=true
   Dim Total, PointsthatExpire, PointsthatDonotExpire
    PointsthatExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 1)
    PointsthatDonotExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 2)
    SubTotal = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 5, "WebElement", 0).getroproperty("innertext")         
    Total=CCur(PointsthatExpire) + CCur(PointsthatDonotExpire)
    If (CCur(Total) > CCur(SubTotal)) or (CCur(Total) = CCur(SubTotal)) Then
    LogMessage "RSLT","Verification","In Rewards Redemption SubTotal is as expected.", True
    bVerifySubTotal=true
    Else
    LogMessage "WARN","Verification","In Rewards Redemption SubTotal is not as expected. Actual -"&SubTotal&", Expected- "&Total, False
    bVerifySubTotal = false
    End IF
    GESubTotal=bVerifySubTotal
End Function


'[Verify SubTotal for lesserthan]
Public Function LSubTotal(strValidationMessage)
   bDevPending=false
   bVerifySubTotal=true
    PointsthatExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 1)
    PointsthatDonotExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 2)
    SubTotal = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 5, "WebElement", 0).getroproperty("innertext")         
    If cint(PointsthatExpire) + cint(PointsthatDonotExpire) < SubTotal Then
    wait 3
       If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (RewardsRedemption_page.lblRewardsRedemptionInlineMessage(), strValidationMessage, "Validation Message")Then 'Need to add the lbl object Sunder
           'LogMessage "WARN","Verification","In Rewards Redemption Inline message is not as expected. Actual -"&strActual&", Expected- "&strValidationMessage, False
           bVerifySubTotal=false
        Else
        'LogMessage "RSLT","Verification","In Rewards Redemption Inline message is as expected. Actual -"&strActualComment&", Expected- "&strValidationMessage, True
         bVerifySubTotal=True
       End If
   End If
End IF
    LSubTotal=bVerifySubTotal
End Function

'[Verify Expired point for Rewards]
Public Function ExpiredPoints_RR(strValidationMessage)
   bDevPending=false
   bVerifyExpiryPoints=true
       ExpirePoint = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 3)
       RedeemExpirePoints = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOption_RedemExpPoints").ChildItem(2,2, "WebElement", 0).getroproperty("innertext")   
    If  clng(RedeemExpirePoints) > cint(ExpirePoint) and cint(ExpirePoint) <> 0 or cint(ExpirePoint) = 0   Then
       If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (RewardsRedemption_page.lblRewardsRedemptionInlineMessage(), strValidationMessage, "Validation Message")Then 'Need to add the lbl object Sunder
		  'LogMessage "WARN","Verification","Rewards Inline message failed.Actual :"&strActual&" Expected : "&strValidationMessage&"" ,false          
          bVerifyExpiryPoints=false
        Else
         bVerifyExpiryPoints=True
       End If
   End If
End If
    ExpiredPoints_RR=bVerifyExpiryPoints
End Function


'[Verify valid Expired points for Rewards]
Public Function ValidExpiryPoints()
   bDevPending=false
   bValidExpiryPoints=true
       ExpirePoint = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 3)
       RedeemExpirePoints = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOption_RedemExpPoints").ChildItem(2,2, "WebElement", 0).getroproperty("innertext")   
       int_ExpirePoint=ExpirePoint
       int_RedeemExpirePoints=RedeemExpirePoints
   ' If RedeemExpirePoints=cint(ExpirePoint) or RedeemExpirePoints < cint(ExpirePoint) or RedeemExpirePoints = 0  or RedeemExpirePoints = 1 or RedeemExpirePoints = ""  Then
     'If clng(int_RedeemExpirePoints)=clng(int_ExpirePoint) or clng(int_RedeemExpirePoints) < clng(int_ExpirePoint) or int_RedeemExpirePoints = "0"  or int_RedeemExpirePoints = "1" or int_RedeemExpirePoints = ""  Then
	If int_RedeemExpirePoints <> "" Then
		If clng(int_RedeemExpirePoints)=clng(int_ExpirePoint) or clng(int_RedeemExpirePoints) < clng(int_ExpirePoint) or int_RedeemExpirePoints = "0"  or int_RedeemExpirePoints = "1" or int_RedeemExpirePoints = ""  Then
    LogMessage "RSLT","Verification","In Rewards Redemption 'RedeemExpirePoints' is as expected.", True
    	bValidExpiryPoints=true
    else
    LogMessage "WARN","Verification","In Rewards Redemption 'RedeemExpirePoints' is not as expected.", False    
'    End If condition Then
     bValidExpiryPoints = false	
    End If
     Else
      If int_RedeemExpirePoints=clng(int_ExpirePoint) or int_RedeemExpirePoints < clng(int_ExpirePoint) or int_RedeemExpirePoints = "0"  or int_RedeemExpirePoints = "1" or int_RedeemExpirePoints = ""  Then    
    LogMessage "RSLT","Verification","In Rewards Redemption 'RedeemExpirePoints' is as expected.", True
    	bValidExpiryPoints=true
    else
    LogMessage "WARN","Verification","In Rewards Redemption 'RedeemExpirePoints' is not as expected.", False    
'    End If condition Then
     bValidExpiryPoints = false	
    End If
	End If  
'  If int_RedeemExpirePoints=clng(int_ExpirePoint) or int_RedeemExpirePoints < clng(int_ExpirePoint) or int_RedeemExpirePoints = "0"  or int_RedeemExpirePoints = "1" or int_RedeemExpirePoints = ""  Then    
'    LogMessage "RSLT","Verification","In Rewards Redemption 'RedeemExpirePoints' is as expected.", True
'    	bValidExpiryPoints=true
'    else
'    LogMessage "WARN","Verification","In Rewards Redemption 'RedeemExpirePoints' is not as expected.", False    
''    End If condition Then
'     bValidExpiryPoints = false	
'    End If
     ValidExpiryPoints=bValidExpiryPoints
End Function

'[Validate and calculate TotalAvailablepoints for Rewards ]
Public Function TotalAvailablePointsandSubTotal(strQty)
   bDevPending=false
   bTotalAvailablePointsandSubTotal=true
    PointsthatExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 1)
    PointsthatDonotExpire = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 2)
    ExpirePoint = RewardsRedemption_page.tblRelationshipPointSummaryContent.GetCellData(1, 3)
    RWOptionPoints=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 3, "WebElement", 0).getroproperty("innertext")         
    Qty=strQty
    TotalAvailablepoints=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblTotalAvailablePoints").ChildItem(1, 2, "WebElement", 0).getroproperty("innertext")   
    
    If clng(PointsthatExpire)+clng(PointsthatDonotExpire)+clng(ExpirePoint) = clng(TotalAvailablepoints) Then
        LogMessage "RSLT","Verification","In Rewards Redemption 'TotalAvailablepoints' is as expected.", True
        bTotalAvailablePointsandSubTotal=true
    Else
        LogMessage "WARN","Verification","In Rewards Redemption 'TotalAvailablepoints' is not as expected.", False
        bTotalAvailablePointsandSubTotal = false
    End IF
    
    TotalAvailablePointsandSubTotal=bTotalAvailablePointsandSubTotal
    
End Function


'[Validate and calculate SubTotal points for Rewards ]
Public Function ValidPoints_SubTotal(strQty)
   bDevPending=false
   bPoints_SubTotal=true

    RWOptionPoints=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 3, "WebElement", 0).getroproperty("innertext")         
    Qty= strQty
    SubTotal=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 5, "WebElement", 0).getroproperty("innertext")         
    
    If (RWOptionPoints * Qty) = clng(SubTotal) Then
        LogMessage "RSLT","Verification","In Rewards Redemption 'SubTotal' is as expected.", True
        bPoints_SubTotal=true
    Else
        LogMessage "WARN","Verification","In Rewards Redemption 'SubTotal' is not as expected.", False
        bPoints_SubTotal = false
    End IF
    
    ValidPoints_SubTotal=bPoints_SubTotal
    
End Function

'[Verify Rewards Redemption button exist]
Public Function VerifyButtonTemporarLimitDE()
   bDevPending=false
   bVerifyButtonTemporarLimit=true
    If (RewardsRedemption_page.btnRewardsRedemption().exist(2)) Then
      LogMessage "RSLT","Verification","Reward Redemption button is available as expected.", True
      bclickButtonAddressChange=true
    else
      LogMessage "WARN","Verification","Reward Redemption button is not available. Expected to be  present.", False
      bVerifyButtonTemporarLimit = false
    End If      
   VerifyButtonTemporarLimitDE=bVerifyButtonTemporarLimit
End Function

'[Verify Field Rewards Redemption Inline Message displayed as]
Public Function verifyRewardsRedemptionInlineMessageText(strVerificationMessage)
   bDevPending=false
   bverifyRewardsRedemptionInlineMessageText=true
   WaitForICallLoading
   If Not IsNull(strVerificationMessage) Then
       If Not VerifyInnerText (RewardsRedemption_page.lblRewardsRedemptionInlineMessage(), strVerificationMessage, "RewardsRedemptionInlineMessage")Then
           LogMessage "WARN","Verification","Rewards Inline message failed.Actual :"&strActual&" Expected : "&strVerificationMessage&"" ,false
           bverifyRewardsRedemptionInlineMessageText=false
       End If
   End If
   verifyRewardsRedemptionInlineMessageText=bverifyRewardsRedemptionInlineMessageText
End Function

'[Click Close button on Request Submitted Popup for Rewards Redemption]
Public Function verifybtnClose_RequestSubmitted_RR()
	bverifybtnClose_RequestSubmitted_RR=true
	RewardsRedemption_page.btnClose_RequestSubmitted.click
	  If Err.Number<>0 Then
       bverifybtnClose_RequestSubmitted_RR=false
            LogMessage "WARN","Verification","Failed to Click Close Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
	verifybtnClose_RequestSubmitted_RR=bverifybtnClose_RequestSubmitted_RR
End Function


'[Verify Field KnowledgeBase on Rewards Redemption SR Screen displayed as]
Public Function verifyKnowledgeBase_RR(strKBLink)
   bDevPending=false
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strKBLink) Then
		
		Set oDesc_KB = Description.Create()
			oDesc_KB("micclass").Value = "Link"
		
			strKBLink_Actual=RewardsRedemption_page.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
			strExpectedLink=Replace(strKBLink_Actual,"@","=")
       If not MatchStr(strKBLink_Actual, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_RR=bVerifyKnowledgeBaseText
End Function

'[Verify Points and Subtotal displayed in Rewards as]
Public Function verifyPoints(strPoints,strSubTotal)
   bDevPending=false
   bverifyPoints=true
    Points = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 3, "WebElement", 0).getroproperty("innertext")         
   SubTotal = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1, 5, "WebElement", 0).getroproperty("innertext")         
    If Points = strPoints Then
      If Subtotal = strSubTotal Then
       'If Not IsNull(strValidationMessage) Then
       'If Not VerifyInnerText (RewardsRedemption_page.lblRewardsRedemptionInlineMessage(), strValidationMessage, "Validation Message")Then 'Need to add the lbl object Sunder
		 LogMessage "RSLT","Verification","Points & Subtotal field validation Passed",true
         bverifyPoints=True
        Else
           LogMessage "WARN","Verification","Points & Subtotal field validation failed" ,false
           bverifyPoints=false
       End If
End if
    verifyPoints=bverifyPoints
End Function


'[Click add button in Rewards Option table]
Public Function clickAddBtn_RR()
   bDevPending=false
   bclickAddBtn_RR=true
   wait 2
   Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").Image("btnadd_plus").Click
'   Set oDesc=Description.Create
'      oDesc("micclass").Value = "WebElement"
'      oDesc("class").Value = "v-image v-widget"
'      oDesc("file name").Value = "add_plus.png"
'      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(1,6, "WebElement", 0)
'      objComboBox.click
   If Err.Number<>0 Then
       bclickAddBtn_RR=false
            LogMessage "WARN","Verification","Failed to Click Link : KnowledgeBase" ,false
       Exit Function
   End If
   clickAddBtn_RR=bclickAddBtn_RR
End Function


'''''''''''''''''''''''''

'[Select third Rewards Option Category Combobox as]
Public Function selectThirdRewardsOptionCategoryComboBox_RR(strCategory2)
	bDevPending=False
   Dim bselectRewardsOptionCategoryComboBox:bselectRewardsOptionCategoryComboBox=true
   wait 5
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=3 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("All") Then             
             Set oDesc=Description.Create
		      oDesc("micclass").Value = "WebElement"
		      oDesc("class").Value = "v-filterselect-button"
		      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
		      set lstObj=objComboBox.ChildObjects(oDesc)
		      lstObj(0).Click
		     		            
		      Set oDescCombo=Description.Create
		      oDescCombo("micclass").Value = "WebElement"
		      oDescCombo("class").Value = "gwt-MenuItem.*"
		      Wait 2
		      ' oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
		      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
		       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
		  
		      intItems=lstCombo.Count
		      
		      For iCount=0 to intItems-1
		        Dim strTemp:strTemp=""
		     
		      strTemp=lstCombo(iCount).GetRoProperty("text")
		      If strTemp = strCategory2 Then
		      	 
		          lstCombo(iCount).click
		          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
		          If  result <>  strCategory2    Then
		          	LogMessage "WARN","Verifiation","Rewards Option third Category not selected",false
					bselectRewardsOptionCategoryComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option third Category selected successfully",True
					bselectRewardsOptionCategoryComboBox=true
		          End If
		          Exit Function
		         'lstCombo(iCount).click
		          Exit for
		      End If
		      Next             
             
         End If
        Next  
'  If strTemp = strCategory Then
'            Exit for
'       End If        
    Next
	selectThirdRewardsOptionCategoryComboBox_RR=bselectRewardsOptionCategoryComboBox
End Function

'[Select third Rewards Option Product Combobox as]
Public Function selectThirdRewardsOptionProductComboBox_RR(strProduct2)
	bDevPending=False
   Dim bselectRewardsOptionProductComboBox:bselectRewardsOptionProductComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=3 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("Please Select") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      Wait 2
       'oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
         ''''''''''''''''''''''''''' Select the item from 11 - 20 in combobox dropdown
     
'      Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").Webelement("lblNext").Click
'      wait 2
      
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
     
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strProduct2 Then
          lstCombo(iCount).click
          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
          		          If  result <>  strProduct2    Then
		          	LogMessage "WARN","Verifiation","Rewards Option third Product not selected",false
					bselectRewardsOptionProductComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option third Product selected successfully",True
					bselectRewardsOptionProductComboBox=true
		          End If
          Exit Function
'             lstCombo(iCount).FireEvent "onclick"
      End If
      Next             
             
         End If
        Next        
    Next
	selectThirdRewardsOptionProductComboBox_RR=bselectRewardsOptionProductComboBox
End Function

'[Select third Rewards Option Qty Combobox as]
Public Function selectThirdRewardsOptionQtyComboBox_RR(strQty2)
	bDevPending=False
   Dim bselectRewardsOptionQtyComboBox:bselectRewardsOptionQtyComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=3 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("1") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,4, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      wait 2
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strQty2 Then
           lstCombo(iCount).click
           wait 2
           result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")  
		          If  result =  strQty2    Then
		          	LogMessage "RSLT","Verifiation","Rewards Option third Quantity selected successfully",true
					bselectRewardsOptionQtyComboBox=true
					else
					LogMessage "WARN","Verification","Rewards Option third Quantity not selected ",false
					bselectRewardsOptionQtyComboBox=false
		          End If
           Exit Function
       End If
      Next             
             
         End If
        Next        
    Next
	selectThirdRewardsOptionQtyComboBox_RR=bselectRewardsOptionQtyComboBox
End Function

'''''''''''''''''''''''''''''''

'[Select fourth Rewards Option Category Combobox as]
Public Function selectFourthRewardsOptionCategoryComboBox_RR(strCategory3)
	bDevPending=False
   Dim bselectRewardsOptionCategoryComboBox:bselectRewardsOptionCategoryComboBox=true
   wait 5
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=4 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("All") Then             
             Set oDesc=Description.Create
		      oDesc("micclass").Value = "WebElement"
		      oDesc("class").Value = "v-filterselect-button"
		      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
		      set lstObj=objComboBox.ChildObjects(oDesc)
		      lstObj(0).Click
		     		            
		      Set oDescCombo=Description.Create
		      oDescCombo("micclass").Value = "WebElement"
		      oDescCombo("class").Value = "gwt-MenuItem.*"
		      Wait 2
		      ' oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
		      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
		       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
		  
		      intItems=lstCombo.Count
		      
		      For iCount=0 to intItems-1
		        Dim strTemp:strTemp=""
		     
		      strTemp=lstCombo(iCount).GetRoProperty("text")
		      If strTemp = strCategory3 Then
		      	 
		          lstCombo(iCount).click
		          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
		          If  result <>  strCategory3    Then
		          	LogMessage "WARN","Verifiation","Rewards Option fourth Category not selected",false
					bselectRewardsOptionCategoryComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option fourth Category selected successfully",True
					bselectRewardsOptionCategoryComboBox=true
		          End If
		          Exit Function
		         'lstCombo(iCount).click
		          Exit for
		      End If
		      Next             
             
         End If
        Next  
'  If strTemp = strCategory Then
'            Exit for
'       End If        
    Next
	selectFourthRewardsOptionCategoryComboBox_RR=bselectRewardsOptionCategoryComboBox
End Function

'[Select fourth Rewards Option Product Combobox as]
Public Function selectFourthRewardsOptionProductComboBox_RR(strProduct3)
	bDevPending=False
   Dim bselectRewardsOptionProductComboBox:bselectRewardsOptionProductComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=4 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("Please Select") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      Wait 2
       'oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
      
     ''''''''''''''''''''''''''' Select the item from 11 - 20 in combobox dropdown
     
      Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").Webelement("lblNext").Click
      wait 2
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      wait 2
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
     
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strProduct3 Then
          lstCombo(iCount).click
          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
          		          If  result <>  strProduct3    Then
		          	LogMessage "WARN","Verifiation","Rewards Option fourth Product not selected",false
					bselectRewardsOptionProductComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option fourth Product selected successfully",True
					bselectRewardsOptionProductComboBox=true
		          End If
          Exit Function
'             lstCombo(iCount).FireEvent "onclick"
      End If
      Next             
             
         End If
        Next        
    Next
	selectFourthRewardsOptionProductComboBox_RR=bselectRewardsOptionProductComboBox
End Function

'[Select fourth Rewards Option Qty Combobox as]
Public Function selectFourthRewardsOptionQtyComboBox_RR(strQty3)
	bDevPending=False
   Dim bselectRewardsOptionQtyComboBox:bselectRewardsOptionQtyComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=4 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("1") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,4, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      wait 2
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      wait 2
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strQty3 Then
           lstCombo(iCount).click
           wait 2
           result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")  
		          If  result =  strQty3    Then
		          	LogMessage "RSLT","Verifiation","Rewards Option fourth Quantity selected successfully",true
					bselectRewardsOptionQtyComboBox=true
					else
					LogMessage "WARN","Verification","Rewards Option fourth Quantity not selected ",false
					bselectRewardsOptionQtyComboBox=false
		          End If
           Exit Function
       End If
      Next             
             
         End If
        Next        
    Next
	selectFourthRewardsOptionQtyComboBox_RR=bselectRewardsOptionQtyComboBox
End Function


''''''''''''''''''''''''''''''''''

'[Select second Rewards Option Category Combobox as]
Public Function selectsecondRewardsOptionCategoryComboBox_RR(strCategory1)
	bDevPending=False
   Dim bselectRewardsOptionCategoryComboBox:bselectRewardsOptionCategoryComboBox=true
   wait 5
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=2 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("All") Then             
             Set oDesc=Description.Create
		      oDesc("micclass").Value = "WebElement"
		      oDesc("class").Value = "v-filterselect-button"
		      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
		      set lstObj=objComboBox.ChildObjects(oDesc)
		      lstObj(0).Click
		     		            
		      Set oDescCombo=Description.Create
		      oDescCombo("micclass").Value = "WebElement"
		      oDescCombo("class").Value = "gwt-MenuItem.*"
		      Wait 2
		      ' oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
		      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
		       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
		  
		      intItems=lstCombo.Count
		      
		      For iCount=0 to intItems-1
		        Dim strTemp:strTemp=""
		     
		      strTemp=lstCombo(iCount).GetRoProperty("text")
		      If strTemp = strCategory1 Then
		      	 
		          lstCombo(iCount).click
		          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
		          If  result <>  strCategory1    Then
		          	LogMessage "WARN","Verifiation","Rewards Option second Category not selected",false
					bselectRewardsOptionCategoryComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option second Category selected successfully",True
					bselectRewardsOptionCategoryComboBox=true
		          End If
		          Exit Function
		         'lstCombo(iCount).click
		          Exit for
		      End If
		      Next             
             
         End If
        Next  
'  If strTemp = strCategory Then
'            Exit for
'       End If        
    Next
	selectsecondRewardsOptionCategoryComboBox_RR=bselectRewardsOptionCategoryComboBox
End Function

'[Select second Rewards Option Product Combobox as]
Public Function selectsecondRewardsOptionProductComboBox_RR(strProduct1)
	bDevPending=False
   Dim bselectRewardsOptionProductComboBox:bselectRewardsOptionProductComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=2 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("Please Select") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,j, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      Wait 2
       'oDescCombo("class").Value = "v-filterselect-suggestmenu.*"
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
      
      ''''''''''''''''''''''''''' Select the item from 11 - 20 in combobox dropdown
'     
'      Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").Webelement("lblNext").Click
'      wait 2
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
     
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strProduct1 Then
          lstCombo(iCount).click
          wait 2
		          result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, j, "WebEdit", 0).getroproperty("value")  
          		          If  result <>  strProduct1    Then
		          	LogMessage "WARN","Verifiation","Rewards Option second Product not selected",false
					bselectRewardsOptionProductComboBox=false
					else
					LogMessage "RSLT","Verification","Rewards Option second Product selected successfully",True
					bselectRewardsOptionProductComboBox=true
		          End If
          Exit Function
'             lstCombo(iCount).FireEvent "onclick"
      End If
      Next             
             
         End If
        Next        
    Next
	selectsecondRewardsOptionProductComboBox_RR=bselectRewardsOptionProductComboBox
End Function

'[Select second Rewards Option Qty Combobox as]
Public Function selectsecondRewardsOptionQtyComboBox_RR(strQty1)
	bDevPending=False
   Dim bselectRewardsOptionQtyComboBox:bselectRewardsOptionQtyComboBox=true
	 'set btestcount=Browser("ICall").Page("I.Serve").WebTable("Category")
	 set btestcount=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent")
    tableRow=btestcount.GetROProperty("rows")
    tableColumn=btestcount.GetROProperty("cols")
    For i=2 to tableRow
        For j = 1 To tableColumn            
          actualMonth=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")         
         If actualMonth = ("1") Then             
             Set oDesc=Description.Create
      oDesc("micclass").Value = "WebElement"
      oDesc("class").Value = "v-filterselect-button"
      set objComboBox=Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i,4, "WebElement", 0)
      set lstObj=objComboBox.ChildObjects(oDesc)
      lstObj(0).Click
      Set oDescCombo=Description.Create
      oDescCombo("micclass").Value = "WebElement"
      oDescCombo("class").Value = "gwt-MenuItem.*"
      wait 2
      'oDescCombo("class").Value = "selectRewardsOptionCategoryComboBox"
       set lstCombo=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDescCombo)
      intItems=lstCombo.Count
      wait 2
        For iCount=0 to intItems-1
        Dim strTemp:strTemp=""
      strTemp=lstCombo(iCount).GetRoProperty("text")
      If strTemp = strQty1 Then
           lstCombo(iCount).click
           wait 2
           result = Browser("Browser_iCall_BlockCancelCard").Page("iCall_RewardsRedemption").WebTable("tblRewardsOptionContent").ChildItem(i, 4, "WebEdit", 0).getroproperty("value")  
		          If  result =  strQty1    Then
		          	LogMessage "RSLT","Verifiation","Rewards Option second Quantity selected successfully",true
					bselectRewardsOptionQtyComboBox=true
					else
					LogMessage "WARN","Verification","Rewards Option second Quantity not selected ",false
					bselectRewardsOptionQtyComboBox=false
		          End If
           Exit Function
       End If
      Next             
             
         End If
        Next        
    Next
	selectsecondRewardsOptionQtyComboBox_RR=bselectRewardsOptionQtyComboBox
End Function





