'*****This is auto generated code using code generator please Re-validate ****************
''******** For Current Credit Limit Table
'Dim strRunTimeRtlCrLimit:strRunTimeRtlCrLimit="" 'This Variable is used to Store Relationship credit limit in runtime
'Dim strRunTimeAcctTotalCrLimit:strRunTimeAcctTotalCrLimit="" 'This Variable is used to Store Account Total Credit Limit in runtime
'Dim strRunTimeEmbosserCrLimit:strRunTimeEmbosserCrLimit="" 'This Variable is used to Store Embosser Credit Limit in runtime
'
''******* For Outstanding Balances Table
'Dim strRunTimeRtlOutStandingBal:strRunTimeRtlOutStandingBal="" 'This Variable is used to Store Relationship Outstanding Balance
'Dim strRunTimeAcctOutStandingBal:strRunTimeAcctOutStandingBal="" 'This Variable is used to Store Account Outstanding Balance
'Dim strRunTimeEmbosserOutStandingBal:strRunTimeEmbosserOutStandingBal="" 'This Variable is used to Store Embosser Outstanding Balance

'[Click Button Capping of Limit]
Public Function clickButtonCappingofLimit()
   bDevPending=False
   CappingLimit_Page.btnCappingofLimit.click
   Wait(5)
   waitForIcallLoading
   If Err.Number<>0 Then
       clickButtonCappingofLimit=false
            LogMessage "WARN","Verification","Failed to Click Button : CappingofLimit" ,false
       Exit Function
   End If
   waitForIcallLoading
   clickButtonCappingofLimit=true
End Function

'[Select Combobox Capping Type as]
Public Function selectCappingTypeComboBox(strCappingType)
   bDevPending=false
   bSelectCappingTypeComboBox=true
   waitForIcallLoading
   If Not IsNull(strCappingType) Then
       If Not (selectItem_Combobox (CappingLimit_Page.lstCappingType(), strCappingType))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From CappingType drop down list" ,false
           bSelectCappingTypeComboBox=false
       End If
   End If
   waitForIcallLoading
   selectCappingTypeComboBox=bSelectCappingTypeComboBox
End Function

'[Select Combobox Capping Reason as]
Public Function selectCappingReasonComboBox(strReasonForCapping)
   bDevPending=false
   bselectCappingReasonComboBox=true
   waitForIcallLoading
   If Not IsNull(ReasonForCapping) Then
       If Not (selectItem_Combobox (CappingLimit_Page.lstReasonForCapping(), strReasonForCapping))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From CappingType drop down list" ,false
           selectCappingReasonComboBox=false
       End If
   End If
   waitForIcallLoading
   selectCappingReasonComboBox=bselectCappingReasonComboBox
End Function

'[Select Combobox Capping Option as]
Public Function selectCappingOptionComboBox(strCappingOption)
   bDevPending=false
   bselectCappingOptionComboBox=true
   waitForIcallLoading
   If Not IsNull(CappingOption) Then
       If Not (selectItem_Combobox (CappingLimit_Page.lstCappingOption(), strCappingOption))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From CappingOption drop down list" ,false
           selectCappingOptionComboBox=false
       End If
   End If
   waitForIcallLoading
   selectCappingOptionComboBox=bselectCappingOptionComboBox
End Function

'[Get selected item from combo box CappingType]
Public Function getCappingTypeSelectedItem()
   'bDevPending=false
   getCappingTypeSelectedItem=getVadinCombo_SelectedItem(CappingLimit_Page.lstCappingType)
End Function

'[Verify Combobox CappingType displayed as]
Public Function verifyCappingTypeText(strExpectedText)
   'bDevPending=false
   bVerifyCappingTypeText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (CappingLimit_Page.lstCappingType(), strExpectedText, "CappingType")Then
           bVerifyCappingTypeText=false
       End If
   End If
   verifyCappingTypeText=bVerifyCappingTypeText
End Function

'[Get CappingInlineMessage Label Text]
Public Function getCappingInlineMessageText()
   bDevPending=False
   getCappingInlineMessageText=CappingLimit_Page.lblCappingInlineMessage.GetRoProperty("innertext")
End Function

'[Verify Field Capping Inline Message displayed as]
Public Function verifyCappingInlineMessageText(strCappingInlineMessage)
   bDevPending=false
   bVerifyCappingInlineMessageText=true
   WaitForICallLoading   
   For iCount = 1 To 180 Step 1
		If Not CappingLimit_Page.lblCappingInlineMessage.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(strCappingInlineMessage) Then
				If Not VerifyInnerText (CappingLimit_Page.lblCappingInlineMessage(), strCappingInlineMessage, "CappingInlineMessage")Then
					bVerifyCappingInlineMessageText=false
				End If
			End If
			Exit for
		End if
	Next
	WaitForICallLoading
   verifyCappingInlineMessageText=bVerifyCappingInlineMessageText
End Function

'[Get amount for validation from Balance and Limit page]
Public Function getAmount()
	strRunTimeRtlCrLimit=BalancesAndLimits.lblRelationship_CreditLimit.GetROProperty("innertext")
	strRunTimeAcctTotalCrLimit=BalancesAndLimits.lblTotalCreditLimit.GetROProperty("innertext")
	strRunTimeEmbosserCrLimit=BalancesAndLimits.lblCardLimit_CreditLimit.GetROProperty("innertext")

	strRunTimeRtlOutStandingBal=BalancesAndLimits.lblRelationship_OutstandingBalance.GetROProperty("innertext")
	strRunTimeAcctOutStandingBal=BalancesAndLimits.lblOutstandingBalance.GetROProperty("innertext")
	
	strAvlEmbosserCrLimit=BalancesAndLimits.lblCardLimit_AvailableLimit.GetROProperty("innertext")
	'***** For Embosser outstanding. If CrLimit>0 then CrLimit-AvlCrLimit. If CrLimit=0 then 0.
	If strRunTimeEmbosserCrLimit>0 Then
		strRunTimeEmbosserOutStandingBal=strRunTimeEmbosserCrLimit-strAvlEmbosserCrLimit
	Else
		strRunTimeEmbosserOutStandingBal=""	
	End If		
End Function

'[Verify Table SelectedCardsContent displayed]
Public Function verifySelectedCardsContentTabledisplayed()
   bDevPending=false
   verifySelectedCardsContentdisplayed= CappingLimit_Page.tblSelectedCardsContent.Exist(1)
End Function

'[Verify Table SelectedCardsContent has following Columns]
Public Function verifySelectedCardsContentTableColumns(arrColumnNameList)
   bDevPending=false
   verifySelectedCardsContentTableColumns=verifyTableColumns(CappingLimit_Page.tblSelectedCardsContent,arrColumnNameList)
End Function

'[Verify row Data in Table SelectedCards for Capping Limit]
Public Function verifytblSelectedCardsContent_CL(arrRowDataList)
   bDevPending=false
   WaitForICallLoading
   verifytblSelectedCardsContent_CL=verifyTableContentList(CappingLimit_Page.tblSelectedCardsHeader,CappingLimit_Page.tblSelectedCardsContent,arrRowDataList,"SelectedCardsContent" , false,null ,null,null)
   WaitForICallLoading
End Function

'[Click <Column Name> link in Table SelectedCardsContent]
Public Function clickSelectedCardsContent_link(arrRowDataList)
   bDevPending=false
   clickSelectedCardsContent_link=selectTableLink(CappingLimit_Page.tblSelectedCardsContentHeader,CappingLimit_Page.tblSelectedCardsContentContent,arrRowDataList,"SelectedCardsContent" ,"Column Name",bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1 ,CappingLimit_Page.lnkPrevious)
End Function

'[Verify Table SelectedCardsHeader displayed]
Public Function verifySelectedCardsHeaderTabledisplayed()
   bDevPending=false
   verifySelectedCardsHeaderdisplayed= CappingLimit_Page.tblSelectedCardsHeader.Exist(1)
End Function

'[Verify Table SelectedCardsHeader has following Columns]
Public Function verifySelectedCardsHeaderTableColumns(arrColumnNameList)
   bDevPending=false
   verifySelectedCardsHeaderTableColumns=verifyTableColumns(CappingLimit_Page.tblSelectedCardsHeader,arrColumnNameList)
End Function

'[Verify row Data in Table SelectedCardsHeader]
Public Function verifytblSelectedCardsHeader_RowData(arrRowDataList)
   bDevPending=false
   verifytblSelectedCardsHeader_RowData=verifyTableContentList(CappingLimit_Page.tblSelectedCardsHeaderHeader,CappingLimit_Page.tblSelectedCardsHeaderContent,arrRowDataList,"SelectedCardsHeader" , bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1,CappingLimit_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table SelectedCardsHeader]
Public Function clickSelectedCardsHeader_link(arrRowDataList)
   bDevPending=false
   clickSelectedCardsHeader_link=selectTableLink(CappingLimit_Page.tblSelectedCardsHeaderHeader,CappingLimit_Page.tblSelectedCardsHeaderContent,arrRowDataList,"SelectedCardsHeader" ,"Column Name",bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1 ,CappingLimit_Page.lnkPrevious)
End Function

'[Verify Table CurrentCreditLimitContent displayed]
Public Function verifyCurrentCreditLimitContentTabledisplayed()
   bDevPending=false
   verifyCurrentCreditLimitContentdisplayed= CappingLimit_Page.tblCurrentCreditLimitContent.Exist(1)
End Function

'[Verify Table CurrentCreditLimitContent has following Columns]
Public Function verifyCurrentCreditLimitContentTableColumns(arrColumnNameList)
   bDevPending=false
   verifyCurrentCreditLimitContentTableColumns=verifyTableColumns(CappingLimit_Page.tblCurrentCreditLimitContent,arrColumnNameList)
End Function

'[Verify row Data in Table Current Credit Limit for Cappling Limit]
Public Function verifytblCurrentCreditLimit_CL(arrRowDataList)
   bDevPending=false
   bverifytblCurrentCreditLimit_CL=true
  Dim arrColumns,arrValues,intSize
  
  For iRowCount=0 to Ubound(arrRowDataList,1)
	  intSize=Ubound(arrRowDataList,2)
	ReDim arrColumns(intSize)
	ReDim arrValues(intSize)
'	For iCount=0 to intSize
'		arrTemp=Split(arrRowDataList(iRowCount,iCount),":")
'		arrColumns(iCount)=arrTemp(0)
'		If arrTemp(1) = "RUNTIME_Relationship" Then
'		   arrValues(iCount)=	checkNull(replace (arrTemp(1),"RUNTIME_Relationship",strRunTimeRtlCrLimit))
'		ElseIf arrTemp(1) = "RUNTIME_Account" Then
'		   arrValues(iCount)=	checkNull(replace (arrTemp(1),"RUNTIME_Account",strRunTimeAcctTotalCrLimit))
'		ElseIf arrTemp(1) = "RUNTIME_Embosser" Then
'		   arrValues(iCount)=	checkNull(replace (arrTemp(1),"RUNTIME_Embosser",strRunTimeEmbosserCrLimit))		   
'		Else
'		   arrValues(iCount)=checkNull(arrTemp(1))	
'		End If										
'	Next  
	
	Dim lstlstCurrentCreditLimit
    'lstlstCurrentCreditLimit (0)="(Relationship:"&strRunTimeRtlCrLimit
    'lstlstCurrentCreditLimit (1)="Account:"& strRunTimeAcctTotalCrLimit
    'lstlstCurrentCreditLimit (2)="Card:"& strRunTimeEmbosserCrLimit&")|"
    lstlstCurrentCreditLimit=CheckNull("(Relationship:"&strRunTimeRtlCrLimit&"|Account:"&strRunTimeAcctTotalCrLimit&"|Card:"&strRunTimeEmbosserCrLimit&")|")
	
   ' verifytblCurrentCreditLimitContent_CL=verifyTableContentList(CappingLimit_Page.tblCurrentCreditLimitContentHeader,CappingLimit_Page.tblCurrentCreditLimitContentContent,arrRowDataList,"CurrentCreditLimitContent" , bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1,CappingLimit_Page.lnkPrevious)
	
	intRow=verifyTableContentList(CappingLimit_Page.tblCurrentCreditLimitHeader,CappingLimit_Page.tblCurrentCreditLimitContent,lstlstCurrentCreditLimit,"CurrentCurrentLimit" , false,null,null,null)
	
	'intRow=getRowForColumns(CappingLimit_Page.tblCurrentCreditLimitHeader,CappingLimit_Page.tblCurrentCreditLimitContent,arrColumns,arrValues)
	If Not intRow Then
		LogMessage "WARN","Verification","Expected Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" not found in table",false
		bverifytblCurrentCreditLimit_CL= False
	else
		LogMessage "RSLT","Verification","Expected Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" found in table at Row "&intRow&" in table",true
		bverifytblCurrentCreditLimit_CL= True
	End If
  Next
  verifytblCurrentCreditLimit_CL=bverifytblCurrentCreditLimit_CL  
End Function

'[Click <Column Name> link in Table CurrentCreditLimitContent]
Public Function clickCurrentCreditLimitContent_link(arrRowDataList)
   bDevPending=false
   clickCurrentCreditLimitContent_link=selectTableLink(CappingLimit_Page.tblCurrentCreditLimitContentHeader,CappingLimit_Page.tblCurrentCreditLimitContentContent,arrRowDataList,"CurrentCreditLimitContent" ,"Column Name",bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1 ,CappingLimit_Page.lnkPrevious)
End Function

'[Verify Table CurrentCreditLimitHeader displayed]
Public Function verifyCurrentCreditLimitHeaderTabledisplayed()
   bDevPending=false
   verifyCurrentCreditLimitHeaderdisplayed= CappingLimit_Page.tblCurrentCreditLimitHeader.Exist(1)
End Function

'[Verify Table CurrentCreditLimitHeader has following Columns]
Public Function verifyCurrentCreditLimitHeaderTableColumns(arrColumnNameList)
   bDevPending=false
   verifyCurrentCreditLimitHeaderTableColumns=verifyTableColumns(CappingLimit_Page.tblCurrentCreditLimitHeader,arrColumnNameList)
End Function

'[Verify row Data in Table CurrentCreditLimitHeader]
Public Function verifytblCurrentCreditLimitHeader_RowData(arrRowDataList)
   bDevPending=false
   verifytblCurrentCreditLimitHeader_RowData=verifyTableContentList(CappingLimit_Page.tblCurrentCreditLimitHeaderHeader,CappingLimit_Page.tblCurrentCreditLimitHeaderContent,arrRowDataList,"CurrentCreditLimitHeader" , bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1,CappingLimit_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table CurrentCreditLimitHeader]
Public Function clickCurrentCreditLimitHeader_link(arrRowDataList)
   bDevPending=false
   clickCurrentCreditLimitHeader_link=selectTableLink(CappingLimit_Page.tblCurrentCreditLimitHeaderHeader,CappingLimit_Page.tblCurrentCreditLimitHeaderContent,arrRowDataList,"CurrentCreditLimitHeader" ,"Column Name",bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1 ,CappingLimit_Page.lnkPrevious)
End Function

'[Verify Table OutstandingBalancesContent displayed]
Public Function verifyOutstandingBalancesContentTabledisplayed()
   bDevPending=false
   verifyOutstandingBalancesContentdisplayed= CappingLimit_Page.tblOutstandingBalancesContent.Exist(1)
End Function

'[Verify Table OutstandingBalancesContent has following Columns]
Public Function verifyOutstandingBalancesContentTableColumns(arrColumnNameList)
   bDevPending=false
   verifyOutstandingBalancesContentTableColumns=verifyTableColumns(CappingLimit_Page.tblOutstandingBalancesContent,arrColumnNameList)
End Function

'[Verify row Data in Table Outstanding Balances for Capping Limit]
Public Function verifytblOutstandingBalances_CL(arrRowDataList)
   bDevPending=false
   bverifytblOutstandingBalances_CL=true
  Dim arrColumns,arrValues,intSize
  
  For iRowCount=0 to Ubound(arrRowDataList,1)
	  intSize=Ubound(arrRowDataList,2)
	'arrTemp=arrPlanData(iRowCount)
	ReDim arrColumns(intSize)
	ReDim arrValues(intSize)
'	For iCount=0 to intSize
'		arrTemp=Split(arrRowDataList(iRowCount,iCount),":")
'		arrColumns(iCount)=arrTemp(0)
'		If arrTemp(1) = "RUNTIME_Relationship" Then
'		   arrValues(iCount)=	checkNull(replace (arrTemp(1),"RUNTIME_Relationship",strRunTimeRtlCrLimit))
'		ElseIf arrTemp(1) = "RUNTIME_Account" Then
'		   arrValues(iCount)=	checkNull(replace (arrTemp(1),"RUNTIME_Account",strRunTimeAcctTotalCrLimit))
'		ElseIf arrTemp(1) = "RUNTIME_Embosser" Then
'		   arrValues(iCount)=	checkNull(replace (arrTemp(1),"RUNTIME_Embosser",strRunTimeEmbosserCrLimit))		   
'		Else
'		   arrValues(iCount)=checkNull(arrTemp(1))	
'		End If										
'	Next  
	
	Dim lstlstOutstandingBalance
    'lstOutstandingBalance (0)="Relationship:"&strRunTimeRtlOutStandingBal
    'lstOutstandingBalance (1)="Account:"& strRunTimeAcctOutStandingBal
    'lstOutstandingBalance (2)="Card:"& strRunTimeEmbosserCrLimit
    
    lstlstOutstandingBalance=CheckNull("(Relationship:"&strRunTimeRtlOutStandingBal&"|Account:"&strRunTimeAcctOutStandingBal&"|Card:"&strRunTimeEmbosserCrLimit&")|")
    
    If lstlstOutstandingBalance(0,2) = "Card:0.00" Then
    	lstlstOutstandingBalance(0,2) = "Card:"&EMPTY
    End If
	
	intRow=verifyTableContentList(CappingLimit_Page.tblOutstandingBalancesHeader,CappingLimit_Page.tblOutstandingBalancesContent,lstlstOutstandingBalance,"OutstandingBalance" , false,null,null,null)
	
	
   'verifytblOutstandingBalancesContent_CL=verifyTableContentList(CappingLimit_Page.tblOutstandingBalancesContentHeader,CappingLimit_Page.tblOutstandingBalancesContentContent,arrRowDataList,"OutstandingBalancesContent" , bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1,CappingLimit_Page.lnkPrevious)
	'intRow=getRowForColumns(CappingLimit_Page.tblOutstandingBalancesHeader,CappingLimit_Page.tblOutstandingBalancesContent,arrColumns,arrValues)
	If intRow =-1  Then
		LogMessage "WARN","Verification","Expected Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" not found in table",false
		bverifytblOutstandingBalances_CL= False
	else
		LogMessage "RSLT","Verification","Expected Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" found in table at Row "&intRow&" in table" ,true
		bverifytblOutstandingBalances_CL= True
	End If
  Next
  verifytblOutstandingBalances_CL=bverifytblOutstandingBalances_CL   
End Function

'[Click <Column Name> link in Table OutstandingBalancesContent]
Public Function clickOutstandingBalancesContent_link(arrRowDataList)
   bDevPending=false
   clickOutstandingBalancesContent_link=selectTableLink(CappingLimit_Page.tblOutstandingBalancesContentHeader,CappingLimit_Page.tblOutstandingBalancesContentContent,arrRowDataList,"OutstandingBalancesContent" ,"Column Name",bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1 ,CappingLimit_Page.lnkPrevious)
End Function

'[Verify Table OutstandingBalancesHeader displayed]
Public Function verifyOutstandingBalancesHeaderTabledisplayed()
   bDevPending=false
   verifyOutstandingBalancesHeaderdisplayed= CappingLimit_Page.tblOutstandingBalancesHeader.Exist(1)
End Function
'[Verify Table OutstandingBalancesHeader has following Columns]
Public Function verifyOutstandingBalancesHeaderTableColumns(arrColumnNameList)
   bDevPending=false
   verifyOutstandingBalancesHeaderTableColumns=verifyTableColumns(CappingLimit_Page.tblOutstandingBalancesHeader,arrColumnNameList)
End Function
'[Verify row Data in Table OutstandingBalancesHeader]
Public Function verifytblOutstandingBalancesHeader_RowData(arrRowDataList)
   bDevPending=false
   verifytblOutstandingBalancesHeader_RowData=verifyTableContentList(CappingLimit_Page.tblOutstandingBalancesHeaderHeader,CappingLimit_Page.tblOutstandingBalancesHeaderContent,arrRowDataList,"OutstandingBalancesHeader" , bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1,CappingLimit_Page.lnkPrevious)
End Function

'[Click <Column Name> link in Table OutstandingBalancesHeader]
Public Function clickOutstandingBalancesHeader_link(arrRowDataList)
   bDevPending=false
   clickOutstandingBalancesHeader_link=selectTableLink(CappingLimit_Page.tblOutstandingBalancesHeaderHeader,CappingLimit_Page.tblOutstandingBalancesHeaderContent,arrRowDataList,"OutstandingBalancesHeader" ,"Column Name",bPagination,CappingLimit_Page.lnkNext ,CappingLimit_Page.lnkNext1 ,CappingLimit_Page.lnkPrevious)
End Function

'[Select Combobox ReasonForCapping as]
Public Function selectReasonForCappingComboBox(strReasonForCapping)
   bDevPending=false
   bSelectReasonForCappingComboBox=true
   If Not IsNull(strReasonForCapping) Then
       If Not (selectItem_Combobox (CappingLimit_Page.lstReasonForCapping(), strReasonForCapping))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From ReasonForCapping drop down list" ,false
           bSelectReasonForCappingComboBox=false
       End If
   End If
   selectReasonForCappingComboBox=bSelectReasonForCappingComboBox
End Function

'[Get selected item from combo box ReasonForCapping]
Public Function getReasonForCappingSelectedItem()
   bDevPending=false
   getReasonForCappingSelectedItem=getVadinCombo_SelectedItem(CappingLimit_Page.lstReasonForCapping)
End Function

'[Verify Combobox Reason For Capping displayed as]
Public Function verifyReasonForCappingText(strExpectedText)
   bDevPending=false
   bVerifyReasonForCappingText=true
   If Not IsNull(strExpectedText) Then
       If Not verifyComboSelectItem (CappingLimit_Page.lstReasonForCapping(), strExpectedText, "Reason For Capping")Then
           bVerifyReasonForCappingText=false
       End If
   End If
   verifyReasonForCappingText=bVerifyReasonForCappingText
End Function

'[Set TextBox Capped Amount on Capping Limit Screen to]
Public Function setCappedAmountTextbox_CL(strCappedAmount)
   bDevPending=false
   If not isNull(strCappedAmount) Then
	   CappingLimit_Page.txtCappedAmount.Set(strCappedAmount)
   End If
   If Err.Number<>0 Then
       setCappedAmountTextbox_CL=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Capped Amount" ,false
       Exit Function
   End If
   WaitForICallLoading
   setCappedAmountTextbox_CL=true
End Function

'[Get Description Label Text]
Public Function getDescriptionText()
   bDevPending=false
   getDescriptionText=CappingLimit_Page.lblDescription.GetRoProperty("innertext")
End Function

'[Verify Field Description displayed on Capping Limit Screen as]
Public Function verifyDescriptionText_CL(strExpectedText)
   bDevPending=false
   bVerifyDescriptionText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CappingLimit_Page.lblDescription(), strExpectedText, "Description")Then
           bVerifyDescriptionText=false
       End If
   End If
   verifyDescriptionText_CL=bVerifyDescriptionText
End Function

'[Click Link KnowledgeBase]
Public Function clickLinkKnowledgeBase()
   bDevPending=false
   CappingLimit_Page.lnkKnowledgeBase.click
   If Err.Number<>0 Then
       clickLinkKnowledgeBase=false
            LogMessage "WARN","Verification","Failed to Click Link : KnowledgeBase" ,false
       Exit Function
   End If
   clickLinkKnowledgeBase=true
End Function

'[Verify Field KnowledgeBase on Capping Limit SR Screen displayed as]
Public Function verifyKnowledgeBase_CL(strExpectedLink)
   bDevPending=false
   bVerifyKnowledgeBaseText=true
   If Not IsNull(strExpectedLink) Then
		
		Set oDesc_KB = Description.Create()
			'oDesc_KB("micclass").Value = "Link"
		
			'strKBLink=CappingLimit_Page.lnkKnowledgeBase.ChildObjects(oDesc_KB)(0).GetROProperty("href")
			strKBLink=CappingLimit_Page.lnkKnowledgeBase.GetROProperty("href")
			strExpectedLink=Replace(strExpectedLink,"@","=")
       If not MatchStr(strKBLink, strExpectedLink)Then
		   LogMessage "RSLT","Verification","Knowledge base link does not matched with expected. Actual : "&strKBLink&" Expected "&strExpectedLink,false
           bVerifyKnowledgeBaseText=false
	   else
	 		LogMessage "RSLT","Verification","Knowledge base link matrched with expected",true
       End If
   End If
   verifyKnowledgeBase_CL=bVerifyKnowledgeBaseText
End Function

'[Click Button AddNotes]
Public Function clickButtonAddNotes()
   bDevPending=false
   CappingLimit_Page.btnAddNotes.click
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
   getCommentText=CappingLimit_Page.txtComment.GetRoProperty("innertext")
End Function

'[Verify Field Comment displayed as]
Public Function verifyCommentText(strExpectedText)
   bDevPending=false
   bVerifyCommentText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( CappingLimit_Page.txtComment(), strExpectedText, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText=bVerifyCommentText
End Function


'[Set TextBox on Capping Limit Comment to]
Public Function setCommentTextbox_CL(strComment)
   bDevPending=false
   CappingLimit_Page.txtComment.Set(strComment)
   If Err.Number<>0 Then
       setCommentTextbox_CL=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_CL=true
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
   getNotesText=CappingLimit_Page.txtNotes.GetRoProperty("innertext")
End Function

'[Verify Field Notes displayed as]
Public Function verifyNotesText(strExpectedText)
   bDevPending=false
   bVerifyNotesText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyField( CappingLimit_Page.txtNotes(), strExpectedText, "Notes")Then
           bVerifyNotesText=false
       End If
   End If
   verifyNotesText=bVerifyNotesText
End Function

'[Perform Add Notes by clicking Add Notes Button on Capping Limit Screen]
Public Function addNote_CL(strNote)
   bDevPending=false
   baddNote_CL=true
	Dim baddNote_CL:addNote_CL=true
	
	If not isNull(strNote) Then
		CappingLimit_Page.btnAddNotes.click
		WaitForICallLoading
           If Not CappingLimit_Page.popupNotes.exist(5)Then
				LogMessage "WARN","Verification","New Note dialog Box displayed",false
				baddNote_CL=false
			 else
			 strMessage=CappingLimit_Page.lblMaxAllowed.GetROProperty("innerText")
				If not strMessage="Max allowed - 3000" Then
					LogMessage "WARN","Verification","Add New Comment popup dislog incorrectly displayed max allowed character count for comment. Expected : Max allowed - 3000 and Actual: "&strMessage, false
					baddNote_CL=false					
				End If
			   'ServiceRequest.txtNewComment.set strNote			   
			   CappingLimit_Page.txtNotesDescription.set strNote
			   CappingLimit_Page.btnNotesSave.Click
				   'ServiceRequest.clickSave_Popup
			  WaitForIcallLoading
		   End If 
		End If 
	addNote_CL=baddNote_CL
End Function

'[Verify Button AddNote is disabled on Capping Limit Screen]
Public Function VerifybtnAddNoteDisable_CL()
	bDevPending=False
   Dim bVerifybtnSubmit_CL:bVerifybtnSubmit_CL=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	intBtnSubmit=Instr(CappingLimit_Page.btnAddNotes.Object.GetAttribute("disabled"),("disabled"))
	If  not intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Add Note button is disable as per expectation.",True
		bVerifybtnSubmit_CL=true
	Else
		LogMessage "WARN","Verifiation","Add Note button is enable. Expected to be disable.",false
		bVerifybtnSubmit_CL=false
	End If
	VerifybtnAddNoteDisable_CL=bVerifybtnSubmit_CL
End Function

'[Click Button OK_ValidationMsg]
Public Function clickButtonOK_ValidationMsg()
   bDevPending=false
   CappingLimit_Page.btnOK_ValidationMsg.click
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
   getValidationMessageText=CappingLimit_Page.lblValidationMessage.GetRoProperty("innertext")
End Function

'[Verify Field ValidationMessage displayed as]
Public Function verifyValidationMessageText(strExpectedText)
   bDevPending=false
   bVerifyValidationMessageText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CappingLimit_Page.lblValidationMessage(), strExpectedText, "ValidationMessage")Then
           bVerifyValidationMessageText=false
       End If
   End If
   verifyValidationMessageText=bVerifyValidationMessageText
End Function

'[Verify Button Cancel is enabled on Capping Limit Screen]
Public Function VerifybtnCancel_CL()
	bDevPending=False
   Dim bVerifybtnCancel_CL:bVerifybtnCancel_CL=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	intBtnSubmit=Instr(CappingLimit_Page.btnCancel.GetROproperty("outerhtml"),("v-disabled"))
	If  intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Cancel button is enable as per expectation.",True
		bVerifybtnCancel_CL=true
	Else
		LogMessage "WARN","Verifiation","Cancel button is disable. Expected to be enable.",false
		bVerifybtnCancel_CL=false
	End If
	VerifybtnCancel_CL=bVerifybtnCancel_CL
End Function

'[Click Button Cancel]
Public Function clickButtonCancel_CL()
   bDevPending=false
   CappingLimit_Page.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_CL=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonCancel_CL=true
End Function

'[Verify Button Submit is enabled on Capping Limit Screen]
Public Function VerifybtnSubmit_CL()
	bDevPending=False
   Dim bVerifybtnSubmit_CL:bVerifybtnSubmit_CL=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	intBtnSubmit=Instr(CappingLimit_Page.btnSubmit.GetROproperty("outerhtml"),("v-disabled"))
	If  intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Submit button is enable as per expectation.",True
		bVerifybtnSubmit_CL=true
	Else
		LogMessage "WARN","Verifiation","Submit button is disable. Expected to be enable.",false
		bVerifybtnSubmit_CL=false
	End If
	VerifybtnSubmit_CL=bVerifybtnSubmit_CL
End Function

'[Verify Button Submit is disabled on Capping Limit Screen]
Public Function VerifybtnSubmitDisable_CL()
	bDevPending=False
   Dim bVerifybtnSubmit_CL:bVerifybtnSubmit_CL=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	'intBtnSubmit=Instr(CappingLimit_Page.btnSubmit.GetROproperty("outerhtml"),("v-disabled"))
	intBtnSubmit=Instr(CappingLimit_Page.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
	If  not intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Submit button is disable as per expectation.",True
		bVerifybtnSubmit_CL=true
	Else
		LogMessage "WARN","Verifiation","Submit button is enable. Expected to be disable.",false
		bVerifybtnSubmit_CL=false
	End If
	VerifybtnSubmitDisable_CL=bVerifybtnSubmit_CL
End Function

'[Click Button Submit on Capping Limit]
Public Function clickButtonSubmit_CL()
   bDevPending=false
   For iCountOne = 1 To 180 Step 1
		If Not CappingLimit_Page.btnSubmit.Exist(0.5) Then
			Wait(0.5)
		else
			CappingLimit_Page.btnSubmit.click
			Exit for
		End if
	Next
   If Err.Number<>0 Then
       clickButtonSubmit_CL=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
   WaitForIcallLoading
   clickButtonSubmit_CL=true
End Function

'[Verify Confirmation Popup on Capping Limit]
Public Function verifyConfirmationPopup_CL()
	verifyConfirmationPopup_CL=true
	If Not verifyInnerText(CappingLimit_Page.lblConfirmationMsg(), "Are you sure you want to discard the changes (if any) and leave this page?", "Verify Pop up confirmation") Then
		verifyConfirmationPopup_CL=false
	End If
	CappingLimit_Page.btnYes_Confirmation.click
	  If Err.Number<>0 Then
       verifyConfirmationPopup_CL=false
            LogMessage "WARN","Verification","Failed to Click Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
End Function

'[Verify Capping Type Combobox has Items]
Public Function verifyCappingTypeComboboxItems(lstItems)
   bDevPending=false
   bverifyCappingTypeComboboxItems=true
   If Not IsNull(lstItems) Then
       If Not verifyComboboxItems (CappingLimit_Page.lstCappingType(), lstItems, "Capping Type")Then
           bverifyCappingTypeComboboxItems=false
       End If
   End If
   verifyCappingTypeComboboxItems=bverifyCappingTypeComboboxItems
End Function

'[Verify Capping Option Combobox has Items]
Public Function verifyCappingOptionComboboxItems(lstItems)
   bDevPending=false
   bverifyCappingTypeComboboxItems=true
   WaitForIcallLoading
   For iCountTre = 1 To 180 Step 1
		If Not CappingLimit_Page.lstReasonForCapping.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(lstItems) Then
				If Not verifyComboboxItems (CappingLimit_Page.lstReasonForCapping(), lstItems, "Capping Option")Then
					bverifyCappingTypeComboboxItems=false
				End If
			End If
			Exit for
		End if
	Next  
   verifyCappingOptionComboboxItems=bverifyCappingTypeComboboxItems
End Function

'[Verify Capping Option is not available]
Public Function verifyCappingOptionExist(lstItems)
   bDevPending=false
   bverifyCappingOptionExist=true
   If (CappingLimit_Page.lstCappingOption().Exist) Then
      LogMessage "WARN","Verification","Capping Option is available on screen. Expected should not available on screen" ,false
 	  bverifyCappingOptionExist=false
   End If 
   verifyCappingOptionExist=bverifyCappingOptionExist
End Function

'[Verify Reason for Capping Combobox has Items]
Public Function verifyReasonForCappingComboboxItems(lstReasonForCapping)
   bDevPending=false
   bverifyReasonForCappingComboboxItems=true
   For iCountTwo = 1 To 180 Step 1
		If Not CappingLimit_Page.lstReasonForCapping.Exist(0.5) Then
			Wait(0.5)
		else
			If Not IsNull(lstReasonForCapping) Then
				If Not verifyComboboxItems (CappingLimit_Page.lstReasonForCapping(), lstReasonForCapping, "Reason For Capping")Then
					bverifyReasonForCappingComboboxItems=false
				End If
			End If
			Exit for
		End if
	Next   
   verifyReasonForCappingComboboxItems=bverifyReasonForCappingComboboxItems
End Function

'[Verify popup Message displayed on Capping Limit screen]
Public Function verifyPopupValidationMessage_CL(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (CappingLimit_Page.lblValidationMessage(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
       CappingLimit_Page.btnOK_ValidationMsg.Click
   End If
   
   WaitForIcallLoading
   verifyPopupValidationMessage_CL=bVerifyValidationMessageText
End Function

'[Verify Validation Message displayed on Capping Limit as]
Public Function verifyValidationMessage_CL(strValidationMessage)
   bDevPending=False
   bVerifyValidationMessageText=true
   If Not IsNull(strValidationMessage) Then
       If Not VerifyInnerText (CappingLimit_Page.lblCappingLimitInlineMessage(), strValidationMessage, "Validation Message")Then
           bVerifyValidationMessageText=false
       End If
   End If
   If (CappingLimit_Page.btnOK_ValidationMsg().Exist) Then
   	CappingLimit_Page.btnOK_ValidationMsg.Click
   End If
   
   WaitForIcallLoading
   verifyValidationMessage_CL=bVerifyValidationMessageText
End Function

'[Verify Popup Request Submitted exist for Capping Limit]
Public Function verifyPopupRequestSubmitted_CL(bExist)
   bDevPending=false
   bActualExist=CappingLimit_Page.popupRequestSubmitted.Exist(4)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted Exists As Expected" ,true
       verifyPopupRequestSubmitted_CL=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :RequestSubmitted does not Exists As Expected" ,true
       verifyPopupRequestSubmitted_CL=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted does not Exists As Expected" ,False
       verifyPopupRequestSubmitted_CL=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :RequestSubmitted Still Exists" ,False
       verifyPopupRequestSubmitted_CL=False
   End If
End Function

'[Verify Field CardNumber on Request Submitted Popup for Capping Limit displayed as]
Public Function verifyCardNumber_RequestSubmitted_CL(strCardNumber)
   bDevPending=false
   bVerifyCardNumber_RequestSubmittedText=true
   insertDataStore "NewSAUsedCard", ""&strCardNumber
   If Not IsNull(strCardNumber) Then
       If Not VerifyInnerText (CappingLimit_Page.lblCardNumber_RequestSubmitted(), strCardNumber, "CardNumber_RequestSubmitted")Then
           bVerifyCardNumber_RequestSubmittedText=false
       End If
   End If
   verifyCardNumber_RequestSubmitted_CL=bVerifyCardNumber_RequestSubmittedText
End Function

'[Verify Field ProductDescription on Request Submitted Popup for Capping Limit displayed as]
Public Function verifyProductDescription_RequestSubmitted_CL(strProductDescription)
   bDevPending=false
   bVerifyProductDescription_RequestSubmittedText=true
   If Not IsNull(strProductDescription) Then
       If Not VerifyInnerText (CappingLimit_Page.lblProductDescription_RequestSubmitted(), strProductDescription, "ProductDescription_RequestSubmitted")Then
           bVerifyProductDescription_RequestSubmittedText=false
       End If
   End If
   verifyProductDescription_RequestSubmitted_CL=bVerifyProductDescription_RequestSubmittedText
End Function

'[Click Button RefreshStatus For Capping Limit]
Public Function clickButtonRefreshStatus_CL()
   bDevPending=false
   CappingLimit_Page.btnRefreshStatus.click
	WaitForICallLoading
    		'Get Status
		If CappingLimit_Page.lblStatus_RequestSubmitted.getROProperty("innertext")="In Progress" then 
			bStatus=true
		 else
			bStatus=false
		End If		
	
	While  bStatus AND (iCount<60)
		CappingLimit_Page.btnRefreshStatus.click
		wait 1
        	'Get Status
			strStatus=CappingLimit_Page.lblStatus_RequestSubmitted.getROProperty("innertext")
			If Trim(strStatus)="In Progress" then 
				bStatus=true
			 else
				LogMessage "WARN","Verification","Status displayed as  :"&strStatus ,true
				bStatus=false
			End If
		wait 5
		intBtnRefreshStatus=Instr(CappingLimit_Page.btnRefreshStatus.GetROproperty("outerhtml"),"v-disabled")
		If intBtnRefreshStatus<>0 Then
			LogMessage "WARN","Verification","Button : RefreshStatus is disabled" ,true
			bStatust=true
		End If
		iCount=iCount+1
	  Wend	

   If Err.Number<>0 Then
       
            LogMessage "WARN","Verification","Failed to Click Button : RefreshStatus" ,false
			clickButtonRefreshStatus_CL=false
       Exit Function
   End If
   WaitForICallLoading
  
   clickButtonRefreshStatus_CL=true
End Function

'[Verify Field Status_RequestSubmitted For Capping Limit displayed as]
Public Function verifyStatus_RequestSubmitted_CL(strExpectedText)
   bDevPending=false
   bVerifyStatus_RequestSubmittedText=true
   If Not IsNull(strExpectedText) Then
       If Not VerifyInnerText (CappingLimit_Page.lblStatus_RequestSubmitted(), strExpectedText, "Status_RequestSubmitted")Then
           bVerifyStatus_RequestSubmittedText=false
       End If
   End If
   verifyStatus_RequestSubmitted_CL=bVerifyStatus_RequestSubmittedText
End Function

'[Click Link SRNumber on Request Submitted Popup for Capping Limit]
Public Function clickLinkSRNumber_RequestSubmitted_CL()
   bDevPending=false
   gstrRuntimeSRNumStep="Click Link SRNumber on Request Submitted Popup for Capping Limit"
   strSelectedSR=CappingLimit_Page.lnkSRNumber_RequestSubmitted.GetRoProperty("innerText")
	If strSelectedSR<>"" Then
		 insertDataStore "SelectedSRLink", strSelectedSR	
	   CappingLimit_Page.lnkSRNumber_RequestSubmitted.click
	 else
   		LogMessage "RSLT","Verification","SR Number did not displayed on Request Submitted pop up",false
	End If
   WaitForIcallLoading
   If Err.Number<>0 Then
       clickLinkSRNumber_RequestSubmitted_CL=false
            LogMessage "WARN","Verification","Failed to Click Link : SRNumber_RequestSubmitted" ,false
       Exit Function
   End If
   clickLinkSRNumber_RequestSubmitted_CL=true
End Function

'[Click Close button on Request Submitted Popup for Capping Limit]
Public Function verifybtnClose_RequestSubmitted_CL()
	bverifybtnClose_RequestSubmitted_CL=true
	CappingLimit_Page.btnCancel_RequestSubmitted.click
	  If Err.Number<>0 Then
       bverifybtnClose_RequestSubmitted_CL=false
            LogMessage "WARN","Verification","Failed to Click Close Button : Yes on Confirmation popup" ,false
       Exit Function
   End If
   WaitForICallLoading
	verifybtnClose_RequestSubmitted_CL=bverifybtnClose_RequestSubmitted_CL
End Function

'[Verify Confirmation Popup content]
Public Function verifyConfirmationPopup_Content_CL(strVerificationMessage)
	bverifyConfirmationPopup_Content_CL=true
	var=CappingLimit_Page.popupValidationMessage.getroproperty("innertext")
	If instr(1,var, strVerificationMessage) = 0 Then
		LogMessage "WARN","Verification",strVerificationMessage &": Failed" ,false	
		bverifyConfirmationPopup_Content_CL=false
		else
		LogMessage "RSLT","Verification",strVerificationMessage &": Passed" ,true
			bverifyConfirmationPopup_Content_CL=true
	End If
	'CappingLimit_Page.btnYes_Confirmation.click
'	  If Err.Number<>0 Then
'       bverifyConfirmationPopup_Content_CL=false
'            LogMessage "WARN","Verification",strVerificationMessage &": Failed" ,false
'       Exit Function
'   End If
	verifyConfirmationPopup_Content_CL=bverifyConfirmationPopup_Content_CL
End Function

'[Verify Submit Button is disabled on Capping Limit Screen]
Public Function VerifybtnSubmitDisable1_CL()
	bDevPending=False
   Dim bVerifybtnSubmit_CL:bVerifybtnSubmit_CL=true
	'CashlineCancellation.tblSelectedCardsHeader.Click
	intBtnSubmit=Instr(CappingLimit_Page.btnSubmit.Object.GetAttribute("disabled"),("disabled"))
	If  not intBtnSubmit=0 Then
		LogMessage "RSLT","Verification","Submit button is disable as per expectation.",True
		bVerifybtnSubmit_CL=true
	Else
		LogMessage "WARN","Verifiation","Submit button is enable. Expected to be disable.",false
		bVerifybtnSubmit_CL=false
	End If
	VerifybtnSubmitDisable1_CL=bVerifybtnSubmit_CL
End Function
