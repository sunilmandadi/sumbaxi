'[Verify list of values displayed in Reason for Blocking dropdown]
Public Function VerifylstAReasonBlocking_BC(lstBlockReason)
bVerifyValues = False
If Not IsNull(lstBlockReason) Then
	bVerifyValues = verifyComboboxItems(coBlockCard_Page.lstBlockReason,lstBlockReason,"Reason for Blocking")		
End If
VerifylstAReasonBlocking_BC = bVerifyValues	
End Function

'[Set Reason For Blocking dropdown as]
Public Function SetReasonforBlocking_BC(strBlockReason)
wait(2)
coBlockCard_Page.txtBlockReason().set strBlockReason
If Err.Number <> 0 Then 
  LogMessage "WARN","Verification","Failed to Set Reason for Blocking in text box", False
  SetReasonforBlocking_BC = False
  Exit Function
Else
  SetReasonforBlocking_BC = True
End If
End Function

'[Verify table row Selected Card displayed for Single Card]
Public Function VerifytblSelectedCardSingle_BC(lstSelectedCard)
	wait(1)
	bverifySelectedCards = True 
	If NOT isNull(lstSelectedCard) Then
		bverifySelectedCards = VerifyTableSingleRowData(coCancelCard_Page.tblSelectedCardHeader,coCancelCard_Page.tblSelectedCardBody,lstSelectedCard,"Validation Failed Cards")
	   If coCommon_Page.btnOK.Exist Then 
	      coCommon_Page.btnOK.Click
	   End IF
	End IF
	VerifytblSelectedCardSingle_BC = bverifySelectedCards
End Function

'[Verify table row Selected Card displayed for Multiple Cards]
Public Function VerifytblSelectedCardMultiple_BC(lstSelectedCard,lstSelectedCard1)
	wait(1)
	bverifySelectedCards = True 
	If NOT isNull(lstSelectedCard) Then
		bverifySelectedCards = VerifyTableSingleRowData(coCancelCard_Page.tblSelectedCardHeader,coCancelCard_Page.tblSelectedCardBody,lstSelectedCard,"Validation Failed Cards")
	End IF
	If NOT isNull(lstSelectedCard1) AND bverifySelectedCards Then
		bverifySelectedCards = VerifyTableSingleRowData(coCancelCard_Page.tblSelectedCardHeader,coCancelCard_Page.tblSelectedCardBody,lstSelectedCard1,"Validation Failed Cards")
	End If	
	If coCommon_Page.btnOK.Exist Then 
	   coCommon_Page.btnOK.Click
	End IF
	VerifytblSelectedCardMultiple_BC = bverifySelectedCards
End Function

'[Click On Select All CheckBox displayed in Block All cards STP Page]
Public Function SelectAllcheckbox_BlockAllCards()
	coBlockCard_Page.tblCheckBox.Click 
	WaitForIServeLoading
	If Err.Number <> 0 Then
	  SelectAllcheckbox_BlockAllCards = False
	  LogMessage "WARN","Verification","Failed to tick checkbox in table header", False
	  Exit Function
	Else
	  SelectAllcheckbox_BlockAllCards = True
	End If
End Function

'[Verify Validation Failed table displayed in Block All cards STP Page]
Public Function VerifyFailedMessage_BlockAllCards(lstVaidtionFailedMsg,lstVaidtionFailedMsg1)
	wait(1)
	bverifyFailedMsg = True 
	If NOT isNull(lstVaidtionFailedMsg) Then
		bverifyFailedMsg = VerifyTableSingleRowData(coCancelCard_Page.tblValidationFailedHeader,coCancelCard_Page.tblValidationFailedContent,lstVaidtionFailedMsg,"Validation Failed Cards")
	End IF
	If NOT isNull(lstVaidtionFailedMsg1) AND bverifyFailedMsg Then
		bverifyFailedMsg = VerifyTableSingleRowData(coCancelCard_Page.tblValidationFailedHeader,coCancelCard_Page.tblValidationFailedContent,lstVaidtionFailedMsg1,"Validation Failed Cards")
	End If	
	If coCommon_Page.btnOK.Exist Then 
	   coCommon_Page.btnOK.Click
	End IF
	VerifyFailedMessage_BlockAllCards = bverifyFailedMsg
End Function
