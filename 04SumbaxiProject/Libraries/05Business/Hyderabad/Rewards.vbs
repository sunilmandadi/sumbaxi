'[Select radiobutton DBS points or rebates]
Public Function selectRadioButton_rewards(strRewards)
	bDevPending=true
	bselectRadioButton_rewards=true
	intRadio_Rewards=Instr(Rewards.rbgShow.GetROproperty("class"),"disabled-area")
	If intRadio_Rewards = 0 Then
		bselectRadioButton_rewards=SelectRadioButtonGrp(strRewards, Rewards.rbgShow, Array("DBS Points","Rebates"))
	Else
		LogMessage "RSLT","Verifiation","Radio button is disabled by default.",True
	End If
	If Err.Number<>0 Then
       bselectRadioButton_rewards=false
          LogMessage "WARN","Verification","Failed to select radiobutton or radiobutton is disabled" ,True
       Exit Function
   End If
   selectRadioButton_rewards=bselectRadioButton_rewards
End Function

'[Verify the row data for Relationship summary table in rewards page]
Public Function verifyrowdata_RelationshipSummRewards(arrRowDataList)
	bverifyrowdata_RelationshipSummRewards = true
	verifyrowdata_RelationshipSummRewards = verifyTableContentList(Rewards.tblRelationshipSummaryHeader,Rewards.tblRelationshipSummaryContent,arrRowDataList,"Relationship Summary",false,null,null,null)
	verifyrowdata_RelationshipSummRewards = bverifyrowdata_RelationshipSummRewards
End Function

'[Verify the row data for Account Details table in rewards page]
Public Function verifyrowdata_AccountDetailsRewards(arrRowDataList)
	bverifyrowdata_AccountDetailsRewards = true
	verifyrowdata_AccountDetailsRewards = verifyTableContentList(Rewards.tblAccountDetailsHeader,Rewards.tblAccountDetailsContent,arrRowDataList,"Account Details",true,Rewards.lnkNext,Rewards.lnkNext1,Rewards.lnkPrevious)
	verifyrowdata_AccountDetailsRewards = bverifyrowdata_AccountDetailsRewards
End Function

'[Click on Points that Expire link from Account Details table]
Public Function clickPointsExpire_Rewards(lstRowData)
	bclickPointsExpire_Rewards = true
	clickPointsExpire_Rewards = selectTableLink(Rewards.tblAccountDetailsHeader,Rewards.tblAccountDetailsContent,lstRowData,"Account Details","Points That Expire",true,Rewards.lnkNext,Rewards.lnkNext1,Rewards.lnkPrevious)
	clickPointsExpire_Rewards = bclickPointsExpire_Rewards
End Function

'[Verify the pop up Points Expiry Details]
Public Function verifyPointsExpiryDetails(bExist)
	bDevPending=false
   bActualExist=Rewards.popupPointsExpiryDetails.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Points Expiry Details Exists As Expected" ,true
       verifyPointsExpiryDetails=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Points Expiry Details does not Exists As Expected" ,true
       verifyPointsExpiryDetails=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :Points Expiry Details does not Exists As Expected" ,False
       verifyPointsExpiryDetails=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :Points Expiry Details Still Exists" ,False
       verifyPointsExpiryDetails=False
   End If
End Function

'[Verify the row data for Points Expiry Details table]
Public Function verifyrowdata_PointsExpDetailsRewards(arrRowDataList)
	bverifyrowdata_PointsExpDetailsRewards = true
	verifyrowdata_PointsExpDetailsRewards = verifyTableContentList(Rewards.tblPointsExpiryHeader,Rewards.tblPointsExpiryContent,arrRowDataList,"Points Expiry Details",false,null,null,null)
	Rewards.btnOK.Click
	verifyrowdata_PointsExpDetailsRewards = bverifyrowdata_PointsExpDetailsRewards	
End Function

'[Verify the click for redemption history link]
Public Function ClickViewRedemptionHist_Rewards()
	bClickViewRedemptionHist_Rewards=true
	Rewards.lnkRedemptionHistory.click
	WaitForICallLoading
	If not Rewards.popupRedemptionHistory.Exist Then
		bClickViewRedemptionHist_Rewards=false
	End If
	ClickViewRedemptionHist_Rewards=bClickViewRedemptionHist_Rewards
End Function

'[Verify the pop up View Redemption History]
Public Function verifyViewRedemptionHistory(bExist)
	bDevPending=false
   bActualExist=Rewards.popupRedemptionHistory.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Redemption History Details Exists As Expected" ,true
       verifyViewRedemptionHistory=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Redemption History Details does not Exists As Expected" ,true
       verifyViewRedemptionHistory=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :Redemption History Details does not Exists As Expected" ,False
       verifyViewRedemptionHistory=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :Redemption History Details Still Exists" ,False
       verifyViewRedemptionHistory=False
   End If
End Function

'[Verify the row data for Redemption History table]
Public Function verifyrowdata_RedemptionHistoryRewards(arrRowDataList)
	bverifyrowdata_RedemptionHistoryRewards = true
	verifyrowdata_RedemptionHistoryRewards = verifyTableContentList(Rewards.tblRedemptionHistoryDetailsHeader,Rewards.tblRedemptionHistoryDetailsContent,arrRowDataList,"Redemption History Details",false,null,null,null)
	Rewards.btnOK.Click
	verifyrowdata_RedemptionHistoryRewards = bverifyrowdata_RedemptionHistoryRewards	
End Function

'[Verify the row data for Summary table in Rebates page]
Public Function verifyrowdata_SummaryRebates(arrRowDataList)
	bverifyrowdata_SummaryRebates = true
	verifyrowdata_SummaryRebates = verifyTableContentList(Rewards.tblRebateSummaryHeader,Rewards.tblRebateSummaryContent,arrRowDataList,"Summary",false,null,null,null)
	verifyrowdata_SummaryRebates = bverifyrowdata_SummaryRebates
End Function

'[Verify the row data for rebates breakdown in Rebates page]
Public Function verifyrowdata_RebatesBreakDown(arrRowDataList)
	bverifyrowdata_RebatesBreakDown = true
	verifyrowdata_RebatesBreakDown = verifyTableContentList(Rewards.tblRebateBreakdownHeader,Rewards.tblRebateBreakdownContent,arrRowDataList,"Rebates Breakdown",false,null,null,null)
	verifyrowdata_RebatesBreakDown = bverifyrowdata_RebatesBreakDown
End Function

'[Verify the click for link cash rebates transaction history]
Public Function ClickCashRebatesTrnHistory_Rewards()
	bClickCashRebatesTrnHistory_Rewards=true
	Rewards.lnkCashRebateHistory.click
	WaitForICallLoading
	If not Rewards.popupCashRebatesTransactionHistory.Exist Then
		bClickCashRebatesTrnHistory_Rewards=false
	End If
	ClickCashRebatesTrnHistory_Rewards=bClickCashRebatesTrnHistory_Rewards
End Function

'[Verify the row data for Cash Rebates Transaction table]
Public Function verifyrowdata_CashRebatesTransaction(arrRowDataList)
	bverifyrowdata_CashRebatesTransaction = true
	verifyrowdata_CashRebatesTransaction = verifyTableContentList(Rewards.tblRebateSummaryHeader,Rewards.tblRebateSummaryContent,arrRowDataList,"Cash Rebates Transaction History",false,null,null,null)
	Rewards.btnOK.Click
	verifyrowdata_CashRebatesTransaction = bverifyrowdata_CashRebatesTransaction	
End Function
