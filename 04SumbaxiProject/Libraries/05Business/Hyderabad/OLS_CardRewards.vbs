'[Verify and Click Cards Reward Link from leftMenu]
Public Function ClickLink_CardRewards()
bClickLink_CardRewards=true
	bcAccountOverview_LeftMenu.btnCardRewards.Click
	WaitForIcallLoading
	If Err.Number<>0 Then
       ClickLink_CardRewards=false
       LogMessage "WARN","Verification","Failed to Click Link  : Card Rewards" ,false
       Exit Function
	End If
	Wait 1
	waitForIcallLoading	
ClickLink_CardRewards = bClickLink_CardRewards
End Function 

'[Verify row Data for Rewards Summary Table]
Public Function verifyRewardsSummary_RowData(arrRowDataList)
   verifyRewardsSummary_RowData=verifyTableContentList(CardRewards.tblRewardsSummaryHeader,CardRewards.tblRewardsSummaryContent,arrRowDataList,"Rewards Summary",false,null,null,null)
End Function

'[Click on the View Campaign Details link]
Public Function clicklink_ViewCampaignDetails()
   bDevPending=true
   CardRewards.lnkViewCampaignDetails.click
   Wait(10)
   waitForIcallLoading
   If Err.Number<>0 Then
       clicklink_ViewCampaignDetails=false
            LogMessage "WARN","Verification","Failed to Click Link : View Campaign Details" ,false
       Exit Function
   End If
   clicklink_ViewCampaignDetails=true
End Function

'[Verify the pop up Campaign Summary]
Public Function verifyPopupCampaignSummaryexist(bExist)
   bDevPending=false
   bActualExist=CardRewards.popupCampaignSummary.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Campaign Summary Exists As Expected" ,true
       verifyPopupCampaignSummaryexist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Campaign Summary does not Exists As Expected" ,true
       verifyPopupCampaignSummaryexist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :Campaign Summary does not Exists As Expected" ,False
       verifyPopupCampaignSummaryexist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :Campaign Summary Still Exists" ,False
       verifyPopupCampaignSummaryexist=False
   End If
End Function

'[Verify Table Campaign Summary has following Columns]
Public Function verifyCampaignSummaryTableColumns(arrColumnNameList)
	verifyCampaignSummaryTableColumns = verifyTableColumns(CardRewards.tblCampaignSummaryHeader,arrColumnNameList)
End Function

'[Verify row Data for Campaign Summary Table]
Public Function verifyCampaignSummary_RowData(arrRowDataList)
   verifyCampaignSummary_RowData=verifyTableContentList(CardRewards.tblCampaignSummaryHeader,CardRewards.tblCampaignSummaryContent,arrRowDataList,"Campaign Summary",True,CardRewards.lnkNext,CardRewards.lnkNext,CardRewards.lnkPrevious)
End Function

'[Click Button Ok on pop up Campaign Summary]
Public Function clickbtnOk_CampaignSummary()
   bclickbtnOk_CampaignSummary=true
   CardRewards.btnOK_CampSummary.click
   If Err.Number<>0 Then
       clickbtnOk_CampaignSummary=false
            LogMessage "WARN","Verification","Failed to Click Button :Ok" ,false
       Exit Function
   End If
   clickbtnOk_CampaignSummary=bclickbtnOk_CampaignSummary
End Function

'[Verify Table Membership Details has following Columns]
Public Function verifyMembershipDetailsTableColumns(arrColumnNameList)
	verifyMembershipDetailsTableColumns=verifyTableColumns(CardRewards.tblMembershipDetailsHeader,arrColumnNameList)
End Function

'[Verify row Data for Membership Details Table]
Public Function verifyMembershipDetails_RowData(arrRowDataList)
   verifyMembershipDetails_RowData = verifyTableContentList(CardRewards.tblMembershipDetailsHeader,CardRewards.tblMembershipDetailsContent,arrRowDataList,"Membership Details" ,  false,null ,null,null)
End Function

'[Click on View link in Membership Details table]
Public Function clickFeeDetails(arrowDataList)
   	clickFeeDetails = selectTableLink(CardRewards.tblMembershipDetailsHeader,CardRewards.tblMembershipDetailsContent,arrowDataList,"Membership Details","Fee Details",false,null,null,null)
End Function

'[Verify the pop up Frequent Flyer Details]
Public Function verifyPopupFrequentFlyerDetailsexist(bExist)
   bDevPending=false
   bActualExist=CardRewards.popupFrequentFlyerDetails.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Frequent Flyer Exists As Expected" ,true
       verifyPopupFrequentFlyerDetailsexist=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Frequent Flyer does not Exists As Expected" ,true
       verifyPopupFrequentFlyerDetailsexist=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :Frequent Flyer does not Exists As Expected" ,False
       verifyPopupFrequentFlyerDetailsexist=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :Frequent Flyer Still Exists" ,False
       verifyPopupFrequentFlyerDetailsexist=False
   End If
End Function

'[Verify row Data for Frequent Flyer details Table]
Public Function verifyFrequentFlyerDetails_RowData(arrRowDataList)
   verifyFrequentFlyerDetails_RowData=verifyTableContentList(CardRewards.tblFreqFlyerDetailsHeader,CardRewards.tblFreqFlyerDetailsContent,arrRowDataList,"Frequent Flyer Details" ,  false,null ,null,null)
End Function

'[Click Button Ok on pop up Frequent Flyer]
Public Function clickbtnOk_FrequentFlyer()
   bclickbtnOk_FrequentFlyer=true
   CardRewards.btnOK_FreqFlyer.click
   If Err.Number<>0 Then
       clickbtnOk_FrequentFlyer=false
            LogMessage "WARN","Verification","Failed to Click Button :Ok" ,false
       Exit Function
   End If
   clickbtnOk_FrequentFlyer=bclickbtnOk_FrequentFlyer
End Function

'[Click on Active Rebate link in Rewards Summary table]
Public Function clickActiveRebates(arrowDataList)
   	clickActiveRebates=selectTableLink(CardRewards.tblRewardsSummaryHeader,CardRewards.tblRewardsSummaryContent,arrowDataList,"Rewards Summary","Active $",false,null,null,null)
End Function

'[Verify the pop up Active Rebates in Rewards Summary Table]
Public Function verifyPopupActiveRebates(bExist)
   bDevPending=false
   bActualExist=CardRewards.popupCashRebateExpiryDetails.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Active Rebates Exists As Expected" ,true
       verifyPopupActiveRebates=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Active Rebates does not Exists As Expected" ,true
       verifyPopupActiveRebates=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :Active Rebates does not Exists As Expected" ,False
       verifyPopupActiveRebates=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :Active Rebates Still Exists" ,False
       verifyPopupActiveRebates=False
   End If
End Function

'[Click Button Ok on pop up Active Rebates]
Public Function clickbtnOk_ActiveRebates()
   bclickbtnOk_ActiveRebates=true
   CardRewards.btnOK_Rebates.click
   If Err.Number<>0 Then
       clickbtnOk_ActiveRebates=false
            LogMessage "WARN","Verification","Failed to Click Button :Ok" ,false
       Exit Function
   End If
   clickbtnOk_ActiveRebates=bclickbtnOk_ActiveRebates
End Function

'[Click on Points that expire link in Rewards Summary table]
Public Function clickPointsExpire(arrowDataList)
   	clickPointsExpire = selectTableLink(CardRewards.tblRewardsSummaryHeader,CardRewards.tblRewardsSummaryContent,arrowDataList,"Rewards Summary","Points That Will Expire",false,null,null,null)
End Function

'[Verify the pop up Point Expire in Rewards Summary Table]
Public Function verifyPopupPointsExpire(bExist)
   bDevPending=false
   bActualExist=CardRewards.popupPointsExpiryDetails.Exist(2)
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Points Expire Exists As Expected" ,true
       verifyPopupPointsExpire=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :Points Expire does not Exists As Expected" ,true
       verifyPopupPointsExpire=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :Points Expire does not Exists As Expected" ,False
       verifyPopupPointsExpire=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :Points Expire Still Exists" ,False
       verifyPopupPointsExpire=False
   End If
End Function

'[Click Button Ok on pop up Points Expiry Details]
Public Function clickbtnOK_PointsExpire()
   CardRewards.btnOK_PointsExpire.click
   If Err.Number<>0 Then
       clickbtnOK_PointsExpire=false
            LogMessage "WARN","Verification","Failed to Click Button:Ok" ,false
       Exit Function
   End If
   clickbtnOK_PointsExpire=true
End Function

'[Click the Credit Points link from Action of rewards summary table]
Public Function selectCreditPoints_Rewards(lstRewardsSummary)
   bselectCreditPoints_Rewards=true
	selectCreditPoints_Rewards= selectTableSubMenu(CardRewards.tblRewardsSummaryHeader,CardRewards.tblRewardsSummaryContent,lstRewardsSummary,"Rewards Summary","NULL",False,NULL,NULL,NULL,"Credit Points",bDisabled)
	If bDisabled Then
		LogMessage "RSLT", "Verification","Redemption action menu is not enabled",false
		bselectCreditPoints_Rewards=false
	End If
	WaitForICallLoading
	Wait 1
    selectCreditPoints_Rewards=bselectCreditPoints_Rewards
End Function
