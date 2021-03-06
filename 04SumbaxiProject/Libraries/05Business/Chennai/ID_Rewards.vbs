'[Verify fields displayed below Rewards section]
Public Function VerifyRewardDetails_Rewards(arrLblValPairs,strProductType)
scrollPageDown 3
bverifyfieldvalues = VerifyIDLabelValuePairs(coRewards_Page.lblAccordionRewards,arrLblValPairs,strProductType,"Reward Details")
VerifyRewardDetails_Rewards = bverifyfieldvalues
End Function

'[Select Row from Rewards scheme table displayed]
Public Function SelectRow_Rewards(lstRowData)
SelectRow_Rewards = SelectTableRow(coRewards_Page.tblRewardsSchemeHeader,coRewards_Page.tblRewardsSchemeContent,lstRowData,"Rewards Scheme","SCHEME ID",False,False)
End Function

'[Verify default value displayed in Reward Balances dropdown]
Public Function VerifyDefaultBalances_Rewards(strBalance)
bVerifyFieldBalance = True	
If Not IsNull(strBalance) Then
   scrollPageDown 1
   If Not verifyFieldValue(coRewards_Page.txtBalances(),strBalance,"Balances") Then
    bVerifyFieldBalance = False
   End If
End If
VerifyDefaultBalances_Rewards = bVerifyFieldBalance
End Function

'[Verify list of values in Balances dropdown displayed under Rewards section]
Public Function VerifyBalancesDropdown_Rewards(lstBalances)
bVerifyValues = True
If Not IsNull(lstBalances) Then
scrollPageDown 1
bVerifyValues = verifyComboboxItems(coRewards_Page.lstBalances(),lstBalances,"Balances")		
End If
VerifyBalancesDropdown_Rewards = bVerifyValues
End Function

'[Select Reward Balances from dropdown displayed]
Public Function SelectBalancesCombobox_Rewards(strItem)
scrollPageDown 1
SelectBalancesCombobox_Rewards = SelectComboBoxItem(coRewards_Page.txtBalances,strItem,"Balances")
End Function

'[Verify Reward balances displayed in the table]
Public Function VerifyBalances_Rewards(lstRowData)
scrollPageDown 1
VerifyBalances_Rewards = VerifyTableSingleRowData(coRewards_Page.tblPointsBalHeader,coRewards_Page.tblPointsBalContent,lstRowData,"Reward Balances")
End Function
