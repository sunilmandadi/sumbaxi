Dim BalancesAndLimits
Set BalancesAndLimits = cBalancesAndLimits()

'This is the Screen BalancesAndLimits

Public Function cBalancesAndLimits()
    Set cBalancesAndLimits = New clsBalancesAndLimits
End Function

Class clsBalancesAndLimits

        Private Sub Class_Initialize()
        End Sub

        Private Sub Class_Terminate()
        End Sub

        '******************************** Object Initialization ******************************************************************


        Public Function pageExists()

           If  (lblCurrentBalance().exist) Then
               pageExists = true
            else
              pageExists = false
           End If

        End Function


        Public Function lblAccount()
           Set lblAccounts =  Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblAccount")
        End Function
		Public Function lblRelationship()
           Set lblRelationship =  Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship")
        End Function

        Public Function lblCurrentBalance()
           Set lblCurrentBalance = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_CurrentBalance")
        End Function
        
		Public Function lblBalanceLimits_PendingPayments()
			Set lblBalanceLimits_PendingPayments = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_PendingPayments")
		End Function
		
		Public Function lblRelationship_PendingPayments()
			Set lblRelationship_PendingPayments = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_lblPendingPayments")
		End Function

        Public Function lblPendingDebits()
           Set lblPendingDebits = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_PendingDebits")
        End Function

        Public Function lblPendingCredits()
           Set lblPendingCredits = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_PendingCredit")
        End Function

        Public Function lblOutstandingBalance()
           Set lblOutstandingBalance = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_OutstandingBalance")
        End Function

        Public Function lblTotalCreditLimit()
           Set lblTotalCreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_TotalCreditLimit")
        End Function

        Public Function lblAvailableLimit()
           Set lblAvailableLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_TotalAvailableLimit")
        End Function

        Public Function lblCashAdvance_Current()
           Set lblCashAdvance_Current = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCashAdvance_Current")
        End Function

        Public Function lblCashAdvance_Outstanding()
           Set lblCashAdvance_Outstanding = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCashAdvance_Outstanding")
        End Function

        Public Function lblCashAdvance_CreditLimit()
           Set lblCashAdvance_CreditLimit =Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCashAdvance_CreditLimit")
        End Function

        Public Function lblCashAdvance_AvailableLimit()
           Set lblCashAdvance_AvailableLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCashAdvance_AvailableLimit")
        End Function

        Public Function lblRetail_Current()
           Set lblRetail_Current = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_Current")
        End Function

        Public Function lblRetail_Outstanding()
           Set lblRetail_Outstanding = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_Outstanding")
        End Function

        Public Function lblRetail_CreditLimit()
           Set lblRetail_CreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_CreditLimit")
        End Function

        Public Function lblRetail_AvailableLimit()
           Set lblRetail_AvailableLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_AvailableLimit")
        End Function

        Public Function lblRetail_TempCreditLimit()
           Set lblRetail_TempCreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_TempCreditLimit")
        End Function

        Public Function lblRetail_EffectiveDate()
           Set lblRetail_EffectiveDate = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_EffectiveDate")
        End Function

        Public Function lblRetail_ExpiryDate()
           Set lblRetail_ExpiryDate = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_ExpiryDate")
        End Function

        Public Function lblRetail_ChangeReason()
           Set lblRetail_ChangeReason = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRetail_ChangeReason")
        End Function

        Public Function lblCardLimit_WithdrawalPerDay()
           Set lblCardLimit_WithdrawalPerDay = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_CashWithdrawalLimitPerDay")
        End Function

        Public Function lblCardLimit_WithdrawalPerTransaction()
           Set lblCardLimit_WithdrawalPerTransaction = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_CashWithdrawalLimitPerTransaction")
        End Function

        Public Function lblCardLimit_EligibleTransactions()
           Set lblCardLimit_EligibleTransactions = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_EligibleTransactions")
        End Function

        Public Function lblCardLimit_CreditLimit()
           Set lblCardLimit_CreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_CreditLimit")
        End Function

        Public Function lblCardLimit_AvailableLimit()
           Set lblCardLimit_AvailableLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_AvailableLimit")
        End Function

        Public Function lblCardLimit_TempCreditLimit()
           Set lblCardLimit_TempCreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_tempCreditLimit")
        End Function

        Public Function lblCardLimit_EffectiveDate()
           Set lblCardLimit_EffectiveDate = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_EffectiveDate")
        End Function

        Public Function lblCardLimit_ExpiryDate()
           Set lblCardLimit_ExpiryDate = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_ExpiryDate")
        End Function

        Public Function lblRelationship_CurrentBalance()
           Set lblRelationship_CurrentBalance = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_CurrentBalance")
        End Function

        Public Function lblRelationship_PendingDebits()
           Set lblRelationship_PendingDebits = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_PendingDebit")
        End Function

        Public Function lblRelationship_PendingCredits()
           Set lblRelationship_PendingCredits = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_PendingCredits")
        End Function

        Public Function lblRelationship_OutstandingBalance()
           Set lblRelationship_OutstandingBalance = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_OutstandingBalance")
        End Function

        Public Function lblRelationship_CreditLimit()
           Set lblRelationship_CreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_CreditLimit")
        End Function

        Public Function lblRelationship_TempCreditLimit()
           Set lblRelationship_TempCreditLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_tempCreditLimit")
        End Function

        Public Function lblRelationship_AvailableCreditLimit()
           Set lblRelationship_AvailableCreditLimit =  Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_AvailableLimit")
        End Function

        Public Function lblRelationship_EffectiveDate()
           Set lblRelationship_EffectiveDate = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_EffectiveDate")
        End Function

        Public Function lblRelationship_ExpiryDate()
           Set lblRelationship_ExpiryDate = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblRelationship_ExpiryDate")
        End Function
        
        Public Function lblAccountBalance_AvailableBalance()
           Set lblAccountBalance_AvailableBalance = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblAccountBalance_AvailableBalance")
        End Function
        
        Public Function lblAccountBalance_LedgerBalance()
           Set lblAccountBalance_LedgerBalance = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblAccountBalance_LedgerBalance")
        End Function
        
        Public Function lblAccountBalance_EarmarkAmount()
           Set lblAccountBalance_EarmarkAmount = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblAccountBalance_EarmarkAmount")
        End Function
        
        Public Function lblAccountBalance_Signals()
           Set lblAccountBalance_Signals = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblAccountBalance_Signals")
        End Function
        
        Public Function lblHoldBalance_HalfDay()
           Set lblHoldBalance_HalfDay = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblHoldBalance_HalfDay")
        End Function
        
        Public Function lblHoldBalance_OneDay()
           Set lblHoldBalance_OneDay = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblHoldBalance_OneDay")
        End Function
        
        Public Function lblHoldBalance_TwoDays()
           Set lblHoldBalance_TwoDays = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblHoldBalance_TwoDays")
        End Function
        
        Public Function lblHoldBalance_LessThanTwoDays()
           Set lblHoldBalance_LessThanTwoDays = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblHoldBalance_LessThanTwoDays")
        End Function
        
        Public Function lblReturnedCheque_CurrentMonth()
           Set lblReturnedCheque_CurrentMonth = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblReturnedCheque_CurrentMonth")
        End Function
        
        Public Function lblReturnedCheque_LastMonth()
           Set lblReturnedCheque_LastMonth = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblReturnedCheque_LastMonth")
        End Function
        
        Public Function lblReturnedCheque_Last2Months()
           Set lblReturnedCheque_Last2Months = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblReturnedCheque_Last2Months")
        End Function
        
        Public Function lblLimits_OverdraftLimit()
           Set lblLimits_OverdraftLimit = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblLimits_OverdraftLimit")
        End Function
        
        Public Function lblLimits_AccruedOverdraft()
           Set lblLimits_AccruedOverdraft = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblLimits_AccruedOverdraft")
        End Function
        
        Public Function lblLimits_ActionIcon()
           Set lblLimits_ActionIcon = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblLimits_ActionIcon")
        End Function
        
        Public Function popupSignalDetails()
           Set popupSignalDetails = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails")
        End Function
        
        Public Function btnOK_popupEarMarkDetails()
           Set btnOK_popupEarMarkDetails = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails").WebButton("btnOK_popupSignalDetails")
        End Function
        
        Public Function btnOK_popupSignalDetails()
        	Set btnOK_popupSignalDetails = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebButton("btnOK_popupSignalDetails")
        End Function
        
        Public Function lnkNextEarMark()
           Set lnkNextEarMark = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails").WebElement("lnkNext")
        End Function
        
        Public Function lnkNext1EarMark()
           Set lnkNext1EarMark = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails").WebElement("lnkNext1")
        End Function
        
        Public Function lnkPreviousEarMark()
        	Set lnkPreviousEarMark = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails").WebElement("lnkPrevious")
        End Function
        
        Public Function lnkPreviousSignal()
           Set lnkPreviousSignal = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebElement("lnkPrevious")
        End Function
        
        Public Function lnkNextSignal()
        	Set lnkNextSignal = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebElement("lnkNext")
        End Function
        
        Public Function lnkNext1Signal()
        	Set lnkNext1Signal = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebElement("lnkNext1")
        End Function
        
        Public Function rdg_LatestHistory()
           Set rdg_LatestHistory = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebRadioGroup("rdg_LatestHistory")
        End Function
        
        Public Function tblProductsListContent()
           Set tblProductsListContent = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebElement("tblProductsListContent")
        End Function
        
        Public Function tblProductsListHeader()
           Set tblProductsListHeader = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebElement("tblProductsListHeader")
        End Function
        
        Public Function popupEarmarkDetails()
           Set popupEarmarkDetails = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails")
        End Function
        
        Public Function tblEarmarkHeader()
           Set tblEarmarkHeader = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails").WebElement("tblEarmarkHeader")
        End Function
        
        Public Function tblEarmarkContent()
           Set tblEarmarkContent = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupEarmarkDetails").WebElement("tblEarmarkContent")
        End Function
        
        Public Function lblBalanceLimits_CurrentBalance_Label()
           Set lblBalanceLimits_CurrentBalance_Label = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblBalanceLimits_CurrentBalance_Label")
        End Function
        
        Public Function lblEarmarkLink()
        	Set lblEarmarkLink = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblEarmarkLink")
        End Function
        
        Public Function lblSignalsLink()
        	Set lblSignalsLink = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblSignalsLink")
        End Function
        
        Public Function rbHistory()
        	Set rbHistory = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("popupSignalDetails").WebElement("rbHistory")
        End Function
        
        
        '**************1602 Changes*******************************
        
        Public Function lblCardLimits_RTLPerDay()
           Set lblCardLimits_RTLPerDay = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_RTLPerDay")
        End Function
        
        Public Function lblCardLimits_RTLPerMonth()
           Set lblCardLimits_RTLPerMonth = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_RTLPerMonth")
        End Function
        
        Public Function lblCardLimits_RTLPerYear()
           Set lblCardLimits_RTLPerYear = Browser("Browser_iCall_Home").Page("BalancesAndLimits").WebElement("lblCardLimits_RTLPerYear")
        End Function
        

        '******************************** End of Object Initialization ******************************************************************

        '*****************************Buttons & Link Clicks on the Page **********************************************************
        Public Sub clickBalansesandLimits()
            lnkBalansesandLimits().Click
        End Sub

        '*****************************End of Buttons & Link Clicks on the Page **********************************************************

        '*****************************Function on the Screen **********************************************************

        Public Function verifyBalanceAndLimits( strCurrentBalance, strPendingDebits, strPendingCredits, strOutstandingBalance, strTotalCreditLimit, strAvailableLimit,  _
		    strCashAdvance_Current, strCashAdvance_Outstanding, strCashAdvance_CreditLimit, strCashAdvance_AvailableLimit,  _
		    strRetail_Current, strRetail_Outstanding, strRetail_CreditLimit, strRetail_AvailableLimit, strRetail_TempCreditLimit, strRetail_EffectiveDate, strRetail_ExpiryDate, strRetail_ChangeReason, _
		    strCardLimit_WithdrawalPerDay, strCardLimit_WithdrawalPerTransaction, strCardLimit_EligibleTransactions, strCardLimit_CreditLimit, strCardLimit_AvailableLimit, strCardLimit_TempCreditLimit, strCardLimit_EffectiveDate, strCardLimit_ExpiryDate, strCardLimits_RTLPerDay, strCardLimits_RTLPerMonth, strCardLimits_RTLPerYear, _
		     strRelationship_CurrentBalance, strRelationship_PendingDebits, strRelationship_PendingCredits, strRelationship_OutstandingBalance, strRelationship_CreditLimit, strRelationship_TempCreditLimit, strRelationship_AvailableCreditLimit, strRelationship_EffectiveDate, strRelationship_ExpiryDate)
					bVerifyBalanceAndLimits=true
				bcAccountOverview_LeftMenu.clickBalanceLimits()
				WaitForICallLoading
				If Not pageExists() Then
					LogMessage "WARN","Verification","Statement Details page does not displayed",false
					bVerifyBalanceAndLimits=false
				Else
					LogMessage "RSLT","Verification","Statement Details page displayed Successfully",true
				End If
               If Not IsNull(strCurrentBalance) Then
					If Not verifyInnerText(lblCurrentBalance() , strCurrentBalance, "Current Balance")Then
								bVerifyBalanceAndLimits = False
						End If
                End If

                If Not IsNull(strPendingDebits) Then
					If Not verifyInnerText(lblPendingDebits() , strPendingDebits, "Pending Debits")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strPendingCredits) Then
					If Not verifyInnerText( lblPendingCredits(), strPendingCredits, "Pending Credit")Then
								bVerifyBalanceAndLimits = False
					End If		              
                End If

                If Not IsNull(strOutstandingBalance) Then
					If Not verifyInnerText(  lblOutstandingBalance(), strOutstandingBalance, "Outstanding Balance")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strTotalCreditLimit) Then
					If Not verifyInnerText(  lblTotalCreditLimit(), strTotalCreditLimit, "Total Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strAvailableLimit) Then
					If Not verifyInnerText(  lblAvailableLimit(), strAvailableLimit, "Available Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCashAdvance_Current) Then
					If Not verifyInnerText( lblCashAdvance_Current(), strCashAdvance_Current, "Cash Advance Current")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCashAdvance_Outstanding) Then
                  If Not verifyInnerText( lblCashAdvance_Outstanding(), strCashAdvance_Outstanding, "Cash Advance Outstanding")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCashAdvance_CreditLimit) Then
					If Not verifyInnerText(lblCashAdvance_CreditLimit(), strCashAdvance_CreditLimit, "Cash Advance Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCashAdvance_AvailableLimit) Then
                  If Not verifyInnerText(lblCashAdvance_AvailableLimit(), strCashAdvance_AvailableLimit, "Cash Advance Available Limit")Then
								bVerifyBalanceAndLimits = False
					End If
					
                End If

                If Not IsNull(strRetail_Current) Then
					If Not verifyInnerText( lblRetail_Current(), strRetail_Current, "Retail Current")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_Outstanding) Then
                  If Not verifyInnerText(lblRetail_Outstanding(), strRetail_Outstanding, "Retail Outstanding")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_CreditLimit) Then
					 If Not verifyInnerText(lblRetail_CreditLimit(), strRetail_CreditLimit, "Retail Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_AvailableLimit) Then
					If Not verifyInnerText(lblRetail_AvailableLimit(), strRetail_AvailableLimit,"Retail Available Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_TempCreditLimit) Then
                  If Not verifyInnerText(lblRetail_TempCreditLimit(), strRetail_TempCreditLimit,"Retail Temporary Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_EffectiveDate) Then
                  If Not verifyInnerText(lblRetail_EffectiveDate(), strRetail_EffectiveDate,"Retail Effective Date")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_ExpiryDate) Then
                  If Not verifyInnerText( lblRetail_ExpiryDate(), strRetail_ExpiryDate,"Retail Expiry Date")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRetail_ChangeReason) Then
                  If Not verifyInnerText(lblRetail_ChangeReason(), strRetail_ChangeReason,"Retail Change Reason")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_WithdrawalPerDay) Then
                  If Not verifyInnerText(lblCardLimit_WithdrawalPerDay(), strCardLimit_WithdrawalPerDay,"Card Limit Withdrawal Limit Per Day")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_WithdrawalPerTransaction) Then
                  If Not verifyInnerText(lblCardLimit_WithdrawalPerTransaction(), strCardLimit_WithdrawalPerTransaction,"Card Limit Withdrawal Limit Per Transaction")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_EligibleTransactions) Then
                  If Not verifyInnerText(lblCardLimit_EligibleTransactions(), strCardLimit_EligibleTransactions,"Card Limit Eligible Transaction")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_CreditLimit) Then
                   If Not verifyInnerText(lblCardLimit_CreditLimit(), strCardLimit_CreditLimit,"Card Limit Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_AvailableLimit) Then
					 If Not verifyInnerText(lblCardLimit_AvailableLimit(), strCardLimit_AvailableLimit,"Card Limit Available Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_TempCreditLimit) Then
                  If Not verifyInnerText(lblCardLimit_TempCreditLimit(), strCardLimit_TempCreditLimit,"Card Limit Temporary Credit  Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_EffectiveDate) Then
                  If Not verifyInnerText(lblCardLimit_EffectiveDate(), strCardLimit_EffectiveDate,"Card Limit Effective Date")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strCardLimit_ExpiryDate) Then
					If Not verifyInnerText(lblCardLimit_ExpiryDate(), strCardLimit_ExpiryDate,"Card Limit Expiry Date")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_CurrentBalance) Then
                   If Not verifyInnerText( lblRelationship_CurrentBalance(), strRelationship_CurrentBalance,"Relationship Current Balance")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_PendingDebits) Then
                   If Not verifyInnerText(lblRelationship_PendingDebits(), strRelationship_PendingDebits,"Relationship Pending Debits")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_PendingCredits) Then
                    If Not verifyInnerText(lblRelationship_PendingCredits(), strRelationship_PendingCredits,"Relationship Pending Credits")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

					lblRelationship_OutstandingBalance().click
                If Not IsNull(strRelationship_OutstandingBalance) Then
                   If Not verifyInnerText(lblRelationship_OutstandingBalance(), strRelationship_OutstandingBalance,"Relationship Outstanding Balance")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_CreditLimit) Then
                   If Not verifyInnerText(lblRelationship_CreditLimit(), strRelationship_CreditLimit,"Relationship Card Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_TempCreditLimit) Then
					Print (strRelationship_TempCreditLimit)
                   If Not verifyInnerText(lblRelationship_TempCreditLimit(), strRelationship_TempCreditLimit,"Relationship Temporary Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_AvailableCreditLimit) Then
                   If Not verifyInnerText(lblRelationship_AvailableCreditLimit(), strRelationship_AvailableCreditLimit,"Relationship Available Credit Limit")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_EffectiveDate) Then
                    If Not verifyInnerText(lblRelationship_EffectiveDate(), strRelationship_EffectiveDate,"Relationship Effective Date")Then
								bVerifyBalanceAndLimits = False
					End If
                End If

                If Not IsNull(strRelationship_ExpiryDate) Then
                   If Not verifyInnerText(lblRelationship_ExpiryDate(), strRelationship_ExpiryDate,"Relationship Expiry Date")Then
								bVerifyBalanceAndLimits = False
					End If
                End If
                
                '***************1602 changes*********************
				If Not IsNull(strCardLimits_RTLPerDay) Then
				        If Not verifyInnerText(lblCardLimits_RTLPerDay(), strCardLimits_RTLPerDay,"Retail Txn Limit: Per Day")Then
						bVerifyBalanceAndLimits = False
						End If
				End If
				
				If Not IsNull(strCardLimits_RTLPerMonth) Then
				        If Not verifyInnerText(lblCardLimits_RTLPerMonth(), strCardLimits_RTLPerMonth,"Retail Txn Limit: Per Month")Then
						bVerifyBalanceAndLimits = False
						End If
				End If
				
				If Not IsNull(strCardLimits_RTLPerYear) Then
				        If Not verifyInnerText(lblCardLimits_RTLPerYear(), strCardLimits_RTLPerYear,"Retail Txn Limit: Per Year")Then
						bVerifyBalanceAndLimits = False
						End If
				End If
                verifyBalanceAndLimits = bVerifyBalanceAndLimits

        End Function

        '*****************************End of Function on the Screen **********************************************************

End Class
