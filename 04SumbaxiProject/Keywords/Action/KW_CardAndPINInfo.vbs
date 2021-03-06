
'Keyword Generation
Class KW_CardAndPINInfo
    Public Function runKey (strKeyArgument)
		Dim bCardAndPINInfo

	   Select Case gstrAction

				   Case "VerifyCardAndPINInfo":
						 Dim cVerifyCardAndPINInfo
						Set cVerifyCardAndPINInfo = New KW_VerifyCardAndPINInfo
						 bCardAndPINInfo =	cVerifyCardAndPINInfo.runKey(strKeyArgument)

					Case Else  LogMessage "RSLT", "Verification", "Action " +gstrAction + " is not valid for the Keyword CardAndPINInfo ", False

	   End Select

       If  bCardAndPINInfo Then
            LogMessage "RSLT", "Verification", "Keyword CardAndPINInfo  with Action " + gstrAction + " executed successfully", True
			runKey =True
        else
            LogMessage "RSLT", "Verification", "Keyword CardAndPINInfo  with Action " + gstrAction + " failed during execution", False
            runKey =False
        End If

    End Function

End Class


'Keyword Generation
Class KW_VerifyCardAndPINInfo
    Public Function runKey (strKeyArgument)
        Dim strSummary_CardNumber, strSummary_EmbossedName, strSummary_CardStatus, strSummary_ActionStatus, strSummary_Reason, strSummary_DateTime, strSummary_TaggedBy, strSummary_Brand, strSummary_PINTries, strSummary_PINIssued, strSummary_LastPINIssuedDate, strDetails_OverseasWdl, strDetails_CashLineLink, strDetails_FirstIssued, strDetails_ExpiryDate, strDetails_LastReplaced, strDetails_NoOfCardIssued, strDetails_IssuerID, strDetails_PINGenDate, strDetails_ActivationDate, strDetails_BOI, strDetails_LastTransactionDate, strDetails_LastUpdatedOn, strDetails_LastServiceType, strDetails_FPC, strReplaceHist_OldCardNumber, strReplaceHist_NewCardNumber,strCPFISLinkage

        strSummary_CardNumber  = checkNull(strKeyArgument(0))
        strSummary_EmbossedName  = checkNull(strKeyArgument(1))
		If UCase(strKeyArgument(2))="BLANK"Then
			strSummary_CardStatus  = checkNull(strKeyArgument(2))
			elseIf UCase(strKeyArgument(2))="NULL" then
				strSummary_CardStatus  = checkNull(strKeyArgument(2))
			else
			 strSummary_CardStatus=strKeyArgument(2)
		End If
        strSummary_ActionStatus  = checkNull(strKeyArgument(3))
        strSummary_Reason  = checkNull(strKeyArgument(4))
        strSummary_DateTime  = checkNull(strKeyArgument(5))
        strSummary_TaggedBy  = checkNull(strKeyArgument(6))
        strSummary_Brand  = checkNull(strKeyArgument(7))
        strSummary_PINTries  = checkNull(strKeyArgument(8))
        strSummary_PINIssued  = checkNull(strKeyArgument(9))
        strSummary_LastPINIssuedDate  = checkNull(strKeyArgument(10))
		If UCase(strKeyArgument(11))="BLANK"Then
			  strDetails_OverseasWdl  = checkNull(strKeyArgument(11))
			elseIf UCase(strKeyArgument(11))="NULL" then
				  strDetails_OverseasWdl  = checkNull(strKeyArgument(11))
			else
			  strDetails_OverseasWdl  =strKeyArgument(11)
		End If
      
        strDetails_CashLineLink  = checkNull(strKeyArgument(12))
		strDetails_AccountNo  = checkNull(strKeyArgument(13))
        strDetails_FirstIssued  = checkNull(strKeyArgument(14))
        strDetails_ExpiryDate  = checkNull(strKeyArgument(15))
        strDetails_LastReplaced  = checkNull(strKeyArgument(16))
        strDetails_NoOfCardIssued  = checkNull(strKeyArgument(17))
        strDetails_IssuerID  = checkNull(strKeyArgument(18))
        strDetails_PINGenDate  = checkNull(strKeyArgument(19))
        strDetails_ActivationDate  = checkNull(strKeyArgument(20))
        strDetails_BOI  = checkNull(strKeyArgument(21))
        strDetails_LastTransactionDate  = checkNull(strKeyArgument(22))
        strDetails_LastUpdatedOn  = checkNull(strKeyArgument(23))
        strDetails_LastServiceType  = checkNull(strKeyArgument(24))
        strDetails_FPC  = checkNull(strKeyArgument(25))
        strReplaceHist_OldCardNumber  = checkNull(strKeyArgument(26))
        strReplaceHist_NewCardNumber  = checkNull(strKeyArgument(27))
		strCPFISLinkage = checkNull(strKeyArgument(28))
        Dim bVerifyCardAndPINInfo
        bVerifyCardAndPINInfo  = VerifyCardAndPINInfo(strSummary_CardNumber, strSummary_EmbossedName, strSummary_CardStatus, strSummary_ActionStatus, strSummary_Reason, strSummary_DateTime, strSummary_TaggedBy, strSummary_Brand, strSummary_PINTries, strSummary_PINIssued, strSummary_LastPINIssuedDate, strDetails_OverseasWdl, strDetails_CashLineLink, strDetails_AccountNo, strDetails_FirstIssued, strDetails_ExpiryDate, strDetails_LastReplaced, strDetails_NoOfCardIssued, strDetails_IssuerID, strDetails_PINGenDate, strDetails_ActivationDate, strDetails_BOI, strDetails_LastTransactionDate, strDetails_LastUpdatedOn, strDetails_LastServiceType, strDetails_FPC, strReplaceHist_OldCardNumber, strReplaceHist_NewCardNumber,strCPFISLinkage)

		If bVerifyCardAndPINInfo Then
			runKey = True
		Else
			runKey = False
		End If

   End Function


End Class

'strSummary_CardNumber	Summary_EmbossedName	Summary_CardStatus	Summary_ActionStatus	Summary_Reason	Summary_DateTime	Summary_TaggedBy	Summary_Brand	Summary_PINTries	Summary_PINIssued	Summary_LastPINIssuedDate	Details_OverseasWdl	Details_CashLineLink	Details_FirstIssued	Details_ExpiryDate	Details_LastReplaced	Details_NoOfCardIssued	Details_IssuerID	Details_PINGenDate	Details_ActivationDate	Details_BOI	Details_LastTransactionDate	Details_LastUpdatedOn	Details_LastServiceType	Details_FPC	ReplaceHist_OldCardNumber	ReplaceHist_NewCardNumber
