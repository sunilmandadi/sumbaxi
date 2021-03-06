'Keyword Generation
Class KW_TriggerCTI
    Public Function runKey (strKeyArgument)

        Dim  strAuthStatus, strCIN, strCINSFX, strCCNumber, strTransactionIndicator, strSourceOwner, strCallerID, strCustName, strExpVerificationStatus, strExpIdentificationStatus, strExpAuthenticationStatus
		bAppStateLogout=true
		strAuthStatus=checkNull(strKeyArgument(0))
		strCIN=checkNull(strKeyArgument(1))
		strCINSFX=checkNull(strKeyArgument(2))
		strCCNumber  = checkNull(strKeyArgument(3))
        strTransactionIndicator  = checkNull(strKeyArgument(4))
        strSourceOwner  = checkNull(strKeyArgument(5))
        strCallerID  = checkNull(strKeyArgument(6))
        strCustName  = checkNull(strKeyArgument(7))
		strExpVerificationStatus = checkNull(strKeyArgument(8))
		strExpIdentificationStatus = checkNull(strKeyArgument(9))
		strExpAuthenticationStatus = checkNull(strKeyArgument(10))

        Dim bTriggerCTI
        bTriggerCTI  =TriggerCTI(strAuthStatus, strCIN, strCINSFX, strCCNumber, strTransactionIndicator, strSourceOwner, strCallerID, strCustName, strExpVerificationStatus, strExpIdentificationStatus, strExpAuthenticationStatus)

        If  bTriggerCTI Then
            LogMessage "RSLT", "Verification", "Keyword Manual Authentication executed successfully", True
            runKey =True
        else
            LogMessage "RSLT", "Verification", "Keyword ManualAuthentication  failed during execution", False
            runKey =False
        End If

    End Function

End Class
