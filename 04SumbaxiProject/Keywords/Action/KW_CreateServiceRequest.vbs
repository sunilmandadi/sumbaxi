'Keyword Generation
Class KW_CreateServiceRequest
    Public Function runKey (strKeyArgument)
        Dim bShortCut,strName,strCIN,strSegment,bVerifyCombo,strRelatedTo,strAccountNo
		Dim strType,strSubType,strProduct,strExpectd_AssignTo,strAssignedTo_New,strStatus,strSubStatus,bOnceAndDone,strComment
    Dim strNotes,bAddNotes,strExpectedProgressStatus,strActivityType,strActivitySubType,strAssignedTo,strActivity_Comment,strErrorMessage
		bShortCut= checkNull(strKeyArgument(0))
        strName  = checkNull(strKeyArgument(1))
        strCIN  = checkNull(strKeyArgument(2))
        strSegment  = checkNull(strKeyArgument(3))
        bVerifyCombo  = checkNull(strKeyArgument(4))
        strRelatedTo  = checkNull(strKeyArgument(5))
        strAccount  = checkNull(strKeyArgument(6))
        strType  = checkNull(strKeyArgument(7))
        strSubType  = checkNull(strKeyArgument(8))
		strDescription=checkNull(strKeyArgument(9))
		strKnoledgeBase=checkNull(strKeyArgument(10))
        strProduct  = checkNull(strKeyArgument(11))
        strExpectd_AssignTo  = checkNull(strKeyArgument(12))
		strAssignedTo_New  = checkNull(strKeyArgument(13))
        strStatus  = checkNull(strKeyArgument(14))
        strSubStatus  = checkNull(strKeyArgument(15))
		strRequestExecuted=checkNull(strKeyArgument(16))
        bOnceAndDone  = checkNull(strKeyArgument(17))
        strComment  = checkNull(strKeyArgument(18))
        strNotes  = checkNull(strKeyArgument(19))
        bAddNotes  = checkNull(strKeyArgument(20))
        strExpectedProgressStatus  = checkNull(strKeyArgument(21))
        
        strActivityType  = checkNull(strKeyArgument(22))
        strActivitySubType  = checkNull(strKeyArgument(23))
        strActivityAssignedTo  = checkNull(strKeyArgument(24))
        strActivity_Comments  = checkNull(strKeyArgument(25))
		bSubmit=checkNull(strKeyArgument(26))
		strErrorMessage=checkNull(strKeyArgument(27))
        Dim bCreateServiceRequest
        bCreateServiceRequest  = CreateServiceRequest(bShortCut,strName,strCIN,strSegment,bVerifyCombo,strRelatedTo,strAccount,strType,strSubType,strDescription,strKnoledgeBase,strProduct,strExpectd_AssignTo,strAssignedTo_New,_ 
   strStatus,strSubStatus,strRequestExecuted,bOnceAndDone,strComment, strNotes,bAddNotes,strExpectedProgressStatus,strActivityType,strActivitySubType,	strActivityAssignedTo,strActivity_Comments,bSubmit,strErrorMessage)

        If  bCreateServiceRequest Then
            LogMessage "RSLT", "Verification", "Keyword CreateServiceRequest executed successfully", True
            runKey =True
        else
            runKey =False
        End If

    End Function


End Class

Class KW_VerifyServiceRequest_Shortcut
    Public Function runKey (strKeyArgument)
        Dim strProductType,strLeftMenuName,lstButtonNames, strButtonToClick,strRelatedTo,strType,lstSubType
		
		strProductType= checkNull(strKeyArgument(0))
        strLeftMenuName  = checkNull(strKeyArgument(1))
        lstButtonNames  = checkNull(strKeyArgument(2))
        strButtonToClick  = checkNull(strKeyArgument(3))
        strRelatedTo  = checkNull(strKeyArgument(4))
        strType  = checkNull(strKeyArgument(5))
        lstSubType  = checkNull(strKeyArgument(6))
      
        Dim bVerifyServiceRequest_Shortcut
        bVerifyServiceRequest_Shortcut = VerifyServiceRequest_Shortcut(strProductType,strLeftMenuName,lstButtonNames, strButtonToClick,strRelatedTo,strType,lstSubType)

        If  bVerifyServiceRequest_Shortcut Then
            LogMessage "RSLT", "Verification", "Keyword CreateServiceRequest_Shortcut executed successfully", True
            runKey =True
        else
            runKey =False
        End If

    End Function


End Class


Class KW_EditServiceRequest
    Public Function runKey (strKeyArgument)
        Dim bShortCut,strName,strCIN,strSegment,bVerifyCombo,strRelatedTo,strAccountNo
		Dim strServiceRequestNo, strAssignedTo,strStatus,strSubStatus, strRequestExecuted,strComment,strNotes
	    Dim lstActivityDetails, strActivityAssignTo,strActivityStatus, strActivityResolution,bActivityOnceandDone,strActivityComment,strActivityCreatedBy
    
		strServiceRequestNo= checkNull(strKeyArgument(0))
		If Ucase(strServiceRequestNo)="DATASTORE" Then
			strServiceRequestNo=fetchFromDataStore ("CreateServiceRequest","Blank", "ServiceRequestID")(0)
		End If
        strAssignedTo  = checkNull(strKeyArgument(1))
        strStatus  = checkNull(strKeyArgument(2))
        strSubStatus  = checkNull(strKeyArgument(3))
        strRequestExecuted  = checkNull(strKeyArgument(4))
		strCreatedBy = checkNull(strKeyArgument(5))
        strComment  = checkNull(strKeyArgument(6))
        strNotes  = checkNull(strKeyArgument(7))
        lstActivityDetails  = checkNull(strKeyArgument(8))
		strLoggedInUser =  checkNull(strKeyArgument(9))
        strActivityAssignTo  = checkNull(strKeyArgument(10))
		strActivityStatus=checkNull(strKeyArgument(11))
		strActivityResolution=checkNull(strKeyArgument(12))
        bActivityOnceandDone  = checkNull(strKeyArgument(13))
        strActivityComment  = checkNull(strKeyArgument(14))
		strActivityCreatedBy  = checkNull(strKeyArgument(15))
		bSubmit= checkNull(strKeyArgument(16))
		strExpectedProgress=checkNull(strKeyArgument(17))
       

        Dim bCreateServiceRequest
        bCreateServiceRequest  = EditServiceRequest(strServiceRequestNo, strAssignedTo,strStatus,strSubStatus, strRequestExecuted,strCreatedBy,strComment,strNotes,_ 
	    lstActivityDetails, strLoggedInUser,strActivityAssignTo,strActivityStatus, strActivityResolution,bActivityOnceandDone,strActivityComment,strActivityCreatedBy,bSubmit,strExpectedProgress  )

        If  bCreateServiceRequest Then
            LogMessage "RSLT", "Verification", "Keyword EditServiceRequest executed successfully", True
            runKey =True
        else
            runKey =False
        End If

    End Function


End Class

