'*****This is auto generated code using code generator please Re-validate ****************

'[Verify row Data in Table for Mobile Search]
Public Function verifytblSelectedContent_MobileSearch(arrRowDataList)
   bDevPending=false 
	   verifytblSelectedContent_MobileSearch=verifyTableContentList(MobileSearch.tblMobileSearcResultsHeader_Mobile,MobileSearch.tblMobileSearchResultsContent_Mobile,arrRowDataList,"SelectedMobileSearchContent",False,NULL,NULL,NULL)
End Function

'[Set text on MobileSearch textfield for Mobile Search]
Public Function setMobileTxt_MobileSearch(strMobileSearchtxt)
	bsetMobileTxt_MobileSearch=true
	MobileSearch.clickByMobile
	If Not IsNull(strMobileSearchtxt) Then
		MobileSearch.txtMobileSearch().set strMobileSearchtxt
	End If
	WaitForICallLoading
	MobileSearch.clickByMobilebtn
	
	setMobileTxt_MobileSearch=bsetMobileTxt_MobileSearch
End Function

'[Verify error message displays for MobileSearch]
Public Function verifyErrorMessage_MobileSearch(strErrorMessage)
bverifyErrorMessage_MobileSearch = true
	If Not IsNull(strErrorMessage) Then
 		If Not VerifyInnerText (MobileSearch.lblMobileSearchErrMsg(), strErrorMessage, "MobileSearch")Then
			   bverifyErrorMessage_OrgID = false
		End If
	End If
	verifyErrorMessage_MobileSearch = bverifyErrorMessage_MobileSearch
End Function

'[Verify Search result message for MobileSearch]
Public Function verifySearchResultMessage_MobileSearch(strErrorMessage)
bverifySearchResultMessage_MobileSearch = true
WaitForICallLoading
	If Not IsNull(strErrorMessage) Then
 		If Not VerifyInnerText (MobileSearch.lblSearchResultMessageMobile(), strErrorMessage, "MobileSearchResult")Then
			   bverifyErrorMessage_OrgID = false
		End If
	End If
	verifySearchResultMessage_MobileSearch = bverifySearchResultMessage_MobileSearch
End Function

'[Verify InfoWarn]
'Public Function verifyInfoWarn(strInfoMsg)
'bverifyInfoWarn = true
	'If Not IsNull(strInfoMsg) Then
 	'	If Not verifyFieldValue(ReferralRequest.txtInfoWarn(), strInfoMsg, "InfoWarnMsg")Then
	'		   bverifyInfoWarn = false
	'	End If
	'End If
	'verifyInfoWarn = bverifyInfoWarn
'End Function
