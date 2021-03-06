'*****This is auto generated code using code generator please Re-validate ****************

Dim strRefId, strtCreateDateTime, strMonthName, strRoleRefId, strGrpRefId, strFunMapRefId

'[Select Verification link in Quick Tasks menu list]
Public Function clicklink_Verification()
bDevPending=false
UserProfileAdmin.lnkVerification.click
If Err.Number<>0 Then
clicklink_Verification=false
LogMessage "WARN","Verification","Failed to Click Link : Verification" ,false
Exit Function
End If
WaitForICallLoading
clicklink_Verification=true
End Function

'[Select AccessControl link in Quick Tasks menu list]
Public Function clicklink_lnkAccessControl()
bDevPending=false
UserProfileAdmin.lnkAccessControl.click
If Err.Number<>0 Then
clicklink_lnkAccessControl=false
LogMessage "WARN","Verification","Failed to Click Link : AccessControl" ,false
Exit Function
End If
WaitForICallLoading
clicklink_lnkAccessControl=true
End Function

'[Select Authorize link in Quick Tasks menu list]
Public Function clicklink_lnkAuthorize()
bDevPending=false
UserProfileAdmin.lnkAuthorize.click
If Err.Number<>0 Then
clicklink_lnkAuthorize=false
LogMessage "WARN","Verification","Failed to Click Link : Authorize" ,false
Exit Function
End If
WaitForICallLoading
clicklink_lnkAuthorize=true
End Function

'[Set Channel Combobox as]
Public Function selectChannelComboBox(strChannel)
bselectChannelComboBox=true
If Not IsNull(strChannel) Then
UserProfileAdmin.txtChannel().set strChannel
End If
waitForIcallLoading
selectChannelComboBox=bselectChannelComboBox
End Function

'[Set Functional Map Channel Combobox as]
Public Function selectChannelFuncMapComboBox(strChannel)
bselectChannelFuncMapComboBox=true
If Not IsNull(strChannel) Then
UserProfileAdmin.txtFuncMapChannel().set strChannel
End If
waitForIcallLoading
selectChannelFuncMapComboBox=bselectChannelFuncMapComboBox
End Function

'[Verify Functional Map Channel Combobox]
Public Function verifyFuncMapComboBoxChnl(lstChnlFuncMap)
   bDevPending=false
   bverifyFuncMapComboBoxChnl=true
   If Not IsNull(lstChnlFuncMap) Then
       If Not (verifyComboList (lstChnlFuncMap, UserProfileAdmin.txtFuncMapChannel()))Then
            LogMessage "WARN","Verification","Failed to select :"&lstChnlFuncMap&" From Func Map drop down list" ,false
           bverifyFuncMapComboBoxChnl=false
       End If
   End If
   waitForIcallLoading
   verifyFuncMapComboBoxChnl=bverifyFuncMapComboBoxChnl
End Function

'[Set Parameter Type Combobox as]
Public Function selectParameterTypeComboBox(strParameter)
bselectParameterTypeComboBox=true
If Not IsNull(strChannel) Then
'waitForIcallLoading
Set wshShell = CreateObject( "WScript.Shell" )
'UserProfileAdmin.txtParameterType().FireEvent.onfocus
UserProfileAdmin.txtParameterType().set strParameter
waitForIcallLoading
UserProfileAdmin.txtParameterType().set strParameter
waitForIcallLoading
UserProfileAdmin.txtParameterType().set strParameter
UserProfileAdmin.txtParameterType().FireEvent "onclick"
'UserProfileAdmin.txtParameterType().set strParameter
'WshShell.SendKeys "{TAB}"
waitForIcallLoading
End If
selectParameterTypeComboBox=bselectParameterTypeComboBox
End Function

'[Verify Field GrpDesc displayed as]
Public Function verifyGrpDesc(strGrpDesc)
bDevPending=false
bverifyGrpDesc=true
If Not IsNull(strGroupId) Then
If Not VerifyInnerText (UserProfileAdmin.lblGrpDescription(), strGrpDesc, "GrpDesc")Then
bverifyGrpDesc=false
End If
End If
verifyGrpDesc=bverifyGrpDesc
End Function 

'[Verify pink panel Field UsrGrpDesc displayed as]
Public Function verifyPnkPnlGrpDesc(strGrpDesc)
bDevPending=false
bverifyPnkPnlGrpDesc=true
If Not IsNull(strGroupId) Then
If Not VerifyInnerText (UserProfileAdmin.lblGrpDescription(), strGrpDesc, "GrpDesc")Then
bverifyPnkPnlGrpDesc=false
End If
End If
verifyPnkPnlGrpDesc=bverifyPnkPnlGrpDesc
End Function 

'[Click Button Add New Values on Verification Screen]
Public Function clickButtonAddNewValues_UserProfileAdmin()
bDevPending=false
UserProfileAdmin.btnAddNewValues.click
If Err.Number<>0 Then
clickButtonAddNewValues_UserProfileAdmin=false
LogMessage "WARN","Verification","Failed to Click Button : AddNewValues" ,false
Exit Function
End If
waitForIcallLoading
clickButtonAddNewValues_UserProfileAdmin=true
End Function

'[Click User Profile Button Save Submit]
Public Function clickButtonSave()
bDevPending=true
strtCreateDateTime = now
UserProfileAdmin.btnSave.click
If Err.Number<>0 Then
clickButtonSave=false
LogMessage "WARN","Verification","Failed to Click Button : Save" ,false
Exit Function
End If
waitForIcallLoading
clickButtonSave=true
End Function

'[Click User Profile Button Save]
Public Function clickButtonSave1()
bDevPending=true
waitForIcallLoading
UserProfileAdmin.btnSave.click
If Err.Number<>0 Then
clickButtonSave1=false
LogMessage "WARN","Verification","Failed to Click Button : Save" ,false
Exit Function
End If
waitForIcallLoading
clickButtonSave1=true
End Function

'[Click User Profile Button Submit]
Public Function clickButtonSubmitUsrProf()
bDevPending=true
waitForIcallLoading
UserProfileAdmin.btnSubmit.click
If Err.Number<>0 Then
clickButtonSubmitUsrProf=false
LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
Exit Function
End If
waitForIcallLoading
clickButtonSubmitUsrProf=true
End Function

'[Click User Profile Button Ok]
Public Function clickButtonOk()
bDevPending=true
waitForIcallLoading
strRefIdVal=UserProfileAdmin.lblIRefID.GetROProperty("innertext")
arrTemp=Split(strRefIdVal,":")
'msgbox arrTemp(1)
strRefId=arrTemp(1)
'msgbox strRefId
UserProfileAdmin.btnOK.click
If Err.Number<>0 Then
clickButtonOk=false
LogMessage "WARN","Verification","Failed to Click Button : OK" ,false
Exit Function
End If
waitForIcallLoading
clickButtonOk=true
End Function

'[Click Approve/Reject hyperlink on Authorize page]
Public Function viewAuthorizeDetails(strChannel, strUserName)
Dim bviewAuthorizeDetails: viewAuthorizeDetails = True
lstAuthorizeDetails = checknull("Channel:"&strChannel&"|Reference ID:"&strRefId&"|Created / Modified By:"&strUserName&"")
If MatchStr(strRefId,"FM.*") Then
	strFunMapRefId=strRefId
	ElseIf MatchStr(strRefId,"UR.*") Then
	strRoleRefId=strRefId
	ElseIf MatchStr(strRefId,"UG.*") Then
	strGrpRefId=strRefId
End If

bviewAuthorizeDetails=selectTableLink(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstAuthorizeDetails,"UserProfileAdmin" ,"Action",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
viewAuthorizeDetails=bviewAuthorizeDetails
End Function

'[Click User Profile View Link]
Public Function lnkViewUserProfile(strGrpId, strDesc, strtChnlId, strDBServer)
Dim blnkViewUserProfile: lnkViewUserProfile = True
strQuery="select CREATED_DATETIME from iserve_user_groups_wrk where GROUP_ID='"&strGrpId&"' order by CREATED_DATETIME desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)
lstViewUserProfile = checknull("Group ID:"&strGrpId&"|Description:"&strDesc&"|Channel:"&strtChnlId&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
blnkViewUserProfile=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewUserProfile,"UserProfileAdmin" ,"Action", "View",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkViewUserProfile=blnkViewUserProfile
End Function

'[Click User Profile Modify Link]
Public Function lnkModifyUserProfile(strGrpId, strDesc, strtChnlId, strDBServer)
Dim blnkModifyUserProfile: lnkModifyUserProfile = True
strQuery="select CREATED_DATETIME from iserve_user_groups_wrk where GROUP_ID='"&strGrpId&"' order by CREATED_DATETIME desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)
lstModifyUserProfile = checknull("Group ID:"&strGrpId&"|Description:"&strDesc&"|Channel:"&strtChnlId&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
blnkModifyUserProfile=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyUserProfile,"UserProfileAdmin" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkModifyUserProfile=blnkModifyUserProfile
End Function

'[Set User profile GroupId as]
Public Function selectGrpId(strGrpId)
bselectGrpId=true
If Not IsNull(strGrpId) Then
UserProfileAdmin.txtUsrGrpGroupId().set strGrpId
End If
waitForIcallLoading
selectGrpId=bselectGrpId
End Function

'[Set User profile GroupDesc as]
Public Function selectGrpDesc(strGrpDesc)
bselectGrpDesc=true
If Not IsNull(strGrpDesc) Then
UserProfileAdmin.txtUsrGrpDesc().set strGrpDesc
End If
waitForIcallLoading
selectGrpDesc=bselectGrpDesc
End Function

'[Set User profile GroupChannel as]
Public Function selectGrpChnl(strGrpChnl)
bselectGrpChnl=true
If Not IsNull(strGrpChnl) Then
UserProfileAdmin.txtusrGrpChannel().set strGrpChnl
End If
waitForIcallLoading
selectGrpChnl=bselectGrpChnl
End Function

'[Set User profile Group UserRole as]
Public Function selectGrpusrRole(strGrpUsrRole)
bselectGrpusrRole=true
If Not IsNull(strGrpUsrRole) Then
UserProfileAdmin.txtUsrGrpUsrRole().set strGrpUsrRole
UserProfileAdmin.txtUsrGrpUsrRole().FireEvent "onmousedown"
UserProfileAdmin.txtUsrGrpUsrRole().FireEvent "onclick"
UserProfileAdmin.txtUsrGrpUsrRole().set strGrpUsrRole
End If
waitForIcallLoading
selectGrpusrRole=bselectGrpusrRole
End Function

'[Set User profile GroupStatus as]
Public Function selectGrpStatus(strGrpSts)
bselectGrpStatus=true
If Not IsNull(strGrpSts) Then
UserProfileAdmin.txtUsrGrpStatus().set strGrpSts
End If
waitForIcallLoading
selectGrpStatus=bselectGrpStatus
End Function

'[Verify User Group validation message displayed as]
Public Function verifyUserGrplValMsg(strUsrGrpMsg)
bDevPending=false
bverifyUserGrplValMsg=true
If Not IsNull(strUsrGrpMsg) Then
If Not VerifyInnerText (UserProfileAdmin.lblGrpValidationMsg(), strUsrGrpMsg, "UserGrpValidationMsg")Then
bverifyUserGrplValMsg=false
End If
End If
verifyUserGrplValMsg=bverifyUserGrplValMsg
End Function

'[Set User profile Approve Reject Comments as]
Public Function selectApprvRjct(strApprvRjct)
bselectApprvRjct=true
If Not IsNull(strApprvRjct) Then
UserProfileAdmin.txtApprvRjctCmnt().set strApprvRjct
End If
waitForIcallLoading
selectApprvRjct=bselectApprvRjct
End Function

'[Verify Field User Profile GrpId displayed as]
Public Function verifyUserProfileGrpId(strGrpId)
bDevPending=false
bverifyUserProfileGrpId=true
If Not IsNull(strGrpId) Then
If Not verifyFieldValue (UserProfileAdmin.lblUsrGrpGroupId(), strGrpId, "GrpId")Then
bverifyUserProfileGrpId=false
End If
End If
verifyUserProfileGrpId=bverifyUserProfileGrpId
End Function 

'[Verify Field User Profile GrpDesc displayed as]
Public Function verifyUserProfileGrpDesc(strGrpDesc)
bDevPending=false
bverifyUserProfileGrpDesc=true
If Not IsNull(strGrpDesc) Then
If Not verifyFieldValue (UserProfileAdmin.lblUsrGrpDesc(), strGrpDesc, "GrpDesc")Then
bverifyUserProfileGrpDesc=false
End If
End If
verifyUserProfileGrpDesc=bverifyUserProfileGrpDesc
End Function 

'[Verify Field User Profile GrpChnl displayed as]
Public Function verifyUserProfileGrpChnl(strGrpChnl)
bDevPending=false
bverifyUserProfileGrpChnl=true
If Not IsNull(strGrpChnl) Then
If Not verifyFieldValue (UserProfileAdmin.lblusrGrpChannel(), strGrpChnl, "GrpChnl")Then
bverifyUserProfileGrpChnl=false
End If
End If
verifyUserProfileGrpChnl=bverifyUserProfileGrpChnl
End Function 

'[Verify Field User Profile GrpRole displayed as]
Public Function verifyUserProfileGrpRole(strGrpRole)
bDevPending=false
bverifyUserProfileGrpRole=true
If Not IsNull(strGrpRole) Then
If Not verifyFieldValue (UserProfileAdmin.lblUsrGrpUsrRole(), strGrpRole, "GrpRole")Then
bverifyUserProfileGrpRole=false
End If
End If
verifyUserProfileGrpRole=bverifyUserProfileGrpRole
End Function 

'[Verify Field User Profile GrpStatus displayed as]
Public Function verifyUserProfileGrpStatus(strGrpsts)
bDevPending=false
bverifyUserProfileGrpStatus=true
If Not IsNull(strGrpsts) Then
If Not verifyFieldValue (UserProfileAdmin.lblUsrGrpStatus(), strGrpsts, "GrpStatus")Then
bverifyUserProfileGrpStatus=false
End If
End If
verifyUserProfileGrpStatus=bverifyUserProfileGrpStatus
End Function

'[Verify Field User Profile RefId displayed as]
Public Function verifyUserProfileGrpRefId(strGrpRefId)
bDevPending=false
bverifyUserProfileGrpRefId=true
If Not IsNull(strGrpRefId) Then
If Not verifyFieldValue (UserProfileAdmin.lblIRefID(), strGrpRefId, "GrpRefId")Then
bverifyUserProfileGrpRefId=false
End If
End If
verifyUserProfileGrpRefId=bverifyUserProfileGrpRefId
End Function

'[Verify Field User Profile Old GrpId displayed as]
Public Function verifyOldUserProfileGrpId(strGrpId)
bDevPending=false
bverifyOldUserProfileGrpId=true
If Not IsNull(strGrpId) Then
If Not verifyFieldValue (UserProfileAdmin.lblOldUsrGrpGroupId(), strGrpId, "OldGrpId")Then
bverifyOldUserProfileGrpId=false
End If
End If
verifyOldUserProfileGrpId=bverifyOldUserProfileGrpId
End Function 

'[Verify Field User Profile Old GrpDesc displayed as]
Public Function verifyOldUserProfileGrpDesc(strGrpDesc)
bDevPending=false
bverifyOldUserProfileGrpDesc=true
If Not IsNull(strGrpDesc) Then
If Not verifyFieldValue (UserProfileAdmin.lblOldUsrGrpDesc(), strGrpDesc, "OldGrpDesc")Then
bverifyOldUserProfileGrpDesc=false
End If
End If
verifyOldUserProfileGrpDesc=bverifyOldUserProfileGrpDesc
End Function 

'[Verify Field old User Profile GrpChnl displayed as]
Public Function verifyOldUserProfileGrpChnl(strGrpChnl)
bDevPending=false
bverifyOldUserProfileGrpChnl=true
If Not IsNull(strGrpChnl) Then
If Not verifyFieldValue (UserProfileAdmin.lblOldusrGrpChannel(), strGrpChnl, "OldGrpChnl")Then
bverifyOldUserProfileGrpChnl=false
End If
End If
verifyOldUserProfileGrpChnl=bverifyOldUserProfileGrpChnl
End Function 

'[Verify Field User Profile Old GrpRole displayed as]
Public Function verifyOldUserProfileGrpRole(strGrpRole)
bDevPending=false
bverifyOldUserProfileGrpRole=true
If Not IsNull(strGrpRole) Then
If Not verifyFieldValue (UserProfileAdmin.lblOldUsrGrpUsrRole(), strGrpRole, "OldGrpRole")Then
bverifyOldUserProfileGrpRole=false
End If
End If
verifyOldUserProfileGrpRole=bverifyOldUserProfileGrpRole
End Function 

'[Verify Field User Profile Old GrpStatus displayed as]
Public Function verifyOldUserProfileGrpStatus(strGrpsts)
bDevPending=false
bverifyOldUserProfileGrpStatus=true
If Not IsNull(strGrpsts) Then
If Not verifyFieldValue (UserProfileAdmin.lblOldUsrGrpStatus(), strGrpsts, "OldGrpStatus")Then
bverifyOldUserProfileGrpStatus=false
End If
End If
verifyOldUserProfileGrpStatus=bverifyOldUserProfileGrpStatus
End Function

'[Click User profile Button Reject in Approve/reject screen]
Public Function clickButtonReject_ApproveReject()
bDevPending=false
UserProfileAdmin.btnReject.click
If Err.Number<>0 Then
clickButtonReject_ApproveReject=false
LogMessage "WARN","Verification","Failed to Click Button : Reject in Approve/reject screen",false 
Exit Function
End If
clickButtonReject_ApproveReject=true
End Function

'[Click Button Approve in Approve/reject screen]
Public Function clickButtonApprove_ApproveReject()
bDevPending=false
UserProfileAdmin.btnApprove.click
If Err.Number<>0 Then
clickButtonApprove_ApproveReject=false
LogMessage "WARN","Verification","Failed to Click Button : Approve in Approve/reject screen",false 
Exit Function
End If
clickButtonApprove_ApproveReject=true
End Function


'[Click Button close on User profile admin popup Screen]
Public Function clickButtonClose_AddNewValues()
bDevPending=false
UserProfileAdmin.btnClose.click
If Err.Number<>0 Then
clickButtonClose_AddNewValues=false
LogMessage "WARN","Verification","Failed to Click Button : Close In Add New Values",false 
Exit Function
End If
clickButtonClose_AddNewValues=true
End Function

'[Click Button cancel on User profile admin popup Screen]
Public Function clickButtonCancel_AddNewValues()
bDevPending=false
UserProfileAdmin.btnCancel1.click
If Err.Number<>0 Then
clickButtonCancel_AddNewValues=false
LogMessage "WARN","Verification","Failed to Click Button : Cancel In Add New Values",false 
Exit Function
End If
clickButtonCancel_AddNewValues=true
End Function


'**************Added by Kalai as part of Other Auth Categeroy 4112016************************************
''[Select Verification link in Quick Tasks menu list]
'Public Function clicklink_Verification()
'bDevPending=false
'UserProfileAdmin.lnkVerification.click
'	If Err.Number<>0 Then
'	clicklink_Verification=false
'	LogMessage "WARN","Verification","Failed to Click Link : Verification" ,false
'	Exit Function
'	End If
'WaitForICallLoading
'clicklink_Verification=true
'End Function

'[Set text on Auth Category textfield for Add new Values Popup]
Public Function setAuthCategory_AddNewValues(strAuthCategoryTxt)
bsetAuthCategory_AddNewValues=true
	If Not IsNull(strAuthCategoryTxt) Then
	UserProfileAdmin.txtAuthCategory().set strAuthCategoryTxt
	End If
WaitForICallLoading
setAuthCategory_AddNewValues=bsetAuthCategory_AddNewValues
End Function

'[Set text on Description textfield for Add new Values Popup]
Public Function setDescription_AddNewValues(strDescriptionTxt)
bsetDescription_AddNewValues=true
	If Not IsNull(strDescriptionTxt) Then
	UserProfileAdmin.txtDescription().set strDescriptionTxt
	End If
WaitForICallLoading
setDescription_AddNewValues=bsetDescription_AddNewValues
End Function

'[Set text on IVR Hotline textfield for Add new Values Popup]
Public Function setIVRHotline_AddNewValues(strIVRHotlineTxt)
bsetIVRHotline_AddNewValues=true
	If Not IsNull(strIVRHotlineTxt) Then
	UserProfileAdmin.txtIVRHotline().set strIVRHotlineTxt
	End If
WaitForICallLoading
setIVRHotline_AddNewValues=bsetIVRHotline_AddNewValues
End Function

'[Set text on IVR Menu textfield for Add new Values Popup]
Public Function setIVRMenu_AddNewValues(strIVRMenuTxt)
bsetIVRMenu_AddNewValues=true
	If Not IsNull(strIVRMenuTxt) Then
	UserProfileAdmin.txtIVRMenu().set strIVRMenuTxt
	End If
WaitForICallLoading
setIVRMenu_AddNewValues=bsetIVRMenu_AddNewValues
End Function

'[Click Button Save on Add New Values popup Screen]
Public Function clickButtonSave_AddNewValues()
bDevPending=false
UserProfileAdmin.btnSave.click
	If Err.Number<>0 Then
	clickButtonSave_AddNewValues=false
	LogMessage "WARN","Verification","Failed to Click Button : Save In Add New Values",false 
	Exit Function
	End If
clickButtonSave_AddNewValues=true
End Function

'[Verify Field AuthCategory in View Auth displayed as]
Public Function verifyAuthCategoryText(strAuthCategoryTxt)
   bDevPending=false
   bVerifyAuthCategoryText=true
   If Not IsNull(strAuthCategoryTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblAuthCategory(), strAuthCategoryTxt, "Auth Category")Then
	   bVerifyAuthCategoryText=false
	End If
   End If
   verifyAuthCategoryText=bVerifyAuthCategoryText
End Function

'[Verify Field Description in View Auth displayed as]
Public Function verifyDescriptionText(strDescriptionTxt)
   bDevPending=false
   bVerifyDescriptionText=true
   If Not IsNull(strDescriptionTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblDescription(), strDescriptionTxt, "Description")Then
	   bVerifyDescriptionText=false
	End If
   End If
   verifyDescriptionText=bVerifyDescriptionText
End Function

'[Verify Field IVR Hotline in View Auth displayed as]
Public Function verifyIVRHotlineText(strIVRHotlineTxt)
   bDevPending=false
   bVerifyIVRHotlineText=true
   If Not IsNull(strIVRHotlineTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblIVRHotline(), strIVRHotlineTxt, "IVR Hotline")Then
	   bVerifyIVRHotlineText=false
	End If
   End If
   verifyIVRHotlineText=bVerifyIVRHotlineext
End Function

'[Verify Field IVR Menu in View Auth displayed as]
Public Function verifyIVRMenuText(strIVRMenuTxt)
   bDevPending=false
   bVerifyIVRMenuText=true
   If Not IsNull(strIVRMenuTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblIVRMenu(), strIVRMenuTxt, "IVR Menu")Then
	   bVerifyIVRMenuText=false
	End If
   End If
   verifyIVRMenuText=bVerifyIVRMenuText
End Function

'[Verify Field Status in View Auth displayed as]
Public Function verifyStatusText(strStatusTxt)
   bDevPending=false
   bVerifyStatusText=true
   If Not IsNull(strStatusTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblStatus(), strStatusTxt, "Status")Then
	   bVerifyStatusText=false
	End If
   End If
   verifyStatusText=bVerifyStatusText
End Function

'[Verify Field Created By in View Auth displayed as]
Public Function verifyCreatedByText(strCreatedByTxt)
   bDevPending=false
   bVerifyCreatedByText=true
   If Not IsNull(strCreatedByTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblCreatedBy(), strCreatedByTxt, "Created By")Then
	   bVerifyCreatedByText=false
	End If
   End If
   verifyCreatedByText=bVerifyCreatedByText
End Function

'[Verify Field Created Date in View Auth displayed as]
Public Function verifyCreatedDateText(strCreatedDateTxt)
   bDevPending=false
   bVerifyCreatedDateText=true
   If Not IsNull(strCreatedDateTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblCreateDate(), strCreatedDateTxt, "Created Date")Then
	   bVerifyCreatedDateText=false
	End If
   End If
   verifyCreatedDateText=bVerifyCreatedDateText
End Function

'[Verify Field Updated By in View Auth displayed as]
Public Function verifyUpdatedByText(strUpdatedByTxt)
   bDevPending=false
   bVerifyUpdatedByText=true
   If Not IsNull(strUpdatedByTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblLastUpdatedBy(), strUpdatedByTxt, "Updated By")Then
	   bVerifyUpdatedByText=false
	End If 
   End If
   verifyUpdatedByText=bVerifyUpdatedByText
End Function

'[Verify Field Updated Date in View Auth displayed as]
Public Function verifyCreatedDateText(strUpdatedDateTxt)
   bDevPending=false
   bVerifyUpdatedDateText=true
   If Not IsNull(strUpdatedDateTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblLastUpdatedDate(), strUpdatedDateTxt, "Updated Date")Then
	   bVerifyUpdatedDateText=false
	End If
   End If
   verifyUpdatedDateText=bVerifyUpdatedDateText
End Function

'[Verify Field Approved By in View Auth displayed as]
Public Function verifyApprovedByText(strApprovedByTxt)
   bDevPending=false
   bVerifyApprovedByText=true
   If Not IsNull(strApprovedByTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblLastApprovedBy(), strApprovedByTxt, "Approved By")Then
	   bVerifyApprovedByText=false
	End If 
   End If
   verifyApprovedByText=bVerifyApprovedByText
End Function

'[Verify Field Approved Date in View Auth displayed as]
Public Function verifyApprovedDateText(strApprovedDateTxt)
   bDevPending=false
   bVerifyApprovedDateText=true
   If Not IsNull(strApprovedDateTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblLastApprovedDate(), strApprovedDateTxt, "Approved Date")Then
	   bVerifyApprovedDateText=false
	End If
   End If
   verifyApprovedDateText=bVerifyApprovedDateText
End Function

'[Verify Field Auth category old value in Approve/reject screen]
Public Function verifyAutholdvalueText(strAutholdvalueTxt)
   bDevPending=false
   bVerifyAutholdvalueText=true
   If Not IsNull(strAutholdvalueTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblAuthCategoryOldValue(), strAutholdvalueTxt, "Auth oldvalue Txt")Then
	   bVerifyAutholdvalueText=false
	End If
   End If
   verifyAutholdvalueText=bVerifyAutholdvalueText
End Function

'[Verify Field Description old value in Approve/reject screen]
Public Function verifyDescriptionoldvalueText(strDescriptionoldvalueTxt)
   bDevPending=false
   bVerifyDescriptionoldvalueText=true
   If Not IsNull(strDescriptionoldvalueTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblDescriptionOldValue(), strDescriptionoldvalueTxt, "Description oldvalue Txt")Then
	   bVerifyDescriptionoldvalueText=false
	End If
   End If
   verifyDescriptionoldvalueText=bVerifyDescriptionoldvalueText
End Function

'[Verify Field IVR Hotline old value in Approve/reject screen]
Public Function verifyIVRHotlineoldvalueText(strIVRHotlineoldvalueTxt)
   bDevPending=false
   bVerifyIVRHotlineoldvalueText=true
   If Not IsNull(strIVRHotlineoldvalueTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblIVRHotlineOldValue(), strIVRHotlineoldvalueTxt, "IVR Hotline oldvalue Txt")Then
	   bVerifyIVRHotlineoldvalueText=false
	End If
   End If
   verifyIVRHotlineoldvalueText=bVerifyIVRHotlineoldvalueText
End Function

'[Verify Field IVR Menu old value in Approve/reject screen]
Public Function verifyIVRMenuoldvalueText(strIVRMenuoldvalueTxt)
   bDevPending=false
   bVerifyIVRMenuoldvalueText=true
   If Not IsNull(strIVRMenuoldvalueTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblIVRMenuOldValue(), strIVRMenuoldvalueTxt, "IVR Menu oldvalue Txt")Then
	   bVerifyIVRMenuoldvalueText=false
	End If
   End If
   verifyIVRHotlineoldvalueText=bVerifyIVRHotlineoldvalueText
End Function

'[Verify Field Status old value in Approve/reject screen]
Public Function verifyStatusuoldvalueText(strStatusoldvalueTxt)
   bDevPending=false
   bVerifyStatusoldvalueText=true
   If Not IsNull(strStatusoldvalueTxt) Then
     If Not verifyFieldValue (UserProfileAdmin.lblIVRMenuOldValue(), strStatusoldvalueTxt, "Status oldvalue Txt")Then
	   bVerifyStatusoldvalueText=false
	End If
   End If
   verifyStatusoldvalueText=bVerifyStatusoldvalueText
End Function
'************************End********************************************************************************

'[Select SR Related To Combobox as]
Public Function selectSRRelatedComboBox(strSRRelated)
	strSelectSRRelatedComboBox=true
	If Not IsNull(strSRRelated) Then
		UserProfileAdmin.txtSRRelatedTo().set strSRRelated
	End If
	waitForIcallLoading
	selectSRRelatedComboBox=strSelectSRRelatedComboBox
End Function


'[Select User Profile New SR Related To as]
Public Function selectNewSRRelated(strNewSRRelatedTo)
	bselectNewSRRelated=true
	If Not IsNull(strNewSRRelatedTo) Then
		UserProfileAdmin.txtNewSRRelatedTo().set strNewSRRelatedTo
	End If
	waitForIcallLoading
	selectNewSRRelated=bselectNewSRRelated
End Function

'[Verify Field SR Related To displayed as]
Public Function verifyPASRRelatedTo(strRelatedTo)
   bDevPending=false
   bverifyPASRRelatedTo=true
   If Not IsNull(strRelatedTo) Then
     If Not verifyFieldValue (UserProfileAdmin.txtNewSRRelatedTo(), strRelatedTo, "Related To")Then
	   bverifyPASRRelatedTo=false
	End If
   End If
   verifyPASRRelatedTo=bverifyPASRRelatedTo
End Function

'[Select SR Type To Combobox as]
Public Function selectSRTypeComboBox(strSRType)
	strselectSRTypeComboBox=true
	If Not IsNull(strSRType) Then
		UserProfileAdmin.txtSRType().set strSRType
	End If
	waitForIcallLoading
	selectSRTypeComboBox=strselectSRTypeComboBox
End Function


'[Set New SR Type as]
Public Function setNewSRType(strNewSRType)
	strsetNewSRType=true
	If Not IsNull(strNewSRType) Then
		UserProfileAdmin.txtNewSRType().set strNewSRType
	End If
	waitForIcallLoading
	setNewSRType=strsetNewSRType
End Function

'[Verify Field SR Type displayed as]
Public Function verifyPASRType(strSRType)
   bDevPending=false
   bverifyPASRType=true
   If Not IsNull(strSRType) Then
     If Not verifyFieldValue (UserProfileAdmin.txtNewSRType(), strSRType, "SRType")Then
	   bverifyPASRType=false
	End If
   End If
   verifyPASRType=bverifyPASRType
End Function

'[Set SR Sub Type as]
Public Function setSRSubType(strSRSubType)
	strsetSRSubType=true
	If Not IsNull(strSRSubType) Then
		UserProfileAdmin.txtSRSubType().set strSRSubType
	End If
	waitForIcallLoading
	setSRSubType=strsetSRSubType
End Function

'[Verify Field SR Sub Type displayed as]
Public Function verifyPASRSubType(strSRSubType)
   bDevPending=false
   bverifyPASRSubType=true
   If Not IsNull(strSRSubType) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRSubType(), strSRSubType, "SRSubType")Then
	   bverifyPASRSubType=false
	End If
   End If
   verifyPASRSubType=bverifyPASRSubType
End Function

'[Select Assigned To ICall Combobox as]
Public Function selectAssignedToICallComboBox(strAssignedToICall)
	strSelectAssignedToICallComboBox=true
	If Not IsNull(strAssignedToICall) Then
		UserProfileAdmin.txtSRAssignedToICall().set strAssignedToICall
	End If
	waitForIcallLoading
	selectAssignedToICallComboBox=strSelectAssignedToICallComboBox
End Function

'[Verify Field Assigned to ICall displayed as]
Public Function verifyPASRAsgndTo(strSRAsgndTo)
   bDevPending=false
   bverifyPASRAsgndTo=true
   If Not IsNull(strSRAsgndTo) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRAssignedToICall(), strSRAsgndTo, "SRAssignedTo")Then
	   bverifyPASRAsgndTo=false
	End If
   End If
   verifyPASRAsgndTo=bverifyPASRAsgndTo
End Function

'[Set CRM Related To as]
Public Function setCRMRelatedTo(strCRMRelatedTo)
	strSetCRMRelatedTo=true
	If Not IsNull(strCRMRelatedTo) Then
		UserProfileAdmin.txtCRMRelatedTo().set strCRMRelatedTo
	End If
	waitForIcallLoading
	setCRMRelatedTo=strSetCRMRelatedTo
End Function

'[Verify Field CRM Related To displayed as]
Public Function verifyCRMRelatedTo(strCRMRltdTo)
   bDevPending=false
   bverifyCRMRelatedTo=true
   If Not IsNull(strCRMRltdTo) Then
     If Not verifyFieldValue (UserProfileAdmin.txtCRMRelatedTo(), strCRMRltdTo, "CRMRelatedTO")Then
	   bverifyCRMRelatedTo=false
	End If
   End If
   verifyCRMRelatedTo=bverifyCRMRelatedTo
End Function

'[Set CRM Type as]
Public Function selectCRMType(strCRMType)
	bselectCRMType=true
	If Not IsNull(strCRMType) Then
		UserProfileAdmin.txtCRMType().set strCRMType
	End If
	waitForIcallLoading
	selectCRMType=bselectCRMType
End Function

'[Verify Field CRM Type displayed as]
Public Function verifyCRMType(strCRMtype)
   bDevPending=false
   bverifyCRMType=true
   If Not IsNull(strCRMtype) Then
     If Not verifyFieldValue (UserProfileAdmin.txtCRMType(), strCRMtype, "CRMType")Then
	   bverifyCRMType=false
	End If
   End If
   verifyCRMType=bverifyCRMType
End Function

'[Set CRM Sub Type as]
Public Function selectCRMSubType(strCRMSubType)
	bselectCRMSubType=true
	If Not IsNull(strCRMSubType) Then
		UserProfileAdmin.txtCRMSubType().set strCRMSubType
	End If
	waitForIcallLoading
	selectCRMSubType=bselectCRMSubType
End Function

'[Verify Field CRM Sub Type displayed as]
Public Function verifyCRMSubType(strCRMSubType)
   bDevPending=false
   bverifyCRMSubType=true
   If Not IsNull(strCRMSubType) Then
     If Not verifyFieldValue (UserProfileAdmin.txtCRMSubType(), strCRMSubType, "CRMSubType")Then
	   bverifyCRMSubType=false
	End If
   End If
   verifyCRMSubType=bverifyCRMSubType
End Function

'***********Added by Kalyan 29112016********************

'[Set GWF Project Code as]
Public Function selectGWFProjCode(strGWFProjCode)
	bselectGWFProjCode=true
	If Not IsNull(strGWFProjCode) Then
		UserProfileAdmin.txtSRGWFProjCode().set strGWFProjCode
	End If
	waitForIcallLoading
	selectGWFProjCode=bselectGWFProjCode
End Function

'[Verify Field GWF Project Code displayed as]
Public Function verifyGWFProjCode(strGWFProjCde)
   bDevPending=false
   bverifyGWFProjCode=true
   If Not IsNull(strGWFProjCde) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRGWFProjCode(), strGWFProjCde, "GWFProjCode")Then
	   bverifyGWFProjCode=false
	End If
   End If
   verifyGWFProjCode=bverifyGWFProjCode
End Function

'[Set GWF Project Description as]
Public Function selectGWFProjDesc(strGWFProjDesc)
	bselectGWFProjDesc=true
	If Not IsNull(strGWFProjDesc) Then
		UserProfileAdmin.txtSRGWFProjDesc().set strGWFProjDesc
	End If
	waitForIcallLoading
	selectGWFProjDesc=bselectGWFProjDesc
End Function

'[Verify Field GWF Project Desc displayed as]
Public Function verifyGWFProjDesc(strGWFProjDesc)
   bDevPending=false
   bverifyGWFProjDesc=true
   If Not IsNull(strGWFProjDesc) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRGWFProjDesc(), strGWFProjDesc, "GWFProjDesc")Then
	   bverifyGWFProjDesc=false
	End If
   End If
   verifyGWFProjDesc=bverifyGWFProjDesc
End Function

'[Set GWF Product Code as]
Public Function selectGWFProdCode(strGWFProdCode)
	bselectGWFProdCode=true
	If Not IsNull(strGWFProdCode) Then
		UserProfileAdmin.txtSRGWFProdCode().set strGWFProdCode
	End If
	waitForIcallLoading
	selectGWFProdCode=bselectGWFProdCode
End Function

'[Verify Field GWF Product Code displayed as]
Public Function verifyGWFProdCode(strGWFProdcode)
   bDevPending=false
   bverifyGWFProdCode=true
   If Not IsNull(strGWFProdcode) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRGWFProdCode(), strGWFProdcode, "GWFProdCode")Then
	   bverifyGWFProdCode=false
	End If
   End If
   verifyGWFProdCode=bverifyGWFProdCode
End Function

'[Set GWF Product Desc as]
Public Function selectGWFProdDesc(strGWFProdDesc)
	bselectGWFProdDesc=true
	If Not IsNull(strGWFProdDesc) Then
		UserProfileAdmin.txtSRGWFProdDesc().set strGWFProdDesc
	End If
	waitForIcallLoading
	selectGWFProdDesc=bselectGWFProdDesc
End Function

'[Verify Field GWF Product Desc displayed as]
Public Function verifyGWFProdDesc(strGWFProdDesc)
   bDevPending=false
   bverifyGWFProdDesc=true
   If Not IsNull(strGWFProdDesc) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRGWFProdDesc(), strGWFProdDesc, "GWFProdDesc")Then
	   bverifyGWFProdDesc=false
	End If
   End If
   verifyGWFProdDesc=bverifyGWFProdDesc
End Function


'[Set GWF Doc Type Code as]
Public Function selectGWFDocType(strGWFDocTyp)
	bselectGWFDocType=true
	If Not IsNull(strGWFDocTyp) Then
		UserProfileAdmin.txtSRGWFDocTypCode().set strGWFDocTyp
	End If
	waitForIcallLoading
	selectGWFDocType=bselectGWFDocType
End Function

'[Verify Field GWF Doc Type displayed as]
Public Function verifyGWFDocType(strGWFDoctyp)
   bDevPending=false
   bverifyGWFDocType=true
   If Not IsNull(strGWFDoctyp) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRGWFDocTypCode(), strGWFDoctyp, "GWFDocType")Then
	   bverifyGWFDocType=false
	End If
   End If
   verifyGWFDocType=bverifyGWFDocType
End Function

'[Set GWF Doc Type Desc as]
Public Function selectGWFDocTypeDesc(strGWFDocTypDesc)
	bselectGWFDocTypeDesc=true
	If Not IsNull(strGWFDocTypDesc) Then
		UserProfileAdmin.txtSRGWFDocTypDesc().set strGWFDocTypDesc
	End If
	waitForIcallLoading
	selectGWFDocTypeDesc=bselectGWFDocTypeDesc
End Function

'[Verify Field GWF Doc Type Description displayed as]
Public Function verifyGWFDocTypeDesc(strGWFDoctypDesc)
   bDevPending=false
   bverifyGWFDocTypeDesc=true
   If Not IsNull(strGWFDoctypDesc) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRGWFDocTypDesc(), strGWFDoctypDesc, "GWFDocTypeDesc")Then
	   bverifyGWFDocTypeDesc=false
	End If
   End If
   verifyGWFDocTypeDesc=bverifyGWFDocTypeDesc
End Function

'************************End*************************************

'[Select TM Approval Indicator Combobox as]
Public Function selectTMApprovalIndicator(strTMApprovalIndicator)
	strSelectTMApprovalIndicator=true
	If Not IsNull(strTMApprovalIndicator) Then
		UserProfileAdmin.txtTMApprovalIndicator().set strTMApprovalIndicator
	End If
	waitForIcallLoading
	selectTMApprovalIndicator=strSelectTMApprovalIndicator
End Function

'[Set Minimum Identification Points as]
Public Function setMinIdentificationPoints(strMinIdentificationPoints)
	strSetMinIdentificationPoints=true
	If Not IsNull(strMinIdentificationPoints) Then
		UserProfileAdmin.txtMinIdentificationPoints().set strMinIdentificationPoints
	End If
	waitForIcallLoading
	setMinIdentificationPoints=strSetMinIdentificationPoints
End Function

'[Verify Field Minimum Identification Points displayed as]
Public Function verifyMinIdentificationPoints(strMinIdentificationPoints)
   bDevPending=false
   bverifyMinIdentificationPoints=true
   If Not IsNull(strMinIdentificationPoints) Then
     If Not verifyFieldValue (UserProfileAdmin.txtMinIdentificationPoints(), strMinIdentificationPoints, "MinIdentificationPoints")Then
	   bverifyMinIdentificationPoints=false
	End If
   End If
   verifyMinIdentificationPoints=bverifyMinIdentificationPoints
End Function

'[Set Minimum Authentication Points as]
Public Function SetMinimumAuthenticationPoints(strMinimumAuthenticationPoints)
	strSetMinimumAuthenticationPoints=true
	If Not IsNull(strMinimumAuthenticationPoints) Then
		UserProfileAdmin.txtMinAuthenticationPoints().set strMinimumAuthenticationPoints
	End If
	waitForIcallLoading
	SetMinimumAuthenticationPoints=strSetMinimumAuthenticationPoints
End Function

'[Verify Field Minimum Authentication Points displayed as]
Public Function verifyMinimumAuthenticationPoints(strMinimumAuthenticationPoints)
   bDevPending=false
   bverifyMinimumAuthenticationPoints=true
   If Not IsNull(strMinimumAuthenticationPoints) Then
     If Not verifyFieldValue (UserProfileAdmin.txtMinAuthenticationPoints(), strMinimumAuthenticationPoints, "MinimumAuthenticationPoints")Then
	   bverifyMinimumAuthenticationPoints=false
	End If
   End If
   verifyMinimumAuthenticationPoints=bverifyMinimumAuthenticationPoints
End Function

'[Select TPIN Indicator Combobox as]
Public Function selectTPINIndicatorCombobox(strTPINIndicator)
	strSelectTPINIndicatorCombobox=true
	If Not IsNull(strTPINIndicator) Then
		UserProfileAdmin.txtTPINIndicator().set strTPINIndicator
	End If
	waitForIcallLoading
	selectTPINIndicatorCombobox=strSelectTPINIndicatorCombobox
End Function

'[Select New SR Related Status Combobox as]
'Public Function selectNewSRRelatedStatusCombobox(strlblStatus)
'	strselectNewSRRelatedStatusCombobox=true
'	If Not IsNull(strlblStatus) Then
'		UserProfileAdmin.lblStatus().set strlblStatus
'	End If
'	waitForIcallLoading
'	selectNewSRRelatedStatusCombobox=strselectNewSRRelatedStatusCombobox
'End Function
'


'[Select New SR Related Route To Combobox as]
Public Function selectNewSRRelatedRouteToCombobox(strlstRouteTo)
	strselectNewSRRelatedRouteToCombobox=true
	If Not IsNull(strlstRouteTo) Then
		UserProfileAdmin.txtRouteTo().set strlstRouteTo
	End If
	waitForIcallLoading
	selectNewSRRelatedRouteToCombobox=strselectNewSRRelatedRouteToCombobox
End Function

'[Set SR Description as]
Public Function SelectSRDescription(strtxtSRDescription)
	bSelectSRDescription=true
	If Not IsNull(strtxtSRDescription) Then
		UserProfileAdmin.txtSRDescription().set strtxtSRDescription
	End If
	waitForIcallLoading
	SelectSRDescription=bSelectSRDescription
End Function

'[Verify Field C3 SR Description displayed as]
Public Function verifySRDescription(strtxtSRDescription)
   bDevPending=false
   bverifySRDescription=true
   If Not IsNull(strtxtSRDescription) Then
     If Not verifyFieldValue (UserProfileAdmin.txtSRDescription(), strtxtSRDescription, "txtSRDescription")Then
	   bverifySRDescription=false
	End If
   End If
   verifySRDescription=bverifySRDescription
End Function

''[Select Available Channels WebList as]
'Public Function selectAvailableChannelsWebList(strwlstAvailableChannels)
'	strselectAvailableChannelsWebList=true
'	If Not IsNull(strwlstAvailableChannels) Then
'		UserProfileAdmin.wlstAvailableChannels().Select strwlstAvailableChannels
'		strselectAvailableChannelsWebList = clickbtnMoveOneLeft()		
'	End If
'	waitForIcallLoading
'	selectAvailableChannelsWebList=strselectAvailableChannelsWebList
'End Function
'
''[Click On Move One Left Button in WebList]
'Public Function clickbtnMoveOneLeft()
'	bDevPending=false
'	UserProfileAdmin.btnMoveOneLeft.click
'	If Err.Number<>0 Then
'		clickbtnMoveOneLeft=false
'		LogMessage "WARN","Verification","Failed to Click Button : btnMoveOneLeft" ,false
'		Exit Function
'	End If
'	clickbtnMoveOneLeft=true
'End Function
'
'[Select IA Related To Combobox as]
Public Function selectIARelatedToCombobox(strIARelatedTo)
	strselectIARelatedToCombobox=true
	If Not IsNull(strIARelatedTo) Then
		UserProfileAdmin.txtIARelatedTo().set strIARelatedTo
		'UserProfileAdmin.txtIARelatedTo().set strIARelatedTo
	End If
	waitForIcallLoading
	selectIARelatedToCombobox=strselectIARelatedToCombobox
End Function

'[Set New IA Related To as]
Public Function SetNewIARelated(strtxtNewIARelated)
	strSetNewIARelated=true
	If Not IsNull(strtxtNewIARelated) Then
		UserProfileAdmin.txtNewIARelatedTo().set strtxtNewIARelated
	End If
	waitForIcallLoading
	SetNewIARelated=strSetNewIARelated
End Function

'[Select IA Type as]
Public Function SelectIAtype(strtxtIAType)
	strSelectIAtype=true
	If Not IsNull(strtxtIAType) Then
		UserProfileAdmin.txtIAType().set strtxtIAType
	End If
	waitForIcallLoading
	SelectIAtype=strSelectIAtype
End Function

'[Set IA New Type as]
Public Function SetIANewType(strIANewType)
	strSetIANewType=true
	If Not IsNull(strIANewType) Then
		UserProfileAdmin.txtIANewType().set strIANewType
	End If
	waitForIcallLoading
	SetIANewType=strSetIANewType
End Function

'[Set IA New Sub Type as]
Public Function setIANewSubType(strIANewSubType)
	strSetIANewSubType = true
	If Not IsNull(strIANewSubType) Then
		UserProfileAdmin.txtIANewSubType().set strIANewSubType
	End If
	waitForIcallLoading	
	setIANewSubType = strSetIANewSubType
End Function

'[Set Filter IA Related To as]
Public Function setFilterIARelatedTo(strtxtFilterIARelatedTO)
	strsetFilterIARelatedTo = true
	waitForIcallLoading
	If Not IsNull(strtxtFilterIARelatedTO) Then
		UserProfileAdmin.txtFilterIARelatedTO().set strtxtFilterIARelatedTO
	End If
	waitForIcallLoading	
	setFilterIARelatedTo = strsetFilterIARelatedTo
End Function

'[Select Filter IA Type as]
Public Function SelectFilterIAType(strtxtFilterIAType)
	strSelectFilterIAType=true
	waitForIcallLoading
	If Not IsNull(strtxtFilterIAType) Then
		UserProfileAdmin.txtFilterIAType().set strtxtFilterIAType
		UserProfileAdmin.txtFilterIAType().set strtxtFilterIAType
	End If
	waitForIcallLoading
	SelectFilterIAType=strSelectFilterIAType
End Function

'[Select Filter IA Status as]
Public Function SelectFilterIAStatus(strIASts)
	strSelectFilterIAStatus=true
	If Not IsNull(strIASts) Then
		UserProfileAdmin.txtIAStatus().set strIASts
	End If
	waitForIcallLoading
	SelectFilterIAStatus=strSelectFilterIAStatus
End Function

'[Click IA Related To View Link]
Public Function lnkViewIARelatedTo(strIASubType, strUserName, strDBServer)
Dim blnkViewIARelatedTo: blnkViewIARelatedTo = True
strQuery="select CREATED_DATETIME from cca_sr_ia_wrk where REQUEST_SUB_TYPE='"&strIASubType&"' order by CREATED_DATETIME desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)   		
lstViewIARelatedTo = checknull("IA Sub Type:"&strIASubType&"|Created By:"&strUserName&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")

blnkViewIARelatedTo=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewIARelatedTo,"UserProfileAdmin" ,"Action", "View",false,NULL ,NULL ,NULL)
WaitForICallLoading
lnkViewIARelatedTo=blnkViewIARelatedTo
End Function

'[Click SR Related To View Link]
Public Function lnkViewSRRelatedTo(strSRSubType, strUserName, strDBServer)
Dim blnkViewSRRelatedTo: blnkViewSRRelatedTo = True
strQuery="select CREATED_DATETIME from cca_sr_ia_wrk where REQUEST_SUB_TYPE='"&strSRSubType&"' order by CREATED_DATETIME desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)   		
lstViewSRRelatedTo = checknull("SR Sub Type:"&strSRSubType&"|Created By:"&strUserName&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")

blnkViewSRRelatedTo=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewSRRelatedTo,"UserProfileAdmin" ,"Action", "View",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkViewSRRelatedTo=blnkViewSRRelatedTo
End Function

'[Click IA Related To Modify Link]
Public Function lnkModifyIARelatedTo(strIASubType, strUserName, strDBServer)
Dim blnkModifyIARelatedTo: blnkModifyIARelatedTo = True
strQuery="select CREATED_DATETIME from cca_sr_ia_wrk where REQUEST_SUB_TYPE='"&strIASubType&"' order by CREATED_DATETIME desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)   		
lstModifyIARelatedTo = checknull("IA Sub Type:"&strIASubType&"|Created By:"&strUserName&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")

blnkModifyIARelatedTo=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyIARelatedTo,"UserProfileAdmin" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkModifyIARelatedTo=blnkModifyIARelatedTo
End Function

'[Click SR Related To Modify Link]
Public Function lnkModifySRRelatedTo(strSRSubType, strUserName, strDBServer)
Dim blnkModifySRRelatedTo: blnkModifySRRelatedTo = True
strQuery="select CREATED_DATETIME from cca_sr_ia_wrk where REQUEST_SUB_TYPE='"&strSRSubType&"' order by CREATED_DATETIME desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)   		
lstModifySRRelatedTo = checknull("SR Sub Type:"&strSRSubType&"|Created By:"&strUserName&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")

blnkModifySRRelatedTo=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifySRRelatedTo,"UserProfileAdmin" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkModifySRRelatedTo=blnkModifySRRelatedTo
End Function

'[Verify Field IA Related To displayed as]
Public Function verifyIARelatedTo(strFIARelatedTo)
bDevPending=false
bverifyIARelatedTo=true
If Not IsNull(strFIARelatedTo) Then
	If Not verifyFieldValue (UserProfileAdmin.txtVrfIARelated(), strFIARelatedTo, "IA Related To")Then
		bverifyIARelatedTo=false
	End If
End If
verifyIARelatedTo=bverifyIARelatedTo
End Function

'[Verify Field IA Status displayed as]
Public Function verifyIASts(strIASts)
	bDevPending=false
 	bverifyIASts=true
	If Not IsNull(strIASts) Then
	If Not verifyFieldValue (UserProfileAdmin.txtIAStatus(), strIASts, "IA Status")Then
		bverifyIASts=false
	End If
	End If
	verifyIASts=bverifyIASts
End Function


'[Verify Field IA Type displayed as]
Public Function verifyIAType(strverifyIAType)
bDevPending=false
bverifyIAType=true
If Not IsNull(strverifyIAType) Then
	If Not verifyFieldValue (UserProfileAdmin.txtVrfIAType(), strverifyIAType, "IA Type")Then
		bverifyIAType=false
	End If
End If
verifyIAType=bverifyIAType
End Function

'[Verify Field IA New Sub Type displayed as]
Public Function verifyIANewSubType(stryIANewSubType)
bDevPending=false
bverifyIANewSubType=true
If Not IsNull(stryIANewSubType) Then
	If Not verifyFieldValue (UserProfileAdmin.txtVrfNewSubType(), stryIANewSubType, "New Sub Type")Then
		bverifyIANewSubType=false
	End If
End If
verifyIANewSubType=bverifyIANewSubType
End Function

'*****************Added by Kalyan for User Role dated 9112016******************

'[Set UserRole textbox as]
Public Function selectUsrRoleTxtBox(strUsrRole)
bselectUsrRoleTxtBox=true
If Not IsNull(strUsrRole) Then
UserProfileAdmin.txtUserRole().set strUsrRole
End If
waitForIcallLoading
selectUsrRoleTxtBox=bselectUsrRoleTxtBox
End Function

'[Set UserRole Desc textbox as]
Public Function selectUsrRoleDesc(strUsrRoleDesc)
bselectUsrRoleDesc=true
If Not IsNull(strUsrRoleDesc) Then
UserProfileAdmin.txtRoleDesc().set strUsrRoleDesc
End If
waitForIcallLoading
selectUsrRoleDesc=bselectUsrRoleDesc
End Function

'[Set UserRole LandingPage combobox as]
Public Function selectUsrRoleLandingPage(strUsrRoleLandPge)
bselectUsrRoleLandingPage=true
If Not IsNull(strUsrRoleLandPge) Then
UserProfileAdmin.lstLandingPage().set strUsrRoleLandPge
End If
waitForIcallLoading
selectUsrRoleLandingPage=bselectUsrRoleLandingPage
End Function

'[Set UserRole Status combobox as]
Public Function selectUsrRoleStatus(strUsrRoleSts)
bselectUsrRoleStatus=true
If Not IsNull(strUsrRoleSts) Then
UserProfileAdmin.lstStatus().set strUsrRoleSts
UserProfileAdmin.lstStatus().FireEvent "onclick"
End If
waitForIcallLoading
selectUsrRoleStatus=bselectUsrRoleStatus
End Function

'[Set UserRole type combobox as]
Public Function selectUsrRoleType(strUsrRoleType)
bselectUsrRoleType=true
If Not IsNull(strUsrRoleType) Then
UserProfileAdmin.lstRoleType().set strUsrRoleType
End If
waitForIcallLoading
selectUsrRoleType=bselectUsrRoleType
End Function

'[Set UserRole Staff combobox as]
Public Function selectUsrRoleStaff(strUsrRoleStaff)
bselectUsrRoleStaff=true
If Not IsNull(strUsrRoleStaff) Then
UserProfileAdmin.txtStaffAccess().set strUsrRoleStaff
End If
waitForIcallLoading
selectUsrRoleStaff=bselectUsrRoleStaff
End Function

'[Set UserRole left select as]
Public Function selectUsrRoleLftSlct(strUsrRoleLftSlct)
bselectUsrRoleLftSlct=true
If Not IsNull(strUsrRoleLftSlct) Then
UserProfileAdmin.lstUsrRoleChnlTypeLeftSlct().set strUsrRoleLftSlct
End If
waitForIcallLoading
selectUsrRoleLftSlct=bselectUsrRoleLftSlct
End Function

'[Set UserRole Right select as]
Public Function selectUsrRoleRgtSlct(strUsrRoleRgtSlct)
bselectUsrRoleRgtSlct=true
If Not IsNull(strUsrRoleRgtSlct) Then
UserProfileAdmin.lstUsrRoleChnlTypeRightSlct().set strUsrRoleRgtSlct
End If
waitForIcallLoading
selectUsrRoleRgtSlct=bselectUsrRoleRgtSlct
End Function

''[Verify Channel listbox  is enabled]
'Public Function verifyChnlLstBox_Enable()
'	bverifyChnlLstBox_Enable=true	
'	'intbtnNewAddIA =UserProfileAdmin.wlstAvailableChannels().GetROProperty("disabled")
'	intbtnNewAddIA=Instr(UserProfileAdmin.wlstAvailableChannels.GetROproperty("outerhtml"),("disabled-area"))
'	If intbtnNewAddIA=0 Then
'		LogMessage "RSLT","Verification","NewIA button is enabled as expected.",True
'		bverifyChnlLstBox_Enable=true
'	Else
'		LogMessage "WARN","Verifiation","NewIA button is disabled. Expected enabled.",false
'		bverifyChnlLstBox_Enable=false
'	End If
'	verifyChnlLstBox_Enable=bverifyChnlLstBox_Enable
'End Function
'

'[Verify Field User Role lbl Channel displayed as]
Public Function verifyUserRolelblChnl(strUsrRleChnl)
bDevPending=false
bverifyUserRolelblChnl=true
If Not IsNull(strUsrRleChnl) Then
If Not verifyFieldValue (UserProfileAdmin.lblChannel(), strUsrRleChnl, "UserRoleChannel")Then
bverifyUserRolelblChnl=false
End If
End If
verifyUserRolelblChnl=bverifyUserRolelblChnl
End Function

'[Verify Field User Role lbl displayed as]
Public Function verifyUserRolelbl(strUsrRle)
bDevPending=false
bverifyUserRolelbl=true
If Not IsNull(strUsrRle) Then
If Not verifyFieldValue (UserProfileAdmin.lblUserRole(), strUsrRle, "UserRole")Then
bverifyUserRolelbl=false
End If
End If
verifyUserRolelbl=bverifyUserRolelbl
End Function

'[Verify Field User Role lbl desc displayed as]
Public Function verifyUserRolelblDesc(strUsrRleDesc)
bDevPending=false
bverifyUserRolelblDesc=true
If Not IsNull(strUsrRleDesc) Then
If Not verifyFieldValue (UserProfileAdmin.lblRoleDesc(), strUsrRleDesc, "UserRoleDesc")Then
bverifyUserRolelblDesc=false
End If
End If
verifyUserRolelblDesc=bverifyUserRolelblDesc
End Function

'[Verify Field User Role Landing Page displayed as]
Public Function verifyUserRolelblLandgPge(strUsrRleLndgPge)
bDevPending=false
bverifyUserRolelblLandgPge=true
If Not IsNull(strUsrRleLndgPge) Then
If Not verifyFieldValue (UserProfileAdmin.lblLandingPage(), strUsrRleLndgPge, "UserRoleLndgPge")Then
bverifyUserRolelblLandgPge=false
End If
End If
verifyUserRolelblLandgPge=bverifyUserRolelblLandgPge
End Function

'[Verify Field User Role lbl Status displayed as]
Public Function verifyUserRolelblSts(strUsrRleSts)
bDevPending=false
bverifyUserRolelblSts=true
If Not IsNull(strUsrRleSts) Then
If Not verifyFieldValue (UserProfileAdmin.lblStatus1(), strUsrRleSts, "UserRoleSts")Then
bverifyUserRolelblSts=false
End If
End If
verifyUserRolelblSts=bverifyUserRolelblSts
End Function

'[Verify Field User Role Type displayed as]
Public Function verifyUserRolelblType(strUsrRleType)
bDevPending=false
bverifyUserRolelblType=true
If Not IsNull(strUsrRleType) Then
If Not verifyFieldValue (UserProfileAdmin.lblRoleType(), strUsrRleType, "UserRoleType")Then
bverifyUserRolelblType=false
End If
End If
verifyUserRolelblType=bverifyUserRolelblType
End Function

'[Verify Field User Role Staff displayed as]
Public Function verifyUserRolelblStaff(strUsrRleStaff)
bDevPending=false
bverifyUserRolelblStaff=true
If Not IsNull(strUsrRleStaff) Then
If Not verifyFieldValue (UserProfileAdmin.lblStaffAccess(), strUsrRleStaff, "UserRoleStaff")Then
bverifyUserRolelblStaff=false
End If
End If
verifyUserRolelblStaff=bverifyUserRolelblStaff
End Function

'[Verify User Role validation message displayed as]
Public Function verifyUserRoleValidation(strUsrRleMsg)
bDevPending=false
bverifyUserRoleValidation=true
If Not IsNull(strUsrRleMsg) Then
If Not verifyInnerText (UserProfileAdmin.lblValidationMsg(), strUsrRleMsg, "UserRoleValidationMsg")Then
bverifyUserRoleValidation=false
End If
End If
verifyUserRoleValidation=bverifyUserRoleValidation
End Function

'[Click User Role View Link]
Public Function lnkViewUserRole(strUsrRle, strDesc, strtCreatedBy, strDBServer)
Dim blnkViewUserRole: lnkViewUserRole = True
strQuery="select created_date from ISERVE_USER_ROLE_WRK where USER_ROLE='"&strUsrRle&"' order by CREATED_DATE desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)   		
lstViewUserProfile = checknull("User Role:"&strUsrRle&"|Role Description:"&strDesc&"|Created By:"&strtCreatedBy&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
'lstViewUserProfile = checknull("User Role:"&strUsrRle&"|Description:"&strDesc&"|Created By:"&strtCreatedBy&"")
blnkViewUserRole=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewUserProfile,"UserRoleAdmin" ,"Action", "View",True,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkViewUserRole=blnkViewUserRole
End Function

'[Click User Role Modify Link]
Public Function lnkModifyUserRole(strUsrRle, strDesc, strtCreatedBy, strDBServer)
Dim blnkModifyUserRole: lnkModifyUserRole = True
'strQuery="select created_date from ISERVE_USER_ROLE_WRK where USER_ROLE='"&strUsrRle&"'"
strQuery="select created_date from ISERVE_USER_ROLE where USER_ROLE='"&strUsrRle&"' order by CREATED_DATE desc LIMIT 1"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
lstModifyUserProfile = checknull("User Role:"&strUsrRle&"|Role Description:"&strDesc&"|Created By:"&strtCreatedBy&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
'lstModifyUserProfile = checknull("User Role:"&strUsrRle&"|Description:"&strDesc&"|Created By:"&strtCreatedBy&"")
blnkModifyUserRole=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyUserProfile,"UserRoleAdmin" ,"Action", "Modify",True,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkModifyUserRole=blnkModifyUserRole
End Function


'********************************End*******************************************

'*****************Added by Kalyan for User Profile dated 9112016******************

'[Set UserProfile Channel as]
Public Function selectUsrProfileChnl(strUsrprofChnl)
bselectUsrProfileChnl=true
If Not IsNull(strUsrprofChnl) Then
UserProfileAdmin.lstUsrProfChnl().set strUsrprofChnl
End If
waitForIcallLoading
selectUsrProfileChnl=bselectUsrProfileChnl
End Function

'[Set UserProfile Role as]
Public Function selectUsrProfileRole(strUsrprofRole)
bselectUsrProfileRole=true
If Not IsNull(strUsrprofRole) Then
UserProfileAdmin.lstUsrProfRole().set strUsrprofRole
End If
waitForIcallLoading
selectUsrProfileRole=bselectUsrProfileRole
End Function

'[Set UserProfile Mngr1BankId as]
Public Function selectUsrProfileMgr1BnkId(strUsrprofMgrBnkId)
bselectUsrProfileMgr1BnkId=true
If Not IsNull(strUsrprofMgrBnkId) Then
UserProfileAdmin.txtUsrProfMngr1BnkId().set strUsrprofMgrBnkId
End If
waitForIcallLoading
selectUsrProfileMgr1BnkId=bselectUsrProfileMgr1BnkId
End Function

'[Set UserProfile RACF as]
Public Function selectUsrProfileRACF(strUsrprofRACF)
bselectUsrProfileRACF=true
If Not IsNull(strUsrprofRACF) Then
UserProfileAdmin.txtRACFID().set strUsrprofRACF
End If
waitForIcallLoading
selectUsrProfileRACF=bselectUsrProfileRACF
End Function
'[Set UserProfile Emp No as]
Public Function selectUsrProfileEmpNo(strUsrprofEmpNo)
bselectUsrProfileEmpNo=true
If Not IsNull(strUsrprofEmpNo) Then
UserProfileAdmin.txtEmpNo().set strUsrprofEmpNo
End If
waitForIcallLoading
selectUsrProfileEmpNo=bselectUsrProfileEmpNo
End Function

'[Set UserProfile UserGroup as]
Public Function selectUsrProfileUsrGrp(strUsrprofUsrGrp)
bselectUsrProfileUsrGrp=true
If Not IsNull(strUsrprofUsrGrp) Then
UserProfileAdmin.lstUsrProfGroup().set strUsrprofUsrGrp
End If
waitForIcallLoading
selectUsrProfileUsrGrp=bselectUsrProfileUsrGrp
End Function

'[Set UserProfile Location as]
Public Function selectUsrProfileUsrLctn(strUsrprofUsrLctn)
bselectUsrProfileUsrLctn=true
If Not IsNull(strUsrprofUsrLctn) Then
UserProfileAdmin.lstUsrProfGroup().set strUsrprofUsrLctn
End If
waitForIcallLoading
selectUsrProfileUsrLctn=bselectUsrProfileUsrLctn
End Function

'[Set UserProfile Srch 1bankId as]
Public Function selectUsrProfileSrchBankId(strUsrprofSrchBnkId)
bselectUsrProfileSrchBankId=true
If Not IsNull(strUsrprofSrchBnkId) Then
UserProfileAdmin.txt1BankIDSearch().set strUsrprofSrchBnkId
End If
waitForIcallLoading
selectUsrProfileSrchBankId=bselectUsrProfileSrchBankId
End Function

'[Click Button Modify User Profile Screen]
Public Function clickButtonModify_UserProfile()
bDevPending=false
UserProfileAdmin.btnModify.click
	If Err.Number<>0 Then
	clickButtonModify_UserProfile=false
	LogMessage "WARN","Verification","Failed to Click Button : Modify",false 
	Exit Function
	End If
clickButtonModify_UserProfile=true
End Function

'[Click Button Search User Profile Screen]
Public Function clickButtonSearch_UserProfile()
bDevPending=false
UserProfileAdmin.btnSearch.click
	If Err.Number<>0 Then
	clickButtonSearch_UserProfile=false
	LogMessage "WARN","Verification","Failed to Click Button : Search",false 
	Exit Function
	End If
clickButtonSearch_UserProfile=true
End Function

'[Set UserProfile Func Map Role as]
Public Function selectUsrProfileFuncMapRole(strUsrprofFuncMap)
bselectUsrProfileFuncMapRole=true
If Not IsNull(strUsrprofFuncMap) Then
UserProfileAdmin.lstUsrFctnMapRole().set strUsrprofFuncMap
UserProfileAdmin.lstUsrFctnMapRole().set strUsrprofFuncMap
End If
waitForIcallLoading
selectUsrProfileFuncMapRole=bselectUsrProfileFuncMapRole
End Function

'[Set UserProfile Func Map Group as]
Public Function selectUsrProfileFuncMapGrp(strUsrprofFuncMapGrp)
bselectUsrProfileFuncMapGrp=true
If Not IsNull(strUsrprofFuncMapGrp) Then
UserProfileAdmin.lstUsrFctnMapUsrGrp().set strUsrprofFuncMapGrp
UserProfileAdmin.lstUsrFctnMapUsrGrp().set strUsrprofFuncMapGrp
End If
waitForIcallLoading
selectUsrProfileFuncMapGrp=bselectUsrProfileFuncMapGrp
End Function

'[Set UserProfile Func Map Type as]
Public Function selectUsrProfileFuncMap(strUsrprofFucnMap)
bselectUsrProfileFuncMap=true
If Not IsNull(strUsrprofFucnMap) Then
UserProfileAdmin.lstUsrFctnMapFucnType().set strUsrprofFucnMap
UserProfileAdmin.lstUsrFctnMapFucnType().set strUsrprofFucnMap
End If
waitForIcallLoading
selectUsrProfileFuncMap=bselectUsrProfileFuncMap
End Function

'[Click Add/Modify hyperlink on Configure page]
Public Function ClickAddModifyUserProfile(strName, str1BankId)
Dim bClickAddModifyUserProfile: ClickAddModifyUserProfile = True

lstAddModify = checknull("Name:"&strName&"|1Bank ID:"&str1BankId&"")

bClickAddModifyUserProfile=selectTableLink(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstAddModify,"UserProfileAdmin" ,"Action",false,NULL ,NULL ,NULL)
WaitForICallLoading
ClickAddModifyUserProfile=bClickAddModifyUserProfile
End Function

'[Select Available Channels WebList as]
Public Function selectAvailableChannelsWebList(strwlstAvailableChannels)
strselectAvailableChannelsWebList=true
If Not IsNull(strwlstAvailableChannels) Then
UserProfileAdmin.wlstAvailableChannels().Select strwlstAvailableChannels
strselectAvailableChannelsWebList = clickbtnMoveOneLeft() 
End If
waitForIcallLoading
selectAvailableChannelsWebList=strselectAvailableChannelsWebList
End Function

'[Select Available Segments WebList as]
Public Function selectAvailableSegmentsWebList(strwlstAvailableSegmnts)
strselectAvailableSegmentsWebList=true
If Not IsNull(strwlstAvailableSegmnts) Then
UserProfileAdmin.wlstAvailableSegments().Select strwlstAvailableSegmnts
strselectAvailableSegmentsWebList = clickbtnMoveSegOneLeft() 
End If
waitForIcallLoading
selectAvailableSegmentsWebList=strselectAvailableSegmentsWebList
End Function

'[Select Available Channels SRIA WebList as]
Public Function selectAvailableChannelsSRIAWebList(strwlstAvailableChannels)
strselectAvailableChannelsSRIAWebList=true
If Not IsNull(strwlstAvailableChannels) Then
UserProfileAdmin.wlstAvailableChannelsIASR().Select strwlstAvailableChannels
strselectAvailableChannelsSRIAWebList = clickbtnMoveOneLeft() 
End If
waitForIcallLoading
selectAvailableChannelsSRIAWebList=strselectAvailableChannelsSRIAWebList
End Function

'[Click On Move One Left Button in WebList]
Public Function clickbtnMoveOneLeft()
bDevPending=false
UserProfileAdmin.btnMoveOneLeft.click
If Err.Number<>0 Then
clickbtnMoveOneLeft=false
LogMessage "WARN","Verification","Failed to Click Button : btnMoveOneLeft" ,false
Exit Function
End If
clickbtnMoveOneLeft=true
End Function

'[Click On Move One Left Button in Segment WebList]
Public Function clickbtnMoveSegOneLeft()
bDevPending=false
UserProfileAdmin.btnMoveSegOneLeft.click
If Err.Number<>0 Then
clickbtnMoveSegOneLeft=false
LogMessage "WARN","Verification","Failed to Click Button : btnMoveSegOneLeft" ,false
Exit Function
End If
clickbtnMoveSegOneLeft=true
End Function

'[Select Available Mapped WebList as]
Public Function selectAvailableMappedsWebList(strwlstMappedChannels)
strselectAvailableMappedsWebList=true
If Not IsNull(strwlstMappedChannels) Then
UserProfileAdmin.wlstMappedChannel().Select strwlstMappedChannels
strselectAvailableMappedsWebList = clickbtnMoveOneRight() 
End If
waitForIcallLoading
selectAvailableMappedsWebList=strselectAvailableMappedsWebList
End Function

'[Select Available Mapped Segment WebList as]
Public Function selectAvailableMappedSegWebList(strwlstMappedSeg)
strselectAvailableMappedSegWebList=true
If Not IsNull(strwlstMappedSeg) Then
UserProfileAdmin.wlstMappedSegments().Select strwlstMappedSeg
strselectAvailableMappedSegWebList = btnMoveSegOneRight()
End If
waitForIcallLoading
selectAvailableMappedSegWebList=strselectAvailableMappedSegWebList
End Function

'[Verify Channel WebList  is enabled]
Public Function verifyChnlWebLst_Enable(strChkVal)
	bverifyChnlWebLst_Enable=true	
	If strChkVal="false" Then
		Exit Function
	End If
	intChnlWebLst =InStr(UserProfileAdmin.wlstAvailableChannels.GetROProperty("disabled"),0)
	If not intbtnAddIA=0 Then
		LogMessage "RSLT","Verification","Channel Weblist is enabled as expected.",True
		bverifyChnlWebLst_Enable=true
	Else
		LogMessage "WARN","Verifiation","Channel Weblist is disabled. Expected enabled.",false
		bverifyChnlWebLst_Enable=false
	End If
	verifyChnlWebLst_Enable=bverifyChnlWebLst_Enable
End Function

'[Verify Channel WebList  is disabled]
Public Function verifyChnlWebLst_Disable(strChkVal)
	bverifyChnlWebLst_Disable=true
	If strChkVal="false" Then
		Exit Function
	End If	
	intChnlWebLst =InStr(UserProfileAdmin.wlstAvailableChannels.GetROProperty("disabled"),0)
	If intbtnAddIA=0 Then
		LogMessage "RSLT","Verification","Channel Weblist is disabled as expected.",True
		bverifyChnlWebLst_Disable=true
	Else
		LogMessage "WARN","Verifiation","Channel Weblist is enabled. Expected disabled.",false
		bverifyChnlWebLst_Disable=false
	End If
	verifyChnlWebLst_Disable=bverifyChnlWebLst_Disable
End Function

'[Click On Move One Right Button in WebList]
Public Function clickbtnMoveOneRight()
bDevPending=false
UserProfileAdmin.btnMoveOneRight.click
If Err.Number<>0 Then
clickbtnMoveOneRight=false
LogMessage "WARN","Verification","Failed to Click Button : btnMoveOneRight" ,false
Exit Function
End If
clickbtnMoveOneRight=true
End Function

'[Click On Move One Right Button in Segment WebList]
Public Function clickbtnMoveSegOneRight()
bDevPending=false
UserProfileAdmin.btnMoveSegOneRight.click
If Err.Number<>0 Then
clickbtnMoveSegOneRight=false
LogMessage "WARN","Verification","Failed to Click Button : btnMoveSegOneRight" ,false
Exit Function
End If
clickbtnMoveSegOneRight=true
End Function

'[Verify User Profile Modify link disabled]
Public Function verifyModifyLink_Disabled(strUsrRle,strDesc,strtCreatedBy,strDBServer)
bverifyModifyLink_Disabled=true
strQuery="select created_date from ISERVE_USER_ROLE_WRK where USER_ROLE='"&strUsrRle&"'"
strtCreateDateTime=getDBValForColumn_MARIADB_FE(strQuery, strDBServer)
getDateMnthYear(strtCreateDateTime(0))
'getDateMnthYear(strtCreateDateTime)
lstModifyUserProfile = checknull("User Role:"&strUsrRle&"|Description:"&strDesc&"|Created By:"&strtCreatedBy&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
bverifyModifyLink_Disabled=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyUserProfile,"UserRoleAdmin" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
verifyModifyLink_Disabled=bverifyModifyLink_Disabled
End Function

'[Execute DB Query to delete User Role Data From DB]
Public Function executeUsrRole_DBQuery(strUsrRle)
bexecuteUsrRole_DBQuery=true
strDeleteQuery1 = "Delete from CCA_ROLE_LANDING where USER_ROLE='"&strUsrRle&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

strDeleteQuery2 = "Delete from CCA_USER_ACCESS where USER_ROLE='"&strUsrRle&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery2)

strDeleteQuery8="Delete from user_role_mapped_segments where role_id in (select rec_id from iserve_user_role where USER_ROLE='"&strUsrRle&"')"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery8)

strDeleteQuery9="Delete from user_role_mapped_channel where USER_ROLE='"&strUsrRle&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery9)

'strDeleteQuery3 = "Delete from ISERVE_USER_ROLE_TABLE where USER_ROLE='"&strUsrRle&"'"
strDeleteQuery3 = "Delete from iserve_user_role where USER_ROLE='"&strUsrRle&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery3)

strDeleteQuery4 = "Delete from ISERVE_USER_ROLE_WRK where USER_ROLE='"&strUsrRle&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery4)

strDeleteQuery5 = "Delete from CCA_USER_VIEW_CONFIG where VIEW_NAME='"&strUsrRle&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery5)

strDeleteQuery6 = "Delete from CCA_PARAM_AUDIT_VALUE where AUDIT_ID in (select AUDIT_ID from CCA_PARAM_AUDIT_LOG where REFERENCE_ID='"&strRoleRefId&"')"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery6)

strDeleteQuery7 = "Delete from CCA_PARAM_AUDIT_LOG where REFERENCE_ID='"&strRoleRefId&"'"
bexecuteUsrRole_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery7)
WaitForICallLoading
executeUsrRole_DBQuery=bexecuteUsrRole_DBQuery
End Function

'[Execute DB Query to delete User Group Data From DB]
Public Function executeUsrGrp_DBQuery(strUsrGrp)
bexecuteUsrGrp_DBQuery=true
strDeleteQuery1 = "Delete from iserve_user_groups where GROUP_ID='"&strUsrGrp&"'"
bexecuteUsrGrp_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

strDeleteQuery2 = "Delete from iserve_user_groups_wrk where GROUP_ID='"&strUsrGrp&"'"
bexecuteUsrGrp_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery2)

strDeleteQuery6 = "Delete from CCA_PARAM_AUDIT_VALUE where AUDIT_ID in (select AUDIT_ID from CCA_PARAM_AUDIT_LOG where REFERENCE_ID='"&strGrpRefId&"')"
bexecuteUsrGrp_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery6)

strDeleteQuery7 = "Delete from CCA_PARAM_AUDIT_LOG where REFERENCE_ID='"&strGrpRefId&"'"
bexecuteUsrGrp_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery7)
WaitForICallLoading
executeUsrGrp_DBQuery=bexecuteUsrGrp_DBQuery
End Function

'[Execute DB Query to delete Functional Mapping Data From DB]
Public Function executeFuncMap_DBQuery(strUsrRole, strUsrGrp)
bexecuteFuncMap_DBQuery=true
strDeleteQuery1 = "Delete from iserve_function_access where ACCESS_GROUP='"&strUsrRole&"' AND USER_GRP='"&strUsrGrp&"'"
bexecuteFuncMap_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

strDeleteQuery2 = "Delete from iserve_function_access_wrk where ACCESS_GROUP='"&strUsrRole&"' AND USER_GRP='"&strUsrGrp&"'"
bexecuteFuncMap_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery2)

strDeleteQuery6 = "Delete from CCA_PARAM_AUDIT_VALUE where AUDIT_ID in (select AUDIT_ID from CCA_PARAM_AUDIT_LOG where REFERENCE_ID='"&strFunMapRefId&"')"
bexecuteFuncMap_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery6)

strDeleteQuery7 = "Delete from CCA_PARAM_AUDIT_LOG where REFERENCE_ID='"&strFunMapRefId&"'"
bexecuteFuncMap_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery7)

strDeleteQuery8 = "Delete from user_role_mapped_channel where USER_ROLE='"&strUsrRole&"'"
bexecuteFuncMap_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery8)
WaitForICallLoading
executeFuncMap_DBQuery=bexecuteFuncMap_DBQuery
End Function

'[Execute DB Query to delete User Profile Data From DB]
Public Function executeUsrProf_DBQuery(strUsrProfName)
bexecuteUsrProf_DBQuery=true
strDeleteQuery1 = "Delete from cca_user_profile where USERLANID='"&strUsrProfName&"'"
bexecuteUsrProf_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

WaitForICallLoading
executeUsrProf_DBQuery=bexecuteUsrProf_DBQuery
End Function

'[Execute DB Query to verify User Profile Audit Entry]
Public Function executeAuditEntry_DBQuery(strDBServer, arrExpectedData)
bexecuteAuditEntry_DBQuery=true
strQuery="select PARAMETER_TYPE, CHANGE_TYPE from cca_param_audit_log where REFERENCE_ID='"&strRefId&"' and CHANGE_TYPE='ADD' limit 1"
bexecuteAuditEntry_DBQuery=CompareDBValue_icall(strDBServer,strQuery, arrExpectedData)
WaitForICallLoading
executeAuditEntry_DBQuery=bexecuteAuditEntry_DBQuery
End Function

'[Execute DB Query to delete Configured IA From DB]
Public Function executeConfigIA_DBQuery(strIARelTo, strIAType, strIASubType)
bexecuteConfigIA_DBQuery=true
strDeleteQuery1 = "Delete from cca_sr_ia_wrk where RELATED_TO='"&strIARelTo&"'"
bexecuteConfigIA_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

strDeleteQuery2 = "Delete from cca_act_int_channel where IA_RELATED_TO='"&strIARelTo&"'"
bexecuteConfigIA_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery2)

strDeleteQuery3 = "Delete from cca_prm_act_subtype where sub_type='"&strIASubType&"'"
bexecuteConfigIA_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery3)

strDeleteQuery4 = "Delete from cca_prm_act_type where type='"&strIAType&"'"
bexecuteConfigIA_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery4)

strDeleteQuery5 = "Delete from cca_prm_act_relto where related_to='"&strIARelTo&"'"
bexecuteConfigIA_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery5)

WaitForICallLoading
executeConfigIA_DBQuery=bexecuteConfigIA_DBQuery
End Function

'[Execute DB Query to delete Configured SR From DB]
Public Function executeConfigSR_DBQuery(strSRRelTo, strSRType, strSRSubType)
bexecuteConfigSR_DBQuery=true
strDeleteQuery1 = "Delete from cca_sr_ia_wrk where RELATED_TO='"&strSRRelTo&"'"
bexecuteConfigSR_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

strDeleteQuery2 = "Delete from cca_prm_sr_icall_c3_intg where TYPE='"&strSRType&"'"
bexecuteConfigSR_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery2)

strDeleteQuery3 = "Delete from cca_prm_sr_subtype where Req_Sub_Type='"&strSRSubType&"'"
bexecuteConfigSR_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery3)

strDeleteQuery4 = "Delete from cca_prm_sr_type where Req_Type='"&strSRType&"'"
bexecuteConfigSR_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery4)

strDeleteQuery5 = "Delete from cca_prm_sr_relto where related_to='"&strSRRelTo&"'"
bexecuteConfigSR_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery5)

WaitForICallLoading
executeConfigSR_DBQuery=bexecuteConfigSR_DBQuery
End Function

'[Execute DB Query to verify SR Mapped Channels]
Public Function executeSRMapdChnls_DBQuery(strDBServer, arrExpectedData, strSRType)
bexecuteSRMapdChnls_DBQuery=true
strQuery="select CREATED_By, CHANNEL from cca_prm_sr_icall_c3_intg where TYPE='"&strSRType&"'"
bexecuteSRMapdChnls_DBQuery=CompareDBMultipleRowValues_icall(strDBServer,strQuery, arrExpectedData)
WaitForICallLoading
executeSRMapdChnls_DBQuery=bexecuteSRMapdChnls_DBQuery
End Function

'[Check Channel Type Enabled Disabled]
Public Function CheckChnlType_Disabled(strObjectStatus)
	Set oDesc=Description.Create
	oDesc("micclass").Value = "WebElement"
	oDesc("class").Value = "view-container.*"
	oDesc("outertext").Value = "Channel Type.*"
	Set objChnlLst= Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDesc)
	intBtnCount=objChnlLst.Count
	If intbtnCount=0 Then
	CheckChnlType_Disabled=false
	LogMessage "RSLT","Verification","Expected Channel Type Field does not displayed",false
	else
	bDisabled =matchStr(objChnlLst(0).GetROProperty("outerhtml"),"view-container layout-column flex-55 disabled-area")
	End If
	If Ucase(Trim(strObjectStatus)) = "ENABLED" Then
	If bDisabled Then
	'Fail
	LogMessage "WARN","Verification","Object is disabled. Expected should be enabled",False
	CheckChnlType_Disabled = False
	Else
	'Pass
	LogMessage "RSLT","Verification","Object is enabled as expected",True
	CheckChnlType_Disabled = True						
	End If
	ElseIf Ucase(Trim(strObjectStatus)) = "DISABLED" Then
	If bDisabled Then
	'Pass
	LogMessage "RSLT","Verification","Object is disabled as expected",True
	CheckChnlType_Disabled = True
	Else
	'Fail
	LogMessage "WARN","Verification","Object is enabled. Expected should be disabeld",False
	CheckChnlType_Disabled = False
	End If	
	End If
End Function

'[Check Segments List Enabled Disabled]
Public Function CheckSegments_Disabled(strObjectStatus)
	Set oDesc=Description.Create
	oDesc("micclass").Value = "WebElement"
	oDesc("class").Value = "view-container.*"
	oDesc("outertext").Value = "Segments >> > < << "
	Set objSegLst= Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDesc)
	intBtnCount=objSegLst.Count
	If intbtnCount=0 Then
	CheckSegments_Disabled=false
	LogMessage "RSLT","Verification","Expected Segment List does not displayed",false
	else
	bDisabled =matchStr(objSegLst(0).GetROProperty("outerhtml"),"view-container layout-column flex-55 disabled-area")
	End If
	If Ucase(Trim(strObjectStatus)) = "ENABLED" Then
	If bDisabled Then
	'Fail
	LogMessage "WARN","Verification","Object is disabled. Expected should be enabled",False
	CheckSegments_Disabled = False
	Else
	'Pass
	LogMessage "RSLT","Verification","Object is enabled as expected",True
	CheckSegments_Disabled = True						
	End If
	ElseIf Ucase(Trim(strObjectStatus)) = "DISABLED" Then
	If bDisabled Then
	'Pass
	LogMessage "RSLT","Verification","Object is disabled as expected",True
	CheckSegments_Disabled = True
	Else
	'Fail
	LogMessage "WARN","Verification","Object is enabled. Expected should be disabeld",False
	CheckSegments_Disabled = False
	End If	
	End If
End Function

'[Check Modify Button Function Mapping Enabled Disabled]
Public Function CheckFuncMapModify_Disabled(strObjectStatus)
	Set oDesc=Description.Create
	oDesc("micclass").Value = "WebElement"
	oDesc("class").Value = "view-container.*"
	oDesc("innertext").Value = "CancelModify"
	Set objFunMapLst= Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDesc)
	intBtnCount=objFunMapLst.Count
	If intbtnCount=0 Then
	CheckFuncMapModify_Disabled=false
	LogMessage "RSLT","Verification","Expected Function Map does not displayed",false
	else
	bDisabled =matchStr(objSegLst(0).GetROProperty("outerhtml"),"layout-align-end-center layout-row flex-60 ng-hide")
	End If
	If Ucase(Trim(strObjectStatus)) = "ENABLED" Then
	If bDisabled Then
	'Fail
	LogMessage "WARN","Verification","Object is disabled. Expected should be enabled",False
	CheckFuncMapModify_Disabled = False
	Else
	'Pass
	LogMessage "RSLT","Verification","Object is enabled as expected",True
	CheckFuncMapModify_Disabled = True						
	End If
	ElseIf Ucase(Trim(strObjectStatus)) = "DISABLED" Then
	If bDisabled Then
	'Pass
	LogMessage "RSLT","Verification","Object is disabled as expected",True
	CheckFuncMapModify_Disabled = True
	Else
	'Fail
	LogMessage "WARN","Verification","Object is enabled. Expected should be disabeld",False
	CheckSegments_Disabled = False
	End If	
	End If
End Function

'[Verify Func Map label Msg displayed as]
Public Function verifyFuncMapLblMsg(strFuncMapLblMsg)
bDevPending=false
bverifyFuncMapLblMsg=true
If Not IsNull(strFuncMapLblMsg) Then
If Not verifyFieldValue (UserProfileAdmin.lblFuncMapMsg(), strFuncMapLblMsg, "FuncMapLblMsg")Then
bverifyFuncMapLblMsg=false
End If
End If
verifyFuncMapLblMsg=bverifyFuncMapLblMsg
End Function

'[Verify Func Master Accnt Access Flag displayed as]
Public Function verifyFuncAccsFlag(strFuncAcss)
bDevPending=false
bverifyFuncAccsFlag=true
If Not IsNull(strFuncAcss) Then
If Not verifyFieldValue (UserProfileAdmin.txtAccntAccsFlg(), strFuncAcss, "FuncAccessFlag")Then
bverifyFuncAccsFlag=false
End If
End If
verifyFuncAccsFlag=bverifyFuncAccsFlag
End Function

'[Verify Func Master Staff Access Flag displayed as]
Public Function verifyFuncStaffAccsFlag(strFuncStaff)
bDevPending=false
bverifyFuncStaffAccsFlag=true
If Not IsNull(strFuncStaff) Then
If Not verifyFieldValue (UserProfileAdmin.txtStaffAccsFlg(), strFuncStaff, "FuncStaffAccessFlag")Then
bverifyFuncStaffAccsFlag=false
End If
End If
verifyFuncStaffAccsFlag=bverifyFuncStaffAccsFlag
End Function

'[Verify Func Master Segment Access Flag displayed as]
Public Function verifyFuncSegmntAccsFlag(strFuncSegmnt)
bDevPending=false
bverifyFuncSegmntAccsFlag=true
If Not IsNull(strFuncSegmnt) Then
If Not verifyFieldValue (UserProfileAdmin.txtSegmntAccsFlg(), strFuncSegmnt, "FuncSegmntAccessFlag")Then
bverifyFuncSegmntAccsFlag=false
End If
End If
verifyFuncSegmntAccsFlag=bverifyFuncSegmntAccsFlag
End Function

'[Click Modify hyperlink on Function Master page]
Public Function ClickModifyFuncMstr(strFuncID, strFuncType, strChannel, strStaffAcs, strSegSegmntScs, strAccntAccs)
Dim bClickModifyFuncMstr: ClickModifyFuncMstr = True

lstModify = checknull("FUNC_ID:"&strFuncID&"|FUNC_TYPE:"&strFuncType&"|CHANNEL:"&strChannel&"|STAFF_ACCESS:"&strStaffAcs&"|SEGMENT_ACCESS:"&strSegSegmntScs&"|ACCOUNT_ACCESS_FLAG:"&strAccntAccs&"")

bClickModifyFuncMstr=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModify,"FunctionMaster" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
ClickModifyFuncMstr=bClickModifyFuncMstr
End Function

'[Set UserProfile Funct Master Accnt Accs Flag as]
Public Function selectUsrProfileAccntAccs(strUsrprofAcntFlag)
bselectUsrProfileAccntAccs=true
If Not IsNull(strUsrprofAcntFlag) Then
UserProfileAdmin.txtAccntAccsFlg().set strUsrprofAcntFlag
End If
waitForIcallLoading
selectUsrProfileAccntAccs=bselectUsrProfileAccntAccs
End Function

'[Set UserProfile Funct Master Segmnt Accs Flag as]
Public Function selectUsrProfileSegmntAccs(strUsrprofSegmntFlg)
bselectUsrProfileSegmntAccs=true
If Not IsNull(strUsrprofSegmntFlg) Then
UserProfileAdmin.txtSegmntAccsFlg().set strUsrprofSegmntFlg
End If
waitForIcallLoading
selectUsrProfileSegmntAccs=bselectUsrProfileSegmntAccs
End Function

'[Set UserProfile Funct Master Staff Accs Flag as]
Public Function selectUsrProfilStaffAccs(strUsrprofStaffFlg)
bselectUsrProfilStaffAccs=true
If Not IsNull(strUsrprofStaffFlg) Then
UserProfileAdmin.txtStaffAccsFlg().set strUsrprofStaffFlg
End If
waitForIcallLoading
selectUsrProfilStaffAccs=bselectUsrProfilStaffAccs
End Function

'[Click View hyperlink on Function Master page]
Public Function ClickViewFuncMstr(strFuncID, strFuncType, strChannel, strStaffAcs, strSegSegmntScs, strAccntAccs)
Dim bClickViewFuncMstr: ClickViewFuncMstr = True

lstView = checknull("FUNC_ID:"&strFuncID&"|FUNC_TYPE:"&strFuncType&"|CHANNEL:"&strChannel&"|STAFF_ACCESS:"&strStaffAcs&"|SEGMENT_ACCESS:"&strSegSegmntScs&"|ACCOUNT_ACCESS_FLAG:"&strAccntAccs&"")

bClickViewFuncMstr=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstView,"FunctionMaster" ,"Action", "View",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
ClickViewFuncMstr=bClickViewFuncMstr
End Function

'*********************************End********************************************
'***************************Functions added for MA questions Kalai 101116*************************

'[Set text on Category textfield for Add new Values Popup]
Public Function setCategory_AddNewValues(strCategoryTxt)
bsetCategory_AddNewValues=true
	If Not IsNull(strIVRMenuTxt) Then
	UserProfileAdmin.txtCategory().set strCategoryTxt
	End If
WaitForICallLoading
setCategory_AddNewValues=bsetCategory_AddNewValues
End Function

'[Set text on QuestionType textfield for Add new Values Popup]
Public Function setQuestionType_AddNewValues(strQuestionTypeTxt)
bsetQuestionType_AddNewValues=true
	If Not IsNull(strQuestionTypeTxt) Then
	UserProfileAdmin.txtQuestionType().set strQuestionTypeTxt
	End If
WaitForICallLoading
setQuestionType_AddNewValues=bsetQuestionType_AddNewValues
End Function

'[Set text on Question textfield for Add new Values Popup]
Public Function setQuestion_AddNewValues(strQuestionTxt)
bsetQuestion_AddNewValues=true
	If Not IsNull(strQuestionTxt) Then
	UserProfileAdmin.txtQuestion().set strQuestionTxt
	End If
WaitForICallLoading
setQuestion_AddNewValues=bsetQuestion_AddNewValues
End Function

'[Set text on ToolTipInfo textfield for Add new Values Popup]
Public Function setToolTipInfo_AddNewValues(strToolTipInfoTxt)
bsetToolTipInfo_AddNewValues=true
	If Not IsNull(strToolTipInfoTxt) Then
	UserProfileAdmin.txtToolTipInfo().set strToolTipInfoTxt
	End If
WaitForICallLoading
setToolTipInfo_AddNewValues=bsetToolTipInfo_AddNewValues
End Function

'[Set text on Points textfield for Add new Values Popup]
Public Function setPoints_AddNewValues(strPointsTxt)
bsetPointsInfo_AddNewValues=true
	If Not IsNull(strPointsTxt) Then
	UserProfileAdmin.txtPoints().set strPointsTxt
	End If
WaitForICallLoading
setPoints_AddNewValues=bsetPoints_AddNewValues
End Function

'[Set text on AutoPassIndicator textfield for Add new Values Popup]
Public Function setAutoPassIndicator_AddNewValues(strAutoPassIndicatorTxt)
bsetAutoPassIndicatorInfo_AddNewValues=true
	If Not IsNull(strAutoPassIndicatorTxt) Then
	UserProfileAdmin.txtAutoPassIndicator().set strAutoPassIndicatorTxt
	End If
WaitForICallLoading
setAutoPassIndicator_AddNewValues=bsetAutoPassIndicator_AddNewValues
End Function

'[Set text on ProductType textfield for Add new Values Popup]
Public Function setProductType_AddNewValues(strProductTypeTxt)
bsetProductTypeInfo_AddNewValues=true
	If Not IsNull(strProductTypeTxt) Then
	UserProfileAdmin.txtProductType().set strProductTypeTxt
	End If
WaitForICallLoading
setProductType_AddNewValues=bsetProductType_AddNewValues
End Function

'[Set text on Status textfield for Add new Values Popup]
Public Function setStatus_AddNewValues(strStatusTxt)
bsetStatus_AddNewValues=true
	If Not IsNull(strStatusTxt) Then
	UserProfileAdmin.txtStatus().set strStatusTxt
	End If
WaitForICallLoading
setStatus_AddNewValues=bsetStatus_AddNewValues
End Function

'[Set text on Channel textfield for Add new Values Popup]
Public Function setChannel_AddNewValues(strChannelTxt)
bsetChannel_AddNewValues=true
	If Not IsNull(strChannelTxt) Then
	UserProfileAdmin.txtChannel().set strStatusTxt
	End If
WaitForICallLoading
setChannel_AddNewValues=bsetChannel_AddNewValues
End Function

'[Set text on UserRole textfield for Add new Values Popup]
Public Function setUserRole_AddNewValues(strUserRoleTxt)
bsetUserRole_AddNewValues=true
	If Not IsNull(strUserRoleTxt) Then
	UserProfileAdmin.txtUserRole().set strUserRoleTxt
	End If
WaitForICallLoading
setUserRole_AddNewValues=bsetUserRole_AddNewValues
End Function

'[Set text on UserGroup textfield for Add new Values Popup]
Public Function setUserGroup_AddNewValues(strUserGroupTxt)
bsetUserGroup_AddNewValues=true
	If Not IsNull(strUserGroupTxt) Then
	UserProfileAdmin.txtUserGroup().set strUserGroupTxt
	End If
WaitForICallLoading
setUserGroup_AddNewValues=bsetUserGroup_AddNewValues
End Function

'[Set text on InteractionMode textfield for Add new Values Popup]
Public Function setInteractionMode_AddNewValues(strInteractionModeTxt)
bsetInteractionMode_AddNewValues=true
	If Not IsNull(strInteractionModeTxt) Then
	UserProfileAdmin.txtInteractionMode().set strInteractionModeTxt
	End If
WaitForICallLoading
setInteractionMode_AddNewValues=bsetInteractionMode_AddNewValues
End Function

'[Click MA Questions View Link]
Public Function lnkViewMAQuestions(strQuestionTxt, strCategoryTxt, strProductTypeTxt)
Dim blnkViewMAQuestions: lnkViewMAQuestions = True
   		strMonth=Month(strtCreateDateTime)
         strYear=Year(strtCreateDateTime)
         strDay=Day(strtCreateDateTime)
         If strDay>=0 and strDay<10 Then
         	strDay=0&strDay
         End If
         strHour=Hour(strtCreateDateTime)
         'If strHour=-1 Then
         '	strHour=23
         'End If
         If strHour>=0 and strHour<10 Then
         	strHour=0&strHour
         End If
         strMin=Minute(strtCreateDateTime)
         If strMin>=0 and strMin<10 Then
         	strMin=0&strMin
         End If
         strMonthName=MonthName(strMonth, True)

lstViewMAQuestions = checknull("Question:"&strQuestionTxt&"|Category:"&strCategoryTxt&"|Product Type:"&strProductTypeTxt&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")

blnkViewMAQuestions=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewMAQuestions,"UserProfileAdmin" ,"Action", "View",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkViewMAQuestions=blnkViewMAQuestions
End Function

'[Click MA Questions Modify Link]
Public Function lnkModifyMAQuestions(strQuestionTxt, strCategoryTxt, strProductTypeTxt)
Dim blnkModifyMAQuestions: lnkModifyMAQuestions = True
lstModifyMAQuestions = checknull("Question:"&strQuestionTxt&"|Category:"&strCategoryTxt&"|ProductType:"&strProductTypeTxt&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
blnkViewUserProfile=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyOtherAuthCategory,"UserProfileAdmin" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkModifyMAQuestions=blnkModifyOtherMAQuestions
End Function

'[Verify Question Id txt displayed as]
Public Function verifyQstnIdText(strQuestion)
	bverifyQstnIdText=true
	strQstnId=getDBValForColumn("select Question_id from iservesgdev.cca_ma_questions where question="&strQuestion&"")
	If Not verifyFieldValue(UserProfileAdmin.txtParameterType(), strQstnId, "QuestionId") Then
		bverifyQstnIdText=false
	End If
	verifyQstnIdText=bverifyQstnIdText
End Function

'[Verify txt AutoPassInd displayed as]
Public Function verifyAutoPassIndText(strAutoPassIndText)
	bverifyAutoPassIndText=true
	If Not verifyFieldValue(UserProfileAdmin.lblAutoPassInd(), strAutoPassIndText, "AutoPassInd Text") Then
		bverifyAutoPassIndText=false
	End If
	verifyAutoPassIndText=bverifyAutoPassIndText
End Function


'[Verify txt AutoPassIndOldValue displayed as]
Public Function verifyAutoPassIndicatorOldValueText(strAutoPassIndicatorOldValueText)
	bverifyAutoPassIndicatorOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblAutoPassIndicatorOldValue(), strAutoPassIndicatorOldValueText, "AutoPassIndicatorOldValue Text") Then
		bverifyAutoPassIndicatorOldValueText=false
	End If
	verifyAutoPassIndicatorOldValueText=bverifyAutoPassIndicatorOldValueText
End Function

'[Verify txt Category displayed as]
Public Function verifyCategoryText(strCategoryText)
	bverifyCategoryText=true
	If Not verifyFieldValue(UserProfileAdmin.lblCategory(), strCategoryText, "Category Text") Then
		bverifyCategoryText=false
	End If
	verifyCategoryText=bverifyCategoryText
End Function


'[Verify txt CategoryOldValue displayed as]
Public Function verifyCategoryOldValueText(strCategoryOldValueText)
	bverifyCategoryOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblCategoryOldValue(), strCategoryOldValueText, "CategoryOldValue Text") Then
		bverifyCategoryOldValueText=false
	End If
	verifyCategoryOldValueText=bverifyCategoryOldValueText
End Function

'[Verify txt Channel displayed as]
Public Function verifyChannelText(strChannelText)
	bverifyChannelText=true
	If Not verifyFieldValue(UserProfileAdmin.lblChannel(), strChannelText, "Channel Text") Then
		bverifyChannelText=false
	End If
	verifyChannelText=bverifyChannelText
End Function

'[Verify txt ChannelOldValue displayed as]
Public Function verifyChannelOldValueText(strChannelOldValueText)
	bverifyChannelOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblChannelOldValue(), strChannelOldValueText, "ChannelOldValue Text") Then
		bverifyChannelOldValueText=false
	End If
	verifyChannelOldValueText=bverifyChannelOldValueText
End Function


'[Verify txt InteractionMode displayed as]
Public Function verifyInteractionModeText(strInteractionModeText)
	bverifyInteractionModeText=true
	If Not verifyFieldValue(UserProfileAdmin.lblInteractionMode(), strInteractionModeText, "InteractionMode Text") Then
		bverifyInteractionModeText=false
	End If
	verifyInteractionModeText=bverifyInteractionModeText
End Function

'[Verify txt InteractionModeOldValue displayed as]
Public Function verifylblInteractionModeOldValueText(strlblInteractionModeOldValueText)
	bverifylblInteractionModeOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblInteractionModeOldValue(), strlblInteractionModeOldValueText, "InteractionModeOldValue Text") Then
		bverifylblInteractionModeOldValueText=false
	End If
	verifylblInteractionModeOldValueText=bverifylblInteractionModeOldValueText
End Function

'[Verify txt Points displayed as]
Public Function verifylblPointsText(strlblPointsText)
	bverifylblPointsText=true
	If Not verifyFieldValue(UserProfileAdmin.lblPoints(), strlblPointsText, "Points Text") Then
		bverifylblPointsText=false
	End If
	verifylblPointsText=bverifylblPointsText
End Function

'[Verify txt Points OldValue displayed as]
Public Function verifylblPointsOldValueText(strlblPointsOldValueText)
	bverifylblPointsOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblPointsOldValue(), strlblPointsOldValueText, "PointsOldValue Text") Then
		bverifylblPointsOldValueText=false
	End If
	verifylblPointsOldValueText=bverifylblPointsOldValueText
End Function

'[Verify txt PreviousQuestion displayed as]
Public Function verifylblPreviousQuestionText(strlblPreviousQuestionText)
	bverifylblPreviousQuestionText=true
	If Not verifyFieldValue(UserProfileAdmin.lblPreviousQuestion(), strlblPreviousQuestionText, "PreviousQuestion Text") Then
		bverifylblPreviousQuestionText=false
	End If
	verifylblPreviousQuestionText=bverifylblPreviousQuestionText
End Function

'[Verify txt ProductGroup displayed as]
Public Function verifylblProductGroupText(strlblProductGroupText)
	bverifylblProductGroupText=true
	If Not verifyFieldValue(UserProfileAdmin.lblProductGroup(), strlblProductGroupText, "ProductGroup Text") Then
		bverifylblProductGroupText=false
	End If
	verifylblProductGroupText=bverifylblProductGroupText
End Function

'[Verify txt ProductGroupOldValue displayed as]
Public Function verifylblProductGroupOldValueText(strlblProductGroupOldValueText)
	bverifylblProductGroupOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblProductGroup(), strlblProductGroupOldValueText, "ProductGroupOldValue Text") Then
		bverifylblProductGroupOldValueText=false
	End If
	verifylblProductGroupOldValueText=bverifylblProductGroupOldValueText
End Function

'[Verify txt ProductType displayed as]
Public Function verifylblProductTypeText(strlblProductTypeText)
	bverifylblProductTypeText=true
	If Not verifyFieldValue(UserProfileAdmin.lblProductType(), strlblProductTypeText, "ProductType Text") Then
		bverifylblProductTypeText=false
	End If
	verifylblProductTypeText=bverifylblProductTypeText
End Function

'[Verify txt ProductType OldValue displayed as]
Public Function verifylblProductTypeOldValueText(strlblProductTypeOldValueText)
	bverifylblProductTypeOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblProductTypeOldValue(), strlblProductTypeOldValueText, "ProductType OldValue Text") Then
		bverifylblProductTypeOldValueText=false
	End If
	verifylblProductTypeOldValueText=bverifylblProductTypeOldValueText
End Function

'[Verify txt Question displayed as]
Public Function verifylblQuestionText(strlblQuestionText)
	bverifylblQuestionText=true
	If Not verifyFieldValue(UserProfileAdmin.lblQuestion(), strlblQuestionText, "Question Text") Then
		bverifylblQuestionText=false
	End If
	verifylblQuestionText=bverifylblQuestionText
End Function

'[Verify txt QuestionOldValue displayed as]
Public Function verifylblQuestionOldValueText(strlblQuestionOldValueText)
	bverifylblQuestionOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblQuestionOldValue(), strlblQuestionOldValueText, "QuestionOldValue Text") Then
		bverifylblQuestionOldValueText=false
	End If
	verifylblQuestionOldValueText=bverifylblQuestionOldValueText
End Function

'[Verify txt QuestionId displayed as]
Public Function verifylblQuestionIdText(strlblQuestionIdText)
	bverifylblQuestionIdText=true
	If Not verifyFieldValue(UserProfileAdmin.lblQuestionId(), strlblQuestionIdText, "QuestionId Text") Then
		bverifylblQuestionIdText=false
	End If
	verifylblQuestionIdText=bverifylblQuestionIdText
End Function

'[Verify txt QuestionIdOldValue displayed as]
Public Function verifylblQuestionIdOldValueText(strlblQuestionIdOldValueText)
	bverifylblQuestionIdOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblQuestionIdOldValue(), strlblQuestionIdOldValueText, "QuestionIdOldValue Text") Then
		bverifylblQuestionIdOldValueText=false
	End If
	verifylblQuestionIdOldValueText=bverifylblQuestionIdOldValueText
End Function

'[Verify txt QuestionType displayed as]
Public Function verifylblQuestionTypeText(strlblQuestionTypeText)
	bverifylblQuestionTypeText=true
	If Not verifyFieldValue(UserProfileAdmin.lblQuestionType(), strlblQuestionTypeText, "QuestionType Text") Then
		bverifylblQuestionTypeText=false
	End If
	verifylblQuestionTypeText=bverifylblQuestionTypeText
End Function

'[Verify txt QuestionTypeOldValue displayed as]
Public Function verifylblQuestionTypeOldValueText(strlblQuestionTypeOldValueText)
	bverifylblQuestionTypeOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblQuestionTypeOldValue(), strlblQuestionTypeOldValueText, "QuestionType OldValue Text") Then
		bverifylblQuestionTypeOldValueText=false
	End If
	verifylblQuestionTypeOldValueText=bverifylblQuestionTypeOldValueText
End Function

'[Verify txt Status1 displayed as]
Public Function verifylblStatus1Text(strlblStatus1Text)
	bverifylblStatus1Text=true
	If Not verifyFieldValue(UserProfileAdmin.lblStatus1(), strlblStatus1Text, "Status1 Text") Then
		bverifylblStatus1Text=false
	End If
	verifylblStatus1Text=bverifylblStatus1Text
End Function

'[Verify txt StatusOldValue1 displayed as]
Public Function verifylblStatusOldValue1Text(strlblStatusOldValue1Text)
	bverifylblStatusOldValue1Text=true
	If Not verifyFieldValue(UserProfileAdmin.lblStatusOldValue1(), strlblStatusOldValue1Text, "StatusOldValue1 Text") Then
		bverifylblStatusOldValue1Text=false
	End If
	verifylblStatusOldValue1Text=bverifylblStatusOldValue1Text
End Function

'[Verify txt ToolTipInfo displayed as]
Public Function verifylblToolTipInfoText(strlblToolTipInfoText)
	bverifylblToolTipInfoText=true
	If Not verifyFieldValue(UserProfileAdmin.lblToolTipInfo(), strlblToolTipInfoText, "ToolTipInfo Text") Then
		bverifylblToolTipInfoText=false
	End If
	verifylblToolTipInfoText=bverifylblToolTipInfoText
End Function

'[Verify txt lblToolTipInfoOldValue displayed as]
Public Function verifylblToolTipInfoOldValueText(strlblToolTipInfoOldValueText)
	bverifylblToolTipInfoOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblToolTipInfoOldValue(), strlblToolTipInfoOldValueText, "ToolTipInfoOldValue Text") Then
		bverifylblToolTipInfoOldValueText=false
	End If
	verifylblToolTipInfoOldValueText=bverifylblToolTipInfoOldValueText
End Function

'[Verify txt UserGroup displayed as]
Public Function verifylblUserGroupText(strlblUserGroupText)
	bverifylblUserGroupText=true
	If Not verifyFieldValue(UserProfileAdmin.lblUserGroup(), strlblUserGroupText, "UserGroup Text") Then
		bverifylblUserGroupText=false
	End If
	verifylblUserGroupText=bverifylblUserGroupText
End Function

'[Verify txt UserGroupOldValue displayed as]
Public Function verifylblUserGroupOldValueText(strlblUserGroupOldValueText)
	bverifylblUserGroupOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblUserGroupOldValue(), strlblUserGroupOldValueText, "UserGroup OldValue Text") Then
		bverifylblUserGrouOldValueText=false
	End If
	verifylblUserGroupOldValueText=bverifylblUserGroupOldValueText
End Function

'[Verify txt UserRole displayed as]
Public Function verifylblUserRoleText(strlblUserRoleText)
	bverifylblUserRoleText=true
	If Not verifyFieldValue(UserProfileAdmin.lblUserRole(), strlblUserRoleText, "UserRole Text") Then
		bverifylblUserRoleText=false
	End If
	verifylblUserRoleText=bverifylblUserRoleText
End Function

'[Verify txt UserRoleOldValue displayed as]
Public Function verifylblUserRoleOldValueText(strlblUserRoleOldValueText)
	bverifylblUserRoleOldValueText=true
	If Not verifyFieldValue(UserProfileAdmin.lblUserRoleOldValue(), strlblUserRoleOldValueText, "UserRoleOldValue Text") Then
		bverifylblUserRoleOldValueText=false
	End If
	verifylblUserRoleOldValueText=bverifylblUserRoleOldValueText
End Function

'[Verify Category Combobox has Items]
Public Function verifyCategoryComboboxItems(lstCategoryItems)
bverifyCategoryComboboxItems=true
If Not IsNull(lstCategoryItems) Then
If Not verifyComboboxItems (UserProfileAdmin.lblCategory(), lstCategoryItems, "Category List")Then
bverifyCategoryComboboxItems=false
End If
End If
verifyCategoryComboboxItems=bverifyCategoryComboboxItems
End Function

'[Verify Question Type Combobox has Items]
Public Function verifyCategoryComboboxItems(lstQuestionTypeItems)
bverifyQuestionTypeComboboxItems=true
If Not IsNull(lstQuestionTypeItems) Then
If Not verifyComboboxItems (UserProfileAdmin.lblQuestionType(), lstQuestionTypeItems, "QuestionType List")Then
bverifyQuestionTypeComboboxItems=false
End If
End If
verifyQuestionTypeComboboxItems=bverifyQuestionTypeComboboxItems
End Function

'[Execute DB Query to delete MA Questions Data From DB]
Public Function executeMAQues_DBQuery(strQuestionTxt)
bexecuteMAQues_DBQuery=true

strDeleteQuery1 = "Delete from `iservesgdev`.`cca_ma_questions` where `QUESTION` = '"&strQuestionTxt&"'"
bexecuteMAQues_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery1)

strDeleteQuery2 = "Delete from `iservesgdev`.`cca_ma_questions-wrk` where `QUESTION` = '"&strQuestionTxt&"'"
bexecuteMAQues_DBQuery=ExecuteDBQueryToDeleteRecords(strDeleteQuery2)

WaitForICallLoading
executeMAQues_DBQuery=bexecuteMAQues_DBQuery
End Function

'************************Added for Number of Display Questions-Kalai 151116***********************************

'[Set text on CodeValue textfield for Modify Popup]
Public Function setCodeValue_Modify(strCodeValueTxt)
bsetCodeValue_Modify=true
	If Not IsNull(strCodeValueTxt) Then
	UserProfileAdmin.txtCodeValue().set strCodeValueTxt
	End If
WaitForICallLoading
setCodeValue_Modify=bsetCodeValue_Modify
End Function

'[Verify Field CodeKey displayed as]
Public Function verifyCodeKey(strCodeKey)
	bDevPending=false
	bverifyCodeKey=true
	If Not IsNull(strCodeKey) Then
		If Not verifyFieldValue (UserProfileAdmin.lblCodeKey(), strCodeKey, "CodeKey")Then
		bverifyCodeKey=false
		End If
	End If
verifyCodeKey=bverifyCodeKey
End Function 

'[Verify Field CodeKey OldValue displayed as]
Public Function verifyCodeKeyOldValue(strCodeKeyOldValue)
bDevPending=false
bverifyCodeKeyOldValue=true
If Not IsNull(strCodeKeyOldValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblCodeKeyOldValue(), strCodeKeyOldValue, "CodeKey OldValue")Then
bverifyCodeKeyOldValue=false
End If
End If
verifyCodeKeyOldValue=bverifyCodeKeyOldValue
End Function 

'[Verify Field CodeValue displayed as]
Public Function verifyCodeValue(strCodeValue)
bDevPending=false
bverifyCodeValue=true
If Not IsNull(strCodeValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblCodeValue(), strCodeValue, "CodeValue")Then
bverifyCodeValue=false
End If
End If
verifyCodeValue=bverifyCodeValue
End Function 

'[Verify Field CodeValue OldValue displayed as]
Public Function verifyCodeValueOldValue(strCodeValueOldValue)
bDevPending=false
bverifyCodeValueOldValue=true
If Not IsNull(strCodeValueOldValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblCodeValueOldValue(), strCodeValueOldValue, "CodeValue OldValue")Then
bverifyCodeValueOldValue=false
End If
End If
verifyCodeValueOldValue=bverifyCodeValueOldValue
End Function 

'[Verify Field Language displayed as]
Public Function verifyLanguage(strLanguage)
bDevPending=false
bverifyLanguage=true
If Not IsNull(strLanguage) Then
If Not verifyFieldValue (UserProfileAdmin.lblLanguage(), strLanguage, "Language")Then
bverifyLanguage=false
End If
End If
verifyLanguage=bverifyLanguage
End Function 

'[Verify Field Language OldValue displayed as]
Public Function verifyLanguageOldValue(strLanguageOldValue)
bDevPending=false
bverifyLanguageOldValue=true
If Not IsNull(strLanguageOldValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblLanguageOldValue(), strLanguageOldValue, "Language OldValue")Then
bverifyLanguageOldValue=false
End If
End If
verifyLanguageOldValue=bverifyLanguageOldValue
End Function 

'[Verify Field ShortKey displayed as]
Public Function verifyShortKey(strShortKey)
bDevPending=false
bverifyShortKey=true
If Not IsNull(strShortKey) Then
If Not verifyFieldValue (UserProfileAdmin.lblShortKey(), strShortKey, "ShortKey")Then
bverifyShortKey=false
End If
End If
verifyShortKey=bverifyShortKey
End Function 

'[Verify Field ShortKey OldValue displayed as]
Public Function verifyShortKeyOldValue(strShortKeyOldValue)
bDevPending=false
bverifyShortKeyOldValue=true
If Not IsNull(strShortKeyOldValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblShortKeyOldValue(), strShortKeyOldValue, "ShortKey OldValue")Then
bverifyShortKeyOldValue=false
End If
End If
verifyShortKeyOldValue=bverifyShortKeyOldValue
End Function 

'[Verify Field ShortValue displayed as]
Public Function verifyShortValue(strShortValue)
bDevPending=false
bverifyShortValue=true
If Not IsNull(strShortValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblShortValue(), strShortValue, "ShortValue")Then
bverifyShortValue=false
End If
End If
verifyShortValue=bverifyShortValue
End Function 

'[Verify Field ShortValue OldValue displayed as]
Public Function verifyShortValueOldValue(strShortValueOldValue)
bDevPending=false
bverifyShortValueOldValue=true
If Not IsNull(strShortValueOldValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblShortValueOldValue(), strShortValueOldValue, "ShortValue OldValue")Then
bverifyShortValueOldValue=false
End If
End If
verifyShortValueOldValue=bverifyShortValueOldValue
End Function 

'[Verify Field Status NoofQues displayed as]
Public Function verifyStatusNoofQues(strStatusNoofQues)
bDevPending=false
bverifyStatusNoofQues=true
If Not IsNull(strStatusNoofQues) Then
If Not verifyFieldValue (UserProfileAdmin.lblStatusNoofQues(), strStatusNoofQues, "Status NoofQues")Then
bverifyStatusNoofQues=false
End If
End If
verifyStatusNoofQues=bverifyStatusNoofQues
End Function 

'[Verify Field Status NoofQues OldValue displayed as]
Public Function verifyStatusNoofQuesOldValue(strStatusNoofQuesOldValue)
bDevPending=false
bverifyStatusNoofQuesOldValue=true
If Not IsNull(strStatusNoofQuesOldValue) Then
If Not verifyFieldValue (UserProfileAdmin.lblStatusNoofQuesOldValue(), strStatusNoofQuesOldValue, "StatusNoofQues OldValue")Then
bverifyStatusNoofQuesOldValue=false
End If
End If
verifyStatusNoofQuesOldValue=bverifyStatusNoofQuesOldValue
End Function

'[Verify Field Error Msg NoofDisplayQues displayed as]
Public Function verifyErrorNoofDisplayQues(strErrorNoofDisplayQues)
bDevPending=false
bverifyErrorNoofDisplayQues=true
If Not IsNull(strErrorNoofDisplayQues) Then
If Not VerifyInnerText (UserProfileAdmin.lblErrorNoofDisplayQues(), strErrorNoofDisplayQues, "Error Msg NoofDisplayQues")Then
bverifyErrorNoofDisplayQues=false
End If
End If
verifyErrorNoofDisplayQues=bverifyErrorNoofDisplayQues
End Function 

'[Verify AddNewValues button is disabled]
Public Function verifybtnAddNewValues_Disable()
	bverifybtnAddNewValues_Disable=true	
	intbtnAddNewValues =UserProfileAdmin.btnAddNewValues().GetROProperty("disabled")
	If  intbtnAddNewValues=1 Then
		LogMessage "RSLT","Verification","AddNewValues button is disabled as expected.",True
		bverifybtnAddNewValues_Disable=true
	Else
		LogMessage "WARN","Verifiation","AddNewValues button is enabled. Expected disabled.",false
		bverifybtnAddNewValues_Disable=false
	End If
	verifybtnAddNewValues_Disable=bverifybtnAddNewValues_Disable
End Function

'[Click No Of Display Questions View Link]
Public Function lnkViewNoofDisplayQues(strCodeKeyTxt, strCodeValueTxt)
	Dim blnkViewNoofDisplayQues: lnkViewNoofDisplayQues = True
	'getDateMnthYear()
	lstViewNoofDisplayQues = checknull("Code Key:"&strCodeKeyTxt&"|Code Value:"&strCodeValueTxt&"")
	blnkViewNoofDisplayQues=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewNoofDisplayQues,"UserProfileAdmin" ,"Action", "View",false,null ,null ,null)
	WaitForICallLoading
	lnkViewNoofDisplayQues=blnkViewNoofDisplayQues
End Function

'[Click No Of Display Questions Modify Link]
Public Function lnkModifyNoofDisplayQues(strCodeKeyTxt, strCodeValueTxt)
	Dim blnkModifyNoofDisplayQues: lnkModifyNoofDisplayQues = True
	'getDateMnthYear()
	lstModifyNoofDisplayQues = checknull("Code Key:"&strCodeKeyTxt&"|Code Value:"&strCodeValueTxt&"")
	blnkModifyNoofDisplayQues=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyNoofDisplayQues,"UserProfileAdmin" ,"Action", "Modify",false,null ,null ,null)
	WaitForICallLoading
	lnkModifyNoofDisplayQues=blnkModifyNoofDisplayQues
End Function

'[Verify User Profile Modify link  is disabled]
Public Function verifyModifyLink_Disabled(strCodeKeyTxt, strCodeValueTxt)
bverifyModifyLink_Disabled=true
'getDateMnthYear()
lstModifyUserProfile = checknull("Code Key:"&strCodeKeyTxt&"|Code Value:"&strCodeValueTxt&"")
bverifyModifyLink_Disabled=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyUserProfile,"UserRoleAdmin" ,"Action", "Modify",false,null ,null ,null)
WaitForICallLoading
verifyModifyLink_Disabled=bverifyModifyLink_Disabled
End Function

'************************End********************************************************************************
'''***********************newly added other Auth kalai************************************
'[Click Other Auth Category View Link]
Public Function lnkViewOtherAuthCategory(strAuthCategoryTxt, strDescriptionTxt, strIVRHotlineTxt)
Dim blnkViewOtherAuthCategory: lnkViewOtherAuthCategory = True
   		strMonth=Month(strtCreateDateTime)
         strYear=Year(strtCreateDateTime)
         strDay=Day(strtCreateDateTime)
         If strDay>=0 and strDay<10 Then
         	strDay=0&strDay
         End If
         strHour=Hour(strtCreateDateTime)
         'If strHour=-1 Then
         '	strHour=23
         'End If
         If strHour>=0 and strHour<10 Then
         	strHour=0&strHour
         End If
         strMin=Minute(strtCreateDateTime)
         If strMin>=0 and strMin<10 Then
         	strMin=0&strMin
         End If
         strMonthName=MonthName(strMonth, True)

lstViewOtherAuthCategory = checknull("Auth Category:"&strAuthCategoryTxt&"|Description:"&strDescriptionTxt&"|IVR Hotline:"&strIVRHotlineTxt&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")

blnkViewOtherAuthCategory=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstViewOtherAuthCategory,"UserProfileAdmin" ,"Action", "View",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkViewOtherAuthCategory=blnkViewOtherAuthCategory
End Function

'[Click Other Auth Category Modify Link]
Public Function lnkModifyOtherAuthCategory(strAuthCategoryTxt, strDescriptionTxt, strIVRHotlineTxt)
Dim blnkModifyOtherAuthCategory: lnkModifyOtherAuthCategory = True
lstModifyOtherAuthCategory = checknull("Auth Category:"&strAuthCategoryTxt&"|Description:"&strDescriptionTxt&"|IVR Hotline:"&strIVRHotlineTxt&"|Created Date:"&strDay&" "&strMonthName&" "&strYear&" "&strHour&""&strMin&"")
blnkViewUserProfile=selectTableLinkWithLinkName(UserProfileAdmin.tblUserProfileHeader,UserProfileAdmin.tblUserProfileContent,lstModifyOtherAuthCategory,"UserProfileAdmin" ,"Action", "Modify",true,UserProfileAdmin.lnkNext ,UserProfileAdmin.lnkNext1 ,UserProfileAdmin.lnkPrevious)
WaitForICallLoading
lnkModifyOtherAuthCategory=blnkModifyOtherAuthCategory
End Function

'[Verify Field Description PinkPanel in UserProfile displayed as]
Public Function verifyDescriptionPinkPanelText(strDescriptionPinkPanelTxt)
   bDevPending=false
   bVerifyDescriptionPinkPanelText=true
   If Not IsNull(strDescriptionPinkPanelTxt) Then
     If Not VerifyInnerText (UserProfileAdmin.lblDescriptionPinkPanel(), strDescriptionPinkPanelTxt, "Description PinkPanel")Then
	   bVerifyDescriptionPinkPanelText=false
	End If
   End If
   verifyDescriptionPinkPanelText=bVerifyDescriptionPinkPanelText
End Function

'*********************************************************************************************************************

