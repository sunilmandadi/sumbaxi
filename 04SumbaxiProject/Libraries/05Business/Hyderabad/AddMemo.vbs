'*****This is auto generated code using code generator please Re-validate ****************

'[Select Combobox LetterCode as]
Public Function selectLetterCodeComboBox(strLetterCode)
   bDevPending=false
   bSelectLetterCodeComboBox=true
   If Not IsNull(strLetterCode) Then
       If Not (selectItem_Combobox (AddMemo_Page.lstLetterCode(), strLetterCode))Then
            LogMessage "WARN","Verification","Failed to select :"&strControlName&" From LetterCode drop down list" ,false
           bSelectLetterCodeComboBox=false
       End If
   End If
   selectLetterCodeComboBox=bSelectLetterCodeComboBox
End Function

'[Get selected item from combo box LetterCode]
Public Function getLetterCodeSelectedItem()
   bDevPending=false
   getLetterCodeSelectedItem=getVadinCombo_SelectedItem(AddMemo_Page.lstLetterCode)
End Function

'[Verify Combobox LetterCode displayed as]
Public Function verifyLetterCodeText(strLetterCode)
   bDevPending=false
   bVerifyLetterCodeText=true
   If Not IsNull(strLetterCode) Then
       If Not verifyComboSelectItem (AddMemo_Page.lstLetterCode(), strLetterCode, "LetterCode")Then
           bVerifyLetterCodeText=false
       End If
   End If
   verifyLetterCodeText=bVerifyLetterCodeText
End Function


'[Verify Field Comment on Add Memo displayed as]
Public Function verifyCommentText_AM(strComment)
   bDevPending=false
   bVerifyCommentText=true
   If Not IsNull(strComment) Then
       If Not VerifyField(AddMemo_Page.txtComment(), strComment, "Comment")Then
           bVerifyCommentText=false
       End If
   End If
   verifyCommentText_AM=bVerifyCommentText
End Function


'[Set TextBox on Add Memo Comment to]
Public Function setCommentTextbox_AM(strComment)
   bDevPending=false
   AddMemo_Page.txtComment.Set(strComment)
   If Err.Number<>0 Then
       setCommentTextbox_AM=false
            LogMessage "WARN","Verification","Failed to Set Text Box :Comment" ,false
       Exit Function
   End If
   setCommentTextbox_AM=true
End Function

'[Click Button Submit on Add Memo]
Public Function clickButtonSubmit_AM()
   bDevPending=false
    WaitForICallLoading	
   AddMemo_Page.btnSubmit.click
    If Err.Number<>0 Then
       clickButtonSubmit_AM=false
            LogMessage "WARN","Verification","Failed to Click Button : Submit" ,false
       Exit Function
   End If
'   '*************** Capturing time stamp to open Memo for this SR by Manish
'	strRunTimeTimeStamp_Step="Click Button Submit on Add Memo"
'	'strDate="04 Feb 2017"
'	
'	'var = "04/02/2017"
'	
'	var=runtimeDate
'	var_month=mid(var,4,2)
'	var_month_change=MonthName(var_month,True)
'	var_month_format=replace(var,var_month,var_month_change)
'	strDate=replace(var_month_format,"/"," ")
'	
'	'strDate= FormatDateTime(Now(),vbLongDate)
'	'strTempTime=FormatDateTime(Now(),vbShortTime)
'	
'	strTempTime_Replace=Replace(runtimeDate,":","-")
'
'	'strTempTime_Replace = "21-26"
' 	strTimeStamp=strDate&" "&strTempTime
'	insertDataStore "TimeStamp", strTimeStamp
   clickButtonSubmit_AM=true
End Function


'[Click Button Cancel on Add Memo]
Public Function clickButtonCancel_AM()
   bDevPending=false
   AddMemo_Page.btnCancel.click
   If Err.Number<>0 Then
       clickButtonCancel_AM=false
            LogMessage "WARN","Verification","Failed to Click Button : Cancel" ,false
       Exit Function
   End If
   clickButtonCancel_AM=true
End Function

'[Click Button RefreshStatus_RequestSubmitted on Add Memo Screen]
Public Function clickButtonRefreshStatus_AM()
   bDevPending=false
   AddMemo_Page.btnRefreshStatus.click
   If Err.Number<>0 Then
       clickButtonRefreshStatus_AM=false
            LogMessage "WARN","Verification","Failed to Click Button : RefreshStatus" ,false
       Exit Function
   End If
   clickButtonRefreshStatus_AM=true
End Function

'[Click Button OK_ValidationMsg on Add Memo]
Public Function clickButtonOK_ValidationMsg_AM()
   bDevPending=false
   AddMemo_Page.btnOK_ValidationMsg.click
   If Err.Number<>0 Then
       clickButtonOK_ValidationMsg_AM=false
            LogMessage "WARN","Verification","Failed to Click Button : OK_ValidationMsg" ,false
       Exit Function
   End If
   clickButtonOK_ValidationMsg_AM=true
End Function

'[Verify Popup ValidationMessage exist on Add Memo]
Public Function verifyPopupValidationMessageexist_AP(bExist)
   bDevPending=false
   bActualExist=strTestAppFrameClass.Exist()
   If bExist And  bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage Exists As Expected" ,true
       verifyPopupValidationMessageexist_AP=True
   ElseIf not bExist And  not bActualExist  Then
       LogMessage "RSLT","Verification","Popup :ValidationMessage does not Exists As Expected" ,true
       verifyPopupValidationMessageexist_AP=True
   ElseIf bExist And  not bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage does not Exists As Expected" ,False
       verifyPopupValidationMessageexist_AP=False
   ElseIf not bExist And   bActualExist  Then
       LogMessage "WARN","Verification","Popup :ValidationMessage Still Exists" ,False
       verifyPopupValidationMessageexist_AP=False
   End If
End Function

'[Verify Add Memo Popup Header]
Public Function verifyPopupHeader_AP()
   bDevPending=false
   var=AddMemo_Page.popupHeader.getroproperty("innertext")
   If var <> "Add Memo" or Err.Number<>0 Then
       verifyPopupHeader_AP=false
            LogMessage "WARN","Verification","Popup Header does not Exists As Expected - Add Memo" ,false
       Exit Function
   End If
   verifyPopupHeader_AP=true
End Function

'[Verify Field Comment on Add Memo popup]
Public Function verifyCommentText_ViewSR_AM(strComment,strCreatedBy)
   bDevPending=false
   verifyCommentText_ViewSR_AM=true
   varComment=AddMemo_Page.popupValidationMessage.getroproperty("innertext")
   If Not IsNull(strComment) Then
       If instr(varComment,strComment)=0 Then
           verifyCommentText_ViewSR_AM=false
       End If
   End If
   If Not IsNull(strCreatedBy) Then
       If instr(varComment,strCreatedBy)=0 Then
           verifyCommentText_ViewSR_AM=false
       End If
   End If
   AddMemo_Page.btnOK.click
   verifyCommentText_ViewSR_AM=verifyCommentText_ViewSR_AM
End Function
