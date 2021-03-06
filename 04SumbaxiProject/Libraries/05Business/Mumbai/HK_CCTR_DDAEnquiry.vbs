'[Verify Direct Debit Account link is enabled]
Public Function verifyDDALink_Enable()
	VerifyObjectEnabledDisabled HK_CCTR_DDAEnquiry_Page.lnkDirectDebitAccount(),"Enable","Direct Debit Account link"
End Function
'[Verify Direct Debit Account link is disabled]
Public Function verifyDDALink_Disable()
	VerifyObjectEnabledDisabled HK_CCTR_DDAEnquiry_Page.lnkDirectDebitAccount(),"Disable","Direct Debit Account link"
End Function
'[Click on Direct Debit Account link]
Public Function clickLinkDDA()
	ClickOnObject HK_CCTR_DDAEnquiry_Page.lnkDirectDebitAccount(),"Direct Debit Account link"
	WaitForICallLoading
End Function
'[Verify DDA Enquiry Pink Panel Details]
Public Function verifyPinkPanelDDAEnquiry(strDDAName,strIDNumber,strOwner)
	blnverifyPinkPanelDDAEnquiry=true
	If HK_CCTR_DDAEnquiry_Page.lblDDAName().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Name Label is displayed as expected.",True
		blnverifyPinkPanelDDAEnquiry=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Name Label",false
		blnverifyPinkPanelDDAEnquiry=false
	End If
	If HK_CCTR_DDAEnquiry_Page.lblDDAIDNumber().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","ID Number Label is displayed as expected.",True
		blnverifyPinkPanelDDAEnquiry=true
	Else
		LogMessage "WARN","Verifiation","Failed to display ID Number Label",false
		blnverifyPinkPanelDDAEnquiry=false
	End If
	If HK_CCTR_DDAEnquiry_Page.lblDDAOwner().Exist(gWaitTime) Then
		LogMessage "RSLT","Verification","Owner Label is displayed as expected.",True
		blnverifyPinkPanelDDAEnquiry=true
	Else
		LogMessage "WARN","Verifiation","Failed to display Owner Label",false
		blnverifyPinkPanelDDAEnquiry=false
	End If
	
	If Not verifyInnerText_Pattern(HK_CCTR_DDAEnquiry_Page.weleDDAName(), strDDAName, "Name Text") Then
		blnverifyPinkPanelDDAEnquiry=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_DDAEnquiry_Page.weleDDAIDNumber(), strIDNumber, "ID Number Text") Then
		blnverifyPinkPanelDDAEnquiry=false
	End If
	If Not verifyInnerText_Pattern(HK_CCTR_DDAEnquiry_Page.weleDDAOwner(), strOwner, "Owner Text") Then
		blnverifyPinkPanelDDAEnquiry=false
	End If
	verifyPinkPanelDDAEnquiry=blnverifyPinkPanelDDAEnquiry
End Function
'[Click on DDA Refresh Button]
Public Function clickRefreshButton()
	blnclickRefreshButton=true
	HK_CCTR_DDAEnquiry_Page.btnDDARefresh().Click
	WaitForICallLoading
	If Err.Number<>0 Then
		blnclickRefreshButton=false
		LogMessage "WARN","Verification","Failed to Click Button :Refresh" ,false
	Else
		LogMessage "RSLT","Verification","Clicked on Refresh Button as expected.",True
		blnclickRefreshButton=true
	End If
	clickRefreshButton=blnclickRefreshButton
End Function
'[Verify DDA Enquiry Action Combobox has Items]
Public Function verifyActionComboboxItems(lstItems)
	blnverifyActionComboboxItems=true
	
	If Not IsNull(lstItems) Then
		If Not verifyComboboxItems (HK_CCTR_DDAEnquiry_Page.lstDDAEnqAction(),lstItems,"DDA Enquiry Action") Then
			blnverifyActionComboboxItems=false
		End If
	End If
	verifyActionComboboxItems=blnverifyActionComboboxItems
End Function
'[Select Combobox DDA Enquiry Action]
Public Function selectDDAEnquiryActionComboBox(strAction)

	blnselectDDAEnquiryActionComboBox=true
	If Not IsNull(strAction) Then
		If Not (selectItem_Combobox (HK_CCTR_DDAEnquiry_Page.lstDDAEnqAction(),strAction))Then
			LogMessage "WARN","Verification","Failed to select :"&strControlName&" From Action drop down list" ,false
			blnselectDDAEnquiryActionComboBox=false
		End If
	End If
	WaitForICallLoading
	selectDDAEnquiryActionComboBox=blnselectDDAEnquiryActionComboBox
End Function
'[Verify row Data in Table for Action in DDA Enquiry Page]
Public Function verifytblContentActionDDAEnquiry(arrActionRowDataList)

   verifytblContentActionDDAEnquiry=verifyTableContentList(HK_CCTR_DDAEnquiry_Page.tblDDAEnquiryActionHeader(),HK_CCTR_DDAEnquiry_Page.tblDDAEnquiryActionContent(),arrActionRowDataList,"Action - DDA Enquiry",false,NULL,NULL,NULL)
End Function
'[Verify error message is displayed Upon selecting Action dropdown when no record found in DDA Enquiry]
Public Function verifyErrorMessageActionDDAEnquiry()
	blnverifyErrorMessageActionDDAEnquiry = true
	If Not IsNull(strErrorMessage) Then
		If Not VerifyInnerText (HK_CCTR_DDAEnquiry_Page.weleNoRecordErrMsg(),"No Records Found", "Error Message DDA Enquiry") Then
			blnverifyErrorMessageActionDDAEnquiry = false
		End If
	End If
	verifyErrorMessageActionDDAEnquiry = blnverifyErrorMessageActionDDAEnquiry
End Function
