'[Verify list of values displayed for Activation Type dropdown]
Public Function VerifylstActivationType_CA(lstActivationType)
bVerifyValues = False
If Not IsNull(lstActivationType) Then
bVerifyValues = verifyComboboxItems(coCardActivation_Page.lstActivationType,lstActivationType,"Activation Type")		
End If
VerifylstActivationType_CA = bVerifyValues	
End Function

'[Select Activation Type dropdown]
Public Function SelectActivationType_CA(strActivationType)
bVerifySelectActivationType = False
coCardActivation_Page.lstActivationType.Click
If Err.Number <> 0 Then
	LogMessage "WARN","Verification","Unable to Click dropdown Activation Type displayed in the page", False
Else 
	bVerifySelectActivationType = SelectItemfromList(strActivationType,"Activation Type")
End IF
SelectActivationType_CA  = bVerifySelectActivationType	
End Function

'[Verify table row Selected Card displayed for Card Activation]
Public Function VerifytblSelectedCard_CA(lstSelectedCard)
	VerifytblSelectedCard_CA = VerifyTableSingleRowData(coCardActivation_Page.tblSelectedCardHeader,coCardActivation_Page.tblSelectedCardBody,lstSelectedCard,"Selected Card")
End Function
