'[Verify table row Selected Card displayed for Card Replacement]
Public Function VerifytblSelectedCardReplacement(lstSelectedCard)
	VerifytblSelectedCardReplacement = VerifyTableSingleRowData(coCardReplacement_Page.tblSelectedCardReplacementHeader,coCardReplacement_Page.tblSelectedCardReplacementBody,lstSelectedCard,"Selected Card")
End Function

'[Verify Card Replacement Label Replacement Option]
Public Function VerifyCR_ReplacementOpt(strReplcmntOpt)
	bVerify=True
	
	If Not IsNull(strReplcmntOpt) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblReplacementOptn(),strReplcmntOpt,"Replacement Option label Name")Then
				bVerify = False
			End If
	End If
	
				
	VerifyCR_ReplacementOpt=bVerify
	
End Function

Public Function  verifyPlaceHolder(objField, strExpectedValue, strFieldName)
Dim strObservedValue
strObservedValue = objField.GetROProperty("placeholder")
'strObservedValue = Replace(strObservedValue,",","")

If  Ucase(Trim(strObservedValue)) = UCase(Trim(strExpectedValue)) Then
LogMessage "RSLT", "Verification","Field " & strFieldName & " matching with the expected value. Expected: "& strExpectedValue &" , Actual: "& strObservedValue, True
verifyPlaceHolder = True
else
LogMessage "WARN", "Verification","Field " & strFieldName & " not matching with the expected value. Expected: "& strExpectedValue &" , Actual: "& strObservedValue, False
verifyPlaceHolder = False
End If
End Function

'[Verify Card Replacement Label Value Replacement Option]
Public Function VerifyCR_ReplacementOptVal(strReplcmntOptVal)
	bVerify=True
	
	If Not IsNull(strReplcmntOptVal) Then
	
		If Not verifyPlaceHolder(coCardReplacement_Page.lblReplacementOptnVal(),strReplcmntOptVal,"Replacement Option label value Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_ReplacementOptVal=bVerify
	
End Function

'[Verify Card Replacement Label Selected Cards]
Public Function VerifyCR_SlctdCards(strSlctdCards)
	bVerify=True
	
	If Not IsNull(strSlctdCards) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblSelectCards(),strSlctdCards,"Selected Cards label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_SlctdCards=bVerify
	
End Function

'[Verify Card Replacement Label Embosser Name]
Public Function VerifyCR_EmbosserName(strEmbosserName)
	bVerify=True
	
	If Not IsNull(strEmbosserName) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblEmbosserName(),strEmbosserName,"Embosser Name label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_EmbosserName=bVerify
	
End Function

'[Verify Card Replacement Label Value Embosser Name]
Public Function VerifyCR_EmbosserNameVal(strEmbosserNameVal)
	bVerify=True
	
	If Not IsNull(strEmbosserNameVal) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblEmbosserNameval(),strEmbosserNameVal,"Embosser Name label Value Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_EmbosserNameVal=bVerify
	
End Function

'[Verify Card Replacement Label New Expiry Date]
Public Function VerifyCR_NewExpDt(strNewExpDt)
	bVerify=True
	
	If Not IsNull(strNewExpDt) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblNewExpDt(),strNewExpDt,"New Expiry Date label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_NewExpDt=bVerify
	
End Function

'[Verify Card Replacement Label Value New Expiry Date]
Public Function VerifyCR_NewExpDtVal(strNewExpDtVal)
	bVerify=True
	
	If Not IsNull(strNewExpDtVal) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblNewExpDtVal(),strNewExpDtVal,"New Expiry Date label Value Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_NewExpDtVal=bVerify
	
End Function

'[Verify Card Replacement Label Delivery Address]
Public Function VerifyCR_DelivryAdrs(strDelivryAdrs)
	bVerify=True
	
	If Not IsNull(strDelivryAdrs) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblDelivryAdrs(),strDelivryAdrs,"Delivery Address label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_DelivryAdrs=bVerify
	
End Function

'[Verify Card Replacement Label Value Delivery Address]
Public Function VerifyCR_DelivryAdrsVal(strDelivryAdrsVal)
	bVerify=True
	
	If Not IsNull(strDelivryAdrsVal) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblDelivryAdrsVal(),strDelivryAdrsVal,"Delivery Address label Value Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_DelivryAdrsVal=bVerify
	
End Function

'[Verify Card Replacement Label Address]
Public Function VerifyCR_Adrs(strAdrs)
	bVerify=True
	
	If Not IsNull(strAdrs) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblAdrs(),strAdrs,"Address label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_Adrs=bVerify
	
End Function

'[Verify Card Replacement Label Value Address]
Public Function VerifyCR_AdrsVal(strAdrsVal)
	bVerify=True
	
	If Not IsNull(strAdrsVal) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblAdrsVal(),strAdrsVal,"Address label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_AdrsVal=bVerify
	
End Function

'[Verify Card Replacement Label Desc]
Public Function VerifyCR_Desc(strDesc)
	bVerify=True
	
	If Not IsNull(strDesc) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblDesc(),strDesc,"Desc label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_Desc=bVerify
	
End Function

'[Verify Card Replacement Label Value Desc]
Public Function VerifyCR_DescVal(strDescVal)
	bVerify=True
	
	If Not IsNull(strDescVal) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblDescVal(),strDescVal,"Desc label value Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_DescVal=bVerify
	
End Function

'[Verify Card Replacement Label Comments]
Public Function VerifyCR_comments(strCmnts)
	bVerify=True
	
	If Not IsNull(strCmnts) Then
	
		If Not verifyInnerText(coCardReplacement_Page.lblComments(),strCmnts,"Comments label Name")Then
				bVerify = False
			End If
	End If
				
	VerifyCR_comments=bVerify
	
End Function

'[Click Card Replacement Submit button]
Public Function ClickSubmitCardReplacement()
	bVerify=True
	coCardReplacement_Page.btnSubmit.click	
	WaitForIServeLoading	
	ClickSubmitCardReplacement=bVerify
End Function
