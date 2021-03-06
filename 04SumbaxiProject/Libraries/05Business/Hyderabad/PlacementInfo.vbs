'[Verify FD Placement details in the table displayed]
Public Function verifyPlacementdetails_FDPI(lstlstPlacementDetails)
   bverifyPlacementdetails=verifyTableContentList(PlacementInfo.tblPlacementdetailsHeader, PlacementInfo.tblPlacementdetailsContent,lstlstPlacementDetails,"Placement Details",false,NULL,NULL,NULL)
   verifyPlacementdetails_FDPI = bverifyPlacementdetails
End Function

'[Click on View hyperlink in Placement details table]
Public Function verifyViewHyperlink_FDPI(lstPlacementDetails)
	bverifyViewHyperlink = selectTableLink(PlacementInfo.tblPlacementdetailsHeader, PlacementInfo.tblPlacementdetailsContent,lstPlacementDetails, "Placement Details", "Maturity Instruction", false, null, null, null)
	verifyViewHyperlink_FDPI = bverifyViewHyperlink
End Function

'[Verify maturity Instruction details displayed on clicking View link from the table]
Public Function verifytextMaturityInstruction_FDPI(strMaturityInst)
	bVerifytext = True 
	StrtxtMI = PlacementInfo.PopupMItxtMessage.GetROProperty("innertext")
	StrActualtext= Instr(StrtxtMI, strMaturityInst)
	If StrActualtext = 1  Then
	   LogMessage "RSLT","Verification","Maturity Instruction details message text displayed as expected. Expected : "&strMaturityInst&" Actual : "&StrtxtMI&" ",True
	   bVerifytext = True 
	Else 
	   LogMessage "RSLT","Verification","Maturity Instruction details message text doesnt displayed as expected. Expected : "&strMaturityInst&" Actual : "&StrtxtMI&"",False
	   bVerifytext = False
	End If
	verifytextMaturityInstruction_FDPI = bVerifytext
End Function

'[Verify Account Summary field values displayed as]
Public Function verifyAccSummaryfields_FYFDPI(strSignal,strBalSGDEqu,strBalUSDEqu)
	bverifyAccSummaryfields=true
	If Not IsNull (strSignal) Then
		If Not verifyInnerText(PlacementInfo.lblSignal(),strSignal, "Name") Then
			LogMessage "WARN","Verification","Signal field value not displayed as expected:"&strSignal&"" ,false
			bverifyAccSummaryfields=false
		End If
	End If
	If Not IsNull (strBalSGDEqu) Then
		If Not verifyInnerText(PlacementInfo.lblBalSGDEquivalent(),strBalSGDEqu, "CIN") Then
			LogMessage "WARN","Verification","Balance SGD Equivalent not displayed as expected:"&strBalSGDEqu&"" ,false
			bverifyAccSummaryfields=false
		End If
	End If
	If Not IsNull (strBalUSDEqu) Then
		If Not verifyInnerText(PlacementInfo.lblBalUSDEquialent(),strBalUSDEqu, "Segment") Then
			LogMessage "WARN","Verification","Balance USD Equivalent not displayed as expected:"&strBalUSDEqu&"" ,false
			bverifyAccSummaryfields=false
		End If
	End If
	verifyAccSummaryfields_FYFDPI=bverifyAccSummaryfields
End Function

'[Verify FCFD Account Summary in the table displayed]
Public Function verifyAccountSummary_FYFDPI(lstlstAccountSummary)
   bverifyAccSummary=verifyTableContentList(PlacementInfo.tblAccountSummaryHeader, PlacementInfo.tblAccountSummaryContent,lstlstAccountSummary,"Account Summary",false,NULL,NULL,NULL)
   verifyAccountSummary_FYFDPI = bverifyAccSummary
End Function

'[Click on Currency hyperlink in Account Summary table]
Public Function verifyCurrencyHyperlink_FDPI(lstAccountSummary)
	bverifyCurrencyHyperlink = selectTableLink(PlacementInfo.tblAccountSummaryHeader, PlacementInfo.tblAccountSummaryContent,lstAccountSummary, "Account Summary", "Currency", false, null, null, null)
	verifyCurrencyHyperlink_FDPI = bverifyCurrencyHyperlink
End Function

'[Verify table name displayed with the currency name selected]
Public Function verifyPlacementSummarytablename_FYFDPI(strCurrency)
  bverifyPlacementSummarytablename = True	
 	  If Not VerifyInnerText (PlacementInfo.lbltablename(), strCurrency, "Placement Summary - "&strCurrency&"")Then
 	     bverifyPlacementSummarytablename=false
      End If
	verifyPlacementSummarytablename_FYFDPI = bverifyPlacementSummarytablename
End Function

'[Verify Placement Summary details for the selected currency in the table displayed as]
Public Function verifyPlacementSummary_FYFDPI(lstlstPlacementSummary)
   bverifyPlacementSummary=verifyTableContentList(PlacementInfo.tblPlacementSummaryHeader, PlacementInfo.tblPlacementSummaryContent,lstlstPlacementSummary,"Placement Summary",false,NULL,NULL,NULL)
   verifyPlacementSummary_FYFDPI = bverifyPlacementSummary
End Function

'[Click on Deposits hyperlink in Placement Summary details table]
Public Function verifyDepositsHyperlink_FDPI(lstPlacementSummary)
	bverifyDepositsHyperlink = selectTableLink(PlacementInfo.tblPlacementSummaryHeader, PlacementInfo.tblPlacementSummaryContent,lstPlacementSummary, "Placement Summary", "Deposit No.", false, null, null, null)
	verifyDepositsHyperlink_FDPI = bverifyDepositsHyperlink
End Function

'[Verify Deposit details for the selected deposit Number in the popup displayed as]
Public Function verifyPlacementDetails_FYFDPI(lstlstPlacementdetails)
	bverifyplacementdepositdetails = true
	intSize = Ubound(lstlstPlacementdetails)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstlstPlacementdetails(Iterator),":")(0))
		arrValue = trim(Split(lstlstPlacementdetails(Iterator),":")(1))
		arrValue = Replace(arrValue,"@",":")
		Select Case (arrLabel)
			Case "Deposit Number"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblDepositNo(), arrValue, "Deposit Number")Then
			       LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "Scheme"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblSchemeName(), arrValue, "Scheme")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "Scheme Type"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblSchemeType(), arrValue, "Scheme Type")Then
			       	  LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	  bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "Premature Withdrawal"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblPrematureWithdrawl(), arrValue, "Premature Withdrawal")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If			    
			Case "Period"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblPeriod(), arrValue, "Period")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "Days"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblDays(), arrValue, "Days")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "SGD Value Rate"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblSGDValueRate(), arrValue, "SGD Value Rate")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If
			    
			Case "SGD Equivalent"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblSGDEquivalent(), arrValue, "SGD Equivalent")Then
			       	  LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	  bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "USD Value Rate"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblUSDValueRate(), arrValue, "USD Value Rate")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "USD Equivalent"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblUSDEquivalent(), arrValue, "USD Equivalent")Then
			       	   LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	   bverifyplacementdepositdetails = False
			       End If
			    End If
			Case "Principal Currency and Amount"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblPrinCurrencyAmt(), arrValue, "Principal Currency and Amount")Then
			       	  LogMessage "WARN","Verification","Deposit Placement details:"&arrValue&" is not displayed as expected",false
			       	  bverifyplacementdepositdetails = False
			       End If
			    End If		    
			Case "Maturity Instructions"
				StrtxtMI = PlacementInfo.txtMaturityInstructions.GetROProperty("innertext")
				StrActualtext= Instr(StrtxtMI, arrValue)
				If StrActualtext <> 0  Then
				   LogMessage "RSLT","Verification","Maturity Instruction details message text displayed as expected. Expected : "&arrValue&" Actual : "&StrtxtMI&" ",True
				Else 
				   LogMessage "RSLT","Verification","Maturity Instruction details message text doesnt displayed as expected. Expected : "&arrValue&" Actual : "&StrtxtMI&"",False
				   bverifyplacementdepositdetails = False
				End If
		End Select
	Next
	PlacementInfo.PopupPDbtnOK.click
	verifyPlacementDetails_FYFDPI = bverifyplacementdepositdetails
End Function

'[Click on Earmark hyperlink in Placement Summary details table]
Public Function verifyEarmarkHyperlink_FDPI(lstPlacementSummary)
	bverifyDepositsHyperlink = selectTableLink(PlacementInfo.tblPlacementSummaryHeader, PlacementInfo.tblPlacementSummaryContent,lstPlacementSummary, "Placement Summary", "Earmark", false, null, null, null)
	verifyEarmarkHyperlink_FDPI = bverifyDepositsHyperlink
End Function

'[Verify Earmark details for the selected Earmark Amount in the popup displayed as]
Public Function verifyEarmarkDetails_FYFDPI(lstEarmarkdetails)
	bverifyEarmarkDetails = true
	intSize = Ubound(lstEarmarkdetails)
	For Iterator = 0 To intSize Step 1
		arrLabel = trim(Split(lstEarmarkdetails(Iterator),":")(0))
		arrValue = trim(Split(lstEarmarkdetails(Iterator),":")(1))
		Select Case (arrLabel)
			Case "Deposit Number"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblDepositNumber(), arrValue, "Deposit Number")Then
			       LogMessage "WARN","Verification","Earmark Placement details is not displayed as expected. Expected:"&arrValue&"",false
			       bverifyEarmarkDetails=false
			       End If
			    End If
			Case "Principal Amount"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblPrincipalAmt(), arrValue, "Principal Amount")Then
			       	    LogMessage "WARN","Verification","Earmark Placement details is not displayed as expected. Expected:"&arrValue&"",false
			           bverifyEarmarkDetails=false
			       End If
			    End If
			Case "MaxEarmark Amount"
				If Not IsNull(arrValue) Then
			       If Not VerifyInnerText (PlacementInfo.lblMaxEarmarkAmt(), arrValue, "Max. Earmark Amount")Then
			       	   LogMessage "WARN","Verification","Earmark Placement details is not displayed as expected. Expected:"&arrValue&"",false
			           bverifyEarmarkDetails=false
			       End If
			    End If
		End Select
	Next
	verifyEarmarkDetails_FYFDPI = bverifyEarmarkDetails
End Function

'[Verify Earmark table details for the selected Earmark Amount in the popup displayed as]
Public Function verifyEarmarktabledetails_FYFDPI(lstlstEarmarkdetails)
   bverifyEarmarkdetails = True 
   If Not IsNull(lstlstEarmarkdetails) Then
   	   bverifyEarmarkdetails=verifyTableContentList(PlacementInfo.tblEarmarkDetailsHeader, PlacementInfo.tblEarmarkDetailsContent,lstlstEarmarkdetails,"Earmark Details",false,NULL,NULL,NULL)
   End If
   PlacementInfo.PopupEMbtnOK.click
   verifyEarmarktabledetails_FYFDPI = bverifyEarmarkdetails
End Function
