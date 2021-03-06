'********************************* Scrips developed as part of IR1603 by  on 07th July 2016***************************************

'[Verify header name displayed for Base Summary table]
Public Function verifylabelname_WealthPortfolio(strBaseNumber)
  bverifylabelname_WealthPortfolio = True	
  WaitForICallLoading
	labelname = "Base No : "
	strExpBaseNumber  = labelname&strBaseNumber
	If not IsNull(strBaseNumber) Then 
	  Wait(10)
 	  If Not VerifyInnerText (Wealth.lblBaseNumber(), strExpBaseNumber, "Base No:"&strBaseNumber)Then
 	     bverifylabelname_WealthPortfolio=false
      End If
    End IF
	verifylabelname_WealthPortfolio = bverifylabelname_WealthPortfolio
End Function

'[Verify header name displayed for Portfolio Summary table]
Public Function verifyheaderPortfolioSummary_Wealth(strPortfolioNumber)
  verifylabelname_PortfolioSummary = True
	labelname = "Portfolio Summary : "
	strExplabelName  = labelname&strPortfolioNumber
 	  If Not VerifyInnerText (Wealth.lblPortfolioSummary(), strExplabelName, "Portfolio Summary:"&strBaseNumber)Then
 	     verifylabelname_PortfolioSummary=false
      End If
	verifyheaderPortfolioSummary_Wealth = verifylabelname_PortfolioSummary
End Function

'[Verify Base summray details table displayed]
Public Function VerifyBaseSummary_Wealth(StrProtfolioNumber, lstlstPortfolioSummary)
	bVerifyBaseSummary = True 
	bverifyPortfolioSummarytable=verifyTableContentList(Wealth.tblProtfolioHeader,Wealth.tblProtfolioContent,lstlstPortfolioSummary,"Base Summary table",false,null,null,null)   
	If bverifyPortfolioSummarytable = True Then
	clickPortfolioNumber=selectTableLink(Wealth.tblProtfolioHeader,Wealth.tblProtfolioContent,Array("Portfolio #:"&StrProtfolioNumber),"Base Summary table","Portfolio #",true,null,null,null)	    
	Else 
	bVerifyBaseSummary = False 
	End If
	VerifyBaseSummary_Wealth = bVerifyBaseSummary
End Function

'[Verify Portfolio Holding details for Derviatives Asset types selected from Portfolio Summary table]
Public Function VerifyDerivativesAssetTypesPHD_Wealth(strAssetType, lstlstDerivatives,lstPortfolioSummary)
 bverifyPHD = True 
 intRCPortfolioSummary = getRecordsCountForColumn(Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,"Asset Type")
	For j = 0 To intRCPortfolioSummary - 1
		Set objAllRowsSummary = getAllRows(Wealth.tblProtfolioSummaryContent)
			strActAssetType = getCellTextFor(Wealth.tblProtfolioSummaryHeader,objAllRowsSummary(j),j, "Asset Type") 
	 	If Trim(strActAssetType) = Trim(strAssetType) Then
	   		'Call clickVaddinLink_tblCell (Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,j,"Asset Type")
	   		bverifyPortfolioSummary = selectTableLink(Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,lstPortfolioSummary,"Portfolio Summary",strAssetType,false,null,null,null)
	   		If bverifyPortfolioSummary = True Then
				If not IsNull(lstlstDerivatives) Then 		
		   		   verifytablePHD_Derivatives=verifyTableContentList(Wealth.tblPHDHeader,Wealth.tblPHDContent,lstlstDerivatives,"Portfolio Holiding details",false,null,null,null)   				
		   		   If verifytablePHD_Derivatives = False Then
		   		   	   bverifyPHD = False  
		   		   	   Exit Function 
		   		   Else
		   		   	   Wealth.btnOK.click
		   		   	   Exit For
		   		   End If
 	   			End If	   			
	   		End If
	   	End If
	Next 
	VerifyDerivativesAssetTypesPHD_Wealth = bverifyPHD
End Function

'[Verify Portfolio Holding details for Blank Columns to be displayed on selected Asset types]
Public Function verifyBlankColumnAssetType_Wealth(StrAssetType,StrAssetSubType)
	bverifyBlankColumns = True
	Dim j
	Dim i
	intRecordCount = getRecordsCountForColumn(Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,"Asset Type")	
    For i = 0 To intRecordCount - 1
    	j = 1
		Set objAllRows=getAllRows(Wealth.tblProtfolioSummaryContent)
		ActAssetType=getCellTextFor(Wealth.tblProtfolioSummaryHeader,objAllRows(i),i,"Asset Type")
		ActRefCurrency=getCellTextFor(Wealth.tblProtfolioSummaryHeader,objAllRows(i),j,"Amount In Reference Currency")
		
		If StrAssetType = ActAssetType  Then	
		Call selectTableLink(Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,Array("Asset Type:"&strAssetType),"Portfolio Summary",strAssetType,false,null,null,null)
		K =1 
			Set objAllRowsPHD = getAllRows(Wealth.tblPHDContent)
			strCurrency = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Currency")
			strAccount = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Account #")
			strDescription = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Description")
			strQuantity = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Quantity")
			strCostPrice = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Cost Price")
			strMarketPrice = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Market Price")
			strLastPriceDate = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Last Price Date")
			strMarketValue = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Market Value")
			strUnrealizedPL = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Unrealized P/L")
			strAmtOriginalCur = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Amount In Original Currency")
			strAccuredInterest = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Accrued Interest")
			strAmtReferenceCur = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"Amount In Reference Currency")
			StrTotalCashInvestment = getCellTextFor(Wealth.tblPHDHeader,objAllRowsPHD(i),K,"% of Total Cash and Cash Investment")
			
			If strAssetType = "Cash and Cash Investment" Then
				If StrAssetSubType =  "Cash (MCSA)" Then		
					If strDescription = "" And strQuantity = "" And strCostPrice = "" And strMarketPrice = "" And strLastPriceDate = "" And strMarketValue ="" And strUnrealizedPL = "" Then
					LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" -"&StrAssetSubType&" are displayed as blank as expected", True
					else
				  	LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" -"&StrAssetSubType&" are not displayed as blank as expected", False
				  	bverifyBlankColumns = False
				  	Exit Function
					End If	
	 			ElseIf StrAssetSubType = "Deposit"  Then		
					If strAccount = "" And strDescription = "" And strQuantity = "" And strCostPrice = "" And strMarketPrice = "" And strLastPriceDate = "" And strMarketValue ="" And strUnrealizedPL = "" Then
					LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" -"&StrAssetSubType&" are displayed as blank as expected", True
					else
				  	LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" -"&StrAssetSubType&" are not displayed as blank as expected", False
				  	bverifyBlankColumns = False
				  	Exit Function
					End If	
				ElseIf StrAssetSubType = "Cash Investment"  Then		
					If strAccount = "" And strDescription = "" And strQuantity = "" And strCostPrice = "" And strMarketPrice = "" And strLastPriceDate = ""  And strMarketValue ="" And strUnrealizedPL = "" And strAmtOriginalCur = "" Then
					LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" -"&StrAssetSubType&" are displayed as blank as expected", True
					else
				  	LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" -"&StrAssetSubType&" are not displayed as blank as expected", False
				  	bverifyBlankColumns = False
				  	Exit Function
					End If	
				End If
			End If
		
			If strAssetType = "Equity" OR strAssetType = "Fixed Income" OR strAssetType = "Fund" OR strAssetType = "FX" OR strAssetType = "Loan" OR strAssetType = "Fixed Income Derivatives" OR strAssetType = "FX Derivatives" Then
				If strAccount = "" And strAmtOriginalCur ="" And StrTotalCashInvestment = "" And strAccuredInterest = "" Then
				LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" are displayed as blank as expected", True
				else
			  	LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" are not displayed as blank as expected", False
			  	bverifyBlankColumns = False
			  	Exit Function
				End If	
			End IF
		
			If strAssetType = "Equity Derivatives" Then
				If strAccount = "" And strQuantity = "" And strCostPrice = "" And strMarketPrice = "" And strLastPriceDate = "" ANd strMarketValue ="" And strUnrealizedPL = "" And StrTotalCashInvestment = "" Then
				LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" are displayed as blank as expected", True
				else
			  	LogMessage "RSLT", "Verification", "All the columns in Portfolio Holding details table for "&strAssetType&" are not displayed as blank as expected", False
			  	bverifyBlankColumns = False
			  	Exit Function
				End If	
			End IF	 
    	End IF	
    Next 
    verifyBlankColumnAssetType_Wealth = bverifyBlankColumns
End Function

'[Verify Portfolio Holding details for Cash Investment asset type selected from Portfolio Summary table]
Public Function VerifyPHDCashInvestment_Wealth(strAssetType,lstPortfolioSummary,lstlstCashInvestment, lstlstDeposits, lstlstOtherCashInvestment)
 bVerifyPHDCashInvestment = True 
 intRCPortfolioSummary = getRecordsCountForColumn(Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,"Asset Type")
	For j = 0 To intRCPortfolioSummary - 1
		Set objAllRowsSummary = getAllRows(Wealth.tblProtfolioSummaryContent)
			strActAssetType = getCellTextFor(Wealth.tblProtfolioSummaryHeader,objAllRowsSummary(j),j, "Asset Type") 
	 	If Trim(strActAssetType) = Trim(strAssetType) Then
	   		bverifyPortfolioSummary = selectTableLink(Wealth.tblProtfolioSummaryHeader,Wealth.tblProtfolioSummaryContent,lstPortfolioSummary,"Portfolio Summary",strAssetType,false,null,null,null)
	   		If bverifyPortfolioSummary = True Then
				If not IsNull(lstlstCashInvestment) Then
	   			   verifytablePHD_Cash=verifyTableContentList(Wealth.tblPHDHeader,Wealth.tblPHDContent,lstlstCashInvestment,"Cash Investments",false,null,null,null)   				
	   			End If
	   			If not IsNull(lstlstDeposits) Then
	   			 'select radio button Deposit
		   			selectRadioButton_Wealth=SelectRadioButtonGrp("Deposit",Wealth.rbtnGroup_ProfileHolding, Array("Cash (MCSA)","Deposit", "Cash Investments"))   
		   			verifytablePHD_Deposits=verifyTableContentList(Wealth.tblPHDHeader,Wealth.tblPHDContent,lstlstDeposits,"Deposits",false,null,null,null)   				
	   			End If 			   		
	   			If lstlstOtherCashInvestment <> "" Then
	   			'select radio button Cash Investment
		   			selectRadioButton_Wealth=SelectRadioButtonGrp("Cash Investments",Wealth.rbtnGroup_ProfileHolding, Array("Cash (MCSA)","Deposit", "Cash Investments"))   
		   			verifytablePHD_OtherCash=verifyTableContentList(Wealth.tblPHDHeader,Wealth.tblPHDContent,lstlstOtherCashInvestment,"Other Cash Investments",false,null,null,null)   				
	   			End If 
	   			Wealth.btnOK.click
 	   		End If	   			
	   	End If
	Next 
VerifyPHDCashInvestment_Wealth = bVerifyPHDCashInvestment
End Function

'[Verify error message displayed in wealth Page]
Public Function verifyErrorMessage_Wealth(strErrorMessage)
  bErrorMessage = True 
  If Not VerifyInnerText (Wealth.ErrorMsg(), strErrorMessage, "Error Message") Then
     bErrorMessage = false
  End If
 verifyErrorMessage_Wealth = bErrorMessage
End Function

'[Verify Wealth Product displayed in the Account Overview table only for valid ProductCode]
Public Function verifyProductAccountOverview_Wealth(strProductName)
    bflagProductFound = True   
	Set ListProductnames = Description.Create
	ListProductnames("xpath").value = "//div[contains(@class,'dt-group-row')]"
	Set ProductList = cCustomerOverview.tblProductsListContent().ChildObjects(ListProductnames)
	ctListofProducts = ProductList.Count
	For it = 0 To ctListofProducts - 1 Step 1
		strActualProduct = ProductList(it).GetRoProperty("innertext")
		If Trim(strActualProduct) = Trim(strProductName) Then
			LogMessage "WARN","Verification","Wealth Product found in the Accounts Overview section for Product Code not 100", False
			bflagProductFound = false
		End If
	Next	
	verifyProductAccountOverview_Wealth = bflagProductFound
End Function
