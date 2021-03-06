'[Verify row Data in Table for Fund Holdings in UT Holdings Page]
Public Function verifytblContentFundHoldingsUTHoldPage(arrFundHoldRowDataList)

   verifytblContentFundHoldingsUTHoldPage=verifyTableContentList(SG_UT_Holdings_Page.tblFundHoldingsHeader(),SG_UT_Holdings_Page.tblFundHoldingsContent(),arrFundHoldRowDataList,"Fund Holdings - UT Holdings",false,NULL,NULL,NULL)
End Function

'[Click On Fund Code in Table for Fund Holdings in UT Holdings Page]
Public Function clickFundCodeInUTHoldings(strFundCodeDetails)
	Dim bclickFundCodeInUTHoldings:bclickFundCodeInUTHoldings = True
	bAccountDetails = selectTableLink(SG_UT_Holdings_Page.tblFundHoldingsHeader(),SG_UT_Holdings_Page.tblFundHoldingsContent(),strFundCodeDetails,"UT Holdings" ,"Fund Code",False,NULL,NULL,NULL)
	WaitForICallLoading
		
	clickFundCodeInUTHoldings = bclickFundCodeInUTHoldings
End Function
'[Verify text fund code is displayed in UT Holdings Page]
Public Function VerifyTextFundCode(strFundcodeValue)
	blnVerifyTextFundCode = true
	If Not IsNull(strFundcodeValue) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleFundCode(),strFundcodeValue, "Fund Code") Then
			blnVerifyTextFundCode = false
		End If
	End If
	VerifyTextFundCode = blnVerifyTextFundCode
End Function
'[Verify text fund name is displayed in UT Holdings Page]
Public Function VerifyTextFundName(strFundNameValue)
	blnVerifyTextFundName = true
	If Not IsNull(strFundNameValue) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleFundName(),strFundNameValue, "Fund Name") Then
			blnVerifyTextFundName = false
		End If
	End If
	VerifyTextFundName = blnVerifyTextFundName
End Function
'[Verify text fund currency is displayed in UT Holdings Page]
Public Function VerifyTextFundCurrency(strFundCurrencyValue)
	blnVerifyTextFundCurrency = true
	If Not IsNull(strFundCurrencyValue) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleFundCurrency(),strFundCurrencyValue, "Fund Currency") Then
			blnVerifyTextFundCurrency = false
		End If
	End If
	VerifyTextFundCurrency = blnVerifyTextFundCurrency
End Function
'[Verify text Available Holdings is displayed in UT Holdings Page]
Public Function VerifyTextAvailableHoldings(strAvailableHoldingsValue)
	blnVerifyTextAvailableHoldings = true
	If Not IsNull(strAvailableHoldingsValue) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAvailableHoldings(),strAvailableHoldingsValue, "Available Holdings(Units)") Then
			blnVerifyTextAvailableHoldings = false
		End If
	End If
	VerifyTextAvailableHoldings = blnVerifyTextAvailableHoldings
End Function
'[Verify text Net Asset Value is displayed in UT Holdings Page]
Public Function VerifyTextNetAssetValue(strNAV)
	blnVerifyTextNetAssetValue = true
	If Not IsNull(strNAV) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleNAV(),strNAV, "Net Asset Value") Then
			blnVerifyTextNetAssetValue = false
		End If
	End If
	VerifyTextNetAssetValue = blnVerifyTextNetAssetValue
End Function
'[Verify text NAV Date is displayed in UT Holdings Page]
Public Function VerifyTextNAVDate(strNAVDate)
	blnVerifyTextNAVDate = true
	If Not IsNull(strNAVDate) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleNAVDate(),strNAVDate, "NAV Date") Then
			blnVerifyTextNAVDate = false
		End If
	End If
	VerifyTextNAVDate = blnVerifyTextNAVDate
End Function
'[Verify text Settlement Mode is displayed in UT Holdings Page]
Public Function VerifySettlementMode(strSettlementMode)
	blnVerifySettlementMode = true
	If Not IsNull(strSettlementMode) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleSettlementMode(),strSettlementMode, "Settlement Mode") Then
			blnVerifySettlementMode = false
		End If
	End If
	VerifySettlementMode = blnVerifySettlementMode
End Function
'[Verify text Market Value Of Total Holdings is displayed in UT Holdings Page]
Public Function VerifyMarketValueOfTotalHoldings(strMarketValueOfTotalHoldings)
	blnVerifyMarketValueOfTotalHoldings = true
	If Not IsNull(strMarketValueOfTotalHoldings) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleMarketValueOfTotalHoldings(),strMarketValueOfTotalHoldings, "Market Value Of Total Holdings") Then
			blnVerifyMarketValueOfTotalHoldings = false
		End If
	End If
	VerifyMarketValueOfTotalHoldings = blnVerifyMarketValueOfTotalHoldings
End Function
'[Verify text Market Value in SGD is displayed in UT Holdings Page]
Public Function VerifyMarketValueinSGD(strMarketValueinSGD)
	blnVerifyMarketValueinSGD = true
	If Not IsNull(strMarketValueinSGD) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleMarketValueinSGD(),strMarketValueinSGD, "Market Value in SGD") Then
			blnVerifyMarketValueinSGD = false
		End If
	End If
	VerifyMarketValueinSGD = blnVerifyMarketValueinSGD
End Function
'[Verify Cost Price in SGD is displayed in UT Holdings Page]
Public Function VerifyCostPrice(strCostPrice)
	blnVerifyCostPrice = true
	If Not IsNull(strCostPrice) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleCostPrice(),strCostPrice, "Cost Price") Then
			blnVerifyCostPrice = false
		End If
	End If
	VerifyCostPrice = blnVerifyCostPrice
End Function
'[Verify Unrealized Profit is displayed in UT Holdings Page]
Public Function VerifyUnrealizedProfit(strUnrealizedProfit)
	blnVerifyUnrealizedProfit = true
	If Not IsNull(strCostPrice) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleUnrealized(),strUnrealizedProfit, "Cost Price") Then
			blnVerifyUnrealizedProfit = false
		End If
	End If
	VerifyUnrealizedProfit = blnVerifyUnrealizedProfit
End Function
'[Verify Dealing CutOff for Valuation is displayed in UT Holdings Page]
Public Function VerifyDealingCutOffforValuation(strDealingCutOff)
	blnVerifyDealingCutOffforValuation = true
	If Not IsNull(strDealingCutOff) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleDealing(),strDealingCutOff, "Dealing Cut-Off for Valuation") Then
			blnVerifyDealingCutOffforValuation = false
		End If
	End If
	VerifyDealingCutOffforValuation = blnVerifyDealingCutOffforValuation
End Function

'[Verify row Data in Table for transcation history in UT Transcation History Page]
Public Function verifytblContentTranHistoryUTTranHistPage(arrTranHistoryRowDataList)

   verifytblContentTranHistoryUTTranHistPage=verifyTableContentList(SG_UT_Holdings_Page.weleTranHistinUTHeader(),SG_UT_Holdings_Page.weleTranHistinUTContent(),arrTranHistoryRowDataList,"Transcation History - UT",false,NULL,NULL,NULL)
End Function
'[Click on Transcation History link in UT Page]
Public Function ClickTranHistLinkUTPage()
 
   SG_UT_Holdings_Page.lnkTranHistoryUT().click
   WaitForICallLoading
   If Err.Number<>0 Then
       ClickTranHistLinkUTPage=false
            LogMessage "WARN","Verification","Failed to Click Link : Transcation History" ,false
       Exit Function
   End If
   ClickTranHistLinkUTPage=true
End Function
'[Click on Regular Savings Plan link in UT Page]
Public Function ClickRSPUTPage()
 
   SG_UT_Holdings_Page.lnkRSP().click
   WaitForICallLoading
   If Err.Number<>0 Then
       ClickRSPUTPage=false
            LogMessage "WARN","Verification","Failed to Click Link : Regular Savings Plan" ,false
       Exit Function
   End If
   ClickRSPUTPage=true
End Function
'[Verify row Data in Table for Regular Savings Plan in UT Transcation History Page]
Public Function verifytblContentRSPUTPage(arrRegularSavingsPlanRowDataList)

   verifytblContentRSPUTPage=verifyTableContentList(SG_UT_Holdings_Page.weleRSPHeader(),SG_UT_Holdings_Page.weleRSPContent(),arrRegularSavingsPlanRowDataList,"Regular Savings Plan - UT",false,NULL,NULL,NULL)
End Function
'[Click on Dividend Details link in UT Page]
Public Function ClickDividendDetails()
 
   SG_UT_Holdings_Page.lnkDividendDetails().click
   WaitForICallLoading
   If Err.Number<>0 Then
       ClickDividendDetails=false
            LogMessage "WARN","Verification","Failed to Click Link : Dividend Details" ,false
       Exit Function
   End If
   ClickDividendDetails=true
End Function
'[Verify row Data in Table for Dividend Details in Page]
Public Function verifytblContentDividendDetails(arrDividendDetailsRowDataList)

   verifytblContentDividendDetails=verifyTableContentList(SG_UT_Holdings_Page.tblDividendDetailsHeader(),SG_UT_Holdings_Page.tblDividendDetailsContent(),arrDividendDetailsRowDataList,"Dividend Details - UT",false,NULL,NULL,NULL)
End Function
'[Click on Key Info link in UT Page]
Public Function ClickKeyInfoInUTPage()
 
   SG_UT_Holdings_Page.lnkKeyInfoUTPage().click
   WaitForICallLoading
   If Err.Number<>0 Then
       ClickKeyInfoInUTPage=false
            LogMessage "WARN","Verification","Failed to Click Link : Key Info" ,false
       Exit Function
   End If
   ClickKeyInfoInUTPage=true
End Function

'[Click On Relationship in Table for Accound Holders in UT KeyInfo Page]
Public Function clickRelationShipUTKeyInfo(strRelationshipDetails)
	Dim bclickRelationShipUTKeyInfo:bclickRelationShipUTKeyInfo = True
	bclickRelationShipUTKeyInfo = selectTableLink(SG_UT_Holdings_Page.tblRelationshipHeader(),SG_UT_Holdings_Page.tblRelationshipContent(),strRelationshipDetails,"UT KeyInfo" ,"Relationship",False,NULL,NULL,NULL)
	WaitForICallLoading
		
	clickRelationShipUTKeyInfo = bclickRelationShipUTKeyInfo
End Function

'[Verify row Data in Table for Related Customer in Key Info Page]
Public Function verifytblContentRelatedCustomerKeyInfo(arrRelatedCustomerRowDataList)

   verifytblContentRelatedCustomerKeyInfo=verifyTableContentList(SG_UT_Holdings_Page.tblRelatedCustomersHeader(),SG_UT_Holdings_Page.tblRelatedCustomersContent(),arrRelatedCustomerRowDataList,"Related Customers - UT",false,NULL,NULL,NULL)
End Function

'[Click on Ok button in Key Info UT Page]
Public Function ClickOkBtnKeyInfoInUTPage()
 
   SG_UT_Holdings_Page.btnOkKeyInfo().click
   WaitForICallLoading
   If Err.Number<>0 Then
       ClickOkBtnKeyInfoInUTPage=false
            LogMessage "WARN","Verification","Failed to Click Link : Ok Button" ,false
       Exit Function
   End If
   ClickOkBtnKeyInfoInUTPage=true
End Function

'[Verify text Account Short Name is displayed in UT Key Info Page]
Public Function VerifyAccountShortName(strAccountShortName)
	blnVerifyAccountShortName = true
	If Not IsNull(strAccountShortName) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAccountShortName(),strAccountShortName, "Account Short Name") Then
			blnVerifyAccountShortName = false
		End If
	End If
	VerifyAccountShortName = blnVerifyAccountShortName
End Function
'[Verify text Account Signatory Type is displayed in UT Key Info Page]
Public Function VerifyAccountSignatoryType(strAccountSignatoryType)
	blnVerifyAccountSignatoryType = true
	If Not IsNull(strAccountSignatoryType) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAccountSignatoryType(),strAccountSignatoryType, "Account Signatory Type") Then
			blnVerifyAccountSignatoryType = false
		End If
	End If
	VerifyAccountSignatoryType = blnVerifyAccountSignatoryType
End Function
'[Verify text Account Type is displayed in UT Key Info Page]
Public Function VerifyAccountType(strAccountType)
	blnVerifyAccountType = true
	If Not IsNull(strAccountType) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAccountType(),strAccountType, "Account Type") Then
			blnVerifyAccountType = false
		End If
	End If
	VerifyAccountType = blnVerifyAccountType
End Function
'[Verify text Primary CIN is displayed in UT Key Info Page]
Public Function VerifyPrimaryCIN(strPrimaryCIN)
	blnVerifyPrimaryCIN = true
	If Not IsNull(strPrimaryCIN) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePrimaryCIN(),strPrimaryCIN, "Primary CIN") Then
			blnVerifyPrimaryCIN = false
		End If
	End If
	VerifyPrimaryCIN = blnVerifyPrimaryCIN
End Function
'[Verify text Account Information Status is displayed in UT Key Info Page]
Public Function VerifyAccountInformationStatus(strAccountInformationStatus)
	blnVerifyAccountInformationStatus = true
	If Not IsNull(strAccountInformationStatus) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAccountInfoStatus(),strAccountInformationStatus, "Status") Then
			blnVerifyAccountInformationStatus = false
		End If
	End If
	VerifyAccountInformationStatus = blnVerifyAccountInformationStatus
End Function
'[Verify text Opening Date is displayed in UT Key Info Page]
Public Function VerifyOpeningDate(strOpeningDate)
	blnVerifyOpeningDate = true
	If Not IsNull(strOpeningDate) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleOpeningDate(),strOpeningDate, "Opening Date") Then
			blnVerifyOpeningDate = false
		End If
	End If
	VerifyOpeningDate = blnVerifyOpeningDate
End Function
'[Verify text Closure Date is displayed in UT Key Info Page]
Public Function VerifyClosureDate(strClosureDate)
	blnVerifyClosureDate = true
	If Not IsNull(strClosureDate) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleClosureDate(),strClosureDate, "Closure Date") Then
			blnVerifyClosureDate = false
		End If
	End If
	VerifyClosureDate = blnVerifyClosureDate
End Function
'[Click on Address and Account Linkage link in UT Page]
Public Function ClickAdressAndAccountLinkageInUTPage()
 
   SG_UT_Holdings_Page.lnkAddressAndAccountLinkage().click
   WaitForICallLoading
   If Err.Number<>0 Then
       ClickAdressAndAccountLinkageInUTPage=false
            LogMessage "WARN","Verification","Failed to Click Link : Address and Account Linkage" ,false
       Exit Function
   End If
   ClickAdressAndAccountLinkageInUTPage=true
End Function
'[Verify text Adress Name is displayed in UT Address Page]
Public Function VerifyAdressName(strAddressName)
	blnVerifyAdressName = true
	If Not IsNull(strAddressName) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAddressName(),strAddressName, "Address Name") Then
			blnVerifyAdressName = false
		End If
	End If
	VerifyAdressName = blnVerifyAdressName
End Function
'[Verify text Adress Type is displayed in UT Address Page]
Public Function VerifyAdressType(strAddressType)
	blnVerifyAdressType = true
	If Not IsNull(strAddressType) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAddressType(),strAddressType, "Address Type") Then
			blnVerifyAdressType = false
		End If
	End If
	VerifyAdressType = blnVerifyAdressType
End Function
'[Verify text Adress CIN is displayed in UT Address Page]
Public Function VerifyAdressCIN(strAddressCIN)
	blnVerifyAdressCIN = true
	If Not IsNull(strAddressCIN) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAddressCIN(),strAddressCIN, "Address CIN") Then
			blnVerifyAdressCIN = false
		End If
	End If
	VerifyAdressCIN = blnVerifyAdressCIN
End Function
'[Verify text Adress Block is displayed in UT Address Page]
Public Function VerifyAdressBlock(strAddressBlock)
	blnVerifyAdressBlock = true
	If Not IsNull(strAddressBlock) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAddressBlock(),strAddressBlock, "Address Block") Then
			blnVerifyAdressBlock = false
		End If
	End If
	VerifyAdressBlock = blnVerifyAdressBlock
End Function
'[Verify text Adress Level is displayed in UT Address Page]
Public Function VerifyAdressLevel(strAddressLevel)
	blnVerifyAdressLevel = true
	If Not IsNull(strAddressLevel) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAddressLevel(),strAddressLevel, "Address Level") Then
			blnVerifyAdressLevel = false
		End If
	End If
	VerifyAdressLevel = blnVerifyAdressLevel
End Function
'[Verify text Adress Unit is displayed in UT Address Page]
Public Function VerifyAdressUnit(strAddressUnit)
	blnVerifyAdressUnit = true
	If Not IsNull(strAddressUnit) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAddressUnit(),strAddressUnit, "Address Unit") Then
			blnVerifyAdressUnit = false
		End If
	End If
	VerifyAdressUnit = blnVerifyAdressUnit
End Function
'[Verify text Adress Line1 is displayed in UT Address Page]
Public Function VerifyAdressLine1(strAdressLine1)
	blnVerifyAdressLine1 = true
	If Not IsNull(strAdressLine1) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAdressLevel1(),strAdressLine1, "Address Level 1") Then
			blnVerifyAdressLine1 = false
		End If
	End If
	VerifyAdressLine1 = blnVerifyAdressLine1
End Function
'[Verify text Adress Line2 is displayed in UT Address Page]
Public Function VerifyAdressLine2(strAdressLine2)
	blnVerifyAdressLine2 = true
	If Not IsNull(strAdressLine2) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAdressLevel2(),strAdressLine2, "Address Level 2") Then
			blnVerifyAdressLine2 = false
		End If
	End If
	VerifyAdressLine2 = blnVerifyAdressLine2
End Function
'[Verify text Adress Line3 is displayed in UT Address Page]
Public Function VerifyAdressLine3(strAdressLine3)
	blnVerifyAdressLine3 = true
	If Not IsNull(strAdressLine3) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAdressLevel3(),strAdressLine3, "Address Level 3") Then
			blnVerifyAdressLine1 = false
		End If
	End If
	VerifyAdressLine3 = blnVerifyAdressLine3
End Function
'[Verify text Adress Line4 is displayed in UT Address Page]
Public Function VerifyAdressLine4(strAdressLine4)
	blnVerifyAdressLine4 = true
	If Not IsNull(strAdressLine4) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAdressLevel4(),strAdressLine4, "Address Level 4") Then
			blnVerifyAdressLine1 = false
		End If
	End If
	VerifyAdressLine4 = blnVerifyAdressLine4
End Function
'[Verify text Adress Postal Code is displayed in UT Address Page]
Public Function VerifyAdressPostalCode(strAdressPostalCode)
	blnVerifyAdressPostalCode = true
	If Not IsNull(strAdressPostalCode) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePostalCode(),strAdressPostalCode, "Postal Code") Then
			blnVerifyAdressPostalCode = false
		End If
	End If
	VerifyAdressPostalCode = blnVerifyAdressPostalCode
End Function
'[Verify text Adress Last Updated Date is displayed in UT Address Page]
Public Function VerifyAdressLastUpdatedDate(strAdressLastUpdatedDate)
	blnVerifyAdressLastUpdatedDate = true
	If Not IsNull(strAdressLastUpdatedDate) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleLastUpdatedDate(),strAdressLastUpdatedDate, "Last Updated Date") Then
			blnVerifyAdressLastUpdatedDate = false
		End If
	End If
	VerifyAdressLastUpdatedDate = blnVerifyAdressLastUpdatedDate
End Function
'[Verify text Adress Last Updated By is displayed in UT Address Page]
Public Function VerifyAdressLastUpdatedBy(strAdressLastUpdatedBy)
	blnVerifyAdressLastUpdatedBy = true
	If Not IsNull(strAdressLastUpdatedBy) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleLastUpdatedBy(),strAdressLastUpdatedBy, "Last Updated By") Then
			blnVerifyAdressLastUpdatedBy = false
		End If
	End If
	VerifyAdressLastUpdatedBy = blnVerifyAdressLastUpdatedBy
End Function

'[Click On CIN Suffix in Table for Accound Holders in UT KeyInfo Page]
Public Function clickCIFTKeyInfo(strRelationshipDetails)
	Dim bclickCIFTKeyInfo:bclickCIFTKeyInfo = True
	bclickCIFTKeyInfo = selectTableLink(SG_UT_Holdings_Page.tblRelationshipHeader(),SG_UT_Holdings_Page.tblRelationshipContent(),strRelationshipDetails,"UT KeyInfo" ,"CIN/CIN Suffix",False,NULL,NULL,NULL)
	WaitForICallLoading
		
	clickCIFTKeyInfo = bclickCIFTKeyInfo
End Function

'[Verify text Customer Identification Number is displayed in UT Address Page]
Public Function VerifyCustomerIdentificationNumber(strCustomerIdentificationNumber)
	blnVerifyCustomerIdentificationNumber = true
	If Not IsNull(strCustomerIdentificationNumber) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleCustomerIdentificationNumber(),strCustomerIdentificationNumber, "Customer Identification Number") Then
			blnVerifyCustomerIdentificationNumber = false
		End If
	End If
	VerifyCustomerIdentificationNumber = blnVerifyCustomerIdentificationNumber
End Function
'[Verify text Customer Name is displayed in UT Address Page]
Public Function VerifyCustomerIdentificationName(strCustomerIdentificationName)
	blnVerifyCustomerIdentificationName = true
	If Not IsNull(strCustomerIdentificationName) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleCustomerIdentificationName(),strCustomerIdentificationName, "Customer Identification Name") Then
			blnVerifyCustomerIdentificationName = false
		End If
	End If
	VerifyCustomerIdentificationName = blnVerifyCustomerIdentificationName
End Function
'[Verify text Alias or Birthday is displayed in UT Address Page]
Public Function VerifyAliasOrBirthday(strAliasOrBirthday)
	blnVerifyAliasOrBirthday = true
	If Not IsNull(strAliasOrBirthday) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.weleAliasOrBirthDay(),strAliasOrBirthday, "Alias or Birthday") Then
			blnVerifyAliasOrBirthday = false
		End If
	End If
	VerifyAliasOrBirthday = blnVerifyAliasOrBirthday
End Function
'[Verify text Personal Info Salutation is displayed in UT Address Page]
Public Function VerifyPersonalInfoSalutation(strPersonalInfoSalutation)
	blnVerifyPersonalInfoSalutation = true
	If Not IsNull(strPersonalInfoSalutation) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoSalutation(),strPersonalInfoSalutation, "Personal Info Salutation") Then
			blnVerifyPersonalInfoSalutation = false
		End If
	End If
	VerifyPersonalInfoSalutation = blnVerifyPersonalInfoSalutation
End Function
'[Verify text Personal Info Date Of Birth is displayed in UT Address Page]
Public Function VerifyPersonalInfoDOB(strPersonalInfoDOB)
	blnVerifyPersonalInfoDOB = true
	If Not IsNull(strPersonalInfoDOB) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoDOB(),strPersonalInfoDOB, "Personal Info Date Of Birth") Then
			blnVerifyPersonalInfoDOB = false
		End If
	End If
	VerifyPersonalInfoDOB = blnVerifyPersonalInfoDOB
End Function
'[Verify text Personal Ethinic Type is displayed in UT Address Page]
Public Function VerifyPersonalInfoEthinicType(strPersonalInfoEthinicType)
	blnVerifyPersonalInfoEthinicType = true
	If Not IsNull(strPersonalInfoEthinicType) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoEthinicType(),strPersonalInfoEthinicType,"Personal Ethinic Type") Then
			blnVerifyPersonalInfoEthinicType = false
		End If
	End If
	VerifyPersonalInfoEthinicType = blnVerifyPersonalInfoEthinicType
End Function
'[Verify text Personal Info Sex is displayed in UT Address Page]
Public Function VerifyPersonalInfoSex(strPersonalInfoSex)
	blnVerifyPersonalInfoSex = true
	If Not IsNull(strPersonalInfoSex) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoSex(),strPersonalInfoSex,"Personal Info Sex") Then
			blnVerifyPersonalInfoSex = false
		End If
	End If
	VerifyPersonalInfoSex = blnVerifyPersonalInfoSex
End Function
'[Verify text Personal Info Marital Status is displayed in UT Address Page]
Public Function VerifyPersonalInfoMaritalStatus(strPersonalInfoMaritalStatus)
	blnVerifyPersonalInfoMaritalStatus = true
	If Not IsNull(strPersonalInfoMaritalStatus) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoMaritalStatus(),strPersonalInfoMaritalStatus,"Personal Info Marital Status") Then
			blnVerifyPersonalInfoMaritalStatus = false
		End If
	End If
	VerifyPersonalInfoMaritalStatus = blnVerifyPersonalInfoMaritalStatus
End Function
'[Verify text Personal Info Nationality is displayed in UT Address Page]
Public Function VerifyPersonalInfoNationality(strPersonalInfoNationality)
	blnVerifyPersonalInfoNationality = true
	If Not IsNull(strPersonalInfoNationality) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoNationality(),strPersonalInfoNationality,"Personal Info Nationality") Then
			blnVerifyPersonalInfoNationality = false
		End If
	End If
	VerifyPersonalInfoNationality = blnVerifyPersonalInfoNationality
End Function
'[Verify text Personal Info Country of Residence is displayed in UT Address Page]
Public Function VerifyPersonalInfoCountryofResidence(strPersonalInfoCountryofResidence)
	blnVerifyPersonalInfoCountryofResidence = true
	If Not IsNull(strPersonalInfoCountryofResidence) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoCountryofResidence(),strPersonalInfoCountryofResidence,"Personal Info Country of Residence") Then
			blnVerifyPersonalInfoCountryofResidence = false
		End If
	End If
	VerifyPersonalInfoCountryofResidence = blnVerifyPersonalInfoCountryofResidence
End Function
'[Verify text Personal Info Employer Name is displayed in UT Address Page]
Public Function VerifyPersonalInfoEmployerName(strPersonalInfoEmployerName)
	blnVerifyPersonalInfoEmployerName = true
	If Not IsNull(strPersonalInfoCountryofResidence) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoEmployerName(),strPersonalInfoEmployerName,"Personal Info Employer Name") Then
			blnVerifyPersonalInfoEmployerName = false
		End If
	End If
	VerifyPersonalInfoEmployerName = blnVerifyPersonalInfoEmployerName
End Function
'[Verify text Personal Info Occupation is displayed in UT Address Page]
Public Function VerifyPersonalInfoOccupation(strPersonalInfoOccupation)
	blnVerifyPersonalInfoOccupation = true
	If Not IsNull(strPersonalInfoOccupation) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoOccupation(),strPersonalInfoOccupation,"Personal Info Employer Name") Then
			blnVerifyPersonalInfoOccupation = false
		End If
	End If
	VerifyPersonalInfoOccupation = blnVerifyPersonalInfoOccupation
End Function
'[Verify text Personal Info Segment is displayed in UT Address Page]
Public Function VerifyPersonalInfoSegment(strPersonalInfoSegment)
	blnVerifyPersonalInfoSegment = true
	If Not IsNull(strPersonalInfoSegment) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoSegment(),strPersonalInfoSegment,"Personal Info Segment") Then
			blnVerifyPersonalInfoSegment = false
		End If
	End If
	VerifyPersonalInfoSegment = blnVerifyPersonalInfoSegment
End Function
'[Verify text Personal Info CDP Account No is displayed in UT Address Page]
Public Function VerifyPersonalInfoCDPAccountNo(strPersonalInfoCDPAccountNo)
	blnVerifyPersonalInfoCDPAccountNo = true
	If Not IsNull(strPersonalInfoCDPAccountNo) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welePersonalInfoCDPAccountNo(),strPersonalInfoCDPAccountNo,"Personal Info CDP Account No") Then
			blnVerifyPersonalInfoCDPAccountNo = false
		End If
	End If
	VerifyPersonalInfoCDPAccountNo = blnVerifyPersonalInfoCDPAccountNo
End Function
'[Verify text MP SMSMMS is displayed in UT Page]
Public Function VerifyMPSMSMMS(strMPSmsMms)
	blnVerifyMPSMSMMS = true
	If Not IsNull(strMPSmsMms) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welestrMPSmsMms(),strMPSmsMms,"MP SMSMMS") Then
			blnVerifyMPSMSMMS = false
		End If
	End If
	VerifyMPSMSMMS = blnVerifyMPSMSMMS
End Function
'[Verify text MP PhoneMobile is displayed in UT Page]
Public Function VerifyMPPhoneMobile(strMPPhoneMobile)
	blnVerifyMPPhoneMobile = true
	If Not IsNull(strMPPhoneMobile) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welestrMPPhoneMobile(),strMPPhoneMobile,"MP PhoneMobile") Then
			blnVerifyMPPhoneMobile = false
		End If
	End If
	VerifyMPPhoneMobile = blnVerifyMPPhoneMobile
End Function
'[Verify text MP Fax is displayed in UT Page]
Public Function VerifyMPFax(strMPFax)
	blnVerifyMPFax = true
	If Not IsNull(strMPFax) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welestrMPFax(),strMPFax,"MP Fax") Then
			blnVerifyMPFax = false
		End If
	End If
	VerifyMPFax = blnVerifyMPFax
End Function
'[Verify text MP Email is displayed in UT Page]
Public Function VerifyMPEmail(strMPEmail)
	blnVerifyMPEmail = true
	If Not IsNull(strMPEmail) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welestrMPEmail(),strMPEmail,"MP Email") Then
			blnVerifyMPEmail = false
		End If
	End If
	VerifyMPEmail = blnVerifyMPEmail
End Function
'[Verify text MP Direct Sales is displayed in UT Page]
Public Function VerifyMPDirectSales(strMPDirectSales)
	blnVerifyMPDirectSales = true
	If Not IsNull(strMPDirectSales) Then
		If Not VerifyInnerText (SG_UT_Holdings_Page.welestrMPDirectSales(),strMPDirectSales,"MP Direct Sales") Then
			blnVerifyMPDirectSales = false
		End If
	End If
	VerifyMPDirectSales = blnVerifyMPDirectSales
End Function
