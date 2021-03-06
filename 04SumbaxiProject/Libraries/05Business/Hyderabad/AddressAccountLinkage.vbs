'[Verify the fields of address and Account Linkage]
	Public Function verifyfields_AddressAcctLinkage(strProductCode,strCardNumber,strName,strBlock,strLevelUnit,strUnit,strAddressLine1,strAddressLine2,strAddressLine3,strAddressLine4,strPostalCode,strLastUpdatedDate)
	bverifyfields_AddressAcctLinkage=true
  If Not IsNull(strName) Then
		If strName = "RUNTIME" Then
			getAccountInfo_CA strProductCode,strCardNumber				
			strCISName=strRunTimeShortName
    		strIServeName=bcVerify_AccountAndAddress.lblName.getroproperty("innertext")
    		If  Ucase(Trim(strCISName)) = UCase(Trim(strIServeName)) Then
    			LogMessage "RSLT", "Verification","For Name value successfully matched with the expected value. Expected: "+ strCISName &" , Actual: "& strIServeName, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Name value not matching with the expected value. Expected: "+ strCISName &" , Actual: "& strIServeName, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblName(),strName,"Name")Then
				bverifyfields_AddressAcctLinkage = False
			End If
    	End If
    End If	
    'commented by Aniket
'    If Not IsNull(strAddressType) Then
'    	If strAddressType = "RUNTIME" Then
'			strCISAddressType=strRunTimeAddressType
'    		strIServeAddressType=bcVerify_AccountAndAddress.lblAddressType.getroproperty("innertext")
'    		If  Ucase(Trim(strCISAddressType)) = UCase(Trim(strIServeAddressType)) Then
'    			LogMessage "RSLT", "Verification","For Address Type value successfully matched with the expected value. Expected: "+ strCISAddressType &" , Actual: "& strIServeAddressType, True
'				bverifyfields_AddressAcctLinkage = True
'			else
'				LogMessage "WARN", "Verification","For Address Type value not matching with the expected value. Expected: "+ strCISAddressType &" , Actual: "& strIServeAddressType, False
'				bverifyfields_AddressAcctLinkage = False
'			End If
'		Else
'			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressType(),strAddressType,"Address Type")Then
'				bverifyfields_AddressAcctLinkage = False
'			End If
'		End If	
'    End If
'    If Not IsNull(strAddressCIN) Then
'    	If strAddressCIN = "RUNTIME" Then
'    		getProductInfo_CA (strCIN)
'			strCISAddressCIN=strRunTimeAddressCIN
'    		strIServeAddressCIN=bcVerify_AccountAndAddress.lblAddressCIN.getroproperty("innertext")
'    		If  Ucase(Trim(strCISAddressCIN)) = UCase(Trim(strCISAddressCIN)) Then
'    			LogMessage "RSLT", "Verification","For Address CIN value successfully matched with the expected value. Expected: "+ strCISAddressCIN &" , Actual: "& strIServeAddressCIN, True
'				bverifyfields_AddressAcctLinkage = True
'			else
'				LogMessage "WARN", "Verification","For Address CIN value not matching with the expected value. Expected: "+ strCISAddressCIN &" , Actual: "& strIServeAddressCIN, False
'				bverifyfields_AddressAcctLinkage = False
'			End If
'		Else
'			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressCIN(),strAddressCIN,"Address CIN")Then
'				bverifyfields_AddressAcctLinkage = False
'			End If
'		End If	
'    End If
    If Not IsNull(strBlock) Then
    	If strBlock = "RUNTIME" Then
			strCISBlock=strRunTimeBlock
    		strIServeBlock=bcVerify_AccountAndAddress.lblBlock.getroproperty("innertext")
    		If  Ucase(Trim(strCISBlock)) = UCase(Trim(strIServeBlock)) Then
    			LogMessage "RSLT", "Verification","For Block value successfully matched with the expected value. Expected: "+ strCISBlock &" , Actual: "& strIServeBlock, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Block value not matching with the expected value. Expected: "+ strCISBlock &" , Actual: "& strIServeBlock, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblBlock(),strBlock,"Block")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strLevelUnit) Then
    	If strLevelUnit = "RUNTIME" Then
			strCISLevel=strRunTimeLevel
    		strIServeLevel=bcVerify_AccountAndAddress.lblLevelUnit.getroproperty("innertext")
    		If  Ucase(Trim(strCISLevel)) = UCase(Trim(strIServeLevel)) Then
    			LogMessage "RSLT", "Verification","For Level value successfully matched with the expected value. Expected: "+ strCISLevel &" , Actual: "& strIServeLevel, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Level value not matching with the expected value. Expected: "+ strCISLevel &" , Actual: "& strIServeLevel, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblLevelUnit(),strLevelUnit,"Level")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strUnit) Then
    	If strUnit = "RUNTIME" Then
			strCISUnit=strRunTimeUnit
    		strIServeUnit=bcVerify_AccountAndAddress.lblUnit.getroproperty("innertext")
    		If  Ucase(Trim(strCISUnit)) = UCase(Trim(strIServeUnit)) Then
    			LogMessage "RSLT", "Verification","For Unit value successfully matched with the expected value. Expected: "+ strCISUnit &" , Actual: "& strIServeUnit, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Unit value not matching with the expected value. Expected: "+ strCISUnit &" , Actual: "& strIServeUnit, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblUnit(),strUnit,"Unit")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine1) Then
    	If strAddressLine1 = "RUNTIME" Then
			strCISAddressLine1=strRunTimeAddressLine1
    		strIServeAddressLine1=bcVerify_AccountAndAddress.lblAddressLine1.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine1)) = UCase(Trim(strIServeAddressLine1)) Then
    			LogMessage "RSLT", "Verification","For Address Line1 value successfully matched with the expected value. Expected: "+ strCISAddressLine1 &" , Actual: "& strIServeAddressLine1, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line1 value not matching with the expected value. Expected: "+ strCISAddressLine1 &" , Actual: "& strIServeAddressLine1, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine1(),strAddressLine1,"Address Line1")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine2) Then
    	If strAddressLine2 = "RUNTIME" Then
			strCISAddressLine2=strRunTimeAddressLine2
    		strIServeAddressLine2=bcVerify_AccountAndAddress.lblAddressLine2.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine2)) = UCase(Trim(strIServeAddressLine2)) Then
    			LogMessage "RSLT", "Verification","For Address Line2 value successfully matched with the expected value. Expected: "+ strCISAddressLine2 &" , Actual: "& strIServeAddressLine2, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line2 value not matching with the expected value. Expected: "+ strCISAddressLine2 &" , Actual: "& strIServeAddressLine2, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine2(),strAddressLine2,"Address Line2")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine3) Then
    	If strAddressLine3 = "RUNTIME" Then
			strCISAddressLine3=strRunTimeAddressLine3
    		strIServeAddressLine3=bcVerify_AccountAndAddress.lblAddressLine3.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine3)) = UCase(Trim(strIServeAddressLine3)) Then
    			LogMessage "RSLT", "Verification","For Address Line3 value successfully matched with the expected value. Expected: "+ strCISAddressLine3 &" , Actual: "& strIServeAddressLine3, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line3 value not matching with the expected value. Expected: "+ strCISAddressLine3 &" , Actual: "& strIServeAddressLine3, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine3(),strAddressLine3,"Address Line3")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strAddressLine4) Then
    	If strAddressLine4 = "RUNTIME" Then
			strCISAddressLine4=strRunTimeAddressLine4
    		strIServeAddressLine4=bcVerify_AccountAndAddress.lblAddressLine4.getroproperty("innertext")
    		If  Ucase(Trim(strCISAddressLine4)) = UCase(Trim(strIServeAddressLine4)) Then
    			LogMessage "RSLT", "Verification","For Address Line4 value successfully matched with the expected value. Expected: "+ strCISAddressLine4 &" , Actual: "& strIServeAddressLine4, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Address Line4 value not matching with the expected value. Expected: "+ strCISAddressLine4 &" , Actual: "& strIServeAddressLine4, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine4(),strAddressLine4,"Address Line4")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strPostalCode) Then
    	If strPostalCode = "RUNTIME" Then
			strCISPostalCode=strRunTimePostalCode
    		strIServePostalCode=bcVerify_AccountAndAddress.lblPostalCode.getroproperty("innertext")
    		If  Ucase(Trim(strCISPostalCode)) = UCase(Trim(strIServePostalCode)) Then
    			LogMessage "RSLT", "Verification","For Postal Code value successfully matched with the expected value. Expected: "+ strCISPostalCode &" , Actual: "& strIServePostalCode, True
				bverifyfields_AddressAcctLinkage = True
			else
				LogMessage "WARN", "Verification","For Postal Code value not matching with the expected value. Expected: "+ strCISPostalCode &" , Actual: "& strIServePostalCode, False
				bverifyfields_AddressAcctLinkage = False
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblPostalCode(),strPostalCode,"Postal Code")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strLastUpdatedDate) Then
    	If strLastUpdatedDate = "RUNTIME" Then
			strCISLastUpdatedDate=strRunTimeLastUpdatedInfo
    		strIServeLastUpdatedDate=bcVerify_AccountAndAddress.lblLastUpdatedDate.getroproperty("innertext")
    		strIServeLastUpdatedDate=FormatDateTime(strIServeLastUpdatedDate, 2)
    		If len(Month(CDate(strIServeLastUpdatedDate)))=1 Then
			
				strMonth="0"&Month(CDate(strIServeLastUpdatedDate))
			else
				strMonth=""&Month(CDate(strIServeLastUpdatedDate))
			End If
			strIServeLastUpdatedDate=Day(CDate(strIServeLastUpdatedDate))& "/"&strMonth& "/"&Year(CDate(strIServeLastUpdatedDate))&""
			
    		If Not matchStr(strCISLastUpdatedDate,strIServeLastUpdatedDate) Then    			
				bverifyfields_AddressAcctLinkage = False			
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblLastUpdatedDate(),strLastUpdatedDate,"Last Updated Date")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
    If Not IsNull(strLastUpdatedBy) Then
    	If strLastUpdatedBy = "RUNTIME" Then
			strCISLastUpdatedBy=strRunTimeLastUpdatedInfo
    		strIServeLastUpdatedBy=bcVerify_AccountAndAddress.lblLastUpdatedBy.getroproperty("innertext")    		  		
    		If Not matchStr(strCISLastUpdatedBy,strIServeLastUpdatedBy) Then    			
				bverifyfields_AddressAcctLinkage = False			
			End If
		Else
			If Not verifyInnerText(bcVerify_AccountAndAddress.lblLastUpdatedBy(),strLastUpdatedBy,"Last Updated By")Then
				bverifyfields_AddressAcctLinkage = False
			End If
		End If	
    End If
'    If Not IsNull(strSavingAccountNo) Then
'    	If strSavingAccountNo = "RUNTIME" Then
'			strCISLinkedSA=strRunTimeLinkedSA
'    		strIServeLinkedSA=bcVerify_AccountAndAddress.lblSavingAccountNo.getroproperty("innertext")
'    		If  Ucase(Trim(strCISLinkedSA)) = UCase(Trim(strIServeLinkedSA)) Then
'    			LogMessage "RSLT", "Verification","For Linked SA value successfully matched with the expected value. Expected: "+ strCISLinkedSA &" , Actual: "& strIServeLinkedSA, True
'				bverifyfields_AddressAcctLinkage = True
'			else
'				LogMessage "WARN", "Verification","For Linked SA value not matching with the expected value. Expected: "+ strCISLinkedSA &" , Actual: "& strIServeLinkedSA, False
'				bverifyfields_AddressAcctLinkage = False
'			End If
'		Else	
'			If Not verifyInnerText(bcVerify_AccountAndAddress.lblSavingAccountNo(),strSavingAccountNo,"Saving Account No")Then
'				bverifyfields_AddressAcctLinkage = False
'			End If
'		End If	
'    End If
    verifyfields_AddressAcctLinkage=bverifyfields_AddressAcctLinkage
End Function

'[Verify the row data for linked account table]
Public Function verifyrowdata_AddAccLinkage(arrRowDataList)
	bverifyrowdata_AddAccLinkage = true
	If IsNull(arrRowDataList) Then
		verifyrowdata_AddAccLinkage = verifyTableContentList(bcVerify_AccountAndAddress.tblLinkedAccountHeader,bcVerify_AccountAndAddress.tblLinkedAccountContent,arrRowDataList,"Linked Account",false,null,null,null)	
	End If
verifyrowdata_AddAccLinkage = bverifyrowdata_AddAccLinkage
End Function

'[Verify fields displayed in address and Account Linkage Enquiry Page for Cards]
Public Function verifyAddressAcctLinkage_CC(strName,strBlock,strLevelUnit,strAddressLine1,strAddressLine2,strAddressLine3,strAddressLine4,strPostalCode,strLastUpdatedDate)
bverifyfields_AddressAcctLinkage=true
  If Not IsNull(strName) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblName(),strName,"Name")Then
	bverifyfields_AddressAcctLinkage = False
	End If
  End If	
  If Not IsNull(strBlock) Then  
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblBlock(),strBlock,"Block")Then
	bverifyfields_AddressAcctLinkage = False
	End If
  End If	
  If Not IsNull(strLevelUnit) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblLevelUnit(),strLevelUnit,"Level")Then
	bverifyfields_AddressAcctLinkage = False
	End If
  End If	
  If Not IsNull(strAddressLine1) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine1(),strAddressLine1,"Address Line1")Then
	bverifyfields_AddressAcctLinkage = False
	End If
  End If
  If Not IsNull(strAddressLine2) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine2(),strAddressLine2,"Address Line2")Then
		bverifyfields_AddressAcctLinkage = False
	End If
  End If
  If Not IsNull(strAddressLine3) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine3(),strAddressLine3,"Address Line3")Then
		bverifyfields_AddressAcctLinkage = False
	End If
  End If
  If Not IsNull(strAddressLine4) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblAddressLine4(),strAddressLine4,"Address Line4")Then
			bverifyfields_AddressAcctLinkage = False
	End If
  End If
  If Not IsNull(strPostalCode) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblPostalCode(),strPostalCode,"Postal Code")Then
		bverifyfields_AddressAcctLinkage = False
	End If
  End If
  If Not IsNull(strLastUpdatedDate) Then
	If Not verifyInnerText(bcVerify_AccountAndAddress.lblLastUpdatedDate(),strLastUpdatedDate,"Last Updated Date")Then
		bverifyfields_AddressAcctLinkage = False
	End If
  End If
  verifyAddressAcctLinkage_CC=bverifyfields_AddressAcctLinkage
End Function
