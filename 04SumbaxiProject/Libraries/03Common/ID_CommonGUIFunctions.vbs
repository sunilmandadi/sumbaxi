'###############################################################################################
'# Name: VerifyTabExist()
'# Description: Function to verify the specific tab name from the list of dynamic tab displayed 
'# Author:
'# Date: 15-September-2017
'# Input Parameters: strTabName
'# 		strTabName = Name of the tab name to verify
'# Output Parameters: None
'# e.g: VerifyTabExist "FAILED SR"
'###############################################################################################
Public Function VerifyTabExist(strTabName)	
	Set oDescTab = Description.Create
	oDescTab("class").Value = "mat-tab-links"
	oDescTab("micClass").Value = "WebElement"
	Set objtablist = gObjIServePage.ChildObjects(oDescTab)
	
	Set oTabCaption = Description.Create
	oTabCaption("class").Value = "mat-tab-link tab.*"
	oTabCaption("micClass").value = "WebElement"
	Set oTabCaption = objtablist(0).ChildObjects(oTabCaption)
	iCount = oTabCaption.Count
	
	ReDim arrCols(iCount)
	For i = 0 To iCount-1
	  Dim strTabCaption
	  strTabCaption = oTabCaption(i).GetROProperty("innertext")
	  If Trim(UCase(strTabCaption)) = Trim(UCase(strTabName)) Then
		  VerifyTabExist = True
		  LogMessage "RSLT","Verification","Expected Tab : "&strTabName&" displayed successfully", True
	      Exit Function
      End If
	Next	
	
	LogMessage "WARN","Verification","Expected Tab : "&strTabName&" is not displayed", False
	
	Set oDescTab = Nothing
	Set objtablist = Nothing
	Set oTabCaption = Nothing
	VerifyTabExist = False
End Function

'################################################################################################
'# Name: VerifyListOfTabs()
'# Description: Function to verify list of tab names displayed 
'# Author:
'# Date: 18-September-2017
'# Input Parameters: lstTabnames
'#	 lstTabnames = List of tab names seperated with "|"
'# Output Parameters: None
'# e.g: VerifyListOfTabs "INTERACTION ACTIVITIES|SERVICE REQUEST|RECENT APPLICATIONS"
'################################################################################################
Public Function VerifyListOfTabs(lstTabnames)	
	VerifyListOfTabs = False
		
	Set oDescTab = Description.Create
	oDescTab("class").Value = "mat-tab-list"
	oDescTab("micClass").Value = "WebTabStrip"
	Set objtablist = gObjIServePage.ChildObjects(oDescTab)
	
	Set oTabCaption = Description.Create
	oTabCaption("class").Value = "mat-tab-label.*"
	oTabCaption("micClass").value = "WebElement"
	Set oTabCaption = objtablist(0).ChildObjects(oTabCaption)
	iCount = oTabCaption.Count
	
	'arrTabName = Split(lstTabnames,"|")	
	For Iterator = 0 To Ubound(lstTabnames)
		strTabName = lstTabnames(Iterator)	
		ReDim arrCols(iCount)
		For i = 1 To iCount-1
		  Dim strTabCaption
		  strTabCaption = oTabCaption(i).GetROProperty("innertext")
		  If Trim(UCase(strTabCaption)) = Trim(UCase(strTabName)) Then	
		  	 VerifyListOfTabs = True	
		  	 LogMessage "RSLT","Verification","Expected Tab : "&strTabName&" displayed successfully", True	
			 Exit For	
		  Else 
			 VerifyListOfTabs = False
	      End IF
	    Next	
		If VerifyListOfTabs <> True Then
		   LogMessage "RSLT","Verification","Expected Tab : "&strTabName&" is not displayed", False
		   Exit Function
		End If	
	 Next
	 
	Set oDescTab = Nothing
	Set objtablist = Nothing
	Set oTabCaption = Nothing
End Function
'################################################################################################
'# Name: ClickTab()
'# Description: Function to click on Tab available in Dashboard page 
'# Author:
'# Date: 12-December-2017
'# Input Parameters: Tabname
'# Output Parameters: None
'# e.g: ClickTab "RECENT APPLICATIONS"
'################################################################################################
Public Function ClickTab(strTabName)	
	ClickTab = True
		
	Set oDescTab = Description.Create
	oDescTab("class").Value = "mat-tab-list"
	oDescTab("micClass").Value = "WebTabStrip"
	Set objtablist = gObjIServePage.ChildObjects(oDescTab)
	
	Set oTabCaption = Description.Create
	oTabCaption("class").Value = "mat-tab-label.*"
	oTabCaption("micClass").value = "WebElement"
	Set oTabCaption = objtablist(0).ChildObjects(oTabCaption)
	iCount = oTabCaption.Count
	
		
		For i = 1 To iCount-1
		  Dim strTabCaption
		  strTabCaption = oTabCaption(i).GetROProperty("innertext")
		  If Trim(UCase(strTabCaption)) = Trim(UCase(strTabName)) Then
			 oTabCaption(i).Click
			 If Err.Number <> 0 Then
				ClickTab = False
				LogMessage "RSLT","Verification","Expected Tab : "&strTabName&" is not displayed", False
			 Else
				ClickTab = True
				LogMessage "RSLT","Verification","Expected Tab : "&strTabName&" is clicked successfully", True
			End If
			Exit For
	      End IF
	    Next	
			 
	 
	Set oDescTab = Nothing
	Set objtablist = Nothing
	Set oTabCaption = Nothing
End Function
'################################################################################################
'# Name: VerifyTabClose()
'# Description: Function to Close single tab by clicking cross icon
'# Authore
'# Date: 20-October-2017
'# Input Parameters: strTabName
'#	 strTabName = Tab name to be closed 
'# Output Parameters: None
'# e.g: VerifyTabClose("New IA")
'################################################################################################
Public Function TabClose(strTabName)	
	Set oDescTab = Description.Create
	oDescTab("class").Value = "mat-tab-links"
	oDescTab("micClass").Value = "WebElement"
	Set objtablist = gObjIServePage.ChildObjects(oDescTab)
	
	Set oTabCaption = Description.Create
	oTabCaption("class").Value = "mat-tab-link tab.*"
	oTabCaption("micClass").value = "WebElement"
	Set oTabCaption = objtablist(0).ChildObjects(oTabCaption)
	
	Set oCloseIcon = Description.Create
	oCloseIcon("class").Value = "removeIcon.*"
	oCloseIcon("micClass").value = "WebElement"
	Set oCloseIcon = objtablist(0).ChildObjects(oCloseIcon)
	iCount = oCloseIcon.Count
	
	For i = 0 To iCount-1
	  strTabCaption = oTabCaption(i+1).GetROProperty("innertext")
	  If Trim(UCase(strTabCaption)) = Trim(UCase(strTabName)) Then
		  oCloseIcon(i).Click	
		  LogMessage "RSLT","Verification","Tab : "&strTabName&" is closed successfully", true
		  TabClose = True
	      Exit Function
      End If
	Next	
	LogMessage "WARN","Verification","Tab : "&strTabName&" not found", false
	Set oDescTab = Nothing
	Set objtablist = Nothing
	Set oTabCaption = Nothing
	Set oCloseIcon = Nothing
	TabClose = False
End Function
'################################################################################################
'# Name: CloseAlltabs()
'# Description: Function to Close all tabs having crossmarks
'# Author:
'# Date: 20-October-2017
'# Input Parameters: None
'# Output Parameters: True/False
'################################################################################################
Public Function CloseAlltabs()	
	Set oDescTab = Description.Create
	oDescTab("class").Value = "mat-tab-links"
	oDescTab("micClass").Value = "WebElement"
	Set objtablist = gObjIServePage.ChildObjects(oDescTab)
	
	Set oCloseIcon = Description.Create
	oCloseIcon("class").Value = "removeIcon.*"
	oCloseIcon("micClass").value = "WebElement"
	Set oCloseIcon = objtablist(0).ChildObjects(oCloseIcon)
	iCount = oCloseIcon.Count
	
	For i = 0 To iCount-1
		oCloseIcon(i).Click	
		If Err.Number <> 0 Then
			CloseAlltabs = False
			LogMessage "RSLT","Verification","Failed to Click Close Icon",False
		End If
	Next	
	LogMessage "WARN","Verification","All tabs having close icon are closed successfully", True
	CloseAlltabs = True

	Set oDescTab = Nothing
	Set objtablist = Nothing
	Set oCloseIcon = Nothing
End Function
'################################################################################################
'# Name: verifyDateRange()
'# Description: Function to verify the date range between from and To Date displayed 
'# Author: 
'# Date: 18-September-2017
'# Input Parameters: ObjFromDate,ObjToDate,strDateRange,StrToDate
'# 		ObjFromDate = From Date object reference 
'# 		ObjToDate = To Date object reference 
'# 		strDateRange = No of days different between from and TO date (eg: 30 days, 31 days) 
'# 		StrToDate = Default To Date to be populated in the text box (eg : Today or Today-1)
'# Output Parameters: True or False
'# e.g: verifyDateRange(ObjFromDate,ObjToDate,30,TODAY)
'################################################################################################
Public Function VerifyDateRange(ObjFromDate,ObjToDate,strDateRange,StrToDate)
	bverifyDateRange = False
	strActFromDate = ObjFromDate.GetROProperty("value")
	strActToDate =  ObjToDate.GetROProperty("value")	
	
	If Trim(Ucase(StrToDate)) = "TODAY" Then
		strExpActToDate = NOW()	   	
		strmonth = Monthname(Month(strExpActToDate))
		strmonth = Mid(strmonth,1,3)
		strDay = Day(strExpActToDate)
		If len(strDay) = 1 Then
		   strDay = "0"&strDay
		End If
		strYear = Year(strExpActToDate)
		strExpActToDate = strDay &" " &strmonth &" " &strYear
	Else 
		strExpActToDate = StrToDate	
	End If
	
	If strExpActToDate = strActToDate Then
	     DaysRange = DateDiff("d", strActFromDate, strActToDate)
	     If Trim(DaysRange) = Trim(strDateRange) Then
			 LogMessage "RSLT","Verification","From and To Dates are displayed within "&strDateRange&" days range" ,True
			 bverifyDateRange = True 
		 Else 
		     bverifyDateRange = False 
	     End If	   		
	End If
	
	VerifyDateRange = bverifyDateRange
End Function
'################################################################################################
'# Name: VerifyDateSearchRecordsdisplayed()
'# Description: Function to verify if the table records are displayed based on dates searched
'# Author: 
'# Date: 19-September-2017
'# Input Parameters: tblheader,tblContent,strFromDate,strToDate,strColName
'# 		tblheader = Data Table header object
'# 		tblContent = Data Table row object
'# 		strFromDate = get the From date set in FROM Date field  
'# 		strToDate = get the TO date set in TO Date field 
'# 		strColName = Column name in table based which has data always
'# Output Parameters: True or False
'# e.g: VerifyDateSearchRecordsdisplayed(Objtblheader,ObjtblContent,strFromDate,strToDate,strColName)
'################################################################################################
Public Function VerifyDateSearchRecordsdisplayed(Objtblheader,ObjtblContent,strFromDate,strToDate,strColName)
   VerifyDateSearchRecordsdisplayed = False
   intRecordCount = getIDRecordsCountForColumn(Objtblheader,ObjtblContent,strColName)
   
   For i = 0 To intRecordCount - 1
		Set objAllRows = getIDAllRows(ObjtblContent)
		strCreatedOn = getIDCellTextFor(Objtblheader,objAllRows(i),i,strColName)					
		If Not IsNull(strToDate) Then
			ActualCreatedOn = FormatDateTime(strCreatedOn,vbShortDate)
			DaysRange = DateDiff("d", strFromDate, ActualCreatedOn)
			DaysRange1 = DateDiff("d", ActualCreatedOn, strToDate)
			If DaysRange < 0 Or DaysRange1 < 0  Then
			    LogMessage "RSLT","Verification","Records displayed in table Row "&i&" is not within the date range selected" ,False
			    VerifyDateSearchRecordsdisplayed = False
			    Exit Function 
			Else 
				LogMessage "WARN","Verification","Records displayed in table Row "&i&" is within the date range selected" ,True
				VerifyDateSearchRecordsdisplayed = True
			End If
		End If
   Next 
   
   Set objAllRows = Nothing
 End Function
'################################################################################################
'# Name: VerifySearchRecordsdisplayed()
'# Description: Function to verify if the table records are displayed based on dates searched
'# Author: 
'# Date: 19-September-2017
'# Input Parameters: tblheader,tblContent,strFromDate,strToDate,strColName
'# 		tblheader = Data Table header object
'# 		tblContent = Data Table row object
'# 	    strColName = Mandatory column name in table 
'#      StrExpValue = Expected value of the column name in table
'# Output Parameters: True or False
'# e.g: VerifySearchRecordsdisplayed(Objtblheader,ObjtblContent,strColName,StrExpValue)
'################################################################################################
Public Function VerifySearchRecordsdisplayed(Objtblheader,ObjtblContent,strColName,StrExpValue)
  VerifySearchRecordsdisplayed = False
  intRecordCount = getIDRecordsCountForColumn(Objtblheader,ObjtblContent,strColName)
  
  For i = 0 To intRecordCount - 1
	  Set objAllRows = getIDAllRows(ObjtblContent)
	  strActValue = getIDCellTextFor(Objtblheader,objAllRows(i),i,strColName)		 
	  If Instr(StrExpValue,strActValue) = 0 Then
	     LogMessage "RSLT","Verification","Records displayed in table is based on the search criteria",True
	     VerifySearchRecordsdisplayed = True 
	  Else 
	     LogMessage "RSLT","Verification","Records displayed in table is not based on the search criteria",False
	     VerifySearchRecordsdisplayed = False 
	  End If
  Next 
  
  Set objAllRows = Nothing
 End Function
'################################################################################################
'# Name: SelectCheckBoxAndVerify_ID()
'# Description: Function to Select checkbox and then check if the checked in ticked
'# Author: 
'# Date: 25-September-2017
'# Input Parameters: oCurrentObj,strObjectName
'# 		oCurrentObj = checkbox parent object
'# 	    strObjectName = Name of the object 
'# Output Parameters: True or False
'# e.g: SelectCheckBoxAndVerify_ID(oCurrentObj,strObjectName)
'################################################################################################
Public Function SelectCheckBoxAndVerify_ID(oCurrentObj,strObjectName)
	bselectCheckBoxFromOptions = False
	Set chkBox = Description.Create
	chkBox("class").value = "mat-checkbox-inner-container|datatable-checkbox"
	
	Set chkBoxObject = oCurrentObj.childObjects(chkBox)
	chkBoxObject(0).Click	
	
	If instr(1,oCurrentObj.getROProperty("class"),"mat-checkbox-checked") <> "0" Then
		LogMessage "RSLT","Verification",strObjectName & " checkbox : checked successfully", True 
		bselectCheckBoxFromOptions = True
	Else
		LogMessage "WARN","Verification",strObjectName & " checkbox : Failed to select the checkbox", False
	End If 
	
	SelectCheckBoxAndVerify_ID = bselectCheckBoxFromOptions
	
	Set oCurrentObj = Nothing
	Set chkBoxObject = Nothing
	Set chkBox = Nothing 
End Function
'################################################################################################
'# Name: VerifyObjectCheckedUnchecked()
'# Description: Function to verify an object is Checked/Unchecked and then write the result to log file.
'# Author: 
'# Date: 25-Sep-2017
'# Input Parameters: oCurrentObj,strCheckFlag,strObjectName
'# 		oCurrentObj = Object Checkbox
'# 		strCheckFlag = True/False
'# 		strObjectName = Name of the object to be verified (eg :OnceDone)
'# Output Parameters: None
'# E.g.: VerifyObjectCheckedUnchecked("WebElement","TRUE","Once Done")
'################################################################################################
Public Function VerifyObjectCheckedUnchecked(oCurrentObj,strCheckFlag,strObjectName)	
	bverifyCheck = True
	If oCurrentObj.Exist Then
		If instr(1,oCurrentObj.getROProperty("class"),"mat-checkbox-checked") <> "0" Then
			bverifyObjectChecked = True
		Else
			bverifyObjectChecked = False
		End If
		
		Select Case strCheckFlag
			Case "Checked"
				If bverifyObjectChecked Then
					LogMessage "RSLT","Verification",strObjectName & " is Checked as expected.",True
					bverifyCheck = True
				Else
					LogMessage "WARN","Verification",strObjectName & " is UnChecked. Expected Checked.",False
					bverifyCheck = False
				End If
			Case "Unchecked"
				If Not bverifyObjectChecked Then
					LogMessage "RSLT","Verification",strObjectName & " is not Checked as expected.",True
					bverifyCheck = True
				Else
					LogMessage "WARN","Verification",strObjectName & " is Checked. Expected Unchecked.",False
					bverifyCheck = False
				End If
		End Select
	Else
		LogMessage "WARN","Step Failure","Expected object " & strObjectName &" does not exist in current Webpage.",False
		bverifyCheck = False		
	End If
	
	Set oCurrentObj = Nothing
	VerifyObjectCheckedUnchecked = bverifyCheck
End Function
'################################################################################################
'# Name: SelectComboBoxItem()
'# Description: Function to Select a particular value from the combo box list 
'# Author: 
'# Date: 27-Sep-2017
'# Input Parameters: objComboBox,StrObjName,strItem
'# 		objComboBox = Object combo box 
'# 		StrObjName = Name of the object to be verified 
'# 		strItem = item to be selected from the combo box list 
'# Output Parameters: True/False
'# E.g.: SelectComboBoxItem(WebElement,"Open","Status")
'################################################################################################
Public Function SelectComboBoxItem(objComboBox,strItem,StrObjName)
	Setting.WebPackage("ReplayType") = 2
	objComboBox.Click		

	Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "mat-select-panel.*|mat-autocomplete-panel.*|mat-menu-panel.*"
	Set lstComboItems = gObjIServePage.ChildObjects(oDesc1)
	
	Set oDescCombo = Description.Create
	oDescCombo("micclass").Value = "WebElement"
	oDescCombo("class").Value = "mat-option|mat-option mat-selected|mat-menu-item"
	Set lstCombo = lstComboItems(0).ChildObjects(oDescCombo)
	
	intItems = lstCombo.Count
	Setting.WebPackage("ReplayType") = 1
	Dim strTemp1
	Dim strSelectedValue
	For iCount = 0 to intItems-1
		strTemp1 = lstCombo(iCount).GetRoProperty("text")
		If Ucase(Trim(strTemp1)) = Ucase(Trim(strItem))Then
			lstCombo(iCount).click
			WaitForIServeLoading
			'use GetRoProperty("value") for combobox object 
			strSelectedValue = objComboBox.GetRoProperty("value")
			If strSelectedValue <> "" Then
				If Trim(Ucase(strSelectedValue)) = Trim(Ucase(strItem)) Then
					LogMessage "RSLT","Verification","Item "&strItem&" selected from "&StrObjName&" listbox sucessfully.Item Index is "&intItemIndex,True
					SelectComboBoxItem = True 
					Exit For
				Else 
					SelectComboBoxItem = False		
				End IF
			Else 
				'use GetRoProperty("Outertext") for listbox object
				strSelectedValue = objComboBox.GetRoProperty("Outertext")
				If Trim(Ucase(strSelectedValue)) = Trim(Ucase(strItem)) Then
					LogMessage "RSLT","Verification","Item "&strItem&" selected from "&StrObjName&" listbox sucessfully.Item Index is "&intItemIndex,True
					SelectComboBoxItem = True 
					Exit For
				Else 
					SelectComboBoxItem = False		
				End IF				
			End If	
		End If
		intItemIndex = intItemIndex+1
	Next
	If SelectComboBoxItem <> True  Then
	   LogMessage "WARN","Verification","Item "&strItem&" Not found in "&StrObjName&" combobox",False
	End If				
	Set oDesc1 = Nothing 
	Set lstComboItems = Nothing 
	Set oDescCombo  = Nothing 
	Set lstCombo  = Nothing 	
End Function
'################################################################################################
'# Name: SelectItemfromList()
'# Description: Function to Select a single menu from ListMenu
'# Author: 
'# Date: 27-Sep-2017
'# Input Parameters: strItem,StrObjName 
'# 		strItem = item to be selected from the combo box list 
'# 		StrObjName = Name of the object to be verified 
'# Output Parameters: True/False
'# E.g.: SelectItemfromList("Cancel Card", "Overview Panel")
'################################################################################################
Public Function SelectItemfromList(strItem,StrObjName)
	bVerifylistItem = False 	

	Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "mat-select-panel.*|mat-autocomplete-panel.*|mat-menu-panel.*"
	Set lstComboItems = gObjIServePage.ChildObjects(oDesc1)
	
	Set oDescCombo = Description.Create
	oDescCombo("micclass").Value = "WebElement"
	oDescCombo("class").Value = "mat-option|mat-option mat-selected|mat-menu-item"
	Set lstCombo = lstComboItems(0).ChildObjects(oDescCombo)
	
	intItems = lstCombo.Count
	
	Dim strTemp1

	For iCount = 0 to intItems-1
		strTemp1 = lstCombo(iCount).GetRoProperty("text")
		
		If Trim(strTemp1) = Trim(strItem)Then
			lstCombo(iCount).click
			If Err.Number <> 0 Then
				bVerifylistItem = False
			Else
				bVerifylistItem = True
				LogMessage "RSLT","Verification","Item "&strItem&" selected from the "&StrObjName&" combobox", True
				Exit For 
			End If
			WaitForIServeLoading		
		End If
		
		intItemIndex = intItemIndex+1
	Next
	Setting.WebPackage("ReplayType") = 1
	If NOT bVerifylistItem Then
	   bVerifylistItem = False
	   LogMessage "WARN","Verification","Item "&strItem&" not selected from the "&StrObjName&" combobox", False
	End If
	
	SelectItemfromList = bVerifylistItem	
	
	Set oDesc1 = Nothing 
	Set lstComboItems = Nothing 
	Set oDescCombo  = Nothing 
	Set lstCombo  = Nothing 	
End Function
'################################################################################################
'# Name: verifyComboboxItems()
'# Description: Function to Verify the list of values displayed in combo box list is matched with expected 
'# Author: 
'# Date: 27-Sep-2017
'# Input Parameters: objComboBox,StrObjName,strItem
'# 		objComboBox = Object combo box 
'# 		lstItems = Closed|Failed|Open|Pending|Pending Submission
'# 		strComboboxName = label Name of the combo box to be verified 
'# Output Parameters: True/False
'# E.g.: verifyComboboxItems(Webelement,lstItems,"Status")
'################################################################################################
Public function verifyComboboxItems(objComboBox,lstItems,strComboboxName)
   Dim bVerifyDropDownListItems:bVerifyDropDownListItems = True
	   arrListItems = getItemsList_ComboBox(objComboBox)

   If not Ubound(arrListItems) = Ubound(lstItems) Then
		LogMessage "RSLT","Verification","Number of Items displayed in drop down list does not matched with expected",False
		verifyComboboxItems = False
		Exit function
   End If
   For iCount = 0 to UBound(lstItems)
		strItem = lstItems(iCount)
		If not ArrayFind(arrListItems,strItem) Then
			bVerifyDropDownListItems = False
			LogMessage "WARN","Verification","Item "&strItem&" does not displayed in Combobox "&strComboboxName,False
		Else
			LogMessage "RSLT","Verification","Item "&strItem&" displayed in Combobox "&strComboboxName,True
		End If
   Next
   verifyComboboxItems = bVerifyDropDownListItems
   Set objComboBox = Nothing
End Function
'#############################################################################################################################################################################
'# Name: getItemsList_ComboBox()
'# Description: Function to get the number of items from the combo box list
'# Author: 
'# Date: 27-Sep-2017
'# Input Parameters: objComboBox
'# 		objComboBox = Object combo box 
'# Output Parameters: Array 
'# E.g.: getItemsList_ComboBox(Webelement)
'#############################################################################################################################################################################
Public Function getItemsList_ComboBox(objComboBox)
	Setting.WebPackage("ReplayType") = 2
	objComboBox.Click
	Setting.WebPackage("ReplayType") = 1
    'get Combo box content list
	Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebList"
	'oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "mat-autocomplete.*|mat-select-content.*"
	'oDesc1("id").Value = "md-autocomplete.*"
	Set lstComboItems = gObjIServePage.ChildObjects(oDesc1)
	
	Set oDescCombo = Description.Create
	oDescCombo("micclass").Value = "WebElement"
	oDescCombo("class").Value= "mat-option|mat-option mat-active|mat-option mat-selected|mat-option-ripple.*"
	'oDescCombo("class").Value= "mat-option|mat-option mat-active|mat-option mat-selected"
	set lstCombo = lstComboItems(0).ChildObjects(oDescCombo)
	intItems = lstCombo.Count	
	
	'Get Count of Combo Items
	ReDim arrComboItems(Cint(strTotalItems)-1)
	Dim intItemIndex : intItemIndex = 0
	If strTotalItems = strCurrentItems Then
		ReDim arrComboItems(Cint(intItems)-1)
		For iCount= 0 to intItems-1
			strTemp = lstCombo(iCount).GetRoProperty("text")
			If strTemp <> "" Then
				arrComboItems(intItemIndex) = strTemp	
				intItemIndex = intItemIndex+1
			End If
		Next
		ReDim Preserve arrComboItems(Cint(intItemIndex)-1)
	End If
	Setting.WebPackage("ReplayType") = 2
	objComboBox.Click
	Setting.WebPackage("ReplayType") = 1
'	Dim mySendKey
'	set mySendKey = CreateObject("WScript.Shell")
'	mySendKey.SendKeys("{ENTER}")
'	mySendKey.SendKeys("{ENTER}")
'	Set mySendKey = nothing
	
	getItemsList_ComboBox = arrComboItems
	
	Set oDesc1 = Nothing 
	Set lstComboItems = Nothing 
	Set oDescCombo  = Nothing 
	Set lstCombo  = Nothing 	
End Function

'################################################################################################
'# Name: verifyComboboxItems1()
'# Description: Function to Verify the list of values displayed in combo box list is matched with expected 
'# Author: 
'# Date: 4 April 2018
'# Input Parameters: objComboBox,StrObjName,strItem
'# 		objComboBox = Object combo box
'#      objComboItems = Object container
'# 		lstItems = Closed|Failed|Open|Pending|Pending Submission
'# 		strComboboxName = label Name of the combo box to be verified 
'# Output Parameters: True/False
'# E.g.: verifyComboboxItems1(Webelement,Webelement,lstItems,"Status")
'################################################################################################
Public function verifyComboboxItems1(objComboBox,objComboItems,lstItems,strComboboxName)
   Dim bVerifyDropDownListItems:bVerifyDropDownListItems = True
	   arrListItems = getItemsList_ComboBox1(objComboBox,objComboItems)
   If not Ubound(arrListItems) = Ubound(lstItems) Then
		LogMessage "RSLT","Verification","Number of Items displayed in drop down list does not matched with expected"
		verifyDropDownListItems=false
		Exit function
   End If
   For iCount = 0 to UBound(lstItems)
		strItem = lstItems(iCount)
		If not ArrayFind(arrListItems,strItem) Then
			bVerifyDropDownListItems = False
			LogMessage "WARN","Verification","Item "&strItem&" does not displayed in Combobox "&strComboboxName,False
		Else
			LogMessage "RSLT","Verification","Item "&strItem&" displayed in Combobox "&strComboboxName,True
		End If
   Next
   verifyComboboxItems1 = bVerifyDropDownListItems
   Set objComboBox = Nothing
End Function
'###########################################################################################################################################
'# Name: getItemsList_ComboBox1()
'# Description: Function to get the number of items from the combo box list
'# Author: 
'# Date: 4 April 2018
'# Input Parameters: objComboBox
'# 		objComboBox = Object combo box 
'# Output Parameters: Array 
'# E.g.: getItemsList_ComboBox1(Webelement,Webelement)
'###########################################################################################################################################
Public Function getItemsList_ComboBox1(objComboBox,lstComboItems)
	objComboBox.Click
    'get Combo box content list
	Set oDescCombo = Description.Create
	oDescCombo("micclass").Value = "WebElement"
	oDescCombo("class").Value= "mat-option|mat-option mat-active|mat-option mat-selected|mat-option-ripple.*"
	set lstCombo = lstComboItems.ChildObjects(oDescCombo)
	intItems = lstCombo.Count	
	
	'Get Count of Combo Items
	ReDim arrComboItems(Cint(strTotalItems)-1)
	Dim intItemIndex : intItemIndex = 0
	If strTotalItems = strCurrentItems Then
		ReDim arrComboItems(Cint(intItems)-1)
		For iCount= 0 to intItems-1
			strTemp = lstCombo(iCount).GetRoProperty("text")
			If strTemp <> "" Then
				arrComboItems(intItemIndex) = strTemp	
				intItemIndex = intItemIndex+1
			End If
		Next
		ReDim Preserve arrComboItems(Cint(intItemIndex)-1)
	End If
	
	objComboBox.Click
	getItemsList_ComboBox1 = arrComboItems
	
	Set oDesc1 = Nothing 
	Set lstComboItems = Nothing 
	Set oDescCombo  = Nothing 
	Set lstCombo  = Nothing 	
End Function

'################################################################################################
'# Name: SetEditBoxInsideTable()
'# Description: Function to set text inside the table textbox and verify if the value set is displayed  
'# Author: 
'# Date: 04-Oct-2017
'# Input Parameters: Objtextbox,strSetText
'# 		Objtextbox = Object Textbox 
'# 		strSetText = The value to be set inside the text box for searching
'# Output Parameters: True/False
'# E.g.: SetEditBoxInsideTable(Objtextbox,"Customer01")
'################################################################################################
Public function SetEditBoxInsideTable(Objtextbox,strSetText)
	Dim StrActText
	Set oSearchText = Description.Create
	oSearchText("html tag").Value = "INPUT"
	oSearchText("class").value = ".*mat-input-element"
	
	Set oSearchEditBox = Objtextbox.ChildObjects(oSearchText)
	' Set value inside the text box 
	oSearchEditBox(0).set strSetText
	' Check if the value set is displayed in textbox appropriately
	StrActText = oSearchEditBox(0).GetROProperty("value")
		
	If Ucase(Trim(strSetText)) = Ucase(Trim(StrActText)) Then
	   LogMessage "RSLT","Verification","Textbox inside the table is set with value "&strSetText&" as Expected", True
	   SetEditBoxInsideTable = True
	Else
	   LogMessage "RSLT","Verification","Textbox inside the table is doesnt set with value "&strSetText&" as Expected", False
	   SetEditBoxInsideTable = False
	End If

	Set Objtextbox = Nothing 
	Set oSearchText = Nothing 
	Set oSearchEditBox = Nothing 
End Function
'################################################################################################
'# Name: VerifyInfowarntext()
'# Description: Function is to verify the text displayed inside the infowarn message text box 
'# Author: 
'# Date: 04-Oct-2017
'# Input Parameters: ObjInfowarn,strSetText
'# 		ObjInfowarn = Object Infowarn 
'# 		strSetText = string text message expected to be displayed for Inforwarn
'# Output Parameters: True/False
'# E.g.: VerifyInfowarntext(Objtextbox,"No Records found")
'################################################################################################
Public function VerifyInfowarntext(ObjInfowarn,strInfoMsgtext)
	Dim strActInfotext
	Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "iw-name-message"
	
	Set ObjInfowarn = ObjInfowarn.ChildObjects(oDesc1)
	strActInfotext = ObjInfowarn(0).GetRoProperty("Outertext")

	If Ucase(Trim(strInfoMsgtext)) = Ucase(Trim(strActInfotext)) Then
	   LogMessage "RSLT","Verification","Infowarn text message "&strInfoMsgtext&" displayed as expected.", True
	   VerifyInfowarntext = True
	Else
	   LogMessage "RSLT","Verification","Infowarn text message "&strInfoMsgtext&" not displayed as expected.", False
	   VerifyInfowarntext = False
	End If

	Set oDesc1 = Nothing 
	Set ObjInfowarn = Nothing  
End Function
'###############################################################################################
'# Name: VerifyIDLabelValuePairs()
'# Description: Function to verify label and value pairs
'# Author: 
'# Date: 26-Oct-2017
'# Input Parameters: objWebkitBox,arrLblValPairs, strProductType and strAccordionName
'# 		objWebkitBox = web object of the entier section , where to verify all label and values
'# 		arrLblValPairs = label and value pairs to validate
'#		strProductType =  Saving Accounts or Current Accounts or Credit Card etc... [Used for Reporting]
'# 		strAccordionName = Grey panel or Balance & Limits or Key Info etc .....[Used for Reporting]
'# Output Parameters: True/False
'# e.g: VerifyIDLabelValuePairs objWebkitBox,"Account Sign Type:Single|Account Number:1150080361|Account Status:Active|Account Opening Date:24 May 2012|Currency:IDR|Branch Code - Name:|Product Type:SA|Product Sub Type:DCLSB|Staff Payroll Account:NO", "Saving Account", "Grey Panel"
'###############################################################################################
Public Function VerifyIDLabelValuePairs(objWebkitBox,arrLblValPairs,strProductType,strAccordionName)
    Dim arrAppLblValPairs,oDiv,oDivChild,intChildCnt,i,strInnerText
    Dim strExpLbl,strExpVal,strExpLblValPair
    
    bIDLabelValuePairs1 = False
    bIDLabelValuePairs2 = True
    
    Set oDiv = Description.Create
    oDiv("class").Value = ".*isrv-view-holder"
    Set oDivChild = objWebkitBox.ChildObjects(oDiv)
    intChildCnt = oDivChild.Count-1
    
    ReDim Preserve arrLblValPairs(intChildCnt)
    
    For i = 0 To intChildCnt Step 1
        strInnerText = Trim(oDivChild(i).GetROProperty("innertext"))
        If (strInnerText <> "") And (arrLblValPairs(i) <> "" )Then
            strExpLbl = Split(arrLblValPairs(i), ":")(0)
            strExpVal = Split(arrLblValPairs(i), ":")(1)
            strExpLblValPair = Replace(arrLblValPairs(i),":"," ")
            If (Instr(strInnerText,strExpLbl) > 0) And (Instr(strInnerText,strExpVal) > 0) Then
                LogMessage "RSLT","Verification","Label and Value Verification for - Product Type : |" & strProductType &_
                "| - Section Name: |" & strAccordionName & "| - Label Name and Value : |" & strExpLblValPair &_
                "| is displayed as expected",True
                bIDLabelValuePairs1 = True
            Else
                LogMessage "WARN","Verification","Label and Value Verification for - Product Type : |" & strProductType &_
                "| - Section Name: |" & strAccordionName & "| - Label Name and Value : |" & strExpLblValPair &_
                "| is not displayed as expected, Actual: |" & strInnerText  & "|",False
                bIDLabelValuePairs2 = False
            End If    
        End If
    Next
    
    If bIDLabelValuePairs1 And bIDLabelValuePairs2 Then
        VerifyIDLabelValuePairs = True
    Else
        VerifyIDLabelValuePairs = False
    End If
    
    Set oDivChild = Nothing
    Set oDiv = Nothing
End Function
'###############################################################################################
'# Name: VerifyIDLabelValuePairsRandom()
'# Description: Function to verify label and value pairs
'# Author: 
'# Date: 01-Dec-2017
'# Input Parameters: objWebkitBox,arrLblValPairs, strProductType and strAccordionName
'# 		objWebkitBox = web object of the entier section , where to verify all label and values
'# 		arrLblValPairs = label and value pairs to validate
'#		strProductType =  Saving Accounts or Current Accounts or Credit Card etc... [Used for Reporting]
'# 		strAccordionName = Grey panel or Balance & Limits or Key Info etc .....[Used for Reporting]
'# Output Parameters: True/False
'# e.g: VerifyIDLabelValuePairs objWebkitBox,"Account Sign Type:Single|Account Number:1150080361|Account Status:Active|Account Opening Date:24 May 2012|Currency:IDR|Branch Code - Name:|Product Type:SA|Product Sub Type:DCLSB|Staff Payroll Account:NO", "Saving Account", "Grey Panel"
'###############################################################################################
Public Function VerifyIDLabelValuePairsRandom(objWebkitBox,arrLblValPairs,strProductType,strAccordionName)
	Dim arrAppLblValPairs,oDiv,oDivChild,intChildCnt,i,strInnerText
	Dim strExpLbl,strExpVal,strExpLblValPair
	
	bIDLabelValueAllPairs=True
	cntUnMatch=0	
	Set oDiv = Description.Create
	oDiv("class").Value = ".*isrv-view-holder.*|isrv-input-holder"
	Set oDivChild = objWebkitBox.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count-1
	
	If (Not IsNull(arrLblValPairs)) And IsArray(arrLblValPairs) Then
		size = Ubound(arrLblValPairs)
	Else
		LogMessage "WARN","Verification","There are no Lable Values are passed or not passed them as array",False
		VerifyIDLabelValuePairsRandom = False
		Exit Function
	End If
	
	For intSize = 0 To size Step 1
		bIDLabelValuePair = False
		strExpLblValPair = Trim(Replace(arrLblValPairs(intSize),":"," ",1,1))
		
		For i = 0 To intChildCnt Step 1
			
			strInnerText = Trim(oDivChild(i).GetROProperty("innertext"))
			
			If (strInnerText <> "") And (arrLblValPairs(intSize) <> "" )Then
				If Instr(1,strInnerText,strExpLblValPair) > 0 Then
					bIDLabelValuePair = True
					Exit For
				End If	
			End If
		Next
		If bIDLabelValuePair Then
			LogMessage "RSLT","Verification","Label and Value Verification for - Product Type : |" & strProductType &_
			"| - Section Name: |" & strAccordionName & "| - Label Name and Value : |" & strExpLblValPair &_
			"| is displayed as expected",True
		Else
			LogMessage "WARN","Verification","Label and Value Verification for - Product Type : |" & strProductType &_
			"| - Section Name: |" & strAccordionName & "| - Label Name and Value : |" & strExpLblValPair &_
			"| is not displayed as expected, Actual: |" & oDivChild(intSize).GetROProperty("innertext")  & "|",False
			cntUnMatch=cntUnMatch+1
		End If
	Next
		
	If Cint(cntUnMatch)>0 Then
		bIDLabelValueAllPairs = False
	End If
	VerifyIDLabelValuePairsRandom=bIDLabelValueAllPairs
	Set oDivChild = Nothing
	Set oDiv = Nothing
End Function

'#######################################################################################################################################################
'# Name: VerifytablePagination()
'# Description: Function is to verify the display of Next & Previous button on page level in table 
'# Author: 
'# Date: 12-Oct-2017
'# Input Parameters: Objtblheader,ObjtblContent,objFristPage,objPreviousPage,objNextPage,objLastPage,NoOfRows
'# 		Objtblheader = table header object reference
'# 		ObjtblContent = table Contect object reference 
'# 		objFristPage = object reference for First Arrow in Pagination of table footer 
'# 		objPreviousPage = object reference for Previous Arrow in Pagination of table footer 
'# 		objNextPage = object reference for Next Arrow in Pagination of table footer 
'# 		objLastPage = object reference for Last Arrow in Pagination of table footer 
'# 		NoOfRows = Maximum No of rows expected to display per page
'# Output Parameters: True/False
'#########################################################################################################################################################
Public Function VerifytablePagination(Objtblheader,ObjtblContent,objFristPage,objPreviousPage,objNextPage,objLastPage,strcolumnname,NoOfRows)
	Dim iCheck
	bverifypagination = False
	intRecordCount = getIDRecordsCountForColumn(Objtblheader,ObjtblContent,strcolumnname)	
	intRecordCount = intRecordCount-1' Adding one becoz the table has the scroll will shown +1 records always
	iCheck  = Cint(NoOfRows)

	bFirstPageDiasbled = matchStr(objFristPage(0).GetROProperty("class"),"disabled")
	bPreviousPageDisabled = matchStr(objPreviousPage(0).GetROProperty("class"),"disabled")
	bNextPageDisabled = matchStr(objNextPage(0).GetROProperty("class"),"disabled")
	bLastPageDiasbled = matchStr(objLastPage(0).GetROProperty("class"),"disabled")
	
	If intRecordCount = iCheck Then			
		  If (bFirstPageDiasbled AND bPreviousPageDisabled) = True Then
		  	LogMessage "RSLT","Verification","Previous Arrows are disabled for first page as expected", True 
		  	bverifypagination = True
			Else 
			LogMessage "WARN","Verification","Previous Arrows are not disabled for first page as expected", False
			bverifypagination = False
		  End If	
		  If bLastPageDiasbled = False Then 
		 	 Setting.WebPackage("ReplayType") = 2
		  	 objLastPage(0).Click
		  	 wait(1)
		  	 bLastPageDiasbled = matchStr(objLastPage(0).GetROProperty("class"),"disabled")
		  	 bNextPageDisabled = matchStr(objNextPage(0).GetROProperty("class"),"disabled")
			 If bLastPageDiasbled AND bNextPageDisabled Then
				LogMessage "WARN","Verification","Next and Last Arrows are disabled as expected", True
				bverifypagination = True
			 Else 
			    LogMessage "WARN","Verification","Next and Last Arrows are not disabled as expected", False
				bverifypagination = False
			 End If
		  End If
	ElseIf intRecordCount < iCheck  Then	
		If bNextPageDisabled AND bLastPageDiasbled AND bFirstPageDiasbled AND bPreviousPageDisabled  Then
			LogMessage "RSLT","Verification","Next and Previous Arrows are disabled as expected when number of records is less than "&iCheck, True 
			bverifypagination = True
		Else 
			LogMessage "WARN","Verification","Next and Previous Arrows are not disabled as expected when number of records is less than "&iCheck, False
			bverifypagination = False
		End If		
	End IF
	Setting.WebPackage("ReplayType") = 1
	 
	VerifytablePagination = bverifypagination
	 
	Set objFristPage = Nothing
	Set objNextPage = Nothing
	Set objLastPage = Nothing
	Set objPreviousPage = Nothing
End Function
'###############################################################################################
'# Name: VerifyActionLinks()
'# Description: Function to verify Action links displayed 
'# Author: 
'# Date: 25-Oct-2017
'# Input Parameters: arrLblValPairs and strProductType
'# 		arrLblValPairs = label and value pairs to validate in grey panel
'#		strProductType =  SA or CA or Credit Card etc... [Used for Reporting]
'# Output Parameters: None
'# e.g: VerifyActionLinks "Block Card|Cancel Card|Card Activation|Card Replacement|Temp Limit Increase|Unblock Card(CSO)|Unblock Card(Retention)", "SA", "Object"
'###############################################################################################
Public Function VerifyActionLinks(lstActions,strProductType,oDesc)
	bVerifyActionLinks1 = False
	bVerifyActionLinks2 = True
	
	Set oDiv = Description.Create
	oDiv("class").Value = "mat-primary mat-button|mat-warn mat-button.*"

	Set oDivChild = oDesc.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count-1
	
	For i = 0 To Ubound(lstActions) Step 1
		strLink = Trim(lstActions(i))

		For j = 0 To oDivChild.Count-1 Step 1
		
			strInnerText = Trim(oDivChild(j).GetROProperty("innertext"))

			If strInnerText = strLink Then
				LogMessage "RSLT","Verification","Action Link: "&strLink&" displayed as expected in Actions section",True
				bVerifyActionLinks1 = True
				Exit For
			Else 
				bVerifyActionLinks1 = False
			End IF
		Next
					
		If Not bVerifyActionLinks1 Then 
	   	   bVerifyActionLinks2 = False
	   	   LogMessage "WARN","Verification","Action Link: "&strLink&" not displayed in Actions section",False
		End IF 
	Next
	
	If bVerifyActionLinks1 AND bVerifyActionLinks2 Then
	   VerifyActionLinks = True
	Else 
	   VerifyActionLinks = False
	End If
	
	Set oDivChild = Nothing
	Set oDiv = Nothing
	Set oDesc = Nothing
End Function
'###############################################################################################
'# Name: ClickActionLinks()
'# Description: Function to Click Action links displayed 
'# Author: 
'# Date: 20-Dec-17
'# Input Parameters: arrLblValPairs and strProductType
'# 		strLinkName = Action link name to be clicked
'#		strProductType =  SA or CA or Credit Card etc... [Used for Reporting]
'# Output Parameters: None
'# e.g: ClickActionLinks "Block Card", "SA", "Object"
'###############################################################################################
Public Function ClickActionLinks(strLinkName,strProductType,oDesc)
	Set oDiv = Description.Create
	oDiv("class").Value = "mat-primary mat-button|mat-warn mat-button.*"

	Set oDivChild = oDesc.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count-1
	
	For j = 0 To oDivChild.Count-1 Step 1
		If Ucase(Trim(oDivChild(j).GetROProperty("innertext"))) = Ucase(Trim(strLinkName)) Then
			oDivChild(j).Click
			LogMessage "RSLT","Verification","Clicked Action Link: "&strLinkName&" displayed below Actions section",True
			ClickActionLinks = True
			Exit Function
		Else 
			ClickActionLinks = False
		End IF
	Next
	If NOT ClickActionLinks Then
	   LogMessage "WARN","Verification","Unable to click Action Link: "&strLinkName&" displayed below Actions section",False		
	End If
	
	Set oDivChild = Nothing
	Set oDiv = Nothing
	Set oDesc = Nothing
End Function
'###############################################################################################
'# Name: VerifyAccordionHeader()
'# Description: Function to verify list of Accordion titles displayed
'# Author: 
'# Date: 01-Nov-2017
'# Input Parameters: ObjAccordionGroup and lstAccordion
'# 		ObjAccordionGroup = parent object id of the Accordion list displayed 
'#		lstAccordion =  name of Accordion title {order should be same as displayed in page} 
'# Output Parameters: None
'# e.g: VerifyAccordionHeader "object", "Additional Card Info|Key Info|Balances & Limits|Transaction Details|Statement|Memo|Rewards"
'###############################################################################################
Public Function VerifyAccordionHeader(ObjAccordionGroup,lstAccordion)
	bverifyAccordion1 = False
	bverifyAccordion2 = True
	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-heading.*|panel panel-default.*"

	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	
	For j = 0 To Ubound(lstAccordion)
		strExpAccordionHeader = Split(lstAccordion(j),":")(0)
		strExpAccordionState  = Split(lstAccordion(j),":")(1)
		
		For k = 0 To oDivChild.Count-1	
			strActAccordionHeader = oDivChild(k).GetRoProperty("innertext")
			strActAccordionHeader = Trim(Replace(strActAccordionHeader,"Open | Closed",""))	
			
			If Instr(strExpAccordionHeader,strActAccordionHeader) > 0 Then	
				If Trim(strExpAccordionState) = "Enable"  AND Instr(1,oDivChild(k).GetRoProperty("innerhtml"),"disable-accordion") = 0 Then 
					LogMessage "RSLT","Verification","Accordion Header: "&strExpAccordionHeader&" is displayed in expected state: "&strExpAccordionState, True
					bverifyAccordion1 = True
					Exit For
				ElseIf Trim(strExpAccordionState) = "Disable" AND Instr(1,oDivChild(k).GetRoProperty("innerhtml"),"disable-accordion") > 0 Then
					LogMessage "RSLT","Verification","Accordion Header: "&strExpAccordionHeader&" is displayed in expected state: "&strExpAccordionState, True
					bverifyAccordion1 = True
					Exit For
				Else 
					LogMessage "WARN","Verification","Accordion Header: "&strExpAccordionHeader&" is not displayed in expected state: "&strExpAccordionState, False
		 			bverifyAccordion2 = False
					Exit For		 			
				End IF
			End IF	
		 Next 
			If Not bverifyAccordion1 Then
			LogMessage "WARN","Verification","Accordion Header: "&strExpAccordionHeader&" not found", False
			End If		
	 Next
	
	If bverifyAccordion1 And bverifyAccordion2 Then
	   VerifyAccordionHeader = True
	Else
	   VerifyAccordionHeader = False
	End If
	
	Set oDiv = Nothing 
	Set oDivChild = Nothing
End Function
'###############################################################################################
'# Name: VerifyAccordionRefresh()
'# Description: Function is to click on Refresh icon for the selected Accordion title
'# Author: 
'# Date: 02-Nov-2017
'# Input Parameters: ObjAccordionGroup and strAccordion
'# 		ObjAccordionGroup = parent object id of the Accordion list displayed 
'#		strAccordion = Accordion name for which Refresh icon to be clicked
'# Output Parameters: None
'# e.g: VerifyAccordionRefresh "object", "Key Info"
'###############################################################################################
Public Function VerifyAccordionRefresh(ObjAccordionGroup,strAccordion)
	bVerifyRefresh = False
	
	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-title"
	
	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	
	Set oRefresh = Description.Create
	oRefresh("Class").Value = ".*mat-button"
	
	Set oChildRefresh = ObjAccordionGroup.ChildObjects(oRefresh)

	For j = 0 To oDivChild.Count-1 Step 1
		strActAccordionHeader = oDivChild(j).GetRoProperty("innertext")
		If Trim(strAccordion) = Trim(strActAccordionHeader) Then
		   oChildRefresh(j).Click
		   LogMessage "RSLT","Verification","Refresh Icon clicked for the Accordion titled:"&strActAccordionHeader, True
		   bVerifyRefresh = True
		   Exit For
		End If
	Next
	
	If Not bVerifyRefresh Then  
		LogMessage "WARN","Verification","Refresh Icon not found for the Accordion titled:"&strActAccordionHeader, False
	End If
	
	VerifyAccordionRefresh = bVerifyRefresh
	
	Set oDiv = Nothing 
	Set oDivChild = Nothing
	Set oRefresh = Nothing 
	Set oChildRefresh = Nothing
End Function
'###############################################################################################
'# Name: ClickExpandIcon()
'# Description: Function is to click on Expand Icon displayed 
'# Author: 
'# Date: 02-Nov-2017
'# Input Parameters: ObjAccordionGroup and strAccordion
'# 		ObjAccordionGroup = parent object id of the Accordion list displayed 
'#		strAccordion = Accordion name for which Expand icon to be clicked
'# Output Parameters: None
'# e.g: ClickExpandIcon "object", "Key Info"
'###############################################################################################
Public Function ClickExpandIcon(ObjAccordionGroup,strAccordion)
	bClickExpandIcon = False
	
	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-title"

	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count
	
	Set oExpandIcon = Description.Create
	oExpandIcon("Class").Value = "expandoIcon"
	
	Set oChildExpandIcon = ObjAccordionGroup.ChildObjects(oExpandIcon)
	
	For j = 0 To intChildCnt-1 Step 1
		strActAccordionHeader = Trim(Replace(oDivChild(j).GetRoProperty("innertext"),"Open | Closed",""))
		If Instr(Trim(strAccordion),strActAccordionHeader) > 0 Then
		   oChildExpandIcon(j).Click
		   If Err.Number <> 0 Then
			  LogMessage "WARN","Verification","Failed to Click Accordion titled:"&strActAccordionHeader, False
			  Exit For
		   Else 
			  LogMessage "RSLT","Verification","Expand icon clicked for the Accordion titled:"&strActAccordionHeader, True
			  bClickExpandIcon = True
			  Exit For
		   End If
		End If
	Next
	
	ClickExpandIcon = bClickExpandIcon
	
	Set oDiv = Nothing
	Set oDivChild = Nothing
	Set oExpandIcon = Nothing
	Set oChildExpandIcon = Nothing
End Function
'###############################################################################################
'# Name: ClickMultipleAccordions()
'# Description: Function is to expand or collapse multiple Accordions displayed in page 
'# Author: 
'# Date: 21-Nov-2017
'# Input Parameters: ObjAccordionGroup and lstAccordion
'# 		ObjAccordionGroup = parent object id of the Accordion list displayed 
'#		lstAccordion = list of Accordion Names splitted with "|" 
'# Output Parameters: None
'# e.g: ClickMultipleAccordions "object", "Additional Card Info|Key Info"
'###############################################################################################
Public Function ClickMultipleAccordions(ObjAccordionGroup,lstAccordion)
	bClickExpandIcon = False

	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-title"

	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count

	Set oExpandIcon = Description.Create
	oExpandIcon("Class").Value = "expandoIcon"

	Set oChildExpandIcon = ObjAccordionGroup.ChildObjects(oExpandIcon)

	For i = 0 To Ubound(lstAccordion)
		strAccordion = Split(lstAccordion(i),"|")(0)
	
		For j = 0 To intChildCnt-1 Step 1
			strActAccordionHeader = Trim(Replace(oDivChild(j).GetRoProperty("innertext"),"Open | Closed",""))
			If Instr(Trim(strAccordion),strActAccordionHeader) > 0 Then
			   oChildExpandIcon(j).Click
			   If Err.Number <> 0 Then
				  LogMessage "WARN","Verification","Failed to Click Accordion titled:"&strActAccordionHeader, False
				  Exit For
			   Else 
				  LogMessage "RSLT","Verification","Expand icon clicked for the Accordion titled:"&strActAccordionHeader, True
				  bClickExpandIcon = True
				  Exit For
			   End If
			End IF
		Next	
	Next

	ClickMultipleAccordions = bClickExpandIcon

	Set oDiv = Nothing
	Set oDivChild = Nothing
	Set oExpandIcon = Nothing
	Set oChildExpandIcon = Nothing
End Function
'###############################################################################################
'# Function Name: VerifyTableSingleRowData
'# Description: Function to verify Header and single row value from data table which having no pagination 
'# Author: 
'# Date: 11-Sept-2017
'# Input Parameter(s): tblHeader,tblRow,lstlstAccountData,strTableName
'#		tblHeader = Data Table header object
'#		tblContent = Data Table row object
'# 		lstlstAccountData = Header and Row data as array list , Eg:(Name in English:RIJU|PRIMARY ID TYPE:ID Number – KTP|PRIMARY ID NUMBER:50209799999999999999)|
'#		strTableName = Name of the Data Table for Reporting, Eg: Customer Search
'# Output Parameter(s): True or False
'###############################################################################################
Public Function VerifyTableSingleRowData(tblHeader,tblContent,lstlstAccountData,strTableName)
	Dim bVerifyData,arrColumns,arrValues,intSize
	
	VerifyTableSingleRowData = True
	
	For iRowCount = 0 To Ubound(lstlstAccountData,1)
	
		intSize = Ubound(lstlstAccountData,2)
	
		ReDim arrColumns(intSize)
		ReDim arrValues(intSize)			
	
		For iCount = 0 To intSize
	
			If 	Instr(lstlstAccountData(iRowCount,iCount),"Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Created On")<>0 or _ 
				Instr(lstlstAccountData(iRowCount,iCount),"Due Date")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Created Date")<>0 or _ 
				Instr(lstlstAccountData(iRowCount,iCount),"Survey Response Date")<>0  or _
				Instr(lstlstAccountData(iRowCount,iCount),"Survey Listing")<>0 or _ 
				Instr(lstlstAccountData(iRowCount,iCount),"Date & Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Date and Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Sent Date")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Template Id")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Memo Details")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Created Date & Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Entered Date & Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Exit Date & Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Survey Listing")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Created Date/Time")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Event Date")<>0 or _
				Instr(lstlstAccountData(iRowCount,iCount),"Event Description")<>0 Then
				
				strTemp = ""
				
				If lstlstAccountData(iRowCount,iCount) = "" Then
					Exit For
				End If
			
				arrTemp = Split(lstlstAccountData(iRowCount,iCount),":")
				arrColumns(iCount) = arrTemp(0)
				
				For iTemp = 0 to Ubound(arrTemp)
					If 	arrTemp(iTemp) = "Completed Date/Time" or _
						arrTemp(iTemp)="Created Date/Time" or _
						arrTemp(iTemp)="Created On" or _
						arrTemp(iTemp)="Due Date" or _
						arrTemp(iTemp)="Created Date" or _
						arrTemp(iTemp)="Survey Response Date" or _
						arrTemp(iTemp)="Survey Listing" or _
						arrTemp(iTemp)="Date and Time" or _
						arrTemp(iTemp)= "Date & Time" or _
						arrTemp(iTemp)= "Payment Due Date" or _
						arrTemp(iTemp)= "Sent Date" or _
						arrTemp(iTemp)= "Template Id" or _
						arrTemp(iTemp)= "Memo Details" or _
						arrTemp(iTemp)= "Created Date & Time" or _
						arrTemp(iTemp)= "Entered Date & Time" or _
						arrTemp(iTemp)= "Exit Date & Time" or _
						arrTemp(iTemp)= "Survey Listing" or _
						arrTemp(iTemp)= "Event Date" or _
						arrTemp(iTemp)= "Event Description" Then
						
						arrTemp(iTemp) = checkNull(Replace(arrTemp(1),"-",":"))
						If strTemp = "" Then
							strTemp = checkNull(arrTemp(iTemp))
						else
							strTemp = strTemp &":"& checkNull(arrTemp(iTemp))
						End If
					End If
				Next
				arrValues(iCount) = strTemp
			Else
				arrTemp = Split(lstlstAccountData(iRowCount,iCount),":")
				arrColumns(iCount) = arrTemp(0)
				arrValues(iCount) = checkNull(arrTemp(1))
			End If
	
		Next
	
		intRow = getIDRowForColumns(tblHeader,tblContent,arrColumns, arrValues)
		
		If intRow = -1  Then
			LogMessage "WARN","Verification","Failed : Expected "& strTableName &" Data ["&ArrayToString(arrValues,",")&"] for respective column Names ["&ArrayToString(arrColumns,",")&"] are not found in  "& strTableName &" table",False
			VerifyTableSingleRowData = False
		Else
			LogMessage "RSLT","Verification","Expected  "& strTableName &" Data ["&ArrayToString(arrValues,",")&"] for respective column Names ["&ArrayToString(arrColumns,",")&"] found in "& strTableName &" table",True
		End If
	Next
End Function


Public Function getIDRowForColumns(objTableHeader,objContentTable,arrColumnName, arrValue)
	Set objAllRows = getIDAllRows(objContentTable)	
	intRow = objAllRows.Count   
	intColCount = UBound(arrColumnName)
	
	Dim arrColIndex, arrCellVal
	ReDim arrCellVal (intColCount)
	ReDim arrColIndex(intColCount)
	
	For i = 0 To intRow-1
		For j = 0 To intColCount
			Dim strColName, strCellVal			
			strColName = arrColumnName(j)			
			If not isEmpty(strColName) Then			
				strCellVal = Trim(getIDCellTextFor(objTableHeader,objAllRows(i),i,strColName))
				If isNull(arrValue(j)) Then
					strCellVal = Null
				End If		
				strCellVal = Replace(strCellVal,"(","")
				strCellVal = Replace(strCellVal,")","")
				arrCellVal(j) = strCellVal
				If (strCellVal="undefined") OR isnull(strCellVal) Then
					getIDRowForColumns = -1
					Exit Function
				End If
			End If
		Next
		
		Print(Vbtab & vbtab & "*******")
		Print (Vbtab & vbtab & arrCellVal(0))
		Print (Vbtab & vbtab & arrValue(0))
		
		If compareArray(arrValue,arrCellVal) Then
			getIDRowForColumns = i
			Exit Function
		End If
	Next
	getIDRowForColumns = -1
	Set objAllRows = Nothing
End Function

Public Function getIDAllRows(objContentTable)
	Set odesc_AllRows = Description.Create
	odesc_AllRows("class").value = "datatable-row-wrapper"
	Set getIDAllRows = objContentTable.ChildObjects(odesc_AllRows)
	Set odesc_AllRows = Nothing
End Function

Public Function getIDCellTextFor(objTableHeader,objRow,intRow,strColName)
	Dim intCol	
	intCol = getIDColIndex(objTableHeader,strColName)	
	Set odesc_Cell = Description.Create
	odesc_Cell("xpath").value= ".//div[contains(@class,'datatable-body-cell')]"
	Set allCellInRow = objRow.childObjects(odesc_Cell) 'Get all cell object for row
	getIDCellTextFor = allCellInRow(intCol).getRoProperty("innerText") 'allCellInrow get text for indexarray	
	Set allCellInRow = Nothing
	Set odesc_Cell = Nothing
End Function

Public Function getIDColIndex(objHeaderTable,strColName)
	Dim intCol
	Dim arrCols
	Set odesc_colHeaderCell = Description.Create
	odesc_colHeaderCell("xpath").value = ".//datatable-header-cell[contains(@class,'datatable-header-cell')]"
	odesc_colHeaderCell("visible").value = True
	Set tableColumnsObj = objHeaderTable.ChildObjects(odesc_colHeaderCell)	
	intCol = tableColumnsObj.Count
	ReDim arrCols(intCol)
	For i = 0 To intCol-1
		Dim strColHeader
		strColHeader = tableColumnsObj(i).GetROProperty("innertext")
		If Trim(UCase(strColHeader)) = Trim(UCase(strColName)) Then
			getIDColIndex = i
			Exit Function
		End If
	Next
	Set tableColumnsObj = Nothing
	Set odesc_colHeaderCell = Nothing
End Function

Public function getIDRecordsCountForColumn(objTableHeader,objContentTbl,strColumnName)
	Set objAllRows = getIDAllRows(objContentTbl)
	intRow = objAllRows.Count   
	intCount = 0
	iBlankRows = 0
	For i = 0 to intRow-1
		Dim strCellVal
		strCellVal = getIDCellTextFor(objTableHeader,objAllRows(i),i,strColumnName)
		Print "strCellVal: "&strCellVal
		If isNull(strCellVal) OR strCellVal="" Then
			iBlankRows = iBlankRows+1
		else
			intCount = intCount+1
			print (strCellVal)	
		End If
		Print("*******")
	Next
	getIDRecordsCountForColumn = intCount
	Set objAllRows = Nothing
End Function

Public Function SelectDateFromIDCalendar(objCalendar,strDate)
    
    SelectDateFromIDCalendar = False
    
    Set objBrowser = gObjIServePage
    objBrowser.RefreshObject
    
'    If objBrowser.WebEdit("xpath:=(//input[contains(@class,'datepicker-')])[1]").GetROProperty("disabled") = 1 Then
'    	Msgbox "Date Picker is in disabled state.Unable to select date"
'    	Exit Function
'    End If
    '[ Added By Raghu - 03/01/2018 - to verify the Calender text box is in enabled mode or not]
    If objCalendar.GetROProperty("disabled") = 1 Then
    	Msgbox "Date Picker is in disabled state.Unable to select date"
    	Exit Function
    End If
    
    arrDate = Split(strDate," ")
    strDay = arrDate(0)
    strMonth = arrDate(1)
    strYear = arrDate(2)
    
    Setting.WebPackage("ReplayType") = 2
    objCalendar.Click
    
    'To get calendar on page
    Set oDesc1 = Description.Create
    oDesc1("class").Value = "datepicker-panel"
    Set objCalendar = objBrowser.ChildObjects(oDesc1)
	
    'To Change Year
    Set oDesc2 = Description.Create
    oDesc2("class").Value="period-switch__period.*"
    Set objYear = objCalendar(0).ChildObjects(oDesc2)
    
    objYear(0).Click
    
    'To Change Month
    strCurrCalYear = Trim(objYear(0).GetRoProperty("innertext"))
    strCurrCalYear = Split(strCurrCalYear," ")(1)
    intYearDiff = CInt(strYear - strCurrCalYear)
    
    If intYearDiff > 0 Then
    	For i = 1 To intYearDiff Step 1
    		objBrowser.WebElement("xpath:=(//span[contains(@class,'angle-right')])[2]").Click
    	Next
    ElseIf intYearDiff < 0 Then
    	For i = -1 To intYearDiff Step -1
    		objBrowser.WebElement("xpath:=(//span[contains(@class,'angle-left')])[2]").Click
    	Next
    End If   
    
    'To Create Month collection
    Set oDesc = Description.Create
    oDesc("class").Value = "date-set__dates"
    Set objDateCollection = objCalendar(0).ChildObjects(oDesc)
        
    'To Change Month
    Set oDesc = Description.Create
    oDesc("xpath").Value = "//month-selector//li[contains(@class,'date-set__date')]"
    Set objDateCollection = objDateCollection(0).ChildObjects(oDesc)
    intTabCount = objDateCollection.count
    For iCount = 0 To intTabCount-1 
        strCalendarvalue = Trim(objDateCollection(iCount).GetRoProperty("innertext"))
        If Instr(strCalendarvalue,strMonth) <> 0 Then
           objDateCollection(iCount).Click
           Exit For
        End If
    Next
    Set objDateCollection = Nothing
    Set oDesc = Nothing    
    
    'To Create date collection
    Set oDesc = Description.Create
    oDesc("class").Value = "day-selector"
    Set objDateCollection = objCalendar(0).ChildObjects(oDesc)    
    
    'To Change Day
    Set oDesc = Description.Create
    oDesc("xpath").Value = "//day-selector//li[contains(@class,'current-month day-selector')]"
    Set objDateCollection = objDateCollection(0).ChildObjects(oDesc)
    intTabCount = objDateCollection.count
    For iCount = 0 To intTabCount-1
        strCalendarvalue = Trim(objDateCollection(iCount).GetRoProperty("innertext"))
        If Instr(strCalendarvalue,strDay) <> 0 Then
           objDateCollection(iCount).Click
           SelectDateFromIDCalendar = True
           Exit Function
        End If
    Next
    Setting.WebPackage("ReplayType") = 1
End Function
'#########################################################################################################################################################################################
'# Function Name: SelectTableRow
'# Description: This function is used to get the row number of the record in table
'# Author: 
'# Date: 22-Sept-2017
'# Input Parameter(s): (ObjtblHeader,ObjtblContent,lstRowData,strTableName,strColumnName,bPagination,objNext)
'#		tblHeader = Data Table header object
'#		tblContent = Data Table Content object
'# 		lstRowData = Header and Row data as array list , Eg:IA NUMBER:APPT-20170922-000019|CREATED ON:22 Sep 2017|SUB STATUS:Approved
'#		strTableName = Name of the Data Table for Reporting, Eg: IA SUmmary table
'# 		bPagination = Set True if the table content has more than 1 row else set to False
'#		objNext = Object table footer Page next (Set True for more than 1 record in table) 
'#		'eg :  SelectTableRow(ObjtblHeader,ObjtblContent,lstRowData,"IA Summary table","IA NUMBER",False,False)
'# Output Parameter(s): True or False
'#########################################################################################################################################################################################
Public function SelectTableRow(ObjtblHeader,ObjtblContent,lstRowData,strTableName,strColumnName,bPagination,objNext)
	Dim bVerifyData,arrColumns,arrValues,intSize
	intSize = Ubound(lstRowData)
	ReDim arrColumns(intSize)
	ReDim arrValues(intSize)
	For iCount = 0 to intSize
		arrTemp = Split(lstRowData(iCount),":")
		arrColumns(iCount) = arrTemp(0)
		arrValues(iCount) = checkNull(arrTemp(1))		
		If arrColumns(iCount) = "CREATED ON" Then
		   arrValues(iCount) = Replace(arrValues(iCount),"#",":")
		End If
	Next
	If bPagination Then
		Do 	
		intRow = getIDRowForColumns(ObjtblHeader,ObjtblContent,arrColumns,arrValues)
		If Not intRow = -1 Then
			Exit Do
		End If
		bNextEnabled = matchStr(objNext.GetROProperty("class"),"disabled")
		If Not bNextEnabled Then
			objNext.Click
			intTablePage = intTablePage + 1
			WaitForIServeLoading
		End If
		Loop while Not bNextEnabled
	Else
		intRow = getIDRowForColumns(ObjtblHeader,ObjtblContent,arrColumns,arrValues)
		'[ : 15th Dec 2017 : Following If Block added to handle scroll issue in Overview Page - Deposites accordion]
		If Instr(strTableName, "Deposits") > 0 Then
			If intRow = -1 Then
				scrollAccordionDownByTagName
				intRow = getIDRowForColumns(ObjtblHeader,ObjtblContent,arrColumns,arrValues)
			End If
		End If
	End If
	If intRow = -1 Then
		LogMessage "WARN","Verification","Expected "& strTableName &" Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" not found in  "& strTableName &" table", False
		SelectTableRow = False
		Exit Function
	Else
		LogMessage "RSLT","Verification","Expected "& strTableName &" Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" found in "& strTableName &" table at Row "&intRow&" on table page number"&intTablePage, True
	End If
	
	   bSelectTableRow = ClickColumnValueInTable(ObjtblHeader,ObjtblContent,intRow,strColumnName)

	WaitForIServeLoading		
	SelectTableRow = bSelectTableRow
End Function

Public Function scrollAccordionDownByID(strId)
	intBodyPixel = Replace(gObjIServePage.RunScript("document.getElementById('"&strId&"').getElementsByClassName('datatable-body')[0].style.height"),"px","")
	intScrollPixel = Replace(gObjIServePage.RunScript("document.getElementById('"&strId&"').getElementsByClassName('datatable-scroll')[0].style.height"),"px","")
	If (intScrollPixel - intBodyPixel) > 5 Then
		gObjIServePage.RunScript("document.getElementById('"&strId&"').getElementsByClassName('datatable-body')[0].scrollTop = 1000000")
		Wait(3)
	End If
End Function

Public Function scrollAccordionDownByTagName
	intBodyPixel = Replace(gObjIServePage.RunScript("document.getElementsByTagName('ngx-datatable')[0].getElementsByClassName('datatable-body')[0].style.height"),"px","")
	intScrollPixel = Replace(gObjIServePage.RunScript("document.getElementsByTagName('ngx-datatable')[0].getElementsByClassName('datatable-scroll')[0].style.height"),"px","")
	If (intScrollPixel - intBodyPixel) > 5 Then
		gObjIServePage.RunScript("document.getElementsByTagName('ngx-datatable')[0].getElementsByClassName('datatable-body')[0].scrollTop = 1000000")
		Wait(3)
	End If
End Function

Public Function scrollPageDown(intNoOfTimesToGoDown)
    x = gObjIServePage.WebElement("xpath:=//nav[@id='isrv_multi_tab_holder']/following-sibling::div").GetROProperty("abs_x")
    y = gObjIServePage.WebElement("xpath:=//nav[@id='isrv_multi_tab_holder']/following-sibling::div").GetROProperty("abs_y")
    For i = 1 To intNoOfTimesToGoDown Step 1
        gObjBrowserWindow.WinObject("regexpwndtitle:=Chrome Legacy Window").Click x+1325,y+190
        gObjBrowserWindow.Type  micDwn
    Next
End Function

Public Function scrollPageUp(intNoOfTimesToGoUp)
    x = gObjIServePage.WebElement("xpath:=//nav[@id='isrv_multi_tab_holder']/following-sibling::div").GetROProperty("abs_x")
    y = gObjIServePage.WebElement("xpath:=//nav[@id='isrv_multi_tab_holder']/following-sibling::div").GetROProperty("abs_y")
    For i = 1 To intNoOfTimesToGoUp Step 1
        gObjBrowserWindow.WinObject("regexpwndtitle:=Chrome Legacy Window").Click x+1325,y+190
        gObjBrowserWindow.Type  micUp
    Next
End Function

'#########################################################################################################################################################################################
'# Function Name: ClickColumnValueInTable
'# Description: This function is used to get the row number of the record in table and Click on the column value of specific row
'# Author: 
'# Date: 22-Sept-2017
'# Input Parameter(s): (ObjtblHeader,ObjtblContent,intRow,strColName)
'#		tblHeader = Data Table header object
'#		tblContent = Data Table Content object
'# 		intRow   = Data table row  number 
'# 		strColName = Column name in table to identify the column index 
'#		'eg :  ClickColumnValueInTable(ObjtblHeader,ObjtblContent,"1","IA NUMBER")	
'# Output Parameter(s): True or False
'#########################################################################################################################################################################################
Public Function ClickColumnValueInTable(ObjtblHeader,ObjtblContent,intRow,strColName)	
	Dim intCol,objCountInCell
	ClickColumnValueInTable = True 
	
	Set objAllRows = getIDAllRows(ObjtblContent)  
	Set objRow = objAllRows(intRow)
	
	'[ : Added below condition to click on link inside the table. E.g "Verify" Link present in Customer Search Result table] 
	If Ucase(Trim(strColName)) = "VERIFY" Then
		Set odesc_Cell = Description.Create
	    odesc_Cell("xpath").value = "//*[@id='verify']"
	    Set objCountInCell = objRow.childObjects(odesc_Cell)  
		objCountInCell(0).Click	  	    
	Else
		Set odesc_Cell = Description.Create
	    odesc_Cell("xpath").value = ".//*[contains(@class,'datatable-body-cell-label')]"
	    Set objCountInCell = objRow.childObjects(odesc_Cell)		    
	    intCol = getIDColIndex(ObjtblHeader,strColName)	
		objCountInCell(intCol).click	    
	End If
		
	WaitForIServeLoading
	
	If Err.Number <> 0 Then
		ClickColumnValueInTable = False
		LogMessage "RSLT","Verification","Failed to Click Column Name : "&strColName ,False
	End If

	Set objAllRows = Nothing 
	Set objRow = Nothing
	Set odesc_Cell = Nothing
	set objCountInCell = Nothing
End Function
'#############################################################################################
'# Function Name: SetObjectFirstPage
'# Description: This function is used to get the object of the First Arrow in Page footer 
'# Author: 
'# Date: 13-Oct-2017
'# Input Parameter(s): (ObjPageFooter)
'#		ObjPageFooter = Object id for Page footer
'# Output Parameter(s): Object id of first double Arrow 
'###############################################################################################
Public Function SetObjectFirstPage(ObjPageFooter)
  Set oDesc1 = Description.Create
      oDesc1("micclass").Value = "WebElement"
	  oDesc1("class").Value = "pager"
  Set PageFooter = ObjPageFooter.ChildObjects(oDesc1)
  Set oPreviousDoubleArrow = Description.Create
	oPreviousDoubleArrow("micclass").Value = "WebElement"
	oPreviousDoubleArrow("class").Value = "firstPage.*"
  Set SetObjectFirstPage = PageFooter(0).ChildObjects(oPreviousDoubleArrow)	
End Function	
'###########################################################################################
'# Function Name: SetObjectPreviousPage
'# Description: This function is used to get the object of the previous arrow in Page footer
'# Author: 
'# Date: 13-Oct-2017
'# Input Parameter(s): (ObjPageFooter)
'#		ObjPageFooter = Object id for Page footer
'# Output Parameter(s): Object id of First Single Arrow 
'############################################################################################
Public Function SetObjectPreviousPage(ObjPageFooter)
  Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "pager"
  Set PageFooter = ObjPageFooter.ChildObjects(oDesc1)	
  Set oPreviousArrow = Description.Create
	oPreviousArrow("micclass").Value = "WebElement"
	oPreviousArrow("class").Value = "previousPage.*"
  Set SetObjectPreviousPage = PageFooter(0).ChildObjects(oPreviousArrow)	
End Function
'###########################################################################################
'# Function Name: SetObjectNextPage
'# Description: This function is used to get the object of the Next arrow of Page footer
'# Author: 
'# Date: 13-Oct-2017
'# Input Parameter(s): (ObjPageFooter)
'#		ObjPageFooter = Object id for Page footer
'# Output Parameter(s): Object id of Next Arrow 
'############################################################################################
Public Function SetObjectNextPage(ObjPageFooter)
  Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "pager"
  Set PageFooter = ObjPageFooter.ChildObjects(oDesc1)
  Set oNextArrow = Description.Create
	oNextArrow("micclass").Value = "WebElement"
	oNextArrow("class").Value = "nextPage.*"
  Set SetObjectNextPage = PageFooter(0).ChildObjects(oNextArrow)
End Function	
'###########################################################################################
'# Function Name: SetObjectLastPage
'# Description: This function is used to get the object of the last arrow of Page footer
'# Author: 
'# Date: 13-Oct-2017
'# Input Parameter(s): (ObjPageFooter)
'#		ObjPageFooter = Object id for Page footer
'# Output Parameter(s): Object id of Next double Arrow 
'############################################################################################	
Public Function SetObjectLastPage(ObjPageFooter)
 Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").Value = "pager"
 Set PageFooter = ObjPageFooter.ChildObjects(oDesc1)
 Set oNextDoubleArrow = Description.Create
	oNextDoubleArrow("micclass").Value = "WebElement"
	oNextDoubleArrow("class").Value = "lastPage.*"
 Set SetObjectLastPage = PageFooter(0).ChildObjects(oNextDoubleArrow)
End Function
'###########################################################################################
'# Function Name: SetObjPanelRow
'# Description: This function is used to get the object Panel row in customer overview Page 
'# Author: 
'# Date: 16-Oct-2017
'# Input Parameter(s): (ObjPanelHeader)
'#		ObjPageFooter = Object id for whole panel header
'# Output Parameter(s): Object id of Panel Row
'############################################################################################
Public Function SetObjPanelRow(ObjPanelHeader)	
Set oDesc1 = Description.Create
	oDesc1("micclass").Value = "WebElement"
	oDesc1("class").value= ".*closed-inter-class"
Set SetObjPanelRow = ObjPanelHeader.ChildObjects(oDesc1)	 
End Function
'###########################################################################################
'# Function Name: SetObjStatus
'# Description: This function is used to get the object id of Status displayed in Overview Page 
'# Author: 
'# Date: 16-Oct-2017
'# Input Parameter(s): (ObjPanelHeader)
'#		ObjPageFooter = Object id for whole panel header
'# Output Parameter(s): Object id of Status Field
'############################################################################################
Public Function SetObjStatus(ObjPanelHeader)
 Set oDesc1 = Description.Create
	 oDesc1("micclass").Value = "WebElement"
	 oDesc1("class").value= "status_automation.*"
 Set SetObjStatus = ObjPanelHeader.ChildObjects(oDesc1)
End Function
'###########################################################################################
'# Function Name: SetObjPanelRow
'# Description: This function is used to get the object id of SubStatus displayed in Overview Page   
'# Author: 
'# Date: 16-Oct-2017
'# Input Parameter(s): (ObjPanelHeader)
'#		ObjPageFooter = Object id for whole panel header
'# Output Parameter(s): Object id of Sub Status field
'############################################################################################
Public Function SetObjSubStatus(ObjPanelHeader)
 Set oDesc1 = Description.Create
	 oDesc1("micclass").Value = "WebElement"
	 oDesc1("class").value= "subStatus_automation.*"
 Set SetObjSubStatus = ObjPanelHeader.ChildObjects(oDesc1)
End Function
'###########################################################################################
'# Function Name: SetObjTriplets
'# Description: This function is used to get the object id of triples (Related TO, Type, SubType) displayed in Overview Page  
'# Author: 
'# Date: 16-Oct-2017
'# Input Parameter(s): (ObjPanelHeader)
'#		ObjPageFooter = Object id for whole panel header
'# Output Parameter(s): Object id of triplets (Related TO, Type & Sub-Type) 
'############################################################################################
Public Function SetObjTriplets(ObjPanelHeader)
 Set oDesc1 = Description.Create
	 oDesc1("micclass").Value = "WebElement"
	 oDesc1("class").value= "triplet_automation"
 Set SetObjTriplets = ObjPanelHeader.ChildObjects(oDesc1)
End Function
'###########################################################################################
'# Function Name: SetObjNumber
'# Description: This function is used to get the object id of the IA Number or SR Number displayed in Overview Page  
'# Author: 
'# Date: 16-Oct-2017
'# Input Parameter(s): (ObjPanelHeader)
'#		ObjPageFooter = Object id for whole panel header
'# Output Parameter(s): Object id of IA or SR Number 
'############################################################################################
Public Function SetObjNumber(ObjPanelHeader)
 Set oDesc1 = Description.Create
	 oDesc1("micclass").Value = "WebElement"
	 oDesc1("class").value= "activityId_automation"
 Set SetObjNumber = ObjPanelHeader.ChildObjects(oDesc1)
End Function
'###########################################################################################
'# Function Name: SetObjCreatedDate
'# Description: This function is used to get the object id of the Created Dated displayed in Overview Page  
'# Author: 
'# Date: 16-Oct-2017
'# Input Parameter(s): (ObjPanelHeader)
'#		ObjPageFooter = Object id for whole panel header
'# Output Parameter(s): Object id of Created Date field 
'############################################################################################
Public Function SetObjCreatedDate(ObjPanelHeader)
 Set oDesc1 = Description.Create
	 oDesc1("micclass").Value = "WebElement"
	 oDesc1("class").value= "createdDate_automation"
 Set SetObjCreatedDate = ObjPanelHeader.ChildObjects(oDesc1)
End Function
'###########################################################################################
'# Function Name: getDefaultSelectedRadioButton
'# Description: This function is used to get default selected Radio button in New SR page. 
'# Author: 
'# Date: 17-Nov-2017
'# Input Parameter(s): (objRadioGrp) : Radio group object
'#		strExpdBtnToBeSelected = expected button to be selected by default
'#		strSelectionLbl = Label name under which button to be selected
'# E.g : 
'#	Set objRadioGrp = gObjIServePage.WebElement("xpath:=//*[@id='newSR_status_radio']")
'#	getDefaultSelectedRadioButton objRadioGrp,"Closed","Status"
	
'#	Set objRadioGrp = gObjIServePage.WebElement("xpath:=//*[@id='newSR_pripority_group']")
'#	getDefaultSelectedRadioButton objRadioGrp,"High", "Priority"
	
'#	Set objRadioGrp = gObjIServePage.WebElement("xpath:=//*[@id='newSR_followupRequired_group']")
'#	getDefaultSelectedRadioButton objRadioGrp,"Yes", "Followup Required"

'# Output Parameter(s): True or False

'############################################################################################
Public Function getDefaultSelectedRadioButton(objRadioGrp,strExpdBtnToBeSelected,strSelectionLbl)
	Set oDesc = Description.Create
	oDesc("html tag").value = "md-radio-button"
	Set oChild = objRadioGrp.ChildObjects(oDesc)
	iCount = oChild.Count-1
	For i = 0 To iCount Step 1
		intClass = oChild(i).GetROProperty("class")
		If Instr(1,intClass,"checked") > 0 Then
			strActlBtnSelected = Trim(oChild(i).GetROProperty("innertext"))
			If strActlBtnSelected = strExpdBtnToBeSelected  Then
				LogMessage "RSLT","Verification","As Expected By Default: " &strExpdBtnToBeSelected& " button is selected for "&strSelectionLbl, True
				getDefaultSelectedRadioButton = True
				Exit For
			Else
				LogMessage "WARN","Verification","Failed: By Default: " &strActlBtnSelected& " button is selected for "&strSelectionLbl&", Expected default button Selection: "&strExpdBtnToBeSelected, False
				getDefaultSelectedRadioButton = False
			End If
		End If
	Next
	Set oChild = Nothing
	Set oDesc =	Nothing
End Function
'############################################################################################
'# Function Name: VerifyAccordiondisplay
'# Description: This function is used to verify if there is any error displayed below any accordions  
'# Author:  
'# Date: 24-Nov-2017
'# Input Parameter(s): (objAccordionSection) : Accordion object
'#		strProduct = Accordion name  {used for Reporting}
'# E.g : 
'# Set objRadioGrp = gObjIServePage.WebElement("xpath:=//*[@id='newSR_status_radio']")
'# VerifyAccordiondisplay objAccordionSection,"Credit Cards"
'# Output Parameter(s): True or False
'############################################################################################
Public Function VerifyAccordiondisplay(objAccordionSection,strProduct)
	bverifyAccordions = False	
	Set odesc_Cell = Description.Create
	odesc_Cell("xpath").value = ".//*[@id='errorTitle_h6']"
	
	Set objErrortext = objAccordionSection.ChildObjects(odesc_Cell)	  
	
	If objErrortext.Count = 0  Then
	   LogMessage "RSLT","Verification","No Error Found for Accordion Selected:"&strProduct ,True
	   bverifyAccordions = True 
	Else
		 If Instr(Ucase(objErrortext(0).GetROProperty("innertext")),"ERROR") > 0 Then 
			LogMessage "WARN","Verification","Error Found for Accordion Selected:"&strProduct, False
			bverifyAccordions = False 
		 Else
			LogMessage "WARN","Verification","Error Found for Accordion Selected :"&strProduct& "not displayed with proper text", False
			bverifyAccordions = False
		 End IF			
		 
	End If		   
		
	VerifyAccordiondisplay = bverifyAccordions
	
	Set odesc_Cell = Nothing
	Set objErrortext = Nothing
End Function

Public Function ExpandSingleAccordion(ObjAccordionGroup,strAccordion)
	
	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-title"

	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count
	
	Set oExpandIcon = Description.Create
	oExpandIcon("xpath").Value = "//*[@class='expandoIcon']/md-icon"
	oExpandIcon("visible").Value = "True"
	
	Set oChildExpandIcon = ObjAccordionGroup.ChildObjects(oExpandIcon)
	
	For j = 0 To intChildCnt-1 Step 1
		strActAccordionHeader = Trim(Replace(oDivChild(j).GetRoProperty("innertext"),"Open | Closed",""))
		If Instr(Trim(strAccordion),strActAccordionHeader) > 0 Then
		  If Instr(oChildExpandIcon(j).GetRoProperty("class"), "plus") > 0 Then
			 	oChildExpandIcon(j).Click
		   		Exit For
			End If
		End If
	Next
	
	ExpandSingleAccordion = intChildCnt
	Set oDiv = Nothing
	Set oDivChild = Nothing
	Set oExpandIcon = Nothing
	Set oChildExpandIcon = Nothing
End Function

Public Function CollapseSingleAccordion(ObjAccordionGroup,strAccordion)
	bClickExpandIcon = False
	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-title"

	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count
	
	Set oExpandIcon = Description.Create
	oExpandIcon("xpath").Value = "//*[@class='expandoIcon']/md-icon"
	oExpandIcon("visible").Value = "True"
	
	Set oChildExpandIcon = ObjAccordionGroup.ChildObjects(oExpandIcon)
	
	For j = 0 To intChildCnt-1 Step 1
		strActAccordionHeader = Trim(Replace(oDivChild(j).GetRoProperty("innertext"),"Open | Closed",""))
		If Instr(Trim(strAccordion),strActAccordionHeader) > 0 Then
		  If Not Instr(oChildExpandIcon(j).GetRoProperty("class"), "plus") > 0 Then
			 	oChildExpandIcon(j).Click
				bClickExpandIcon = True
		   		Exit For
			End If
		End If
	Next
	
	CollapseSingleAccordion = bClickExpandIcon
	
	Set oDiv = Nothing
	Set oDivChild = Nothing
	Set oExpandIcon = Nothing
	Set oChildExpandIcon = Nothing
End Function

Public Function CollapseMultipleAccordions(ObjAccordionGroup,ByVal lstAccordion)
	bClickExpandIcon = False
	
	lstAccordion = Split(lstAccordion,"|")
	
	Set oDiv = Description.Create
	oDiv("Class").Value = "panel-title"

	Set oDivChild = ObjAccordionGroup.ChildObjects(oDiv)
	intChildCnt = oDivChild.Count

	Set oExpandIcon = Description.Create
	oExpandIcon("xpath").Value = "//*[@class='expandoIcon']/md-icon"
	oExpandIcon("visible").Value = "True"

	Set oChildExpandIcon = ObjAccordionGroup.ChildObjects(oExpandIcon)

	For i = 0 To Ubound(lstAccordion)
		strAccordion = lstAccordion(i)
		For j = 0 To intChildCnt-1 Step 1
			strActAccordionHeader = Trim(Replace(oDivChild(j).GetRoProperty("innertext"),"Open | Closed",""))
			If Instr(Trim(strAccordion),strActAccordionHeader) > 0 Then
				If Not Instr(oChildExpandIcon(j).GetRoProperty("class"), "plus") > 0 Then
				 	oChildExpandIcon(j).Click
			   		Exit For
				End If
			End If
		Next	
	Next

	CollapseMultipleAccordions = bClickExpandIcon

	Set oDiv = Nothing
	Set oDivChild = Nothing
	Set oExpandIcon = Nothing
	Set oChildExpandIcon = Nothing
End Function

Public Function IndonesiaCustomerVerification(iNoOfIdentQues,iNoOfAuthQues)
 	Set ObjAccordionGroup = gObjIServePage.WebElement("xpath:=//*[@id='verificationTab']")
 	gObjIServePage.WebButton("xpath:=.//*[@id='verification_button']").Click
 	WaitForIServeLoading
 	selectIdentificationQuestions ObjAccordionGroup,iNoOfIdentQues
 	selectAuthenticationQuestions ObjAccordionGroup,iNoOfAuthQues
 	gObjIServePage.WebButton("xpath:=(.//*[@id='maverification_auth']/following-sibling::div//button)[2]").Click
 	WaitForIServeLoading
 	Set ObjAccordionGroup = Nothing
 End Function
 
Public Function selectIdentificationQuestions(ObjAccordionGroup,iNoOfIdentQues)
 	
 	Set oDescMD = Description.Create
 	oDescMD("xpath").Value = ".//accordion[@id='maverification_identi']//div[contains(@class,'panel-collapse')]"
 	oDescMD("Visible").Value = "True"
 	
 	Set oDesc = Description.Create	
 	oDesc("html tag").Value = "md-radio-button"
 	oDesc("class").Value = "green mat-radio-button"
 	oDesc("Visible").Value = "True"
 	
 	Wait(4)
 	
 	Set oChild = gObjIServePage.WebElement(oDescMD).ChildObjects(oDesc)
 	Setting.WebPackage("ReplayType") = 2
 	For i = 0 To iNoOfIdentQues-1 Step 1
 		oChild(i).Click
 		WaitForIServeLoading
 	Next
 	Setting.WebPackage("ReplayType") = 1
 	
 	Set oChild = Nothing
 	Set oDesc = Nothing
 	Set oDescMD = Nothing
 	
 	CollapseSingleAccordion ObjAccordionGroup,"Identification"
 End Function
 
 Public Function selectAuthenticationQuestions()
 	Set objTop = gObjIServePage.WebElement("xpath:=(.//accordion[@id='maverification_auth']//div[contains(@class,'panel-collapse')]//md-card)[4]","visible:=True")
 
 	Set oChild = ClickAuthSectionQuestion(objTop)

 	Setting.WebPackage("ReplayType") = 2
	 	For i = 0 To 0 Step 1
	 		oChild(i).Click
	 		WaitForIServeLoading
	 	Next
	 	Setting.WebPackage("ReplayType") = 1
	
 	Set objTop = Nothing
 	Set oChild = Nothing
 End Function
 
 Public Function ClickAuthSectionQuestion(objSection) 
 	Set oDesc = Description.Create	
 	oDesc("html tag").Value = "md-radio-button"
 	oDesc("class").Value = "green mat-radio-button"
 	oDesc("Visible").Value = "True"
 	
 	Set oChild = objSection.ChildObjects(oDesc)
 	
 	Set ClickAuthSectionQuestion = oChild
 	
 	Set oChild = Nothing
 	Set oDesc = Nothing
 End Function
 
 Public Function setCustomQuestion()  
 	Set objTop = gObjIServePage.WebElement("xpath:=(.//accordion[@id='maverification_auth']//div[contains(@class,'panel-collapse')]//md-card)[6]","visible:=True")
 	Set oChild = ClickAuthSectionQuestion(objTop)
 	gObjIServePage.WebEdit("xpath:=.//TEXTAREA[@id='MA_Custom_DD']").Set "Test Question"
 	WaitForIServeLoading
 	Wait(3)
 	Setting.WebPackage("ReplayType") = 2
 	For i = 0 To 0 Step 1
 		oChild(i).Click
 		WaitForIServeLoading
 	Next
 	Setting.WebPackage("ReplayType") = 1 

 	Set objTop = Nothing
 	Set oChild = Nothing
 End Function

'############################################################################################
'# Function Name: VerifytblPagination
'# Description: This function is used to verify the pagination for the tables 
'# Author:  
'# Date: 20-Dec-2017
'# Input Parameter(s): (Objtbl) : Table object
'#		strcolumnname = Column name that has the data for all the rows
'#      NoOfRows = Number of Rows per Page
'#      strAccordian = Accordianname
'# E.g : 
'# Set Objtbl = gObjIServePage.WebElement("xpath:=//*[@id='bankTransHist_TransHist_table']")
'# strcolumnname = Transaction Date
'# VerifytblPagination Objtbl,strcolumnname,5,'Transaction Details"
'# Output Parameter(s): True or False
'############################################################################################
 
 Public Function VerifytblPagination(Objtbl,strcolumnname,NoOfRows,strAccordian)
	Dim iCheck
	'Browser("title:=I.*").Page("title:=I.*")
	bverifypagination = False
	htmlid = Trim(Objtbl.GetROProperty("html id"))
	If Not (htmlid="") Then
		Set tblHeader = Description.Create
	    tblHeader("xpath").Value = "//*[@id='"+htmlid+"']//datatable-header"
	    Set tblBody = Description.Create
	    tblBody("xpath").Value = "//*[@id='"+htmlid+"']//datatable-body"
	    Set tblPager = Description.Create
	    tblPager("xpath").Value = "//*[@id='"+htmlid+"']//datatable-pager"
	 Else
	 	Set tblHeader = Description.Create
	    tblHeader("xpath").Value = "(//ngx-datatable//datatable-header)[1]"
	    Set tblBody = Description.Create
	    tblBody("xpath").Value = "(//ngx-datatable//datatable-body)[1]"
	    Set tblPager = Description.Create
	    tblPager("xpath").Value = "(//ngx-datatable//datatable-pager)[1]"
	End If
	
	Set tblhdr = gObjIServePage.ChildObjects(tblHeader)
	Set tblbdy = gObjIServePage.ChildObjects(tblBody)
	Set TblFooter = gObjIServePage.ChildObjects(tblPager)
	intRecordCount = getIDRecordsCountForColumn(tblhdr(0),tblbdy(0),strcolumnname)	
	intRecordCount = intRecordCount-1' Adding one becoz the table has the scroll will shown +1 records always
	iCheck  = Cint(NoOfRows)
	Set ObjFirstPg = SetObjectFirstPage(TblFooter(0))
	Set ObjPreviousPg = SetObjectPreviousPage(TblFooter(0))
	Set ObjNextPg = SetObjectNextPage(TblFooter(0))
	Set ObjLastPg = SetObjectLastPage(TblFooter(0))
	bFirstPageDiasbled = matchStr(ObjFirstPg(0).GetROProperty("class"),"disabled")
	bPreviousPageDisabled = matchStr(ObjPreviousPg(0).GetROProperty("class"),"disabled")
	bNextPageDisabled = matchStr(ObjNextPg(0).GetROProperty("class"),"disabled")
	bLastPageDiasbled = matchStr(ObjLastPg(0).GetROProperty("class"),"disabled")
	
	If intRecordCount = iCheck Then			
		  If (bFirstPageDiasbled AND bPreviousPageDisabled) = True Then
		  	LogMessage "RSLT","Verification","Previous Arrows are disabled for first page as expected in "+strAccordian, True 
		  	bverifypagination = True
			Else 
			LogMessage "WARN","Verification","Previous Arrows are not disabled for first page as expected in "+strAccordian, False
			bverifypagination = False
		  End If	
		  If bLastPageDiasbled = False Then 
		 	 Setting.WebPackage("ReplayType") = 2
		  	 ObjLastPg(0).Click
		  	 wait(1)
		  	 bLastPageDiasbled = matchStr(ObjLastPg(0).GetROProperty("class"),"disabled")
		  	 bNextPageDisabled = matchStr(ObjNextPg(0).GetROProperty("class"),"disabled")
			 If bLastPageDiasbled AND bNextPageDisabled Then
				LogMessage "WARN","Verification","Next and Last Arrows are disabled as expected in "+strAccordian, True
				bverifypagination = True
			 Else 
			    LogMessage "WARN","Verification","Next and Last Arrows are not disabled as expected in "+strAccordian, False
				bverifypagination = False
			 End If
		  End If
	ElseIf intRecordCount < iCheck  Then	
		If bNextPageDisabled AND bLastPageDiasbled AND bFirstPageDiasbled AND bPreviousPageDisabled  Then
			LogMessage "RSLT","Verification","Next and Previous Arrows are disabled as expected when number of records is less than "&iCheck&" in "+strAccordian, True 
			bverifypagination = True
		Else 
			LogMessage "WARN","Verification","Next and Previous Arrows are not disabled as expected when number of records is less than "&iCheck&" in "+strAccordian, False
			bverifypagination = False
		End If		
	End IF
	Setting.WebPackage("ReplayType") = 1
	 
	VerifytblPagination = bverifypagination
	 
	Set objFristPage = Nothing
	Set objNextPage = Nothing
	Set objLastPage = Nothing
	Set objPreviousPage = Nothing
End Function

'###############################################################################################
'# Function Name: CustomerSearch
'# Description: Function to Search a customer from Dashboard page
'# Author: 
'# Date: 11-Sept-2017
'# Input Parameter(s): strSearchType,strSearchVal
'#		strSearchType = CIF Number, Email ID, Phone Number etc
'#		strSearchVal = Respective Search Type value 
'# Output Parameter(s): True or False
'###############################################################################################
Public Function CustomerSearch(strSearchType,strSearchVal)

	SetValue coDashboard_Page.txtSearchType,strSearchType,"Search ID Type"
	SetValue coDashboard_Page.txtSearchVal,strSearchVal,"Search Value"
	
	If coDashboard_Page.btnSearch.GetROProperty("disabled") = 0 Then
		ClickOnObject coDashboard_Page.btnSearch,"Customer Search Button"
	Else
		LogMessage "WARN","Verification","Search Button is disabled, Unable to perform search operation. ",False
		CustomerSearch = False
	End If
	
	WaitForIServeLoading
	
	If Err.Number = 0 Then
		LogMessage "RSLT","Verification","Search criteria entered as expected",True
		CustomerSearch = True
	Else
		LogMessage "WARN","Verification","Search criteria is not entered as expected",False
		CustomerSearch = False
	End If

End Function


