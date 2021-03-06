Dim bcUserProfile
Set bcUserProfile = cUserProfile()

'This is the Screen UserProfile

Public Function cUserProfile()
    Set cUserProfile = New clsUserProfile
End Function

Class clsUserProfile

        Private Sub Class_Initialize()
        End Sub

        Private Sub Class_Terminate()
        End Sub

        '******************************** Object Initialization ******************************************************************

       Public Function pageExists()
           If  (lnkConfigureUser().exist(1)) Then
               pageExists = true
			 '  LogMessage "RSLT", "Verification", "User Profile screen exists.", True
            else
              pageExists = false
			  'LogMessage "RSLT", "Verification", "User Profile screen does not exists.", False
           End If
        End Function
        
        Public Function ReportsPageExists()
           If (txtAuditReportFromDate().exist(0)) Then
           		ReportsPageExists = True
			LogMessage "RSLT", "Verification", "Reports Page screen exists.", True           		
           	Else
           		ReportsPageExists = False
           		LogMessage "RSLT", "Verification", "Reports Page screen does not exists.", False
           End If
        	
        End Function

        Public Function UserProfile()
           'Set UserProfile = 
        End Function

		'For Search window
        Public Function txt1BankID_Search()
           Set txt1BankID_Search = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txt1BankIDSearch")
        End Function

        Public Function btnSearch()
           Set btnSearch = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnSearch")
        End Function

        Public Function lnkConfigureUser()
           Set lnkConfigureUser = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebButton("btnConfigureUser")
        End Function

        Public Function tblSearchResultHeader()
           Set tblSearchResultHeader = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebTable("tblSearchResultHeader")
        End Function

        Public Function tblSearchResultContent()
           Set tblSearchResultContent = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebTable("tblSearchResultContent")
        End Function

        Public Function tblDirectReportsHeader()
           Set tblDirectReportsHeader = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebTable("tblDirectReportsHeader")
        End Function

        Public Function tblDirectReportsContent()
           Set tblDirectReportsContent = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebTable("tblDirectReportsContent")
        End Function

		 Public Function dlgUserConfigure()
           Set dlgUserConfigure = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("dlgUserConfuguration")
        End Function

		 Public Function lnkNext()
           Set lnkNext = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lnkNext")
        End Function
		Public Function lnkNext1()
           Set lnkNext1= Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lnkNext1")
        End Function
        Public Function lnkPrevious()
           Set lnkPrevious = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lnkPrevious")
        End Function
		'-----------------------------------------------

        Public Function lblUserConfuguration()
           'Set lblUserConfuguration = 
        End Function

        Public Function txt1BankID()
           Set txt1BankID = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txt1BankID")

        End Function

        Public Function txtLANID()
           Set txtLANID = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txtLANID")

        End Function

        Public Function txtName()
           Set txtName = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txtName")

        End Function

        Public Function ddlRole()
           Set ddlRole = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lstRole")
        End Function

        Public Function txtManager1BankID()
           Set txtManager1BankID = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txtManager1BankID")
        End Function
	    Public Function lblConfigureErrorMsg()
           Set lblConfigureErrorMsg = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblConfigureErrorMsg")
        End Function

		 Public Function lblApproverErrorMsg()
           Set lblApproverErrorMsg = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblApproverErrorMsg")
        End Function
		 Public Function lblSearchErrorMsg()
           Set lblSearchErrorMsg = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblSearchErrorMsg")
        End Function

        Public Function txtManagerLANID()
           Set txtManagerLANID = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txtManagerLANID")
        End Function

        Public Function ddlGroup()
           Set ddlGroup = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lstGroup")
        End Function

        Public Function ddlChannel()
           Set ddlChannel = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lstChannel")
        End Function

        Public Function ddlLocation()
           Set ddlLocation = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lstLocation")

        End Function
		'------------------

        Public Function rbtC3User()
           'Set rbtC3User = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebRadioGroup("rbtC3User")
			Set rbtC3User = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("rbtC3User")

        End Function

		 Public Function rbtStaffAccess()
           'Set rbtStaffAccess = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebRadioGroup("rbtStaffAccess")
			Set rbtStaffAccess =Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("rbtStaffAccess")

        End Function

		 Public Function rbtStatus()
           'Set rbtStatus = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebRadioGroup("rbtStatus").Select
			Set rbtStatus = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("rbtStatus")

        End Function
		'------------------

        Public Function lblCreatedDate()
           Set lblCreatedDate = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblCreateDate")

        End Function

        Public Function lblCreatedBy()
           Set lblCreatedBy = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblCreatedBy")

        End Function

        Public Function lblLastApprovedDate()
           Set lblLastApprovedDate = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblLastApprovedDate")

        End Function

        Public Function lblLastApprovedBy()
           Set lblLastApprovedBy = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblLastApprovedBy")

        End Function

        Public Function lblLastUpdatedDate()
           Set lblLastUpdatedDate = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblLastUpdatedDate")

        End Function

        Public Function lblLastUpdatedBy()
           Set lblLastUpdatedBy = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblLastUpdatedBy")
        End Function

        Public Function btnSubmit()
           Set btnSubmit = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnSubmit")
        End Function

		Public Function btnOK()
           Set btnOK = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnOK")
        End Function

        Public Function btnCancel()
           Set btnCancel = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnCancel")
        End Function

		Public Function btnimgverify()
           Set btnimgverify = Browser("Browser_iCall_Home").Page("iCall_UserProfile").Image("imgverify_new")
        End Function


		Public Function lnkApprove()
           Set lnkApprove =Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lnkApprove")
        End Function

		Public Function lnkConfigure()
           Set lnkConfigure =Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lnkConfigure")
		End Function

		Public Function btnApprove()
           Set btnApprove =Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnApprove")
		End Function

		Public Function btnReject()
           Set btnReject =Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnReject")
		End Function

		
		'------------END View/Edit User configuration-------------
		'Reports related
		 Public Function lnkReports()
           Set lnkReports = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebButton("btnReports")
        End Function
       
		Public Function lblReportNote()
           Set lblReportNote = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblReportNote")
        End Function

        Public Function txtAuditReportFromDate()
           Set txtAuditReportFromDate =Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txtFrom_Date")
        End Function

        Public Function txtAuditReportToDate()
           Set txtAuditReportToDate = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebEdit("txtTo_Date")
        End Function

        Public Function lblReportErrorMsg()
           Set lblReportErrorMsg = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("lblReportErrorMsg")
        End Function

        Public Function txtDeactivatedUserReportToDate()
           'Set txtDeactivatedUserReportToDate = 
        End Function

        Public Function btnGoAuditReport()
           Set btnGoAuditReport= Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnGo")
        End Function

        Public Function btnGoDeactivatedUserReport()
           'Set btnGoDeactivatedUserReport = 
        End Function

        Public Function btnDownload_Audit()
           Set btnDownload_Audit = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnDownload_Audit")
        End Function

        Public Function btnDownload_InvalidUsers()
           Set btnDownload_InvalidUsers = Browser("Browser_iCall_Home").Page("iCall_UserProfile").WebElement("btnDownload_InvalidUsers")
        End Function

		Public Function dlgFileDownload()
           Set dlgFileDownload = Browser("Browser_iCall_Home").Dialog("dlgFileDownload")
        End Function
		Public Function btnFileDownload_Save()
           Set btnFileDownload_Save =Browser("Browser_iCall_Home").Dialog("dlgFileDownload").WinButton("btnFileDownload_Save")
        End Function
		Public Function lblFileName()
           Set lblFileName = Browser("Browser_iCall_Home").Dialog("dlgFileDownload").WinObject("lblFileName")
        End Function
			
		
        Public Function btn_SaveDlg()
        	Set btn_SaveDlg = Dialog("DlgSaveDialog").Dialog("Save As").WinButton("btn_SaveDlg")
        End Function
        
        Public Function txt_FileName()
        	Set txt_FileName = Dialog("DlgSaveDialog").Dialog("Save As").WinEdit("txt_FileName")

        End Function
        
        Public Function btn_OK()
        	Set btn_OK = Dialog("DlgSaveDialog").Dialog("Save As").Dialog("Save As").WinButton("btn_OK")
        End Function


        '******************************** End of Object Initialization ******************************************************************

        '*****************************Buttons & Link Clicks on the Page **********************************************************
        Public Sub clickSearch()
            btnSearch().Click
        End Sub

        Public Sub clickConfigureUser()
            lnkConfigureUser().Click
        End Sub

        Public Sub clickSubmit()
            btnSubmit().Click
        End Sub

        Public Sub clickCancel()
            btnCancel().Click
        End Sub

        Public Sub clickGoAuditReport()
            btnGoAuditReport().Click
        End Sub

        Public Sub clickGoDeactivatedUserReport()
            btnGoDeactivatedUserReport().Click
        End Sub

        Public Sub clickDownloadAuditReport()
            btnDownload_Audit().Click
        End Sub

        Public Sub clickDownloadDeactivatedUserReport()
            btnDownloadDeactivatedUserReport().Click
        End Sub

        '*****************************End of Buttons & Link Clicks on the Page **********************************************************

        '*****************************Function on the Screen **********************************************************

		'Function to search customer
		Public function searchUserProfile(strUserName, lstSearchResult,strDirectReportsCount, lstlstDirectReports, strAction,strSearchError,strApproverError)
		    Dim arrSearchResultCol, intRow
			Dim bSearchUserProfile:bSearchUserProfile=true
		   	ReDim arrSearchResultCol (4)
			lstSearchResultCol = array("Name","1Bank ID", "Role","Status", "Action")
			
			lnkConfigureUser().click()
			waitForIcallLoading
				If Not IsNull(strUserName) Then
                    txt1BankID_Search().set strUserName
                End If
			btnSearch().click()
			waitForIcallLoading
			'Verifiy the search result		
			If not IsNull(lstSearchResult) Then
				 intRow = getRowForColumns (tblSearchResultHeader,tblSearchResultContent,lstSearchResultCol, lstSearchResult)
				if intRow=-1 then
                    	LogMessage "WARN","Verification","Expected record "&ArrayToString(lstSearchResult,",")&" does not displayed in search result table",false
						bSearchUserProfile = false
				else
					LogMessage "RSLT","Verification","Expected record "&ArrayToString(lstSearchResult,",")&" successfully displayed in search result table",true
				End If
				
			End If
			If not isNull(strSearchError) Then
				   If Not VerifyInnerText( lblSearchErrorMsg(), strSearchError, "Error Message") Then
								bSearchUserProfile = False
					 else
							bSearchUserProfile=true
					 End If
					 searchUserProfile=bSearchUserProfile
					 Exit Function
			End If
				
			' if  lstlstDirectReports is not NULL search & verify the 'Direct Report' table
			If Not IsNull(strDirectReportsCount) Then
				
				intRecordCount=getRecordsCountForColumn (tblDirectReportsHeader,tblDirectReportsContent,"Name")
				If intRecordCount<=10 Then
					LogMessage "RSLT","Verification","Number of records displayed per page matched with expected. Expected Count is less than or equal to 10", true
					bSearchUserProfile=true
				Else
					LogMessage "WARN","Verification","Number of records displayed per page is more than 5 record. Expected Count is less than or equal to 10, Actual "&intRecordCount, false
					bSearchUserProfile=false
				End If

				If Not verifySummaryCount(tblDirectReportsHeader,tblDirectReportsContent,"Name",strDirectReportsCount,"Direct Reports",true)	 then
					bSearchUserProfile=false
				End If
			else
				If tblDirectReportsHeader.Exist(0) then
					LogMessage "WARN","Verification","Direct Reports table displayed unexpectadly",false
					bSearchUserProfile=false
				End If
			End If
             If Not IsNull(lstlstDirectReports) Then
				 '(tblDirectReportsHeader,tblDirectReportsContent,arrSearchResultCol, lstlstDirectReports(icRow))
				If not verifyTableContentList(tblDirectReportsHeader, tblDirectReportsContent, lstlstDirectReports, "Direct Report'", true) then
						bSearchUserProfile = false
					End If
				 End If

			'Default search & verify the 'Search Result' table
			If not isNull(strAction) Then
				If intRow <> -1 Then
					clickVaddinLink_tblCell tblSearchResultHeader,tblSearchResultContent,intRow, "Action"
					WaitForICallLoading
					If not dlgUserConfigure.Exist(0) Then
							bSearchUserProfile = false
							LogMessage "WARN","Verification","User Configure dialog does not displayed as expected",false
					 else
						strTemp=dlgUserConfigure.GetROProperty("outertext")
						If Instr(1,strTemp,"Error")<>0 Then
							bError=True
								LogMessage "WARN","Verification","Error dialog displayed",true
						else
								LogMessage "WARN","Verification","User Configure dialog displayed successfully",true
						End If
						
					End If
					'bSearchUserProfile = true
				End If
			End If
		If Not iSNull(strApproverError) Then
			 If Not VerifyInnerText( lblApproverErrorMsg(), strApproverError, "Approver Error Message") Then
					bSearchUserProfile = False
			 else
					bSearchUserProfile=true
			 End If
			 	btnOK.Click
		End If
			searchUserProfile=bSearchUserProfile
		End Function



	'Function to Verify  Approve/Reject User Profile
	Public Function VerifyUserProfile(str1BankID, strLANID, strName, strRole, strManager1BankID, strManagerLANID,strGroup, strChannel,strLocation,_ 
	    strC3User, strStaffAccess, strStatus,strCreateDate,strCreateBy,strLastApprovedDate,strLastApprovedBy,strLastUpdatedDate,strLastUpdatedBy,strAction)

			Dim bVerifyProfile:bVerifyProfile=true

				

				If Not IsNull(str1BankID) Then
                     If Not VerifyField( txt1BankID(), str1BankID, "1BankID") Then
							bVerifyProfile = False
					 End If	
				Else
					LogMessage "WARN","Verification","1bank Id is mandatory for this keyword hence exiting keyword execution ",true	
					VerifyUserProfile=true
					Exit Function  
                End If
		'str1BankID = lbl1BankID()
				If Not IsNull(strLANID) Then
                    If Not VerifyField( txtLANID(), strLANID, "LAN ID") Then
								bVerifyProfile = False
					 End If
                End If
		'strLANID = lblLANID()
				If Not IsNull(strName) Then
                    If Not VerifyField( txtName(), strName, "Name") Then
								bVerifyProfile = False
					 End If
                End If
	'Verify the Role DDL and select 
		If Not IsNull(strRole) Then
			If not verifyVadinComboProperties(ddlRole,true,strRole,"Role")	 Then
					bVerifyProfile = False
			 End If
		End If
	'Verify Manager 1Bank ID
		If Not IsNull(strManager1BankID) Then
			If not VerifyField(txtManager1BankID,strManager1BankID,"Manager 1Bank ID" )Then
				bVerifyProfile=False
			End If			
		End If
		'strManager1BankID = txtManager1BankID()\				\
		If Not IsNull(strManagerLANID) Then
			If Not VerifyField( txtManagerLANID(), strManagerLANID, "Manager LAN ID") Then
				bVerifyProfile = False
			 End If
        End If

		If Not IsNull(strGroup) Then
			If not verifyVadinComboProperties(ddlGroup,true,strGroup,"Group")	 Then
            	bVerifyProfile = False
			 End If
		End If

		
		If Not IsNull(strChannel) Then
			If not verifyVadinComboProperties(ddlChannel,true,strChannel,"Channel")	 Then
                	bVerifyProfile = False
			 End If
		End If
		If Not IsNull(strLocation) Then
			If not verifyVadinComboProperties(ddlLocation,true,strLocation,"Location")	 Then
                bVerifyProfile = False
			End If
		End If
		'Verify Default Selection
		If not VerifyRadioButtonGrpSelection(strC3User, rbtC3User(), array("Yes", "No")) Then
			LogMessage "WARN","Verification","Radio button "&strC3User&" not selected by default as expected for C3 User",false
			bVerifyProfile = False
		else
			LogMessage "RSLT","Verification","Radio button "&strC3User&" selected by default as expected for C3 User",true
		End If
		'verify the radio button & do a selection
		If not VerifyRadioButtonGrpSelection(strStaffAccess, rbtStaffAccess(), array("Yes", "No")) Then
			LogMessage "WARN","Verification","Radio button "&strStaffAccess&" not selected by default as expected for Staff Access",false
			bVerifyProfile = False
		else
			LogMessage "RSLT","Verification","Radio button "&strStaffAccess&" selected by default as expected for Staff Access",true
		End If
		'verify the radio button & do a selection
		If not VerifyRadioButtonGrpSelection(strStatus, rbtStatus(), array("Active", "Inactive")) Then
			LogMessage "WARN","Verification","Radio button "&strStatus&" not selected by default as expected for Status",false
			bVerifyProfile = False
		else
			LogMessage "RSLT","Verification","Radio button "&strStatus&" selected by default as expected for Status",true
		End If
        'Verify Dates
     		If Not isNull(strCreateDate) Then
			If strCreateDate<>"" Then
				If Ucase(strCreateDate)="TODAY" Then
					strCreateDate=""&Day(CDate(Now)) & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
				End If
				strCreateDatePattern=strCreateDate&" ([0-2][0-9]:[0-9][0-9])"
			End If
			If Not verifyInnerText_Pattern( lblCreatedDate(), strCreateDatePattern, "Created Date") Then
						bVerifyProfile = False
			 End If
		End If
		
		If Not isNull(strCreateBy) Then
			If Not verifyInnerText( lblCreatedBy, strCreateBy, "Created By") Then
						bVerifyProfile = False
			 End If
		End If
			
		If Not isNull(strLastApprovedDate) Then
			If strLastApprovedDate<>"" Then
				If Ucase(strLastApprovedDate)="TODAY" Then
					strLastApprovedDate=""& Day(CDate(Now)) & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
				End If
				strLastApprovedDatePattern=strLastApprovedDate &" ([0-2][0-9]:[0-9][0-9])"
			End If
			If Not verifyInnerText_Pattern( lblLastApprovedDate(), strLastApprovedDatePattern, "Last Approved Date") Then
						bVerifyProfile = False
			 End If
		End If
			
		If Not isNull(strLastApprovedBy) Then
			If Not verifyInnerText( lblLastApprovedBy(), strLastApprovedBy, "Last Approved By") Then
						bVerifyProfile = False
			 End If
		End If
			
		If Not isNull(strLastUpdatedDate) Then
			If strLastUpdatedDate<>"" Then
				If Ucase(strLastUpdatedDate)="TODAY" Then
					strLastUpdatedDate=""& Day(CDate(Now)) & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
				End If
				strLastUpdatedDatePattern=strLastUpdatedDate &" ([0-2][0-9]:[0-9][0-9])"
			End If
			If Not verifyInnerText_Pattern( lblLastUpdatedDate(), strLastUpdatedDatePattern, "Last Updated Date") Then
						bVerifyProfile = False
			 End If
		End If

		If Not isNull(strLastUpdatedBy) Then
			If Not verifyInnerText( lblLastUpdatedBy(), strLastUpdatedBy, "Last Updated By") Then
						bVerifyProfile = False
			 End If
		End If

		If Ucase(strAction)="REJECT" Then
			btnReject().click()
			WaitForICallLoading			
		End If

		If Ucase(strAction)="APPROVE" Then
			btnApprove().click()
		End If

		If Ucase(strAction)="CANCEL" Then
			btnCancel().click()
		End If

		WaitForIcallLoading
		lstSearchResultCol = array("Name","1Bank ID", "Role","Status", "Action")
		lstSearchResult=Array(strName,str1BankID,strRole,strStatus,"Approve")
		 intRow = getRowForColumns (tblSearchResultHeader,tblSearchResultContent,lstSearchResultCol, lstSearchResult)
		if intRow<>-1 then
				LogMessage "WARN","Verification","Expected record "&ArrayToString(lstSearchResult,",")&" does not moved from Pending List table",false
				bSearchUserProfile = false
		else
			LogMessage "RSLT","Verification","Expected record "&ArrayToString(lstSearchResult,",")&" successfully moved from Pending List table",true
		End If
'		If Not isNull(strErrorMessage) Then
'			If Not VerifyInnerText( lblConfigureErrorMsg(), strErrorMessage, "Error Message") Then
'						bVerifyProfile = False
'			 End If
'		End If
			
		VerifyUserProfile = bVerifyProfile
	End Function

	'Create a new profile
	Public Function CreateProfile(str1BankID, strLANID, strName, lstRoles, strRole, strManager1BankID, strInvalidManagerID,strManagerLANID,_ 
	   lstGroup, strGroup,lstChannel, strChannel,lstLocation, strLocation, strC3User, strStaffAccess, strMakerID,strStatus,strCreateDate,strCreateBy,_ 
	   strLastApprovedDate,strLastApprovedBy,strLastUpdatedDate,strLastUpdatedBy,strAction,strErrorMessage)
				Dim bCreateProfile : bCreateProfile=true

				If Not IsNull(str1BankID) Then
                     If Not VerifyField( txt1BankID(), str1BankID, "1BankID") Then
							bCreateProfile = False
					 End If
				 Else
					LogMessage "WARN","Verification","1bank Id is mandatory for this keyword hence exiting keyword execution ",true	
					CreateProfile=true
					Exit Function  
                End If
		'str1BankID = lbl1BankID()
				If Not IsNull(strLANID) Then
                    If Not VerifyField( txtLANID(), strLANID, "LAN ID") Then
								bCreateProfile = False
					 End If
                End If
		'strLANID = lblLANID()
				If Not IsNull(strName) Then
                    If Not VerifyField( txtName(), strName, "Name") Then
								bCreateProfile = False
					 End If
                End If
	'Verify the Role DDL and select 
		If not IsNull(lstRoles) Then
			If Not verifyDropDownListItems(ddlRole,lstRoles) Then
					bCreateProfile = False
			End If
		End If
		If Not IsNull(strRole) Then
			If not selectItem_Combobox(ddlRole,strRole)	 Then
				LogMessage "WARN","Verification","Failed to Select Role : "&strRole, false
					bCreateProfile = False
			 End If
		End If
	'Verify Manager 1Bank ID
		If Not IsNull(strManager1BankID) Then
			txtManager1BankID().set strManager1BankID
			wait 1
			btnimgverify().click
			waitForIcallLoading
			If Not IsNull(strInvalidManagerID) Then
				    If Not VerifyInnerText( lblConfigureErrorMsg(), strInvalidManagerID, "Error Message") Then
'								btnCancel().click()
'								WaitForIcallLoading
							CreateProfile = False
								Exit Function
					 End If
					CreateProfile=true
					Exit Function
			End If
			
		End If
		'strManager1BankID = txtManager1BankID()\				\
		If Not IsNull(strManagerLANID) Then
			If Not VerifyField( txtManagerLANID(), strManagerLANID, "Manager LAN ID") Then
				bCreateProfile = False
			 End If
        End If

		If not IsNull(lstGroup) Then
			If Not verifyDropDownListItems(ddlGroup,lstGroup) Then
					bCreateProfile = False
			End If
		End If
		If Not IsNull(strGroup) Then
			txtName.Click
			If not selectItem_Combobox(ddlGroup,strGroup)	 Then
				LogMessage "WARN","Verification","Failed to Select Group : "& strGroup, false
					bCreateProfile = False
			 End If
		End If

		If not IsNull(lstChannel) Then
			If Not verifyDropDownListItems(ddlChannel,lstChannel) Then
					bCreateProfile = False
			End If
		End If
		If Not IsNull(strChannel) Then
			txtName.Click
			If not selectItem_Combobox(ddlChannel,strChannel)	 Then
				LogMessage "WARN","Verification","Failed to Select Channel : "& strChannel, false
					bCreateProfile = False
			 End If
		End If

		If not IsNull(lstLocation) Then
			If Not verifyDropDownListItems(ddlLocation,lstLocation) Then
					bCreateProfile = False
			End If
		End If

		If Not IsNull(strLocation) Then
			txtName.Click
			If not selectItem_Combobox(ddlLocation,strLocation)	 Then
				LogMessage "WARN","Verification","Failed to Select Location : "& strLocation, false
					bCreateProfile = False
			 End If
		End If
		'Verify Default Selection
		If not VerifyRadioButtonGrpSelection("Yes", rbtC3User(), array("Yes", "No")) Then
			LogMessage "WARN","Verification","Radio button YES not selected by default as expected for C3 User",true
		'	bCreateProfile = False
		else
			LogMessage "RSLT","Verification","Radio button YES selected by default as expected for C3 User",true
		End If
		If Not IsNull(strC3User) Then
			If Not SelectRadioButtonGrp(strC3User, rbtC3User(), array("Yes", "No")) then
				bCreateProfile = False
			End If					
		End If
		
		'verify the radio button & do a selection
		If not VerifyRadioButtonGrpSelection("No", rbtStaffAccess(), array("Yes", "No")) Then
			LogMessage "WARN","Verification","Radio button NO not selected by default as expected for Staff Access",true
			'bCreateProfile = False
		else
			LogMessage "RSLT","Verification","Radio button NO selected by default as expected for Staff Access",true
		End If
		If Not IsNull(strStaffAccess) Then
			If Not SelectRadioButtonGrp(strStaffAccess, rbtStaffAccess(), array("Yes", "No")) then
				bCreateProfile = False
			End If					
		End If		

		'verify the radio button & do a selection
		If not VerifyRadioButtonGrpSelection("Active", rbtStatus(), array("Active", "Inactive")) Then
			LogMessage "WARN","Verification","Radio button ACTIVE not selected by default as expected for Status",true
			'bCreateProfile = False
		else
			LogMessage "RSLT","Verification","Radio button ACTIVE selected by default as expected for Status",true
		End If
		If Not IsNull(strStatus) Then
			If Not SelectRadioButtonGrp(strStatus, rbtStatus(), array("Active", "Inactive")) then
				bCreateProfile = False
			End If					
		End If	

        'Verify Dates
     		If Not isNull(strCreateDate) Then
			If strCreateDate<>"" Then
				If Ucase(strCreateDate)="TODAY" Then
					strCreateDate=""& Day(CDate(Now)) & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
				End If
				strCreateDatePattern=strCreateDate&" ([0-2][0-9]:[0-9][0-9])"
			End If
			If Not verifyInnerText_Pattern( lblCreatedDate(), strCreateDatePattern, "Created Date") Then
						bCreateProfile = False
			 End If
		End If
		
		If Not isNull(strCreateBy) Then
			If Not verifyInnerText( lblCreatedBy, strCreateBy, "Created By") Then
						bCreateProfile = False
			 End If
		End If
			
		If Not isNull(strLastApprovedDate) Then
			If strLastApprovedDate<>"" Then
				If Ucase(strLastApprovedDate)="TODAY" Then
					strLastApprovedDate=""& Day(CDate(Now)) & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
				End If
				strLastApprovedDatePattern=strLastApprovedDate &" ([0-2][0-9]:[0-9][0-9])"
			End If
			If Not verifyInnerText_Pattern( lblLastApprovedDate(), strLastApprovedDatePattern, "Last Approved Date") Then
						bCreateProfile = False
			 End If
		End If
			
		If Not isNull(strLastApprovedBy) Then
			If Not verifyInnerText( lblLastApprovedBy(), strLastApprovedBy, "Last Approved By") Then
						bCreateProfile = False
			 End If
		End If
			
		If Not isNull(strLastUpdatedDate) Then
			If strLastUpdatedDate<>"" Then
				If Ucase(strLastUpdatedDate)="TODAY" Then
					strLastUpdatedDate=""& Day(CDate(Now)) & " "&monthName(Month(CDate(Now)),true) &" " &Year(CDate(Now))
				End If
				strLastUpdatedDatePattern=strLastUpdatedDate &" ([0-2][0-9]:[0-9][0-9])"
			End If
			If Not verifyInnerText_Pattern( lblLastUpdatedDate(), strLastUpdatedDatePattern, "Last Updated Date") Then
						bCreateProfile = False
			 End If
		End If

		If Not isNull(strLastUpdatedBy) Then
			If Not verifyInnerText( lblLastUpdatedBy(), strLastUpdatedBy, "Last Updated By") Then
						bCreateProfile = False
			 End If
		End If

		If Ucase(strAction)="SUBMIT" Then
			btnSubmit().click()
			WaitForIcallLoading
			If  isNull(strErrorMessage)Then
				lstSearchResultCol = array("Name","1Bank ID", "Role","Maker 1Bank ID","Status", "Action")
				lstSearchResult=Array(strName,str1BankID,strRole,strMakerID,"Pending Approval","Approve")
				 intRow = getRowForColumns (tblSearchResultHeader,tblSearchResultContent,lstSearchResultCol, lstSearchResult)
				if intRow=-1 then
						LogMessage "WARN","Verification","Submited record "&ArrayToString(lstSearchResult,",")&" does not displayed in Pending List table",false
						bCreateProfile = false
				else
					LogMessage "RSLT","Verification","Submitted record "&ArrayToString(lstSearchResult,",")&" successfully displayed in Pending List table",true
				End If
			End If
		End If
		If Ucase(strAction)="CANCEL" Then
			btnCancel().click()
			WaitForIcallLoading
		End If

		
		If Not isNull(strErrorMessage) Then
			If  lblConfigureErrorMsg.Exist(0) Then
				If Not VerifyInnerText( lblConfigureErrorMsg(), strErrorMessage, "Error Message") Then
				bCreateProfile = False
				 End If
			 else
				If Not VerifyInnerText( lblApproverErrorMsg(), strErrorMessage, "Warning Message") Then
					bCreateProfile = False
				else
					bCreateProfile=true
			  End If
			 	btnOK.Click
			 End If
		End If
	
		CreateProfile =bCreateProfile
	End Function
	   Public Function verifyReport_InvalidUsers(lstExpectedData)
		  Dim bCreateReport_InvalidUsers:bCreateReport_InvalidUsers=true
			lnkReports.Click		
		'clickReports()
			WaitForICallLoading
			If Not  ReportsPageExists() Then		
				LogMessage "WARN","Verification","Report page does not displayed",false
				verifyReport_InvalidUsers=false
				Exit Function
			Else
			LogMessage "RSLT","Verification","Report page displayed successfully",true
	      End If
				btnDownload_InvalidUsers.Click
				waitForIcallLoading
			 If Not dlgFileDownload.Exist(5) Then
				  LogMessage "WARN","Verification","File download dialog does not displayed",false
				verifyReport_InvalidUsers=false
				Exit Function
		  End if

		    btnFileDownload_Save.Click
		    WaitForICallLoading
	    
		   SysDate = FormatDateTime(Date, 1)
		   
		   Dim strFileName
		   Dim strFolder:strFolder = "D:\Temp\"
		   
		   strFileName = strFolder & txt_FileName.GetROProperty("text")		   
		   'Format FileName
		   sTime = Split(Time, ":", -1, 1)(0) & "." & Split(Time, ":", -1, 1)(1) & "." & Split(Time, ":", -1, 1)(2)
		   sDate = Split(SysDate, ",", -1, 1)(1) & Split(SysDate, ",", -1, 1)(2)
		   sFile=Split(strFileName, ".", -1, 1)(0)&"_"&sDate&"-"&sTime&".csv"   
		   strFileName = sFile    


		   txt_FileName.set strFileName		
		   WaitForICallLoading
		   
		   Set objFSO = CreateObject("Scripting.FileSystemObject")
	
			' Check that the strDirectory folder exists
			If Not objFSO.FolderExists(strFolder) Then
				 Set oFolder = oFSO.CreateFolder(strFolder)
			End If	            
	            
		      btn_SaveDlg.Click
		      WaitForICallLoading
			If Not verifyCSV(strFileName,lstExpectedData) Then
				verifyReport_InvalidUsers=false
				Exit function
			End If
				verifyReport_InvalidUsers=true
	   End Function
'******************************************************************************************************************************************************************************
		'To verify table content from the list
		Public Function verifyCSV(strFileName,lstExpectedData)
			Const ForReading = 1
			Set objFSO = CreateObject("Scripting.FileSystemObject")

			Dim bVerifyCSV:bVerifyCSV=true
			Set objTextFile = objFSO.OpenTextFile (strFileName, ForReading)
			
			iRow = 0
			
			Dim RecordArray
			
			Do Until objTextFile.AtEndOfStream
			
				strCurrrentLine = objTextFile.ReadLine
				ReDim Preserve lstLineArray(iRow)
				lstLineArray (iRow)= Ucase(strCurrrentLine)

					iRow=iRow+1	
			Loop

			For iCount=0 to Ubound(lstExpectedData)
				strTemp=lstExpectedData(iCount)
				If ArrayFind(lstLineArray,Ucase(strTemp)) Then
					LogMessage "RSLT","Verification","Expected record " &strTemp& " found in Report file", true
				Else
					LogMessage "WARN","Verification","Expected record not found  "&strTemp&" in Report file", False
					bVerifyCSV=false
				End If
			Next
			verifyCSV=bVerifyCSV
		End Function

		Public function verifyTableContentList(tblHeader,tblContent,lstlstAccountData,strTableName,bPagination)

		   Dim bVerifyData,arrColumns,arrValues,intSize
			intTablePage=0
			For iRowCount=0 to Ubound(lstlstAccountData,1)
				intSize=Ubound(lstlstAccountData,2)
				'arrTemp=arrPlanData(iRowCount)
				ReDim arrColumns(intSize)
				ReDim arrValues(intSize)
				For iCount=0 to intSize
						arrTemp=Split(lstlstAccountData(iRowCount,iCount),":")
						arrColumns(iCount)=arrTemp(0)
						arrValues(iCount)=checkNull(arrTemp(1))
					
				Next
				If bPagination Then
					Do 
							tblContent.RefreshObject
							intRow=	getRowForColumns (tblHeader,tblContent,arrColumns, arrValues)
						If not intRow=-1 Then
							Exit Do
						End If
						bNextEnabled =matchStr(lnkNext1.GetROProperty("outerhtml"),"v-disabled")
	
								If Not bNextEnabled Then
									lnkNext.Click
									intTablePage=intTablePage+1
									WaitForICallLoading
								End If
						Loop while Not  bNextEnabled
				else
						intRow=	getRowForColumns (tblHeader,tblContent,arrColumns, arrValues)
				End If

			
				If intRow =-1  Then
						LogMessage "WARN","Verification","Expected "& strTableName &" Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" not found in  "& strTableName &" table",false
						bVerifyData= False
					else
							LogMessage "RSLT","Verification","Expected  "& strTableName &" Data "&ArrayToString(arrValues,",")&" for respective column Names "&ArrayToString(arrColumns,",")&" found in "& strTableName &" table at Row "&intRow&" on table page number"&intTablePage,true
						bVerifyData= True
				End If
				If bPagination Then
					For i=0 to intTablePage
						lnkPrevious.Click
						WaitForIcallLoading
					Next
				End If
			Next	
			
					verifyTableContentList=bVerifyData
		End Function
	'
	
		' To open user detail pop up window
        Public Function viewUserDetails(lstUserProfile)
            	Dim bviewUserDetails:bviewUserDetails = True 			

                intSize=Ubound(lstUserProfile)
				ReDim arrColumns(intSize)
				ReDim arrValues(intSize)
	
				For iCount=0 to intSize
						arrTemp=Split(lstUserProfile(iCount),":")
						arrColumns(iCount)=arrTemp(0)
						arrValues(iCount)=checkNull(arrTemp(1))
			
				Next
					   'Verify Acc Details exists
					Dim intRow,bNextEnabled,intTablePage
					bNextEnabled=false
					intTablePage=1
					intRow=	getRowForColumns (tblSearchResultHeader,tblSearchResultContent,arrColumns, arrValues)
					If intRow =-1  Then
						LogMessage "WARN","Verfication","Expected Account Details "&ArrayToString(arrValue,",")&" for respective column Names "&ArrayToString(arrColumnName,",")&" not found in Account Details table",false
						bviewMemoDetails=false
						Exit Function
					else
							LogMessage "RSLT","Verfication","Expected Account Details "&ArrayToString(arrValue,",")&" for respective column Names "&ArrayToString(arrColumnName,",")&" found in Account Details table at Row "&intRow&" on table page number"&intTablePage,true
							bviewMemoDetails=true
					End If

					clickVaddinLink_tblCell tblSearchResultHeader,tblSearchResultContent,intRow, "Action"
					WaitForICallLoading

                viewMemoDetails=bviewMemoDetails
		End Function
        '*****************************End of Function on the Screen **********************************************************

' Select a radiobutton - local function
 Public Function SelectRadioButtonGrp(strItem, objRadioButtons, arrayOptions)
 'Public Function SelectRadioButtonGrp(strItem, arrayOptions)
	Dim oDesc, oChild, allItems
	Set oDesc=Description.Create()
	oDesc("micclass").Value="WebElement"
	Set oChild=objRadioButtons.ChildObjects(oDesc)

	allItems = oChild.Count
	bDisabled=false
	For i = 0 to allItems-1
		print oChild.Item(i).GetRoProperty("innertext")
		print trim(oChild.Item(i).GetRoProperty("micclass"))
		If (oChild.Item(i).GetRoProperty("innertext")=strItem) Then
			If oChild.Item(i).GetRoProperty("class")="v-radiobutton v-select-option" Then
				If (trim(oChild.Item(i+1).GetRoProperty("micclass"))="WebRadioGroup") Then
					'Dim arrayOptions
					'arrayOptions = array("Yes", "No")
					' i+1 is used as we have the radiobutton grp at  +1 level
					
					selectRadioGroup oChild.Item(i+1), strItem, arrayOptions
					

				End If
				If Ucase(gstrBrowser)="CHROME" Then '23-10-2013 ***Changed to handle Chrome - 
						selectRadioGroup oChild.Item(1), strItem, arrayOptions
				End If
				bDisabled=true
			End If
		End If
	
	Next
	SelectRadioButtonGrp=bDisabled
 End Function

  Public Function VerifyRadioButtonGrpSelection(strItem, objRadioButtons, arrayOptions)
 'Public Function SelectRadioButtonGrp(strItem, arrayOptions)
	Dim oDesc, oChild, allItems, iIndex, iSelectedIndex, bVerified
	Set oDesc=Description.Create()
	oDesc("micclass").Value="WebElement"
	Set oChild=objRadioButtons.ChildObjects(oDesc)

	allItems = oChild.Count
	bVerified=false
	For i = 0 to allItems-1
		print oChild.Item(i).GetRoProperty("innertext")
		print trim(oChild.Item(i).GetRoProperty("micclass"))
		If (oChild.Item(i).GetRoProperty("innertext")=strItem) Then
			If matchstr(oChild.Item(i).GetRoProperty("class"),"v-radiobutton v-select-option.*") Then
				If Ucase(gstrBrowser)="CHROME" Then '23-10-2013 ***Changed to handle Chrome - 
				
					iSelectedIndex =  oChild.Item(1).GetRoProperty("selected item index")
					iIndex= IndexOf(arrayOptions, strItem)
					iIndex = iIndex+1
				else 
					'If (trim(oChild.Item(i+1).GetRoProperty("micclass"))="WebRadioGroup") Then

					iSelectedIndex =  oChild.Item(i+1).GetRoProperty("selected item index")
					iIndex= IndexOf(arrayOptions, strItem)
					iIndex = iIndex+1
					
				End If
				If cstr(iIndex) =  iSelectedIndex Then
					bVerified = true
				End If			
			End If
		End If
	
	Next
	VerifyRadioButtonGrpSelection=bVerified
 End Function

  Public Function VerifyRadioButtonGrpProperty(strItem, objRadioButtons, arrayOptions)
	Set oDesc=Description.Create()
	oDesc("micclass").Value="WebElement"
	Set oChild=objRadioButtons.ChildObjects(oDesc)
	allItems = oChild.Count
	bDisabled=false
	For i = 0 to allItems-1
		print oChild.Item(i).GetRoProperty("innertext")
		print trim(oChild.Item(i).GetRoProperty("micclass"))
		If (oChild.Item(i).GetRoProperty("innertext")=strItem) Then
			If oChild.Item(i).GetRoProperty("class")="v-radiobutton v-select-option v-radiobutton-disabled v-disabled" Then
				bDisabled=true
				Exit for
			End If
		End If
	
	Next
	VerifyRadioButtonGrpProperty=bDisabled
 End Function


Public function selectItem_Combobox(objComboBox,strItem)
   	  Set oDesc=Description.Create
	  oDesc("micclass").Value = "WebElement"
	  oDesc("class").Value = "v-filterselect-button"
		set lstObj=objComboBox.ChildObjects(oDesc)
	
		lstObj(0).Click
		wait 1
	Set oDescCombo=Description.Create
	  oDescCombo("micclass").Value = "WebElement"
	  oDescCombo("class").Value = "gwt-MenuItem.*"
	  'oDescCombo("id").Value="VAADIN_COMBOBOX_OPTIONLIST"
	set lstCombo=Browser("Browser_iCall_Home").Page("iCall_UserProfile").ChildObjects(oDescCombo)
	
		intItems=lstCombo.Count
		For iCount=0 to intItems-1
			Dim strTemp:strTemp=""
			strTemp=lstCombo(iCount).GetRoProperty("text")
			If strTemp=strItem Then
					lstCombo(iCount).click
					LogMessage "RSLT","Verification","Item "&strItem&" selected from combobox sucessfully. Item Index is "& intItemIndex,true
					selectItem_Combobox=true
					Exit Function
			End If
			intItemIndex=intItemIndex+1
		Next
		LogMessage "WARN","Verification","Item "&strItem&" Not found in combobox",false
	Dim intItemIndex
	intItemIndex=0
	
	selectItem_Combobox=false
End Function

Public function verifyDropDownListItems(objComboBox,lstItems)
   Dim bVerifyDropDownListItems:bVerifyDropDownListItems=true
	arrListItems= getItemsList_ComboBox(objComboBox)
	iItemCount=Ubound(lstItems)
	If not Ubound(arrListItems)=Ubound(lstItems) Then
		If (lstItems(iItemCount)<>"" ) Then
			LogMessage "RSLT","Verification","Number of Items displayed in drop down list does not matched with expected list. Expected :"& ArrayToString(lstItems,","), false
			verifyDropDownListItems=false
			Exit function
		End If
	End If
   For iCount=0 to UBound(lstItems)
		strDBItem=lstItems(iCount)
		If strDBItem<>"" Then
			If not ArrayFind(arrListItems,strDBItem) Then
					bVerifyDropDownListItems=false
					LogMessage "RSLT","Verification","Item "&strDBItem&" does not displayed on UI",true
			End If
		End If
   Next
	verifyDropDownListItems=bVerifyDropDownListItems
End Function
Public Function getItemsList_ComboBox(objComboBox)
     Set oDesc=Description.Create
	  oDesc("micclass").Value = "WebElement"
	  oDesc("class").Value = "v-filterselect-button"
		set lstObj=objComboBox.ChildObjects(oDesc)
	
			lstObj(0).Click
			wait 2
	Set oDescCombo=Description.Create
	  oDescCombo("micclass").Value = "WebElement"
	  oDescCombo("class").Value = "gwt-MenuItem.*"
	  'oDescCombo("id").Value="VAADIN_COMBOBOX_OPTIONLIST"
	set lstCombo=Browser("Browser_iCall_Home").Page("iCall_UserProfile").ChildObjects(oDescCombo)
	
		intItems=lstCombo.Count
		ReDim arrComboItems(Cint(intItems)-1)
	'Get Count of Combo Items
		For iCount=0 to intItems-1
			Dim strTemp:strTemp=""
			strTemp=lstCombo(iCount).GetRoProperty("text")
			arrComboItems(intItemIndex)=strTemp
			intItemIndex=intItemIndex+1
				wait 1
		Next
	
	lstObj(0).Click
	wait 2
	getItemsList_ComboBox=arrComboItems
End Function
Public Function getVadinCombo_SelectedItem(objVadinCombo)
	getVadinCombo_SelectedItem=objVadinCombo.WebEdit("micclass:=WebEdit").GetRoProperty("value")
   '//*[@id="customer.serviceRequestDetails.assignedToField"]/input
End Function
'This will validate if  combox is read only or not along with selected text
Private Function verifyVadinComboProperties(objComboBox,bDisabled,strExpectedVal,strFieldName)
	Dim strTemp, strActual,bVerifyVadinComboProperties
	bVerifyVadinComboProperties=true
	strActual=getVadinCombo_SelectedItem(objComboBox)
	strTemp=objComboBox.GetRoProperty("class")
	If bDisabled Then
		If InStr (1,strTemp,"v-readonly")<>0  then
			LogMessage "RSLT","Verification","Field "&strFieldName&" is disabled as expceted",true
		else
				LogMessage "WARN","Verification","Field "&strFieldName&" is not disabled as expceted",false
				bVerifyVadinComboProperties=false
		End If
	Else
		If InStr (1,strTemp,"v-readonly")=0  then
			LogMessage "RSLT","Verification","Field "&strFieldName&" is Enabled as expceted",true
		else
				LogMessage "WARN","Verification","Field "&strFieldName&" is  Not Enabled as expceted",false
				bVerifyVadinComboProperties=false
		End If
	End If
	If strActual=strExpectedVal Then
		LogMessage "RSLT","Verification","Expected value "&strExpectedVal&" matched with actual value for combobox"&strFieldName, true
	Else
		LogMessage "WARN","Verification","Expected value "&strExpectedVal&" does not matched with actual value "&strActual&" for combobox"&strFieldName, false
			bVerifyVadinComboProperties=false
	End If
	verifyVadinComboProperties=	bVerifyVadinComboProperties
End Function
Public Function verifySummaryCount(tblHeader,tblContent,strColumnName,strExpectedTranCount,strTableName,bPagination)
		   	'Get Row count

			Dim intRow,intRowtemp
			intRow=0
			bNextPageExist=true
			If bPagination Then
				Do 
					intRowTemp=0
					tblContent.RefreshObject
					Set newTable=tblContent
					intRowTemp=getRecordsCountForColumn (tblHeader,newTable,strColumnName)
					

					intRow=intRow+intRowTemp
					bNextPageExist =matchStr(lnkNext1().GetROProperty("outerhtml"),"v-disabled")
			
							If Not bNextPageExist Then
								lnkNext().Click 
								intTablePage=intTablePage+1
								WaitForICallLoading
							End If
				Loop while Not  bNextPageExist
				
			Else
					intRowTemp=0
					intRowTemp=getRecordsCountForColumn (tblHeader,tblContent,strColumnName)
					intRow=intRow+intRowTemp
			End If
			If CInt(strExpectedTranCount)=CInt(intRow) Then
				LogMessage "RSLT","Verification","Number of recodrs displayed for "&strTableName&" matched with expected, Expecetd Count is "&strExpectedTranCount&" and Actual Count is "&intRow,true
				verifySummaryCount=true
			Else
				LogMessage "WARN","Verification","Number of recodrs displayed for  "&strTableName&" does not matched with expected, Expecetd Count is "&strExpectedTranCount&" and Actual Count is "&intRow,false
				verifySummaryCount=false
			End If
			If bPagination Then
				 For i=0 to intTablePage
						lnkPrevious().Click
						WaitForIcallLoading
				 Next
			End If
			
		End Function

Public Function VerifyAuditTrailReport(strFromDate, strToDate, strMessage, strMakers1BankID, strMakersRecordCreateDate, strCheckers1BankID, strCheckersRecordDate, strUsers1BankID, lstAttributeName, lstAttributeOldValue, lstAttributeNewValue)
		Dim bVerifyAuditTrailReport:bVerifyAuditTrailReport = True
		
		If Not  pageExists() Then		
			LogMessage "WARN","Verification","Admin Home Page does not displayed",false
			VerifyAuditTrailReport=false
			Exit Function
		 Else
				LogMessage "RSLT","Verification","Admin Home Page displayed successfully",true
		End If
		lnkReports.Click		
		'clickReports()
		WaitForICallLoading
	
	    If Not  ReportsPageExists() Then		
			LogMessage "WARN","Verification","Report page does not displayed",false
			VerifyAuditTrailReport=false
			Exit Function
	    Else
			LogMessage "RSLT","Verification","Report page displayed successfully",true
	      End If
	
	    If Not IsNull(strFromDate) Then
			txtAuditReportFromDate.Set strFromDate
	    End If	
	    
	    If Not IsNull(strToDate) Then
			txtAuditReportToDate.Set strToDate
	    End If	
	    
	    CheckVaadinObject_Disabled btnDownload_Audit, "DISABLED"    
	    clickGoAuditReport()
	    WaitForICallLoading
	        
	    If Not IsNull(strMessage) Then  ' Error Scenario
	        CheckVaadinObject_Disabled btnDownload_Audit, "DISABLED"
	      	If Not verifyInnerText(lblReportErrorMsg , strMessage, "Error Message - ")Then
		 		VerifyAuditTrailReport = False
		 		Exit Function		
			else
				VerifyAuditTrailReport = True
		 		Exit Function		
		     End If          
			  
	    Else							     'Positive Scenario
	         CheckVaadinObject_Disabled btnDownload_Audit, "ENABLED"
		    clickDownloadAuditReport()
		    WaitForICallLoading
		  If Not dlgFileDownload.Exist(5) Then
			  LogMessage "WARN","Verification","File download dialog does not displayed",false
		  End if

		    btnFileDownload_Save.Click
		    WaitForICallLoading
	    
		   SysDate = FormatDateTime(Date, 1)
		   
		   Dim strFileName
		   Dim strFolder:strFolder = "D:\Temp\"
		   
		   strFileName = strFolder & txt_FileName.GetROProperty("text")		   
		   'Format FileName
		   sTime = Split(Time, ":", -1, 1)(0) & "." & Split(Time, ":", -1, 1)(1) & "." & Split(Time, ":", -1, 1)(2)
		   sDate = Split(SysDate, ",", -1, 1)(1) & Split(SysDate, ",", -1, 1)(2)
		   sFile=Split(strFileName, ".", -1, 1)(0)&"_"&sDate&"-"&sTime&".csv"   
		   strFileName = sFile    


		   txt_FileName.set strFileName		
		   WaitForICallLoading
		   
		   Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	            ' Check that the strDirectory folder exists
	            If Not objFSO.FolderExists(strFolder) Then
	                 Set oFolder = oFSO.CreateFolder(strFolder)
	            End If	            
	            
		      btn_SaveDlg.Click
		      WaitForICallLoading
		      
'	   	   CheckVaadinObject_Disabled btnDownloadAuditReport, "DISABLED"


	   'Logic to check the expected data in the csv file
	If not isNull(strMakers1BankID) Then
			Makers1BankID = strMakers1BankID
			MakersRecordCreateDate = strMakersRecordCreateDate
			Checkers1BankID = strCheckers1BankID
			CheckersRecordDate = strCheckersRecordDate
			Users1BankID = strUsers1BankID
			AttributeName = lstAttributeName
			AttributeOldValue = lstAttributeOldValue
			AttributeNewValue = lstAttributeNewValue
		
			Dim ExpRecordArray
			
			For itemcnt = 0 To Ubound(AttributeName)
				strLineArray = Array(Trim(Makers1BankID),DateValue(MakersRecordCreateDate),Trim(Checkers1BankID),DateValue(CheckersRecordDate),Trim(Users1BankID),Trim(AttributeName(itemcnt)),Trim(AttributeOldValue(itemcnt)),Trim(AttributeNewValue(itemcnt)))
				ReDim ExpTempArray (itemcnt, 7)
				
				If itemcnt = 0 Then
					ExpRecArray = appendTwoDimensionalArray(ExpTempArray,strLineArray)
				Else
					ExpRecArray = appendTwoDimensionalArray(ExpRecArray,strLineArray)		
				End If	
			Next
			
			Const ForReading = 1
	
			Set objTextFile = objFSO.OpenTextFile (strFileName, ForReading)
			
			iRow = 0
			
			Dim RecordArray
			
			Do Until objTextFile.AtEndOfStream
					strCurrrentLine = objTextFile.ReadLine
					strLineArray = Split(strCurrrentLine , ",")
			
					
					If Not isNull(UBound(strLineArray)) And Trim(strLineArray(0)) <> "Makers1BankID" Then  ' NUmber of array column
						'Modify the date column
						strLineArray(1) = DateValue(strLineArray(1))
						strLineArray(3) = DateValue(strLineArray(3))
						
						If iRow = 0 Then
							ReDim arrayTemp(iRow,7)			
							RecordArray = appendTwoDimensionalArray(arrayTemp,strLineArray)
							iRow = iRow + 1					
						Else				
							If Trim(strLineArray(0)) = Trim(Makers1BankID) And _
								DateValue(strLineArray(1)) = DateValue(MakersRecordCreateDate) And _
								Trim(strLineArray(2)) = Trim(Checkers1BankID) And _
								DateValue(strLineArray(3)) = DateValue(CheckersRecordDate) And _
								Trim(strLineArray(4)) = Trim(Users1BankID)		    Then
								
									ReDim arrayTemp(iRow,7)			
									RecordArray = appendTwoDimensionalArray(RecordArray,strLineArray)		
									iRow = iRow + 1							
							End If			
						End If
					End If
			Loop
			
			For i = 0 To UBound(ExpRecArray)
				bRecordExist = False
				TgtArray = FetchRowArrayFrom2DArray(ExpRecArray,i)
	'			strExpectedRow = ArrayToString(TgtArray,"|")
				For j = 0 To UBound(RecordArray)
					SrcArray = FetchRowArrayFrom2DArray(RecordArray,j)			
					
					If ArrayCompare(SrcArray,TgtArray) Then
						
						Print "Expected row "&i+1 &" found in report "& j+1
						LogMessage "RSLT","Verification","Expected record "&i+1 &" found in Profile Admin report in Row "&j+1,True			
						bRecordExist = True
					End If		
				 Next
				If Not bRecordExist Then			
					Print "Expected row "&i+1 &" not found in report"
					LogMessage "WARN","Verification","Expected record "&i+1 &" not found in Profile Admin log ",False
					bVerifyAuditTrailReport = False				
				End If	 
			Next     
	End If
			
			VerifyAuditTrailReport = bVerifyAuditTrailReport
		
	    End If
	End Function


End Class
