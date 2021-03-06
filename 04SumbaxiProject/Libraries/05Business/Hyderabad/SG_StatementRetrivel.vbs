'*****This is auto generated code using code generator please Re-validate ****************

'[Click on Statement Link from Customer Overview Page]
Public Function clickStatement()
	StatementRetrieval.lnkStatement.Click
	 If Err.Number<>0 Then
       clickStatement=false
            LogMessage "WARN","Verification","Failed to Click Button : Statement" ,false
       Exit Function
   End If
   waitForIcallLoading
   clickStatement=true
End Function

'[Validate Statement Link disable on Customer Overview Page]
Public Function validateStatementLink()
	bvalidateStatementLink=true
	intStatementlnk=Instr(StatementRetrieval.lnkStatement.GetROproperty("outerhtml"),("v-disabled"))
		If  not intStatementlnk=0 Then
			LogMessage "RSLT","Verification","Statement Link is disabled as per expectation.",True
			bvalidateStatementLink=true
		Else
			LogMessage "WARN","Verifiation","Statement Link is enable. Expected to be disabled.",false
			bvalidateStatementLink=false
		End If
	validateStatementLink=bvalidateStatementLink
End Function

'[Verify row Data in Table Statement Type]
Public Function verifytblStatementType_RowData(lstlstStatementType)
   bDevPending=false   
   verifytblStatementType_RowData=verifyTableContentList(StatementRetrieval.tblStatementTableHeader,StatementRetrieval.tblStatementTableContent,lstlstStatementType,"Statement Type" , false,null ,null,null)
End Function

'[Click Statement Type link in Table Statement Type]
Public Function clickStatementType_link(arrRowDataList)
   bDevPending=false
   clickStatementType_link=selectTableLink(StatementRetrieval.tblStatementTableHeader,StatementRetrieval.tblStatementTableContent,arrRowDataList,"StatementType" ,"Statement Type",False,NULL ,NULL,NULL)
End Function

'[Click Statement Month link in Table Current Year Statement]
Public Function clickCurrentMonthStatement_link(strExpectedMonth)
   bDevPending=false
   Dim statementCount  
	set statementCount=StatementRetrieval.tblStatementCurrentYearContent
	
	tableRow=statementCount.GetROProperty("rows")
	tableColumn=statementCount.GetROProperty("cols")
	For i=1 to tableRow
		For j = 1 To tableColumn			
		 actualMonth=statementCount.GetCellData(i,j)
		 If actualMonth = strExpectedMonth Then		 	
			 statementCount.ChildItem(i,j,"WebElement",i).click
			 LogMessage "RSLT","Verification","Successfully found expected Statement Month", True
			 clickCurrentMonthStatement_link=true 			 
		 	Exit Function
		 End If
		Next		
	Next
	LogMessage "WARN","Verification","Not able to find expected Statement Month", True
	clickCurrentMonthStatement_link=False 	
End Function

'[Select Previous Year Tab to open Previous Year Statment]
Public Function selectPreviousYear(strPreviousYear)
	bDevPending=false
   Dim yearCount  
	set yearCount=StatementRetrieval.tblYearTable	
	tableRow=yearCount.GetROProperty("rows")
	tableColumn=yearCount.GetROProperty("cols")
	For i=1 to tableRow
		For j = 1 To tableColumn			
		 actualYear=yearCount.GetCellData(i,j)
		 If actualYear = strPreviousYear Then		 	
			 yearCount.ChildItem(i,j,"WebElement",i).click
			 LogMessage "RSLT","Verification","Successfully found expected Previous Year", True
			 selectPreviousYear=true 			 
		 	Exit Function
		 End If
		Next		
	Next
	LogMessage "WARN","Verification","Not able to find expected Previous Year", True
	selectPreviousYear=False
End Function

'[Click Month from Previous Year Table]
Public Function selectPreviousYearMonth(strPreviousYearMonth)
	bDevPending=false
   Dim statementCount  
	set statementCount=StatementRetrieval.tblStatementPreviousYearContent	
	tableRow=statementCount.GetROProperty("rows")
	tableColumn=statementCount.GetROProperty("cols")
	For i=1 to tableRow
		For j = 1 To tableColumn			
		 actualMonth=statementCount.GetCellData(i,j)
		 If actualMonth = strPreviousYearMonth Then		 	
			 statementCount.ChildItem(i,j,"WebElement",i).click
			 LogMessage "RSLT","Verification","Successfully found expected Previous Year Month", True
			 selectPreviousYearMonth=true 			 
		 	Exit Function
		 End If
		Next		
	Next
	LogMessage "WARN","Verification","Not able to find expected Previous Year Month", True
	selectPreviousYearMonth=False
End Function

'[Validate if Statement Open Successfully]
Public Function validateStatement()
	bvalidateStatement=true
	checkStatement=StatementRetrieval.lblStatementDetail.GetROProperty("type")
	If checkStatement = "application/pdf" Then
	  logMessage "RSLT", "Verification", "Selected Statement open successfully.", True
	  bvalidateStatement=true
	Else
	  logMessage "WARN", "Verification", "failed to open selected Statement", false
	  bvalidateStatement=false		
	End If
validateStatement=bvalidateStatement	
End Function




