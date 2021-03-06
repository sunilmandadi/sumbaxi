'+++++++++++++++++++++++++++++++++++ File Header Information ++++++++++++++++++++++++++++++++++++++++++++++
	'<Summary>  This file contains all the class declarations and 
								'functions to read /get data from Excel
                                '</summary>

	
   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' <summary>The class for Excel Engine. This class provides 
	'all the required methods to fetch the TestData from excel data sheet
	' Similar class can be created for any other database, if used as test data provider</summary>
''*********************************************************************************************************************************************************************************
Option Explicit

Public function cExcelDataEngine
  Set cExcelDataEngine=new clsExcelDataEngine
End Function

Class clsExcelDataEngine
	Private mExcelDataEngine
			
	Private sub Class_Initialize()
		Set mExcelDataEngine=cDataEngine()
	End Sub
	
	Private sub Class_Terminate()
		Set mExcelDataEngine=nothing
	End Sub

	Function dbConnection (strFileName)
		'This function make connection with excel database.  i/p = Excel File Name
		
		Dim objAdCon
		Dim bConnection
		Dim bConnected
		
		Set objAdCon = CreateObject("ADODB.Connection")
		
		'Rohit - Made change in the drivers
		
		objAdCon.Open "DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ="&strFileName& ";Readonly=False"
		
		If Err <> 0 Then
			Print "DB CONNECTION ERROR-"&Err.Description
			Reporter.ReportEvent micFail,"Create Connection", "[Connection] Error has occured. Error : " & Err
			Set dbConnection = Nothing
			Exit Function
		End If
		
		bConnected="Connected"
		
		Set dbConnection = objAdCon
	End Function
	
 	Public Function startDataEngine(strTestcaseName, strKWsheetPath)
		  mExcelDataEngine.startDataEngine strTestcaseName, strKWsheetPath
	End Function
	
	Function FetchData(strTestcase,strFileName)
		strSQLStatement="Select  *  from ["& strTestcase &"$] where  ["& strTestcase &"$].Sequence <> '0' AND ["& strTestcase &"$].TestCase <>'NULL'  "
	  'msgbox "Excel Engine"
	'This function connects with the Testcase sheet and returns all the testcase data . i/p = sql statement , excel file name
	Dim objAdsRs, objCon, lstMyResult,objAdRs
	Set objCon=dbConnection (strFileName)
	'msgbox"Test"  &strSQLStatement
	Set objAdRs = CreateObject("ADODB.Recordset") 
	objAdRs.CursorLocation=3                     ' set the cursor to use adUseClient - disconnected recordset
	'On error Resume Next
		objAdRs.Open strSQLStatement, objCon, 1, 3  
	
		'****Report Error if failed to connect with excel  data sheet
		If Err<>0 Then
			iErrNum =err.number
			strErrDescription=err.description
			err.clear
			LogMessage "RSLT", "Input Validations", "Invalid Input data, error occured is "&strErrDescription, false
			 'Reporter.ReportEvent micFail,"Open Recordset", "Error has occured.Error Code : " & Err
			  Set FetchExcelValue = Nothing
			 Exit Function
		End If 	
	   
	
	
	Dim iColCount,iRowCount, i
	iColCount=objAdRs.fields.count ' Get max number of columns 
	iRowCount =objAdRs.RecordCount  'Get max number of Rows
		'msgbox iRowCount
		ReDim MyRowData (0,1)
		ReDim MyRowData (cInt(iRowCount), cInt(iColCount)) 'change dimension of array
	   
		iRowCount=0
		
	' MsgBox iColName
		While objAdRs.EOF=false
				 For i=0 to iColCount-1
						'Print ("Data " &i & ": " & isnull(objAdRs.fields(i)))
						 If (objAdRs.fields(i) <>"") OR (not isnull(objAdRs.fields(i))) Then 
							MyRowData(iRowCount,i)=( objAdRs.fields(i))
						else
							Exit for
						End If
				Next   
				iRowCount=iRowCount+1
				objAdRs.moveNext   	
		Wend
				
		If Err<>0 Then
			iErrNum =err.number
			strErrDescription=err.description
			err.clear
			LogMessage "RSLT", "Input Validations", "Invalid Input data, error occured is "&strErrDescription, false
			 Reporter.ReportEvent micFail,"Open Recordset", "Error has occured.Error Code : " & Err
			  Set FetchExcelValue = Nothing
			 Exit Function
		End If 	
	   
		  'msgbox "" & Ubound (arrCol,2)
		 Set objAdRs.ActiveConnection = Nothing   		
		 'msgbox MyRowData(1,3)
		 Set objCon = Nothing  		
		  'msgbox iRowCount 
		  'msgbox  iColCount
		
		  FetchData =  MyRowData
	End Function
	
	Function FetchExcelValue(strSQLStatement,strFileName)
	'This function connects with the Excel sheet and returns all the data based on query . i/p = sql statement , excel file name
	Dim objAdsRs, objCon, lstMyResult,objAdRs
	Set objCon=dbConnection (strFileName)
	'msgbox"Test"  &strSQLStatement
	Set objAdRs = CreateObject("ADODB.Recordset") 
	objAdRs.CursorLocation=3                     ' set the cursor to use adUseClient - disconnected recordset
	objAdRs.Open strSQLStatement, objCon, 1, 3  
	
	Dim iColCount,iRowCount, i
	iColCount=objAdRs.fields.count ' Get max number of columns 
	iRowCount =objAdRs.RecordCount  'Get max number of Rows
		'msgbox iRowCount
		ReDim MyRowData (0,1)
		ReDim MyRowData (cInt(iRowCount)-1,  cInt(iColCount)-1) 'change dimension of array
	   
		iRowCount=0
		
	' MsgBox iColName
		While objAdRs.EOF=false
	        'MsgBox(objAdRs.GetRows ())
				 For i=0 to iColCount-1
						'msgbox objAdRs.fields(i)
						 If (objAdRs.fields(i) <>"") OR (not isnull(objAdRs.fields(i))) Then
							MyRowData(iRowCount,i)=( objAdRs.fields(i))
						 else
							If  ( Trim(  ( Ucase (objAdRs.fields(giDDDataSheetRecordType)  ) ) ) = "KWD_NEXT" ) OR ( Trim(  ( Ucase (objAdRs.fields(giDDDataSheetRecordType)  ) ) ) = "KWI_NEXT" )Then
								MyRowData(iRowCount,i)=( objAdRs.fields(i))
							else
								Exit For
							End If					
						End If
				Next   
				iRowCount=iRowCount+1
				objAdRs.moveNext   	
		Wend
				
		If Err<>0 Then
			 Reporter.ReportEvent micFail,"Open Recordset", "Error has occured.Error Code : " & Err
			  Set FetchExcelValue = Nothing
			 Exit Function
		End If 	
	   
		  'msgbox "" & Ubound (arrCol,2)
		 Set objAdRs.ActiveConnection = Nothing   		
		 'msgbox MyRowData(1,3)
		 Set objCon = Nothing  		
		  'msgbox iRowCount 
		  'msgbox  iColCount
		
		  FetchExcelValue =  MyRowData
	End Function
	
	Function fetchRowValuesForAColumn(strTableName, strColumnName, strExcelPath)
	
	Dim strSQLQuery
	strSQLQuery =   "Select  [" & strTableName & "$]." & strColumnName & "  from [" & strTableName & "$]"
	
	Dim arrKW
	arrKW = cExcelDataEngine.FetchExcelValue (strSQLQuery,strExcelPath )
	
	Dim arrRowValues
	arrRowValues = fetchFirstElementsOfAllRows (arrKW)
	
	fetchRowValuesForAColumn = arrRowValues
	
	Exit Function
	
	End Function
	
	Function fetchRowValuesForAColumnBasedOnQuery(strTableName, strColumnName, strExcelPath, strSearchColumn, strSearchValue)
	
	Dim strSQLQuery
	strSQLQuery =   "Select  [" & strTableName & "$]." & strColumnName & "  from [" & strTableName & "$] WHERE ["  &  strTableName & "$]." & strSearchColumn & "= '" & strSearchValue  &"'"
	
	Dim arrKW
	arrKW = cExcelDataEngine.FetchExcelValue (strSQLQuery,strExcelPath )
	
	Dim arrRowValues
	arrRowValues = fetchFirstElementsOfAllRows (arrKW)
	
	fetchRowValuesForAColumnBasedOnQuery = arrRowValues
	
	Exit Function
	
	End Function
	
	Function fetchRowValuesForAColumnBasedTwoSearchValues(strTableName, strColumnName, strExcelPath, strSearchColumn1, strSearchValue1, strSearchColumn2, strSearchValue2)
	
	Dim strSQLQuery
	strSQLQuery =   "Select  [" & strTableName & "$]." & strColumnName & "  from [" & strTableName & "$] WHERE ["  &  strTableName & "$]." & strSearchColumn1 & "= '" & strSearchValue1  &"' And [" &   strTableName & "$]." & strSearchColumn2 & "= '" & strSearchValue2  &"'"
	
	Dim arrKW
	arrKW = cExcelDataEngine.FetchExcelValue (strSQLQuery,strExcelPath )
	
	Dim arrRowValues
	arrRowValues = fetchFirstElementsOfAllRows (arrKW)
	
	fetchRowValuesForAColumnBasedTwoSearchValues = arrRowValues
	
	Exit Function
	
	End Function
	
	Function fetchRowValuesForAColumnBasedFiveSearchValues(strTableName, strColumnName, strExcelPath, strSearchColumn1, strSearchValue1, strSearchColumn2, strSearchValue2,strSearchColumn3, strSearchValue3,strSearchColumn4, strSearchValue4, strSearchColumn5, strSearchValue5)
	
	Dim strSQLQuery
	strSQLQuery =   "Select  [" & strTableName & "$]." & strColumnName & "  from [" & strTableName & "$] WHERE ["  &  strTableName & "$]." & strSearchColumn1 & "= '" & strSearchValue1  & "' And [" &   strTableName & "$]." & strSearchColumn2 & "= '" & strSearchValue2  &  "' And [" &   strTableName & "$]." & strSearchColumn3 & "= '" & strSearchValue3 &  "' And [" &   strTableName & "$]." & strSearchColumn4 & "= '" & strSearchValue4 &  "' And [" &   strTableName & "$]." & strSearchColumn5 & "= '" & strSearchValue5 & "'"
	
	Dim arrKW
	arrKW = cExcelDataEngine.FetchExcelValue (strSQLQuery,strExcelPath )
	
	Dim arrRowValues
	arrRowValues = fetchFirstElementsOfAllRows (arrKW)
	
	fetchRowValuesForAColumnBasedFiveSearchValues = arrRowValues
	
	Exit Function
	
	End Function
	
	Function fetchDetailsOfKeyword( strTestcaseName, strExcelPath, strTestcaseSheet)
	
	Dim strSQLQuery
	'strSQLQuery =   "Select  *  from [" & strTestcaseSheet &"$] where  ["& strTestcaseSheet &"$].Keyword ='" & strKeywordName & "' AND  ["& strTestcaseSheet & "$].Testcase ='" & strTestcaseName& "'AND  ["& strTestcaseSheet & "$].Sequence ='" & strParentKWSeq & "'"
	strSQLQuery =   "Select  *  from [" & strTestcaseSheet &"$] where  [" &  strTestcaseSheet & "$].Testcase ='" & strTestcaseName & "'"
	
	Dim arrKW
	arrKW = cExcelDataEngine.FetchExcelValue (strSQLQuery,strExcelPath )
	
	'fetchDetailsOfKeyword = returnColumnValuesForARow (arrKW, 0)
	fetchDetailsOfKeyword = arrKW
	
	Exit Function
	
	End Function
	
	Function getKWDDataFromDDSheet(strFileName,strSheetName)
	
		Dim strData()
		Dim objFS, objExcel, objSheet, objRange
		Dim intTotalRow, intTotalCol
		Dim intRow, intCol
		
		' create the file system object
		Set objFS = CreateObject("Scripting.FileSystemObject")
		
		' ensure that the xls file exists
		If Not objFS.FileExists(strFileName) Then
		
			' issue a fail if the file wasn't found
			Reporter.ReportEvent micFail, "Read XLS", "Unable to read XLS file, file not found: " & strFileName
			' file wasn't found, so exit the function
			Exit Function
		
		End If ' file exists
		
		' create the excel object 
		Set objExcel = CreateObject("Excel.Application")
		
		' open the file
		objExcel.Workbooks.open strFileName
		
		' select the worksheet
		Set objSheet = objExcel.ActiveWorkbook.Worksheets(strSheetName)
		
		' select the used range
		Set objRange = objSheet.UsedRange
		
		' count the number of rows
		intTotalRow=CInt(Split(objRange.Address, "$")(4)) - 1
		
		' count the number of columns
		intTotalCol= objSheet.Range("A1").CurrentRegion.Columns.Count
		
		' redimension the multi-dimensional array to accomodate each row and column
		ReDim strData(intTotalRow, intTotalCol)
		ReDim array2DDDRows (0,1)
		Dim iCountAppend
		iCountAppend = 0
		' for each row
		For intRow = 0 to intTotalRow - 1
	
			Dim strTempDatasetONOFF, strRecordType, strSequence
	
			strTempDatasetONOFF = Ucase (Trim(objSheet.Cells(intRow + 2, (giDDDataSheetDataSetONOFF+1)).Value))
			strRecordType = Ucase (Trim(objSheet.Cells(intRow + 2, (giDDDataSheetRecordType+1) ).Value))
			strSequence = Trim(objSheet.Cells(intRow + 2, (giDDDataSheetSequence+1)).Value)
	
		   If (  (Ucase (Trim (strTempDatasetONOFF))= "ON")  AND  (Ucase (Trim (strRecordType))= "KWD" )  And (strSequence <>  "0") )Then
	
				ReDim arrString(0)
	
				   ' for each column
				For intCol =0 to intTotalCol - 1
			
					' store the data from the cell in the array
					Dim strTemp
					strTemp = Trim(objSheet.Cells(intRow + 2,intcol + 1).Value)
	
					'strData(intRow, intcol) = Trim(objSheet.Cells(intRow + 2,intcol + 1).Value)
					 addItemToArray arrString, strTemp
			
				Next ' column
	
				Dim arrayTemp
	
				If  iCountAppend =0 Then
					arrayTemp = appendTwoDimensionalArray (array2DDDRows, arrString )
				else
					arrayTemp = appendTwoDimensionalArray (arrayTemp, arrString )
				End If
	
				iCountAppend =+1
	
		   End If
			
		
		Next ' row
		
		' close the excel object
		objExcel.DisplayAlerts = False
		objExcel.Quit 
		
		' destroy the other objects 
		Set objFS = Nothing 
		Set objExcel = Nothing
		Set objSheet = Nothing 
		
		' return the array containing the data
		getKWDDataFromDDSheet = arrayTemp
	
	End Function

End Class
