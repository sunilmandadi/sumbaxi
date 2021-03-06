'+++++++++++++++++++++++++++++++++++ File Header Information ++++++++++++++++++++++++++++++++++++++++++++++
	'<Summary>  This file contains all the functions for the  Data Exchange. 
								'The data Exchange functions works as communication channel between two
								'independent Keywords at run time.
								'</summary>


   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

   
   '***********************************************************Keyword Data Exchange Functions**************************************************************************************************


Public Sub  insertAccountMap(strAccountNumber, strAccountKey, strUsed)

	Dim strSQLStatement
	strAccountKey = Trim(UCase(strAccountKey))
	strUsed = Trim(UCase(strUsed))
	strSQLStatement= "Insert Into [Accounts$] (AccountNumber, AccountKey, Used) Values ('"& strAccountNumber & "', '" & strAccountKey & "', '" & strUsed & "') "
	cExcellogger().logToExcel "", gstrAccountMapExcelPath, strSQLStatement

End Sub

Public Function  fetchAccountNumber(strAccountKey)

   strAccountKey = Trim(UCase(strAccountKey))
	Dim arrFetchedValue
	arrFetchedValue = cExcelDataEngine().fetchRowValuesForAColumnBasedTwoSearchValues("Accounts", "AccountNumber", gstrAccountMapExcelPath, "AccountKey",strAccountKey,"Used","NO")
	fetchAccountNumber = arrFetchedValue(0)	

End Function

Public Sub  markAccountUsed(strAccountNumber)

	Dim strSQLStatement
	strSQLStatement= "update [Accounts$] set [Accounts$].Used = 'YES' Where  [Accounts$].AccountNumber='"  & strAccountNumber & "'"
	cExcellogger().logToExcel "", gstrAccountMapExcelPath, strSQLStatement

End Sub


''********************************************************End of Keyword Data Exchange Functions '*********************************************************************************************

Public Sub  insertDataStore(strParameter, strValue)

	Dim strSQLStatement
	Dim strTestNameDataStore, strDataSetDataStore, strKWActionDataStore
	Dim arrayTest

	arrayTest = Split(gstrDataset,"|")
	strTestNameDataStore = arrayTest(0)

	If  (UBound(arrayTest ) > 0 )Then
		strDataSetDataStore = arrayTest(1)
	else
		strDataSetDataStore = "NA"
	End If

	If isNull(gstrAction) Then
		strKWActionDataStore = "Blank"
	else
		strKWActionDataStore = gstrAction
	End If

		

	strSQLStatement = "Insert Into [Common$] (TestcaseName, DataSet,KeywordName,KeywordAction,Parameter,ParaValue) Values ('" + strTestNameDataStore + "','"+ strDataSetDataStore +  "','" + gstrKeyword + "','" + strKWActionDataStore + "','" + strParameter + "','" + strValue + "')"
    cExcellogger().logToExcel "", gstrDataStoreExcelPath, strSQLStatement

End Sub

Public Function  fetchFromDataStore(strKeywordName, strKeywordAction, strParameter)

   Dim strTestNameDataStore, strDataSetDataStore, strKWActionDataStore
	Dim arrayTest

	arrayTest = Split(gstrDataset,"|")
	strTestNameDataStore = arrayTest(0)

	If  (UBound(arrayTest ) > 0 )Then
		strDataSetDataStore = arrayTest(1)
	else
		strDataSetDataStore = "NA"
	End If

	strKWActionDataStore = gstrAction	

   strParameter = Trim(strParameter)

	Dim arrFetchedValue
	arrFetchedValue = cExcelDataEngine().fetchRowValuesForAColumnBasedFiveSearchValues("Common", "ParaValue", gstrDataStoreExcelPath, "TestcaseName", strTestNameDataStore,"DataSet", strDataSetDataStore,"KeywordName",strKeywordName,"KeywordAction",strKeywordAction,"Parameter" , strParameter)

	fetchFromDataStore = arrFetchedValue

End Function
