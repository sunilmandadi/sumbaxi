Print ""
Print "============="
Print "Started ["&Environment.Value("Testcase")&"] Scenario execution at : "& now
Print "============="

Dim bExecKeywords, bCommandLine, bTriggeredFromQC
Dim strProcess, strFeature, strDay,strTestCase,strKWSheetPath, strDataSetFilter, strDataSetName, DD_KWexe

bCommandLine = True
bExecKeywords = True
bTriggeredFromQC = Parameter("bExecute_This_DataSet")

If  bTriggeredFromQC = True Then
	strDataSetName = Environment.Value("TestName") 'In QTP, each OBTAFdataset is created as test case
	Environment.Value("DD_KWexe") = Parameter("bExecute_All_DataSet") 
Else  'Need to be filled when executing from QTP
	strDataSetName ="Data_Story26"	'"<DataSetName>"  'Replace with dataset name to run a particular test case from QTP
	'global variable to execute KW dataset by retriving all the data from the datadrive sheet ELSE we can execute the select dataset from datadrive sheet.
	Environment.Value("DD_KWexe") = True
End If

If (bCommandLine = False) Then
	strProcess = "CardList"
	strFeature = "Story 220"                               
	strDay = "Day1"
	strTestCase = "CardList" 
	strKWSheetPath = "D:\iservetaf\02TestData\Singapore\Enquiry\ENQUIRY_CardList.xlsm" '"D:\OBTAF_QTP\testdata_UAT\Story26_ICALL_CL_TopUp.xlsx"
	strDataSetFilter = 	strDataSetName
Else
	strProcess = Environment.Value("Process")
	strFeature = Environment.Value("Feature")
	strTestCase = Environment.Value("Testcase")
	strKWSheetPath = Environment.Value("KWSheetPath")
	strDataSetFilter =  Environment.Value("KWSheetPath")  'This needs modification in the batch creation process to include data set
End If

Print "Process is "& strProcess	
Print "Feature is "& strFeature
Print "Testcase is "& strTestCase
Print "KWSheetPath is "& strKWSheetPath

gstrProcess = strProcess
gstrFeature = strFeature
gstrDay = strDay
gstrTestCase = strTestCase
gstrKWSheetPath = strKWSheetPath

If Ucase(gstrExecutionFramework) = "OBTAF" Then 	'[This option will execute all Business Driven Testcase]
	Set oTest = cDriveEngine
	oTest.ExecuteKWDrivenTest strTestCase,strKWSheetPath,bExecKeywords, strDataSetFilter
ElseIf Ucase(gstrExecutionFramework) = "BDT" Then  	'[This Option will execute all behaviour Driven Tests]
	bExecKeywords = True
	Set oBDT = cBDTEngine
	oBDT.ExecuteBDTest strTestCase, strKWSheetPath,bExecKeywords
Else
	Msgbox "Invalid Framework Option selected. Please select correct option"
End If

Print ""
Print "============="
Print "Completed ["&Environment.Value("Testcase")&"] Scenario execution at : "& now
Print "============="
