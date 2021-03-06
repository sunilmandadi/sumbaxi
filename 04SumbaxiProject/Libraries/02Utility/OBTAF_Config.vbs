'  ******************************************** Global variables*************************************************
'''<summary>This file holds all the Global variables resused in OBTAF across projects</summary>


Environment.Value("intLoginCounter") = 0
Environment.Value("intBrowserLaunchCounter") = 0
Environment.Value("intMaxLogoutTimeInSeconds") = 6900 '6900 Seconds/115 Minutes/ 1 hr 55 Minutes
Environment.Value("intTotalExecutionTime") = 0
Environment.Value("strPreviousUser") = ""

Dim bSetBaseState
Dim bAppState
Dim bAppStateLogout

Dim bLogOff
Dim bCloseBrowsers

Dim bHasKWDNext
Dim bDemo

Dim bLISAExecution
Dim bExecuteNavigation
Dim bRecoverData
Dim bWordLog
Dim bReLaunchBrowser

bSetBaseState = True 	'[If True executed Base state by invoking AUT]

bAppState = False	 	'[For Batch Run make bAppState = false and for Normal Run bAppState = true]
bLogOff = False 	 	'[For Batch Run make bLogOff = false and for Normal Run bLogOff = true]
bCloseBrowsers = False  '[For Batch Run make bCloseBrowsers = false and for Normal Run bCloseBrowsers = true]
bReLaunchBrowser = False '[IF True Relaunch application will be called]

bHasKWDNext = True
bDemo = False
bAppStateLogout = True 'Runtime veriable to control logging out  at data set level. 'make it true for multiple TCs

'**This is temporary set for LISA executions when C3 not available
bLISAExecution = False
bExecuteNavigation = True
bRecoverData = True
'strSpecialCharacters="$,%,@,#,&,_,-"
bWordLog = False 'True if  Word file logging is to be activated

Dim gstrRunTimeStmtDate'To get Statement date from statement page
Dim gstrRunTimePymtDueDate'To get Payment Due Date from statement page
Dim gstrRuntimeCommentStep ' TO ULSTP Submit Time Stamp
Dim gstrRuntimeInterestRateStep ' To ULSTP Interest Rate
Dim gstrRuntimeAdministrativeFeeStep ' To ULSTP AdministrativeFeeStep
Dim gstrRuntimeEffectiveInterestRateStep ' To ULSTP Effective Interest Rate
Dim gstrParameterNameStep ' To ULSTP Submit Time Stamp

'Dim gstrRuntimeSRNumStep 'view IA
'Dim gstrRuntimeTMCommentStep 'view IA
'Dim gstrRuntimeCommentStep 'view IA
'******** For Current Credit Limit Table
Dim strRunTimeRtlCrLimit'This Variable is used to Store Relationship credit limit in runtime
Dim strRunTimeAcctTotalCrLimit'This Variable is used to Store Account Total Credit Limit in runtime
Dim strRunTimeEmbosserCrLimit'This Variable is used to Store Embosser Credit Limit in runtime

'******* For Outstanding Balances Table
Dim strRunTimeRtlOutStandingBal'This Variable is used to Store Relationship Outstanding Balance
Dim strRunTimeAcctOutStandingBal'This Variable is used to Store Account Outstanding Balance
Dim strRunTimeEmbosserOutStandingBal'This Variable is used to Store Embosser Outstanding Balance
'''***************************************** TestData Sheet Variables*********************************************

'<summary> These variables holds the column structure represented by the test data  sheet/table</summary>		
	
	Dim gstrDay
	Dim gstrDataset
	Dim gstrMachine
	Dim gstrKeyword 
	Dim gstrAction
	Dim gstrStartPage
	Dim gstrEndPage
	Dim gstrBrowser
	
	Dim gstrExceptionType 
	Dim gstrExpDetails 
	Dim gstrExpAction 
	
	Dim gstrExceptionData
	
	Dim arrColumnSeq,gstrApptype
	gstrApptype = "window"
	arrColumnSeq = Array ("TestCase", "TestCaseONOFF","Machine", "Keyword" , "Sequence", "Action","StartPage","EndPage","ExceptionType","ExpDetails", "ExpAction","Arguments")
	
	Dim iDay,iDataSet,iDataSetONOFF,iMachine,iKeyword,iSequence,iAction,iExceptionType,iExpDetails,iExpAction, iArguments, iStartPage, iEndPage, iBrowser
	Dim sTestcase
	'sTestcase=arrStringPos ( arrColumnSeq, "TestcaseName" )
	
	iDay =  arrStringPos ( arrColumnSeq, "Day" )
	iDataSet  = arrStringPos ( arrColumnSeq, "TestCase")
	iDataSetONOFF =  arrStringPos ( arrColumnSeq, "TestCaseONOFF")
	iMachine  =  arrStringPos ( arrColumnSeq, "Machine")
	iKeyword  =  arrStringPos ( arrColumnSeq, "Keyword")
	iSequence  =  arrStringPos ( arrColumnSeq, "Sequence")
	iAction  =  arrStringPos ( arrColumnSeq, "Action")
	
	iStartPage  =  arrStringPos ( arrColumnSeq, "StartPage")
	iEndPage  =  arrStringPos ( arrColumnSeq, "EndPage")
	iBrowser  =  arrStringPos ( arrColumnSeq, "Browser")
	
	iExceptionType  =  arrStringPos ( arrColumnSeq, "ExceptionType")
	iExpDetails  =  arrStringPos ( arrColumnSeq, "ExpDetails")
	iExpAction  =  arrStringPos ( arrColumnSeq, "ExpAction")
	iArguments  =  arrStringPos ( arrColumnSeq, "Arguments")

	''***************************************** TestData Sheet Variables*********************************************
	
	''***************************************** DD Testcase Sheet Variables*******************************************
	
'<summary> These variables holds the column structure represented by the DD Testcase</summary>		  
          
	Dim arrDDTestColumnSeq		
	arrDDTestColumnSeq = Array ( "TestCase", "TestCaseONOFF","Machine", "Keyword" , "Action","Sequence", "DDPath", "DDSheetName")
	
	Dim giDDTestDay, giDDTestTestcase, giDDTestTestcaseONOFF, giDDTestMachine, giDDTestKeyword, giDDTestAction, giDDTestSequence, giDDTestPath, giDDTestDDSheetName
	
	giDDTestDay =  arrStringPos ( arrDDTestColumnSeq, "Day" )
	giDDTestTestcase  = arrStringPos ( arrDDTestColumnSeq, "TestCase")
	giDDTestTestcaseONOFF =  arrStringPos ( arrDDTestColumnSeq, "TestCaseONOFF")
	giDDTestMachine  =  arrStringPos ( arrDDTestColumnSeq, "Machine")
	giDDTestKeyword  =  arrStringPos ( arrDDTestColumnSeq, "Keyword")
	giDDTestAction  =  arrStringPos ( arrDDTestColumnSeq, "Action")
	giDDTestSequence  =  arrStringPos ( arrDDTestColumnSeq, "Sequence")
	giDDTestPath  =  arrStringPos ( arrDDTestColumnSeq, "DDPath")
	giDDTestDDSheetName  =  arrStringPos ( arrDDTestColumnSeq, "DDSheetName")
	
	Dim giTestDDKWConfigDetailsCount, arrayKWConfigDetails
	arrayKWConfigDetails = Array( "StartPage", "EndPage",  "ExceptionType",  "ExpDetails", "ExpAction")
	giTestDDKWConfigDetailsCount = Ubound(arrayKWConfigDetails) + 1
		 
'***************************************** DD Testcase Sheet Variables ********************************************


'  ***************************************** DD DataSheet Variables*********************************************
'<summary> These variables holds the column structure represented by the DD DataSheet</summary>		  
		  
          
		  Dim arrDDDataSheetColumnSeq		
		  arrDDDataSheetColumnSeq = Array ( "DataSet", "DataSetONOFF","RecordType", "Sequence" , "KWConfigStart")

		  Dim giDDDataSheetDataSet, giDDDataSheetDataSetONOFF, giDDDataSheetRecordType, giDDDataSheetSequence, giDDDataSheetKWConfigStart
          
		  giDDDataSheetDataSet =  arrStringPos ( arrDDDataSheetColumnSeq, "DataSet" )
		  giDDDataSheetDataSetONOFF  = arrStringPos ( arrDDDataSheetColumnSeq, "DataSetONOFF")
		  giDDDataSheetRecordType =  arrStringPos ( arrDDDataSheetColumnSeq, "RecordType")
		  giDDDataSheetSequence  =  arrStringPos ( arrDDDataSheetColumnSeq, "Sequence")
		  giDDDataSheetKWConfigStart  =  arrStringPos ( arrDDDataSheetColumnSeq, "KWConfigStart")	
		 

''***************************************** DD DataSheet Variables********************************************

'******************************************Config file variable Initiation***********************************************
Dim gstrCurrentProjectDir
gstrCurrentProjectDir = getMachineEnviromentalVariable("User", "OBTAFProjectRoot")

Dim gstrProjectName
gstrProjectName = getMachineEnviromentalVariable("User", "OBTAFProjectName")

Dim gstrConfigFile
gstrConfigFile = gstrCurrentProjectDir & "\" & gstrProjectName & "\Config\Config.INI"

Dim gstrProjectPath
'gstrProjectPath=getConfig("Config","ProjectDir")


		Dim gstrLogPath
		gstrLogPath = gstrCurrentProjectDir + "\"+getConfig("Excel","ResultLog")'"E:\QTPDemoKWDriven\ResultLog\ResultLogs.xlsx"

		Dim gstrLogPath_wordfile
		gstrLogPath_wordfile=gstrCurrentProjectDir + "\"+getConfig("Word","ResultLog")

		Dim gstrCaptureScreenPath
		gstrCaptureScreenPath = gstrCurrentProjectDir + "\"+getConfig("Config","CaptureScreen ")'
		
		'For upload files
		Dim gstrAttachmentsPath
		gstrAttachmentsPath = gstrCurrentProjectDir + "\"+getConfig("Config","UploadFiles ")

		Dim gstrBackUpChromeLog
		gstrBackUpChromeLog=getConfig("Config","ChromeLogBackup")
		'@@@@@@@@@@ Todo : - modification required:TestData file shound be read from master
	   Dim gstrTestDataPath
       'gstrTestDataPath="F:\QTPDemoKWDriven\TestData\Testdata.xls"

	  Dim gstrMasterSheetPath
      gstrMasterSheetPath = gstrCurrentProjectDir + "\"+getConfig("Excel","MasterSheet ")'
	  
	  'gstrTestDataPath="E:\QTPDemoKWDriven\TestData\Testdata_Flight_Demo.xlsx"

Dim gstrKeywordClassMap
gstrKeywordClassMap=gstrCurrentProjectDir + "\"+gstrProjectName+"\"+getConfig("Excel","ClassMap ") 

Dim gstrTagRepository
gstrTagRepository=gstrCurrentProjectDir + "\"+gstrProjectName+ "\"+getConfig("Excel","TagRepository") '"E:\QTPDemoKWDriven\ObjectRepository\Repository.xlsx"

Dim gstrExceptionMap
gstrExceptionMap=gstrCurrentProjectDir + "\"+gstrProjectName+"\"+getConfig("Excel","ExceptionMap") '"E:\QTPDemoKWDriven\Keywords\ExceptionClassMap.xlsx"

Dim gstrConnectionSrting
gstrConnectionString =getConfig("DB","ConnectionString")  '"DRIVER={Microsoft Excel Driver (*.xls)};DBQ="& gstrLogPath & ";Readonly=True"

Dim gstrAccountMapExcelPath
gstrAccountMapExcelPath=gstrCurrentProjectDir + "\"+gstrProjectName+"\"+"Keywords\AccountMap.xls" '"E:\QTPDemoKWDriven\Keywords\ExceptionClassMap.xlsx"

Dim gstrDataStoreExcelPath
gstrDataStoreExcelPath = gstrCurrentProjectDir + "\"+gstrProjectName+"\"+"Keywords\DataStore.xls" '"E:\QTPDemoKWDriven\Keywords\ExceptionClassMap.xlsx"

Dim gstrExecutionFramework
gstrExecutionFramework = getConfig("Config", "Framework")

Dim gstrExecutionEnvironment
gstrExecutionEnvironment = getConfig("Config", "Environment")

Dim gstrFeatureFile
gstrFeatureFile = gstrCurrentProjectDir & "\" & getConfig("Config", "FeatureFile")

Dim gstrTextLog
gstrTextLog = gstrCurrentProjectDir & "\" & getConfig("Config", "TextLog")

'******************************************Config file variable Initiation***********************************************    
'This is Config parameter. Valid values are 0, 1, 2, 3
'0=Only reporting for RSLT enabled
'1=Only reporting for RSLT and WARN enabled
'2=ROnly reporting for RSLT , WARN , INFO enabled
'3=All Reports enabled

Dim gstrLogLevel
gstrLogLevel=2 'This is Config parameter. Valid values are 0, 1, 2, 3
bLogging =TRUE

'*******************************All below parameters will come from Master sheet********************************
Dim gstrBuild
Dim gstrClient  
Dim gstrTestCase
Dim gstrKWSheetPath
'Dim gstrTestPlanName
Dim gstrProcess
Dim gstrFeature
'Dim gstrBrowser
Dim gstrCheckObjectBefore

'********************************************************************************************************************

' @@@@@ TODO: To be deleted after setting uo batch execution - 
gstrMachine="NotSpecified"
'gstrLogLevel =0  ' Configration parameter
'gstrMachine =".omnitechinfo.com"
gstrBuild = "1.0.0"
'		gstrProcess="InitialFramework"
'		gstrFeature="InitialFramework"
'		gstrActionName =" Set Texts"
'		gstrScriptName = "ExcelComponent"
'		gstrTestPlanName="Framework"
gstrBrowser=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Config" , "BrowserType" )'"IE"
'********************************************************************************************************************		
'Exception Number at runtime

Dim gstrExceptionNumber
Dim gstrExceptionMessage
Dim gstrApplicationURL

'MsgBox gstrCurrentProjectDir  &  "\" & gstrProjectName &  "\Config\Config.ini"

gstrApplicationURL = readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "URL" , "APPURL" )


'Authentication in Application
Dim gstrCSOUserName
gstrCSOUserName =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "CSOUSERNAME" )

Dim gstrCSOPassword
gstrCSOPassword =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "CSOPASSWORD" )

Dim gstrCSOUserName1
gstrCSOUserName1 =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "CSOUSERNAME1" )

Dim gstrCSOPassword1
gstrCSOPassword1 =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "CSOPASSWORD2" )
Dim gstrDefaultUserType, gstrDefaultUserName, gstrDefaultPassword

gstrDefaultUserType=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "DEFAULTUSERTYPE" )
gstrDefaultUserName=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "DEFAULTUSERNAME" )
gstrDefaultPassword=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "Authentication" , "DEFAULTPASSWORD" )

'************************Database Variables************************************************************
Dim gstrICALLDB
gstrICALLDB =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDB" )
Dim gstrSID
gstrSID =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDB_SID" )
Dim gstrICallDBUser
gstrICallDBUser =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBUser" )
Dim gstrICallDBPassword
gstrICallDBPassword =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBPassword" )
Dim gstrICallDb_Env
gstrICallDb_Env =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDb_Env" )
Dim gstrICallDBPort
gstrICallDBPort = readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBPort" )
Dim gstrICallDbName_CC
gstrICallDbName_CC =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CC" )
Dim gstrICallDB_CC
gstrICallDB_CC =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDB_CC" )
Dim gstrICallDBUser_CC
gstrICallDBUser_CC =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBUser_CC" )
Dim gstrICallDBPassword_CC
gstrICallDBPassword_CC =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBPassword_CC" )
Dim gstrSID_CC
gstrSID_CC =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBCC_SID" )

Dim gstrICallDbName_Branch
gstrICallDbName_Branch =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_Branch" )
Dim gstrICallDB_Branch
gstrICallDB_Branch =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDB_Branch" )
Dim gstrICallDBUser_Branch
gstrICallDBUser_Branch =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBUser_Branch" )
Dim gstrICallDBPassword_Branch
gstrICallDBPassword_Branch =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBPassword_Branch" )
Dim gstrICallDBBranch_SID
gstrICallDBBranch_SID =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDBBranch_SID" )



If Ucase(gstrExecutionEnvironment) ="LISA" Then '[Newly Added to Configur LISA DB Instance]
	'Configuration for FE DB	
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_LISA_FE" )
	'Configuration for OL DB	
	gstrICallDbName_CC_OL =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_LISA_OL" )
	
elseIf Ucase(gstrExecutionEnvironment) ="SIT" Then
	'Configuration for FE DB	
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CC_FE_UAT" )
	'Configuration for OL DB	
	gstrICallDbName_CC_OL =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CC_OL_UAT" )

elseIf Ucase(gstrExecutionEnvironment)="UAT" Then
	'Configuration for FE DB	
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CC_FE_UAT" )
	'Configuration for OL DB	
	gstrICallDbName_CC_OL =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CC_OL_UAT" )

elseIf Ucase(gstrExecutionEnvironment)="CS_SIT" Then
	'Configuration for FE DB	
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CS_FE_SIT" )
	'Configuration for OL DB	
	gstrICallDbName_CC_OL =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CS_OL_SIT" )

elseIf Ucase(gstrExecutionEnvironment)="CS_UAT" Then
	'Configuration for FE DB	
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CS_FE_UAT" )
	'Configuration for OL DB	
	gstrICallDbName_CC_OL =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_CS_OL_UAT" )
elseIf Ucase(gstrExecutionEnvironment)="1602_SIT" Then
	'Configuration for FE DB in 1602     
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_1602_FE_SIT" )
	'Configuration for OL DB	
	gstrICallDbName_CC_OL =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_1602_OL_SIT" )
elseIf Ucase(gstrExecutionEnvironment)="1603_SIT" Then
	'Configuration for FE DB in 1603     
	gstrICallDbName_CC_FE =  readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "DB" , "ICallDbName_1603_FE_SIT" )
Else
	msgbox "Invalid Environment Option selected. Please select correct option"
End If 

'*************************************************BDD *********************************************************************
'This Section is for New  Behaviour Driven Test Automation Framework approach: 
Dim gstrTestSheetName:gstrTestSheetName=""
Dim gstrRuntimeArgDict,gstrRuntimeStepKeys
'Set gstrRuntimeArgDict =CreateObject("Scripting.Dictionary")
'Set gstrRuntimeStepKeys=CreateObject("Scripting.Dictionary")
Dim bDevPending:bDevPending=False
Dim arrBDDTestColumnSeq		
arrBDDTestColumnSeq = Array ( "Feature", "Scenario","ONOFF", "Data Drive" ,"Recovery_Option", "NavigationKey","Steps","StepLibrary")
giBDFeature =  arrStringPos ( arrBDDTestColumnSeq, "Feature" )
giBDScenario =  arrStringPos ( arrBDDTestColumnSeq, "Scenario" )
giBDONOFF =  arrStringPos ( arrBDDTestColumnSeq, "ONOFF" )
giBDDataDriven=  arrStringPos ( arrBDDTestColumnSeq, "Data Drive" )
giBDRecoveryOption=arrStringPos ( arrBDDTestColumnSeq, "Recovery_Option" )
giBDNavigationKey=  arrStringPos ( arrBDDTestColumnSeq, "NavigationKey" )
giBDSteps =  arrStringPos ( arrBDDTestColumnSeq, "Steps" )
giBDStepLibrary=arrStringPos ( arrBDDTestColumnSeq, "StepLibrary" )

Dim arrBDDataColumnSeq
arrBDDDataColumnSeq= Array ( "Scenario","Data Set","ONOFF","RecordType", "Arguments" )
giBDDData_Scenario =  arrStringPos ( arrBDDDataColumnSeq, "Scenario" )
giBDDData_DS =  arrStringPos ( arrBDDDataColumnSeq, "Data Set" )
giBDDData_ONOFF =  arrStringPos ( arrBDDDataColumnSeq, "ONOFF" )
giBDDData_RecordType =  arrStringPos ( arrBDDDataColumnSeq, "RecordType" )
giBDDData_Arguments =  arrStringPos ( arrBDDDataColumnSeq, "Arguments" )

Dim gstrNavigationKeyStore
gstrNavigationKeyStore = readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "EXCEL" , "NavigationStore")
'gstrNavigationKeyStore = gstrCurrentProjectDir & "\" & gstrProjectName & "\" & gstrNavigationKeyStore

Select Case UCase(Environment.Value("Region"))
	Case "MUMBAI"
		gstrNavigationKeyStore = gstrCurrentProjectDir & "\" & gstrProjectName & "\" & gstrNavigationKeyStore & "\" & "Mumbai_NavigationKeyStore.xlsm"
	Case "CHENNAI"
		gstrNavigationKeyStore = gstrCurrentProjectDir & "\" & gstrProjectName & "\" & gstrNavigationKeyStore & "\" & "Chennai_NavigationKeyStore.xlsm"
	Case "HYDERABAD"
		gstrNavigationKeyStore = gstrCurrentProjectDir & "\" & gstrProjectName & "\" & gstrNavigationKeyStore & "\" & "Hyderabad_NavigationKeyStore.xlsm"
End Select

'***************************For IBM Percomm Sessions for V+ and KRSP************************************************
'Dim gstrSessionFile, gstrUser_KRSP,gstrPassword_KRSP,gstrClient_VPlus,gstrUser_VPlus,gstrPassword_VPlus,gstrUser_CIS,gstrPassword_CIS
If Ucase(gstrExecutionEnvironment) ="SIT" OR Ucase(gstrExecutionEnvironment) = "1602_SIT" OR Ucase(gstrExecutionEnvironment) = "CS_SIT" OR Ucase(gstrExecutionEnvironment) = "LISA" Then
	 gstrSessionFile=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "SessionFile" )
	 gstrRegion_KRSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "KRSP_Region" ) 
	 gstrUser_KRSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "KRSP_User" )
	 gstrPassword_KRSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "KRSP_Password" )
	
	gstrRegion_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_Region" ) 
	gstrClient_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_Client" )
	gstrUser_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_UserName" )
	gstrPassword_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_Password")
	
	gstrRegion_CIS=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "CIS_Region" ) 
	gstrUser_CIS=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "CIS_User" )
	gstrPassword_CIS=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "CIS_Password" )
	
	gstrRegion_FTSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "FTSP_Region" ) 
	gstrUser_FTSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "FTSP_User" )
	gstrPassword_FTSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "FTSP_Password" )
	
	gstrRegion_DCBX=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DCBX_Region" ) 
	gstrUser_DCBX=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DCBX_User" )
	gstrPassword_DCBX=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DCBX_Password" )
	
	gstrRegion_DACB=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DACB_Region" ) 
	gstrUser_DACB=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DACB_User" )
	gstrPassword_DACB=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DACB_Password" )
	
	elseIf Ucase(gstrExecutionEnvironment)="UAT" OR Ucase(gstrExecutionEnvironment)="CS_UAT" Then
	gstrSessionFile=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "SessionFile" )
	gstrRegion_KRSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "KRSP_Region_UAT" ) 
	gstrUser_KRSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "KRSP_User_UAT" )
	gstrPassword_KRSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "KRSP_Password_UAT" )
	
	gstrRegion_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_Region_UAT" ) 
	gstrClient_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_Client_UAT" )
	gstrUser_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_UserName_UAT" )
	gstrPassword_VPlus=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "VPlus_Password_UAT" )
	
	gstrRegion_CIS=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "CIS_Region_UAT" ) 
	gstrUser_CIS=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "CIS_User_UAT" )
	gstrPassword_CIS=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "CIS_Password_UAT" )
	
	gstrRegion_FTSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "FTSP_Region_UAT" ) 
	gstrUser_FTSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "FTSP_User_UAT" )
	gstrPassword_FTSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "FTSP_Password_UAT" )
	
	gstrRegion_DCBX=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DCBX_Region_UAT" ) 
	gstrUser_DCBX=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DCBX_User_UAT" )
	gstrPassword_DCBX=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DCBX_Password_UAT" )
	
	gstrRegion_DACB=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DACB_Region_UAT" ) 
	gstrUser_DACB=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DACB_User_UAT" )
	gstrPassword_DACB=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "DACB_Password_UAT" )
	
	gstrRegion_GPSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "GPSP_Region_UAT" ) 
	gstrUser_GPSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "GPSP_User_UAT" )
	gstrPassword_GPSP=readFromINIFile(gstrCurrentProjectDir + "\"+gstrProjectName+ "\Config\Config.ini",  "IBM PERCOM" , "GPSP_Password_UAT" )
Else
	msgbox "Invalid Environment Option selected. Please select correct option"
End If 
