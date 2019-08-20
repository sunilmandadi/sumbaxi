@echo off
echo ------------------------------------------------------------
echo Test Case Execution Started at : %date% %Time%
echo ------------------------------------------------------------
echo We are in the Execution Control Batch File

REM Project Root
Set ProjectRoot=%OBTAFProjectRoot%
echo Project Root is %ProjectRoot%

Set Project=%OBTAFProjectName%
echo Project Name is %Project%

echo ------------------------------------------------------------
echo Renaming the existing Testlog ScreenShot and UFTResult
call %OBTAFProjectRoot%\07BatchFiles\Common\PreSetup.bat
echo ------------------------------------------------------------

Set MailSubject=I.SERVE_INDONESIA_NSTP_REGRESSION_EXECUTION_BASELINE_1

Set MailToList=vishnupriya@dbs.com;santhanakrishn1@dbs.com;balakumaran@dbs.com;kolanoorthomas@dbs.com;sunilreddy@dbs.com;srinivasulu@dbs.com
Set MailCCList=sanjay@dbs.com;ranjiniashok@1bank.dbs.com;rohitsawant@dbs.com;raghuramk@dbs.com

Set strDriverScriptPath=%ProjectRoot%\03DriverScript\ISERVETAFEngine
Set strBatchFilePath=%ProjectRoot%\%Project%\Libraries\06Batch
Set strTestDataPathEnq=%ProjectRoot%\02TestData\Indonesia\Enquiry
Set strTestDataPathSTP=%ProjectRoot%\02TestData\Indonesia\STP
Set strTestDataPathNSTP=%ProjectRoot%\02TestData\Indonesia\NSTP


REM ----------------------------Starting ID STP Regression Suite Execution------------------------------

echo --------------------------------------------------
echo Test Case Execution ID_NSTP_ViewIA NSTP Starting : %date% %Time%
echo --------------------------------------------------
echo %strDriverScriptPath%
cscript %strBatchFilePath%\setRuntimeEnv.vbs DBS Indonesia IServe_ID_Regression ID_NSTP_ViewIA %strTestDataPathNSTP%\ID_NSTP_ViewIA.xlsm ID_NSTP_ViewIA
cscript %strBatchFilePath%\QTPexecute.vbs %strDriverScriptPath% %strTestDataPathNSTP%\ID_NSTP_ViewIA.xlsm ID_NSTP_ViewIA
echo --------------------------------------------------
echo Test Case Execution ID_NSTP_ViewIA NSTP completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ID_NSTP_NewIA NSTP Starting : %date% %Time%
echo --------------------------------------------------
echo %strDriverScriptPath%
cscript %strBatchFilePath%\setRuntimeEnv.vbs DBS Indonesia IServe_ID_Regression ID_NSTP_NewIA %strTestDataPathNSTP%\ID_NSTP_NewIA.xlsm ID_NSTP_NEWIA
cscript %strBatchFilePath%\QTPexecute.vbs %strDriverScriptPath% %strTestDataPathNSTP%\ID_NSTP_NewIA.xlsm ID_NSTP_NEWIA
echo --------------------------------------------------
echo Test Case Execution ID_NSTP_NewIA NSTP completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ID_NSTP_EditIA NSTP Starting : %date% %Time%
echo --------------------------------------------------
echo %strDriverScriptPath%
cscript %strBatchFilePath%\setRuntimeEnv.vbs DBS Indonesia IServe_ID_Regression ID_NSTP_EditIA %strTestDataPathNSTP%\ID_NSTP_EditIA.xlsm ID_NSTP_EDITIA
cscript %strBatchFilePath%\QTPexecute.vbs %strDriverScriptPath% %strTestDataPathNSTP%\ID_NSTP_EditIA.xlsm ID_NSTP_EDITIA
echo --------------------------------------------------
echo Test Case Execution ID_NSTP_EditIA NSTP completed : %date% %Time%
echo --------------------------------------------------

REM ----------------------------Ending ID STP Regression Suite Execution------------------------------

cscript %strBatchFilePath%\SendMailResults.vbs %MailSubject% %MailToList% %MailCCList%

echo ------------------------------------------------------------
echo Testcase Execution Ended at : %date% %Time%
echo ------------------------------------------------------------