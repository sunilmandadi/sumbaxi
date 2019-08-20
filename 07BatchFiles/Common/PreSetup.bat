@echo off
REM -------------- Configuration Secion for ISERVE Automation-----------------------------------------------------------------------------------
REM Project Root
Set ProjectRoot=%OBTAFProjectRoot%
echo Project Root is %ProjectRoot%

REM Project
Set Project=%OBTAFProjectName%
echo Project Name is %Project%

REM -------------Clear Cache----------------------------------------------------------------------------------------------
echo %ProjectRoot%\07BatchFiles\Common\ClearCache.bat
call %ProjectRoot%\07BatchFiles\Common\ClearCache.bat
REM ----------------------------------------------------------------------------------------------------------------------

REM -------------Find Current TimeStamp ----------------------------------------------------------------------------------------
set PATH=%ProjectRoot%\11ThirdParty;%PATH%
echo %PATH%
for /f %%i in ('doff ddmmyyhhmmss') do (set ddmmyyhhmmssTimeStamp=%%i)
echo %ddmmyyhhmmssTimeStamp%
REM ----------------------------------------------------------------------------------------------------------------------

REM  -------*Rename existing Testlog and Copy Empty Testlog file *-----------------------------------------------------------
ren %ProjectRoot%\05ResultLog\TESTLOG.xls TESTLOG._%ddmmyyhhmmssTimeStamp%.xls
copy %ProjectRoot%\05ResultLog\emptylog\TESTLOG.xls %ProjectRoot%\05ResultLog\TESTLOG.xls /Y

REM  -------*Rename existing Testlog Doc and Copy Empty Testlog doc file *-----------------------------------------------------------
ren %ProjectRoot%\05ResultLog\TestLOG.doc TestLOG._%ddmmyyhhmmssTimeStamp%.doc
copy %ProjectRoot%\05ResultLog\emptylog\TestLOG.doc %ProjectRoot%\05ResultLog\TestLOG.doc /Y
REM  -------*Rename existing ScreenShot and Create New ScreenShot *-----------------------------------------------------------
ren %ProjectRoot%\05ResultLog\ScreenShots ScreenShots_%ddmmyyhhmmssTimeStamp%
md %ProjectRoot%\05ResultLog\ScreenShots

REM  -------*Rename existing UFTResults and Create New UFTResults *-----------------------------------------------------------
ren %ProjectRoot%\05ResultLog\UFTResults UFTResults_%ddmmyyhhmmssTimeStamp%
md %ProjectRoot%\05ResultLog\UFTResults
REM ----------------------------------------------------------------------------------------------------------------------

REM  -------*Rename existing DataStore and Copy Empty DataStore file *-----------------------------------------------------------
ren %ProjectRoot%\%Project%\Keywords\Datastore.xls Datastore_%ddmmyyhhmmssTimeStamp%.xls
copy %ProjectRoot%\%Project%\Keywords\EmptyDatastore\Datastore.xls %ProjectRoot%\%Project%\Keywords\Datastore.xls /Y
REM ----------------------------------------------------------------------------------------------------------------------