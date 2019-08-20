@echo off
echo ------------------------------------------------------------
echo Test Case Execution Started at : %date% %Time%
echo ------------------------------------------------------------
echo We are in the Execution Control Batch File

REM Project Root
Set ProjectRoot=%OBTAFProjectRoot%
echo Project Root is %ProjectRoot%

REM Project
Set Project=%OBTAFProjectName%
echo Project Name is %Project%

echo ------------------------------------------------------------
echo Performing Clear Cache then renaming the existing Testlog ScreenShot UFTResults and DataStore
echo %ProjectRoot%\07BatchFiles\Common\PreSetup.bat
call %ProjectRoot%\07BatchFiles\Common\PreSetup.bat
echo ------------------------------------------------------------

REM ---------------------------------------------------------------------Starting Regression Suite Execution--------------------------------------------------------------
REM cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs "ApplicationURL" "Framework" "DBServerAddress" "DBName" "DBUserName" "DBPassword" "DBPort"

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Balance and Limits Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression BalandLimits %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BalanceAndLimits.xlsm BalandLimits
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BalanceAndLimits.xlsm BalandLimits
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Balance and Limits completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_RelationshipDetails Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression RelationshipDetails %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_RelationshipDetails.xlsm RelationshipDetails
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_RelationshipDetails.xlsm RelationshipDetails
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_RelationshipDetails completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_LimitsandUsage Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression LimitsandUsage %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_LimitsandUsage.xlsm LimitsandUsage
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_LimitsandUsage.xlsm LimitsandUsage
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_LimitsandUsage completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_TopUpDetails Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression TopUpDetails %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_TopUpDetails.xlsm TopUpDetails
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_TopUpDetails.xlsm TopUpDetails
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_TopUpDetails completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_AuthorizationLog Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression AuthorizationLog %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_AuthorizationLog.xlsm AuthorizationLog
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_AuthorizationLog.xlsm AuthorizationLog
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_AuthorizationLog completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_ChequeInfo Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression ChequeInfo %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_ChequeInfo.xlsm ChequeInfo
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_ChequeInfo.xlsm ChequeInfo
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_ChequeInfo completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_InstallmentPlan Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression InstallmentPlan %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_InstallmentPlan.xlsm InstallmentPlan
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_InstallmentPlan.xlsm InstallmentPlan
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_InstallmentPlan completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_InsuranceInfo Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression InsuranceInfo %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_InsuranceInfo.xlsm InsuranceInfo
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_InsuranceInfo.xlsm InsuranceInfo
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_InsuranceInfo completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_DirectDebitArrangement Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression DirectDebit %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_DirectDebitArrangement.xlsm DirectDebit
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_DirectDebitArrangement.xlsm DirectDebit
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_DirectDebitArrangement completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_CardandPinInfomation Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression CardandPin %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_CardandPinInfomation.xlsm CardandPin
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_CardandPinInfomation.xlsm CardandPin
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_CardandPinInfomation completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Deliquency Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression Deliquency %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_Deliquency.xlsm Deliquency
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_Deliquency.xlsm Deliquency
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Deliquency completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_RecurringArrangement Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression Recurring %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_RecurringArrangement.xlsm Recurring
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_RecurringArrangement.xlsm Recurring
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_RecurringArrangement completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_AddAccountLinkage Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression AddAccLinkage %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_AddAccountLinkage.xlsm AddAccLinkage
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_AddAccountLinkage.xlsm AddAccLinkage
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_AddAccountLinkage completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_KeyInfo_Cards Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression KeyInfo %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_KeyInfo_Cards.xlsm KeyInfo
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_KeyInfo_Cards.xlsm KeyInfo
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_KeyInfo_Cards completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_CASA_AccountDetails Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression AccountDetails %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_CASA_AccountDetails.xlsm AccountDetails
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_CASA_AccountDetails.xlsm AccountDetails
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_CASA_AccountDetails completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution Inq_BankAndEarnEnrollment Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression BankandEarnEnrollment %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_BankAndEarnEnrollment.xlsm BankandEarnEnrollment
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_BankAndEarnEnrollment.xlsm BankandEarnEnrollment
echo --------------------------------------------------
echo Test Case Execution Inq_BankAndEarnEnrollment completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution Inq_BankAndEarnSummary Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression BankandEarnSummary %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_BankAndEarnSummary.xlsm BankandEarnSummary
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_BankAndEarnSummary.xlsm BankandEarnSummary
echo --------------------------------------------------
echo Test Case Execution Inq_BankAndEarnSummary completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution Inq_BankandEarnTransactionHistory Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression BankEarnTransactionHistory %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_BankandEarnTransactionHistory.xlsm BankEarnTransactionHistory
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_BankandEarnTransactionHistory.xlsm BankEarnTransactionHistory
echo --------------------------------------------------
echo Test Case Execution Inq_BankandEarnTransactionHistory completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution Inq_CardRewards Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression CardRewards %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_CardRewards.xlsm CardRewards
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_CardRewards.xlsm CardRewards
echo --------------------------------------------------
echo Test Case Execution Inq_CardRewards completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution Inq_CardRewardsTransaction Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression CardRewardsTransaction %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_CardRewardsTransaction.xlsm CardRewardsTransaction
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\Inq_CardRewardsTransaction.xlsm CardRewardsTransaction
echo --------------------------------------------------
echo Test Case Execution Inq_CardRewardsTransaction completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_OtherPlans Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression OtherPlans %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_OtherPlans.xlsm OtherPlans
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_OtherPlans.xlsm OtherPlans
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_OtherPlans completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Statements Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression Statements %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_Statements.xlsm Statements
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_Statements.xlsm Statements
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Statements completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_TransactionHistory_CC_CL_UL_DC_ATM Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression TransHistory %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_TransactionHistory_CC_CL_UL_DC_ATM.xlsm TransHistory
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_TransactionHistory_CC_CL_UL_DC_ATM.xlsm TransHistory
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_TransactionHistory_CC_CL_UL_DC_ATM completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Submission_AddMemo Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression AddMemo %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_Submission_AddMemo.xlsm AddMemo
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\ENQUIRY_Submission_AddMemo.xlsm AddMemo
echo --------------------------------------------------
echo Test Case Execution ENQUIRY_Submission_AddMemo completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_AvaloqPortfolio Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression AvaloqPortfolio %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_AvaloqPortfolio.xlsm AvaloqPortfolio
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_AvaloqPortfolio.xlsm AvaloqPortfolio
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_AvaloqPortfolio completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_SMS_Threshold Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression SMSThreshold %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_SMS_Threshold.xlsm SMSThreshold
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_SMS_Threshold.xlsm SMSThreshold
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_SMS_Threshold completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_SuspensionDetails Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression SuspensionDetails %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_SuspensionDetails.xlsm SuspensionDetails
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_SuspensionDetails.xlsm SuspensionDetails
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_SuspensionDetails completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_InternetBanking Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression InternetBanking %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_InternetBanking.xlsm InternetBanking
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_InternetBanking.xlsm InternetBanking
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_InternetBanking completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_TMS Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression TokenManagement %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_TMS.xlsm TokenManagement
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_TMS.xlsm TokenManagement
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_TMS completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_TMS Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression TokenManagement %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_TMS.xlsm TokenManagement1
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_TMS.xlsm TokenManagement1
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_TMS completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_BillPayment Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression BillPayment %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_BillPayment.xlsm BillPayment
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_BillPayment.xlsm BillPayment
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_BillPayment completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_RecentApplication Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression RecentApplication %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_RecentApplication.xlsm RecentApplication
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_RecentApplication.xlsm RecentApplication
echo --------------------------------------------------
echo Test Case Execution v completed : %date% %Time%
echo --------------------------------------------------

echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_StandingInstructions Starting : %date% %Time%
echo --------------------------------------------------
echo %ProjectRoot%\03DriverScript\ISERVETAFEngine
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setConfig.vbs https://iservedev.corp.dbs.com:18443/sgiserve/iserve.index.html#/login BDT 10.92.145.85 iservesgr1 iservesgr1 Iservesgr1#12 6603
cscript %ProjectRoot%\%Project%\Libraries\06Batch\setRuntimeEnv.vbs DBS Singapore IServe_Regression SI %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_StandingInstructions.xlsm SI
cscript %ProjectRoot%\%Project%\Libraries\06Batch\QTPexecute.vbs %ProjectRoot%\03DriverScript\ISERVETAFEngine %ProjectRoot%\02TestData\Singapore\Enquiry\SG_ENQUIRY_BankingFacilities_StandingInstructions.xlsm SI
echo --------------------------------------------------
echo Test Case Execution SG_ENQUIRY_BankingFacilities_StandingInstructions completed : %date% %Time%
echo --------------------------------------------------

cscript %ProjectRoot%\%Project%\Libraries\06Batch\SendMailResults.vbs IServe_Regression_ENQUIRY_Execution_1702 sampathkumarsk@dbs.com;ranjiniashok@1bank.dbs.com