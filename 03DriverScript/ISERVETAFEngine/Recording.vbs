int iMedSleep
iMedSleep = 7
int iSmallSleep
iSmallSleep = 3

SystemUtil.Run "C:\Program Files\Internet Explorer\IEXPLORE.EXE","","C:\Documents and Settings\pritam","open"
Browser("DBS Intranet").Page("DBS Intranet").Sync
Browser("DBS Intranet").Navigate "https://obemea-uat.sgp.dbs.com:62444/SSO/ui/SSOLogin.html"
Browser("DBS Intranet").Page("Page").Frame("loginFrame_4").WebEdit("usertxt").Set DataTable("UserName", dtGlobalSheet)
Browser("DBS Intranet").Page("Page").Frame("loginFrame_4").WebEdit("passtxt").SetSecure "4df5c8a46ddc2fb0aa9d8868cbebc4c08214f2a37b0db2eb"
Browser("DBS Intranet").Page("Page").Frame("loginFrame_4").WebButton("Login").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iSmallSleep)
Browser("DBS Intranet").Page("Page").Frame("loginFrame_2").WebList("appSelect").Select "Finacle Core"
Browser("DBS Intranet").Dialog("Microsoft Internet Explorer").WinButton("OK").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iSmallSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_3").WebEdit("menuName").Set "hoaacla"
Browser("DBS Intranet").Page("Page").Sync
wait(iSmallSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_3").WebButton("Go").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iSmallSleep)

Browser("DBS Intranet").Page("Page").Frame("FINW_4").WebEdit("laacop.cifId").Set "2200003"
Browser("DBS Intranet").Page("Page").Frame("FINW_4").WebEdit("laacop.schmCode").Set "ltlaa"
Browser("DBS Intranet").Page("Page").Frame("FINW_4").WebButton("Go").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iSmallSleep)

'*********************************************  Account Interest****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_6").WebElement("A/C Interest").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iMedSleep)


'*********************************************  Loan Details****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_7").WebElement("Loan Details").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_8").WebEdit("lasch.loanAmt").Set "1000"
Browser("DBS Intranet").Page("Page").Frame("FINW_8").WebEdit("lasch.loanPerdMths").Set "12"

Browser("DBS Intranet").Page("Page").Frame("FINW_8").WebEdit("modeOfPayment").Set "ach"
Browser("DBS Intranet").Page("Page").Frame("FINW_8").WebButton("Validate").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iSmallSleep)

'*********************************************  ILA Interest****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_9").WebElement("LA Interest").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_9").WebEdit("laint.prinOvrdPerdMths").Set "12"

'*********************************************  Int Slabs ****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_7").WebElement("Int.Slabs").Click
Browser("DBS Intranet").Page("Page").Sync
wait(iMedSleep)


Browser("DBS Intranet").Page("Page").Frame("FINW_10").WebEdit("linttmacct.intTblCode").Set "AVGPP"
Browser("DBS Intranet").Page("Page").Frame("FINW_13").WebEdit("linttmacct.tenorOfSlabInMnths").Set "3"

'********************************************* Payment Plan ****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_7").WebElement("Payment Plan").Click
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_11").WebEdit("laparm.noOfInstlmnts").Set "4"


'********************************************* Payment  Schedule ****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_11").WebElement("Payment Schedule").Click
wait(iMedSleep)


'********************************************* Amortization  Schedule ****************************************************
Browser("DBS Intranet").Page("Page").Frame("FINW_12").WebButton("Amortization Schedule").Click
wait(iMedSleep)


Browser("DBS Intranet").Dialog("Microsoft Internet Explorer").Activate
Browser("DBS Intranet").Dialog("Microsoft Internet Explorer").WinButton("OK").Click
wait(2)
Browser("DBS Intranet").Dialog("Microsoft Internet Explorer").WinButton("OK").Click
wait(2)
Browser("DBS Intranet").Page("Page").Frame("FINW_13").WebButton("Ok").Click
wait(5)
Browser("DBS Intranet").Page("Page").Frame("FINW_14").WebButton("Submit").Click
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_13").WebElement("Fees").Click
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_7").WebButton("Submit").Click
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Frame("FINW_15").WebButton("Accept").Click
wait(iMedSleep)
Browser("DBS Intranet").Page("Page").Sync

Browser("DBS Intranet").Page("Page").Frame("FINW_5").Link("Logout").Click
Browser("DBS Intranet").Window("Confirm Dialog -- Web").Page("Confirm Dialog").WebButton("Logout").Click
Browser("DBS Intranet").Page("Page").Frame("loginFrame_2").Image("Log out").FireEvent "onmouseover"
Browser("DBS Intranet").Page("Page").Frame("loginFrame_2").Image("Log out").Click
wait(iSmallSleep)
Browser("DBS Intranet").Dialog("Microsoft Internet Explorer").WinButton("OK").Click
wait(iSmallSleep)

Browser("DBS Intranet").WinToolbar("ToolbarWindow32").Press "&File"
Browser("DBS Intranet").WinMenu("ContextMenu").Select "Close"






