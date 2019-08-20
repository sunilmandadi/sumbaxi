'BaseState
Dim Counter
Counter=0
While 1 = 1
	Search
	Counter = Counter+1
	Print Counter& " - " & Now
Wend


Public Function BaseState
	SystemUtil.CloseProcessByName "iexplore.exe"

    
  SystemUtil.Run "C:\Program Files\Internet Explorer\iexplore.exe"
'Browser("Browser").Page("Page").Sync
Wait 3

If Dialog("Windows Internet Explorer").Exist(2) then
	Dialog("Windows Internet Explorer").Close
End If
Wait 3

'Window("Windows Internet Explorer").Maximize
Browser("Browser").Navigate "http://10.92.132.225:7055/icall/"

Browser("Browser").Page("ICall").Sync
Browser("Browser").Page("ICall").WebEdit("USername").Set "ccacso1"
Browser("Browser").Page("ICall").WebEdit("password").Set "Password1"
Browser("Browser").Page("ICall").WebElement("Login").Click
Browser("Browser").Page("ICall").Sync
Wait 2

End Function


Public Function Search
   On Error Resume Next

   If  Not Browser("Browser").Page("ICall").WebEdit("html id:=search.nric").Exist(2) Then
	   BaseState
	   	Wait 2
   End If

   Browser("Browser").Page("ICall").WebEdit("CIN").Set "s8970005a"



Browser("Browser").Page("ICall").WebElement("Search").Click
Browser("Browser").Page("ICall").Sync
Wait 2

Browser("Browser").Page("ICall").WebElement("ICALL USER PB S").Click
Browser("Browser").Page("ICall").Sync
Wait 8

'Browser("Browser").Page("ICall").WebEdit("CIN").Submit
If  Browser("Browser").Page("ICall").WebElement("OK").Exist(2) Then
	Browser("Browser").Page("ICall").WebElement("OK").Click
End If
Browser("Browser").Page("ICall").Sync
'Wait 2

stractual =Browser("Browser").Page("ICall").WebTable("Savings Account").RowCount
'msgbox stractual
 strExpected = 9

  strfilename =  Now
 If stractual <> strExpected  Then
		a = Now
		b = Replace(Replace(a,"/",""),":","")
		strfilename = "D:\Screenshot\"&b&".png"
		Desktop.CaptureBitmap strfilename			 
 End If



' Print Counter

Browser("Browser").Page("ICall").Image("search-icon").Click

End Function