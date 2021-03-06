Class KW_iCallLogin
	Public Function runKey (strKeyArgument)
		Dim strUserType,strUserName, strPassword, strExpectedStatus, strMessage 

			strUserType  = checkNull(strKeyArgument(0))		

			Select Case Trim(UCase(strKeyArgument (1)))
			
				Case "CSOUSERNAME":
					strUserName = checkNull(gstrCSOUserName)
			
				Case "CSOUSERNAME1":
					strUserName = checkNull(gstrCSOUserName1)
			
				Case Else  strUserName = checkNull(strKeyArgument (1))
			
			End Select

			Select Case Trim(UCase(strKeyArgument (2)))
			
				Case "CSOPASSWORD":
					strPassword = checkNull(gstrCSOPassword)
			
				Case "CSOPASSWORD1":
					strPassword = checkNull(gstrCSOPassword1)
			
				Case Else  strPassword = checkNull(strKeyArgument (2))
			
			End Select

			strExpectedStatus = checkNull(strKeyArgument(3))
			strMessage = checkNull(strKeyArgument (4))

		'Call to GUI Function			
		If   login_iCall(strUserType, strUserName, strPassword,strExpectedStatus,strMessage ) Then
			LogMessage "RSLT", "Verification", "Keyword  iCallLogin Executed successfully", true
			runKey = true
		else		      
			runKey = false
		End If		 
	
   End Function
End Class
