Function exception_VerificationErr
	strErrExpected="Verification Error"
 If gstrExceptionType= strErrExpected Then
		 LogMessage "RSLT", "Verification" , "Expexted Verification error got raised ",true
'		LogMessage "RSLT", "DataSetEnd", "Data set End",true
'		LogMessage "RSLT", "KeywordEnd", "Keyword execution ending successfully",true
  else
	'msgbox "Error is: " &err
	  '
	   Dim strErrorMessage
		 strErrorMessage="Error # " & cStr(err.number) & ", " & err.Description  & " occured at  " & err.source
		' Err.Clear   ' Clear the error.
	   '  LogMessage "RSLT", "Keyword Execution",""&strErrorMessage
	   'msgbox "specified object not found"
		LogMessage "RSLT", "Verification" , "Got error "& err.number &" " &strErrorMessage, false
'		LogMessage "RSLT", "DataSetEnd", "Failed to execute Dataset", false
'		LogMessage "RSLT", "KeywordEnd", "Failed to execute Keyword" , false
	   
   End If
End Function


Function RecoveryFunction_ObjectNotFound(Object, Method, Arguments, retVal)
	 Dim strErrDescription, strObjectName, strMethodName, strArguments,strArg, strDescription
	'strDescription=err.Description
	strArguments=""
	strObjectName=Object.tostring()
	strMethodName=Method
	For i=0 to UBound(Arguments)
		strArg=Arguments(i)
	strArguments=""+strArguments+ ", "+strArg+""
	Next
	
	strErrDescription="Error occured in "+strMethodName +" method with arguments "+strArguments+ " on Object " '+ strObjectName 
	
	  LogMessage "WARN","Exception", "Object  not found, "+strErrDescription ,true
	   gstrExceptionNumber=-424
	   gstrExceptionMessage="Object Not Found"
     	 
End Function
 
Function RecoveryFunction_ObjectNotVisible(Object, Method, Arguments, retVal)
'msgbox err

 Dim strErrDescription, strObjectName, strMethodName, strArguments,strArg, strDescription
	'strDescription=err.Description
	strArguments=""
	strObjectName=Object.tostring()
	strMethodName=Method
	For i=0 to UBound(Arguments)
		strArg=Arguments(i)
	strArguments=""+strArguments+ ", "+strArg+""
	Next
	
	strErrDescription="Error occured in "+strMethodName +" method with arguments "+strArguments+ " on Object " '+ strObjectName 
	
	  LogMessage "WARN","Exception", "Object  not visible, "+strErrDescription ,true
	   gstrExceptionNumber=-424
	   gstrExceptionMessage="Object Not Found"
   
End Function 


 
 
Function RecoveryFunction1(Object, Method, Arguments, retVal)
 
End Function 
 

 
Function RecoveryFunction_ListObjectNotFound(Object, Method, Arguments,retVal)
'On Error Resume Next 
Dim strErrDescription, strObjectName, strMethodName, strArguments,strArg, strDescription, strDialogMessage
'strDescription=err.Description
strArguments=""
strObjectName=Object.tostring()
strMethodName=Method
For i=0 to UBound(Arguments)
	strArg=Arguments(i)
strArguments=""+strArguments+ ", "+strArg+""
Next

strErrDescription="Error occured in "+strMethodName +" method with arguments "+strArguments+ " on Object " '+ strObjectName 

	

   LogMessage "WARN","Exception", "List object "+ strArguments + " not found in List Object " ,true
   gstrExceptionNumber=-427
   gstrExceptionMessage="List Object Item Not Found"
   
   'msgbox "gstrExceptionType"&gstrExceptionType
  ' exception_ObjectNotFound strErrExpected
  ' retVal=-427
  'On Error Resume Next 
 'err.raise -427,"",strErrDescription
 Exit Function
 
 'RecoveryFunction_ListObjectNotFound=err
 
End Function 
 

Function RecoveryFunction_ObjectNotUnique(Object, Method, Arguments, retVal)

 Dim strErrDescription, strObjectName, strMethodName, strArguments,strArg, strDescription
	'strDescription=err.Description
	strArguments=""
	strObjectName=Object.tostring()
	strMethodName=Method
	For i=0 to UBound(Arguments)
		strArg=Arguments(i)
	strArguments=""+strArguments+ ", "+strArg+""
	Next
	
	strErrDescription="Error occured in "+strMethodName +" method with arguments "+strArguments+ " on Object " '+ strObjectName 
	
	  LogMessage "WARN","Exception", "Object  not found, "+strErrDescription ,true
	   gstrExceptionNumber=-426
	   gstrExceptionMessage="Object Not Unique"
 
End Function 

Function RecoveryFunction_PopupWindow(Object, Method, Arguments, retVal)

 Dim strErrDescription, strObjectName, strMethodName, strArguments,strArg, strDescription
	'strDescription=err.Description
	strArguments=""
	strObjectName=Object.tostring()
    strMethodName=Method
	For i=0 to UBound(Arguments)
		strArg=Arguments(i)
	strArguments=""+strArguments+ ", "+strArg+""
	Next
	
	strErrDescription="Error occured in "+strMethodName +" method with arguments "+strArguments+ " on Object " '+ strObjectName 
	
	  LogMessage "WARN","Exception", "Pop up window appeared with the error message, "+strErrDescription ,true
	   gstrExceptionNumber=-419
	   gstrExceptionMessage="Popup Window Appeared"
 
End Function 
 
Function PopupWindow(Object)

LogMessage "WARN","Exception", "Pop up window appeared with the error message, "+strErrDescription ,false
 
End Function 
 
