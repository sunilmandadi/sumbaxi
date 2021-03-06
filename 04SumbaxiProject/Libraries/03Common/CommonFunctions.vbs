'%%%%%%%%%%Object Description Creation Functions%%%%%%%%%%%%%%%%
Dim objDescription
initialized = false

level = ""
desc = ""
object = ""
objectDescription = ""

levelsubdescriptiondelimiter = ","
leveldescdelimiter = "|"
objectdelimiter = "|"
leveldelimiter = "|"
objectsDescriptiondelimiter = "|"

webLevels = "Browser|Page|Frame"
webLevelsDesc = "micclass:=Browser|micclass:=Page|micclass:=Frame|"
objects = "Link|WebButton|WebList|WebEdit|WebArea|WebElement|Image"
objectsDescription = "micclass:=Link|micclass:=WebButton|micclass:=WebList|micclass:=WebEdit|micclass:=WebArea|micclass:=WebElement|micclass:=Image"

winLevels ="Window|Dialog|1Window"
winLevelsDesc ="micclass := Window|micclass := Dialog|micclass :=Window"
winObjects="WinEdit|WinButton"
winObjDescription="micclass:=WinEdit|micclass:=WinButton"


Public Function GenerateLevelDesc (levelstr)
   'Returns the level description for the object
 ' msgbox "Levelstr"&levelstr
  'msgbox "Level" &level
	l = IndexOf(level, levelstr)
   'msgbox I
	If l >=0 Then
		fdesc = level(0) & "(" & Quote(desc(0)) & ")."
		If l >= 1 Then
			fdesc = fdesc + level(1) & "(" & Quote(desc(1)) & ")."
			If 2 >= l Then
				If thirdlevel <> "" Then
					'msgbox thirdlevel
					fdesc = fdesc + level(2) & "(" & Quote(desc(2)) & "," & Quote("name:=" & thirdlevel) & ")."
				End If
			End If
		End If
	End If
   GenerateLevelDesc = fdesc
 ' msgbox fdesc 
End Function


'Public Function CloseBrowsers
'	If Browser("micclass:=Browser").Exist (0) Then
'		Browser("micclass:=Browser").Close
'	End If
'	While Browser("micclass:=Browser", "index:=1").Exist (0)
'		Browser("index:=1").Close
'	Wend
'	If Browser("micclass:=Browser").Exist (0) Then
'		Browser("micclass:=Browser").Close
'	End If
'End Function
'
'Public Function Launch (apptype, val)
'   'Initializes framework and launches the application 
'	If "website" = apptype Then
'		thirdlevel = ""
'	    LogMessage  "WARN", "Initialization", "Initializing Framework for  web Site", true
'		level = split(webLevels, leveldelimiter, -1, 1)
'	  '  msgbox level(0)
'		
'		desc = split(webLevelsDesc, leveldescdelimiter, -1, 1)
'	   ' msgbox desc(0)
'		object = split(objects, objectdelimiter, -1, 1)
'	  '   msgbox object(0)
'		objectDescription = split(objectsDescription, objectsDescriptiondelimiter, -1, 1)
'	  '   msgbox objectDescription(0)
'	  '***************Following line needs to be uncommented-
'  CloseBrowsers
'	   Set IE = CreateObject("InternetExplorer.Application")
'	   ' Set IE = CreateObject("MozillaFirefox.Application")
'		IE.visible = true
'		IE.Navigate val
'		While IE.Busy
'			wait 1
'		Wend
'	
'    else
'		 if "window" = apptype then
'	 ''' Insert code to launch standalone application
'			level = split(winLevels, leveldelimiter, -1, 1)
'	  'msgbox level(0)
'		
'		desc = split(winLevelsDesc, leveldescdelimiter, -1, 1)
'	   'msgbox desc(0)
'		object = split(objects, objectdelimiter, -1, 1)
'	    'msgbox object(0)
'		objectDescription = split(objectsDescription, objectsDescriptiondelimiter, -1, 1)
'	  'msgbox objectDescription(0)
'		'CloseBrowsers
'	
'		InvokeApplication "C:\Program Files\Mercury Interactive\QuickTest Professional\samples\flight\app\flight4a.exe"
'		'Set WshShell = WScript.CreateObject("WScript.Application")
'	   ' intReturn = WshShell.Run("Flight" & WScript.ScriptFullName, 1, TRUE)
''		Set App = CreateObject("WScript.Shell")
''		'IE.visible = true
'		'App.Run "C:\Program Files\Mercury Interactive\QuickTest Professional\samples\flight\app\flight4a.exe",1,true
'		'WshShell.AppActivate "Flight"
'	'	App.WinListView("SysListView32").Activate "" &vals
'		End if 
'	End If
'	initialized = true
'	Launch = true
'   ' msgbox level
'	'msgbox objectDescription
'End Function


Public Function IndexOf (ary, str)
   'returns level index for object
	val = -1
	For i = 0 to UBound(ary)
		If ary(i) = str Then
			val = i
		End If
	Next
	IndexOf = val
  ' msgBox "Indexof "&val
End Function

Public Function Quote (txt)
'Inserts the "  (quote)  at the begining and at end of string
	Quote = chr(34) & txt & chr(34)
	'msgbox chr(34)
End Function


Public Function GenerateObjectDescription (obj, prop)
   'gererate property description for the object
	i = IndexOf(object, obj)
	ndesc = ""
   If i <> -1 Then
		ndesc = obj & "(" & Quote(objectDescription(i)) & "," & Quote(prop) & ")."
	End If
	GenerateobjectDescription = ndesc
    Report micPass, "Object Description Creation", "Object description created as " & ndesc
End Function

Public Function GenerateObjectDescription (obj, prop)
	i = IndexOf(object, obj)
	ndesc = ""
   If i <> -1 Then
		ndesc = obj & "(" & Quote(objectDescription(i)) & "," & Quote(prop) & ")."
	End If
	GenerateobjectDescription = ndesc
	 Report micPass, "Object Description Creation", "Object description created as " & ndesc
	'msgbox ndesc
End Function

Public Function AutoSync
   'This function synchronizes with the application opened
    Execute GenerateLevelDesc("Browser") & "Sync"
End Function

Public Function Report (status, objtype, text)
   'This function reports the event logs at various levels
	Reporter.Filter = rtEnableAll
	Reporter.ReportEvent status, objtype, text
	Reporter.Filter = rfDisableAll
End Function


Function getObjectDetails(strObjectName)
	'This function reads Object Class Property and Property Value from repository file.
	Dim strTable
   strTable ="Repository"
   Dim strConnection,strSQLStatement,strClassName,strStatement,strData,strMyObjectDetails
   
     	strSQLStatement="Select  *  from [TagRepository$] where   ObjectName ='"& strObjectName &"' "
	   'strSQLStatement="Select  Class  from ["& strTable &"$] where   Keyword =' "& strKeyword &"' "
		strClassName = "clsExcelDataEngine"
		strStatement = "Set clsObj = New  " &   strClassName
	    Execute strStatement
		strData=clsObj.FetchExcelValue (strSQLStatement,gstrTagRepository)
	   ' Print("strData:"&strData(0,1))
		'strMyObjectDetails=strData(0,1)
	'	Msgbox "Test"
	 'msgbox strData(0,3)
	  getObjectDetails = strData
   
End Function

Function myObject (strObjectName)

      'Dim apptype
  ' apptype="window"
  
'  if gstrApptype="window" then
'	 ''' Insert code to launch standalone application
'			level = split(winLevels, leveldelimiter, -1, 1)
'	  'msgbox  level(0)
'		
'		desc = split(winLevelsDesc, leveldescdelimiter, -1, 1)
'	   'msgbox desc(0)
'		object = split(winObjects, objectdelimiter, -1, 1)
'	    'msgbox object(0)
'		objectDescription = split(winObjDescription, objectsDescriptiondelimiter, -1, 1)
'	  'msgbox objectDescription(0)
'end if
    
' This function generates the description object of specified class and property .
Dim strObjDetails,strClass,strProperty,strObjectVal
strObjDetails=getObjectDetails(strObjectName)
strClass   =strObjDetails(0,1)
'msgbox strClass
strProperty   =strObjDetails(0,2)
strProperty=strProperty &":="
'msgbox strProperty
strObjectVal = strObjDetails(0,3)

   localDesc = ""
	rval = true
	'thirdlevel=""
   ' msgbox "Thirdlevel " & thirdlevel
	If thirdlevel <> "" Then
		localDesc = GenerateLevelDesc(level(2))
	Else
		localDesc = GenerateLevelDesc(level(1))
	End If
	'msgbox localDesc
	AutoSync()

	localDesc = localdesc & GenerateObjectDescription( strClass ,  strProperty & strObjectVal)
   myObject=localDesc
	' msgbox "Obj " &localDesc
  LogMessage "INFO", "Object Description ", "Object Description for object " &strObjectName & " created as  :" & localDesc, true
  
 ' LogMessage micPass, "Object Description ", "DP object created as follows " & localDesc
End Function

Function mywinObject (strObjectName)
   Dim apptype
   apptype="window"
  if "window" = apptype then
	 ''' Insert code to launch standalone application
			level = split(winLevels, leveldelimiter, -1, 1)
	  'msgbox level(0)
		
		desc = split(winLevelsDesc, leveldescdelimiter, -1, 1)
	'   msgbox desc(0)
		object = split(winObjects, objectdelimiter, -1, 1)
	 '   msgbox object(0)
		objectDescription = split(winObjDescription, objectsDescriptiondelimiter, -1, 1)
	  'msgbox objectDescription(0)
end if    
'Function myObject (strClass,strProperty,strObjectVal)
' This function generates the description object of specified class and property .
Dim strObjDetails,strClass,strProperty,strObjectVal
strObjDetails=getObjectDetails(strObjectName)
strClass   =strObjDetails(0,1)
'msgbox strClass
strProperty   =strObjDetails(0,2)
strProperty=strProperty &":="
'msgbox strProperty
strObjectVal = strObjDetails(0,3)

   localDesc = ""
	rval = true
	If thirdlevel <> "" Then
		localDesc = GenerateLevelDesc(level(2))
	Else
		localDesc = GenerateLevelDesc(level(1))
	End If
	'msgbox localDesc
	AutoSync()

	localDesc = localdesc & GenerateObjectDescription( strClass ,  strProperty & strObjectVal)
   myObject=localDesc
  ' msgbox localDesc
  LogMessage "INFO", "Object Description ", "Object Description for object " &strObjectName & " created as  :" & localDesc, true
  
 ' LogMessage micPass, "Object Description ", "DP object created as follows " & localDesc
End Function


'Object Creator function...I/p =testobject name, o/p = TestObject

Function myObjectCreater (strObjectName)
      'Dim apptype
Dim strObjDetails,strClass,strProperty,strObjectVal
strObjDetails=getObjectDetails(strObjectName)
strClass   =strObjDetails(0,1)
msgbox strClass
strProperty   =strObjDetails(0,2)

msgbox strProperty
strObjectVal = strObjDetails(0,3)
msgbox strObjectVal
Set oDesc = Description.Create() 

    oDesc("micclass").Value =strClass '"WebEdit" 

   oDesc(strProperty).Value =strObjectVal
  
'msgbox(oDesc.tostring)
 set EditCollection=Browser("micclass:=Browser").Page("micclass:=Page").ChildObjects(oDesc)

 If  EditCollection.count =0Then
	 LogMessage "RSLT", "Object Identification ", strObjectName+ ":object not found " , false
	 
	 Err.Raise -424    'raise a user-defined error
	Err.Description = "Object Not Found"
	Err.Source = "Object Identification"
	myObjectCreater=nothing
elseif editCollection.count>1 then
  LogMessage "RSLT", "Object Identification ", strObjectName+ ": more than one object with same description found " , false
	 Err.Raise vbObjectError -424    'raise a user-defined error
	Err.Description = "Object Not Unique"
	Err.Source = "Object Identification"
	myObjectCreater=nothing

 End If
Set odescr =EditCollection(0)
 
set  myObjectCreater=odescr
End Function


'

'——————————————————————————-
Public Function IsContextLoaded(ByRef htContext)
'——————————————————————————-
'Function: IsContextLoaded
'Checks that the current GUI context is loaded
'
'Iterates through the htContext (HashTable) items and executes the Exist method with 0 (zero) as parameter.
'
'
'——————————————————————————-
    Dim ix, items, keys, strDetails, strAdditionalRemarks
 
    '—————————————————————————
    items = htContext.Items
    keys = htContext.Keys

    If not gstrCheckObjectBefore Then
		IsContextLoaded=true
		Exit Function
	End If
    For ix = 0 To htContext.Count-1
		msgbox (Browser("name:=Welcome: Mercury Tours").Exist)
        IsContextLoaded =  items(ix).Exist(0)
        strDetails = strDetails & vbNewLine & "Object #" & ix+1 & ": '" & keys(ix) & "' was"
        If IsContextLoaded Then     
            intStatus = micPass
            strDetails = strDetails & ""
            strAdditionalRemarks = ""
        Else 
            intStatus = micWarning
            strDetails = strDetails & " not"
            strAdditionalRemarks = " Please check the object properties."
        End If
        strDetails = strDetails & " found." & strAdditionalRemarks
    Next
    '—————————————————————————
 
    Reporter.ReportEvent intStatus, "IsContextLoaded", strDetails
'——————————————————————————-
End Function

Public Function getMachineEnviromentalVariable(strVariableType, strVariableName)
     'Declare Variables
	  Dim WshShl, Shell, UserVar

	'Set objects
	Set WshShl = CreateObject("WScript.Shell")
	Set Shell = WshShl.Environment(strVariableType)
	getMachineEnviromentalVariable =  Shell(strVariableName)
    
	'Cleanup Objects
	Set WshShl = Nothing
	Set Shell = Nothing
	
	Exit Function
	
End Function

'Public Function getMachineEnviromentalVariable(strVariableType, strVariableName)
'     'Declare Variables
'	  Dim WshShl, Shell, UserVar
'
'	'Set objects
'	Set WshShl = CreateObject("WScript.Shell")
'	Set Shell = WshShl.Environment(strVariableType)
'	getMachineEnviromentalVariable =  Shell(strVariableName)
'
'	'Cleanup Objects
'	Set WshShl = Nothing
'	Set Shell = Nothing
'	
'	Exit Function
'	
'End Function
'

Public function readFromINIFile(strINIFilePath,  strSection , strKey )

	
	 On Error resume next  
	 'On error goto ErrTrap
	 
	 Extern.Declare micInteger,"GetPrivateProfileStringA", "kernel32.dll","GetPrivateProfileStringA", micString, micString, micString, micString+micByRef, micInteger, micString 
	
	Dim key,  key2 
	Dim strValue,i,strConfigVal
	'strSection="Config"
	'strKey="ProjectPath"
	'strValue=String(32, "-") 
	key = String(1024, "-") 
	
	i = Extern.GetPrivateProfileStringA(strSection,strKey,"NOT_FOUND", key, 1024, strINIFilePath) 
	
	strValue = Left(key,i) 
	
	If strValue = "NOT_FOUND" Then
					LogMessage micFail, "Read Config ","An error occured: could not read Project config parameter,check config.ini file"
	End If
			 
	
				If Err.Number <> 0 Then
				   
					
				   ' Err.Raise(Err.Number, , "Error form Fucntions.GetIniSettings " & Err.Description)
					  LogMessage micFail, "Read Config ","An error occured:  " &  err.description 
				End If
				
	   readFromINIFile = strValue
	   Exit Function
End Function

Public Function itemExistsInArray(arrString, strSearchItem)

     itemExistsInArray = InStr(1, vbNullChar & Join(arrString, vbNullChar) & vbNullChar, _
     vbNullChar & strSearchItem & vbNullChar) > 0

	Exit Function
	
End Function

Public Sub addItemToArray(arrString, strSearchItem)

   If  (UBound(arrString) = 0 AND IsEmpty (arrString(0)) ) Then
    
       arrString(0) = strSearchItem

   else

      ReDim Preserve arrString(UBound(arrString)+1) 
       arrString(UBound(arrString)) = strSearchItem

   End If   
     
End Sub


Public Function returnColumnValuesForARow (arrData, iRowCount)
	ReDim arrColData(0)
	iTotalCols = Ubound(arrData, 2)

	For iColCount = 0 to iTotalCols
		 strTempArg=null
		 strTempArg=arrData(iRowCount,iColCount)
		 addItemToArray arrColData, strTempArg
		 
	 Next

	 returnColumnValuesForARow = arrColData

	Exit Function
	
End Function

Public Sub fill2DArrayRowValue (array2D, arrData, iRowCount, iTotalRows)
  Dim iTotalCols

   iTotalCols = Ubound(arrData)
 

   ReDim Preserve array2D(iTotalRows,iTotalCols )

	For iColCount = 0 to iTotalCols
          array2D(iRowCount,iColCount) = arrData(iColCount)
   	 
	 Next

	 'ReDim Preserve array2D(iRowCount )
	
End Sub

Public Function appendTwoDimensionalArray (array2D, arrayObj )
  Dim iTotalCols, iTotalRows
  Dim array2DTemp

If (Ubound (array2D) = 0 ) Then
	If ( isEmpty (array2D(0,0)) And  isEmpty (array2D(0,1))  )Then
		ReDim   array2DAppendZeroIndex(0,Ubound(arrayObj) )

		For  iCol = 0 to Ubound(arrayObj)
            array2DAppendZeroIndex (0, iCol) =   arrayObj(iCol)                     
		Next
        appendTwoDimensionalArray = array2DAppendZeroIndex
		Exit Function

	End If
End If


  array2DTemp = array2D
  iTotalRows = Ubound(array2D) + 1
  iTotalCols = returnMaxNumber ( Ubound(array2D, 2),  Ubound(arrayObj) )

   ReDim   array2DAppended(iTotalRows,iTotalCols )
   ReDim Preserve array2DTemp (Ubound(array2D) , iTotalCols )

	For iRowCount = 0 to iTotalRows

		   For  iColCount = 0 to iTotalCols
			   If   iRowCount  <= Ubound(array2D) Then
				   array2DAppended (iRowCount, iColCount) = array2DTemp(iRowCount, iColCount)
			   else
			        If   iColCount > Ubound(arrayObj)Then
						Exit For
					 else
					    array2DAppended (iRowCount, iColCount) =   arrayObj(iColCount)

					End If
                   
				  
			   End If
                     
		   Next
    
         	 
	 Next

    appendTwoDimensionalArray = array2DAppended
	Exit Function
    	
End Function


Public Function mergeTwoDimensionalArrays (array2D1, array2D2 )
  Dim iTotalCols, iTotalRows
  Dim array2D1Temp, array2D2Temp
  Dim iIncrement
  iIncrement = 0

	If (Ubound (array2D1) = 0 ) Then
		If ( isEmpty (array2D1(0,0)) And  isEmpty (array2D1(0,1))  )Then
		
			mergeTwoDimensionalArrays = array2D2
			Exit Function
	
		End If
	End If


	array2D1Temp = array2D1
	array2D2Temp = array2D2

	iTotalRows = Ubound(array2D1) + Ubound(array2D2) +1	
    iTotalCols = returnMaxNumber ( Ubound(array2D1, 2),  Ubound(array2D2, 2) )

   ReDim   array2DMerged(iTotalRows,iTotalCols )
   ReDim Preserve array2D1Temp (Ubound(array2D1) , iTotalCols )
   ReDim Preserve array2D2Temp (Ubound(array2D2) , iTotalCols )

	For iRowCount = 0 to iTotalRows

		   For  iColCount = 0 to iTotalCols

			   If   iRowCount  <= Ubound(array2D1) Then
				   array2DMerged (iRowCount, iColCount) = array2D1Temp(iRowCount, iColCount)

			   else			   			   		       
					array2DMerged (iRowCount, iColCount) =   array2D2Temp(iIncrement, iColCount)		        
				  
			   End If
                     
		   Next

		   If  iRowCount  > Ubound(array2D1) Then
			     iIncrement = iIncrement + 1
		   End If

		  
         	 
	 Next

    mergeTwoDimensionalArrays = array2DMerged
	Exit Function
    	
End Function

Public Sub copyTwoDimensionalArray (array2DSource, array2DTarget )
  Dim iTotalCols, iTotalRows  
  
  iTotalRows = Ubound(array2DSource)
  iTotalCols =  Ubound(array2DSource,2)

	For iRowCount = 0 to iTotalRows

		   For  iColCount = 0 to iTotalCols
              array2DSource (iRowCount, iColCount) = array2DTarget(iRowCount, iColCount)	  
                     
		   Next    
         	 
	 Next

End Sub


Public Function removeFirstNElements1DArray (array1DSource, iCount )
  Dim  iTotalRows, iRowCount
  
  iTotalRows = Ubound(array1DSource,1)
  ReDim array1DRemoved(iTotalRows-iCount)
 
	For iRowCount = 0  to (iTotalRows-iCount)

		  array1DRemoved(iRowCount) = array1DSource(iRowCount + iCount)
		      	 
	Next

	removeFirstNElements1DArray = array1DRemoved

End Function

Public Function insertArrayIntoArray(arrSource, arrToBeInserted, iIndex )

   Dim  iTotalSource
   Dim  iTotalToBeInserted 
   Dim iInsert
   iInsert = 0
     
    iTotalSource = Ubound(arrSource,1)
	iTotalToBeInserted = Ubound(arrToBeInserted,1)

	Dim iTarget

	iTarget =  iTotalSource + iTotalToBeInserted +1

	ReDim arrTarget ( iTarget )

	For iCount = 0 to iTarget

		   If (iCount <  iIndex) Then
			   arrTarget (iCount) = arrSource (iCount)

			else

			    If  iInsert  <= iTotalToBeInserted Then

					arrTarget (iCount) = arrToBeInserted (iInsert)
					iInsert = iInsert  + 1
                else
                    arrTarget (iCount) = arrSource (iCount - ( iTotalToBeInserted+1) )
				End if

		   End If	
         	 
	 Next
	
	insertArrayIntoArray = arrTarget
	Exit Function
  
     
End Function

Public Function append1DArrayWith1DArray(arrSource, arrToAppend)

   If  isEmpty (arrSource) Then
		append1DArrayWith1DArray = arrToAppend
		Exit Function
    End if

	If isEmpty (arrToAppend)   Then
		append1DArrayWith1DArray = arrSource
		Exit Function
    End if

   	If  (isEmpty(arrSource(0)) And (Ubound(arrSource,1) = 0)) Then
		append1DArrayWith1DArray = arrToAppend
		Exit Function
    End if

	If (isEmpty(arrToAppend(0)) And (Ubound(arrToAppend,1) = 0)) Then
		append1DArrayWith1DArray = arrSource
		Exit Function
    End if

   Dim  iTotalSource
   Dim  iTotalToBeAppended, iAppendStart
   Dim arrTarget

   iAppendStart = 0

    iTotalSource = Ubound(arrSource,1)
	iTotalToBeAppended = Ubound(arrToAppend,1)

	Dim iTarget
	iTarget =  iTotalSource + iTotalToBeAppended +1

	arrTarget = arrSource
	ReDim Preserve arrTarget ( iTarget )

	For iCount = iTotalSource+1 to iTarget
        arrTarget (iCount) = arrToAppend (iAppendStart)
		iAppendStart  = iAppendStart+1   	 
	 Next
	
	append1DArrayWith1DArray = arrTarget
	Exit Function 
     
End Function

Public Function fetchFirstElementsOfAllRows (array2D)
  Dim iTotalCols, iTotalRows 
  ReDim arrString(0)
    
  iTotalRows = Ubound(array2D, 1)

	For iRowCount = 0 to iTotalRows
		   
	   Dim strTemp
	  strTemp = array2D (iRowCount, 0)			  
	  addItemToArray arrString, strTemp	
         	 
	 Next

      fetchFirstElementsOfAllRows =  arrString

	  Exit Function
End Function

Public Function returnMaxNumber (iNum1, iNum2)

    If  iNum1 > iNum2 Then
		returnMaxNumber = iNum1
		Exit Function
	else
       returnMaxNumber = iNum2
	   Exit Function
	End If
   
    	
End Function

Function checkIfNull(strValue)
   If UCase(strValue) ="BLANK" Or  strValue  = "" Or IsNull(strValue) Then
	   checkIfNull =true
	Else
		checkIfNull =false   
   End If
End Function


Public Function compareArray (arrStringSource, arrStringTarget)

   If  Not ( UBound(arrStringSource) = UBound(arrStringTarget) ) Then
      compareArray = False
   Else     
		For iCount = 0 to UBound(arrStringSource)
			If Not matchStr(Trim (arrStringTarget(iCount)), Trim (arrStringSource(iCount)))  Then
'				Reporter.ReportEvent micFail, "Compare Arrays", "The String" + arrStringSource(iCount) + " and the String " +arrStringTarget(iCount) +" not matched" 
				compareArray= False
				Exit Function
			End If
		Next

    compareArray = True

   End If    

End Function



Function matchStr(strString, strPattern)
	Dim objRegEx
	If (strPattern="") Then
		If  (strString="") Then
			matchStr=True
		 else
			matchStr=False
		End If
		
		Exit Function
	End If
	If strString=strPattern Then
			matchStr=True
			Exit Function
	End If

	' create the regular expression
	Set objRegEx = New RegExp
	' set the pattern
	objRegEx.Pattern = strPattern
	' ignore the casing
	objRegEx.IgnoreCase = True
	' perform the search
	matchStr = objRegEx.Test(strString)
	' destroy the object
	Set objRegEx = Nothing
End Function ' LocateText 

'Casting Functions

Function castStringToBoolean(strString)

	If Ucase(Trim (strString)) = "TRUE" Then
        castStringToBoolean = true
	Else
	   castStringToBoolean = false
	End If

End Function


Function castStringToListOfString(strString)
    strString = Trim(strString)
'	strString = Replace(strString, "(","")
'	strString = Replace(strString, ")","")
	If InStr(strString,"(")=1 AND InStrRev(strString,")")=Len(strString) Then
		strString=Mid(strString,2,Len(strString)-2)
	End If
	If  InStr(strString,"(")=1Then
		strString=Mid(strString,2,Len(strString)-1)
	End If
	Dim lstStr
	lstStr = split (strString,"|")
	If Not IsArray(lstStr) Then
		lstStr=Array(strString)
	End If
	For iCount = 0 to UBound(lstStr)
		If lstStr(iCount)="" And iCount=Ubound(lstStr) Then
			Exit for
		End If
		strTemp = checkNull(lstStr(iCount))
		lstStr(iCount) = strTemp
	Next

	castStringToListOfString = lstStr
	
End Function


Function castStringToListOfListOfString(strString)
    ReDim lstStrStr (0,1)
    strString = Trim(strString)

	Dim lstStr
	If InStrRev(strString,")|")=Len(strString)-1 Then
		strString=strString&"("
	End If

	If  InStr(strString,"(")=1 AND  InStrRev(strString,")")=Len(strString)Then
			strString=Mid(strString,2,Len(strString)-2)
    End If
	lstStr = split (strString,")|(")

    For iCount = 0 to UBound(lstStr)
		If lstStr(iCount)="" And iCount=Ubound(lstStr) Then
			Exit for
		End If
		Dim lstStrIndividual



		lstStrIndividual = 	castStringToListOfString(lstStr(iCount))
		Dim arrayTemp

		If  iCount =0 Then
			arrayTemp = appendTwoDimensionalArray (lstStrStr,lstStrIndividual )
		else
		   arrayTemp = appendTwoDimensionalArray (arrayTemp,lstStrIndividual )
		End If
		
	Next	

	castStringToListOfListOfString = arrayTemp
	
End Function


Public Sub setMachineEnviromentalVariable(strVariableType, strVariableName, strVariableValue)
     'Declare Variables
	  Dim WshShl, Shell

	'Set objects
	Set WshShl = CreateObject("WScript.Shell")
	Set Shell = WshShl.Environment(strVariableType)
    
	' create and set the custom variable
	Shell( strVariableName ) = strVariableValue

	'Cleanup Objects
	Set WshShl = Nothing
	Set Shell = Nothing
	
	Exit Sub
	
End Sub


Public Sub clickAjaxComboImages(objParent, iIndex)
   Redim arrayComboImageElements(0)

	Set dpAllTxt = Description.Create 
	dpAllTxt("micClass").value = "WebElement"  
	dpAllTxt("class").value = "jsx30select_display" 
	dpAllTxt("visible").value = true
	 
	Dim childs
	Set childs =objParent.ChildObjects(dpAllTxt)
	iFound =0
	For iCount = 0 to (childs.Count-1)
	
		Dim childElement
		Set childElement = childs(iCount)
		If   (childElement.GetROProperty("class") = "jsx30select_display") Then
			
				addObjectToArray arrayComboImageElements, childElement
			
		End If
        
	Next

'	For iCount =0 to UbOund(arrayComboImageElements)
'		arrayComboImageElements(iCount).Click()
'		wait(5)
'	Next
    arrayComboImageElements(iIndex).Click()
	Wait(5)

End Sub



Public Function addObjectToArray(arrString, strSearchItem)

   If  (UBound(arrString) = 0 AND IsEmpty (arrString(0)) ) Then
    
       Set arrString(0) = strSearchItem

   else

      ReDim Preserve arrString(UBound(arrString)+1) 
       Set arrString(UBound(arrString)) = strSearchItem

   End If

   addObjectToArray = arrString
     
End Function



Public Function typeSpecialText( obj, strValue)

	obj.Click()
	
	Set dr = CreateObject( "Mercury.DeviceReplay" )
        dr.SendString "S6506237B"
    	obj.Set(strValue)	


End Function

Public Function captureFNANumber(Objt)

	Dim strFNANumber,arrVals
	strVal =  Objt.GetROProperty("innertext")
	arrVals = Split(strVal,":")
	strFNANumber = arrVals(1)
	captureFNANumber = strFNANumber

End Function


Public sub selectRadioGroup(objRadioGroup, strOption, arrayOptions)
	Dim iIndex
	iIndex= IndexOf(arrayOptions, strOption)
	'objRadioGroup.Select "#"+cstr(iIndex)
	objRadioGroup(iIndex).Click
End Sub

'Public sub selectRadioGroup(objRadioGroup, strOption, arrayOptions)
'	For i = 0 To objRadioGroup.Count-1
'		strRadioValue=objRadioGroup(i).getroproperty("innertext")
'		If Trim(strRadioValue)=Trim(strOption) Then
'			strRadioValue.Click
'			Exit Sub
'		End If		
'	Next
'End Sub

Public sub selectRadioGroupElement(objRadioGroup, strOption, arrayOptions)
	Dim iIndex
	iIndex= IndexOf(arrayOptions, strOption)

	objRadioGroup.Click "#"+cstr(iIndex)
End Sub


Public Function ArrayToString(arrData,delimiter)
	If IsArray(arrData) Then
		For i = 0 To UBound(arrData)
			If i = 0 Then
				ArrayToString = arrData(i)
			Else
				ArrayToString = ArrayToString+delimiter+arrData(i)
			End If		
		Next
	Else
		ArrayToString = arrData		
	End If
End Function


Public Function Get2DArrayFromTable(objWebTable)

	Dim iPageCount, iTableRowCount, iTableColCount, arr2DArrayTable, arrRowData, arrRowDataTemp
	
	Dim arrTemp()

'	arr2DArrayTable = Array()
'	arrRowData = Array()

	If objWebTable.Exist(0) Then
	
		iTableRowCount = objWebTable.RowCount

		If iTableRowCount > 0 Then
			iTableColCount = objWebTable.ColumnCount(1)
		Else
			iTableColCount = 0
		End If

		If (iTableRowCount) > 0 And (iTableColCount > 0) Then
			For i = 1 to iTableRowCount
					Redim arrRowData(iTableColCount - 1)
					For j = 1 to iTableColCount
						ReDim Preserve arrRowData(j - 1)
						arrRowData(j - 1) = objWebTable.GetCellData(i,j)
					Next
					If i = 1 Then	
							fill2DArrayRowValue  arrTemp , arrRowData, 0, 0
							arr2DArrayTable = arrTemp
					Else
							arr2DArrayTable = appendTwoDimensionalArray(arr2DArrayTable,arrRowData)
					End If
			Next				
		Else
			LogMessage "INFO", "Table does not have data","Table doesnt not have any data", True
			Get2DArrayFromTable = arrTemp  'Blank Array ' No Rows
			Exit Function
		End If

	Else
		LogMessage "RSLT", "Verification","Table object is not available. Error description :- " &  err.description, False
		Exit Function
	End If
	Get2DArrayFromTable = arr2DArrayTable
End Function


Public Function FetchRowArrayFrom2DArray(arr2D, iRow)
		Dim iCol, arrTemp
		iCol = UBound(arr2D, 2)

		Redim arrTemp(iCol)

		For i = 0 To iCol
				arrTemp(i) = arr2D(iRow,i)
		Next
	FetchRowArrayFrom2DArray = arrTemp
End Function
Public Function ArrayCompare (arrStringSource, arrStringTarget)

   If  Not ( UBound(arrStringSource) = UBound(arrStringTarget) ) Then

      ArrayCompare = False

   else     
		For iCount = 0 to UBound(arrStringSource)

		If Not IsNull(arrStringSource(iCount)) Then  'If you want to pass value ignore the array value, pass BLANK or "" . To compare the blank values pass NULL
		'strcomp
			If StrComp(Trim(arrStringTarget(iCount)), Trim(arrStringSource(iCount)), 1)  <>  0 Then
					ArrayCompare= False
					Exit Function
			End If
		End If

		Next

    ArrayCompare = True

   End If    
End Function


Public Function ArrayCompare_Regx (arrStringSource, arrStringTarget)

   If  Not ( UBound(arrStringSource) = UBound(arrStringTarget) ) Then

      ArrayCompare_Regx = False

   else     
		For iCount = 0 to UBound(arrStringSource)

		If Not IsNull(arrStringSource(iCount)) Then  'If you want to pass value ignore the array value, pass BLANK or "" . To compare the blank values pass NULL
		'strcomp
			If Not Matchstr(Trim(arrStringTarget(iCount)), Trim(arrStringSource(iCount))) Then
					ArrayCompare_Regx= False
					Exit Function
			End If
		End If

		Next

    ArrayCompare_Regx = True

   End If    
End Function


Function GetArrayDimension(arr)
    Dim dimensions : dimensions = 0
    On Error Resume Next
    Do While Err.number = 0
        dimensions = dimensions + 1
        UBound arr, dimensions
    Loop
    On Error Goto 0
    GetArrayDimension = dimensions - 1
End Function

'Returns the match count
Function RegExpMatchCount(strPattern, SearchString)
   Dim regEx, Match, Matches   ' Create variable.
   Set regEx = New RegExp   ' Create a regular expression.
   regEx.Pattern = strPattern   ' Set pattern.
   regEx.IgnoreCase = True   ' Set case insensitivity.
   regEx.Global = True   ' Set global applicability.
   Set Matches = regEx.Execute(SearchString)   ' Execute search.
   
'   For Each Match in Matches   ' Iterate Matches collection.
'      RetStr = RetStr & "Match found at position "
'      RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
'      RetStr = RetStr & Match.Value & "'." & vbCRLF
'	   RegExpTest = Match.FirstIndex
'   Next
	RegExpMatchCount = Matches.count
End Function


Public Function AppendToArray(arrToAppend, arrData)
		'Get Dimension of Array
		Dim iArrayDimension
		
		iarrToAppendDimension =  GetArrayDimension(arrToAppend)
		iarrDataDimension = GetArrayDimension(arrData)
		
		If iarrToAppendDimension = 0 Then 'Fresh array. Decide the dimension based on the arrData
			If iarrDataDimension = 1 Then
					iarrDataRow = Ubound(arrData)
					ReDim arrToAppend(iarrDataRow)
					For iRow = 0 to iarrDataRow
						arrToAppend(iRow) = arrData(iRow)
					Next
			elseif iarrDataDimension = 2 Then
					iarrDataRow = Ubound(arrData, 1)
					iarrDataColumn = Ubound(arrData, 2)
					ReDim arrToAppend(iarrDataRow, iarrDataColumn)
					'Copy the data
					For iRow = 0 to iarrDataRow
						For iCol = 0 to iarrDataColumn
							arrToAppend( iRow, iCol ) = arrData(iRow, iCol)
						Next
					Next
			End If
		Else 'Existing array. Change the size
			If iarrToAppendDimension <>  iarrDataDimension Then
				 err.number = "9001"
				 err.description = "Dimension of the array to append and the data array are different"
			Else
					Select Case iarrToAppendDimension
						Case 1:
								'Write code to append the array
								  iarrToAppendRow = Ubound(arrToAppend, 1)
								  iarrDataRow = Ubound(arrData, 1)
								  iNewRow = iarrToAppendRow + iarrDataRow + 1
								  ReDim Preserve arrToAppend(iNewRow)
								For iRow = 0 to iarrDataRow
									arrToAppend(iarrToAppendRow + iRow + 1) = arrData(iRow)
								Next			
						Case 2:
								iarrToAppendRow = Ubound(arrToAppend, 1)
								iarrToAppendColumn = Ubound(arrToAppend, 2)
								iarrDataRow = Ubound(arrData, 1)
								iarrDataColumn = Ubound(arrData, 2)

								'Check both the column size is same
								If Ubound(arrToAppend, 2) = Ubound(arrData, 2) Then
									'Code to redim the size and append data
									iNewRow = iarrToAppendRow + iarrDataRow + 1
									Dim arrOld
									arrOld = arrToAppend
									ReDim arrToAppend(iNewRow , iarrToAppendColumn)
									'Fill old
									For iOldRow = 0 to Ubound(arrOld, 1)
										For iOldCol = 0 to Ubound(arrOld, 2)
												arrToAppend( iOldRow, iOldCol ) = arrOld(iOldRow, iOldCol)
										Next
									Next
									'Fill New
									For iRow = 0 to iarrDataRow
										For iCol = 0 to iarrDataColumn
											arrToAppend( iarrToAppendRow+iRow+1, iCol ) = arrData(iRow, iCol)
										Next
									Next
								End If			
					End Select
			End If
		End If
End Function
