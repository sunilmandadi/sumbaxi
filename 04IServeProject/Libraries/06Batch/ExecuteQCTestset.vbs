dim arg
arg=WScript.Arguments.Count
'msgbox arg
dim strTestFolder,strTestSetName,strHostName,strRunON
strTestFolder=WScript.Arguments(0)
strTestSetName=WScript.Arguments(1)
strHostName=WScript.Arguments(2)
strRunON=WScript.Arguments(2)
'unTestSet "iCall\Release1\SIT\UI Automation","Sanity_UAT_Drop1","","RUN_LOCAL"
RunTestSet strTestFolder,strTestSetName,strHostName,strRunON
Public Sub RunTestSet(tsFolderName, tSetName,HostName, runWhere )
' This example show how to run a test set in three different ways:
' * Run all tests on the local machine (where this code runs).
' * Run the tests on a specified remote machine.
' * Run the tests on the hosts as planned in the test set.
    Dim TSetFact, tsList
    Dim theTestSet
    Dim tsTreeMgr
    Dim tsFolder
    Dim Scheduler
    Dim execStatus
                Set fso = CreateObject ("Scripting.FileSystemObject")
Set stdout = fso.GetStandardStream (1)
'Set stderr = fso.GetStandardStream (2)
stdout.WriteLine "This will go to standard output."
stdout.WriteLine "Folder:"&tsFolderName
'stderr.WriteLine "This will go to error output."
    On Error Resume Next
    errmsg = "RunTestSet"
' Get the test set tree manager from the test set factory.
    'tdc is the global TDConnection object.
                set tdc= createobject("TDApiOle80.TDConnection")
                tdc.InitConnectionEx "http://10.91.65.132:8080/qcbin"
                ' Check status.
stdout.WriteLine "User LoggedIn :"& tdc.LoggedIn 'False
stdout.WriteLine "UserConnected: "& tdc.Connected 'True
stdout.WriteLine "QC Server:"& tdc.ServerName 'http://<qcServer>/qcbin/wcomsrv.dll
tdc.Login "subashuttam", "dbs123"
' Check status.
stdout.WriteLine "User LoggedIn :"& tdc.LoggedIn 'True
'stdout.WriteLine "Project Name :"& tdc.ProjectName 'Empty String
stdout.WriteLine "Projected Connected? :"& tdc.ProjectConnected 'False
' Connect to the project and user.
                tdc.Connect "CBG", "iCall"
stdout.WriteLine "Project Name :"& tdc.ProjectName  'qcProject
stdout.WriteLine "Project Connected ?:"& tdc.ProjectConnected 'True
' Exit status.
    Set TSetFact = tdc.TestSetFactory
   Set tsTreeMgr = tdc.TestSetTreeManager
' Get the test set folder passed as an argument to the example code.
    Dim nPath
    nPath = "Root\" & Trim(tsFolderName)
    On Error Resume Next
    Set tsFolder = tsTreeMgr.NodeByPath(nPath)
stdout.WriteLine "Test Set Path ?:" &tsFolder
    If tsFolder Is Nothing Then
        err.Raise vbObjectError + 1, "RunTestSet", "Could not find folder " & nPath
                                stdout.WriteLine  "Could not find folder " & nPath
        'GoTo RunTestSetErr
    End If
    On Error Resume Next
' Search for the test set passed as an argument to the example code.
    Set tsList = tsFolder.FindTestSets(tSetName)
                stdout.WriteLine "Test sets found:" &tsList.Count
    If tsList.Count > 1 Then
        stdout.WriteLine "FindTestSets found more than one test set: refine search "&tsList.Count
    ElseIf tsList.Count < 1 Then
        stdout.WriteLine "FindTestSets: test set not found"
    End If
    Set theTestSet = tsList.Item(1)
    'Debug.Print theTestSet.ID
' Start the scheduler on the local machine.
    Set Scheduler = theTestSet.StartExecution("")
'Set up for the run depending on where the test instances
' are to execute.
    Select Case runWhere
        Case "RUN_LOCAL"
        ' Run all tests on the local machine.
                               'msgbox "Local"
            Scheduler.RunAllLocally = True
        Case "RUN_REMOTE"
        ' Run tests on a specified remote machine.
            Scheduler.TdHostName = HostName
            ' RunAllLocally must not be set for
            ' remote invocation of tests.
            ' Do not do this:
            ' Scheduler.RunAllLocally = False
        Case "RUN_PLANNED_HOST"
        ' Run on the hosts as planned in the test set.
            Dim TSTestFact, testList
            Dim tsFilter
            Dim TSTst
        ' Get the test instances from the test set.
            Set TSTestFact = theTestSet.TSTestFactory
            Set tsFilter = TSTestFact.Filter
            tsFilter.Filter("TC_CYCLE_ID") = theTestSet.ID
            Set testList = TSTestFact.NewList(tsFilter.Text)
            Debug.Print "Test instances and planned hosts:"
        'For each test instance, set the host to run depending
        ' on the planning in the test set.
            For Each TSTst In testList
                Print "Name: " & TSTst.Name & " ID: " & TSTst.ID & " Planned Host: " & TSTst.HostName
                Scheduler.RunOnHost(TSTst.ID) = TSTst.HostName
            Next
            Scheduler.RunAllLocally = False
    End Select
' Run the tests.
'               WScript.Sleep(1000)
    Scheduler.Run
' Get the execution status object.
    Set execStatus = Scheduler.ExecutionStatus
                Debug.Print "Execution Status:"&execStatus
                stdout.WriteLine "Execution Status:"&execStatus
' Track the events and statuses.
    Dim RunFinished , iter , i
    Dim ExecEventInfoObj , EventsList
    Dim TestExecStatusObj
                iter=0
    While (RunFinished = False) And (iter < 100)
                                iter = iter + 1
        execStatus.RefreshExecStatusInfo "all", True
        RunFinished = execStatus.Finished
        Set EventsList = execStatus.EventsList
        For Each ExecEventInfoObj In EventsList
            stdout.WriteLine  "Event: " & ExecEventInfoObj.EventDate & " " & _
                    ExecEventInfoObj.EventTime & " " & _
                    "Event Type: " & ExecEventInfoObj.EventType & " [Event types: " & _
                    "1-fail, 2-finished, 3-env fail, 4-timeout, 5-manual]"
        Next
        'msgbox  execStatus.Count & " exec status"
        For i = 1 To execStatus.Count
            Set TestExecStatusObj = execStatus.Item(i)
                                                stdout.WriteLine  "Iteration " & iter & " Status: " & _
                       " Test " & TestExecStatusObj.TestID & _
                        " ,Test instance " & TestExecStatusObj.TestInstance & _
                        " ,order " & TestExecStatusObj.TSTestID & " " & _
                        TestExecStatusObj.Message & ", status=" & _
                        TestExecStatusObj.Status
        Next
                                Set TestExecStatusObj = execStatus.Item(i)
                                'stdout.WriteLine  "Iteration " & iter & " Testcase Name: " & TestExecStatusObj.Name &_
         '               "Test instance " & TestExecStatusObj.TestInstance & _
          '              "Execution " &TestExecStatusObj.Message & ", status=" & _
           '             TestExecStatusObj.Status
        'Sleep() has to be declared before it can be used.
        'This is the module level declaration of Sleep():
        'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
        'Sleep (5000)
                                WScript.Sleep (60000)   
    Wend 'Loop While execStatus.Finished = False
    stdout.WriteLine "Scheduler finished around " & CStr(Now)
                If tdc.Connected Then
       tdc.Disconnect
    End If
                If tdc.LoggedIn Then
        tdc.Logout
    End If
                tdc.ReleaseConnection
                stdout.WriteLine "At the end of Test Set, Connected ?" &              tdc.Connected 'False
                Set tdc = nothing
Exit Sub
'RunTestSetErr:
'   ErrHandler err, err.Description, errmsg, NON_FATAL_ERROR
End Sub