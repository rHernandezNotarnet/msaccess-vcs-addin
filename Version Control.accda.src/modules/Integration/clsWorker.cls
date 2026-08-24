VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : This module has two primary purposes.
'           :
'           : 1. To run commands in the `current database` that would otherwise risk
'           : running in the add-in database. (Deleting a database object, for example,
'           : will default to the CodeDB before CurrentDB.)
'           :
'           : 2. Opens up the opportunity for parallel execution when VBScript can run
'           : operations asynchrously. (For example, scanning through a large number of
'           : source files can be slow, but this gives us the ability to scan multiple
'           : folders at the same time which should improve performance on SSD drives.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Integration")


' Worker Job queue of jobs in progress
Public Queue As Dictionary

' Last result returned by a worker callback
Private m_varLastResult As Variant

' Generic placeholder objects so we can compile the VBScript in the VBA IDE
Private WScript As Object


'---------------------------------------------------------------------------------------
' Procedure : Run_UninstallAddin
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Remove the actual add-in file and lock file during uninstall.
'---------------------------------------------------------------------------------------
'
Public Function Run_UninstallAddin()
    CallWorker PROJECT_NAME, 10, "Uninstall", """" & GetInstalledAddInFileName & """"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Run_BuildAndInstall
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Rebuild the VCS add-in from source files and install the result.
'           : Runs disconnected: quits the current Access instance, opens a new one
'           : to perform the build, then launches the install via /cmd INSTALL.
'---------------------------------------------------------------------------------------
'
Public Function Run_BuildAndInstall(strSourcePath As String)
    CallWorker PROJECT_NAME, 300, "BuildAndInstall", _
        """" & strSourcePath & """"
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDatabaseAccessible
' Author    : Adam Waller
' Date      : 4/2/2026
' Purpose   : Check if the database file is accessible to external clients by running
'           : an out-of-process DAO connection test via the worker script. This detects
'           : the JET/ACE engine-level lock state (error 3734) that an in-process check
'           : cannot detect.
'           : Pass the path of the database to test. Falling back to the caller's
'           : CurrentProject tests the add-in file itself whenever the operation is
'           : driven from the add-in, which is not what the caller is asking about.
'---------------------------------------------------------------------------------------
'
Public Function IsDatabaseAccessible(Optional strDatabasePath As String) As Boolean

    Dim strPath As String

    strPath = strDatabasePath
    If Len(strPath) = 0 Then strPath = CurrentProject.FullName
    m_varLastResult = Empty
    CallWorker "DbAccessCheck", 10, "CheckDatabaseAccessible", """" & strPath & """"
    Me.WaitForQueue 10, "CheckDatabaseAccessible", 0.1

    ' Three outcomes, not two. An empty result means the worker never reported anything
    ' at all, which is not the same as an answer of "not accessible" even though we take
    ' the same conservative action. Without this, a worker that died is indistinguishable
    ' in the log from a database that really was locked.
    If IsEmpty(m_varLastResult) Then
        Log.Error eelWarning, _
            T("No worker result for {0}; assuming it is not accessible.", var0:=strPath), _
            ModuleName(Me) & ".IsDatabaseAccessible"
        IsDatabaseAccessible = False
    Else
        IsDatabaseAccessible = CBool(m_varLastResult)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteDatabaseObject
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Delete a database object from the current database. (The reason we are
'           : not just using the native command is because when it is initiated from
'           : the add-in, it will delete any object with the same name in the add-in
'           : before deleting the one in the current database. This will cause the
'           : add-in file to be currupted, and the operation will fail.)
'---------------------------------------------------------------------------------------
'
Public Function WIP_DeleteDatabaseObject(intType As AcObjectType, strName As String, Optional blnWaitForQueue As Boolean = True)

    ' Don't use the following method from the add-in! Instead, we need to call
    ' it from an external process so it deletes from the current database first.
    'DoCmd.DeleteObject acModule, "basModule"
    CallWorker strName, 5, "DeleteDatabaseObject", intType, strName
    If blnWaitForQueue Then Me.WaitForQueue 5, "DeleteDatabaseObject", 0.25

End Function


'---------------------------------------------------------------------------------------
' Procedure : RenameDatabaseObject
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Rename the specified database object
'---------------------------------------------------------------------------------------
'
Public Function WIP_RenameDatabaseObject(strNewName As String, intType As AcObjectType, _
    strOldName As String, Optional blnWaitForQueue As Boolean = True)
    CallWorker strNewName, 3, "RenameDatabaseObject", strNewName, intType, strOldName
    If blnWaitForQueue Then Me.WaitForQueue 3, "RenameDatabaseObject", 0.1
End Function


'---------------------------------------------------------------------------------------
' Procedure : WaitForQueue
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Wait for the job queue to finish processing, checking at each specified
'           : interval. (Don't make the poll interval too high, or the workers may not
'           : be able to complete their callbacks.)
'---------------------------------------------------------------------------------------
'
Public Sub WaitForQueue(dblTimeout As Double, Optional strAction As String = "All", Optional sngPollSeconds As Single = 0.5)

    Dim cJob As clsJob
    Dim varKey As Variant
    Dim blnFound As Boolean
    Dim dblStart As Double
    Dim dblPollStart As Double
    Dim colRemove As Collection
    Dim varJob As Variant

    ' No need to wait if there is nothing in the queue
    If Me.Queue.Count = 0 Then Exit Sub

    ' Get start time so we can monitor the timeout
    dblStart = Perf.MicroTimer
    Set colRemove = New Collection

    ' Monitor performance of waiting for the job queue to clear
    Perf.OperationStart "Wait for Job Queue"

    ' Loop while periodically checking the queue
    Do
        ' Pause for the poll interval
        dblPollStart = Perf.MicroTimer
        Do While dblPollStart + sngPollSeconds > Perf.MicroTimer
            ' A tight DoEvents loop is required here. The worker
            ' VBScript makes many individual COM calls back into
            ' Access (property access, VBE iteration, etc.), each
            ' needing message pump dispatch. Sleep would add ~100ms
            ' latency per round-trip, compounding to seconds of delay.
            DoEvents
        Loop
        DoEvents

        ' Pick up anything delivered through the file channel. A job that does not need
        ' the Access instance answers by writing a file, so nothing arrives on its own
        ' and the wait loop has to come looking for it.
        PollJobResults

        ' We can exit early if there is nothing in the queue
        If Me.Queue.Count = 0 Then Exit Do

        ' Reset flag for this iteration
        blnFound = False

        ' Loop through queue of jobs
        For Each varKey In Me.Queue.Keys
            Set cJob = Me.Queue(varKey)

            ' Check for job timeout (all jobs)
            If Perf.MicroTimer > cJob.Timeout Then
                ' Add to list of items to remove
                colRemove.Add CStr(varKey)
            Else
                ' Look for matching actions
                If cJob.Action = strAction Then
                    blnFound = True
                    Exit For
                End If
            End If
        Next varKey
        Set cJob = Nothing

        ' Log and remove any individual jobs that have timed out
        For Each varJob In colRemove
            Set cJob = Me.Queue(varJob)
            With cJob
                Log.Error eelError, "Job " & .KeyHash & " (" & .Action & ") timed out after " & _
                    Round(Perf.MicroTimer - .Start, 0) & " seconds.", ModuleName(Me) & ".WaitForQueue"
            End With
            Set cJob = Nothing
            If Me.Queue.Exists(varJob) Then Me.Queue.Remove varJob
        Next varJob

        ' If nothing found of a specified action, we can exit the loop
        If (strAction <> "All") And (Not blnFound) Then Exit Do

        ' Check the overall timeout before we start another iteration
        If dblStart + dblTimeout < Perf.MicroTimer Then
            Log.Error eelError, "Timed out waiting for job queue", ModuleName(Me) & ".WaitForQueue"
            Log.Add " (" & Me.Queue.Count & " jobs still in the queue.)", False
            Exit Do
        End If
    Loop
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Verify that the work script is installed and unchanged.
'---------------------------------------------------------------------------------------
'
Private Function VerifyWorker() As Boolean

    Static blnVerified As Boolean

    Dim strContent As String
    Dim strFilePath As String

    ' Only need to verify this once per session.
    If blnVerified Then
        VerifyWorker = True
        Exit Function
    End If

    ' Compare hash of file with hash of content
    strContent = GetWorkerScriptContent
    strFilePath = WorkerFilePath
    If FSO.FileExists(strFilePath) Then
        If GetFileHash(WorkerFilePath) = GetStringHash(strContent, True) Then
            ' Existing file matches template.
            blnVerified = True
        End If
    End If

    ' Update the file if needed.
    If Not blnVerified Then
        WriteFile strContent, strFilePath, "Windows-1252"
        blnVerified = True
    End If
    VerifyWorker = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetWorkerScriptContent
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Return the content of the script, ready to save to a `Worker.vbs` file.
'---------------------------------------------------------------------------------------
'
Private Function GetWorkerScriptContent() As String

    Dim strContent As String
    Dim varLines As Variant
    Dim strLine As String
    Dim blnSave As Boolean
    Dim lngLine As Long

    ' Get VBA code from this class
    strContent = GetCodeVBProject.VBComponents(TypeName(Me)).CodeModule.Lines(1, 99999)

    ' Split into lines
    varLines = Split(strContent, vbCrLf)

    ' Look for the marker where the placeholder starts, and keep everything after that.
    With New clsConcat
        For lngLine = 0 To UBound(varLines) - 1
            strLine = varLines(lngLine)
            If blnSave Then
                .Add strLine, vbCrLf
            ElseIf strLine = "' *** BEGIN WORKER SCRIPT ***" Then
                blnSave = True
                ' Add line to top of VBScript to launch the main sub
                .Add "Main ' Launch main subroutine", vbCrLf, vbCrLf
            End If
        Next lngLine
        ' Return script content
        GetWorkerScriptContent = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : WorkerFilePath
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : File path to the installed worker script.
'---------------------------------------------------------------------------------------
'
Private Function WorkerFilePath() As String
    WorkerFilePath = FSO.GetParentFolderName(modInstall.GetAddInFileName) _
        & PathSep & "Worker.vbs"
End Function


'---------------------------------------------------------------------------------------
' Procedure : CallWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Call the worker script to run a specified operation.
'---------------------------------------------------------------------------------------
'
Private Function CallWorker(strKey As String, dblTimeoutSeconds As Double, _
    strAction As String, ParamArray varParams()) As Long

    Dim intParam As Integer
    Dim strHash As String
    Dim strJobKey As String
    Dim lngHandle As Long
    Dim cJob As clsJob

    ' Make sure this key/action is not already running. The queue key carries a suffix
    ' unique to each launch, so the check is on the deterministic part of it.
    strHash = GetStringHash(strKey & strAction)
    If JobIsQueued(strHash) Then
        Log.Error eelError, "Duplicate job in queue: " & strAction & "." & strKey, _
            ModuleName(Me) & ".CallWorker"
        Exit Function
    End If

    ' Make the key unique to this launch. Without a suffix, the key of a given
    ' key/action pair is identical in every session and forever, so a worker answering
    ' after its caller gave up would satisfy a later job with a stale value, and a
    ' result file left behind by an earlier run would be read as this answer. The
    ' computer name is in the salt because the install folder, and with it the result
    ' folder, can be shared, while the timer and counter behind the suffix are per
    ' process.
    strJobKey = strHash & "-" & UniqueHashSuffix(strKey & strAction & Environ$("COMPUTERNAME"))

    ' Verify that the work script is installed and current
    VerifyWorker

    ' Build command with parameters
    With New clsConcat

        ' Build base command to launch worker
        .Add "wscript """, WorkerFilePath, """"

        ' Add database file as first parameter
        .Add " """, CurrentProject.FullName, """"

        ' Add unique key as second parameter
        .Add " ", strJobKey

        ' Add specified action
        .Add " ", strAction

        ' Add any parameters
        For intParam = 0 To UBound(varParams)
            .Add " ", CStr(varParams(intParam))
        Next intParam

        ' Add to job queue (before launching, just in case it runs very fast)
        Set cJob = New clsJob
        With cJob
            .KeyHash = strJobKey
            .Action = strAction
            .Start = Perf.MicroTimer
            .Timeout = .Start + dblTimeoutSeconds
            ' Where this job's result arrives if the worker reports by file. Assigned
            ' for every job: which channel is used is the worker's call, and polling
            ' for a file that never appears costs one FileExists.
            .ResultFile = JobResultFolder & PathSep & strJobKey & ".result"
        End With
        Me.Queue.Add strJobKey, cJob

        ' Use the Shell command to launch the worker as an external asynchrous
        ' process, and return the process ID to the caller.
        lngHandle = Shell(.GetStr, vbHide)

    End With

    ' Return process handle
    CallWorker = lngHandle

End Function


'---------------------------------------------------------------------------------------
' Procedure : JobIsQueued
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/22/2026
' Purpose   : Return true if a job for this key hash is already in the queue. The queue
'           : key is the hash plus a per-launch suffix, so this matches on the hash
'           : alone and keeps the duplicate check independent of the suffix.
'---------------------------------------------------------------------------------------
'
Private Function JobIsQueued(strKeyHash As String) As Boolean

    Dim varKey As Variant

    ' Both halves of a queue key are lowercase hexadecimal (modHash.GetHash and
    ' modHash.UniqueHashSuffix), so neither half can contain the separator. If the
    ' hash alphabet ever changes, this guard stops being sound.
    For Each varKey In Me.Queue.Keys
        If Left(CStr(varKey), Len(strKeyHash) + 1) = strKeyHash & "-" Then
            JobIsQueued = True
            Exit Function
        End If
    Next varKey

End Function


'---------------------------------------------------------------------------------------
' Procedure : JobResultFolder
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/22/2026
' Purpose   : Folder where workers drop their result files, next to the worker script.
'           : (WriteJobResult builds the same path from WScript.ScriptFullName, so the
'           : two folder names have to stay in step.)
'           : Swept once per operation of files left behind by a job whose caller had
'           : already given up. The sweep is by age rather than by emptying the folder,
'           : because another Access instance may have a job in flight in here.
'---------------------------------------------------------------------------------------
'
Private Function JobResultFolder() As String

    ' Well past the longest job timeout (300 seconds for BuildAndInstall), so a file
    ' this old cannot belong to a job anybody is still waiting for.
    Const StaleJobHours As Double = 1

    Static blnReady As Boolean

    Dim oFS As Object
    Dim strFolder As String
    Dim colStale As Collection
    Dim varFile As Variant

    ' Dereference the FSO getter before opening a local handler: the getter contains
    ' LogUnhandledErrors, so with an error pending it would be reported from in here.
    Set oFS = FSO
    strFolder = oFS.GetParentFolderName(modInstall.GetAddInFileName) & PathSep & "WorkerJobs"
    JobResultFolder = strFolder
    If blnReady Then Exit Function
    blnReady = True

    LogUnhandledErrors
    On Error Resume Next

    ' VerifyPath tells a folder from a file name by the trailing separator, so without
    ' one it would verify the parent instead and silently do nothing here.
    VerifyPath strFolder & PathSep

    ' Collect first and delete afterwards: the Files collection is live, and deleting
    ' while enumerating it can skip entries.
    Set colStale = New Collection
    If oFS.FolderExists(strFolder) Then
        For Each varFile In oFS.GetFolder(strFolder).Files
            If varFile.DateLastModified < Now - (StaleJobHours / 24) Then
                colStale.Add varFile.Path
            End If
        Next varFile
    End If
    For Each varFile In colStale
        DeleteFile CStr(varFile)
    Next varFile

    CatchAny eelWarning, "Error preparing the worker result folder: " & strFolder, _
        ModuleName(Me) & ".JobResultFolder"

End Function


'---------------------------------------------------------------------------------------
' Procedure : PollJobResults
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/22/2026
' Purpose   : Pick up results delivered through the file channel. An action that does
'           : not need the Access instance to do its work reports by writing a file
'           : rather than by calling back over COM, so nothing arrives on its own and
'           : the wait loop has to come looking. The value is handed to ReturnWorker so
'           : that job logging and queue removal stay in one place.
'           : Runs under its own handler on purpose: modBuild.Build wraps everything in
'           : On Error Resume Next, so an error escaping from here would be swallowed
'           : there and the build would carry on as if the check had answered.
'---------------------------------------------------------------------------------------
'
Private Sub PollJobResults()

    Dim oFS As Object
    Dim varKeys As Variant
    Dim varKey As Variant
    Dim cJob As clsJob
    Dim strFile As String
    Dim strContent As String
    Dim strValue As String

    ' Dereference the FSO getter and snapshot the keys before opening a local handler:
    ' the getter contains LogUnhandledErrors, so with an error pending it would be
    ' reported from in here. (modInstall carries the same note for the same reason.)
    Set oFS = FSO
    varKeys = Me.Queue.Keys

    LogUnhandledErrors
    On Error Resume Next

    For Each varKey In varKeys
        If Me.Queue.Exists(varKey) Then
            Set cJob = Me.Queue(varKey)
            strFile = cJob.ResultFile
            Set cJob = Nothing
            If Len(strFile) > 0 Then
                If oFS.FileExists(strFile) Then

                    ' Split on an empty string returns an array with no elements, so
                    ' the first element cannot be read without checking the length.
                    strContent = ReadFile(strFile)
                    strValue = vbNullString
                    If Len(strContent) > 0 Then
                        strValue = Trim(Split(strContent, vbCrLf)(0))
                    End If

                    Select Case strValue
                        Case "1", "0"
                            ' Hand the value over before deleting. DeleteFile has no
                            ' error handling of its own, so a sharing violation must
                            ' not be able to lose a result that did arrive; the age
                            ' sweep collects whatever is left behind.
                            Me.ReturnWorker CStr(varKey), (strValue = "1")
                            DeleteFile strFile
                        Case Else
                            ' Not content we wrote. Retire the job so the caller does
                            ' not also pay the timeout, and keep the file as evidence
                            ' instead of deleting it.
                            Log.Error eelWarning, _
                                T("Unexpected worker result in {0}", var0:=strFile), _
                                ModuleName(Me) & ".PollJobResults"
                            Me.ReturnWorker CStr(varKey), Empty
                            If oFS.FileExists(strFile & ".bad") Then
                                oFS.DeleteFile strFile & ".bad"
                            End If
                            oFS.MoveFile strFile, strFile & ".bad"
                    End Select
                End If
            End If
        End If
    Next varKey

    CatchAny eelWarning, "Error reading worker result files", _
        ModuleName(Me) & ".PollJobResults"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Initialize the worker queue.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set Me.Queue = New Dictionary
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReturnWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Complete any post-call processes and remove from job queue.
'---------------------------------------------------------------------------------------
'
Public Function ReturnWorker(strKey As String, varParams As Variant)

    Dim cJob As clsJob

    ' Guard clause in case the code is stopped while workers are running.
    If Me.Queue Is Nothing Then Exit Function

    ' Make sure this job exists in the queue
    If Not Me.Queue.Exists(strKey) Then
        Log.Error eelError, "Returned worker not found in job queue: " & strKey, _
            ModuleName(Me) & ".ReturnWorker"
        Exit Function
    Else
        Set cJob = Me.Queue(strKey)
    End If

    ' Store the return value so the caller can read it after WaitForQueue
    m_varLastResult = varParams

    ' Log and remove from job queue
    With cJob
        Log.Add "Worker job " & .Action & " (" & .KeyHash & ") completed in " & _
            Round(Perf.MicroTimer - .Start, 2) & " seconds.", False
    End With
    Set cJob = Nothing
    Me.Queue.Remove strKey

End Function


'---------------------------------------------------------------------------------------
' *** BEGIN WORKER SCRIPT ***
'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Main entry point for worker script
'---------------------------------------------------------------------------------------
'
Sub Main()

    Dim objFSO
    Dim objApp
    Dim strLibPath
    Dim strKey
    Dim strValue
    Dim strAction
    Dim blnNeedsApp
    Dim objProject
    Dim varReturn


    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' We are expecting some very specific command line arguments in this worker script.
    ' Three at a minimum: database path, job key and action. Accepting two would raise
    ' an unhandled error further down, where the action is read.
    If WScript.Arguments.Count < 3 Then
        WorkerLog "Invalid arguments provided to worker script"
        WScript.Quit
    End If

    ' The first argument should be the path of the current database
    strValue = WScript.Arguments(0)

    ' The second argument is the key, and the third the action
    strKey = WScript.Arguments(1)
    strAction = WScript.Arguments(2)

    ' A variable that was never assigned is Empty rather than Nothing in VBScript, and
    ' the "Is Nothing" tests below raise an error and take the branch anyway, so the
    ' reference has to be initialized whether or not we go on to bind.
    Set objApp = Nothing

    ' Not every action needs the Access instance. CheckDatabaseAccessible builds its own
    ' DAO engine, opens the file and reports a value: the only thing it ever used the
    ' instance for was handing the answer back. That is a dependency taken out purely in
    ' order to talk, and it is the one that fails when the current database is the add-in
    ' itself, so this action reports through a result file and never binds. An action not
    ' listed here keeps binding, which is the safe default for anything added later.
    blnNeedsApp = True
    If strAction = "CheckDatabaseAccessible" Then blnNeedsApp = False

    If blnNeedsApp Then

        ' Arguments(0) is the target of the bind, so it has to be openable. That is not
        ' an invariant: a full build renames the database to _VCSBackup, and until the
        ' new one exists the path resolves to nothing. An action that does not bind has
        ' no use for it, which is why this guard lives inside here.
        If Not objFSO.FileExists(strValue) Then
            WorkerLog "Database file must exist: " & strValue
            WScript.Quit
        End If

        ' Get a reference to the database application. Binding by file path fails when
        ' the path is the add-in itself: the .accda extension is registered with a
        ' ProgID that has no CLSID, so COM cannot bind it. That is the case whenever the
        ' operation is driven from the add-in file. Leaving the error unhandled shows a
        ' Windows Script Host dialog that blocks this script until somebody dismisses
        ' it, and the caller waits out its whole timeout.
        On Error Resume Next
        Set objApp = GetObject(strValue)
        Err.Clear
        If objApp Is Nothing Then
            On Error GoTo 0
            WorkerLog "Could not attach to Access for " & strValue & " (action " & _
                strAction & "). No result was reported, so the caller will time out."
            WScript.Quit
        End If

        ' Best effort: the VBProjects collection can be in flux while a database is
        ' being reopened, and failing to set the active project is not fatal here.
        Set objApp.VBE.ActiveVBProject = CurrentVBProject(objApp)
        Err.Clear
        On Error GoTo 0

    End If

    ' Determine specified action (More complex actions can be split out into subs)
    Select Case strAction

        Case "DeleteDatabaseObject"
            ' Type, Name
            objApp.DoCmd.DeleteObject WScript.Arguments(3), WScript.Arguments(4)

        Case "RenameDatabaseObject"
            ' New name, Type, Old Name
            objApp.DoCmd.Rename WScript.Arguments(3), WScript.Arguments(4), WScript.Arguments(5)

        Case "Uninstall"
            ' Remove the add-in database file and this worker script.
            ' (This must be done after closing Access)
            ' Args: Access application instance, add-in path
            UninstallAddIn objApp, WScript.Arguments(3)

        Case "BuildAndInstall"
            ' Rebuild the add-in from source and install the result.
            ' (This must be done after closing Access)
            ' Args: Access application instance, source path
            BuildAndInstall objApp, WScript.Arguments(3)

        Case "CheckDatabaseAccessible"
            ' The path to test is passed as a parameter, because Arguments(0) is the
            ' caller's CurrentProject, which is the add-in file itself whenever the
            ' operation is driven from the add-in. Testing that answers the wrong
            ' question. Callers that pass no parameter keep the old behavior.
            If WScript.Arguments.Count > 3 Then
                varReturn = CheckDatabaseAccessible(WScript.Arguments(3))
            Else
                varReturn = CheckDatabaseAccessible(WScript.Arguments(0))
            End If

        Case Else
            WorkerLog "Unknown action: " & strAction

    End Select

    ' Report the result. An action that never bound has no COM channel to answer
    ' through, and answers by writing a file that the caller polls for. An action that
    ' did bind keeps calling back below, because it holds the instance anyway.
    If Not blnNeedsApp Then WriteJobResult objFSO, strKey, varReturn

    ' Make sure we still have a valid reference to the application before
    ' attempting any callback operations.
    If Not objApp Is Nothing Then

        ' GetAddInVBProject leaves its return value unassigned when no matching project
        ' is found, and the VBProjects collection can be in flux while a database is
        ' reopened, so reading .FileName off it raises 424. Uncaught, that shows the
        ' Windows Script Host dialog this script exists to avoid, in the one branch a
        ' build actually reaches.
        Set objProject = Nothing
        On Error Resume Next
        Set objProject = GetAddInVBProject(objApp)
        Err.Clear
        If objProject Is Nothing Then
            WorkerLog "Could not find the add-in VB project to call back into for " & _
                strAction & ". The caller will wait out its timeout."
        Else
            ' Build path to library so we can ensure that we are calling
            ' the add-in function, not a function in the current database.
            strLibPath = objProject.FileName
            ' Remove file extension from library path.
            strLibPath = Left(strLibPath, Len(strLibPath) - Len(".accda"))

            ' Return key to caller to signal job completion
            objApp.Run strLibPath & ".WorkerCallback", strKey, varReturn
            If Err.Number <> 0 Then
                WorkerLog "Callback failed for " & strAction & ": " & Err.Number & " " & _
                    Err.Description
                Err.Clear
            End If
        End If
        On Error GoTo 0
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WorkerLog
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/21/2026
' Purpose   : Append a line to Worker.log, next to the worker script. This script runs
'           : unattended, launched hidden by Shell(), so a failure has to leave a trace
'           : on disk. A dialog cannot be dismissed by anyone and blocks the script,
'           : which turns a failed job into the caller's full timeout.
'---------------------------------------------------------------------------------------
'
Sub WorkerLog(strMessage)

    Dim objLogFSO
    Dim objFile

    On Error Resume Next
    Set objLogFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objLogFSO.OpenTextFile(objLogFSO.GetParentFolderName( _
        WScript.ScriptFullName) & "\Worker.log", 8, True)
    objFile.WriteLine Now & vbTab & strMessage
    objFile.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteJobResult
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/22/2026
' Purpose   : Report a job result through a file instead of a COM callback, for an
'           : action that does not need the Access instance to do its work.
'           : The file is named after the job key, which is unique to each launch, so a
'           : file left behind by an earlier run can never be read as this answer.
'           : Written under a temporary name and renamed, because a rename is atomic on
'           : the same volume: the caller polling for the file cannot catch it half
'           : written. The value goes out as 1 or 0 rather than as a boolean, so nothing
'           : in the channel depends on locale or on string formatting.
'           : The folder is the one clsWorker.JobResultFolder builds, and is created
'           : here as well: a missing folder would leave the caller with no answer and
'           : nothing to look at.
'---------------------------------------------------------------------------------------
'
Sub WriteJobResult(objFSO, strKey, varValue)

    Dim strFolder
    Dim strFile
    Dim strTemp
    Dim objOut

    On Error Resume Next

    ' No value means no answer, and no answer is the caller's third state. Writing a
    ' file here would turn it into a definite "not accessible" instead.
    If IsEmpty(varValue) Then Exit Sub

    strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\WorkerJobs"
    If Not objFSO.FolderExists(strFolder) Then objFSO.CreateFolder strFolder

    ' A CreateFolder that had nothing to do still leaves Err set, and the success path
    ' of this script must not write anything to the log.
    Err.Clear

    strFile = strFolder & "\" & strKey & ".result"
    strTemp = strFile & ".tmp"

    Set objOut = objFSO.CreateTextFile(strTemp, True)
    If ReportJobError("create " & strTemp) Then Exit Sub

    If varValue Then
        objOut.WriteLine "1"
    Else
        objOut.WriteLine "0"
    End If
    objOut.Close
    Set objOut = Nothing
    If ReportJobError("write " & strTemp) Then Exit Sub

    If objFSO.FileExists(strFile) Then objFSO.DeleteFile strFile
    objFSO.MoveFile strTemp, strFile
    If ReportJobError("publish " & strFile) Then Exit Sub

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReportJobError
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/22/2026
' Purpose   : Log the pending error, if any, naming the step that raised it, and report
'           : whether the caller should stop. Checked after each step of WriteJobResult
'           : because a single check at the end would log the last error rather than the
'           : one that actually broke the chain.
'---------------------------------------------------------------------------------------
'
Function ReportJobError(strWhat)
    If Err.Number <> 0 Then
        WorkerLog "Could not " & strWhat & ": " & Err.Number & " " & Err.Description
        Err.Clear
        ReportJobError = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SafeProjectFileName
' Author    : Adam Waller
' Date      : 6/9/2026
' Purpose   : Guarded read of a VBProject's FileName for the isolated worker script.
'           : The VBProjects collection can contain non-VBA entries (registered type
'           : libraries or wizards) that raise an error when .FileName is accessed,
'           : and the project set can be in flux while the database is reopened
'           : (e.g. in shared mode). Returns an empty string on failure. (#709)
'---------------------------------------------------------------------------------------
'
Function SafeProjectFileName(objProj)
    On Error Resume Next
    SafeProjectFileName = objProj.FileName
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : CurrentVBProject
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Get a reference to the VB Project associated with the current database.
'           : (Look for a matching file name between the VB Projects and the database.)
'---------------------------------------------------------------------------------------
'
Function CurrentVBProject(objApp)

    Dim objProj
    Dim strUncPath

    ' Use currently active project by default
    Set CurrentVBProject = objApp.VBE.ActiveVBProject

    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(objApp.CurrentProject.FullName)

    If StrComp(SafeProjectFileName(objApp.VBE.ActiveVBProject), strUncPath, vbTextCompare) <> 0 Then
        ' Search for project with matching filename.
        For Each objProj In objApp.VBE.VBProjects
            If StrComp(SafeProjectFileName(objProj), strUncPath, vbTextCompare) = 0 Then
                Set CurrentVBProject = objProj
                Exit For
            End If
        Next
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddInVBProject
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Get a reference to the add-in project, even if it has the same name as
'           : the current project.
'---------------------------------------------------------------------------------------
'
Function GetAddInVBProject(objApp)

    Dim objProj
    Dim strUncPath
    Dim objFSO
    Dim strProjFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(objApp.CurrentProject.FullName)

    ' Search for project with matching filename.
    For Each objProj In objApp.VBE.VBProjects
        strProjFile = SafeProjectFileName(objProj)
        If Len(strProjFile) > 0 Then
            If StrComp(objFSO.GetBaseName(strProjFile), "Version Control", vbTextCompare) = 0 Then
                ' Check for different file path
                If StrComp(strProjFile, strUncPath, vbTextCompare) <> 0 Then
                    Set GetAddInVBProject = objProj
                    Exit For
                End If
            End If
        End If
    Next

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUncPath
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns the UNC path for a network location (if applicable)
'---------------------------------------------------------------------------------------
'
Function GetUncPath(strPath)

    Dim strDrive
    Dim strUNC
    Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strUNC = strPath
    strDrive = objFSO.GetDriveName(strPath)
    If strDrive <> vbNullString Then
        With objFSO.GetDrive(strDrive)
            If .DriveType = 3 Then  ' Remote
                strUNC = Replace(strPath, strDrive, .ShareName, , 1, vbTextCompare)
            End If
        End With
    End If
    GetUncPath = strUNC

End Function


'---------------------------------------------------------------------------------------
' Procedure : UninstallAddIn
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Remove add-in files that would normally be in use during the uninstall
'           : process. Then remove this worker script file itself.
'           : NOTE: No message is given if the uninstall succeeds.
'---------------------------------------------------------------------------------------
'
Function UninstallAddIn(objApp, strAddInFile)

    Dim objFSO
    Dim strLockFile
    Dim intCnt

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Close the Microsoft Access application
    objApp.Quit
    Set objApp = Nothing

    ' Build lock file name from add-in name
    strLockFile = Replace(strAddInFile, ".accda", ".laccda")
    strLockFile = Replace(strAddInFile, ".accde", ".laccda")

    ' Resume next on any errors while attempting to delete the file.
    On Error Resume Next

    ' Wait up to 10 seconds to delete the main add-in file and database lock file.
    For intCnt = 1 To 10
        If objFSO.FileExists(strLockFile) Then objFSO.DeleteFile strLockFile
        If objFSO.FileExists(strAddInFile) Then objFSO.DeleteFile strAddInFile
        If Not objFSO.FileExists(strAddInFile) Then
            ' Clear any errors from failed delete attempts
            If Err Then Err.Clear
            Exit For
        End If
        WScript.Sleep 1000  ' Pause one second
    Next

    ' Ready or not, here we go with one final attempt. (This time we will show the error message.)
    If objFSO.FileExists(strAddInFile) Then objFSO.DeleteFile strAddInFile
    If Err Then
        MsgBox "Unable to delete add-in file: " & strAddInFile & vbCrLf & vbCrLf & _
            " The file might be in use by another instance of Microsoft Access." & vbCrLf & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Else
        ' Delete this worker script file.
        objFSO.DeleteFile WScript.ScriptFullName
        WScript.Quit
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckDatabaseAccessible
' Author    : Adam Waller
' Date      : 4/2/2026
' Purpose   : Test whether the database file can be opened by an external client.
'           : Uses a separate DAO engine instance (out-of-process) to detect the
'           : JET/ACE engine-level lock state (error 3734) that blocks connections
'           : after build/merge operations involving schema changes.
'---------------------------------------------------------------------------------------
'
Function CheckDatabaseAccessible(strPath)
    Dim objEngine
    Dim objDb
    On Error Resume Next
    Set objEngine = CreateObject("DAO.DBEngine.120")
    Set objDb = objEngine.OpenDatabase(strPath, False, True)
    If Err.Number = 0 Then
        objDb.Close
        Set objDb = Nothing
        CheckDatabaseAccessible = True
    Else
        CheckDatabaseAccessible = False
    End If
    Set objEngine = Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildAndInstall
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Rebuild the VCS add-in from source files and install the result.
'           : Quits the current Access instance, opens a new one to build via the
'           : installed library, polls the build log for completion, then launches
'           : the install phase via /cmd INSTALL.
'---------------------------------------------------------------------------------------
'
Function BuildAndInstall(objApp, strSourcePath)

    Dim objFSO
    Dim objNewAccess
    Dim WshShell
    Dim strMSAccessPath
    Dim strInstalledLib
    Dim strRepoAddinPath
    Dim strLogsPath
    Dim strLockFile
    Dim intCnt

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set WshShell = CreateObject("WScript.Shell")

    ' Derive paths before quitting Access
    strMSAccessPath = objApp.SysCmd(9) & "MSACCESS.EXE"  ' acSysCmdAccessDir = 9

    ' Installed library path (Worker.vbs lives next to the installed .accda)
    strInstalledLib = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\Version Control"

    ' Repo .accda path: source folder is "Version Control.accda.src" so the
    ' .accda sits in the same parent directory.
    Dim strSrcClean
    strSrcClean = strSourcePath
    If Right(strSrcClean, 1) = "\" Then strSrcClean = Left(strSrcClean, Len(strSrcClean) - 1)
    strRepoAddinPath = objFSO.GetParentFolderName(strSrcClean) & "\Version Control.accda"

    ' Build log folder
    If Right(strSourcePath, 1) = "\" Then
        strLogsPath = strSourcePath & "logs\"
    Else
        strLogsPath = strSourcePath & "\logs\"
    End If

    ' Close the current Access instance to release all locks
    objApp.Quit
    Set objApp = Nothing

    ' Build lock file path for the repo .accda
    strLockFile = Replace(strRepoAddinPath, ".accda", ".laccdb", 1, 1, vbTextCompare)

    ' Wait for lock files to clear (up to 10 seconds)
    For intCnt = 1 To 10
        If Not objFSO.FileExists(strLockFile) Then Exit For
        WScript.Sleep 1000
    Next

    ' Snapshot existing build log files
    Dim dictExistingLogs
    Set dictExistingLogs = CreateObject("Scripting.Dictionary")
    SnapshotLogFiles objFSO, strLogsPath, "Build_", dictExistingLogs

    ' --- BUILD PHASE ---
    On Error Resume Next
    Set objNewAccess = CreateObject("Access.Application")
    If Err.Number <> 0 Then
        MsgBox "Failed to create Access application instance." & vbCrLf & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description, vbCritical, "Rebuild Add-In"
        Exit Function
    End If
    On Error GoTo 0

    objNewAccess.Visible = True
    objNewAccess.OpenCurrentDatabase strRepoAddinPath

    ' Close the installer form if AutoExec opened it
    On Error Resume Next
    objNewAccess.DoCmd.Close 2, "frmVCSInstall"  ' acForm = 2
    Err.Clear
    On Error GoTo 0

    ' Trigger build via the installed library
    On Error Resume Next
    objNewAccess.Run strInstalledLib & ".HandleRibbonCommand", "btnBuild", strSourcePath
    If Err.Number <> 0 Then
        Dim strBuildErr
        strBuildErr = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        On Error Resume Next
        objNewAccess.Quit
        On Error GoTo 0
        Set objNewAccess = Nothing
        MsgBox "Failed to trigger build command." & vbCrLf & vbCrLf & _
            strBuildErr, vbCritical, "Rebuild Add-In"
        Exit Function
    End If
    On Error GoTo 0

    ' Poll for build log completion (2s interval, up to 120s)
    Dim strNewLog
    Dim strLogContent
    Dim intWait
    strNewLog = ""
    For intWait = 1 To 60
        WScript.Sleep 2000
        strNewLog = FindNewBuildLog(objFSO, strLogsPath, "Build_", dictExistingLogs)
        If Len(strNewLog) > 0 Then
            strLogContent = ReadTextFile(objFSO, strNewLog)
            If InStr(1, strLogContent, "TOTAL RUNTIME:", vbTextCompare) > 0 Then
                Exit For
            End If
            strNewLog = ""
        End If
    Next

    ' Close the build Access instance
    On Error Resume Next
    objNewAccess.Quit
    On Error GoTo 0
    Set objNewAccess = Nothing
    WScript.Sleep 2000

    ' Check build result
    If Len(strNewLog) = 0 Then
        MsgBox "Build timed out. No completed build log was found." & vbCrLf & vbCrLf & _
            "Check the logs folder for details:" & vbCrLf & strLogsPath, _
            vbCritical, "Rebuild Add-In"
        Exit Function
    End If

    ' Re-read the log to confirm success
    strLogContent = ReadTextFile(objFSO, strNewLog)
    Dim blnSuccess
    blnSuccess = (InStr(1, strLogContent, "Done. (", vbTextCompare) > 0) And _
                  (InStr(1, strLogContent, "CRITICAL:", vbTextCompare) = 0)

    If Not blnSuccess Then
        MsgBox "Build failed. Check the build log for details:" & vbCrLf & vbCrLf & _
            strNewLog, vbCritical, "Rebuild Add-In"
        Exit Function
    End If

    ' --- INSTALL PHASE ---
    ' Wait for lock files from the build to clear
    For intCnt = 1 To 10
        If Not objFSO.FileExists(strLockFile) Then Exit For
        WScript.Sleep 1000
    Next

    ' Launch Access with /cmd INSTALL (blocks until user clicks OK on success dialog)
    WshShell.Run """" & strMSAccessPath & """ """ & strRepoAddinPath & """ /cmd INSTALL", 1, True

End Function


'---------------------------------------------------------------------------------------
' Procedure : SnapshotLogFiles
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Collect existing log file names matching a prefix into a dictionary.
'---------------------------------------------------------------------------------------
'
Sub SnapshotLogFiles(objFSO, strFolder, strPrefix, dictOut)
    Dim objFile
    If Not objFSO.FolderExists(strFolder) Then Exit Sub
    For Each objFile In objFSO.GetFolder(strFolder).Files
        If Left(objFile.Name, Len(strPrefix)) = strPrefix Then
            dictOut.Add objFile.Name, True
        End If
    Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindNewBuildLog
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Find a build log file that wasn't in the original snapshot.
'---------------------------------------------------------------------------------------
'
Function FindNewBuildLog(objFSO, strFolder, strPrefix, dictExisting)
    Dim objFile
    FindNewBuildLog = ""
    If Not objFSO.FolderExists(strFolder) Then Exit Function
    For Each objFile In objFSO.GetFolder(strFolder).Files
        If Left(objFile.Name, Len(strPrefix)) = strPrefix Then
            If Not dictExisting.Exists(objFile.Name) Then
                FindNewBuildLog = objFile.Path
                Exit Function
            End If
        End If
    Next
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadTextFile
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Read entire contents of a text file.
'---------------------------------------------------------------------------------------
'
Function ReadTextFile(objFSO, strPath)
    Dim objStream
    ReadTextFile = ""
    If Not objFSO.FileExists(strPath) Then Exit Function
    Set objStream = objFSO.OpenTextFile(strPath, 1, False)
    If Not objStream.AtEndOfStream Then ReadTextFile = objStream.ReadAll
    objStream.Close
    Set objStream = Nothing
End Function
