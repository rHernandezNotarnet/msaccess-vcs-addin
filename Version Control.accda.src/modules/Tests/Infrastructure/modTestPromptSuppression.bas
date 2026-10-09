Attribute VB_Name = "modTestPromptSuppression"
'---------------------------------------------------------------------------------------
' Module    : modTestPromptSuppression
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Tests that a harness-driven run cannot be blocked by a prompt.
'           : These run in the project being tested, which is not the project driving the
'           : run: the add-in exports its own source, so testing it loads a second copy of
'           : this project, and the operation that owns the run belongs to the other copy.
'           : Nothing here can see that operation, which is exactly the situation a user
'           : database is in. So the contract asserted here is the one every hosted
'           : project gets: modTestAssert.TestRunActive is true, and PromptWouldDisplay is
'           : false for every kind of prompt.
'           : The same module also covers the other half of the rule, which has nothing
'           : to do with a test run: a root nobody is attending never gets a dialog,
'           : whoever asked for it. That is the state every API and MCP call arrives in.
'           : MsgBox2 is only called where suppression is already established, and never
'           : with the user-gesture flag.
'           : Run: ?VCS.RunTests("modTestPromptSuppression")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Private Const ModuleName As String = "modTestPromptSuppression"


Public Sub TestRunAnnouncesItselfToTheHostedProject()
    TestAssert TestRunActive, "the run driving this project announced itself"
End Sub


Public Sub TestRunSuppressesOrdinaryPrompts()
    TestAssert Not PromptWouldDisplay(False), "ordinary prompt suppressed during a test run"
End Sub


Public Sub TestRunSuppressesGesturePromptsToo()
    ' A gesture prompt outlives silent mode when a person is driving, but during a test
    ' run the runner is the caller: there is no gesture, and a dialog stalls the suite.
    TestAssert Not PromptWouldDisplay(True), "gesture prompt suppressed during a test run"
End Sub


Public Sub TestSuppressionOutranksTheLocalOperation()
    Dim eimPrior As eInteractionMode

    ' No root is active in this project, so relaxing the local operation is allowed here.
    ' Suppression must not depend on it.
    eimPrior = Operation.InteractionMode
    Operation.InteractionMode = eimNormal
    TestAssert Not PromptWouldDisplay(False), "prompts stay suppressed under a relaxed local operation"
    Operation.InteractionMode = eimPrior
End Sub


Public Sub TestSilentMsgBox2ReturnsDefaultResult()
    Dim intResult As VbMsgBoxResult

    ' The only live MsgBox2 call in the suite, and it runs only once suppression is
    ' confirmed. If the precondition fails the call is skipped rather than attempted:
    ' a real dialog here stalls the whole run until someone clicks it.
    If PromptWouldDisplay(False) Then
        TestAssert False, "precondition failed: prompts are not suppressed, MsgBox2 call skipped"
        Exit Sub
    End If

    TestAssert True, "precondition: prompts are suppressed"
    intResult = MsgBox2("Suppressed prompt", "line1", , vbYesNo, , vbNo)
    TestAssert intResult = vbNo, "suppressed MsgBox2 returns the unattended default"
End Sub


Public Sub TestSilentMsgBox2AnswersOnlyTheCallerDefault()
    Dim intResult As VbMsgBoxResult

    ' What a suppressed question answers is intDefaultResult and nothing else: the
    ' dialog's own default button is not consulted, and a caller that passes no default
    ' gets vbOK, whatever buttons the question offered. Each call site reads that answer
    ' as its decision, so this pins the decisions an unattended run makes on its own.
    If PromptWouldDisplay(False) Then
        TestAssert False, "precondition failed: prompts are not suppressed, MsgBox2 calls skipped"
        Exit Sub
    End If

    intResult = MsgBox2("Suppressed prompt", , , vbYesNo + vbDefaultButton2)
    TestAssert intResult = vbOK, "a Yes/No question with no default answers vbOK, not its default button"
    intResult = MsgBox2("Suppressed prompt", , , vbOKCancel + vbDefaultButton2)
    TestAssert intResult = vbOK, "an OK/Cancel question with no default answers vbOK, not its default button"
    intResult = MsgBox2("Suppressed prompt", , , vbRetryCancel)
    TestAssert intResult = vbOK, "a Retry/Cancel question with no default answers vbOK, which is neither button"
    intResult = MsgBox2("Suppressed prompt", , , vbYesNoCancel + vbDefaultButton3, , vbYes)
    TestAssert intResult = vbYes, "a caller default wins over the dialog's default button"
    intResult = MsgBox2("Suppressed prompt", , , vbAbortRetryIgnore, , vbAbort)
    TestAssert intResult = vbAbort, "an Abort/Retry/Ignore question answers the caller default"
End Sub


Public Sub TestUnattendedRootSuppressesEveryPrompt()
    ' The state an API or MCP call arrives in: nothing asked for silence, and the root
    ' captured Attended = False from the automation source. Until this was covered, an
    ' error raised inside VCS.ImportObject reached MsgBox2 and blocked the COM caller
    ' until the Access process was killed.
    TestAssert Not PromptPolicyAllows(eimNormal, blnAttended:=False), _
        "ordinary prompt suppressed on an unattended root in normal mode"
    TestAssert Not PromptPolicyAllows(eimNormal, blnAttended:=False, blnUserGesturePrompt:=True), _
        "gesture prompt suppressed on an unattended root: there was no gesture"
    TestAssert Not PromptPolicyAllows(eimSilent, blnAttended:=False), _
        "an unattended root is suppressed in silent mode too"
End Sub


Public Sub TestAttendedRootKeepsItsDialogs()
    ' Suppressing an unattended root must not silence the person at the ribbon.
    TestAssert PromptPolicyAllows(eimNormal, blnAttended:=True), _
        "an attended operation in normal mode still prompts"
    TestAssert Not PromptPolicyAllows(eimSilent, blnAttended:=True), _
        "silent mode still suppresses an ordinary prompt"
    TestAssert PromptPolicyAllows(eimSilent, blnAttended:=True, blnUserGesturePrompt:=True), _
        "a gesture prompt still outlives silent mode when a person is driving"
End Sub


Public Sub TestAutomationRootIsUnattendedAndSuppressed()
    ' Ties the rule to the state clsOperation actually produces for an API caller, so
    ' the two cannot drift apart. Runs on a private instance, clear of the singleton.
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = New clsOperation
    cOp.Source = eosExternalAPI
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "root acquired on the private instance"
    TestAssert Not cOp.Attended, "an external API root is unattended"
    TestAssert cOp.InteractionMode = eimNormal, "nothing asked for silence on the way in"
    TestAssert Not PromptPolicyAllows(cOp.InteractionMode, cOp.Attended), _
        "the state an API call arrives in suppresses prompts"
    cRoot.Complete eorSuccess
End Sub


Public Sub TestApiCallerIsUnattendedBeforeItsRoot()
    ' Some API entry points prompt before they begin a root (MigrateDebugAssert never
    ' begins one; ExecuteTests installs modTestAssert first). No root has captured
    ' anything yet, so the incoming caller has to decide.
    Dim cOp As clsOperation

    Set cOp = New clsOperation
    cOp.Source = eosExternalAPI
    TestAssert Not cOp.Attended, "an API caller with no root yet is unattended"
    cOp.Source = eosMCPTool
    TestAssert Not cOp.Attended, "an MCP caller with no root yet is unattended"
    cOp.Source = eosUserInterface
    TestAssert cOp.Attended, "a ribbon caller with no root is attended"
    cOp.ForceUnattended = True
    TestAssert Not cOp.Attended, "ForceUnattended with no root is unattended"
End Sub


Public Sub TestRibbonAfterApiRootIsAttendedAgain()
    ' The root's captured value describes that root only. Once it finishes, a ribbon
    ' command - which resets Source - must not inherit "unattended" from the API call
    ' that ran before it in the same Access instance.
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = New clsOperation
    cOp.Source = eosExternalAPI
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "root acquired on the private instance"
    TestAssert Not cOp.Attended, "the API root is unattended while it runs"
    cRoot.Complete eorSuccess
    cOp.Source = eosUserInterface
    TestAssert cOp.Attended, "a ribbon command after the API root is attended"
End Sub


Public Sub TestRibbonCommandClearsAStrandedForceUnattended()
    ' A headless entry sets ForceUnattended before it begins its root; Finish clears it,
    ' and so does the entry when its root is refused. A root that never finishes leaves
    ' it set, and outside a root it would make the next ribbon command unattended. The
    ' ribbon reset has to clear it too.
    Dim cOp As clsOperation

    Set cOp = New clsOperation
    cOp.Source = eosExternalAPI
    cOp.ForceUnattended = True
    cOp.ResetForInteractiveCommand
    TestAssert cOp.Source = eosUserInterface, "the ribbon reset restores the interactive source"
    TestAssert Not cOp.ForceUnattended, "the ribbon reset clears a stranded ForceUnattended"
    TestAssert cOp.Attended, "a ribbon command after a stranded ForceUnattended is attended"
End Sub


Public Sub TestLiveRootKeepsItsCapturedAttended()
    ' Inside a root the captured value still holds, whatever Source does meanwhile.
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = New clsOperation
    cOp.Source = eosExternalAPI
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cOp.Source = eosUserInterface
    TestAssert Not cOp.Attended, "a live API root stays unattended when Source changes"
    cRoot.Complete eorSuccess
End Sub
