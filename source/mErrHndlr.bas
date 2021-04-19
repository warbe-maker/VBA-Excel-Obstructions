Attribute VB_Name = "mErrHndlr"
Option Explicit
#Const ErrMsg = "Custom"    ' System = Error displayed by MsgBox,
'                             Custom = Error displayed by fMsgFrm which is
'                                      without the message box's limitations in size
'                                      and with automated adjustment in width and height
' --------------------------------------------------------------------------------------
' Standard  Module mErrHndlr
'           Global error handling for any VBA Project.
'           - When a call stack is maintained by BoP/EoP)
'             - The full path from the entry procedure to the procedure where the
'               error occured is displayed with the error message
'             - The error number is passed from the error source procedure up to
'               the entry procedure which is a significant advantage for an
'               unatended regression test as follows:
'
'               BoP ErrSrc(PROC)
'               On Error Resume Next
'               <tested procedure>
'               Debug.Assert Err.Number = n or in case the error is a programmed
'               application error: Debug.Asser AppErr(Err.Number) = n
'               EoP ErrSrc(PROC)
'
'           - When the Conditional Compile Argument "ExecTrace = 1" is provided, a
'             complete execution trace with precision time tracking is displayed in the
'             imediate window whenever the code execution has returned to the entry
'             procedure (the topmost procedure with a BoP/EoP statement).
'           - The local Conditional Compile Constant 'ErrMsg = "Custom"' allows the use
'             of the dedicate UserForm lErrMsg which provideds a better readability.
'
' Methods:  - ErrHndlr      Either passes on the error to the caller or when
'                           the entry procedure is reached, displays the
'                           error with a complete path from the entry procedure
'                           to the procedure with the error.
'           - BoP           Maintains the call stack at the Begin of a Procedure
'                           (optional when using this common error handler)
'           - EoP           Maintains the call stack at the End of a Procedure,
'                           triggers the display of the Execution Trace when the
'                           entry procedure is finished and the Conditional Compile
'                           Argument ExecTrace = 1
'           - BoT           Begin of Trace. In contrast to BoP this is for any
'                           group of code lines within a procedure
'           - EoT           End of trace corresponding with the BoT.
'           - ErrMsg        Displays the error message in a proper formated manner
'                           The local Conditional Compile Constant 'ErrMsg = "Custom"'
'                           allows the use of the dedicate UserForm fErrMsg which
'                           provideds a significant better readability.
'                           ErrMsg may be used with or without a call stack.
'           - AppErr      Exclusively for an application (programmed) error used
'                           with "Err.Raise AppErr(n). Whereby n is any positive
'                           number from 1 to ... which is translated into a negative
'                           error number by adding vbObjectError. This prevents any
'                           conflict with a VB Error Number. In return such a negative
'                           number is translated back into its original Application
'                           Error Number, e.g. for test purpose with
'                           "Debug.Assert AppErr(Err.Number) = n.
'                           The function is also used in the ErrMsg to indicate a
'                           programmed Application Error in contrast to a VB error.
'
' Usage:                    Private/Public Sub/Function any()
'                           Const PROC = "any"  ' procedure's name as error source
'
'                              On Error GoTo on_error
'                              BoP ErrSrc(PROC)   ' puts the procedure on the call stack
'
'                              ' <any code>
'
' exit_proc:
'                               ' <any "finally" code like re-protecting an unprotected sheet for instance>
'                               EoP ErrSrc(PROC)   ' takes the procedure off from the call stack
'                               Exit Sub/Function
'
' on_error:
'                            #If Debugging = 1 Then
'                                Stop: Resume    ' allows to exactly locate the line where the error occurs
'                            #End If
'
' Note: When the call stack is not maintained the ErrHndlr will display the message
'       immediately with the procedure the error occours. When the call stack is
'       maintained, the error message will display the call path to the error beginning
'       with the first (entry) procedure in which the call stack is maintained all the
'       call sequence down to the procedure where the error occoured.
'
' Uses: - Class Module clsCallStack
'       - Class Module clsCallStackItem
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin January 2020
' -----------------------------------------------------------------------------
' ~~ Begin of Declarations for withdrawing the title bar ------------------------------------
Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "User32.dll" _
                          Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, _
                                                     ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "User32.dll" _
                          Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, _
                                                     ByVal nIndex As Long, _
                                                     ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function DrawMenuBar Lib "User32.dll" (ByVal hWnd As LongPtr) As Long
Private Const GWL_STYLE  As Long = (-16)
Private Const WS_CAPTION As Long = &HC00000
' ~~ End of Declarations for withdrawing the title bar --------------------------------------
Public CallStack    As clsCallStack
Public dicTrace     As Dictionary       ' Procedure execution trancing records

Public Sub BoP(ByVal sErrSource As String)
' ---------------------------------------------
' Begin of Procedure. Maintains the call stack.
' ---------------------------------------------
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    ElseIf CallStack.StackIsEmpty Then
        Set CallStack = Nothing
        Set CallStack = New clsCallStack
    End If
    CallStack.StackPush sErrSource
End Sub

Public Sub BoT(ByVal s As String)
' ---------------------------------------
' Explicit execution trace start for (s).
' ---------------------------------------
#If ExecTrace Then
    CallStack.TraceBegin s
#End If
End Sub

Public Sub EoP(ByVal sErrSource As String)
' -------------------------------------------
' End of Procedure. Maintains the call stack.
' -------------------------------------------
    If Not CallStack Is Nothing Then
        CallStack.StackPop sErrSource
        If CallStack.StackIsEmpty Then
            If CallStack.ErrorPath = vbNullString Then
                Set CallStack = Nothing
            End If
        End If
    End If
End Sub

Public Sub EoT(ByVal s As String)
' -------------------------------------
' Explicit execution trace end for (s).
' -------------------------------------
    CallStack.TraceEnd s
End Sub

Public Sub ErrHndlr(ByVal lErrNo As Long, _
                    ByVal sErrSource As String, _
                    ByVal sErrText As String, _
                    ByVal sErrLine As String)
' -----------------------------------------------
' When the caller (sErrSource) is the entry
' procedure the error is displayed with the path
' to the error. Otherwise the error is raised
' again to pass it on to the calling procedure.
' The .ErrorPath string is maintained with all
' the way up to the calling procedure.
' -----------------------------------------------
Const PROC      As String = "ErrHndlr"
Static sLine    As String   ' provided error line (if any) for the the finally displayed message
Dim sErrMsg     As String
Dim sTitle      As String
Dim a()         As String
Dim sErrPath    As String
Dim s           As String
Dim sIndicate   As String
   
    On Error GoTo on_error
    
    If lErrNo = 0 Then
        Stop: Resume
        mCommon.ErrMsg AppErr(1), sErrSource, "An ""Exit ..."" statement before error handling missing! Error number is 0!", Erl
    End If
    
    If CallStack Is Nothing Then Set CallStack = New clsCallStack
    If sErrLine <> 0 Then sLine = sErrLine
    
    With CallStack
        '~~ Provide a line in the Execution Trace indicating the error by its number and description
        If .ErrorSource = vbNullString Then
            '~~ This is the "entry procedure"
            .ErrorSource = sErrSource
            .SourceErrorNo = lErrNo
            .ErrorNumber = lErrNo
            .ErrorDescription = sErrText
            .ErrorPath = .ErrorPath & sErrSource & " (" & ErrorDetails(lErrNo, sErrLine) & ")" & vbLf
            .TraceError ErrorDetails(lErrNo, sErrLine) & ": " & sErrText
        ElseIf .ErrorNumber <> lErrNo Then
            '~~ The error number had changed while passing it on to the entry procedure
            .ErrorPath = .ErrorPath & sErrSource & " (" & ErrorDetails(lErrNo, sErrLine) & ")" & vbLf
            .TraceError ErrorDetails(lErrNo, sErrLine) & ": " & sErrText
            .ErrorNumber = lErrNo
        Else
            .ErrorPath = .ErrorPath & sErrSource & vbLf
        End If
        '~~ End of trace for the procedure which caused/raised an error
        .TraceEnd sErrSource
        If .EntryProc <> sErrSource Then ' And Not .ErrorPath <> vbNullString Then
            '~~ As long as the entry procedure has not been reached the error is passed to the calling procedure
            Err.Raise lErrNo, sErrSource, sErrText
        
        ElseIf .EntryProc = sErrSource Then
            '~~ When the entry procedure has been reached:
            '~~ - the error path string is comppleted
            '~~ - the error is displayed
            ErrMsg .SourceErrorNo, .ErrorSource, .ErrorDescription, sLine
            
            '~~ ----------------------------------------------------------------------
            '~~ This is the place for any project specific clean-up code which had not
            '~~ been performed because of the interuption caused by the error.
            '~~ Attention: When Component Management is used to automatically update
            '              this module's code when outdated, the Conditional Compile
            '              Argument
            '~~ ----------------------------------------------------------------------
#If ExecTrace Then
        DsplyTrace
#End If
        End If
    End With
    Exit Sub
    
on_error:
#If Debugging Then
    Stop: ' Resume
#End If
    ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Public Sub ErrMsg(ByVal lErrNo As Long, _
                  ByVal sErrSrc As String, _
                  ByVal sErrDesc As String, _
                  ByVal sErrLine As String)
' -------------------------------------------
' Displays the error message by means of
' MsgBox or, when the Conditional Compile
' Argument ErrMsg = "Custom", by means of the
' Common Component fMsg. In any case the path
' to the error may be displayed provided a
' call stack is available.
'
' W. Rauschenberger Berlin March 2020
' --------------------------------------------
Dim sErrMsg     As String
Dim sTitle      As String
Dim a           As Variant
Dim sErrPath    As String
Dim s           As String
Dim sIndicate   As String
Dim sLine       As String
Dim i           As Long
Dim sErrText    As String
Dim sErrInfo    As String
Dim iIndent     As Long

    '~~ Additional info about the error line in case one had been provided
    If sErrLine = vbNullString Or sErrLine = "0" Then
        sIndicate = vbNullString
    Else
        sIndicate = " (at line " & sErrLine & ")"
    End If
    sTitle = sTitle & sIndicate
        
    '~~ Path from the entry procedure (the first which uses BoP/EoP)
    '~~ all the way down to the procedure in which the error occoured.
    '~~ When the call stack had not been maintained the path is empty.
    If Not CallStack Is Nothing Then
        If Not CallStack.ErrorPath = vbNullString Then
            CallStack.TraceEndTime = Now()
            CallStack.StackUnwind
            a = Split(CallStack.ErrorPath, vbLf)
            '~~ In case an error path had been maintained
            '~~ the line indicates the error source
            ReDim Preserve a(UBound(a) - 1)
            sErrSrc = Split(a(LBound(a)), " ")(0)
            If UBound(a) > 1 Then
                iIndent = -1
                For i = UBound(a) To LBound(a) Step -1
                    If i = UBound(a) Then
                        sErrPath = a(i) & vbLf
                    Else
                        sErrPath = sErrPath & Space((iIndent) * 2) & "|_" & a(i) & vbLf
                    End If
                    iIndent = iIndent + 1
                Next i
            End If
        End If
    End If
    
    '~~ Prepare the Title with the error number and the procedure which caused the error
    Select Case lErrNo
        Case Is > 0:    sTitle = "VBA Error " & lErrNo
        Case Is < 0:    sTitle = "Application Error " & AppErr(lErrNo)
    End Select
    sTitle = sTitle & " in:  " & sErrSrc & sIndicate
         
    '~~ Consider the error description may include an additional information about the error
    '~~ possible only when the error is raised by Err.Raise
    If InStr(sErrDesc, MSG_CONCAT) <> 0 Then
        sErrText = Split(sErrDesc, MSG_CONCAT)(0)
        sErrInfo = Split(sErrDesc, MSG_CONCAT)(1)
    Else
        sErrText = sErrDesc
        sErrInfo = vbNullString
    End If
                       
#If ErrMsg = "Custom" Then
    '~~ Display the error message by means of the Common UserForm fMsg
    mCommon.ErrMsg lErrNo:=lErrNo, sTitle:=sTitle, sErrDesc:=sErrText, sErrPath:=sErrPath, sErrInfo:=sErrInfo
#Else
    '~~ Assemble error message to be displayed by MsgBox
    sErrMsg = "Source: " & vbLf & sErrSrc & sIndicate & vbLf & vbLf & _
              "Error: " & vbLf & sErrText
    If sErrPath <> vbNullString Then
        sErrMsg = sErrMsg & vbLf & vbLf & "Call Stack:" & vbLf & sErrPath
    End If
    If sErrInfo <> vbNullString Then
        sErrMsg = sErrMsg & vbLf & "About: " & vbLf & sErrInfo
    End If
    MsgBox sErrMsg, vbCritical, sTitle
#End If
End Sub

Private Sub DsplyTrace()
' ------------------------------------------------------------
' Displays the current execution trace when the current
' procedure is the entry procedure. This condition is
' required since all items possibly remaining on the call
' stack due to an error (which had prevented them being poped)
' are unwinded in order to finish their execution trace.
' Note:
' The call stack is used to display the path to the error when
' the error message is displayed. The execution trace is
' stored in this module's Dictionary (dicTrace).
' ------------------------------------------------------------
    If CallStack Is Nothing Then
        Set CallStack = New clsCallStack
    End If
    CallStack.TraceDsply
    Set CallStack = Nothing
End Sub

Private Function ErrorDetails(ByVal lErrNo As Long, _
                              ByVal sErrLine As String) As String
' -----------------------------------------------------------------
' Returns kind of error, error number, and error line if available.
' -----------------------------------------------------------------
Dim s As String
    If lErrNo < 0 Then
        s = "App error " & AppErr(lErrNo)
    Else
        s = "VB error " & lErrNo
    End If
    If sErrLine <> 0 Then
        s = s & " at line " & sErrLine
    End If
    ErrorDetails = s
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mErrHndlr" & "." & sProc
End Function
