Attribute VB_Name = "mPendingImplementation"
Option Explicit

Public Sub NamedRanges(ByVal nr_service As enObstService, _
                       ByVal nr_ws As Worksheet)
' ------------------------------------------------------------------------------
' Basic obstruction "Named Ranges":
' - nr_service = enEliminate: All references to a name in Worksheet (nr_ws)
'   are turned into a direct range reference and saved in a Collection of
'   name objects
' - nr_service = enRestore: All name objects are established as range names
'   referring to Worksheet (nr_ws)
'
' Note: Names are regarded an obstruction when the Worksheet (nr_ws) is to be
'       copied from one Workbook to another because all named ranges are
'       - pssibly unintended - back-linked to the source Worksheet.
'
' W. Rauschenberger, Berlin Nov 2021
' ------------------------------------------------------------------------------
    Const PROC = "BasicObstNamedRanges"
    
    On Error GoTo eh
    Dim dcNames     As Dictionary   ' Names which refer to a range in the Worksheet nr_ws
    Dim dcCells     As Dictionary   ' Keeps a record of each cell's formula which had been modified
    Dim nm          As Name
    Dim nms         As Names
    Dim cel         As Range
    Dim v           As Variant
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = nr_ws.Parent
    Set nms = wb.Names
    
    Select Case nr_service
        Case enEliminate
            '~~ Collect the relevant names
            If dcNames Is Nothing Then Set dcNames = New Dictionary Else dcNames.RemoveAll
            For Each nm In nms
                dcNames.Add nm.Name, nm
            Next nm
            '~~ Collect cells with a formula referencing on of the collected names
            For Each ws In wb.Sheets
                SheetProtection obs_mode:=enEliminate, obs_ws:=ws
                For Each cel In ws.UsedRange.SpecialCells(xlCellTypeFormulas)
                    Debug.Print "cell with formula: " & cel.Address   ' & ", formula=" & cel.Formula2Local
                    With cel
                        For Each v In dcNames
                            If InStr(.Formula, v) <> 0 Then
                                dcCells.Add cel, .Formula ' At least one name is used
                            End If
                        Next v
                    End With
                Next cel
                SheetProtection obs_mode:=enRestore, obs_ws:=ws
            Next ws
        
        Case enRestore
        
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_10_NamedRanges()
' ------------------------------------------------------------------
' Range names imply a serious problem when a worksheet (wsSource)
' is about to be copied into another Workbook (wbTarget) since all
' names would refer back to the source Workbook (wbSource). To
' avoid this formulas using relevant range names are temporarily
' turned into comments and restored after the Worksheet had been
' copied.
' ------------------------------------------------------------------
    Const PROC = "Test_10_NamedRanges"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)

    mPendingImplementation.NamedRanges nr_service:=enEliminate, nr_ws:=wsTest3
    mPendingImplementation.NamedRanges nr_service:=enRestore, nr_ws:=wsTest3
    
xt: mObstructions.Rewind
    EoP ErrSrc(PROC)
    '~~ Must not display any yet undone cleanup message
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' End of Procedure stub. Handed over to the corresponding procedures in the
' Common Component mTrc (Execution Trace) or mErH (Error Handler) provided the
' components are installed which is indicated by the corresponding Conditional
' Compile Arguments.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 And TrcComp = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' This is a kind of universal error message which includes a debugging option.
' It may be copied into any module - turned into a Private function. When the/my
' Common VBA Error Handling Component (ErH) is installed and the Conditional
' Compile Argument 'CommErHComp = 1' the error message will be displayed by
' means of the Common VBA Message Component (fMsg, mMsg).
'
' Usage: When this procedure is copied as a Private Function into any desired
'        module an error handling which consideres the possible Conditional
'        Compile Argument 'Debugging = 1' will look as follows
'
'            Const PROC = "procedure-name"
'            On Error Goto eh
'        ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC)
'               Case vbYes: Stop: Resume
'               Case vbNo:  Resume Next
'               Case Else:  Goto xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Used:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
' ------------------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
#If CommErHComp Then
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    ErrMsg = mErH.ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    '~~ Translate back the elaborated reply buttons mErrH.ErrMsg displays and returns to the simple yes/No/Cancel
    '~~ replies with the VBA MsgBox.
    Select Case ErrMsg
        Case mErH.DebugOptResumeErrorLine:  ErrMsg = vbYes
        Case mErH.DebugOptResumeNext:       ErrMsg = vbNo
        Case Else:                          ErrMsg = vbCancel
    End Select
#Else
    '~~ When the Common VBA Error Handling Component (ErH) is not used/installed there might still be the
    '~~ Common VBA Message Component (Msg) be installed/used
#If CommMsgComp Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
#Else
    '~~ None of the Common Components is installed/used
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
#End If
#End If
End Function
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mPendingImplementation" & "." & sProc
End Function

Private Sub BoP(ByVal b_proc As String, _
           ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Begin of Procedure stub. Handed over to the corresponding procedures in the
' Common Component mTrc (Execution Trace) or mErH (Error Handler) provided the
' components are installed which is indicated by the corresponding Conditional
' Compile Arguments.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 And TrcComp = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

