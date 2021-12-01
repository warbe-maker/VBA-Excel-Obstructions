Attribute VB_Name = "mTest"
Option Explicit

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
    
#If ErHComp Then
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
#If MsgComp Then
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
    ErrSrc = "mTest" & "." & sProc
End Function

Public Sub Names()
    Const PROC = "Names"
    
    On Error GoTo eh
    Dim nm      As Name
    Dim lRow    As Long
    Dim wb      As Workbook
    Dim ws      As Worksheet
    
    Set wb = wsTest1.Parent
    
    With wsTest1
        .RngNames.ClearContents
        lRow = .RngNames.row - 1
        For Each nm In wb.Names
            Debug.Print nm.RefersTo
            lRow = lRow + 1
            Intersect(.RngNames, .NamesSheet.EntireColumn, .Rows(lRow)).Value = Split(nm.RefersTo, "!")(0)
            Intersect(.RngNames, .NamesReference.EntireColumn, .Rows(lRow)).Value = Split(nm.RefersTo, "!")(1)
            If InStr(nm.Name, "!") <> 0 Then
                Intersect(.RngNames, .NamesName.EntireColumn, .Rows(lRow)).Value = Split(nm.Name, "!")(1)
            Else
                Intersect(.RngNames, .NamesName.EntireColumn, .Rows(lRow)).Value = nm.Name
            End If
            Intersect(.RngNames, .NamesScope.EntireColumn, .Rows(lRow)).Value = "Workbook"
        Next nm
    
        .Sort.SortFields.Clear
        .Sort.SortFields.Add Key _
            :=Range("G3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With .Sort
            .SetRange Range("G3:K58")
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    
    End With
    
xt: mObstructions.Rewind ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub TestSetUp()
' ------------------------------------------------------------------------------
' Setup/prepare all Test-Worksheets.
' Important: Sheet protection prevents that AutoFilter is turned off in another
'            sheet while it is turned on in one sheet
' ------------------------------------------------------------------------------
    
    wsTest1.Protect
    wsTest2.Protect
    wsTest3.Protect
    
    With wsTest1
        .Unprotect
        .TestColHidden1.EntireColumn.Hidden = True
        If .AutoFilterMode = False Then .AutoFilter1.AutoFilter
        .Protect
    End With
    With wsTest2
        .Unprotect
        If .AutoFilterMode = False Then .AutoFilter2.AutoFilter Field:=1 _
                                                              , Criteria1:="<>*Filtered*" _
                                                              , Operator:=xlAnd
        .TestColHidden2.EntireColumn.Hidden = True
        .Protect
    End With
    With wsTest3
        .Unprotect
        If .AutoFilterMode = False Then .AutoFilter3.AutoFilter
        .Protect
    End With
    
    wsTest2.Unprotect
    Application.EnableEvents = True
    
End Sub

Public Sub Test_01_ApplEvents()
' ------------------------------------------------------
' The test procedure will halt at any assertion not met.
' ------------------------------------------------------
    Const PROC = "Test_01_ApplEvents"

    On Error GoTo eh
    Dim ws  As Worksheet
    
    BoP ErrSrc(PROC)
    Set ws = wsTest1
    
    Application.EnableEvents = True
    mObstructions.ApplEvents enEliminate
    Debug.Assert Application.EnableEvents = False
    mObstructions.ApplEvents enRestore
    Debug.Assert Application.EnableEvents = True
    
    '~~ Any subsequent SaveAndOff and CleanUp (usually in nested sub procedures) must not change the status
    Debug.Assert Application.EnableEvents = True
    
    mObstructions.ApplEvents enEliminate
    Debug.Assert Application.EnableEvents = False
    mObstructions.ApplEvents enEliminate
    mObstructions.ApplEvents enEliminate
    mObstructions.ApplEvents enEliminate
    mObstructions.ApplEvents enRestore
    Debug.Assert Application.EnableEvents = False
    mObstructions.ApplEvents enRestore
    Debug.Assert Application.EnableEvents = False
    mObstructions.ApplEvents enRestore
    Debug.Assert Application.EnableEvents = False
    mObstructions.ApplEvents enRestore
    
    Debug.Assert Application.EnableEvents = True
        
xt: mObstructions.Rewind   ' check if something to retore remained
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_02_SheetProtection()
' ------------------------------------------------------------
' Whichever number of sheets, protected or not is unprotected,
' when finally their protection status is restored it is like
' is was at the beginning.
' Assertions proof the correctness of this obstruction
' implementation.
' ------------------------------------------------------------
    Const PROC = "Test_02_SheetProtection"
        
    On Error GoTo eh
    
    TestSetUp
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = True
    
    BoP ErrSrc(PROC)
    
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest1
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest2
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest3
    '~~ Subsequent protection status save operations
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest3
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest1
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest2
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest2
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest3
    mObstructions.SheetProtection sp_service:=enEliminate, sp_ws:=wsTest1
    
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest2
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest3
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest1
    Debug.Assert wsTest1.ProtectContents = False
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = False
    
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest2
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest3
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest1
    Debug.Assert wsTest1.ProtectContents = False
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = False
    
    '~~ Only the final paired protection status 'Restore' operation restores the initial status
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest2
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest3
    mObstructions.SheetProtection sp_service:=enRestore, sp_ws:=wsTest1
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = True
    
xt: '~~ Must not display any yet undone cleanup message
    mObstructions.Rewind
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_03_FilteredRowsHiddenCols()
' --------------------------------------------------------------------
' Assertions proof the correctness of this obstruction implementation.
' Obstructions SaveAndOff and CleanUp allways have to be pairs.
' Unpaired Save/CleanUp will lead to incorrect results!!!!
'
' Important: Filtering can only set off/on sheet wise when all not
'            concerned sheets are protected while AutoFilterMode of
'            one changes.
' --------------------------------------------------------------------
    Const PROC = "Test_03_FilteredRowsHiddenCols"
    
    On Error GoTo eh

    '~~  Setup test sheets and assert the required initial status
    TestSetUp
    With wsTest1
        Debug.Assert .AutoFilterMode = True
        Debug.Assert .ProtectContents = True
        Debug.Assert .TestColHidden1.EntireColumn.Hidden = True
    End With
    With wsTest2
        Debug.Assert .AutoFilterMode = True
        Debug.Assert .ProtectContents = False
        Debug.Assert .TestColHidden2.EntireColumn.Hidden = True
    End With
    With wsTest3
        Debug.Assert .AutoFilterMode = True
        Debug.Assert .ProtectContents = True
        Debug.Assert wsTest3.ProtectContents = True
    End With
    
    BoP ErrSrc(PROC)
    
    '~~ Note! Obstructions are saved-and-turned and restored Worksheetwise
    wsTest1.Activate
    Application.ScreenUpdating = False
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest1
    Application.ScreenUpdating = True
    
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = False
    Debug.Assert wsTest2.AutoFilterMode = True:     Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = True

    Application.ScreenUpdating = False
    mObstructions.FilteredRowsHiddenCols frhc_service:=enRestore, frhc_ws:=wsTest1
    Application.ScreenUpdating = True

    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = True:     Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = True
    
    wsTest2.Activate
    
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest2
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = False
    Debug.Assert wsTest3.AutoFilterMode = True
    
    mObstructions.FilteredRowsHiddenCols frhc_service:=enRestore, frhc_ws:=wsTest2
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = True:     Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = True
    
    wsTest3.Activate
    
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest3
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = True:     Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = False
    
    mObstructions.FilteredRowsHiddenCols frhc_service:=enRestore, frhc_ws:=wsTest3
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = True:     Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = True
        
    '~~ Note! Obstructions are handled Worksheet by Worksheet and have to be restored in reverse order
    '~~       When not performed in reverse order an error is thrown and the service Rewind will do it
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest1
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest2
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest3
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = False
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = False
    Debug.Assert wsTest3.AutoFilterMode = False
    mObstructions.FilteredRowsHiddenCols frhc_service:=enRestore, frhc_ws:=wsTest3
    mObstructions.FilteredRowsHiddenCols frhc_service:=enRestore, frhc_ws:=wsTest2
    mObstructions.FilteredRowsHiddenCols frhc_service:=enRestore, frhc_ws:=wsTest1
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = True:     Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = True
        
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest1
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest2
    mObstructions.FilteredRowsHiddenCols frhc_service:=enEliminate, frhc_ws:=wsTest3
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = False
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = False
    Debug.Assert wsTest3.AutoFilterMode = False
    
    mObstructions.Rewind
    Debug.Assert wsTest1.AutoFilterMode = True:    Debug.Assert wsTest1.TestColHidden1.EntireColumn.Hidden = True
    Debug.Assert wsTest2.AutoFilterMode = True:    Debug.Assert wsTest2.TestColHidden2.EntireColumn.Hidden = True
    Debug.Assert wsTest3.AutoFilterMode = True

xt: mObstructions.Rewind   ' check if something to retore remained
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  Stop: GoTo xt
    End Select
End Sub

Public Sub Test_04_MergedAreas()
    Const PROC = "Test_04_MergedAreas"

    On Error GoTo eh
    Dim wb  As Workbook
    Dim ws  As Worksheet
    Dim r   As Range
    
    BoP ErrSrc(PROC)
    Set r = wsTest2.MergedCellsSelect
    Set ws = r.Worksheet
    Set wb = ws.Parent
    Test_04_MergedAreas_Setup
    ws.Protect
    
    ws.Activate
    Debug.Assert wsTest2.MergedCells1.MergeCells = True
    Debug.Assert wsTest2.MergedCells2.MergeCells = True
    Debug.Assert wsTest2.MergedCells3.MergeCells = True
    mObstructions.MergedAreas mc_service:=enEliminate, mc_range:=r
    
    Debug.Assert wsTest2.MergedCells1.MergeCells = False
    Debug.Assert wsTest2.MergedCells2.MergeCells = False
    Debug.Assert wsTest2.MergedCells3.MergeCells = True ' col dependant not implemented
    
    mObstructions.MergedAreas mc_service:=enRestore, mc_range:=r
    Debug.Assert wsTest2.MergedCells1.MergeCells = True
    Debug.Assert wsTest2.MergedCells2.MergeCells = True
    Debug.Assert wsTest2.MergedCells3.MergeCells = True

xt: mObstructions.Rewind   ' check if something to retore remained
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_04_MergedAreas_Setup()
    Const PROC = "Test_04_MergedAreas_Setup"
    
    On Error GoTo eh
    Application.DisplayAlerts = False
    SheetProtection sp_service:=enEliminate, sp_ws:=wsTest2
    
    If wsTest2.UnMerged1.MergeCells = False Then
        wsTest2.UnMerged1.Merge
        With wsTest2.MergedCells1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
    End If
    
    If wsTest2.UnMerged2.MergeCells = False Then
        wsTest2.UnMerged2.Merge
        With wsTest2.MergedCells2
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
    End If
    
    If wsTest2.UnMerged3.MergeCells = False Then
        wsTest2.UnMerged3.Merge
        With wsTest2.MergedCells3
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
    End If

xt: Application.DisplayAlerts = True
    SheetProtection sp_service:=enRestore, sp_ws:=wsTest2
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_99_All()
' --------------------------------------------------------------------
' Self asserting, unatended regression test.
' --------------------------------------------------------------------
    Const PROC  As String = "Test_99_All"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    '~~  Test sheets setup with assertion of the required initial status
    
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
    mBasic.TimedDoEvents "> " & ErrSrc(PROC)
#End If

    Test_01_ApplEvents
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
    mBasic.TimedDoEvents "Test_01_ApplEvents"
#End If

    Test_02_SheetProtection         ' Basic obstruction also used by other obstructions
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
    mBasic.TimedDoEvents "Test_02_SheetProtection"
#End If

    Test_03_FilteredRowsHiddenCols
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
    mBasic.TimedDoEvents "Test_03_FilteredRowsHiddenCols"
#End If

    Test_04_MergedAreas
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
    mBasic.TimedDoEvents "Test_04_MergedAreas"
    mBasic.TimedDoEvents "< " & ErrSrc(PROC)
#End If

xt: mObstructions.Rewind   ' Should not do anything
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub Test_BoP_EoP()
    Const PROC = "Test_BoP_EoP"
    
    BoP ErrSrc(PROC)
    EoP ErrSrc(PROC)
    
End Sub

