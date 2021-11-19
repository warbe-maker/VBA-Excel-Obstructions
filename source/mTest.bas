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
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
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
              err_dscrptn & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    
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
    ErrSrc = "mTest" & "." & sProc
End Function

Public Sub Names()
    Const PROC = "Names"
    
    On Error GoTo eh
    Dim nm      As Name
    Dim lRow    As Long
    Dim wb      As Workbook

    Set wb = ThisWorkbook
    
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
    
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub TestSetUp()
' -----------------------------------------------
' Setup/prepare all Test-Worksheets.
' -----------------------------------------------
    
    With wsTest1
        .Unprotect
        If .AutoFilterMode = False Then .Range("rngAutoFilter1").AutoFilter
        .TestColHidden.EntireColumn.Hidden = True
        .Protect
    End With
    With wsTest2
        .Unprotect
        If .AutoFilterMode = True Then .AutoFilterMode = False
    End With
    With wsTest3
        .Unprotect
        If .AutoFilterMode = False Then .Range("rngAutoFilter3").AutoFilter
        .Protect
    End With
    Application.EnableEvents = True

End Sub

Public Sub Test_All()
' --------------------------
' Unatended regression test.
' All results asserted
' --------------------------
    Const PROC  As String = "Test_All"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    '~~  Test sheets setup with assertion of the required initial status
    TestSetUp
    Debug.Assert mObstructions.WsColsHidden(wsTest1) = True
    Debug.Assert wsTest1.AutoFilterMode = True
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True
    Debug.Assert wsTest3.ProtectContents = True
    mObstructions.CleanUp ' Let's see if something's still remaining. Investigation due in case!
    
    Test_ObstApplicationEvents
    Test_CellsMerging
    Test_ColHiding
    Test_Obstructions1
    Test_Obstructions2
    Test_ObstNamedRanges
    Test_RowsFiltering
    Test_ObstProtectedSheets
    Test_WsCustomView
    
xt: EoP ErrSrc(PROC)
    Debug.Assert mObstructions.WsColsHidden(wsTest1) = True
    Debug.Assert wsTest1.AutoFilterMode = True
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True
    Debug.Assert wsTest3.ProtectContents = True
    mObstructions.CleanUp ' Let's see if something's still remaining. Investigation due in case!
        
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_CellsMerging()
    Const PROC = "Test_CellsMerging"

    On Error GoTo eh
    BoP ErrSrc(PROC)

    ' still to be done !

xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_ColHiding()
    Const PROC = "Test_ColHiding"

    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    mObstructions.CleanUp bForce:=True ' Enforce remaining obstruction restores (without confirmation)
    TestSetUp
    
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest1
    mObstructions.ObstHiddenColumns xlSaveAndOff, wsTest1
    Debug.Assert wsTest1.TestColHidden.EntireColumn.Hidden = False
        '~~ Subsequent (nested) save/restore request (must not change the status)!
        mObstructions.ObstHiddenColumns xlSaveAndOff, wsTest1
        Debug.Assert wsTest1.TestColHidden.EntireColumn.Hidden = False
        mObstructions.ObstHiddenColumns xlRestore, wsTest1
        Debug.Assert wsTest1.TestColHidden.EntireColumn.Hidden = False
    mObstructions.ObstHiddenColumns xlRestore, wsTest1
    mObstructions.ObstProtectedSheets xlRestore, wsTest1 ' unprotected with hidden cols save and off but not restored ?
    Debug.Assert wsTest1.TestColHidden.EntireColumn.Hidden = True
    
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_ObstApplicationEvents()
' ------------------------------------------------------
' The test procedure will halt at any assertion not met.
' ------------------------------------------------------
    Const PROC = "Test_ObstApplicationEvents"

    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    Application.EnableEvents = True
    mObstructions.ObstApplicationEvents xlSaveAndOff
    Debug.Assert Application.EnableEvents = False
    
    '~~ Any subsequent SaveAndOff and CleanUp (usually in nested sub procedures) must not change the status
    mObstructions.ObstApplicationEvents xlSaveAndOff
    mObstructions.ObstApplicationEvents xlRestore
    Debug.Assert Application.EnableEvents = False
    mObstructions.ObstApplicationEvents xlSaveAndOff
    mObstructions.ObstApplicationEvents xlRestore
    Debug.Assert Application.EnableEvents = False
    
    '~~ The final CleanUp restores the initially saved status
    mObstructions.ObstApplicationEvents xlRestore
    Debug.Assert Application.EnableEvents = True
    
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_ObstNamedRanges()
' ------------------------------------------------------------------
' Range names imply a serious problem when a worksheet (wsSource)
' is about to be copied into another Workbook (wbTarget) since all
' names would refer back to the source Workbook (wbSource). To
' avoid this formulas using relevant range names are temporarily
' turned into comments and restored after the Worksheet had been
' copied.
' ------------------------------------------------------------------
    Const PROC = "Test_ObstNamedRanges"
    
    On Error GoTo eh
    mObstructions.CleanUp

    BoP ErrSrc(PROC)

    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest3
        mObstructions.ObstNamedRanges xlSaveAndOff, wsTest3
        mObstructions.ObstNamedRanges xlRestore, wsTest3
    mObstructions.ObstProtectedSheets xlRestore, wsTest3
    
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_ObstProtectedSheets()
' ------------------------------------------------------------
' Whichever number of sheets, protected or not is unprotected,
' when finally their protection status is restored it is like
' is was at the beginning.
' Assertions proof the correctness of this obstruction
' implementation.
' ------------------------------------------------------------
    Const PROC = "Test_ObstProtectedSheets"
        
    On Error GoTo eh
    
    TestSetUp
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = True
    BoP ErrSrc(PROC)
    
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest1
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest2
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest3
    
    '~~ Any subsequent xlSaveAndOff and xlRestore (usually in nested sub-procedures)
    '~~ must not have any effect on the final result
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest1
    mObstructions.ObstProtectedSheets xlRestore, wsTest1
    
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest1
    mObstructions.ObstProtectedSheets xlRestore, wsTest1
    
    '~~ Assert all sheets are unprotectec
    Debug.Assert wsTest1.ProtectContents = False
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = False
    
    mObstructions.ObstProtectedSheets xlRestore, wsTest1
    mObstructions.ObstProtectedSheets xlRestore, wsTest2
    mObstructions.ObstProtectedSheets xlRestore, wsTest3
    
    '~~ Assert only those initially protected are protected again
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = True

xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_Obstructions1()
    Const PROC = "Test_Obstructions1"

    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    '~~ 1. Test local approach
    mObstructions.Obstructions xlSaveAndOff, ActiveSheet
    ' any "elementary" operation e.g. rows copy
    mObstructions.Obstructions xlRestore, ActiveSheet

xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_Obstructions2()
    Const PROC = "Test_Obstructions2"

    On Error GoTo eh
    Dim cv  As CustomView
    Dim dct As Dictionary
    
    BoP ErrSrc(PROC)
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_RowsFiltering()
' --------------------------------------------------------------------
' Assertions proof the correctness of this obstruction implementation.
' Obstructions SaveAndOff and CleanUp allways have to be pairs.
' Unpaired Save/CleanUp will lead to incorrect results!!!!
' --------------------------------------------------------------------
Const PROC = "Test_RowsFiltering"
Dim cv      As CustomView

    '~~  Setup test sheets and assert the required initial test status
    mObstructions.CleanUp True
    TestSetUp
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    '~~ Note! Obstructions are saved-and-turned and restored Worksheetwise
    wsTest1.Activate
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    mObstructions.ObstFilteredRows xlSaveAndOff, wsTest1
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlRestore, wsTest1
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    
    wsTest2.Activate
    mObstructions.ObstFilteredRows xlSaveAndOff, wsTest2
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlRestore, wsTest2
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    
    wsTest3.Activate
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlSaveAndOff, wsTest3
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = False:    Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlRestore, wsTest3
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
        
    '~~ Note! Obstructions are handled Worksheet by Worksheet
    mObstructions.ObstFilteredRows xlSaveAndOff, wsTest1
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlSaveAndOff, wsTest2
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlSaveAndOff, wsTest3
    Debug.Assert wsTest1.AutoFilterMode = False:    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = False:    Debug.Assert wsTest3.ProtectContents = True
    
    mObstructions.ObstFilteredRows xlRestore, wsTest1
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.AutoFilterMode = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = False:    Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlRestore, wsTest2
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = False:    Debug.Assert wsTest3.ProtectContents = True
    mObstructions.ObstFilteredRows xlRestore, wsTest3
    Debug.Assert wsTest1.AutoFilterMode = True:     Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False:    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True:     Debug.Assert wsTest3.ProtectContents = True
        
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Test_WsCustomView()
Const PROC  As String = "Test_WsCustomView"
    
    '~~  Setup test sheets and assert the required initial status
    TestSetUp
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest1.AutoFilterMode = True
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest2.AutoFilterMode = False
    Debug.Assert wsTest3.ProtectContents = True
    Debug.Assert wsTest3.AutoFilterMode = True
    Debug.Assert Application.EnableEvents = True
    mObstructions.CleanUp
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest1
    mObstructions.WsCustomView xlSaveOnly, wsTest1, bRowsFiltered:=wsTest1.AutoFilterMode
    wsTest1.AutoFilterMode = False
        mObstructions.WsCustomView xlSaveOnly, wsTest1, bRowsFiltered:=wsTest1.AutoFilterMode
        wsTest1.AutoFilterMode = False
        Debug.Assert wsTest1.AutoFilterMode = False
        mObstructions.WsCustomView xlRestore, wsTest1
        Debug.Assert wsTest1.AutoFilterMode = False
    mObstructions.WsCustomView xlRestore, wsTest1
    
    Debug.Assert wsTest1.AutoFilterMode = True
    mObstructions.ObstProtectedSheets xlRestore, wsTest1
    
xt: EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
eh: If mErH.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

