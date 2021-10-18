Attribute VB_Name = "mTest"
Option Explicit

Public Sub Test_All()
' --------------------------
' Unatended regression test.
' All results asserted
' --------------------------
Const PROC  As String = "Test_All"
    
    On Error GoTo on_error
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
    
exit_proc:
    EoP ErrSrc(PROC)
    Debug.Assert mObstructions.WsColsHidden(wsTest1) = True
    Debug.Assert wsTest1.AutoFilterMode = True
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.AutoFilterMode = False
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.AutoFilterMode = True
    Debug.Assert wsTest3.ProtectContents = True
    mObstructions.CleanUp ' Let's see if something's still remaining. Investigation due in case!
Exit Sub
    
on_error:
    mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub Test_Obstructions1()
Const PROC = "Test_Obstructions1"

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    '~~ 1. Test local approach
    mObstructions.Obstructions xlSaveAndOff, ActiveSheet
    ' any "elementary" operation e.g. rows copy
    mObstructions.Obstructions xlRestore, ActiveSheet

exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
    mErH.ErrMsg ErrSrc(PROC)
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
    
    mObstructions.CleanUp

    On Error GoTo on_error
    BoP ErrSrc(PROC)

    mObstructions.ObstProtectedSheets xlSaveAndOff, wsTest3
        mObstructions.ObstNamedRanges xlSaveAndOff, wsTest3
        mObstructions.ObstNamedRanges xlRestore, wsTest3
    mObstructions.ObstProtectedSheets xlRestore, wsTest3
    
exitProc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    mErH.ErrMsg ErrSrc(PROC)
End Sub


Public Sub Test_Obstructions2()
Const PROC = "Test_Obstructions2"
Dim cv  As CustomView
Dim dct As Dictionary

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlSaveAndOff, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    mObstructions.ObstAll obs_mode:=xlRestore, obs_ws:=wsTest1
    
exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
    mErH.ErrMsg ErrSrc(PROC)
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
    
    On Error GoTo on_error
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
    
exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
    mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub Test_ObstApplicationEvents()
' ------------------------------------------------------
' The test procedure will halt at any assertion not met.
' ------------------------------------------------------
Const PROC = "Test_ObstApplicationEvents"

    On Error GoTo on_error
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
    
exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
    mErH.ErrMsg ErrSrc(PROC)

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
    
    TestSetUp
    Debug.Assert wsTest1.ProtectContents = True
    Debug.Assert wsTest2.ProtectContents = False
    Debug.Assert wsTest3.ProtectContents = True
        
    On Error GoTo on_error
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

exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    mErH.ErrMsg ErrSrc(PROC)
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
    
    On Error GoTo on_error
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
        
exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub Test_ColHiding()
Const PROC = "Test_ColHiding"

    On Error GoTo on_error
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
    
exitProc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub Test_CellsMerging()
Const PROC = "Test_CellsMerging"

    On Error GoTo on_error
    BoP ErrSrc(PROC)

    ' still to be done !

exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    mErH.ErrMsg ErrSrc(PROC)
End Sub

Public Sub Names()
Const PROC = "Names"
Dim nm      As Name
Dim row     As Long
Dim wb      As Workbook
Dim ws      As Worksheet

    Set wb = ThisWorkbook
    Set ws = wsTest1
    
    With wsTest1
        .rNames.ClearContents
        row = .rNames.row - 1
        For Each nm In wb.Names
            Debug.Print nm.RefersTo
            row = row + 1
            Intersect(.rNames, .NamesSheet.EntireColumn, .Rows(row)).Value = Split(nm.RefersTo, "!")(0)
            Intersect(.rNames, .NamesReference.EntireColumn, .Rows(row)).Value = Split(nm.RefersTo, "!")(1)
            If InStr(nm.Name, "!") <> 0 Then
                Intersect(.rNames, .NamesName.EntireColumn, .Rows(row)).Value = Split(nm.Name, "!")(1)
            Else
                Intersect(.rNames, .NamesName.EntireColumn, .Rows(row)).Value = nm.Name
            End If
            Intersect(.rNames, .NamesScope.EntireColumn, .Rows(row)).Value = "Workbook"
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
    
exit_proc:
    EoP ErrSrc(PROC)
    mObstructions.CleanUp ' check if something to retore remained
    Exit Sub
    
on_error:
    mErH.ErrMsg ErrSrc(PROC)
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

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mTest" & "." & sProc
End Function
