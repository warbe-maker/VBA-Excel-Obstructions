Attribute VB_Name = "mObstructions"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mObstructions
'           Manages obstructions which hinder vba operations by
'           providing procedured to save and set them off and
'           restoring them. Typical operations prevented are:
'           - Rows move/copy (e.g. by filtered rows)
'           - Range value modifications
' Procedures:
' - Obstructions            summarizes all below and turns them off and
'                           retores them
' - ObstApplicationEvents   Sets it to False and restores the initially
'                           saved status
' - ObstFilteredRows        Turns Autofilter off when active and restores
'                           it by means of a CustomView.
' - ObstHiddenColumns       Displays them and restores them by means of a
'                           CustomView
' - ObstMergedCells         xlSaveAndOff Un-merges, xlRestore re-merges cells
'                           associated with the current Selection.
' - ObstProtectedSheets     Un-protects any number of sheets used in a project
'                           and re-protects them (only) when they
'                           were initially protected.
' - ObstNamedRanges         Saves and restores all formulas in Workbook
'                           which use a RangeName of a certain Worksheet
'                           by commenting and uncommenting the formulas
' ------------------------------------------------------------------------------
' Note 1: A CustomView is the means to restore Autofilter and/or hidden columns
'         of a certain Worksheet. CustomViews to save/restore Autofilter a
'         independant from those for saving/restoring hidden columns.
' Note 2: In order to make an "elementary" operation like "copy row" for
'         instance independent from the environment the statements will be
'         enclosed in Obstructions xlSaveAndOff and Obstructions xlRestore.
'         However, if such an "elemetary" operation is just one amongst others,
'         performed by a more complex operation this one as well should start
'         with Obstructions xlSaveAndOff and end with Obstructions xlRestore
'         but additionally provided with a "global" Dictionary and a "global"
'         CustomView object. The Obstructions call with the "elementary"
'         operation will thus not conflict with the "global" Off/CleanUp since
'         there is no longer an obstruction to turn off and subsequently none
'         to be restored.
'
' Uses the common components:
' mCstmVw, m-Wrkbk (the common components mErH, fMsg, mMsg are only used by the
' mTest module and thus are not required when using mObstructions)
'
' ------------------------------------------------------------------------------
Public Const TEMPCVNAME    As String = "TempObstructionsCustomView_"
Public Enum xlSaveRestore
    xlSaveOnly
    xlSaveAndOff
    xlRestore
    xlOffOnly
    xlOnOnly
End Enum
Private i               As Long
Private dcProt          As Dictionary
Private cllAppEvents    As Collection
Private dcCvsWb         As Dictionary
Private dcMerged        As Dictionary
Private wb              As Workbook
Private bObstructionHiddenCols      As Boolean
Private bObstructionFilteredRows    As Boolean

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Public Sub ObstApplicationEvents(ByVal ae_operation As xlSaveRestore)
' ------------------------------------------------------------------------------
' - ae_operation = xlSaveAndOff
'   Saves the current Application.EnableEvents status and turns it off. Any
'   subsequent execution will just save the status (i.e. adds it to the stack)
' - ae_operation = xlRestore
'   Restores the last saved Application.EnableEvents status and removes the
'   saved item.
' ------------------------------------------------------------------------------
    Const PROC = "ObstApplicationEvents"

    On Error GoTo eh
    
    Select Case ae_operation
        Case xlSaveAndOff
            If cllAppEvents Is Nothing Then Set cllAppEvents = New Collection
            cllAppEvents.Add Application.EnableEvents ' add status to stack
            Application.EnableEvents = False
            
        Case xlRestore
            If cllAppEvents Is Nothing Then GoTo xt
            With cllAppEvents
                If .Count > 0 Then
                    '~~ restore last saved statis item and remove item (take it off from stack)
                    Application.EnableEvents = .Item(cllAppEvents.Count)
                    .Remove .Count
                End If
            End With
    End Select

xt: Exit Sub
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub Borders(ByVal r As Range, _
                   ByVal SaveRestore As xlSaveRestore, _
                   ByRef dct As Dictionary)
' ------------------------------------------------------------------------------
' - xlOffOn = xlSaveOnly: Saves the border properties of
'   Range r into the Dictionary dct.
' - xlOnOff = xlRestore: Restores all border properties
'   for Range r from the Dictionary dct.
' ------------------------------------------------------------------------------
    Const PROC = "Borders"
    
    On Error GoTo eh
    Dim cll             As Collection
    Dim xlBi            As XlBordersIndex

    i = 0
    
    Select Case SaveRestore
        Case xlSaveOnly
            Set dct = Nothing
            Set dct = New Dictionary
            For xlBi = xlDiagonalDown To xlInsideHorizontal ' = 5 to 12
                With r.Borders(xlBi)
                    Set cll = New Collection
                    cll.Add .LineStyle
                    cll.Add .Weight
                    '~~ If there is a Color or a ColorIndex there is no ThemeColor
                    cll.Add .Color
                    cll.Add .ColorIndex
                    
                    If CStr(.Color) = vbNullString And CStr(.ColorIndex) = vbNullString Then
                        cll.Add .ThemeColor
                    Else
                        cll.Add Null
                    End If
                    cll.Add .TintAndShade
                End With
                dct.Add xlBi, cll
            Next xlBi
            
        Case xlRestore
            For i = 0 To dct.Count - 1
                xlBi = dct.Keys(i)
                Set cll = dct.Items(i)
                
                With r.Borders(xlBi)
                    .LineStyle = cll.Item(1)
                    If Not .LineStyle = xlNone Then
                        '~~ Any other border formationg only when there is a border
                        .Weight = cll.Item(2)
                        If Not IsNull(cll.Item(5)) Then
                            .ThemeColor = cll.Item(5)
                        Else
                            '~~ Any color only when there is not ThemeColor
                            .Color = cll.Item(3)
                            .ColorIndex = cll.Item(4)
                        End If
                        If Not IsNull(cll.Item(6)) Then
                            .TintAndShade = cll.Item(6)
                        End If
                    End If  ' LineStyle not is xlNone
                End With
                
            Next i
    End Select
    
    Exit Sub

eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub CleanUp(Optional ByVal bForce As Boolean = False)
' ------------------------------------------------------------------------------
' Does cleanup for all obstructions still waiting for
' a CleanUp, likely not done due to an error.
' When bForce is True all remaining Restores are done
' without notice. This may be used in the error handling of
' the project after the error message had been displayed.
' ------------------------------------------------------------------------------
    Dim cv      As CustomView
    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim v1      As Variant
    Dim v2      As Variant
    Dim cll     As Collection
    Dim sMsg    As String
    Dim dcCvsWs As Dictionary

    '~~ CleanUp CustomViews
    If Not dcCvsWb Is Nothing Then
        For Each v1 In dcCvsWb
            Set dcCvsWs = dcCvsWb.Item(v1)
            For Each v2 In dcCvsWs
                Set ws = v2
                Set cll = dcCvsWs.Item(v2)
                If cll.Count > 0 Then
                    '~~ Remaining restore action!
                    If bForce Then
                        While cll.Count > 0: WsCustomView xlRestore, ws: Wend
                    Else
                        If TypeName(cll.Item(1)) = "Object" Then
                            Set cv = cll.Item(1)
                            If CstmVwExists(wb, cv) Then
                                If MsgBox(" Restore the CustomView of Worksheet yet unrestored '" & ws.Name & "' ?", vbYesNo, "Unrestored CustomView") = vbYes Then
                                    While cll.Count > 0: WsCustomView xlRestore, ws: Wend
                                End If
                            End If
                        End If
                    End If
                End If
            Next v2
        Next v1
        dcCvsWb.RemoveAll
    End If
    
    '~~ CleanUp Protection
    If Not dcProt Is Nothing Then
        If dcProt.Count > 0 Then
            For Each v1 In dcProt
                Set ws = v1
                Set cll = dcProt.Item(v1)
                With cll
                    If .Count > 0 Then
                        '~~ Remaining protection restore
                        If bForce Then
                            If .Item(1) Then ws.Protect Else ws.Unprotect
                        Else
                            If MsgBox("CleanUp the yet unrestored protection status for Worksheet '" & ws.Name & "' ?", vbYesNo, "Unrestored Protection Satus") = vbYes Then
                                If .Item(1) Then ws.Protect Else ws.Unprotect
                            End If
                        End If
                    End If
                End With
            Next v1
        End If
        dcProt.RemoveAll
    End If
    
    '~~ CleanUp ObstApplicationEvents (allway done silent)
    If Not cllAppEvents Is Nothing Then
        With cllAppEvents
            If .Count > 0 Then Application.EnableEvents = .Item(1)
        End With
        Set cllAppEvents = Nothing
    End If
    
    If Not dcMerged Is Nothing Then
        '~~ Merge anything yet not re-merged
    End If
    
    If sMsg <> vbNullString Then
        MsgBox "Attention!" & vbLf & sMsg, vbCritical, "CleanUp of saved and set off bbstructions incomplete!"
    End If

End Sub

Public Function CstmVwExists(ByVal vWb As Variant, _
                             ByVal vCv As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the CustomView (vCv) - which may be a CustomView object or a
' CustoView's name - exists in the Workbook (vwb). If vCv is provided as a
' CustomView object only its name is used for the existence check.
' ------------------------------------------------------------------------------
    Const PROC  As String = "CustomViewExists"      ' This procedure's name for the error handling and execution tracking
    On Error GoTo eh
    
    Dim wb      As Workbook
    Dim sTest   As String

    CstmVwExists = False
    
    If Not mWrkbk.IsObject(vWb) And Not mWrkbk.IsFullName(vWb) And Not mWrkbk.IsName(vWb) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The Workbook (parameter vWb) is neither a Workbook object nor a Workbook's name or fullname)!"
    
    If mWrkbk.IsObject(vWb) Then
        Set wb = vWb
    ElseIf mWrkbk.IsFullName(vWb) Then
        Set wb = mWrkbk.GetOpen(vWb)
    ElseIf mWrkbk.IsName(vWb) Then
        If Not mWrkbk.IsOpen(vWb, wb) _
        Then Err.Raise AppErr(2), ErrSrc(PROC), "The provided Workbook (vWb) '" & vWb & "' is not open!"
    End If
    
    If Not IsCstmVwObject(vCv) And Not IsCstmVwName(vCv) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The CustomView (vCv) is neither a string (CustomView's name) nor a CustomView object!"
    
    If IsCstmVwObject(vCv) Then
        On Error Resume Next
        sTest = vCv.Name
        CstmVwExists = Err.Number = 0
        GoTo xt
    ElseIf IsCstmVwName(vCv) Then
        On Error Resume Next
        sTest = wb.CustomViews(vCv).Name
        CstmVwExists = Err.Number = 0
        GoTo xt
    End If
  
xt: Exit Function
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Function

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString, _
    Optional ByVal err_line As Long = 0)
' ------------------------------------------------------------------------------
' This 'Common VBA Component' uses only a kind of minimum error handling!
' ------------------------------------------------------------------------------
    Dim ErrNo   As Long
    Dim ErrDesc As String
    Dim ErrType As String
    Dim errline As Long
    Dim AtLine  As String
    Dim ErrMsg  As String
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Applicatin error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    If err_dscrptn = vbNullString Then ErrDesc = Err.Description Else ErrDesc = err_dscrptn
    If err_line = 0 Then errline = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    MsgBox Title:=ErrType & ErrNo & " in " & err_source _
         , Prompt:="Error : " & ErrDesc & vbLf & _
                   "Source: " & err_source & AtLine _
         , Buttons:=vbCritical
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mObstructions" & "." & sProc
End Function

Public Sub ObstFilteredRows(ByVal fr_operation As xlSaveRestore, _
                            ByVal fr_ws As Worksheet)
' ------------------------------------------------------------------------------
' - fr_operation = xlSaveAndOff
'   When Autofilter is active a temporary CustomView is created and AutoFilter
'   is turned off.
' - fr_operation = xlRestore
'   Returns to the temporary created CustomView if any and thus restores the
'   Autofilter with all its initial specifications.
' Note:
' - Save/Restore requests may be nested but it is absolutely essential that they
'   are paired!
' - Subsequent SaveAndOff request (e.g. in nested subprocedures) are just
'   "stacked", subsequent Restore requests are just un-stacked - and thus do not
'   cause any problem - provided they are paired.
' - Filtered rows have to be turned off by Worksheet
' - Worksheet's obstructions may be restored in any order order! (Save wsTest1,
'   wsTest2, Restore wsTest2, wsTest1)
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin Dec 2019
' ------------------------------------------------------------------------------
    Const PROC = "ObstFilteredRows"

    On Error GoTo eh

    '~~ The Workbook of the Worksheet may not be found within this Application instance.
    '~~ Application.Workbooks() may thus not be appropriate. GetOpenWorkbook will find
    '~~ it in whichever Application instance
    Set wb = mWrkbk.GetOpen(fr_ws.Parent.Name)
    If dcCvsWb Is Nothing Then Set dcCvsWb = New Dictionary
    
    Select Case fr_operation
        Case xlSaveAndOff
            If fr_ws.AutoFilterMode = True Then
                '~~ Create a CustomView, keep a record of the CustomView and turn filtering off
                WsCustomView xlSaveOnly, fr_ws, bRowsFiltered:=True
                ObstProtectedSheets xlSaveAndOff, fr_ws    ' Possibly nested request ensuring unprotection
                fr_ws.AutoFilterMode = False
                ObstProtectedSheets xlRestore, fr_ws       ' Possibly nested restore ensuring protection status restore
            Else
                WsCustomView xlSaveOnly, fr_ws ' Just add subsequent save request to stack
            End If
        
        Case xlRestore
            '~~ CleanUp the CustomView saved for Worksheet (fr_ws) if any
            If dcCvsWb.Exists(wb) Then ' Only if at least for one Worksheet a CustomView had been saved
                WsCustomView xlRestore, fr_ws
            End If
    End Select

xt: Exit Sub

eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub ObstHiddenColumns(ByVal hc_operation As xlSaveRestore, _
                             ByVal hc_ws As Worksheet)
' ------------------------------------------------------------------------------
' - hc_operation = xlSaveAndOff
'   Create/Save CustomView and display all hidden columns
' - hc_operation = xlRestore
'   CleanUp the saved CustomView.
' Note: May be called in nested subprocedures without a problem.
' ------------------------------------------------------------------------------
    Const PROC      As String = "ObstHiddenColumns"
    
    On Error GoTo eh
    Dim col         As Range
    
    Set wb = mWrkbk.GetOpen(hc_ws.Parent.Name)
    
    Select Case hc_operation
        
        Case xlSaveAndOff
            If WsColsHidden(hc_ws) Then
                '~~ If not one already exists create a CustomView for this Worksheet and keep a record of it
                '~~ and un-hide all hidden columns
                ObstProtectedSheets xlSaveAndOff, hc_ws
                WsCustomView xlSaveOnly, hc_ws, bColsHidden:=True
                For Each col In hc_ws.UsedRange.Columns
                    If col.Hidden Then col.Hidden = False
                Next col
                ObstProtectedSheets xlRestore, hc_ws
            Else
                '~~ Add a subsequent save request to the Worksheet's save stack
                WsCustomView xlSaveOnly, hc_ws, bColsHidden:=True
            End If
        
        Case xlRestore
            WsCustomView xlRestore, hc_ws
    End Select
    
xt: Exit Sub
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Function IsCstmVwName(ByVal v As Variant) As Boolean
    IsCstmVwName = VarType(v) = vbString
End Function

Public Function IsCstmVwObject(ByVal v As Variant) As Boolean

    If VarType(v) = vbObject Then
        If Not TypeName(v) = "Nothing" Then
            IsCstmVwObject = TypeOf v Is CustomView
        End If
    End If
    
End Function

Private Sub Merge(ByVal r As Range, ByVal OffOn As xlSaveRestore)
' ------------------------------------------------------------------------------
' OffOn = xlSaveAndOff: un-merges range r by copying the top left
'                       content to all cells in the merge area.
' OffOn = xlRestore: Re-merges range r.
' ------------------------------------------------------------------------------
    Const PROC As String = "Merge"
    
    On Error GoTo eh
    Dim rSel    As Range
    Dim cel     As Range
    Dim bEvents As Boolean
    Dim row     As Range
    
    With Application
        bEvents = .EnableEvents
        .EnableEvents = False
        .ScreenUpdating = False
        
        If OffOn = xlOffOnly Then
            Set rSel = Selection
            With r
                '~~ Avoid automatic rows height adjustment by setting it explicit
                For Each row In r.Rows
                    With row: .RowHeight = .RowHeight + 0.01: End With
                Next row
                .UnMerge
                '~~ In order not to loose the value in the top left cell
                '~~ e.g. when a row is deleted or moved up or down
                '~~ it is copied to all other cells. Merge will return to
                '~~ the then top left cell's content.
                r.Cells(1, 1).Copy
                For Each cel In r.Cells
                    If cel.Value = vbNullString Then
                        cel.PasteSpecial Paste:=xlPasteAllExceptBorders, _
                                     Operation:=xlNone, _
                                    SkipBlanks:=False, _
                                     Transpose:=False
                     End If
                Next cel
                    
                rSel.Select
                Application.CutCopyMode = False
            End With ' r
            
        ElseIf OffOn = xlOnOnly Then
            '~~ Reset to original row height
            For Each row In r.Rows
                On Error Resume Next
                With row: .RowHeight = .RowHeight - 0.01: End With
            Next row
            .DisplayAlerts = False
            r.Merge
            .DisplayAlerts = True
                
        End If
        .EnableEvents = bEvents
    End With ' application

    Exit Sub
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub ObstMergedCells(ByVal mc_operation As xlSaveRestore, _
                  Optional ByRef mc_global As Variant = Null)
' ------------------------------------------------------------------------------
' - mc_operation = xlSaveAndOff
'   Any merge area associated with a in the current Selection is un-merged by
'   saving the merge areas' address in a temporary range name and additionally
'   in a Dictionary. The content of the top left cell is copied to all cells in
'   the un-merged area to prevent a loss of the merge  area's content even when
'   the top row of it is deleted. The named ranges address is automatically
'   maintained by Excel throughout any rows operations performed  within the
'   originally merged area's top and bottom row. I.e. any row copied or
'   inserted above the top row or below the bottom row will not become part of
'   the retored merge area(s).
' - mc_operation = xlRestore:
'   All merge areas registered by a temporary range name are re-merged, thereby
'   eliminating all duplicated content except the one in the top left cell. When
'   no merge areas are detected neither of the ObstMergedCells call does anything.
'   I.e. no need to check for any merged cells beforehand.
'
' Used by Obstructions, uses Merge
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin June 2019
' ------------------------------------------------------------------------------
    Const PROC     As String = "ObstMergedCells"
    Const sTempName As String = "rngTempMergeAreaName"
    On Error GoTo eh
    
    Static dcLocl   As Dictionary
    Dim dcBorders   As Dictionary
    Dim dc          As Dictionary
    Dim r           As Range
    Dim i           As Long
    Dim k           As Long
    Dim cel         As Range
    Dim sMergeArea  As String
    Dim sName       As String
    Dim vKey        As Variant

    '~~ When provided, the global dictionary is used, else the local
    If Not IsNull(mc_global) Then
        If mc_global Is Nothing Then
            Set mc_global = New Dictionary
        End If
        Set dc = mc_global
    Else
        If dcLocl Is Nothing Then
            Set dcLocl = New Dictionary
        End If
        Set dc = dcLocl
    End If
    
    If mc_operation = xlSaveAndOff Then
        Set r = Selection
        For Each cel In Intersect(r.Worksheet.UsedRange, r.EntireRow).Cells
            With cel
                If .MergeCells Then
                    i = i + 1:  sName = sTempName & i
                    sMergeArea = Replace(.MergeArea.Address(RowAbsolute:=False), "$", vbNullString)
                    If Not dc.Exists(sName) Then
                        '~~ When the added range names are 0
                        '~~ remove any outdated range names beforehand
                        k = 0
                        For Each vKey In dc.Keys
                            If TypeName(vKey) = "String" Then k = k + 1 ' If there are also other things in the dictionary
                        Next vKey
                        If k = 0 Then
                            NamesRemove sTempName, r.Worksheet, False
                        End If
                        Set r = Range(sMergeArea)
                        Application.Names.Add sName, r
                        Borders r, xlSaveOnly, dcBorders   ' Get format properties
                        dc.Add sName, dcBorders            ' Save range name and border properties to dictionary
                        Merge r, xlOffOnly
                    End If
                End If
            End With
        Next cel
            
    ElseIf mc_operation = xlRestore Then
        For i = dc.Count - 1 To 0 Step -1               ' bottom up to allow removing processed items
            If TypeName(dc.Keys(i)) = "String" Then     ' A string typ key indicates a merge area's range name
                sName = dc.Keys(i)                      ' Get range name
                Set dcBorders = dc.Items(i)             ' Get border properties
                Set r = Range(sName)                    ' Set the to-be-merged range
                Merge r, xlOnOnly                       ' Merge the range named sName
                Borders r, xlRestore, dcBorders         ' CleanUp the border proprties
                NamesRemove sName, r.Worksheet, False   ' Remove the no longer required range name
                dc.Remove sName                         ' Remove the no longer required item from the Dictionary
            End If
        Next i
    End If

xt: Exit Sub
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Private Sub NamesRemove(ByVal sName As String, _
               Optional ByVal ws As Worksheet = Nothing, _
               Optional ByVal bConfirm As Boolean = True)
' ------------------------------------------------------------------------------
' Removes the name sName from the list of Names provided
' the name refers to ws which defaults to the ActiveSheet.
' When there is no sName found sName is regarded a generic
' part of names and all names with it are removed.
' ------------------------------------------------------------------------------
    Dim nm          As Name
    Dim sConfirm    As String
    Dim bConfirmed  As Boolean
    Dim sWsName     As String

    If ws Is Nothing Then Set ws = ActiveSheet
    sWsName = "='" & ws.Name & "'!"
    If bConfirm = False Then bConfirmed = True Else bConfirmed = False
    
    For Each nm In Application.Names
        If Left(nm.RefersTo, Len(sWsName)) = sWsName Then
            If nm.Name = sName Then
                '~~ If name is unique delete it right away
                nm.Delete
                GoTo xt
            End If
        End If
    Next nm
    
    '~~ Regard sName as a generic name string
again_confirmed:
    For Each nm In Application.Names
        If Left(nm.RefersTo, Len(sWsName)) = sWsName Then
            '~~ Refers to the rquested sheet
            If Left(nm.Name, Len(sName)) = sName Then
                '~~ Is one of the generic names
                If bConfirmed Then
                    nm.Delete
                Else
                    sConfirm = sConfirm & vbLf & "'" & nm.Name & "'"
                End If
            End If
        End If
    Next nm
    If bConfirmed Then GoTo xt
    If bConfirm And sConfirm <> vbNullString Then
        If MsgBox("Yes if the following renames are to be removed: " & sConfirm, vbYesNo, "Confirm removals") = vbYes Then
            bConfirmed = True
            GoTo again_confirmed:
        Else
            GoTo xt
        End If
    End If

xt:
End Sub

Public Sub All(ByVal mode As xlSaveRestore)
    mObstructions.Obstructions obs_operation:=mode _
                             , obs_ws:=wsTest1 _
                             , obs_application_events:=True _
                             , obs_protected_sheets:=True _
                             , obs_filtered_rows:=True _
                             , obs_hidden_columns:=True _
                             , obs_merged_cells:=True

End Sub

Public Sub Obstructions(ByVal obs_operation As xlSaveRestore, _
                        ByVal obs_ws As Worksheet, _
               Optional ByVal obs_protected_sheets As Boolean = False, _
               Optional ByVal obs_filtered_rows As Boolean = False, _
               Optional ByVal obs_hidden_columns As Boolean = False, _
               Optional ByVal obs_merged_cells As Boolean = False, _
               Optional ByVal obs_named_ranges As Boolean = False, _
               Optional ByVal obs_application_events As Boolean = False, _
               Optional ByVal obs_form_events As Boolean = False)
' --------------------------------------------------------------------
' Saves and restores all obstructions indicated True. It is absolutely
' essential that any Obstructions Save is paired by an exactly corres-
' ponding CleanUp. Nested Save/CleanUp pairs, usually performed in
' nested sub-procedures (which allows independant testing) is fully
' suported. The sequence in which paired Restores are perfomed is not
' relevant as long as they are exactly paired.
'
' Requires: - Reference to "Microsoft Scripting Runtime"
'           - Module mErrHndlr
'           - Module mExists
' --------------------------------------------------------------------
Const PROC          As String = "Obstructions"
    
    On Error GoTo eh
           
    bObstructionHiddenCols = obs_hidden_columns
    bObstructionFilteredRows = obs_filtered_rows

    Select Case obs_operation
        Case xlSaveAndOff
            
            '~~ 1. Save and turn off Application Events
            If obs_application_events Then
                ObstApplicationEvents xlSaveAndOff
            End If
            
            '~~ 2. Save and turn off sheet protection if requested or implicitely required
            If obs_protected_sheets Or obs_filtered_rows Or obs_hidden_columns Or obs_merged_cells Then
                ObstProtectedSheets xlSaveAndOff, obs_ws
            End If
            
            '~~ 3. Save and turn off Autofilter  if applicable
            If bObstructionFilteredRows Then
                ObstFilteredRows xlSaveAndOff, obs_ws
            End If
            
            '~~ 4. Save and turn off hidden columns  if applicable
            If bObstructionHiddenCols Then
                ObstHiddenColumns xlSaveAndOff, obs_ws
            End If
            
            '~~ 5. Save and turn off any effected merged cells if applicable
            If obs_merged_cells Then
                ObstMergedCells xlSaveAndOff
            End If
            
            '~~ 6. Save and turn off used range names
            If obs_named_ranges Then
                ObstNamedRanges xlSaveAndOff, obs_ws
            End If
            
        Case xlRestore
            
            '~~ 1. CleanUp all formulas using a ws range name
            If obs_named_ranges Then
                ObstNamedRanges xlRestore, obs_ws
            End If
            
            '~~ 2. CleanUp merge areas which were initially effected by the Selection
            If obs_merged_cells Then
                ObstMergedCells xlRestore
            End If
            
            '~~ 3. CleanUp Autofilter if applicable
            If obs_filtered_rows Then
                ObstFilteredRows xlRestore, obs_ws
            End If
            
            '~~ 4. CleanUp hidden columns
            If obs_hidden_columns Then
                ObstHiddenColumns xlRestore, obs_ws
            End If
            
            '~~ 5. CleanUp the sheets protection status when it initially was protected
            If obs_protected_sheets Or obs_filtered_rows Or obs_hidden_columns Or obs_merged_cells Then
                ObstProtectedSheets xlRestore, obs_ws
            End If
            
            '~~ 6. CleanUp Application Events status
            If obs_application_events Then
                ObstApplicationEvents xlRestore
            End If
            
    End Select
         
xt: Exit Sub
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub ObstNamedRanges(ByVal nr_operation As xlSaveRestore, _
                           ByVal nr_ws As Worksheet)
' ------------------------------------------------------------------------------
' - nr_operation = xlSaveAndOff
'   All references to a name in the source Worksheet (nr_ws) are turned into a
'   direct range reference.
' Note:
' Names are regarded an obstruction when the Worksheet (nr_ws) is to be copied
' or moved from one Workbook to another. Since the names for ranges in the
' source Workshhet are not copied to the target Workbook all names are subject
' to a decision stay or another decision.
' ------------------------------------------------------------------------------
    Const PROC      As String = "ObstNamedRanges"
    
    On Error GoTo eh
    Dim dcNames     As Dictionary   ' Names which refer to a range in the Worksheet nr_ws
    Dim dcCells     As Dictionary   ' Keeps a record of each cell's formula which had been modified
    Dim nm          As Name
    Dim wsheet      As Worksheet
    Dim cel         As Range
    Dim v           As Variant
    Dim wb          As Workbook

    '~~ The Workbook of the Worksheet may not be found within this Application instance.
    '~~ Application.Workbooks() may thus not be appropriate. GetOpenWorkbook will find
    '~~ it in whichever Application instance
    Set wb = mWrkbk.GetOpen(nr_ws.Parent.Name)
    
    Select Case nr_operation
        Case xlSaveAndOff
            '~~ Collect the relevant names
            If dcNames Is Nothing Then Set dcNames = New Dictionary Else dcNames.RemoveAll
            For Each nm In Application.Names
'                Debug.Print "Collecting Name '" & nm.Name & "'"
'                Debug.Print nm.NameLocal
'                Debug.Print nm.RefersTo
'                Debug.Print nm.RefersToLocal
'                Debug.Print nm.RefersToR1C1
'                Debug.Print nm.RefersToR1C1Local
'                Debug.Print nm.RefersToRange.Address
                dcNames.Add nm.Name, nm
            Next nm
            For Each nm In wb.Names
'                Debug.Print "Collecting Name '" & nm.Name & "'"
'                Debug.Print nm.NameLocal
'                Debug.Print nm.RefersTo
'                Debug.Print nm.RefersToLocal
'                Debug.Print nm.RefersToR1C1
'                Debug.Print nm.RefersToR1C1Local
'                Debug.Print nm.RefersToRange.Address
                On Error Resume Next
                dcNames.Add nm.Name, nm
            Next nm
            For Each nm In nr_ws.Names
'                Debug.Print "Collection Range Name '" & nm.Name & "'"
'                Debug.Print nm.NameLocal
'                Debug.Print nm.RefersTo
'                Debug.Print nm.RefersToLocal
'                Debug.Print nm.RefersToR1C1
'                Debug.Print nm.RefersToR1C1Local
'                Debug.Print nm.RefersToRange.Address
                On Error Resume Next
                dcNames.Add nm.Name, nm
            Next nm
            '~~ Collect cells with a formula referencing on of the collected names
            For Each wsheet In wb.Sheets
                For Each cel In wsheet.UsedRange.SpecialCells(xlCellTypeFormulas).HasFormula
                    With cel
                        For Each v In dcNames
                            If InStr(.Formula, v) <> 0 Then
'                                Debug.Print cel.Address & " formula uses '" & v & "'"
                                dcCells.Add cel, .Formula ' At least one name is used
                            End If
                        Next v
                    End With
                Next cel
            Next wsheet
        
        Case xlRestore
        
    End Select

xt: Exit Sub
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub ObstProtectedSheets(ByVal ps_operation As xlSaveRestore, _
                               ByVal ps_ws As Worksheet)
' ------------------------------------------------------------------------------
' - ps_operation = xlSaveAndOff
'   Keeps (adds) a record of the Worksheet's (ps_ws) protection status and turns
'   protection off.
' - ps_operation = xlRestore
'   CleanUp the sheet's (ps_ws) protection status in case it was initially
'   protected.
'
' Requires Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin June 2019
' ------------------------------------------------------------------------------
    Const PROC     As String = "ObstProtectedSheets"
    
    On Error GoTo eh
    Dim cll         As Collection

    If dcProt Is Nothing Then Set dcProt = New Dictionary
    
    With ps_ws
        Select Case ps_operation
            Case xlSaveAndOff
                If ps_ws.ProtectContents Then
                    If Not dcProt.Exists(ps_ws) Then
                        Set cll = New Collection
                        cll.Add ps_ws.ProtectContents
                        dcProt.Add ps_ws, cll
                    Else
                        Set cll = dcProt.Item(ps_ws)
                        cll.Add ps_ws.ProtectContents ' may be true or false
                        dcProt.Remove ps_ws
                        dcProt.Add ps_ws, cll
                    End If
                Else
                    If dcProt.Exists(ps_ws) Then
                        Set cll = dcProt.Item(ps_ws)
                        cll.Add ps_ws.ProtectContents ' may be true or false
                        dcProt.Remove ps_ws
                        dcProt.Add ps_ws, cll
                    Else ' The sheet were never protected
                    End If
                End If
                ps_ws.Unprotect
            
            Case xlRestore
                If dcProt.Exists(ps_ws) Then
                    Set cll = dcProt.Item(ps_ws)
                    With cll
                        If .Count > 0 Then
                            If .Item(cll.Count) Then
                                ps_ws.Protect
                            Else
                                ps_ws.Unprotect
                            End If
                            .Remove .Count ' take off last saved item from stack
                        End If
                        If cll.Count = 0 Then dcProt.Remove ps_ws
                    End With
                End If
        End Select
    End With
         
xt: Exit Sub
    
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Function WsColsHidden(ByVal ws As Worksheet) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when at least one column in sheet (ws) is hidden
' ------------------------------------------------------------------------------
    Dim col As Range

    WsColsHidden = False
    For Each col In ws.UsedRange.Columns
        If col.Hidden Then
            WsColsHidden = True
            Exit Function
        End If
    Next col
    
End Function

Public Sub WsCustomView(ByVal SaveRestore As xlSaveRestore, _
                        ByVal ws As Worksheet, _
               Optional ByVal bColsHidden As Boolean = False, _
               Optional ByVal bRowsFiltered As Boolean = False)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "WsCustomView"
    On Error GoTo eh
    
    Dim dcCvsWs     As Dictionary
    Dim cv          As CustomView
    Dim cllCv       As Collection
    Dim cllProt     As Collection
    Dim wsTemp      As Worksheet
    Dim v           As Variant

    If dcCvsWb Is Nothing Then Set dcCvsWb = New Dictionary
    
    Select Case SaveRestore
        Case xlSaveOnly
            If Not dcCvsWb.Exists(wb) Then
                '~~ This is the first Save request for a CustomView for a Worksheet (ws) in the Workbook (wb)
                If bRowsFiltered Or bColsHidden Then
                    Set dcCvsWs = New Dictionary
                    Set cllCv = New Collection
                    Set cv = wb.CustomViews.Add(ViewName:=ws.Name & TEMPCVNAME, RowColSettings:=True)
                    cllCv.Add cv  ' keep record of the CustomView saved for the Worksheet (ws)
                    dcCvsWs.Add ws, cllCv ' Add a
                    dcCvsWb.Add wb, dcCvsWs
                End If
            Else ' dcCvsWb.Exists(wb)
                '~~ Apparently at least for one Worksheet in the Workbook a CustomView had already been saved
                Set dcCvsWs = dcCvsWb.Item(wb)
                If Not dcCvsWs.Exists(ws) Then
                    If bRowsFiltered Or bColsHidden Then
                        '~~ The first entry for a Worksheet's Obstruction save request
                        '~~ This is the first save of a CustomView for the Workseet (ws)
                        Set cllCv = New Collection
                        Set cv = wb.CustomViews.Add(ViewName:=ws.Name & TEMPCVNAME, RowColSettings:=True)
                        cllCv.Add cv  ' Save the CustomView created for the Worksheet (ws)
                        dcCvsWs.Add ws, cllCv
                    End If
                Else ' dcCvsWs.Exists(ws)
                    '~~ Apparently a CustomView had already been saved for the Worksheet (ws)
                    '~~ (the first entry for a Worksheet is always the one along with the creation of the CustomView)
                    '~~ thus this subsequent Save request is just added to the CustomView save-stack
                    Set cllCv = dcCvsWs.Item(ws)
                    cllCv.Add vbNullString
                    dcCvsWs.Remove ws
                    dcCvsWs.Add ws, cllCv
                End If
                dcCvsWb.Remove wb
                dcCvsWb.Add wb, dcCvsWs
            End If
        
        Case xlRestore
            '~~ Unstack the Save requests in reverse order. I.e. first all subsequent Save requests
            '~~ are unstacked and finally the created/saved CustomViews is restored
            Set dcCvsWs = dcCvsWb(wb)
            If dcCvsWs.Exists(ws) Then
                '~~ A CustomView had been created and saved for the Workseet (ws)
                Set cllCv = dcCvsWs.Item(ws)
                If cllCv.Count > 0 Then
                    With cllCv
                        If TypeName(.Item(.Count)) = "String" Then
                            '~~ "Unstack" the indication of a subsequent Save request
                            .Remove .Count
                        Else
                            Set cv = .Item(.Count)
                            If CstmVwExists(wb, cv) Then
                                '~~ Temporarily protect all not concerned Worksheets
                                '~~ save their sequence within the Workbook and move
                                '~~ the concerned Worksheet to the front
                                Set cllProt = New Collection
                                For Each wsTemp In wb.Sheets
                                    If Not wsTemp Is ws And wsTemp.ProtectContents = False Then
                                        cllProt.Add wsTemp  ' Collect the sheet for Unprotect
                                        wsTemp.Protect
                                    End If
                                Next wsTemp
                                WsSequence xlSaveOnly, wb
                                ws.Move Before:=wb.Sheets(1)
                                
                                '~~ The re-activated CustomView no can only be applied for the
                                '~~ first Worksheet in the Workbook, which is the one ment.
                                '~~ Activating the CustomView for any other Worksheet will fail
                                '~~ since they are all protected
                                ObstProtectedSheets xlSaveAndOff, ws
                                cv.show
                                ObstProtectedSheets xlRestore, ws
                                cv.Delete
                                dcCvsWs.Remove ws ' CustomView restore done for this Worksheet
                                
                                For Each v In cllProt: v.Unprotect: Next v  '~~ CleanUp the sheet's protection status
                                WsSequence xlRestore, wb                    '~~ CleanUp the Worksheet's initial sequence
                                
                            End If
                        End If
                    End With
                Else
                    Err.Raise 600, ErrSrc(PROC), "Obstruction Restore for Worksheet '" & ws.Name & "' request has no corresponding Save request (Save/Restore unpaired)!"
                End If
                Set cllCv = Nothing
            End If
    End Select

xt: Exit Sub
                 
eh:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Private Sub WsSequence(ByVal SaveRestore As xlSaveRestore, _
                       ByVal wb As Workbook)
' ------------------------------------------------------------------------------
' Saves and restore the sequence in which the Worksheets appear in the
' Worksheet's (ws) Workbook.
' ------------------------------------------------------------------------------
    Static cll  As Collection
    Dim ws      As Worksheet
    Dim i       As Long
    
    Select Case SaveRestore
        Case xlSaveOnly
            Set cll = New Collection
            For Each ws In wb.Sheets
                cll.Add ws
            Next ws
        Case xlRestore
            For i = 2 To cll.Count
                cll(i).Move After:=cll(i - 1)
            Next i
    End Select

End Sub

