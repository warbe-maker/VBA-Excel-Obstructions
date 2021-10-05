Attribute VB_Name = "mObstructions"
Option Explicit
' --------------------------------------------------------------------
' Standard Module mObstructions
'           Manages obstructions which hinder vba operations by
'           providing procedured to save and set them off and
'           restoring them. Typical operations prevented are:
'           - Rows move/copy (e.g. by filtered rows)
'           - Range value modifications
' Procedures:
' - Obstructions    summarizes all below and turns them off and
'                   retores them
' - AppEvents       Sets it to False and restores the initially
'                   saved status
' - FilteredRows    Turns Autofilter off when active and restores
'                   it by means of a CustomView.
' - FormEvents      Saves the actual events status and turns it
'                   off and restores it to the saved status
' - HiddenColumns   Displays them and restores them by means of a
'                   CustomView
' - MergedCells     Un-merges and re-merges cells associated with
'                   the current Selection.
' - SheetProtection Un-protects any number of sheets used in a
'                   project and re-protects them (only) when they
'                   were initially protected.
' - RangeNames      Saves and restores all formulas in Workbook
'                   which use a RangeName of a certain Worksheet
'                   by commenting and uncommenting the formulas
' Note 1: A CustomView is the means to restore Autofilter and/or hidden
'         columns of a certain Worksheet. CustomViews to save/restore
'         Autofilter a independant from those for saving/restoring
'         hidden columns.
' Note 2: In order to make an "elementary" operation like "copy row"
'         for instance independent from the environment the statements
'         will be enclosed in Obstructions xlSaveAndOff and Obstructions
'         xlRestore. However, if such an "elemetary" operation is just
'         one amongst others, performed by a more complex operation this
'         one as well should start with Obstructions xlSaveAndOff and
'         end with Obstructions xlRestore but additionally provided with
'         a "global" Dictionary and a "global" CustomView object. The
'         Obstructions call with the "elementary" operation will thus
'         not conflict with the "global" Off/CleanUp since there is
'         no longer an obstruction to turn off and subsequently none to
'         be restored.
' Uses the common components:
' - mCstmVw, m-Wrkbk; the common components mErH, fMsg, mMsg are only
'                     used by the mTest module and thus are not required
'                     when using mObstructions
' ----------------------------------------------------------------------
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

Public Sub Obstructions(ByVal SaveRestore As xlSaveRestore, _
                        ByVal ws As Worksheet, _
               Optional ByVal bSheetProtection As Boolean = False, _
               Optional ByVal bRowsFiltering As Boolean = False, _
               Optional ByVal bHiddenCols As Boolean = False, _
               Optional ByVal bCellsMerging As Boolean = False, _
               Optional ByVal bRangeNames As Boolean = False, _
               Optional ByVal bAppEvents As Boolean = False, _
               Optional ByVal bFormEvents As Boolean = False)
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
    
    On Error GoTo on_error
    BoP ErrSrc(PROC)
           
    bObstructionHiddenCols = bHiddenCols
    bObstructionFilteredRows = bRowsFiltering

    Select Case SaveRestore
        Case xlSaveAndOff
            
            '~~ 1. Save and turn off Application Events
            If bAppEvents Then
                AppEvents xlSaveAndOff
            End If
            
            '~~ 2. Save and turn off sheet protection if requested or implicitely required
            If bSheetProtection Or bRowsFiltering Or bHiddenCols Or bCellsMerging Then
                SheetProtection xlSaveAndOff, ws
            End If
            
            '~~ 3. Save and turn off Autofilter  if applicable
            If bObstructionFilteredRows Then
                FilteredRows xlSaveAndOff, ws
            End If
            
            '~~ 4. Save and turn off hidden columns  if applicable
            If bObstructionHiddenCols Then
                HiddenColumns xlSaveAndOff, ws
            End If
            
            '~~ 5. Save and turn off any effected merged cells if applicable
            If bCellsMerging Then
                MergedCells xlSaveAndOff
            End If
            
            '~~ 6. Save and turn off used range names
            If bRangeNames Then
                RangeNames xlSaveAndOff, ws
            End If
            
        Case xlRestore
            
            '~~ 1. CleanUp all formulas using a ws range name
            If bRangeNames Then
                RangeNames xlRestore, ws
            End If
            
            '~~ 2. CleanUp merge areas which were initially effected by the Selection
            If bCellsMerging Then
                MergedCells xlRestore
            End If
            
            '~~ 3. CleanUp Autofilter if applicable
            If bRowsFiltering Then
                FilteredRows xlRestore, ws
            End If
            
            '~~ 4. CleanUp hidden columns
            If bHiddenCols Then
                HiddenColumns xlRestore, ws
            End If
            
            '~~ 5. CleanUp the sheets protection status when it initially was protected
            If bSheetProtection Or bRowsFiltering Or bHiddenCols Or bCellsMerging Then
                SheetProtection xlRestore, ws
            End If
            
            '~~ 6. CleanUp Application Events status
            If bAppEvents Then
                AppEvents xlRestore
            End If
            
    End Select
         
exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub AppEvents(ByVal SaveRestore As xlSaveRestore)
' ------------------------------------------------------
' - SaveRestore = xlSaveAndOff: Saves the current
'   Application.EnableEvents status and turns it off.
'   Any subsequent execution will just save the status
'   (adds it to the stack)
' - SaveRestore = xlRestore: Restores the last saved
'   Application.EnableEvents status and removes the
'   saved item.
' ------------------------------------------------------
Const PROC = "AppEvents"

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    Select Case SaveRestore
        Case xlSaveAndOff
            If cllAppEvents Is Nothing Then Set cllAppEvents = New Collection
            cllAppEvents.Add Application.EnableEvents ' add status to stack
            Application.EnableEvents = False
            
        Case xlRestore
            If cllAppEvents Is Nothing Then GoTo exit_proc
            With cllAppEvents
                If .Count > 0 Then
                    '~~ restore last saved statis item and remove item (take it off from stack)
                    Application.EnableEvents = .Item(cllAppEvents.Count)
                    .Remove .Count
                End If
            End With
    End Select

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub FilteredRows(ByVal SaveRestore As xlSaveRestore, _
                        ByVal ws As Worksheet)
' -----------------------------------------------------------
' - SaveRestore = xlSaveAndOff: When Autofilter is active a
'   temporary CustomView is created and Autofilter is turned
'   off.
' - SaveRestore = xlRestore: Returns to the temporary created
'   CustomView (if any) and thus restores the Autofilter with
'   all its initial specifications.
' General:
' - Save/Restore requests may be nested but it is absolutely
'   essential that they are paired!
' - Subsequent SaveAndOff request (e.g. in nested sub-
'   procedures) are just "stacked", subsequent Restore
'   Restore requests are just un-stacked - and thus do not
'   cause any problem - provided they are paired.
' - Filtered rows have to be turned off by Worksheet
' - Worksheet's obstructions may be restored in any order
'   order! (Save wsTest1, wsTest2, Restore wsTest2, wsTest1)
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin Dec 2019
' -------------------------------------------------------------
Const PROC = "FilteredRows"

    On Error GoTo on_error
    BoP ErrSrc(PROC)

    '~~ The Workbook of the Worksheet may not be found within this Application instance.
    '~~ Application.Workbooks() may thus not be appropriate. GetOpenWorkbook will find
    '~~ it in whichever Application instance
    Set wb = mWrkbk.GetOpen(ws.Parent.Name)
    If dcCvsWb Is Nothing Then Set dcCvsWb = New Dictionary
    
    Select Case SaveRestore
        Case xlSaveAndOff
            If ws.AutoFilterMode = True Then
                '~~ Create a CustomView, keep a record of the CustomView and turn filtering off
                WsCustomView xlSaveOnly, ws, bRowsFiltered:=True
                SheetProtection xlSaveAndOff, ws    ' Possibly nested request ensuring unprotection
                ws.AutoFilterMode = False
                SheetProtection xlRestore, ws       ' Possibly nested restore ensuring protection status restore
            Else
                WsCustomView xlSaveOnly, ws ' Just add subsequent save request to stack
            End If
        
        Case xlRestore
            '~~ CleanUp the CustomView saved for Worksheet (ws) if any
            If dcCvsWb.Exists(wb) Then ' Only if at least for one Worksheet a CustomView had been saved
                WsCustomView xlRestore, ws
            End If
    End Select

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub

on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub HiddenColumns(ByVal SaveRestore As xlSaveRestore, _
                         ByVal ws As Worksheet)
' --------------------------------------------------------------
' - SaveRestore = xlSaveAndOff: Create/Save CustomView and
'   display all hidden columns
' - SaveRestore = xlRestore: CleanUp the saved CustomView.
' Note: May be called in nested subprocedures without a problem.
' --------------------------------------------------------------
Const PROC      As String = "HiddenColumns"
Dim col         As Range

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    Set wb = mWrkbk.GetOpen(ws.Parent.Name)
    
    Select Case SaveRestore
        
        Case xlSaveAndOff
            If WsColsHidden(ws) Then
                '~~ If not one already exists create a CustomView for this Worksheet and keep a record of it
                '~~ and un-hide all hidden columns
                SheetProtection xlSaveAndOff, ws
                WsCustomView xlSaveOnly, ws, bColsHidden:=True
                For Each col In ws.UsedRange.Columns
                    If col.Hidden Then col.Hidden = False
                Next col
                SheetProtection xlRestore, ws
            Else
                '~~ Add a subsequent save request to the Worksheet's save stack
                WsCustomView xlSaveOnly, ws, bColsHidden:=True
            End If
        
        Case xlRestore
            WsCustomView xlRestore, ws
    End Select
    
exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub MergedCells(ByVal SaveRestore As xlSaveRestore, _
              Optional ByRef dcGlbl As Variant = Null)
' ------------------------------------------------------------
' - xlOffOn = xlSaveAndOff: Any merge area associated with a
'   in the current Selection is un-merged by saving the merge
'   areas' address in a temporary range name and additionally
'   in a Dictionary.
'   The content of the top left cell is copied to all cells
'   in the un-merged area to prevent a loss of the merge
'   area's content even when the top row of it is deleted.
'   The named ranges address is automatically maintained by
'   Excel throughout any rows operations performed  within the
'   originally merged area's top and bottom row. I.e. any row
'   copied or inserted above the top row or below the bottom
'   row will not become part of the retored merge area(s).
' - xlOffOn = xlRestore: All merge areas registered by a
'   temporary range name are re-merged, thereby eliminating
'   all duplicated content except the one in the top left
'   cell.
' When no merge areas are detected neither of the MergedCells
' call does anything. I.e. no need to check for any merged
' cells beforehand.
'
' Used by Obstructions, uses Merge
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin June 2019
' ------------------------------------------------------------
Const PROC     As String = "MergedCells"
Const sTempName As String = "rngTempMergeAreaName"
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

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    '~~ When provided, the global dictionary is used, else the local
    If Not IsNull(dcGlbl) Then
        If dcGlbl Is Nothing Then
            Set dcGlbl = New Dictionary
        End If
        Set dc = dcGlbl
    Else
        If dcLocl Is Nothing Then
            Set dcLocl = New Dictionary
        End If
        Set dc = dcLocl
    End If
    
    If SaveRestore = xlSaveAndOff Then
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
            
    ElseIf SaveRestore = xlRestore Then
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

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub SheetProtection(ByVal SaveRestore As xlSaveRestore, _
                           ByVal ws As Worksheet)
' --------------------------------------------------------------
' - SaveRestore = xlSaveAndOff: Keeps (adds) a record of the
'   Worksheet's (ws) protection status and turns protection off.
'
' - SaveRestore = xlRestore: CleanUp the sheet's (ws) protection
'   status in case it was initially protected.
'
' Requires Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin June 2019
' ----------------------------------------------------------------------
Const PROC     As String = "SheetProtection"
Static dcLocl   As Dictionary
Dim col         As Range
Dim dc          As Dictionary
Dim i           As Long
Dim cll         As Collection

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    If dcProt Is Nothing Then Set dcProt = New Dictionary
    
    With ws
        Select Case SaveRestore
            Case xlSaveAndOff
                If ws.ProtectContents Then
                    If Not dcProt.Exists(ws) Then
                        Set cll = New Collection
                        cll.Add ws.ProtectContents
                        dcProt.Add ws, cll
                    Else
                        Set cll = dcProt.Item(ws)
                        cll.Add ws.ProtectContents ' may be true or false
                        dcProt.Remove ws
                        dcProt.Add ws, cll
                    End If
                Else
                    If dcProt.Exists(ws) Then
                        Set cll = dcProt.Item(ws)
                        cll.Add ws.ProtectContents ' may be true or false
                        dcProt.Remove ws
                        dcProt.Add ws, cll
                    Else ' The sheet were never protected
                    End If
                End If
                ws.Unprotect
            
            Case xlRestore
                If dcProt.Exists(ws) Then
                    Set cll = dcProt.Item(ws)
                    With cll
                        If .Count > 0 Then
                            If .Item(cll.Count) Then
                                ws.Protect
                            Else
                                ws.Unprotect
                            End If
                            .Remove .Count ' take off last saved item from stack
                        End If
                        If cll.Count = 0 Then dcProt.Remove ws
                    End With
                End If
        End Select
    End With
         
exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub RangeNames(ByVal SaveRestore As xlSaveRestore, _
                      ByVal ws As Worksheet)
' ------------------------------------------------------
' Names are regarded an obstruction when the Worksheet
' (ws) is to be copied or moved from one Workbook to
' another. Since the names for ranges in the source
' Workshhet are not copied to the target Workbook all
' names are subject to a decision stay or another
' decision.
' SaveRestore = xlSaveAndOff: All references to a
' name in the source Worksheet (ws) are turned into a
' direct range reference.
'
' -------------------------------------------------------
Const PROC      As String = "RangeNames"
Static dcLocl   As Dictionary   ' Keeps a record for each name replaced by its reference
Dim dcNames     As Dictionary   ' Names which refer to a range in the Worksheet ws
Dim dcCells     As Dictionary   ' Keeps a record of each cell's formula which had been modified
Dim nm          As Name
Dim wsheet      As Worksheet
Dim cel         As Range
Dim v           As Variant
Dim wb          As Workbook

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    '~~ The Workbook of the Worksheet may not be found within this Application instance.
    '~~ Application.Workbooks() may thus not be appropriate. GetOpenWorkbook will find
    '~~ it in whichever Application instance
    Set wb = mWrkbk.GetOpen(ws.Parent.Name)
    
    Select Case SaveRestore
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
            For Each nm In ws.Names
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

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Private Sub Merge(ByVal r As Range, ByVal OffOn As xlSaveRestore)
' ---------------------------------------------------------------
' OffOn = xlSaveAndOff: un-merges range r by copying the top left
'                       content to all cells in the merge area.
' OffOn = xlRestore: Re-merges range r.
' ---------------------------------------------------------------
Const PROC As String = "Merge"
Dim rSel    As Range
Dim cel     As Range
Dim bEvents As Boolean
Dim row     As Range

    On Error GoTo on_error
    
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
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub Borders(ByVal r As Range, _
                   ByVal SaveRestore As xlSaveRestore, _
                   ByRef dct As Dictionary)
' ------------------------------------------------------
' - xlOffOn = xlSaveOnly: Saves the border properties of
'   Range r into the Dictionary dct.
' - xlOnOff = xlRestore: Restores all border properties
'   for Range r from the Dictionary dct.
' ------------------------------------------------------
Const PROC = "Borders"
Dim cll             As Collection
Dim bo              As Border
Dim xlBi            As XlBordersIndex

    On Error GoTo on_error
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

on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Sub WsCustomView(ByVal SaveRestore As xlSaveRestore, _
                        ByVal ws As Worksheet, _
               Optional ByVal bColsHidden As Boolean = False, _
               Optional ByVal bRowsFiltered As Boolean = False)
' ----------------------------------------------------
'
' ----------------------------------------------------
Const PROC      As String = "WsCustomView"
Dim dcCvsWs     As Dictionary
Dim cv          As CustomView
Dim cllCv       As Collection
Dim cllProt     As Collection
Dim wsTemp      As Worksheet
Dim v           As Variant

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
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
                            If mCstmVw.Exists(wb, cv) Then
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
                                SheetProtection xlSaveAndOff, ws
                                cv.show
                                SheetProtection xlRestore, ws
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

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
                 
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    ErrMsg ErrSrc(PROC)
End Sub

Public Function WsColsHidden(ByVal ws As Worksheet) As Boolean
' -------------------------------------------------------------
' Returns TRUE when at least one column in sheet (ws) is hidden
' -------------------------------------------------------------
Dim col As Range

    WsColsHidden = False
    For Each col In ws.UsedRange.Columns
        If col.Hidden Then
            WsColsHidden = True
            Exit Function
        End If
    Next col
    
End Function
Public Sub CleanUp(Optional ByVal bForce As Boolean = False)
' ----------------------------------------------------------
' Does cleanup for all obstructions still waiting for
' a CleanUp, likely not done due to an error.
' When bForce is True all remaining Restores are done
' without notice. This may be used in the error handling of
' the project after the error message had been displayed.
' ----------------------------------------------------------
Dim cv      As CustomView
Dim wb      As Workbook
Dim ws      As Worksheet
Dim v1      As Variant
Dim v2      As Variant
Dim v3      As Variant
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
                            If mCstmVw.Exists(wb, cv) Then
                                If MsgBox(" CleanUp the yet unrestored CustomView for Worksheet '" & ws.Name & "' ?", vbYesNo, "Unrestored CustomView") = vbYes Then
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
    
    '~~ CleanUp AppEvents (allway done silent)
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

Private Sub WsSequence(ByVal SaveRestore As xlSaveRestore, _
                       ByVal wb As Workbook)
' ----------------------------------------------------------
' Save and restore the sequence of all Worksheets in
' the Worksheet's (ws) Workbook.
' ----------------------------------------------------------
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

Private Sub NamesRemove(ByVal sName As String, _
               Optional ByVal ws As Worksheet = Nothing, _
               Optional ByVal bConfirm As Boolean = True)
' ---------------------------------------------------------
' Removes the name sName from the list of Names provided
' the name refers to ws which defaults to the ActiveSheet.
' When there is no sName found sName is regarded a generic
' part of names and all names with it are removed.
' ---------------------------------------------------------
Const PROC     As String = "NamesRemove"
Dim nm          As Name
Dim sNewName    As String
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
                GoTo exit_proc
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
    If bConfirmed Then GoTo exit_proc
    If bConfirm And sConfirm <> vbNullString Then
        If MsgBox("Yes if the following renames are to be removed: " & sConfirm, vbYesNo, "Confirm removals") = vbYes Then
            bConfirmed = True
            GoTo again_confirmed:
        Else
            GoTo exit_proc
        End If
    End If

exit_proc:
End Sub

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString)
' ------------------------------------------------------
' This Common Component does make use of a Common Error
' Handling module (in order to limit the number of used
' components. Instead it passes on any error to this
' procedure.
' ------------------------------------------------------
    
    If err_no = 0 Then err_no = Err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description

    Application.EnableEvents = True
    Err.Raise Number:=err_no, Source:=err_source, Description:=err_dscrptn

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mObstructions" & "." & sProc
End Function
