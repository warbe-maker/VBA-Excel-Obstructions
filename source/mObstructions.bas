Attribute VB_Name = "mObstructions"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mObstructions
'
' Services to manages obstructions hindering vba operations otherwise, such as
' merged cells for instance to name one of the most ugly ones first.
' Obstructions are managed by an 'Eliminate' and  'Restore service. 'Eliminate'
' is performed at the beginning of a procedure and the 'Restore' service at the
' end. Because the 'Eliminate' service pushes the entry status on a stack from
' where the 'Restorte' service pops it both can be performed on any nested level
' provided thees two services are strictly  p a i r e d ! Due to the stack
' approach there is no need to check the status of any obstruction beforehand.
'
' Public services:
' ------------------------------------------------------------------------------
' - All                    Summarizes all available services allowing to
'                          explicitely ignore some.
' - ApplEvents             'Eliminate' the current Application.
'                          EnableEvents status and restore it to the saved
'                          status
' - FilteredRowsHiddenCols Turns Autofilter off when active and restores
'                          it by means of a CustomView.
' - MergedAreas            enEliminate Un-merges, enRestore re-merges cells
'                          associated with the current Selection.
' - SheetProtection        Un-protects any number of sheets used in a project
'                          and re-protects them (only) when they
'                          were initially protected.
' - ObstNamedRanges        Saves and restores all formulas in Workbook
'                          which use a RangeName of a certain Worksheet
'                          by commenting and uncommenting the formulas
' - Rewind                 Rewinds all 'Saved-and-Set-Off' obstruction to
'                          their initial status. May only be used in case of
'                          an error in order to end up with all obstructions
'                          restored.
' ------------------------------------------------------------------------------
' Note 1: For filtered rows and/or hidden columns a temporary added CustomView
'         is the means to restore a set-off Autofilter and re-hide displayed
'         hidden columns. For this obstruction the very initial 'Eliminate'
'         service pushes the temporary CustomView name to a dedicated stack.
'         Every subsequent (nested) 'Eliminate' consequently pushes a
'         vbNullString on the stack.
' Note 2: The stacking approach makes all obstruction sevices extremly robust.
'         Eliminate                      saves and sets off the obstruction
'         | Eliminate                    just stacks the already set off status
'         | |  Eliminate                 just stacks the already set off status
'         | |  |                         any nested procedure
'         | |  Restore                   restores nothing
'         | Restore                      restores nothing
'         Restore                        restores the initial status
'
' Note 3: The module/component uses the following other common components:
'         mWrkbk (the common components mErH, fMsg, mMsg are only used by the
' mTest module and thus are not required when using mObstructions)
'
' Requires: - Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin, Nov 2021
' See: https://github.com/warbe-maker/Common-Excel-VBA-Obstructions-Services
' ------------------------------------------------------------------------------
Private Const TEMP_MERGED_AREA_NAME As String = "TempObstructionMergeAreaName"
Private Const TEMP_CUSTOM_VIEW_NAME As String = "TempObstructionCustomViewName"

Public Enum enObstService
    enEliminate
    enRestore
End Enum

Private MergedAreasSheetStacks              As Dictionary   ' Sheet specific merged areas stacks
Private SheetProtectionSheetStacks          As Dictionary   ' Sheet spcific protection stacks
Private ApplEventsStack                     As Collection   ' Application.EnableEvents stack
Private FilteredRowsHiddenColsSheetStacks   As Dictionary   ' Sheet specific filter/hidden stacks

Private Property Get SheetStack(Optional ByRef stacks As Dictionary, _
                                Optional ByVal ws As Worksheet) As Collection
' ------------------------------------------------------------------------------
' Returns the sheet (ws) specific stack of the sheet stacks (stacks). when the
' sheet stacks (stacks) Dictionary doesn't exist it is created with an empty
' stack.
' ------------------------------------------------------------------------------
    Dim stack As Collection
    
    If stacks Is Nothing Then Set stacks = New Dictionary
    If Not stacks.Exists(ws) Then
        Set stack = New Collection
        stacks.Add ws, stack
    Else
        Set stack = stacks(ws)
   End If
   Set SheetStack = stack

End Property

Private Property Let SheetStack(Optional ByRef stacks As Dictionary, _
                                Optional ByVal ws As Worksheet, _
                                         ByVal stack As Collection)
' ------------------------------------------------------------------------------
' Replaces in the sheet specific stack (stack) in the Dictionary of sheet stacks
' (stacks) with the provided sheet specific stack (stack).
' When the sheet specific stack (stack) is empty it is removed from the stacks.
' ------------------------------------------------------------------------------
    If stacks.Exists(ws) Then stacks.Remove ws
    If Not BasicStackIsEmpty(stack) Then stacks.Add ws, stack
End Property

Private Property Get TempMergedAreaName(Optional ws As Worksheet) As String
    TempMergedAreaName = TEMP_MERGED_AREA_NAME & Replace(ws.Name, " ", "_")
End Property

Private Function AppErr(ByVal err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error's number never conflicts
' with VB runtime or other errors. The function returns a given positive error
' number (err_no) with the vbObjectError added - thereby turning it into a
' negative value. When the provided error number (err_no) is negative the
' original positive "application" error number is returned.
' ------------------------------------------------------------------------------
    If err_no >= 0 Then AppErr = err_no + vbObjectError Else AppErr = Abs(err_no - vbObjectError)
End Function

Public Sub ApplEvents(ByVal obs_mode As enObstService)
' ------------------------------------------------------------------------------
' Obstruction service 'Application.EnableEvents':
' Eliminate: Pushes the current Application.EnableEvents on the 'ApplEventsStack'
'            and turns it off.
' Restore:   Pops the status form the 'ApplEventsStack' and  sets the
'            EnableEvents accordingly.
'
' W. Rauschenberger, Berlin Nov 2021
' ------------------------------------------------------------------------------
    Const PROC = "ApplEvents"

    On Error GoTo eh
    
    Select Case obs_mode
        Case enEliminate
            BasicStackPush ApplEventsStack, Application.EnableEvents
            Application.EnableEvents = False
            
        Case enRestore
            If Not BasicStackIsEmpty(ApplEventsStack) Then
                Application.EnableEvents = BasicStackPop(ApplEventsStack)
            End If
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub MergedAreas(ByVal obs_mode As enObstService, _
                       ByVal obs_ws As Worksheet, _
              Optional ByVal obs_range As Range)
' ------------------------------------------------------------------------------
' Obstruction service 'Merged Areas/Cells':
' Eliminate: Any merge area concerned by the provided range's (obs_range) rows
'            and columns is un-merged by saving the merge areas' address in a
'            temporary range name and additionally in a Dictionary. The content
'            of the top left cell is copied to all cells in the un-merged area
'            to prevent a content loss even when the top row or the left column
'            is deleted. The named ranges address is automatically maintained by
'            Excel throughout any rows operations performed within the
'            originally merged area's top and bottom row. I.e. any row copied or
'            inserted above the top row or below the bottom row will not become
'            part of the retored merge area(s).
' Restore:   The temporary range name is popped from the sheet specific stack.
'            The ranges are re-merged, thereby eliminating all duplicated
'            content except the one in the top left cell. In contrast to
'            'Eliminate' the range argument (obs_range) is not used but the
'            Worksheet (obs_ws) only.
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin Nov 2021
' ------------------------------------------------------------------------------
    Const PROC = "MergedAreas"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim stack       As Collection
    
            
    Select Case obs_mode
        Case enEliminate
            SheetProtection obs_mode:=enEliminate, obs_ws:=obs_ws
            
            MergedAreas1SaveAndUnMerge obs_range
            
            SheetProtection obs_mode:=enRestore, obs_ws:=obs_ws
    
        Case enRestore
            If MergedAreasSheetStacks.Exists(obs_ws) Then
                SheetProtection obs_mode:=enEliminate, obs_ws:=obs_ws
                
                Set stack = SheetStack(MergedAreasSheetStacks, obs_ws)
                If Not BasicStackIsEmpty(stack) Then
                    Set dct = BasicStackPop(stack)
                    MergedAreas2Restore dct
                End If
                SheetStack(MergedAreasSheetStacks, obs_ws) = stack
                
                SheetProtection obs_mode:=enRestore, obs_ws:=obs_ws
'            Else
'                Err.Raise AppErr(1), ErrSrc(PROC), _
'                          "The corresponding merged cells 'Save' is missing for the merged cells 'Restore' " & _
'                          "for sheet '" & obs_ws.Name & "'! " & _
'                          "P a i r e d  'Save' and 'Restore' operations are obligatory!"
            End If
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub SheetProtection(ByVal obs_mode As enObstService, _
                  Optional ByVal obs_wb As Workbook = Nothing, _
                  Optional ByVal obs_ws As Worksheet = Nothing)
' ------------------------------------------------------------------------------
' Obstruction service 'Sheet Protection':
' The 'Eliminate' and 'Retore' service is either provided for the provided
' Worksheet (obs_ws) or when none is provided for all Worksheets in the provided
' Workbook (obs_wb). When neither a Workbook nor a Worksheet is provided the
' service raises an error
' Eliminate: Pushes the Worksheet's (obs_ws) protection status and turns on the
'            Worksheet's specific stack and sets the protection off. When no
'            Worksheet is provided all sheets protection is set off.
' Restore:   Restores the Worksheets's (obs_ws) protection status by popping the
'            sheet's former protection status off its dedicated stack. When no
'            Worksheet is provided this is done for all sheets in the Workbook
'            (obs_wb).
'
' Requires:  Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin June 2019
' ------------------------------------------------------------------------------
    Const PROC = "BasicObstSheetProtections"
    
    On Error GoTo eh
    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim stack   As Collection

    If obs_wb Is Nothing And obs_ws Is Nothing _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Neither a Worksheet (obs_ws) nor a Workbook (obs_wb) is provided!"
    
    If Not obs_wb Is Nothing And obs_ws Is Nothing Then Set wb = obs_wb
    If obs_wb Is Nothing And Not obs_ws Is Nothing Then Set wb = obs_ws.Parent
    
    If SheetProtectionSheetStacks Is Nothing Then Set SheetProtectionSheetStacks = New Dictionary
    
    For Each ws In wb.Worksheets
        If obs_ws Is Nothing Or Not obs_ws Is Nothing And ws Is obs_ws Then
            With ws
                If Not SheetProtectionSheetStacks.Exists(ws) Then
                    Set stack = New Collection
                    SheetProtectionSheetStacks.Add ws, stack
                End If
                Select Case obs_mode
                    Case enEliminate
                        Set stack = SheetStack(SheetProtectionSheetStacks, ws)
                        BasicStackPush stack, .ProtectContents      ' push True or False on the 'Stack'
                        SheetStack(SheetProtectionSheetStacks, ws) = stack     ' replace the Stack in the SheetProtectionSheetStacks
                        ws.Unprotect
                    
                    Case enRestore
                        If SheetProtectionSheetStacks.Exists(ws) Then
                            Set stack = SheetProtectionSheetStacks(ws)
                            If BasicStackPop(stack) Then ws.Protect Else ws.Unprotect
                            SheetStack(SheetProtectionSheetStacks, ws) = stack
                        Else
                            Err.Raise AppErr(1), ErrSrc(PROC), _
                                      "The corresponding protection status 'Save' is missing for the protection status 'Restore' " & _
                                      "of sheet '" & ws.Name & "'! " & _
                                      "P a i r e d  'Save' and 'Restore' operations are obligatory!"
                        End If
                End Select
            End With
            If Not obs_ws Is Nothing Then GoTo xt ' this Worksheet only
        End If
    Next ws

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function BasicStackIsEmpty(ByVal stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Basic Stack Empty check service. Returns True when either there is no stack
' (stck Is Nothing) or when the stack is empty (items count is 0).
' ----------------------------------------------------------------------------
    If stck Is Nothing Then Set stck = New Collection
    BasicStackIsEmpty = stck.Count = 0
End Function

Private Function BasicStackPop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Basic Stack Pop service. Returns the last item pushed on the stack (stck)
' and removes the item from the stack. When the stack (stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "BasicStackPop"
    
    On Error GoTo eh
    
    If BasicStackIsEmpty(stck) Then GoTo xt
    
    On Error Resume Next
    Set BasicStackPop = stck(stck.Count)
    If Err.Number <> 0 _
    Then BasicStackPop = stck(stck.Count)
    stck.Remove stck.Count

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub BasicStackPush(ByRef stck As Collection, _
                           ByVal stck_item As Variant)
' ----------------------------------------------------------------------------
' Basic Stack Push service. Pushes (adds) an item (stck_item) to the stack
' (stck). When the provided stack (stck) is Nothing the stack is created.
' ----------------------------------------------------------------------------
    Const PROC = "BasicStackPush"
    
    On Error GoTo eh
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_item

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

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

Private Function ColsHidden(ByVal ch_ws As Worksheet, _
                   Optional ByVal ch_check_only As Boolean = True) As Boolean
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim col As Range
    
    For Each col In ch_ws.UsedRange.Columns
        If col.Hidden Then
            If ch_check_only Then
                ColsHidden = True
                Exit Function
            Else
                col.Hidden = False
            End If
        End If
    Next col

End Function


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
    ErrSrc = "mObstructions" & "." & sProc
End Function

Public Sub FilteredRowsHiddenCols(ByVal obs_mode As enObstService, _
                                  ByVal obs_ws As Worksheet)
' ------------------------------------------------------------------------------
' Obstruction service 'Filtered Rows':
' The service covers 'AutoFilter' and 'Hidden Columns'.
' Eliminate: When AutoFilter is active a CustomView with a temporary name is
'            added and pushed on the 'FilteredRowsAndOrColsStack' and
'            'AutoFilter' is turned off and hidden columns are displayed. When
'            neither 'AutoFilter' is active nor a column is hidden a
'            vbNullString is pushed on the sheet specific stack of the
'            'FilteredRowsAndOrColsStacks'.
'
' - obs_mode = enRestore: Restores the temporary CustomView with the name
'   poped from the sheet stack of the 'FilteredRowsHiddenColsSheetStacks' which
'   turns AutoFilter back on and hiddes formerly hidden rows. When a
'   vbNullstring is poped no action is taken.
'
' Note 1: Eliminate/Restore services may be nested but it is absolutely essential
'         that they are paired Worksheet wise!
' Note 2: In contrast to other obstruction services! When this obstruction's
'         Eliminate service is performed for more than one Worksheet the
'         Restore services have to be performed in reverse order. If not an
'         error is raised and it is with the callers to perform an Rewind
'         service which does exactly that.
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin Dec 2019
' ------------------------------------------------------------------------------
    Const PROC = "FilteredRowsHiddenCols"
    
    On Error GoTo eh
    Dim stack   As Collection
    Dim wb      As Workbook
    Dim ws      As Worksheet
    
    Set wb = obs_ws.Parent
    
    Select Case obs_mode
        Case enEliminate
        
            SheetProtection obs_mode:=enEliminate, obs_ws:=obs_ws
            Set stack = SheetStack(FilteredRowsHiddenColsSheetStacks, obs_ws)
            FilteredRowsHiddenCols1Eliminate obs_ws, stack
            SheetStack(FilteredRowsHiddenColsSheetStacks, obs_ws) = stack
            SheetProtection obs_mode:=enRestore, obs_ws:=obs_ws       ' Possibly nested restore ensuring protection status restore
        
        Case enRestore
        
            '~~ Important!
            '~~ Showing a CustomView potentially fails (i.e. stops half way dome) when any sheet is protected.
            '~~ The only way to re-establish a CustomView is to make sure no sheet is protected. By no providing
            '~~ a certain Worksheet but only the Workbook does this with the SheetProtection obstruction service.
            SheetProtection obs_mode:=enEliminate, obs_wb:=obs_ws.Parent   ' unprotect all sheets
            Set stack = SheetStack(FilteredRowsHiddenColsSheetStacks, obs_ws)
            FilteredRowsHiddenCols2Restore obs_ws, stack
            SheetStack(FilteredRowsHiddenColsSheetStacks, ws) = stack
            SheetProtection obs_mode:=enRestore, obs_wb:=obs_ws.Parent     ' re-protect all sheets originally protected
    End Select

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub FilteredRowsHiddenCols1Eliminate( _
            ByVal obs_ws As Worksheet, _
            ByRef obs_stack As Collection)
' ------------------------------------------------------------------------------
' - Pushes a temporary CustomView name on the stack (obs_stck) when either
'   AutoFilter is TRUIE or any columns are hidden,
' - Sets AutoFilter off and displays any hidden columns.
' ------------------------------------------------------------------------------
    Const PROC = "FilteredRowsHiddenCols1Eliminate"
    
    On Error GoTo eh
    Dim TempCustViewName    As String
    Dim wb                  As Workbook
    
    BoP ErrSrc(PROC)
    Set wb = obs_ws.Parent
    TempCustViewName = vbNullString
    
    If obs_ws.AutoFilterMode = True Or ColsHidden(obs_ws) Then
        TempCustViewName = TEMP_CUSTOM_VIEW_NAME & "_" & Replace(obs_ws.Name, " ", "_")
        '~~ Create a CustomView, keep a record of the CustomView and turn filtering off
        On Error Resume Next
        wb.CustomViews(TempCustViewName).Delete ' in case one exists under this name
        On Error GoTo eh
        wb.CustomViews.Add ViewName:=TempCustViewName, RowColSettings:=True
        
        '~~ Turn off AutoFilter and display hidden columns
        obs_ws.AutoFilterMode = False
        ColsHidden obs_ws, False
    End If
    BasicStackPush obs_stack, TempCustViewName

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub FilteredRowsHiddenCols2Restore(ByVal obs_ws As Worksheet, _
                                           ByRef obs_stack As Collection)
' ------------------------------------------------------------------------------
' Pops a temporary CustomView name off from the stack (obs_stack) and when the
' temporary name is not vbNullString shows this Custom View and deletes it.
' ------------------------------------------------------------------------------
    Const PROC = "FilteredRowsHiddenCols2Restore"
    
    On Error GoTo eh
    Dim wb                  As Workbook
    Dim TempCustViewName    As String
    
    BoP ErrSrc(PROC)
    Set wb = obs_ws.Parent
    
    If Not BasicStackIsEmpty(obs_stack) Then
        TempCustViewName = BasicStackPop(obs_stack)
        If TempCustViewName <> vbNullString Then
            '~~ It is absolutely essential that there are no protected sheets in the Workbook
            '~~ when the CustomView is re-shown
            SheetProtection obs_mode:=enEliminate, obs_wb:=wb
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
            mBasic.TimedDoEvents "> wb.CustomViews(TempCustViewName).Show"
#End If
            wb.CustomViews(TempCustViewName).Show
#If Debugging = 1 Then ' The mBasic component is only available (and used) in the development and test environment
            mBasic.TimedDoEvents "< wb.CustomViews(TempCustViewName).Show"
#End If
            wb.CustomViews(TempCustViewName).Delete
            SheetProtection obs_mode:=enRestore, obs_wb:=wb
        End If
    Else
        Err.Raise AppErr(1), ErrSrc(PROC), _
                  "The corresponding filtered rows and/or hidden cols 'Save' is missing for the corresponding 'Restore' " & _
                  "of sheet '" & obs_ws.Name & "'! " & _
                  "P a i r e d  'Save' and 'Restore' operations are obligatory!"
    End If

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt ' clean exit
    End Select
End Sub

Private Sub MergedArea1UnMerge(ByVal obs_range As Range)
' ------------------------------------------------------------------------------
' Saves/Restores a merged cells range (obs_range).
' - obs_mode = enEliminate: un-merges range obs_range by copying the top
'                                left content to all cells in the merge area.
' - obs_mode = enRestore: Re-merges range obs_range.
' ------------------------------------------------------------------------------
    Const PROC As String = "MergedArea"
    
    On Error GoTo eh
    Dim cel     As Range
    Dim rRow    As Range
    Dim ws      As Worksheet
    
    Application.ScreenUpdating = False
    Set ws = obs_range.Parent
    ApplEvents enEliminate
    SheetProtection obs_mode:=enEliminate, obs_ws:=ws
    
    With obs_range
        '~~ Avoid automatic rows height adjustment by setting it explicit
        For Each rRow In obs_range.Rows
            With rRow
                .RowHeight = .RowHeight + 0.01
            End With
        Next rRow
        .UnMerge
        '~~ In order not to loose the value in the top left cell
        '~~ e.g. when a row is deleted or moved up or down
        '~~ it is copied to all other cells. Merge will return to
        '~~ the then top left cell's content.
        obs_range.Cells(1, 1).Copy
        For Each cel In obs_range.Cells
            If cel.Value = vbNullString Then
                cel.PasteSpecial Paste:=xlPasteAllExceptBorders, _
                             operation:=xlNone, _
                            SkipBlanks:=False, _
                             Transpose:=False
             End If
        Next cel
        Application.CutCopyMode = False
    End With ' obs_range
    
xt: SheetProtection obs_mode:=enRestore, obs_ws:=ws
    ApplEvents enRestore
    Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MergedArea2ReMerge(ByVal obs_range As Range)
' ------------------------------------------------------------------------------
' Re-merges range obs_range.
' ------------------------------------------------------------------------------
    Const PROC As String = "MergedArea"
    
    On Error GoTo eh
    Dim bEvents As Boolean
    Dim rRow    As Range
    
    With Application
        bEvents = .EnableEvents
        .EnableEvents = False
        .ScreenUpdating = False
        
        '~~ Reset to original row height
        For Each rRow In obs_range.Rows
            On Error Resume Next
            With rRow: .RowHeight = .RowHeight - 0.01: End With
        Next rRow
        .DisplayAlerts = False ' prevent allert for content in other cells which is ignored
        obs_range.Merge
        .DisplayAlerts = True
                
        .EnableEvents = bEvents
    End With ' application

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MergedAreaBorders1Save(ByVal mab_range As Range, _
                                   ByRef mab_dict As Dictionary)
' ------------------------------------------------------------------------------
' Saves border properties of the range (mab_range) in Dictionary (mab_dict).
' ------------------------------------------------------------------------------
    Const PROC = "MergedAreaBorders1Save"
    
    On Error GoTo eh
    Dim cll     As Collection
    Dim xlBi    As XlBordersIndex
    
    Set mab_dict = Nothing
    Set mab_dict = New Dictionary
    For xlBi = xlDiagonalDown To xlInsideHorizontal ' = 5 to 12
        With mab_range.Borders(xlBi)
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
        mab_dict.Add xlBi, cll
    Next xlBi
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MergedAreaBorders2Restore(ByVal mab_range As Range, _
                                      ByRef mab_dict As Dictionary)
' ------------------------------------------------------------------------------
' Restores border properties for the range (mab_range) obtained from the
' Dictionary (mab_dict).
' ------------------------------------------------------------------------------
    Const PROC = "MergedAreaBorders2Restore"
    
    On Error GoTo eh
    Dim cll     As Collection
    Dim xlBi    As XlBordersIndex
    Dim i       As Long
    
    For i = 0 To mab_dict.Count - 1
        xlBi = mab_dict.Keys()(i)
        Set cll = mab_dict.Items()(i)
        
        With mab_range.Borders(xlBi)
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
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MergedAreas1SaveAndUnMerge(ByVal ma_source As Variant)
' ------------------------------------------------------------------------------
' Pushes a Dictionary on the Worksheet specific Merged Areas stack. Each item of
' the Dictionary is a Dictionary of border properties of a concerned Merged Area
' and a range name of it as the key. Concerned are Merged Areas which are in one
' of the rows or columns of the provided range (ma_source).
' ------------------------------------------------------------------------------
    Const PROC = "MergedAreas1SaveAndUnMerge"
    
    On Error GoTo eh
    Dim dct                 As Dictionary
    Dim dctBorders          As Dictionary
    Dim rMergeArea          As Range
    Dim RangeNameMergedArea As String
    Dim ws                  As Worksheet
    Dim cel                 As Range
    Dim stack               As Collection
    Dim r                   As Range
    Dim sMergeArea          As String
    Dim i                   As Long
    
    Set r = ma_source
    Set ws = r.Worksheet
    SheetProtection obs_mode:=enEliminate, obs_ws:=ws
    ApplEvents obs_mode:=enEliminate ' avoid any interference with Worksheet_Change actions
    
    If BasicStackIsEmpty(SheetStack(MergedAreasSheetStacks, ws)) _
    Then RangeNamesRemove nr_wb:=ws.Parent, nr_ws:=ws, nr_generic_name:=TEMP_MERGED_AREA_NAME
    
    Set dct = Nothing: Set dct = New Dictionary
    For Each cel In Intersect(ws.UsedRange, r.EntireRow).Cells
        With cel
            If .MergeCells Then
                If sMergeArea <> Replace(.MergeArea.Address(RowAbsolute:=False), "$", vbNullString) Then
                    i = i + 1
                    RangeNameMergedArea = TempMergedAreaName(ws) & i
                    sMergeArea = Replace(.MergeArea.Address(RowAbsolute:=False), "$", vbNullString)
                    Set rMergeArea = Range(sMergeArea)
                    RangeNameAdd nm_ws:=ws, nm_wb:=ws.Parent, nm_name:=RangeNameMergedArea, nm_range:=rMergeArea
                    MergedAreaBorders1Save rMergeArea, dctBorders   ' Get format properties
                    dct.Add RangeNameMergedArea, dctBorders            ' Save range name and border properties to dictionary
                    MergedArea1UnMerge rMergeArea
                End If
            End If
        End With
    Next cel
    Set stack = SheetStack(MergedAreasSheetStacks, ws)
    If i = 0 _
    Then BasicStackPush stack, vbNullString _
    Else BasicStackPush stack, dct
    SheetStack(MergedAreasSheetStacks, ws) = stack
    
    ApplEvents obs_mode:=enRestore
    SheetProtection obs_mode:=enRestore, obs_ws:=ws

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub MergedAreas2Restore(ByVal ma_source As Variant)
' ------------------------------------------------------------------------------
' ------------------------------------------------------------------------------
    Const PROC = "MergedAreas"
    
    On Error GoTo eh
    Dim dct                 As Dictionary
    Dim dctBorders          As Dictionary
    Dim rMergeArea          As Range
    Dim RangeNameMergedArea As String
    Dim ws                  As Worksheet
    Dim i                   As Long
    
    Set dct = ma_source
    If Err.Number = 0 Then
        For i = dct.Count - 1 To 0 Step -1                              ' bottom up to allow removing processed items
            If TypeName(dct.Keys()(i)) = "String" Then                  ' A string typ key indicates a merge area's range name
                RangeNameMergedArea = dct.Keys()(i)                     ' Get range name
                Set dctBorders = dct.Items()(i)                         ' Get border properties
                Set rMergeArea = Range(RangeNameMergedArea)             ' Set the to-be-merged range
                Set ws = rMergeArea.Worksheet
                ApplEvents obs_mode:=enEliminate                      ' avoid any interference with Worksheet_Change actions
                
                MergedArea2ReMerge rMergeArea                           ' Merge the range named RangeNameMergedArea
                MergedAreaBorders2Restore rMergeArea, dctBorders        ' CleanUp the border proprties
                RangeNamesRemove nr_wb:=ws.Parent _
                               , nr_ws:=ws _
                               , nr_generic_name:=RangeNameMergedArea   ' Remove the no longer required range name
                               
                ApplEvents obs_mode:=enRestore
            End If
        Next i
    End If

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub All(ByVal obs_mode As enObstService, _
               ByVal obs_ws As Worksheet, _
      Optional ByVal obs_range As Range = Nothing, _
      Optional ByVal obs_application_events As Boolean = True, _
      Optional ByVal obs_filtered_rows_hidden_cols As Boolean = True, _
      Optional ByVal obs_merged_cells As Boolean = True, _
      Optional ByVal obs_sheet_protection As Boolean = True)
' ------------------------------------------------------------------------------
' Obstructions 'All' services in the sense all but ....
' Eliminate: Eliminates all obstruction not explicitely denied by the
'            corresponding argument
' Restore:   Restores all obstructions not explicitely denied by the
'            corresponding argument
' Attention: Eliminate and Restore service need to be exactly paired!
'
' W. Rauschenberger Berlin, Dec 2021
' ------------------------------------------------------------------------------
    Const PROC = "All"
    
    On Error GoTo eh
    Select Case obs_mode
        Case enEliminate
        
            If obs_application_events _
            Then ApplEvents enEliminate
            
            If obs_sheet_protection _
            Then SheetProtection obs_mode:=enEliminate, obs_ws:=obs_ws
            
            If obs_filtered_rows_hidden_cols _
            Then FilteredRowsHiddenCols enEliminate, obs_ws
            
            If obs_merged_cells _
            Then
                '~~ Ensure the provided obs_range argument is a range of the provided Worksheet
                '~~ and throw an error if not
                If obs_range Is Nothing _
                Then Err.Raise AppErr(1), ErrSrc(PROC), _
                               "The range argument (obs_range) required for the 'FilteredRowsHiddenCols' " & _
                               "obstruction service is missing when this service included in the 'All' service " & _
                               "is not explicitely denied (argument obs_filtered_rows_hidden_cols:=False)!"
                If Not obs_range.Worksheet Is obs_ws _
                Then Err.Raise AppErr(1), ErrSrc(PROC), _
                               "The provided range (obs_range) is not one of the provided Worksheet (obs_ws)!" & "||" & _
                               "The '" & ErrSrc(PROC) & "' service only manages merged areas which relate to " & _
                               "a provided range's rows and columns. The argument 'obs_range' is optional only " & _
                               "for all obstruction services but this one - for which the argument 'obs_merged_cells' " & _
                               "defaults to TRUE"

                MergedAreas obs_mode:=enEliminate, obs_ws:=obs_ws, obs_range:=obs_range
            End If
        Case enRestore
                        
            If obs_merged_cells _
            Then MergedAreas obs_mode:=enRestore, obs_ws:=obs_ws, obs_range:=obs_range
            
            If obs_filtered_rows_hidden_cols _
            Then FilteredRowsHiddenCols obs_mode:=enRestore, obs_ws:=obs_ws
            
            If obs_sheet_protection _
            Then SheetProtection obs_mode:=enRestore, obs_ws:=obs_ws
            
            If obs_application_events _
            Then ApplEvents enRestore
            
    End Select

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Rewind()
' ------------------------------------------------------------------------------
' Rewinds all set-off obstructions in order to have all Worksheets finally
' reset to their original status. This is definitely helpfull in case of an
' error which would leave obstructions un-restored otherwise.
' Rewind considers more than one concerned Worksheet and rewinds them in reverse
' order (ws1, ws2, ws3 - ws3, ws2, ws1). This is primarily essential when
' filtered rows and/or hidden columns are restored because this is done by means
' of Custom Views. Though any Custom View has a Workbook scope CustomViews are
' stored per Worksheet.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
    Const PROC = "Rewind"
    
    On Error GoTo eh
    Dim ws                  As Worksheet
    Dim stack               As Collection
    Dim i                   As Long
    
    '~~ Restore Filtered Rows and/or Hidden Columns for all Worksheets in revers order !!
    If Not FilteredRowsHiddenColsSheetStacks Is Nothing Then
        For i = FilteredRowsHiddenColsSheetStacks.Count - 1 To 0 Step -1
            Set ws = FilteredRowsHiddenColsSheetStacks.Keys()(i)
            Set stack = SheetStack(FilteredRowsHiddenColsSheetStacks, ws)
            While Not BasicStackIsEmpty(stack)
                FilteredRowsHiddenCols2Restore obs_ws:=ws, obs_stack:=stack
                SheetStack(FilteredRowsHiddenColsSheetStacks, ws) = stack
                Set stack = SheetStack(FilteredRowsHiddenColsSheetStacks, ws)
            Wend
        Next i
    End If
    
    '~~ Restore Sheet(s) Protection
    If Not SheetProtectionSheetStacks Is Nothing Then
        For i = SheetProtectionSheetStacks.Count - 1 To 0 Step -1
            Set ws = SheetProtectionSheetStacks.Keys()(i)
            Set stack = SheetProtectionSheetStacks(ws)
            While Not BasicStackIsEmpty(stack)
                SheetProtection obs_mode:=enRestore, obs_ws:=ws
            Wend
        Next i
    End If
    
    '~~ Restore Application EnableEvents (allways done silent)
    If Not ApplEventsStack Is Nothing Then
        While Not BasicStackIsEmpty(ApplEventsStack)
            Application.EnableEvents = BasicStackPop(ApplEventsStack)
        Wend
    End If
    
    '~~ Restore Merged Cells/Areas
    If Not MergedAreasSheetStacks Is Nothing Then
        For i = MergedAreasSheetStacks.Count - 1 To 0 Step -1
            Set ws = MergedAreasSheetStacks.Keys()(i)
            SheetProtection enEliminate, ws
            
            Set stack = SheetStack(MergedAreasSheetStacks, ws)
            While Not BasicStackIsEmpty(stack)
                MergedAreas2Restore BasicStackPop(stack) ' popped is a Dictionary of merge area range names
                SheetStack(MergedAreasSheetStacks, ws) = stack
                Set stack = SheetStack(MergedAreasSheetStacks, ws)
            Wend
            SheetProtection enRestore, ws
        
        Next i
    End If
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub RangeNameAdd(ByVal nm_ws As Worksheet, _
                         ByVal nm_name As String, _
                         ByVal nm_range As Range, _
                Optional ByVal nm_wb As Workbook = Nothing, _
                Optional ByVal nm_visible As Boolean = True)
' ------------------------------------------------------------------------------
' Add the name (nm_name) for the range (nm_range)
' a) with Worksheet scope when no Workbook is provided
' b) with Workbook scope when a Workbook is provided
' ------------------------------------------------------------------------------
    Const PROC = "RangeNameAdd"
    
    On Error GoTo eh
    Dim nms As Names
    
    If nm_wb Is Nothing Then Set nms = nm_ws.Names Else Set nms = nm_wb.Names
    nms.Add Name:=nm_name, RefersTo:=nm_ws.Range(nm_range.Address), Visible:=nm_visible
    If Err.Number <> 0 Then
         Err.Raise AppErr(3), ErrSrc(PROC), "Adding a name failed with: Error " & Err.Number & " " & Err.Description & """"
    End If
        
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub RangeNamesRemove(ByVal nr_ws As Worksheet, _
                             ByVal nr_generic_name As String, _
                    Optional ByVal nr_wb As Workbook = Nothing, _
                    Optional ByRef nr_deleted As Long = 0, _
                    Optional ByRef nr_failed As Long = 0)
' ------------------------------------------------------------------------------
' Deletes all range names in Workbook (nr_wb)
' a) which do begin with the generic name (nr_generic_name) when a Workbook is
'    provided (Workbook scope names)
' b) which do begin with "(nr_ws).Name!(generic_name_part)" when no Workbook
'    is provided (Worksheet scope names)
' Note 1: Application.Names and Names both refer to ActiveWorkbook.Names. The
'         argument nr_wb avoids any confusion between ActiveWorkbook,
'         ThisWorkbok or any other Workbook.
' Note 2: Names with Worksheet scope are prefixed with "'(nr_ws).Name'!".
' The procedure is Public because it is also used in the wbRowsTest code.
' ------------------------------------------------------------------------------
    Const PROC          As String = "RangeNamesRemove"
    
    On Error GoTo eh
    Dim nm              As Name
    Dim nms             As Names
    Dim sName           As String
    Dim sComparewith    As String
    
    nr_deleted = 0
    nr_failed = 0
    
    '~~ Determine which scope is ment
    If nr_wb Is Nothing Then
        '~~ Worksheet scope
        sComparewith = "'!" & nr_generic_name
        Set nms = nr_ws.Names
    Else
        '~~ Workbook scope
        sComparewith = nr_generic_name
        Set nms = nr_wb.Names
    End If
    
    '~~ Delete the names in the ment scope
    For Each nm In nms
        sName = nm.Name
        If Len(sName) >= Len(sComparewith) Then
            If Left(sName, Len(sComparewith)) = sComparewith _
            Or InStr(sName, sComparewith) <> 0 Then
                '~~ Delete the generic name
                On Error Resume Next
                nm.Delete           ' Delete the name
                If Err.Number = 0 Then
                    nr_deleted = nr_deleted + 1
                Else
                    nr_failed = nr_failed + 1
                End If
            End If
        End If
nx: Next nm
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

