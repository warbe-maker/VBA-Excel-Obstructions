Attribute VB_Name = "mCommon"
Option Private Module
Option Explicit
#Const CustomMsg = 1 ' = 1 when the fMsg UserForm is to be used instead of the MsgBox
' ------------------------------------------------------------
' Standard Module mCommon Contains declarations and
'          methods common for allmost every VB-Project.
' Methods:
' - AppErr          Converts a positive number into a negative
'                   error number ensuring it not conflicts with
'                   and VB error. A negative error number is
'                   turned back into the original positive
'                   Application Error Number.
' - ArrayCompare    Compares two one-dimensional arrays.
'                   Returns an array with all different
'                   items
' - ArrayIsAllocated    Returns TRUE when the provided array
'                       has at least one item
' - ArrayNoOfDims   Returns the number of dimensions of an
'                   array.
' - ArrayRemoveItem Removes an array's item either by its
'                   index or its element number
' - ArrayToRange    Transferres the content of a one- or two-
'                   dimensional arry to a range
' - ArrayTrim       Removes any leading or trailing empty items.
' - CleanTrim       Clears a string from any unprinable characters.
'
' Requires: Reference to "Microsoft Scripting Runtime"
'
' W. Rauschenberger Berlin March 2020
' ------------------------------------------------------------
' --------------------------------------------------------
' Module mBscs  Basics potentially uesefull in any project
' --------------------------------------------------------
Private Declare PtrSafe Function GetSystemMetrics32 _
  Lib "user32" Alias "GetSystemMetrics" _
    (ByVal nIndex As Long) As Long

'Functions to get DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const LOGPIXELSX = 88               ' Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72  ' A point is defined as 1/72 inches
Private Declare PtrSafe Function GetForegroundWindow _
  Lib "User32.dll" () As Long

Private Declare PtrSafe Function GetWindowLongPtr _
  Lib "User32.dll" Alias "GetWindowLongA" _
    (ByVal hWnd As LongPtr, _
     ByVal nIndex As Long) _
  As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr _
  Lib "User32.dll" Alias "SetWindowLongA" _
    (ByVal hWnd As LongPtr, _
     ByVal nIndex As LongPtr, _
     ByVal dwNewLong As LongPtr) _
  As LongPtr

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16
                
Public Const MSG_CONCAT As String = "||"    ' For concatenating and error with a general message (info) to the error description
Public Const LINE_LEFT  As String = ">"
Public Const LINE_RIGHT As String = "<"
Public Const sApost     As String = "'"
Public Const sQuote     As String = """"
Public Const sKomma     As String = ","
Public Const sBslash    As String = "\"
Public Const sDot       As String = "."
Public Const sColon     As String = ":"
Public Const sEqu       As String = "="
Public Const sSpace     As String = " "
Public Const sExcl      As String = "!"

Private vMsgReply       As Variant

' Common xl constants grouped ----------------------------
Public Enum YesNo   ' ------------------------------------
    xlYes = 1       ' System constants (identical values)
    xlNo = 2        ' grouped for being used as Enum Type.
End Enum            ' ------------------------------------
Public Enum xlOnOff ' ------------------------------------
    xlOn = 1        ' System constants (identical values)
    xlOff = -4146   ' grouped for being used as Enum Type.
End Enum            ' ------------------------------------

Public Enum enDctMode ' Dictionary add/insert modes
    dct_addafter = 1
    dct_addbefore = 2
    dct_ascending = 3
    dct_ascendingignorecase = 4
    dct_descending = 5
    dct_descendingignorecase = 6
    dct_sequence = 7
End Enum


Public Property Let MsgReply(ByVal v As Variant):   vMsgReply = v:          End Property
Public Property Get MsgReply() As Variant:          MsgReply = vMsgReply:   End Property

Public Function Max(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = 0, _
           Optional ByVal v4 As Variant = 0, _
           Optional ByVal v5 As Variant = 0, _
           Optional ByVal v6 As Variant = 0, _
           Optional ByVal v7 As Variant = 0, _
           Optional ByVal v8 As Variant = 0, _
           Optional ByVal v9 As Variant = 0) As Variant
' -----------------------------------------------------
' Returns the maximum (biggest) of all provided values.
' -----------------------------------------------------
Dim dMax As Double
    dMax = v1
    If v2 > dMax Then dMax = v2
    If v3 > dMax Then dMax = v3
    If v4 > dMax Then dMax = v4
    If v5 > dMax Then dMax = v5
    If v6 > dMax Then dMax = v6
    If v7 > dMax Then dMax = v7
    If v8 > dMax Then dMax = v8
    If v9 > dMax Then dMax = v9
    Max = dMax
End Function

Public Function Min(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = Nothing, _
           Optional ByVal v4 As Variant = Nothing, _
           Optional ByVal v5 As Variant = Nothing, _
           Optional ByVal v6 As Variant = Nothing, _
           Optional ByVal v7 As Variant = Nothing, _
           Optional ByVal v8 As Variant = Nothing, _
           Optional ByVal v9 As Variant = Nothing) As Variant
' ------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------
Dim dMin As Double
    dMin = v1
    If v2 < dMin Then dMin = v2
    If TypeName(v3) <> "Nothing" Then If v3 < dMin Then dMin = v3
    If TypeName(v4) <> "Nothing" Then If v4 < dMin Then dMin = v4
    If TypeName(v5) <> "Nothing" Then If v5 < dMin Then dMin = v5
    If TypeName(v6) <> "Nothing" Then If v6 < dMin Then dMin = v6
    If TypeName(v7) <> "Nothing" Then If v7 < dMin Then dMin = v7
    If TypeName(v8) <> "Nothing" Then If v8 < dMin Then dMin = v8
    If TypeName(v9) <> "Nothing" Then If v9 < dMin Then dMin = v9
    Min = dMin
End Function

Public Function BaseName(ByVal v As Variant) As String
' -----------------------------------------------------
' Returns the file name without the extension. v may be
' a file name a file path (full name) a File object or
' a Workbook object.
' -----------------------------------------------------
Const PROC  As String = "BaseName"

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    
    Select Case TypeName(v)
        Case "String":      With New FileSystemObject:  BaseName = .GetBaseName(v):             End With
        Case "Workbook":    With New FileSystemObject:  BaseName = .GetBaseName(v.FullName):    End With
        Case "File":        With New FileSystemObject:  BaseName = .GetBaseName(v.ShortName):  End With
        Case Else:          Err.Raise AppErr(1), ErrSrc(PROC), "The parameter (v) is neither a string nor a File or Workbook object (TypeName = '" & TypeName(v) & "')!"
    End Select

exit_proc:
    EoP ErrSrc(PROC)
    Exit Function
    
on_error:
#If Debugging Then
    Stop: Resume
#End If
    ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Public Function CleanTrim(ByVal s As String, _
                 Optional ByVal ConvertNonBreakingSpace As Boolean = True) As String
' ----------------------------------------------------------------------------------
' Returns the string 's' cleaned from any non-printable characters.
' ----------------------------------------------------------------------------------
Dim l           As Long
Dim asToClean   As Variant
    
    asToClean = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
                     21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 127, 129, 141, 143, 144, 157)
    If ConvertNonBreakingSpace Then s = Replace(s, Chr$(160), " ")
    For l = LBound(asToClean) To UBound(asToClean)
        If InStr(s, Chr$(asToClean(l))) Then s = Replace(s, Chr$(asToClean(l)), vbNullString)
    Next
    CleanTrim = s

End Function


Public Sub DictAdd(ByRef dct As Dictionary, _
                   ByRef vKey As Variant, _
                   ByRef vItem As Variant, _
                   ByVal lMode As enDctMode, _
          Optional ByVal vTarget As Variant)
' --------------------------------------------------
' Universal method to add an item to the Dictionary
' (dct), supporting ascending and descending order,
' case and case insensitive as well as adding items
' before or after an existing item.
' - When the key (vKey) already exists adding will
'   just be skipped without a warning or an error.
' - When the dictionary (dct) is Nothing it is setup
'   on the fly.
'
' W. Rauschenberger, Berlin Mar 2015
' --------------------------------------------------
Dim dctTemp     As Dictionary
Dim vTempKey    As Variant
Dim bAdd        As Boolean

    If dct Is Nothing Then Set dct = New Dictionary
    
    With dct
        If .Count = 0 Or lMode = dct_sequence Then
            '~~> Very first item is just added
            .Add vKey, vItem
            Exit Sub
        Else
            ' ------------------------------------
            ' Let's see whether the new key can be
            ' added directly after the last key
            ' ------------------------------------
            vTempKey = .Keys()(.Count - 1)
            Select Case lMode
                Case dct_ascending
                    If vKey > vTempKey Then
                        .Add vKey, vItem
                        Exit Sub                ' Done!
                    End If
                Case dct_ascendingignorecase
                    If LCase$(vKey) > LCase$(vTempKey) Then
                        .Add vKey, vItem
                        Exit Sub                ' Done!
                    End If
                Case dct_descending
                    If vKey < vTempKey Then
                        .Add vKey, vItem
                        Exit Sub                ' Done!
                    End If
                Case dct_descendingignorecase
                    If LCase$(vKey) < LCase$(vTempKey) Then
                        .Add vKey, vItem
                        Exit Sub                ' Done!
                    End If
            End Select
        End If
    End With

    ' ----------------------------------------------
    ' Since the new key could not simply be added
    ' to the dct it must be added/inserted somewhere
    ' in between or even before the very first key.
    ' ----------------------------------------------
    Set dctTemp = New Dictionary
    bAdd = True
    For Each vTempKey In dct
        With dctTemp
            If bAdd Then
                '~~> Skip this section when already added
                If dct.Exists(vKey) Then
                    '~~> Simply ignore add when already existing
                    bAdd = False
                    Exit Sub
                End If
                Select Case lMode
                    Case dct_ascending
                        If vTempKey > vKey Then
                            .Add vKey, vItem
                            bAdd = False ' Add done
                        End If
                    Case dct_ascendingignorecase
                        If LCase$(vTempKey) > LCase$(vKey) Then
                            .Add vKey, vItem
                            bAdd = False ' Add done
                        End If
                    Case dct_addbefore
                        If vTempKey = vTarget Then
                            '~~> Add before vTarget key has been reached
                            .Add vKey, vItem
                            bAdd = True
                        End If
                    Case dct_descending
                        If vTempKey < vKey Then
                            .Add vKey, vItem
                            bAdd = False ' Add done
                        End If
                    Case dct_descendingignorecase
                        If LCase$(vTempKey) < LCase$(vKey) Then
                            .Add vKey, vItem
                            bAdd = False ' Add done
                        End If
                End Select
            End If
            
            '~~> Transfer the existing item to the temporary dictionary
            .Add vTempKey, dct.Item(vTempKey)
            
            If lMode = dct_addafter And bAdd Then
                If vTempKey = vTarget Then
                    ' ----------------------------------------
                    ' Just add when lMode indicates add after,
                    ' and the vTraget key has been reached
                    ' ----------------------------------------
                    .Add vKey, vItem
                    bAdd = False
                End If
            End If
            
        End With
    Next vTempKey
    
    '~~> Return the temporary dictionary with the new item added
    Set dct = dctTemp
    Set dctTemp = Nothing
    Exit Sub

End Sub

Public Sub ArrayTrimm(ByRef a As Variant)
' ---------------------------------------
' Return the array (a) with all leading
' and trailing blank items removed. Any
' vbCr, vbCrLf, vbLf are ignored.
' When the array contains only blank
' items the returned array is erased.
' ---------------------------------------
Const PROC  As String = "ArrayTrimm"
Dim i       As Long

    On Error GoTo on_error
    
    '~~ Eliminate leading blank lines
    If Not mCommon.ArrayIsAllocated(a) Then Exit Sub
    
    Do While (Len(Trim$(a(LBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
        mCommon.ArrayRemoveItems a, Index:=i
        If Not mCommon.ArrayIsAllocated(a) Then Exit Do
    Loop
    
    If mCommon.ArrayIsAllocated(a) Then
        Do While (Len(Trim$(a(UBound(a)))) = 0 Or Trim$(a(LBound(a))) = " ") And UBound(a) >= 0
            If UBound(a) = 0 Then
                Erase a
            Else
                ReDim Preserve a(UBound(a) - 1)
            End If
            If Not mCommon.ArrayIsAllocated(a) Then Exit Do
        Loop
    End If
exit_proc:
    Exit Sub
    
on_error:
    '~~ Global error handling is used to seamlessly monitor error conditions
#If Debugging Then
    Stop: Resume
#End If
    mCommon.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Public Sub ArrayRemoveItems(ByRef va As Variant, _
                   Optional ByVal Element As Variant, _
                   Optional ByVal Index As Variant, _
                   Optional ByVal NoOfElements = 1)
' ------------------------------------------------------
' Returns the array (va) with the number of elements
' (NoOfElements) removed whereby the start element may be
' indicated by the element number 1,2,... (vElement) or
' the index (Index) which must be within the array's
' LBound to Ubound.
' Any inapropriate provision of the parameters results
' in a clear error message.
' When the last item in an array is removed the returned
' arry is erased (no longer allocated).
'
' Restriction: Works only with one dimensional array.
'
' W. Rauschenberger, Berlin Jan 2020
' ------------------------------------------------------
Const PROC              As String = "ArrayRemoveItems"
Dim a                   As Variant
Dim iElement            As Long
Dim iIndex              As Long
Dim NoOfElementsInArray    As Long
Dim i                   As Long
Dim iNewUBound          As Long

    On Error GoTo on_error
    
    If Not IsArray(va) Then
        Err.Raise AppErr(1), ErrSrc(PROC), "Array not provided!"
    Else
        a = va
        NoOfElementsInArray = UBound(a) - LBound(a) + 1
    End If
    If Not ArrayNoOfDims(a) = 1 Then
        Err.Raise AppErr(2), ErrSrc(PROC), "Array must not be multidimensional!"
    End If
    If Not IsNumeric(Element) And Not IsNumeric(Index) Then
        Err.Raise AppErr(3), ErrSrc(PROC), "Neither FromElement nor FromIndex is a numeric value!"
    End If
    If IsNumeric(Element) Then
        iElement = Element
        If iElement < 1 _
        Or iElement > NoOfElementsInArray Then
            Err.Raise AppErr(4), ErrSrc(PROC), "vFromElement is not between 1 and " & NoOfElementsInArray & " !"
        Else
            iIndex = LBound(a) + iElement - 1
        End If
    End If
    If IsNumeric(Index) Then
        iIndex = Index
        If iIndex < LBound(a) _
        Or iIndex > UBound(a) Then
            Err.Raise AppErr(5), ErrSrc(PROC), "FromIndex is not between " & LBound(a) & " and " & UBound(a) & " !"
        Else
            iElement = ElementOfIndex(a, iIndex)
        End If
    End If
    If iElement + NoOfElements - 1 > NoOfElementsInArray Then
        Err.Raise AppErr(6), ErrSrc(PROC), "FromElement (" & iElement & ") plus the number of elements to remove (" & NoOfElements & ") is beyond the number of elelemnts in the array (" & NoOfElementsInArray & ")!"
    End If
    
    For i = iIndex + NoOfElements To UBound(a)
        a(i - NoOfElements) = a(i)
    Next i
    
    iNewUBound = UBound(a) - NoOfElements
    If iNewUBound < 0 Then Erase a Else ReDim Preserve a(LBound(a) To iNewUBound)
    va = a
    
exit_proc:
    Exit Sub

on_error:
    '~~ Global error handling is used to seamlessly monitor error conditions
#If Debugging Then
    Stop: Resume
#End If
    mCommon.ErrMsg Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Public Function ElementOfIndex(ByVal a As Variant, _
                                ByVal i As Long) As Long
' ------------------------------------------------------
' Returns the element number of index (i) in array (a).
' ------------------------------------------------------
Dim ia  As Long
    For ia = LBound(a) To i
        ElementOfIndex = ElementOfIndex + 1
    Next ia
End Function

Public Sub ArrayToRange(ByVal vArr As Variant, _
                        ByVal r As Range, _
               Optional ByVal bOneCol As Boolean = False)
' -------------------------------------------------------
' Copy the content of the Arry (vArr) to the range (r).
' -------------------------------------------------------
Dim rTarget As Range

    If bOneCol Then
        '~~ One column, n rows
        Set rTarget = r.Cells(1, 1).Resize(UBound(vArr), 1)
        rTarget.Value = Application.Transpose(vArr)
    Else
        '~~ One column, n rows
        Set rTarget = r.Cells(1, 1).Resize(1, UBound(vArr))
        rTarget.Value = vArr
    End If
    
End Sub

Public Function ArrayCompare(ByVal a1 As Variant, _
                             ByVal a2 As Variant) As Variant
' ----------------------------------------------------------
' Returns an array of all lines which are different. Each
' element contains the corresponding elements of both arrays
' in form: linenumber: <line>||<line>
' ----------------------------------------------------------
Dim l               As Long
Dim i               As Long
Dim va()            As Variant

    l = 0
    For i = LBound(a1) To Min(UBound(a1), UBound(a2))
        If a1(i) <> a2(i) Then
            ReDim Preserve va(l)
            va(l) = i & ": " & LINE_LEFT & a1(i) & LINE_RIGHT & MSG_CONCAT & LINE_LEFT & a2(i) & LINE_RIGHT
            l = l + 1
        End If
    Next i
    If UBound(a1) < UBound(a2) Then
        For i = UBound(a1) + 1 To UBound(a2)
            ReDim Preserve va(l)
            va(l) = i & ": " & LINE_LEFT & LINE_RIGHT & MSG_CONCAT & LINE_LEFT & a2(i) & LINE_RIGHT
            l = l + 1
        Next i
    ElseIf UBound(a2) < UBound(a1) Then
        For i = UBound(a2) + 1 To UBound(a1)
            ReDim Preserve va(l)
            va(l) = i & ": " & LINE_LEFT & a1(i) & LINE_RIGHT & MSG_CONCAT & LINE_LEFT & LINE_RIGHT
            l = l + 1
        Next i
    End If

    ArrayCompare = va
    
End Function

Public Function ArrayNoOfDims(arr As Variant) As Integer
' ------------------------------------------------------
' Returns the number of dimensions of an array. An un-
' allocated dynamic array has 0 dimensions. This may as
' as well be tested by means of ArrayIsAllocated.
' ------------------------------------------------------
Dim Ndx As Integer
Dim Res As Integer

    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    ArrayNoOfDims = Ndx - 1

End Function

Public Function ArrayIsAllocated(arr As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = IsArray(arr) And _
                       Not IsError(LBound(arr, 1)) And _
                       LBound(arr, 1) <= UBound(arr, 1)
End Function

Public Function IsCvName(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then IsCvName = True
End Function

Public Function IsCvObject(ByVal v As Variant) As Boolean

    If VarType(v) = vbObject Then
        If Not TypeName(v) = "Nothing" Then
            IsCvObject = TypeOf v Is CustomView
        End If
    End If
    
End Function

Public Function IsPath(ByVal v As Variant) As Boolean
    
    If VarType(v) = vbString Then
        If InStr(v, "\") <> 0 Then
            If InStr(Right(v, 6), ".") = 0 Then
                IsPath = True
            End If
        End If
    End If

End Function

#If CustomMsg Then
Public Sub ErrMsg(Optional ByVal lErrNo As Long = 0, _
                  Optional ByVal sErrSrc As String = vbNullString, _
                  Optional ByVal sErrDesc As String = vbNullString, _
                  Optional ByVal sErrLine As String = vbNullString, _
                  Optional ByVal sTitle As String = vbNullString, _
                  Optional ByVal sErrPath As String = vbNullString, _
                  Optional ByVal sErrInfo As String = vbNullString)
' -------------------------------------------------------------------
' Common error message using fMsg which supports the
' display of an error path in a fixed font textbox.
' -------------------------------------------------------------------
Const PROC      As String = "ErrMsg"
Dim sIndicate   As String
Dim sErrText    As String   ' May be a first part of the sErrDesc

    If lErrNo = 0 _
    Then MsgBox "Exit statement before error handling missing! Error number is 0!", vbCritical, "Application error in " & ErrSrc(PROC) & "!"
    
    '~~ Additional info about the error line in case one had been provided
    If sErrLine = vbNullString Or sErrLine = "0" Then
        sIndicate = vbNullString
    Else
        sIndicate = " (at line " & sErrLine & ")"
    End If
    
    If sTitle = vbNullString Then
        '~~ When no title is provided, one is assembled by provided info
        sTitle = sTitle & sIndicate
        '~~ Distinguish between VBA and Application error
        Select Case lErrNo
            Case Is > 0:    sTitle = "VBA Error " & lErrNo
            Case Is < 0:    sTitle = "Application Error " & AppErr(lErrNo)
        End Select
        sTitle = sTitle & " in:  " & sErrSrc & sIndicate
    End If
    
    If sErrInfo = vbNullString Then
        '~~ When no error information is provided one may be within the error description
        '~~ which is only possible with an application error raised by Err.Raise
        If InStr(sErrDesc, MSG_CONCAT) <> 0 Then
            sErrText = Split(sErrDesc, MSG_CONCAT)(0)
            sErrInfo = Split(sErrDesc, MSG_CONCAT)(1)
        Else
            sErrText = sErrDesc
            sErrInfo = vbNullString
        End If
    Else
        sErrText = sErrDesc
    End If
    
    '~~ Display error message by UserForm fErrMsg
    With fMsg
        .Title = sTitle
        .TitleFontName = "Tahoma"
        .TitleFontSize = 9
'        .ErrSrc = sErrSrc & sIndicate
        .LabelMessage1 = "Error Message/Description:"
        .Message1Proportional = sErrText
        If sErrPath <> vbNullString Then
            .LabelMessage2 = "Error path (call stack):"
            .Message2Proportional = sErrPath
        End If
        If sErrInfo <> vbNullString Then
            .LabelMessage3 = "Info:"
            .Message3Proportional = sErrInfo
        End If
        .Replies = vbOKOnly
        .Show
    End With

End Sub
#Else

Public Sub ErrMsg(ByVal lErrNo As Long, _
                  ByVal sErrSource As String, _
                  ByVal sErrText As String, _
                  ByVal sErrLine As String)
' ---------------------------------------------
' Common error message using MsgBox.
' ---------------------------------------------
Const PROC  As String = "ErrMsg"
Dim sMsg    As String
Dim sTitle  As String

    If lErrNo = 0 _
    Then MsgBox "Exit statement before error handling missing! Error number is 0!", vbCritical, "Application error in " & ErrSrc(PROC) & "!"
    
    '~~ Prepare Title
    If lErrNo < 0 Then
        sTitle = "Application Error " & AppErr(lErrNo)
    Else
        sTitle = "VB Error " & lErrNo
    End If
    sTitle = sTitle & " in " & sErrSource
    If sErrLine <> 0 Then sTitle = sTitle & " (at line " & sErrLine & ")"
    
    '~~ Prepare message
    sMsg = "Error : " & sErrText & vbLf & vbLf & _
           "In : " & sErrSource
    If sErrLine <> 0 Then sMsg = sMsg & " (at line " & sErrLine & ")"
    
    MsgBox sMsg, vbCritical, sTitle

End Sub
#End If

Public Function Msg(ByVal sTitle As String, _
           Optional ByVal sMsgText As String = vbNullString, _
           Optional ByVal bFixed As Boolean = False, _
           Optional ByVal sTitleFontName As String = vbNullString, _
           Optional ByVal lTitleFontSize As Long = 0, _
           Optional ByVal siFormWidth As Single = 0, _
           Optional ByVal sReplies As Variant = vbOKOnly) As Variant
' ------------------------------------------------------------------
' Custom message using the UserForm fMsg. The function returns the
' clicked reply button's caption or the corresponding vb variable
' (vbOk, vbYes, vbNo, etc.) or its caption string.
' ------------------------------------------------------------------
Dim siDisplayHeight As Single
Dim w               As Long
Dim h               As Long
Dim siHeight        As Single

    w = GetSystemMetrics32(0) ' Screen Resolution width in points
    h = GetSystemMetrics32(1) ' Screen Resolution height in points
    
    With fMsg
        .Title = sTitle
        .TitleFontName = sTitleFontName
        .TitleFontSize = lTitleFontSize
        
        If sMsgText <> vbNullString Then
            If bFixed = True _
            Then .Message1Fixed = sMsgText _
            Else .Message1Proportional = sMsgText
        End If
        
        .Replies = sReplies
        If siFormWidth <> 0 Then .Width = Max(.Width, siFormWidth)
        .StartupPosition = 1
        .Width = w * PointsPerPixel * 0.85 'Userform width= Width in Resolution * DPI * 85%
        siHeight = h * PointsPerPixel * 0.2
        .Height = Min(.Height, siHeight)
                     
        .Show
    End With
    
    Msg = vMsgReply
End Function

Public Sub MakeFormResizable()
' ---------------------------------------------------------------------------
' This part is from Leith Ross                                              |
' Found this Code on:                                                       |
' https://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html |
'                                                                           |
' All credits belong to him                                                 |
' ---------------------------------------------------------------------------
Const WS_THICKFRAME = &H40000
Const GWL_STYLE As Long = (-16)
Dim lStyle As LongPtr
Dim hWnd As LongPtr
Dim RetVal

    hWnd = GetForegroundWindow
    
    lStyle = GetWindowLongPtr(hWnd, GWL_STYLE Or WS_THICKFRAME)
    RetVal = SetWindowLongPtr(hWnd, GWL_STYLE, lStyle)
End Sub

Public Function AppErr(ByVal lNo As Long) As Long
' -------------------------------------------------------------------------------
' Attention: This function is dedicated for being used with Err.Raise AppErr()
'            in conjunction with the common error handling module mErrHndlr when
'            the call stack is supported. The error number passed on to the entry
'            procedure is interpreted when the error message is displayed.
' The function ensures that a programmed (application) error numbers never
' conflicts with VB error numbers by adding vbObjectError which turns it into a
' negative value. In return, translates a negative error number back into an
' Application error number. The latter is the reason why this function must never
' be used with a true VB error number.
' -------------------------------------------------------------------------------
    If lNo < 0 Then
        AppErr = lNo - vbObjectError
    Else
        AppErr = vbObjectError + lNo
    End If
End Function

Public Function PointsPerPixel() As Double
' ----------------------------------------
' Return DPI
' ----------------------------------------
Dim hDC             As Long
Dim lDotsPerInch    As Long
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCommon" & "." & sProc
End Function
