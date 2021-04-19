Attribute VB_Name = "mCstmVw"
Option Explicit

Public Function Exists(ByVal vWb As Variant, _
                       ByVal vCv As Variant, _
              Optional ByRef cvResult As CustomView) As Boolean
' -------------------------------------------------------------
' Returns TRUE when the CustomView (vCv) - which may be a
' CustomView object or a CustoView's name - exists in the
' Workbook (wb). If vCv is provided as CustomView object, only
' its name is used for the existence check in Workbook (wb).
' -------------------------------------------------------------
Const PROC  As String = "CustomViewExists"      ' This procedure's name for the error handling and execution tracking
Dim wb      As Workbook
Dim sTest   As String

    On Error GoTo on_error
    BoP ErrSrc(PROC)
    Exists = False
    
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
    
    If Not mCommon.IsCvObject(vCv) And Not mCommon.IsCvName(vCv) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The CustomView (vCv) is neither a string (CustomView's name) nor a CustomView object!"
    
    If mCommon.IsCvObject(vCv) Then
        On Error Resume Next
        sTest = vCv.name
        Exists = Err.Number = 0
        GoTo exit_proc
    ElseIf mCommon.IsCvName(vCv) Then
        On Error Resume Next
        sTest = wb.CustomViews(vCv).name
        Exists = Err.Number = 0
        GoTo exit_proc
    End If
  
exit_proc:
    EoP ErrSrc(PROC)
    Exit Function
    
on_error:
#If Debugging = 1 Then
    Stop: Resume
#End If
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCstmVw" & "." & sProc
End Function
