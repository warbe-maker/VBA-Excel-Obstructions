Attribute VB_Name = "mCstmVw"
Option Explicit

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

Public Function IsCvObject(ByVal v As Variant) As Boolean

    If VarType(v) = vbObject Then
        If Not TypeName(v) = "Nothing" Then
            IsCvObject = TypeOf v Is CustomView
        End If
    End If
    
End Function

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
    
    If Not IsCvObject(vCv) And Not IsCvName(vCv) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The CustomView (vCv) is neither a string (CustomView's name) nor a CustomView object!"
    
    If IsCvObject(vCv) Then
        On Error Resume Next
        sTest = vCv.Name
        Exists = Err.Number = 0
        GoTo exit_proc
    ElseIf IsCvName(vCv) Then
        On Error Resume Next
        sTest = wb.CustomViews(vCv).Name
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
    ErrMsg ErrSrc(PROC)
End Function

Public Function IsCvName(ByVal v As Variant) As Boolean
    If VarType(v) = vbString Then IsCvName = True
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mCstmVw" & "." & sProc
End Function
