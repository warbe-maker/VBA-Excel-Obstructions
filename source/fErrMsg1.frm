VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fErrMsg1 
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   OleObjectBlob   =   "fErrMsg1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fErrMsg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CHAR_WIDTH        As Single = 7.5
Const MIN_FORM_WIDTH    As Single = 100
Const H_RIGHT_MARGIN    As Single = 15
Const V_MARGIN_ELEMENTS As Single = 10
Dim lErrPathLines       As Long
Dim sTitle              As String
Dim sErrSrc             As String
Dim sErrMsg             As String
Dim sErrPath            As String
Dim sErrInfo            As String

Public Property Let Title(ByVal s As String):           sTitle = s:         End Property
Public Property Let ErrSrc(ByVal s As String):          sErrSrc = s:        End Property
Public Property Let ErrMsg(ByVal s As String):          sErrMsg = s:       End Property
Public Property Let CallStack(ByVal s As String):       sErrPath = s:       End Property
Public Property Let ErrInfo(ByVal s As String):         sErrInfo = s:       End Property

Private Sub cmbOk_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
Dim sErrPathMaxLine As String   ' longest line of the error path
Dim v               As Variant
Dim siWidthErrPath  As Single
    
    With Me
        '~~ Title: The extra title label mimics the title bar,
        '~~        is autosized and determines the minimum form width
        With .laTitle
            .AutoSize = True
            .Caption = "  " & sTitle    ' some left margin
            .AutoSize = False
            .Width = .Width + H_RIGHT_MARGIN
        End With
        .laTitleSpaceBottom.Width = .laTitle.Width
        .Width = .laTitle.Width
        
        '~~ Error description: Width adjusted for now to title width
        .laErrMsg.Caption = sErrMsg
        .laErrMsg.Top = .laTitleSpaceBottom.Top + .laTitleSpaceBottom.Height + V_MARGIN_ELEMENTS
        .laErrMsgTag.Top = .laErrMsg.Top
        .laErrMsg.Width = .Width - .laErrMsg.Left - H_RIGHT_MARGIN
        
        '~~ Error Path: Adjust top position and height
        .laErrPath.Top = .laErrMsg.Top + .laErrMsg.Height + V_MARGIN_ELEMENTS
        .laErrPathTag.Top = .laErrPath.Top
        .laErrPath.Caption = sErrPath
        .laErrPath.Visible = False
        .laErrPathTag.Visible = False
        '~~ Get width
        lErrPathLines = UBound(Split(.laErrPath.Caption, vbLf)) + 1
        If lErrPathLines > 0 Then
            .laErrPathTag.Visible = True
            .laErrPathTag.Top = .laErrPath.Top
            With .laErrPath
                .Visible = True
                .Height = lErrPathLines * 11.25
            End With
            '~~ Error Path: Adjust width
            For Each v In Split(.laErrPath.Caption, vbLf)
                If Len(v) > Len(.laErrPathWidthTemplate.Caption) Then .laErrPathWidthTemplate.Caption = v
            Next v
            .laErrPath.Width = .laErrPathWidthTemplate.Width
        Else
            .laErrPath.Visible = False
            .laErrPathTag.Visible = False
        End If
        
        '~~ Adjust the form width to the maximum elements width
        .Width = mCommon.Max(MIN_FORM_WIDTH - .laErrMsg.Left, _
                             .laTitle.Width, _
                             .laErrPath.Width + .laErrPath.Left + H_RIGHT_MARGIN)
        .laTitle.Width = .Width
        .laTitleSpaceBottom.Width = .Width
        .laErrMsg.Width = .Width - .laErrMsg.Left - H_RIGHT_MARGIN
        
        '~~ Error Info: Invisible by default
        If sErrInfo <> vbNullString Then
            With Me
                .laErrInfo.Visible = True
                .laErrInfoTag.Visible = True
                '~~ Adjust top position depending on the visibility of the error path
                If .laErrPath.Visible Then
                    .laErrInfo.Top = .laErrPath.Top + .laErrPath.Height + V_MARGIN_ELEMENTS
                Else
                    .laErrInfo.Top = .laErrMsg.Top + .laErrMsg.Height + V_MARGIN_ELEMENTS
                End If
                .laErrInfo.Width = .Width - .laErrInfo.Left - H_RIGHT_MARGIN
                .laErrInfoTag.Top = .laErrInfo.Top
                '~~ Have the height adjusted by AutoSize
                .laErrInfo.Caption = sErrInfo
            End With
        End If
        
        '~~ Adjust position of the Ok button depending on the above elements visability
        If .laErrInfo.Visible = True Then
            .cmbOk.Top = .laErrInfo.Top + .laErrInfo.Height + (V_MARGIN_ELEMENTS * 2)
        ElseIf .laErrPath.Visible = True Then
            .cmbOk.Top = .laErrPath.Top + .laErrPath.Height + (V_MARGIN_ELEMENTS * 2)
        Else
            .cmbOk.Top = .laErrMsg.Top + .laErrMsg.Height + (V_MARGIN_ELEMENTS * 2)
        End If
        
        '~~ Adjust form height and width
        .Height = .cmbOk.Top + .cmbOk.Height + 30
        
        '~~ Adjust elements width to the final form width
'        .laErrMsg.Width = .Width - .laErrMsg.Left - 10
        .laErrInfo.Width = .Width - (.laErrInfo.Left + 15)
        
        '~~ Center the Ok button
        .cmbOk.Left = (.Width / 2) - (.cmbOk.Width / 2)
    
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
