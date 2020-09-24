VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12690
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const H_MARGIN                  As Single = 15
Const V_MARGIN                  As Single = 10
Const MIN_FORM_WIDTH            As Single = 280
Const MIN_REPLY_WIDTH           As Single = 70
Dim lFixedFontMessageLines      As Long
Dim sTitle                      As String
Dim sErrSrc                     As String
'Dim sFixedFontMessage           As String
Dim vReplies                    As Variant
Dim aReplies                    As Variant
Dim siFormWidth                 As Single
Dim sTitleFontName              As String
Dim sTitleFontSize              As String ' Ignored when sTitleFontName is not provided
Dim sProportionalFontMessage    As String
Dim siTopNextElement            As Single
Dim bWithLabel                  As Boolean
Dim sMsg1Proportional           As String
Dim sMsg2Proportional           As String
Dim sMsg3Proportional           As String
Dim sMsg1Fixed                  As String
Dim sMsg2Fixed                  As String
Dim sMsg3Fixed                  As String
Dim sLabelMessage1              As String
Dim sLabelMessage2              As String
Dim sLabelMessage3              As String

Private Sub UserForm_Initialize()
    siFormWidth = MIN_FORM_WIDTH ' Default
End Sub

Public Property Let ErrSrc(ByVal s As String):                  sErrSrc = s:                                    End Property
Public Property Let FormWidth(ByVal si As Single):              siFormWidth = si:                               End Property
Public Property Let LabelMessage1(ByVal s As String):           sLabelMessage1 = s:                             End Property
Public Property Let LabelMessage2(ByVal s As String):           sLabelMessage2 = s:                             End Property
Public Property Let LabelMessage3(ByVal s As String):           sLabelMessage3 = s:                             End Property
Private Property Get LabelMsg1() As MSForms.Label:              Set LabelMsg1 = Me.laMsg1:                      End Property
Private Property Get LabelMsg2() As MSForms.Label:              Set LabelMsg2 = Me.laMsg2:                      End Property
Private Property Get LabelMsg3() As MSForms.Label:              Set LabelMsg3 = Me.laMsg3:                      End Property
Public Property Let Message1Fixed(ByVal s As String):           sMsg1Fixed = s:                                 End Property
Public Property Let Message1Proportional(ByVal s As String):    sMsg1Proportional = s:                          End Property
Public Property Let Message2Fixed(ByVal s As String):           sMsg2Fixed = s:                                 End Property
Public Property Let Message2Proportional(ByVal s As String):    sMsg2Proportional = s:                          End Property
Public Property Let Message3Fixed(ByVal s As String):           sMsg3Fixed = s:                                 End Property
Public Property Let Message3Proportional(ByVal s As String):    sMsg3Proportional = s:                          End Property
Private Property Get Msg1Fixed() As MSForms.TextBox:            Set Msg1Fixed = Me.tbMsg1Fixed:                 End Property
Private Property Get Msg1Proportional() As MSForms.TextBox:     Set Msg1Proportional = Me.tbMsg1Proportional:   End Property
Private Property Get Msg2Fixed() As MSForms.TextBox:            Set Msg2Fixed = Me.tbMsg2Fixed:                 End Property
Private Property Get Msg2Proportional() As MSForms.TextBox:     Set Msg2Proportional = Me.tbMsg2Proportional:   End Property
Private Property Get Msg3Fixed() As MSForms.TextBox:            Set Msg3Fixed = Me.tbMsg3Fixed:                 End Property
Private Property Get Msg3Proportional() As MSForms.TextBox:     Set Msg3Proportional = Me.tbMsg3Proportional:   End Property

Public Property Let Replies(ByVal v As Variant)
    vReplies = v
    aReplies = Split(v, ",")
End Property

Public Property Let Title(ByVal s As String):                    sTitle = s:                                    End Property
Public Property Let TitleFontName(ByVal s As String):            sTitleFontName = s:                            End Property
Public Property Let TitleFontSize(ByVal l As Long):              sTitleFontSize = l:                            End Property

Private Sub AdjustTextBoxAndFormWidth( _
            ByVal tb As MSForms.TextBox, _
            ByVal sText As String, _
            ByVal bFixed As Boolean)
' ----------------------------------------
'
' ----------------------------------------
Dim sSplit      As String
Dim v           As Variant
Dim siMaxWidth  As Single

    If bFixed Then
        '~~ A fixed font Textbox's width is determined by
        '~~ - the maximum text line length (determined by means of an autosized width-template)
        '~~ - the Title width (minimum to avoid truncation)
        If InStr(sText, vbLf) <> 0 Then sSplit = vbLf
        If InStr(sText, vbCrLf) <> 0 Then sSplit = vbCrLf
        For Each v In Split(sText, sSplit)
            Me.tbMsgFixedWidthTemplate.Value = v
            siMaxWidth = Max(siMaxWidth, Me.Width, Me.tbMsgFixedWidthTemplate.Width, Me.laTitle.Width)
        Next v
        tb.Width = siMaxWidth + 10
        Me.Width = tb.Left + tb.Width + H_MARGIN
    Else
        '~~ The width of a proportional font Textbox is determined by
        '~~ - The width of the Title
        '~~ - The specified minimum Form width
        '~~ - The provided/desired minimum Form width of the caller
        '~~ Adjust Form width
        With Me
            .Width = mCommon.Max(MIN_FORM_WIDTH, _
                                 siFormWidth, _
                                 .laTitle.Width)
            
            tb.Width = .Width - tb.Left - H_MARGIN
            .laTitle.Width = .Width
            .laTitleSpaceBottom.Width = .Width
        End With
    End If

End Sub

Private Sub AdjustTextBoxHeight(ByVal tb As MSForms.TextBox, _
                                ByVal sText As String)
' ------------------------------------------------------------
' Adjust the height of the Textbox (tb) according to the
' number of lines of the text (sText). Consider the maximum
' height of the UserForm.
' ------------------------------------------------------------
Dim sSplit      As String
Dim lLinesText  As Long
Dim lLinesInBox As Long

    If InStr(sText, vbLf) <> 0 Then sSplit = vbLf
    If InStr(sText, vbCrLf) <> 0 Then sSplit = vbCrLf
    lLinesText = UBound(Split(sText, sSplit)) + 1
    With tb
        .SetFocus
        .Value = .Value
        Debug.Print "Number of Textbox lines = " & .LineCount
        Debug.Print "Number of test lines    = " & lLinesText
        lLinesInBox = Max(lLinesText, .LineCount)
        Select Case lLinesInBox
            Case 0
            Case 1:     .Height = 15.2
            Case 2, 3:  .Height = lLinesInBox * 12.4
            Case Else:  .Height = lLinesInBox * 12
        End Select
    End With
    
 End Sub

Private Sub cmbReply1_Click()
    With Me.cmbReply1
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub cmbReply2_Click()
    With Me.cmbReply2
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub cmbReply3_Click()
    With Me.cmbReply3
        Select Case UCase(.Caption)
            Case "OK":      ReplyWith vbOK
            Case "YES":     ReplyWith vbYes
            Case "NO":      ReplyWith vbNo
            Case "CANCEL":  ReplyWith vbCancel
            Case Else:      ReplyWith .Caption
        End Select
    End With
End Sub

Private Sub ReplyWith(ByVal v As Variant)
    mCommon.MsgReply = v
    Unload Me
End Sub

Private Sub SetupMessageFixed( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' -----------------------------------------------------
' Any fixed font message's width is adjusted to the
' maximum line width.
' -----------------------------------------------------
Dim v           As Variant
Dim siMaxWidth  As Single

    If sTextBoxText <> vbNullString Then
        With la
            '~~ Error Path: Adjust top position and height
            If sLabelText <> vbNullString Then
                With la
                    .Caption = sLabelText
                    .Visible = True
                    .Top = siTopNextElement
                    siTopNextElement = .Top + .Height
                End With
            End If
        End With
        
        With tb
            .Value = sTextBoxText
            .Visible = True
            .Top = siTopNextElement
            
            AdjustTextBoxHeight tb, sTextBoxText
            AdjustTextBoxAndFormWidth tb, sTextBoxText, bFixed:=True
            
            siTopNextElement = .Top + .Height + V_MARGIN
        End With
        
        With Me
            .Width = mCommon.Max(MIN_FORM_WIDTH, _
                                 siFormWidth, _
                                 .laTitle.Width, _
                                 tb.Left + tb.Width)
            
            .laTitle.Width = .Width
            .laTitleSpaceBottom.Width = .Width
        End With
    End If
End Sub

Private Sub SetupMessageProportional( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' ---------------------------------------
'
' ---------------------------------------
    If sTextBoxText <> vbNullString Then
        
        '~~ Setup Message Label
        If sLabelText <> vbNullString Then
            With la
                .Caption = sLabelText
                .Visible = True
                .Top = siTopNextElement
                siTopNextElement = .Top + .Height
            End With
        End If
        
        '~~ Setup Message Textbox
        With tb
            .Top = siTopNextElement
            .Visible = True
'            .AutoSize = True
            .WordWrap = True
            .ScrollBars = fmScrollBarsVertical
            .Value = sTextBoxText
'            .Width = Me.Width - .Left - H_MARGIN ' Adjust textbox width to Form width
'            .AutoSize = True
            
            AdjustTextBoxAndFormWidth tb, sTextBoxText, bFixed:=False
            AdjustTextBoxHeight tb, sTextBoxText
            
            siTopNextElement = .Top + .Height + V_MARGIN
        End With
        
    End If

End Sub

Private Sub SetupReply(ByVal cmb As MSForms.CommandButton, _
                       ByVal s As String)
' --------------------------------------------------
' Setup Command Button width and height.
' --------------------------------------------------
    With cmb
        .Top = siTopNextElement + V_MARGIN
        .Visible = True
        .Caption = s
        .Width = Max(MIN_REPLY_WIDTH, .Width)
    End With
End Sub

Private Sub SetupReplyButtons(ByVal vReplies As Variant)
' -------------------------------------------------
' Setup and position the reply buttons
' -------------------------------------------------
Dim lReplies    As Long

    With Me
        '~~ Setup button caption
        Select Case vReplies
            Case vbOKOnly, "OK"
                lReplies = 1
                SetupReply .cmbReply1, "Ok"
            Case vbYesNo
                lReplies = 2
                SetupReply .cmbReply1, "Yes"
                SetupReply .cmbReply2, "No"
            Case vbOKCancel
                lReplies = 2
                SetupReply .cmbReply1, "OK"
                SetupReply .cmbReply2, "Cancel"
            Case vbYesNoCancel
                lReplies = 3
                SetupReply .cmbReply1, "Yes"
                SetupReply .cmbReply2, "No"
                SetupReply .cmbReply3, "Cancel"
            Case Else
                lReplies = UBound(aReplies) + 1
                Select Case lReplies
                    Case 1
                        SetupReply .cmbReply1, aReplies(0)
                    Case 2
                        SetupReply .cmbReply1, aReplies(0)
                        SetupReply .cmbReply2, aReplies(1)
                        .cmbReply2.Visible = True
                    Case 3
                        SetupReply .cmbReply1, aReplies(0)
                        SetupReply .cmbReply2, aReplies(1)
                        SetupReply .cmbReply3, aReplies(2)
                End Select
        End Select
        
        '~~ Setup reply button position and size
        Select Case lReplies
            Case 1
                With .cmbReply1
                    .Left = (Me.Width / 2) - (.Width / 2) ' Center the only reply button
                End With
            Case 2
                With .cmbReply1
                    .Left = (Me.Width / 2) - .Width - V_MARGIN ' left from center
                End With
                With .cmbReply2
                    .Left = (Me.Width / 2) + V_MARGIN ' Right from center
                End With
            Case 3
                With .cmbReply2
                    .Left = (Me.Width / 2) - (.Width / 2) ' Center
                End With
                With .cmbReply1
                    .Left = Me.cmbReply2.Left - .Width - V_MARGIN ' Left from center
                End With
                With .cmbReply3
                    .Left = Me.cmbReply2.Left + Me.cmbReply2.Width + V_MARGIN ' Right from center
                End With
        End Select
        
        .Height = .cmbReply1.Top + .cmbReply1.Height + (V_MARGIN * 5)
    End With

End Sub

Private Sub SetupTitle()
' ----------------------------------------------------------------
' When a font name other than the system's font name is provided
' an extra title label mimics the title bar.
' In any case the title label is used to determine the form width
' by autosize of the label.
' ----------------------------------------------------------------
    
    With Me
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            '~~ A title with a specific font is displayed in a dedicated title label
            With .laTitle   ' Hidden by default
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .Visible = True
            End With
            siTopNextElement = .laTitleSpaceBottom.Top + .laTitleSpaceBottom.Height + V_MARGIN
            
        Else
            .Caption = sTitle
            .laTitleSpaceBottom.Visible = False
            With .laTitle
                '~~ The title label is used to adjust the form width
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.7
                End With
                .Visible = False
            End With
            siTopNextElement = V_MARGIN
        End If
        
        With .laTitle
            '~~ The title label is used to adjust the form width
            .AutoSize = True
            .Caption = "  " & sTitle    ' some left margin
            .AutoSize = False
            .Width = Max(MIN_FORM_WIDTH, .Width + H_MARGIN)
        End With
        .laTitleSpaceBottom.Width = .laTitle.Width
    End With

End Sub

Private Sub UserForm_Activate()
    
    With Me
        SetupTitle
        
        If sMsg1Proportional <> vbNullString _
        Then SetupMessageProportional LabelMsg1, sLabelMessage1, Msg1Proportional, sMsg1Proportional
        If sMsg1Fixed <> vbNullString _
        Then SetupMessageFixed LabelMsg1, sLabelMessage1, Msg1Fixed, sMsg1Fixed
        
        If sMsg2Proportional <> vbNullString _
        Then SetupMessageProportional LabelMsg2, sLabelMessage2, Msg2Proportional, sMsg2Proportional
        If sMsg2Fixed <> vbNullString _
        Then SetupMessageFixed LabelMsg2, sLabelMessage2, Msg2Fixed, sMsg2Fixed
        
        If sMsg3Proportional <> vbNullString _
        Then SetupMessageProportional LabelMsg3, sLabelMessage3, Msg3Proportional, sMsg3Proportional
        If sMsg3Fixed <> vbNullString _
        Then SetupMessageFixed LabelMsg3, sLabelMessage3, Msg3Fixed, sMsg3Fixed
        
        SetupReplyButtons vReplies
    End With

'    MakeFormResizable

End Sub
