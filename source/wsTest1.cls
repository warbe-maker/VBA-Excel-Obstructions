VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsTest1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Friend Property Get NamesName() As Range:       Set NamesName = Me.Range("colNamesName"):           End Property
Friend Property Get NamesRangeType() As Range:  Set NamesRangeType = Me.Range("colNamesRangeType"): End Property
Friend Property Get NamesReference() As Range:  Set NamesReference = Me.Range("colNamesReference"): End Property
Friend Property Get NamesSheet() As Range:      Set NamesSheet = Me.Range("colNamesSheet"):         End Property
Friend Property Get NamesScope() As Range:      Set NamesScope = Me.Range("colNamesScope"):         End Property
Public Property Get RngNames() As Range:        Set RngNames = Me.Range("rngNames"):                End Property
Public Property Get TestColHidden1() As Range:  Set TestColHidden1 = Me.Range("colHiddenForTest1"): End Property
Public Property Get AutoFilter1() As Range:     Set AutoFilter1 = Me.Range("rngAutoFilter1"):       End Property

