VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Property Get TestColHidden2() As Range:      Set TestColHidden2 = Me.Range("colHiddenForTest2"):         End Property
Public Property Get MergedCellsSelect() As Range:   Set MergedCellsSelect = Me.Range("rngMergedCellsSelect"):   End Property
Public Property Get AutoFilter2() As Range:         Set AutoFilter2 = Me.Range("rngAutoFilter2"):               End Property
Public Property Get MergedCells1() As Range:        Set MergedCells1 = Me.Range("rngMergedCells1"):             End Property
Public Property Get MergedCells2() As Range:        Set MergedCells2 = Me.Range("rngMergedCells2"):             End Property
Public Property Get MergedCells3() As Range:        Set MergedCells3 = Me.Range("rngMergedCells3"):             End Property
Public Property Get UnMerged1() As Range:           Set UnMerged1 = Me.Range("rngUnMerged1"):                   End Property
Public Property Get UnMerged2() As Range:           Set UnMerged2 = Me.Range("rngUnMerged2"):                   End Property
Public Property Get UnMerged3() As Range:           Set UnMerged3 = Me.Range("rngUnMerged3"):                   End Property

