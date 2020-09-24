VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateChangedCommComps ThisWorkbook
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Const HOSTED_COMMON_COMPONENTS = "mObstructions"
#If CompMan Then
    mCompMan.ExportChangedComponents ThisWorkbook, sHosted:=HOSTED_COMMON_COMPONENTS
    mCommComp.ManageIrregularUpdates ThisWorkbook, sHosted:=HOSTED_COMMON_COMPONENTS
#End If
End Sub
