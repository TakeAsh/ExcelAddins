VERSION 1.0 CLASS

BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_AddinInstall()
    Module1.RegisterJoinRangeText
    Module1.RegisterJoinRangeTextA
    Module1.RegisterJoinRangeValue
    Module1.RegisterJoinRangeValueA
End Sub
