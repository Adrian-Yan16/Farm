VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' 检查是否选择了设置了数据验证的单元格
    If Not Intersect(Target, Me.Range("D2")) Is Nothing Then ' 假设数据验证设置在A1单元格
        MsgBox "您选择了: " & Target.Value

    End If
End Sub
