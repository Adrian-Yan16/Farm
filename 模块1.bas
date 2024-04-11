Attribute VB_Name = "模块1"
'使用化肥和浇水来促进生长
Public Sub Button_Click()
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("确定要使用吗？", vbYesNo + vbQuestion, "确认操作")
    
    If userResponse = vbYes Then
        Dim ifSelect As VbMsgBoxResult
        ifSelect = MsgBox("请绘制使用区域", vbYesNo + vbQuestion, "确认操作")
        If ifSelect = vbYes Then
            Dim userSelectedRange As Range
            Set userSelectedRange = WaitUserSelectCell()
            
            If Not userSelectedRange Is Nothing Then
                ' 用户已选择区域，你可以在此处使用selectedRange
                'MsgBox "您选择了: " & userSelectedRange.Address
                Dim surplusFund As Long
                surplusFund = IfFundEnough(userSelectedRange, 1)
                
                If surplusFund < 0 Then
                    MsgBox "资金不足!"
                Else
                    Call UnprotectCells
                    Worksheets("Sheet1").Range("B2").Value = surplusFund
                    Call ProtectCells
                End If
                'todo，区域内的产量加倍
            Else
                MsgBox "用户未选择任何区域。"
            End If
        End If
    End If
End Sub

'用户选择区域
Function WaitUserSelectCell() As Range
    Dim selectedRange As Range
    ' 弹出对话框让用户选择单元格区域
    On Error Resume Next ' 处理用户取消选择的情况
    Set selectedRange = Application.InputBox("请选择单元格区域:", Type:=8)
    On Error GoTo 0 ' 取消错误处理
    
    ' 如果用户成功选择了一个区域，则返回该区域
    If Not selectedRange Is Nothing Then
        Set WaitUserSelectCell = selectedRange
    End If
End Function

'遍历单元格填充农作物
Sub FillCellUseProduce(ByRef inputRange As Range, ByRef produce As Variant)
    inputRange.Value = produce
End Sub

'计算资金是否够支付种子或者化肥或浇水
Function IfFundEnough(ByRef inputRange As Range, ByRef price As Long) As Long
    '获取资金
    Dim fund As Long
    fund = Worksheets("Sheet1").Range("B2").Value
    
    '获取农田数量
    Dim cellCount As Long
    cellCount = inputRange.Rows.Count * inputRange.Columns.Count
    
    Dim sum As Long
    sum = price * cellCount
    
    Dim surplusFund As Long
    surplusFund = fund - sum
    ' 剩余资金
    IfFundEnough = surplusFund
    
End Function

'让资金栏无法被更改
Sub ProtectCells()
    ' 解锁允许编辑的单元格
    Worksheets("Sheet1").Range("B2").Locked = True
    'Worksheets("Sheet2").Range("A41:Z60").Locked = True
    
    ' 锁定不允许编辑的单元格，此处假定其余单元格默认锁定
    ' 若需要锁定某些单元格，可以用类似上述的方法设置 Locked = True
    
    ' 保护工作表，设置密码（如果有）
    Worksheets("Sheet1").Protect Password:="Kk124589.", UserInterfaceOnly:=True
End Sub

Sub UnprotectCells()
    Worksheets("Sheet1").Unprotect Password:="Kk124589."
End Sub

'todo 保存农作物：区域
Sub SaveProduceRange(ByRef produceRange As Range, ByRef produce As Variant)
    '保存在sheet2中
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Worksheets("Sheet2")
    For Each cell In sourceSheet.Range("A70:Z70")
        If cell.Value = "" Then
            cell.Value = produce
            Set cellRange = cell.Offset(1, 0)
            cellRange.Value = produceRange.Address
            Exit Sub
        ElseIf cell.Value = produce Then
            While True
                Set cell = cell.Offset(1, 0)
                If cell.Value = "" Then
                    cell.Value = produceRange.Address
                    Exit Sub
                End If
            Wend
        End If
    Next cell
End Sub
