Attribute VB_Name = "模块2"
'查找农作物价格
Function FindPrice(compareString As String) As Long
    Dim sourceSheet As Worksheet
    'Dim priceCell As Range
    Dim lastRow As Long
    
    ' 设置来源工作表
    Set sourceSheet = ThisWorkbook.Worksheets("Sheet2")
    
    ' 获取A列最后一行的行号
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' 遍历A1到E1所在行的单元格（这里假设是第一行）
    For Each cell In sourceSheet.Range("R1:T1")
        ' 比较单元格内容是否与传递进来的compareString相等
        If cell.Value = compareString Then
            ' 如果匹配，获取同一列下一行（即A2）的价格
            Set priceCell = cell.Offset(1, 0)
            ' 返回价格
            FindPrice = priceCell.Value
            Exit Function ' 如果找到了就无需继续遍历
        End If
    Next cell
    
    ' 如果未找到相应的商品，则返回一个特殊值（如Null或一个标记值）
    FindPrice = Null
End Function

'查找地区种植的产物
