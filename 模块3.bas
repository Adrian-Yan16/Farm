Attribute VB_Name = "模块3"
Dim flag As Integer
' 季节
Public Sub UpdateTime()
    ' 定义一个包含四个季节的数组
    Dim seasons(3) As String
    ' 初始化数组元素
    seasons(0) = "春"
    seasons(1) = "夏"
    seasons(2) = "秋"
    seasons(3) = "冬"
    
    ' 声明一个变量用于作为索引
    
    flag = (flag + 1) Mod 4
    ThisWorkbook.Sheets("Sheet1").Range("G2").Value = "季节：" & seasons(flag) + " 时间：30分钟" ' 替换 "Sheet1" 为实际要操作的工作表名称
    Call StartTimer
End Sub

Public Sub StartTimer()
    Application.OnTime Now + TimeValue("00:01:00"), "UpdateTime"
End Sub

'收成
Sub CalProduct()
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Worksheets("Sheet2")
    '目前资金总数
    Set fund = ThisWorkbook.Worksheets("Sheet1").Range("B2")
    '所以农产品总价值
    Dim allFund  As Long
    allFund = fund
    For Each cell In sourceSheet.Range("A70:Z70")
        If cell.Value = "" Then
            Exit For
        Else
            '当前农产品单价
            Dim productPrice As Long
            productPrice = FindPrice(cell.Value)
            '当前农产品总产量
            Dim sumProduct As Long
            sumProduct = 0
            Set rangeCell = cell
            
            Do While True
                Set rangeCell = rangeCell.Offset(1, 0)
                If rangeCell.Value = "" Then
                    allFund = allFund + sumProduct
                    Exit Do
                Else
                    Dim rngTarget As Range
                    '将字符串地址转换为Range对象
                    Set rngTarget = sourceSheet.Range(rangeCell.Value)
                    
                    '获取农田数量
                    Dim cellCount As Long
                    cellCount = rngTarget.Rows.Count * rngTarget.Columns.Count
                    
                    sumProduct = sumProduct + cellCount
                End If
            Loop
            sumProduct = productPrice * sumProduct
        End If
    Next cell
    MsgBox "总赚：" & allFund
    'Call UnprotectCells
    'Worksheets("Sheet1").Range("B2").Value = allFund
    'Call ProtectCells
    

End Sub
