Attribute VB_Name = "ģ��3"
Dim flag As Integer
' ����
Public Sub UpdateTime()
    ' ����һ�������ĸ����ڵ�����
    Dim seasons(3) As String
    ' ��ʼ������Ԫ��
    seasons(0) = "��"
    seasons(1) = "��"
    seasons(2) = "��"
    seasons(3) = "��"
    
    ' ����һ������������Ϊ����
    
    flag = (flag + 1) Mod 4
    ThisWorkbook.Sheets("Sheet1").Range("G2").Value = "���ڣ�" & seasons(flag) + " ʱ�䣺30����" ' �滻 "Sheet1" Ϊʵ��Ҫ�����Ĺ���������
    Call StartTimer
End Sub

Public Sub StartTimer()
    Application.OnTime Now + TimeValue("00:01:00"), "UpdateTime"
End Sub

'�ճ�
Sub CalProduct()
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Worksheets("Sheet2")
    'Ŀǰ�ʽ�����
    Set fund = ThisWorkbook.Worksheets("Sheet1").Range("B2")
    '����ũ��Ʒ�ܼ�ֵ
    Dim allFund  As Long
    allFund = fund
    For Each cell In sourceSheet.Range("A70:Z70")
        If cell.Value = "" Then
            Exit For
        Else
            '��ǰũ��Ʒ����
            Dim productPrice As Long
            productPrice = FindPrice(cell.Value)
            '��ǰũ��Ʒ�ܲ���
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
                    '���ַ�����ַת��ΪRange����
                    Set rngTarget = sourceSheet.Range(rangeCell.Value)
                    
                    '��ȡũ������
                    Dim cellCount As Long
                    cellCount = rngTarget.Rows.Count * rngTarget.Columns.Count
                    
                    sumProduct = sumProduct + cellCount
                End If
            Loop
            sumProduct = productPrice * sumProduct
        End If
    Next cell
    MsgBox "��׬��" & allFund
    'Call UnprotectCells
    'Worksheets("Sheet1").Range("B2").Value = allFund
    'Call ProtectCells
    

End Sub
