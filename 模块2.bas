Attribute VB_Name = "ģ��2"
'����ũ����۸�
Function FindPrice(compareString As String) As Long
    Dim sourceSheet As Worksheet
    'Dim priceCell As Range
    Dim lastRow As Long
    
    ' ������Դ������
    Set sourceSheet = ThisWorkbook.Worksheets("Sheet2")
    
    ' ��ȡA�����һ�е��к�
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' ����A1��E1�����еĵ�Ԫ����������ǵ�һ�У�
    For Each cell In sourceSheet.Range("R1:T1")
        ' �Ƚϵ�Ԫ�������Ƿ��봫�ݽ�����compareString���
        If cell.Value = compareString Then
            ' ���ƥ�䣬��ȡͬһ����һ�У���A2���ļ۸�
            Set priceCell = cell.Offset(1, 0)
            ' ���ؼ۸�
            FindPrice = priceCell.Value
            Exit Function ' ����ҵ��˾������������
        End If
    Next cell
    
    ' ���δ�ҵ���Ӧ����Ʒ���򷵻�һ������ֵ����Null��һ�����ֵ��
    FindPrice = Null
End Function

'���ҵ�����ֲ�Ĳ���
