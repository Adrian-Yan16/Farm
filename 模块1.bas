Attribute VB_Name = "ģ��1"
'ʹ�û��ʺͽ�ˮ���ٽ�����
Public Sub Button_Click()
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("ȷ��Ҫʹ����", vbYesNo + vbQuestion, "ȷ�ϲ���")
    
    If userResponse = vbYes Then
        Dim ifSelect As VbMsgBoxResult
        ifSelect = MsgBox("�����ʹ������", vbYesNo + vbQuestion, "ȷ�ϲ���")
        If ifSelect = vbYes Then
            Dim userSelectedRange As Range
            Set userSelectedRange = WaitUserSelectCell()
            
            If Not userSelectedRange Is Nothing Then
                ' �û���ѡ������������ڴ˴�ʹ��selectedRange
                'MsgBox "��ѡ����: " & userSelectedRange.Address
                Dim surplusFund As Long
                surplusFund = IfFundEnough(userSelectedRange, 1)
                
                If surplusFund < 0 Then
                    MsgBox "�ʽ���!"
                Else
                    Call UnprotectCells
                    Worksheets("Sheet1").Range("B2").Value = surplusFund
                    Call ProtectCells
                End If
                'todo�������ڵĲ����ӱ�
            Else
                MsgBox "�û�δѡ���κ�����"
            End If
        End If
    End If
End Sub

'�û�ѡ������
Function WaitUserSelectCell() As Range
    Dim selectedRange As Range
    ' �����Ի������û�ѡ��Ԫ������
    On Error Resume Next ' �����û�ȡ��ѡ������
    Set selectedRange = Application.InputBox("��ѡ��Ԫ������:", Type:=8)
    On Error GoTo 0 ' ȡ��������
    
    ' ����û��ɹ�ѡ����һ�������򷵻ظ�����
    If Not selectedRange Is Nothing Then
        Set WaitUserSelectCell = selectedRange
    End If
End Function

'������Ԫ�����ũ����
Sub FillCellUseProduce(ByRef inputRange As Range, ByRef produce As Variant)
    inputRange.Value = produce
End Sub

'�����ʽ��Ƿ�֧�����ӻ��߻��ʻ�ˮ
Function IfFundEnough(ByRef inputRange As Range, ByRef price As Long) As Long
    '��ȡ�ʽ�
    Dim fund As Long
    fund = Worksheets("Sheet1").Range("B2").Value
    
    '��ȡũ������
    Dim cellCount As Long
    cellCount = inputRange.Rows.Count * inputRange.Columns.Count
    
    Dim sum As Long
    sum = price * cellCount
    
    Dim surplusFund As Long
    surplusFund = fund - sum
    ' ʣ���ʽ�
    IfFundEnough = surplusFund
    
End Function

'���ʽ����޷�������
Sub ProtectCells()
    ' ��������༭�ĵ�Ԫ��
    Worksheets("Sheet1").Range("B2").Locked = True
    'Worksheets("Sheet2").Range("A41:Z60").Locked = True
    
    ' ����������༭�ĵ�Ԫ�񣬴˴��ٶ����൥Ԫ��Ĭ������
    ' ����Ҫ����ĳЩ��Ԫ�񣬿��������������ķ������� Locked = True
    
    ' �����������������루����У�
    Worksheets("Sheet1").Protect Password:="Kk124589.", UserInterfaceOnly:=True
End Sub

Sub UnprotectCells()
    Worksheets("Sheet1").Unprotect Password:="Kk124589."
End Sub

'todo ����ũ�������
Sub SaveProduceRange(ByRef produceRange As Range, ByRef produce As Variant)
    '������sheet2��
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
