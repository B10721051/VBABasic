Attribute VB_Name = "Module1"
Sub �@�~1()
Attribute �@�~1.VB_Description = "�N�s��ƶq�q�j��p�Ƨ�\n�H�έp���`�M�γ̤j�Ȥγ̤p��"
Attribute �@�~1.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' �@�~1 ����
' �N�s��ƶq�q�j��p�Ƨ� �H�έp���`�M�γ̤j�Ȥγ̤p��
'
' �ֳt��: Ctrl+c
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("B1:B553")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=sum"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = ")"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[49]C[-3])"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = ""
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-5]:R[49]C[-5])"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[1]C[-7]:R[49]C[-7])"
    Range("I2").Select
End Sub
Sub �@�~2()
Attribute �@�~2.VB_Description = "�N�s��ƶq�q�p��j�Ƨ�\n�H�έp���`�M�γ̤j�Ȥγ̤p��"
Attribute �@�~2.VB_ProcData.VB_Invoke_Func = "v\n14"
'
' �@�~2 ����
' �N�s��ƶq�q�p��j�Ƨ� �H�έp���`�M�γ̤j�Ȥγ̤p��
'
' �ֳt��: Ctrl+v
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B50"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B50")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[1]C[-3]:R[49]C[-3])"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[1]C[-5]:R[49]C[-5])"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[1]C[-7]:R[49]C[-7])"
    Range("I2").Select
End Sub
