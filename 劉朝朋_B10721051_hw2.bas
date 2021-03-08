Attribute VB_Name = "Module1"
Sub 作業1()
Attribute 作業1.VB_Description = "將酒精數量從大到小排序\n以及計算總和及最大值及最小值"
Attribute 作業1.VB_ProcData.VB_Invoke_Func = "c\n14"
'
' 作業1 巨集
' 將酒精數量從大到小排序 以及計算總和及最大值及最小值
'
' 快速鍵: Ctrl+c
'
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B553"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
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
Sub 作業2()
Attribute 作業2.VB_Description = "將酒精數量從小到大排序\n以及計算總和及最大值及最小值"
Attribute 作業2.VB_ProcData.VB_Invoke_Func = "v\n14"
'
' 作業2 巨集
' 將酒精數量從小到大排序 以及計算總和及最大值及最小值
'
' 快速鍵: Ctrl+v
'
    Range("B1").Select
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B50"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
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
