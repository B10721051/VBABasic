Attribute VB_Name = "Module1"
Sub sort0308()
Attribute sort0308.VB_Description = "口罩數量排序"
Attribute sort0308.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' sort0308 巨集
' 口罩數量排序
'
' 快速鍵: Ctrl+q
'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
