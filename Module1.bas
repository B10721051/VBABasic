Attribute VB_Name = "Module1"
Sub sort0308()
Attribute sort0308.VB_Description = "�f�n�ƶq�Ƨ�"
Attribute sort0308.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' sort0308 ����
' �f�n�ƶq�Ƨ�
'
' �ֳt��: Ctrl+q
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
