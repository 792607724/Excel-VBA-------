Attribute VB_Name = "模块1"
Sub 标记销售金额小于10000元的数据()
Attribute 标记销售金额小于10000元的数据.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' 标记销售金额小于10000元的数据 宏
'
' 快捷键: Ctrl+Shift+P
'
    Columns("F:F").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=10000"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
