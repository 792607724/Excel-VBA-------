Attribute VB_Name = "ģ��1"
Sub �����������С��20Ԫ������()
Attribute �����������С��20Ԫ������.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' �����������С��20Ԫ������ ��
'
' ��ݼ�: Ctrl+Shift+P
'
    Columns("D:D").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=20"
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
