Sub 按照某个规则分表()

    Dim a
    Dim i As Integer
    '循环列表
    Sheets("工作表1").Select
    For i = 1 To 12
    a = Cells(i, 1).Value
    '根据规则进行数据筛选复制
    Sheets("工作表2").Select
    ActiveSheet.Range("$A$1:$J$1389").AutoFilter Field:=3, Criteria1:=a
    Range("A1:J1389").Select
    Selection.Copy
    '将复制内容保存到新的工作簿并命名
    Workbooks.Add
    ActiveSheet.Paste
    ChDir "这里是文件夹路径"
    ActiveWorkbook.SaveAs Filename:=a & "_文件名后缀.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    Workbooks("工作簿.xlsm").Activate
    Sheets("工作表1").Select
    Next

End Sub



Sub 选项归位()

'将竖着排列的ABCDE选项变成横向排列

Do While ActiveCell.Value > 0

If Left(ActiveCell.Value, 1) = "A" Or Left(ActiveCell.Value, 1) = "a" Then
Selection.Cut
ActiveCell.Offset(-1, 2).Select
ActiveSheet.Paste
ActiveCell.Offset(2, -2).Select


ElseIf Left(ActiveCell.Value, 1) = "B" Or Left(ActiveCell.Value, 1) = "b" Then
Selection.Cut
ActiveCell.Offset(-2, 3).Select
ActiveSheet.Paste
ActiveCell.Offset(3, -3).Select

ElseIf Left(ActiveCell.Value, 1) = "C" Or Left(ActiveCell.Value, 1) = "c" Then
Selection.Cut
ActiveCell.Offset(-3, 4).Select
ActiveSheet.Paste
ActiveCell.Offset(4, -4).Select

ElseIf Left(ActiveCell.Value, 1) = "D" Or Left(ActiveCell.Value, 1) = "d" Then
Selection.Cut
ActiveCell.Offset(-4, 5).Select
ActiveSheet.Paste
ActiveCell.Offset(5, -5).Select

ElseIf Left(ActiveCell.Value, 1) = "E" Or Left(ActiveCell.Value, 1) = "e" Then
Selection.Cut
ActiveCell.Offset(-5, 6).Select
ActiveSheet.Paste
ActiveCell.Offset(6, -6).Select

ElseIf Left(ActiveCell.Value, 1) <> "A" And Left(ActiveCell.Value, 1) <> "B" And Left(ActiveCell.Value, 1) <> "C" And Left(ActiveCell.Value, 1) <> "D" And Left(ActiveCell.Value, 1) <> "E" And Left(ActiveCell.Value, 1) <> "a" And Left(ActiveCell.Value, 1) <> "b" And Left(ActiveCell.Value, 1) <> "c" And Left(ActiveCell.Value, 1) <> "d" And Left(ActiveCell.Value, 1) <> "e" Then
ActiveCell.Offset(1, 0).Select


End If

Loop

End Sub



Sub 删空行()

Dim i As Integer

For i = 1 To 60
 Cells(i, 1).Select
 If ActiveCell.Value = 0 Then
   Selection.EntireRow.Delete
Else
 ActiveCell.Offset(1, 0).Select
End If

Next
End Sub



Sub 多次删空行()

删空行
Cells(1, 1).Select
删空行
Cells(1, 1).Select
删空行
Cells(1, 1).Select
删空行
Cells(1, 1).Select

End Sub

