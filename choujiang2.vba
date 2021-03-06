
Sub 删除重复项()
'
' 删除支持者中重复的用户pin
'
'
    ActiveSheet.Range("A:B").RemoveDuplicates Columns:=1, Header:=xlYes
    
        
End Sub


Sub 删除空白行()

'删除重复项后有空行，影响随机数产生，此宏用于删除空白行


For i = 2 To ActiveSheet.UsedRange.Rows.Count

If Cells(i, 1) = "" Then
  Rows(i).Delete
End If

Next i

End Sub

Sub 生成随机数并排序()

'
' 生成随机数，按照随机数排序
'

'
    Dim a As Integer
    
    a = ActiveSheet.UsedRange.Rows.Count
    
    Range(Cells(2, 4), Cells(a, 4)).Formula = "=Rand()"
    
    Range(Cells(2, 4), Cells(a, 4)).Select
    ActiveWorkbook.Worksheets("用户数据").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("用户数据").Sort.SortFields.Add Key:=Range("D2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("用户数据").Sort
        .SetRange Range(Cells(2, 1), Cells(a, 4))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 
  
End Sub


Sub 生成抽奖号()

Dim a As Integer
    
a = ActiveSheet.UsedRange.Rows.Count

Range("C2").Value = 0
Range("C3").Value = 1


Set SourceRange = Range("C2:C3")
Set fillRange = Range(Cells(2, 3), Cells(a, 3))
SourceRange.AutoFill Destination:=fillRange
    

End Sub



Sub 排重()

删除重复项
删除重复项
删除空白行
删除空白行
删除空白行
删除空白行
删除空白行
删除空白行

End Sub




Sub 奖项()

Dim x As Integer
Dim num

num = Range("L1").Value

Range("A4").Select

For x = 1 To num

Selection.Value = "第" & x & "位"

ActiveCell.Offset(1, 0).Select

Next

End Sub



Sub 中奖号()

Dim first
Dim i
Dim jg
Dim num


Range("B4").Value = "=MOD(H1,N1)"

first = Range("B4").Value
num = Range("L1").Value
jg = Range("N1").Value
xs = Range("P1").Value

Range("B5").Select

For i = 1 To (num - 1)

Selection.Value = Int(first + (xs * i))

ActiveCell.Offset(1, 0).Select

Next

End Sub


Sub 中奖用户()


' 根据中奖号挑选出用户pin

Dim num
num = Range("L1").Value

Range("C4").Select
   
Selection.Value = "=INDEX(用户数据!A:A,MATCH(B4,用户数据!C:C))"


If num > 1 Then

 
Set SourceRange = Range("C4")
Set fillRange = Range(Cells(4, 3), Cells(num + 3, 3))
SourceRange.AutoFill Destination:=fillRange
 
End If
 
End Sub


Sub 中奖订单()


' 根据中奖号挑选出用户pin

Dim num
num = Range("L1").Value

Range("D4").Select
   
Selection.Value = "=INDEX(用户数据!B:B,MATCH(B4,用户数据!C:C))"


If num > 1 Then

 
Set SourceRange = Range("D4")
Set fillRange = Range(Cells(4, 4), Cells(num + 3, 4))
SourceRange.AutoFill Destination:=fillRange
 
End If
 
End Sub


Sub 中奖用户pin隐藏()


' 将用户pin从第3位开始的3个字符用*代替

Dim num
num = Range("L1").Value

Range("E4").Select
   
Selection.Value = "=REPLACE(C4,3,3,""***"")"

If num > 1 Then
Set SourceRange = Range("E4")
Set fillRange = Range(Cells(4, 5), Cells(num + 3, 5))
SourceRange.AutoFill Destination:=fillRange
 
End If
 
End Sub





Sub 开始抽奖()


奖项
中奖号
中奖用户
中奖订单
中奖用户pin隐藏

End Sub



Sub 清空用户数据()
'
' 清空所有用户数据


Dim a

a = ActiveSheet.UsedRange.Rows.Count

Range(Cells(2, 1), Cells(a + 1, 4)).Select

Selection.ClearContents

Range("A2").Select
    
End Sub




Sub 清空抽奖数据()

'
' 清空所有抽奖数据


Dim b

b = Range("J1").Value

Range(Cells(4, 1), Cells(b + 7, 5)).Select

Selection.ClearContents

Range("A4").Select
    
End Sub


Sub 订单号展示()



    Sheets("订单号展示").Select
    Cells.Select
    Selection.EntireRow.Delete


    Sheets("用户数据").Select
    Columns("B:C").Select
    Selection.Copy
    
     
    Sheets("订单号展示").Select
    Range("A1").Select
    ActiveSheet.Paste
 

Dim a, b

a = ActiveSheet.UsedRange.Rows.Count

b = Int((a - 1) / 3) + 1


Range("C1").Value = "订单号"
Range("D1").Value = "抽奖号"
Range("E1").Value = "订单号"
Range("F1").Value = "抽奖号"



Range(Cells(b + 1, 1), Cells(b * 2 - 1, 2)).Select
Selection.Cut
Range("C2").Select
ActiveSheet.Paste


Range(Cells(b * 2, 1), Cells(a, 2)).Select
Selection.Cut
Range("E2").Select
ActiveSheet.Paste



    Range("A:A,C:C,E:E").Select
    Selection.ColumnWidth = 15
    Range("B:B,D:D,F:F").Select
    Selection.ColumnWidth = 8





End Sub




Sub 订单号数字格式()
'
' 订单号从文本格式改为数字格式（去掉 '）
'

'
    Columns("B:B").Select
    Selection.Replace What:="'", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.NumberFormatLocal = "0_);[红色](0)"
   
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub





Sub 中奖号展示()


    Dim r As Integer
    r = Sheets("抽奖").Range("L1").Value

    Sheets("中奖号展示").Select
    Cells.Select
    Selection.EntireRow.Delete
    Range("A1").Value = "中奖号"
    Range("B1").Value = "订单号"
    Range("C1").Value = "京东账号"


    Sheets("抽奖").Select
    Range(Cells(4, 2), Cells(4 + r - 1, 2)).Select
    Selection.Copy
    Sheets("中奖号展示").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
     
    Sheets("抽奖").Select
    Range(Cells(4, 4), Cells(4 + r - 1, 4)).Select
    Selection.Copy
    Sheets("中奖号展示").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
     
    Sheets("抽奖").Select
    Range(Cells(4, 5), Cells(4 + r - 1, 5)).Select
    Selection.Copy
    Sheets("中奖号展示").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
  
  
    Sheets("中奖号展示").Select
    Range("A:A").Select
    Selection.ColumnWidth = 10
    Range("B:B").Select
    Selection.ColumnWidth = 20
    Range("C:C").Select
    Selection.ColumnWidth = 20


End Sub
