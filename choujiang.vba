

Sub 删除重复项()
'
' 删除支持者中重复的用户pin
'
'
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
    
        
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
    
    Range(Cells(2, 3), Cells(a, 3)).Formula = "=Rand()"
    
    Range(Cells(2, 3), Cells(a, 3)).Select
    ActiveWorkbook.Worksheets("用户数据").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("用户数据").Sort.SortFields.Add Key:=Range("C2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("用户数据").Sort
        .SetRange Range(Cells(2, 1), Cells(a, 3))
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

Range("B2").Value = 1
Range("B3").Value = 2


Set SourceRange = Range("B2:B3")
Set fillRange = Range(Cells(2, 2), Cells(a, 2))
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

num = Range("J1").Value

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


Range("B4").Value = "=(B1 - F1) * D1 - Quotient((B1 - F1) * D1, L1) * L1"

first = Range("B4").Value
num = Range("J1").Value
jg = Range("L1").Value

Range("B5").Select

For i = 1 To (num - 1)

Selection.Value = first + (jg * i)

ActiveCell.Offset(1, 0).Select

Next

End Sub


Sub 中奖用户()


' 根据中奖号挑选出用户pin

Dim num
num = Range("J1").Value

Range("C4").Select
   
Selection.Value = "=INDEX(用户数据!A:A,MATCH(B4,用户数据!B:B))"
  
 
Set SourceRange = Range("C4")
Set fillRange = Range(Cells(4, 3), Cells(num + 3, 3))
SourceRange.AutoFill Destination:=fillRange
 
 
End Sub



Sub 开始抽奖()

奖项
中奖号
中奖用户

End Sub
