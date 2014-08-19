Sub 大版积分排序()
'
' 大版积分从大到小排序
'考核前2名，当月获得100分版主积分
'考核3～5名，当月获得80分版主积分
'考核6~8名，当月获得60分版主积分
'


     ActiveCell.Offset(0, -1).Select
    ActiveWorkbook.Worksheets("大版积分").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("大版积分").Sort.SortFields.Add Key:=Range(ActiveCell.Address), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("大版积分").Sort
        .SetRange Range("A10:DD100")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    

   ActiveCell.Offset(0, 1).Select
   ActiveCell.FormulaR1C1 = "100"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "100"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "80"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "80"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "80"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "60"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "60"
   ActiveCell.Offset(1, 0).Select
   ActiveCell.FormulaR1C1 = "60"
   ActiveCell.Offset(1, 0).Select
    
  
End Sub
    


Sub 大版积分得分()

'4、8名以后考核分数大于30分，获得30分版主积分；
'20<=考核分数<30分，获得10分版主积分；
'15<=考核分数<20分，获得5分版主积分。
'

    Do While ActiveCell.Offset(0, -1).Value > 0
  
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]>=30,30,IF(RC[-1]>=20,10,IF(RC[-1]>=15,5,IF(RC[-1]>=0,0))))"
    
    ActiveCell.Offset(1, 0).Select
           
Loop
  
End Sub



Sub 大版积分汇总()

大版积分排序
大版积分得分

End Sub



Sub 小版积分排序()
'
' 小版积分从大到小排序
'


    ActiveCell.Offset(0, -1).Select
    ActiveWorkbook.Worksheets("小版积分").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("小版积分").Sort.SortFields.Add Key:=Range(ActiveCell.Address), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("小版积分").Sort
        .SetRange Range("A10:DD300")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

   ActiveCell.Offset(0, 1).Select
   
End Sub
    


Sub 小版积分得分()

'考核得分>=90分，当月积累100分版主积分
'80<=考核分数<90,当月积累80分版主积分
'70<=考核分数<80,当月积累60分版主积分
'50<=考核分数<70,当月积累40分版主积分
'30<=考核分数<50,当月积累20分版主积分
'15<=考核分数<30,当月积累10分版主积分


    Do While ActiveCell.Offset(0, -1).Value > 0
  
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]>=90,100,IF(RC[-1]>=80,80,IF(RC[-1]>=70,60,IF(RC[-1]>=50,40,IF(RC[-1]>=30,20,IF(RC[-1]>=15,10,IF(RC[-1]>=0,0)))))))"
    
    ActiveCell.Offset(1, 0).Select
           
Loop
  
End Sub



Sub 小版积分汇总()

小版积分排序
小版积分得分

End Sub

