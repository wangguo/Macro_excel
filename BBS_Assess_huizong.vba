Sub 大版考核结果()
'
' 根据得分评定优秀合格
'

 
 Range("F8").Select
 
 Do While ActiveCell.Offset(0, -3).Value > 0
  
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]>=60,""优秀"",IF(RC[-1]>=40,""达标"",IF(RC[-1]>=15,""合格"",IF(RC[-1]>0,""不合格"",""无排名""))))"
    
    ActiveCell.Offset(1, 0).Select
    
       
Loop
  
End Sub


Sub 大版考核排名()
'
' 自动填充相关排名
'

 

 
 Do While ActiveCell.Offset(0, 4).Value > 0
 
  If ActiveCell.Offset(0, 3).Value > 0 Then
  
    ActiveCell.FormulaR1C1 = _
        "=RANK(RC[3],R8C5:R40C5)"
    
    ActiveCell.Offset(1, 0).Select
            
Else

 ActiveCell.Value = "无排名"
 ActiveCell.Offset(1, 0).Select
 
 End If
            
Loop

  
End Sub


Sub 统计大版考核各级人数()

'统计优秀、达标、合格版主的个数


Range("D3").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[2],""优秀"")"

Range("D4").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[2],""达标"")"

Range("D5").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[2],""合格"")"
        
   
End Sub





Sub 大版考核排名汇总()

大版考核结果
Range("B8").Select
大版考核排名
统计大版考核各级人数
   
End Sub









Sub 小版考核结果()
'
' 根据得分评定优秀合格
'

 
 Range("H8").Select
 
 Do While ActiveCell.Offset(0, -4).Value > 0
  
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]>=80,""优秀"",IF(RC[-1]>=50,""达标"",IF(RC[-1]>=15,""合格"",IF(RC[-1]>0,""不合格"",""无排名""))))"
    
    ActiveCell.Offset(1, 0).Select
    
       
Loop
  
End Sub


Sub 小版考核排名()
'
' 自动填充相关排名
'
 

 
 Do While ActiveCell.Offset(0, 5).Value > 0
 
  If ActiveCell.Offset(0, 4).Value > 0 Then
  
    ActiveCell.FormulaR1C1 = _
        "=RANK(RC[4],R8C7:R300C7)"
    
    ActiveCell.Offset(1, 0).Select
            
Else

 ActiveCell.Value = "无排名"
 ActiveCell.Offset(1, 0).Select
 
 End If
            
Loop

  
End Sub


Sub 统计小版考核各级人数()

'统计优秀、达标、合格版主的个数


Range("E3").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[3],""优秀"")"

Range("E4").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[3],""达标"")"

Range("E5").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[3],""合格"")"
        
   
End Sub





Sub 小版考核排名汇总()

小版考核结果
Range("C8").Select
小版考核排名
统计小版考核各级人数
   
End Sub


