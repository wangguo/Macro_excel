Sub ��濼�˽��()
'
' ���ݵ÷���������ϸ�
'

 
 Range("F8").Select
 
 Do While ActiveCell.Offset(0, -3).Value > 0
  
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]>=60,""����"",IF(RC[-1]>=40,""���"",IF(RC[-1]>=15,""�ϸ�"",IF(RC[-1]>0,""���ϸ�"",""������""))))"
    
    ActiveCell.Offset(1, 0).Select
    
       
Loop
  
End Sub


Sub ��濼������()
'
' �Զ�����������
'

 

 
 Do While ActiveCell.Offset(0, 4).Value > 0
 
  If ActiveCell.Offset(0, 3).Value > 0 Then
  
    ActiveCell.FormulaR1C1 = _
        "=RANK(RC[3],R8C5:R40C5)"
    
    ActiveCell.Offset(1, 0).Select
            
Else

 ActiveCell.Value = "������"
 ActiveCell.Offset(1, 0).Select
 
 End If
            
Loop

  
End Sub


Sub ͳ�ƴ�濼�˸�������()

'ͳ�����㡢��ꡢ�ϸ�����ĸ���


Range("D3").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[2],""����"")"

Range("D4").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[2],""���"")"

Range("D5").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[2],""�ϸ�"")"
        
   
End Sub





Sub ��濼����������()

��濼�˽��
Range("B8").Select
��濼������
ͳ�ƴ�濼�˸�������
   
End Sub









Sub С�濼�˽��()
'
' ���ݵ÷���������ϸ�
'

 
 Range("H8").Select
 
 Do While ActiveCell.Offset(0, -4).Value > 0
  
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-1]>=80,""����"",IF(RC[-1]>=50,""���"",IF(RC[-1]>=15,""�ϸ�"",IF(RC[-1]>0,""���ϸ�"",""������""))))"
    
    ActiveCell.Offset(1, 0).Select
    
       
Loop
  
End Sub


Sub С�濼������()
'
' �Զ�����������
'
 

 
 Do While ActiveCell.Offset(0, 5).Value > 0
 
  If ActiveCell.Offset(0, 4).Value > 0 Then
  
    ActiveCell.FormulaR1C1 = _
        "=RANK(RC[4],R8C7:R300C7)"
    
    ActiveCell.Offset(1, 0).Select
            
Else

 ActiveCell.Value = "������"
 ActiveCell.Offset(1, 0).Select
 
 End If
            
Loop

  
End Sub


Sub ͳ��С�濼�˸�������()

'ͳ�����㡢��ꡢ�ϸ�����ĸ���


Range("E3").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[3],""����"")"

Range("E4").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[3],""���"")"

Range("E5").Select
ActiveCell.FormulaR1C1 = "=COUNTIF(C[3],""�ϸ�"")"
        
   
End Sub





Sub С�濼����������()

С�濼�˽��
Range("C8").Select
С�濼������
ͳ��С�濼�˸�������
   
End Sub


