Sub ����()
'
' ����n�´�濼����ϸ������Ϊ�����´�濼����ϸ�������ں�������
'


    Sheets(2).Name = "���´�濼����ϸ"

    Sheets(3).Name = "���´�������"

    Sheets(4).Name = "���´�����÷�"
   
End Sub


Sub �����ĵ�()
'
' ���Ƹ����ĵ�
'

'

    ChDir "D:\����\2014��Ӫ����\�������˱�\3�·ݿ��˱�"
    
    
    Workbooks.Open Filename:="�����������ϸ��.2014-03-01.csv"
    Sheets("�����������ϸ��.2014-03-01").Select
    Sheets("�����������ϸ��.2014-03-01").Copy After:=Workbooks("3�´��������.xlsx").Sheets(4)
    
    Workbooks.Open Filename:="�������������.2014-03-01.csv"
    Sheets("�������������.2014-03-01").Select
    Sheets("�������������.2014-03-01").Copy After:=Workbooks("3�´��������.xlsx").Sheets(5)
    
    Workbooks.Open Filename:="������÷ֱ�.2014-03-01.csv"
    Sheets("������÷ֱ�.2014-03-01").Select
    Sheets("������÷ֱ�.2014-03-01").Copy After:=Workbooks("3�´��������.xlsx").Sheets(6)
   
    
End Sub




Sub �����������ϸ�ϲ�1()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(2).Select
 
 Range("L10").Select
 
 Do While ActiveCell.Offset(0, -11).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-11]&RC[-10]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub

Sub �����������ϸ�ϲ�2()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(5).Select
 
 Range("P2").Select
 
 Do While ActiveCell.Offset(0, -15).Value > 0
 
  
    ActiveCell.FormulaR1C1 = "=RC[-15]&RC[-14]"
    
    ActiveCell.Offset(1, 0).Select
          
 Loop
End Sub



Sub �����������ϸVlookup()
'
' ����vlookup��ʽ���������°�����һ�µĵط�'
'

   Sheets(5).Select
   Range("Q2").Select
    

Do While ActiveCell.Offset(0, -16).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'���´�濼����ϸ'!C[-5],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub �����������ϸ����()
'
' D�а��յ�������
'


    Sheets(5).Select
    Columns("Q:Q").Select
   
   
   
    ActiveWorkbook.Sheets(5).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(5).Sort.SortFields.Add Key:= _
        Range("Q2:Q100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Sheets(5).Sort
        .SetRange Range("A1:Q100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
End Sub


Sub �����������ϸ()
 
 
�����������ϸ�ϲ�1
�����������ϸ�ϲ�2
�����������ϸVlookup
�����������ϸ����

End Sub















Sub ������������ϲ�1()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(3).Select
 
 Range("N2").Select
 
 Do While ActiveCell.Offset(0, -13).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub

Sub ������������ϲ�2()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(6).Select
 
 Range("N2").Select
 
 Do While ActiveCell.Offset(0, -13).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub



Sub �����������Vlookup()
'
' ����vlookup��ʽ���������°�����һ�µĵط�'
'

   Sheets(6).Select
   Range("O2").Select
    

Do While ActiveCell.Offset(0, -14).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'���´�������'!C[-1],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub ���������������()
'
' D�а��յ�������
'


    Sheets(6).Select
    Columns("O:O").Select
   
   
   
    ActiveWorkbook.Sheets(6).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(6).Sort.SortFields.Add Key:= _
        Range("O2:O100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Sheets(6).Sort
        .SetRange Range("A1:O100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
End Sub


Sub �����������()
 
 
������������ϲ�1
������������ϲ�2
�����������Vlookup
���������������

End Sub



















Sub ������÷�Vlookup()
'
' ����vlookup��ʽ���������°���÷ֲ�һ�µĵط�
'

   Sheets(7).Select
   Range("M2").Select
    

Do While ActiveCell.Offset(0, -12).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-12],'���´�����÷�'!C[-12],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub ������÷�����()
'
' D�а��յ�������
'


    Sheets(7).Select
    Columns("M:M").Select
   
   
   
    ActiveWorkbook.Sheets(7).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(7).Sort.SortFields.Add Key:= _
        Range("M2:M100"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Sheets(7).Sort
        .SetRange Range("A1:M100")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
End Sub


Sub ������÷�()
 
 
������÷�Vlookup
������÷�����

End Sub

















Sub ת��2010��2003�ĵ�()

sPath = "C:\"

sFile = Dir(sPath & "*.docx")

While sFile <> ""

With Documents.Open(sPath & sFile)

.SaveAs Filename:=sPath & Replace(sFile, "docx", "doc"), FileFormat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
        True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
        False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
.Close

End With

sFile = Dir

Wend

End Sub


