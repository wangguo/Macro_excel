
Sub С�����()
'
' ��"n��С�濼����ϸ"����Ϊ"����С�濼����ϸ"�����ں�������
'


    Sheets(2).Name = "����С�濼����ϸ"

    Sheets(3).Name = "����С�������"

    Sheets(4).Name = "����С�����÷�"
   
End Sub


Sub С�渴���ĵ�()
'
' ���Ƹ����ĵ�
'

'

    ChDir "D:\����\2014��Ӫ����\�������˱�\2�·ݿ��˱�"
    
    
    Workbooks.Open Filename:="С����������ϸ��.2014-02-01.csv"
    Sheets("С����������ϸ��.2014-02-01").Select
    Sheets("С����������ϸ��.2014-02-01").Copy After:=Workbooks("1��С��������.xlsx").Sheets(4)
    
    Workbooks.Open Filename:="С������������.2014-02-01.csv"
    Sheets("С������������.2014-02-01").Select
    Sheets("С������������.2014-02-01").Copy After:=Workbooks("1��С��������.xlsx").Sheets(5)
    
    Workbooks.Open Filename:="С�����÷ֱ�.2014-02-01.csv"
    Sheets("С�����÷ֱ�.2014-02-01").Select
    Sheets("С�����÷ֱ�.2014-02-01").Copy After:=Workbooks("1��С��������.xlsx").Sheets(6)
   
    
End Sub




Sub С����������ϸ�ϲ�1()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(2).Select
 
 Range("M13").Select
 
 Do While ActiveCell.Offset(0, -12).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-12]&RC[-11]&RC[-10]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub

Sub С����������ϸ�ϲ�2()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(5).Select
 
 Range("P2").Select
 
 Do While ActiveCell.Offset(0, -15).Value > 0
 
  
    ActiveCell.FormulaR1C1 = "=RC[-15]&RC[-14]&RC[-13]"
    
    ActiveCell.Offset(1, 0).Select
          
 Loop
End Sub



Sub С����������ϸVlookup()
'
' ����vlookup��ʽ���������°�����һ�µĵط�'
'

   Sheets(5).Select
   Range("Q2").Select
    

Do While ActiveCell.Offset(0, -16).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'����С�濼����ϸ'!C[-4],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub С����������ϸ����()
'
' D�а��յ�������
'


    Sheets(5).Select
    Columns("Q:Q").Select
   
   
   
    ActiveWorkbook.Sheets(5).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(5).Sort.SortFields.Add Key:= _
        Range("Q2:Q400"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Sheets(5).Sort
        .SetRange Range("A1:Q400")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
End Sub


Sub С����������ϸ()
 
 
С����������ϸ�ϲ�1
С����������ϸ�ϲ�2
С����������ϸVlookup
С����������ϸ����

End Sub















Sub С�����������ϲ�1()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(3).Select
 
 Range("O2").Select
 
 Do While ActiveCell.Offset(0, -14).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-14]&RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub

Sub С�����������ϲ�2()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(6).Select
 
 Range("O2").Select
 
 Do While ActiveCell.Offset(0, -14).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-14]&RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub



Sub С����������Vlookup()
'
' ����vlookup��ʽ���������°�����һ�µĵط�'
'

   Sheets(6).Select
   Range("P2").Select
    

Do While ActiveCell.Offset(0, -15).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'����С�������'!C[-1],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub С��������������()
'
' D�а��յ�������
'


    Sheets(6).Select
    Columns("P:P").Select
   
   
   
    ActiveWorkbook.Sheets(6).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(6).Sort.SortFields.Add Key:= _
        Range("P2:P400"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Sheets(6).Sort
        .SetRange Range("A1:P400")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
End Sub


Sub С����������()
 
 
С�����������ϲ�1
С�����������ϲ�2
С����������Vlookup
С��������������

End Sub










Sub С�����÷ֺϲ�1()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(4).Select
 
 Range("N10").Select
 
 Do While ActiveCell.Offset(0, -13).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub




Sub С�����÷ֺϲ�2()
'
' ѡ��Ԫ�񣬽�xx��xx�ϲ���һ��
'

 Sheets(7).Select
 
 Range("N2").Select
 
 Do While ActiveCell.Offset(0, -13).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub








Sub С�����÷�Vlookup()
'
' ����vlookup��ʽ���������°���÷ֲ�һ�µĵط�
'

   Sheets(7).Select
   Range("O2").Select
    

Do While ActiveCell.Offset(0, -14).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'����С�����÷�'!C[-1],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub С�����÷�����()
'
' D�а��յ�������
'


    Sheets(7).Select
    Columns("M:M").Select
   
   
   
    ActiveWorkbook.Sheets(7).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(7).Sort.SortFields.Add Key:= _
        Range("O2:O400"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Sheets(7).Sort
        .SetRange Range("A1:O400")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
End Sub


Sub С�����÷�()
 

С�����÷ֺϲ�1
С�����÷ֺϲ�2
С�����÷�Vlookup
С�����÷�����

End Sub














Sub �����ĵ�()
'
' ���Ƹ����ĵ�
'

'

    ChDir "D:\����\2014��Ӫ����\�������˱�\2�·ݿ��˱�"
    
    
    Workbooks.Open Filename:="С����������ϸ��.2014-02-01.csv"
    Sheets("С����������ϸ��.2014-02-01").Select
    Sheets("С����������ϸ��.2014-02-01").Copy After:=Workbooks("1��С��������.xlsx").Sheets(4)
    
    Workbooks.Open Filename:="С������������.2014-02-01.csv"
    Sheets("С������������.2014-02-01").Select
    Sheets("С������������.2014-02-01").Copy After:=Workbooks("1��С��������.xlsx").Sheets(5)
    
    Workbooks.Open Filename:="С�����÷ֱ�.2014-02-01.csv"
    Sheets("С�����÷ֱ�.2014-02-01").Select
    Sheets("С�����÷ֱ�.2014-02-01").Copy After:=Workbooks("1��С��������.xlsx").Sheets(6)
   
    
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








