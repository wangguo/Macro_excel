
Sub 小版改名()
'
' 将"n月小版考核明细"改名为"本月小版考核明细"，便于后续操作
'


    Sheets(2).Name = "本月小版考核明细"

    Sheets(3).Name = "本月小版管理动作"

    Sheets(4).Name = "本月小版版面得分"
   
End Sub


Sub 小版复制文档()
'
' 复制各个文档
'

'

    ChDir "D:\桌面\2014运营工作\版主考核表\2月份考核表"
    
    
    Workbooks.Open Filename:="小版主考核明细表.2014-02-01.csv"
    Sheets("小版主考核明细表.2014-02-01").Select
    Sheets("小版主考核明细表.2014-02-01").Copy After:=Workbooks("1月小版主考核.xlsx").Sheets(4)
    
    Workbooks.Open Filename:="小版主管理动作表.2014-02-01.csv"
    Sheets("小版主管理动作表.2014-02-01").Select
    Sheets("小版主管理动作表.2014-02-01").Copy After:=Workbooks("1月小版主考核.xlsx").Sheets(5)
    
    Workbooks.Open Filename:="小版版面得分表.2014-02-01.csv"
    Sheets("小版版面得分表.2014-02-01").Select
    Sheets("小版版面得分表.2014-02-01").Copy After:=Workbooks("1月小版主考核.xlsx").Sheets(6)
   
    
End Sub




Sub 小版主考核明细合并1()
'
' 选择单元格，将xx和xx合并到一起
'

 Sheets(2).Select
 
 Range("M13").Select
 
 Do While ActiveCell.Offset(0, -12).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-12]&RC[-11]&RC[-10]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub

Sub 小版主考核明细合并2()
'
' 选择单元格，将xx和xx合并到一起
'

 Sheets(5).Select
 
 Range("P2").Select
 
 Do While ActiveCell.Offset(0, -15).Value > 0
 
  
    ActiveCell.FormulaR1C1 = "=RC[-15]&RC[-14]&RC[-13]"
    
    ActiveCell.Offset(1, 0).Select
          
 Loop
End Sub



Sub 小版主考核明细Vlookup()
'
' 利用vlookup公式查找与上月版主不一致的地方'
'

   Sheets(5).Select
   Range("Q2").Select
    

Do While ActiveCell.Offset(0, -16).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'本月小版考核明细'!C[-4],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub 小版主考核明细排序()
'
' D列按照倒序排序
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


Sub 小版主考核明细()
 
 
小版主考核明细合并1
小版主考核明细合并2
小版主考核明细Vlookup
小版主考核明细排序

End Sub















Sub 小版主管理动作合并1()
'
' 选择单元格，将xx和xx合并到一起
'

 Sheets(3).Select
 
 Range("O2").Select
 
 Do While ActiveCell.Offset(0, -14).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-14]&RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub

Sub 小版主管理动作合并2()
'
' 选择单元格，将xx和xx合并到一起
'

 Sheets(6).Select
 
 Range("O2").Select
 
 Do While ActiveCell.Offset(0, -14).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-14]&RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub



Sub 小版主管理动作Vlookup()
'
' 利用vlookup公式查找与上月版主不一致的地方'
'

   Sheets(6).Select
   Range("P2").Select
    

Do While ActiveCell.Offset(0, -15).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'本月小版管理动作'!C[-1],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub 小版主管理动作排序()
'
' D列按照倒序排序
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


Sub 小版主管理动作()
 
 
小版主管理动作合并1
小版主管理动作合并2
小版主管理动作Vlookup
小版主管理动作排序

End Sub










Sub 小版版面得分合并1()
'
' 选择单元格，将xx和xx合并到一起
'

 Sheets(4).Select
 
 Range("N10").Select
 
 Do While ActiveCell.Offset(0, -13).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub




Sub 小版版面得分合并2()
'
' 选择单元格，将xx和xx合并到一起
'

 Sheets(7).Select
 
 Range("N2").Select
 
 Do While ActiveCell.Offset(0, -13).Value > 0
  
    ActiveCell.FormulaR1C1 = "=RC[-13]&RC[-12]"
    
    ActiveCell.Offset(1, 0).Select
      
Loop

End Sub








Sub 小版版面得分Vlookup()
'
' 利用vlookup公式查找与上月版面得分不一致的地方
'

   Sheets(7).Select
   Range("O2").Select
    

Do While ActiveCell.Offset(0, -14).Value > 0

    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'本月小版版面得分'!C[-1],1,FALSE)"
    
     ActiveCell.Offset(1, 0).Select
          
 Loop
    
End Sub



Sub 小版版面得分排序()
'
' D列按照倒序排序
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


Sub 小版版面得分()
 

小版版面得分合并1
小版版面得分合并2
小版版面得分Vlookup
小版版面得分排序

End Sub














Sub 复制文档()
'
' 复制各个文档
'

'

    ChDir "D:\桌面\2014运营工作\版主考核表\2月份考核表"
    
    
    Workbooks.Open Filename:="小版主考核明细表.2014-02-01.csv"
    Sheets("小版主考核明细表.2014-02-01").Select
    Sheets("小版主考核明细表.2014-02-01").Copy After:=Workbooks("1月小版主考核.xlsx").Sheets(4)
    
    Workbooks.Open Filename:="小版主管理动作表.2014-02-01.csv"
    Sheets("小版主管理动作表.2014-02-01").Select
    Sheets("小版主管理动作表.2014-02-01").Copy After:=Workbooks("1月小版主考核.xlsx").Sheets(5)
    
    Workbooks.Open Filename:="小版版面得分表.2014-02-01.csv"
    Sheets("小版版面得分表.2014-02-01").Select
    Sheets("小版版面得分表.2014-02-01").Copy After:=Workbooks("1月小版主考核.xlsx").Sheets(6)
   
    
End Sub


Sub 转换2010到2003文档()

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








