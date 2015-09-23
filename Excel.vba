Sub 自动链接()

'自动将单元格中的URL变成链接形式


Do While ActiveCell.Value > 0

    ActiveCell.Select

    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=ActiveCell.Value

    ActiveCell.Offset(1, 0).Select

Loop

End Sub

Sub 删链接()

'清除表中所有链接

ActiveSheet.Hyperlinks.Delete

End Sub





Sub 删图()

Dim pic As Shape

For Each pic In ActiveSheet.Shapes
    If pic.Width <> 0 Then
        pic.Select
        pic.Delete
    End If

Next

End Sub


Function getname(HyperCell As Variant)
    Application.Volatile True
    getname = HyperCell.Hyperlinks(1).Name
End Function

Function geturl(HyperCell As Variant)
Application.Volatile True
    With HyperCell.Hyperlinks(1)
        geturl = IIf(.Address = "", .SubAddress, .Address)
    End With
End Function



Sub 自动补全ID()

Do While ActiveCell.Value > 0

    ActiveCell.Select
        
    ActiveCell.Value = "[url=http://www.csdn.net/blog/" & ActiveCell.Value & "]" & ActiveCell.Value & "[/url]"
     

    ActiveCell.Offset(1, 0).Select

Loop

End Sub


Sub 自动补全资源URL()

Do While ActiveCell.Value > 0

    ActiveCell.Select
    
    
    ActiveCell.Value = "[url=" & ActiveCell.Offset(0, 7).Value & "]" & ActiveCell.Value & "[/url]"
       

    ActiveCell.Offset(1, 0).Select

Loop

End Sub





Sub 删除偶数行值()
'
' 删除偶数行值 宏
'

    
    
    Range("B12").Select
    
 Do While ActiveCell.Value > 0
 
   Selection.EntireRow.Delete
        
    ActiveCell.Offset(1, 0).Select
          
 Loop
 
End Sub
    
 
Sub 数据整理()

'将固定行数的1列值转置

Dim i, j, k As Integer

a = 1
j = 1
k = a

For i = a To 300
       
    Sheets("总数").Select
    Cells(i, 1).Select
    Selection.Copy
    Sheets("整理").Select
    Cells(j, k).Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Select
      
    i = i + 4
    j = j + 1
    
    Next
 Sheets("总数").Select
 Range("A1").Select

End Sub



Sub 批量打开链接()
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim h As Integer

l = InputBox("请输入链接所在列号", Title, 10)
i = InputBox("请输入起始行号", Title, 1)
j = InputBox("请输入终止行号", Title, 10)

For h = 1 To 3
Cells(h, 4).Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
Next
End Sub





Sub 挑选楼层()



Dim i

For i = 1 To 200

Sheets("Sheet2").Select
If Cells(i, 1).Value > 0 Then

Range(Cells(i, 1), Cells(i, 2)).Select

Selection.Copy
  
Sheets("Sheet3").Select
 
ActiveSheet.Paste

ActiveCell.Offset(1, 0).Select

End If

Next i

End Sub

