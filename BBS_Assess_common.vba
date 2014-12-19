Sub 删图()

Dim pic As Shape

For Each pic In ActiveSheet.Shapes
    If pic.Width <> 0 Then
        pic.Select
        pic.Delete
    End If

Next

End Sub


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




Sub 打开多个链接()

Dim i As Integer
For i = 2 To 10   '行号
j = 5             '列号
Cells(i, j).Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
Next

End Sub




Function getname(HyperCell As Variant)

'获取链接名称

    Application.Volatile True
    getname = HyperCell.Hyperlinks(1).Name
End Function

Function geturl(HyperCell As Variant)

'获取链接URL

Application.Volatile True
    With HyperCell.Hyperlinks(1)
        geturl = IIf(.Address = "", .SubAddress, .Address)
    End With
End Function

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


Sub 自动链接_定制()

'自动将单元格中的URL变成链接形式


Do While ActiveCell.Value > 0

    ActiveCell.Select

    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=ActiveCell.Offset(0, 1).Value

    ActiveCell.Offset(1, 0).Select

Loop

End Sub





Sub 生成彩票数()

'取5个号码

    n = 1
    fanwei = 5
        
    For a = 1 To fanwei
        For b = 1 To fanwei
            For c = 1 To fanwei
                For d = 1 To fanwei
                  For e = 1 To fanwei
                     If a <> b And a <> c And a <> d And a <> e And b <> c And b <> d And b <> e And c <> d And c <> e And d <> e Then
                             
            
                Cells(n, 1) = a
                Cells(n, 2) = b
                Cells(n, 3) = c
                Cells(n, 4) = d
                Cells(n, 5) = d
                n = n + 1
                                
                    End If
                Next
            Next
        Next
    Next
Next

End Sub

Sub 生成彩票数2()

'取4个号码

    n = 1
    fanwei = 33
        
    For a = 1 To fanwei
        For b = 1 To fanwei
            For c = 1 To fanwei
                For d = 1 To fanwei
                 
      If a <> b And a <> c And a <> d And b <> c And b <> d And c <> d Then
                             
            
                Cells(n, 1) = a
                Cells(n, 2) = b
                Cells(n, 3) = c
                Cells(n, 4) = d
            
                n = n + 1
                                
                End If
            Next
        Next
    Next
Next


End Sub


Sub 生成连续数()

 n = 2
        
 For a = 1 To 10000
                    
    Cells(n, 1) = a
    n = n + 1
    
  Next

End Sub



Sub 数据展示()

'将有规律的一列数据分为多列显示

Dim i, j, k As Integer
j = 1
k = 1
For i = 1 To 270
    
    Sheets("总数").Select
    Cells(i, 1).Select
    Selection.Copy
    
    Sheets("分类").Select
    Cells(j, k).Select
    ActiveSheet.Paste
    ActiveCell.Offset(1, 0).Select

    i = i + 8
    j = j + 1
     
Next

 Sheets("总数").Select
 Range("A1").Select

End Sub
