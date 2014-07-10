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
    
 
