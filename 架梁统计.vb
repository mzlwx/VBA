Option Base 1
Dim last_row#, i#, j#, n#
Dim a() As Variant

Sub document()
MsgBox ("1. 在工作表【设置】中输入新添加的桥名及其梁板总数" & Chr(13) & "2. 点击【运行】按钮" & Chr(13) & "3. 请勿修改工作表的名称")
End Sub

Sub counter()


'录入桥梁名称及梁板数量

last_row = Sheets("设置").Range("A65535").End(xlUp).Row

'将完成率区域设置未无填充
Sheets("设置").Range(Cells(1, 4), Cells(last_row, 4)).Interior.Pattern = xlNone


If Sheets("设置").Cells(last_row, 1) = "总计" Then
    Rows(last_row).Delete
    last_row = Sheets("设置").Range("A65535").End(xlUp).Row
    n = last_row - 1 'n为桥梁数量
    num (n)

Else
    n = last_row - 1 'n为桥梁数量
    num (n)
    
End If

End Sub

Function num(x) As Integer
ReDim a(n, 3) '定义一个n行3列矩阵，第1列储存桥名，第2列储存梁板总数，第3列储存累计完成量

For i = 2 To n + 1
    a(i - 1, 1) = Sheets("设置").Cells(i, 1)
    a(i - 1, 2) = Sheets("设置").Cells(i, 2)
Next


Dim d_sum#, day$
day = Sheets("统计表").Cells(4, 1)


'统计
Dim last_row2#
last_row2 = Sheets("统计表").UsedRange.Rows.Count



    For i = 1 To last_row2
        For j = 1 To n

            If Sheets("统计表").Cells(i, 1) = a(j, 1) And Sheets("统计表").Cells(i, 8) = 1 Then
        
                a(j, 3) = a(j, 3) + 1 '该桥累计梁板完成量+1
                Sheets("设置").Cells(j + 1, 3) = a(j, 3)
                Sheets("设置").Range(Cells(j + 1, 1), Cells(j + 1, 3)).Interior.Color = 65535 '标黄
                Sheets("设置").Range(Cells(j + 1, 1), Cells(j + 1, 3)).Interior.Pattern = xlNone
                Sheets("统计表").Cells(i, 10) = a(j, 3) / a(j, 2)
                
            If Sheets("统计表").Cells(i, 3) = day Then
                d_sum = d_sum + 1
            Else
                day = Sheets("统计表").Cells(i, 3)
                d_sum = 1
            End If
            
            Sheets("设置").Cells(2, 8) = Sheets("统计表").Cells(i, 3)
            Sheets("设置").Cells(4, 8) = d_sum
                
                
            Else
            End If
            
        Next
    Next
    
            
'统计各桥梁板架设完成率
For i = 2 To n + 1
    Sheets("设置").Cells(i, 4) = a(i - 1, 3) / a(i - 1, 2)
    If Sheets("设置").Cells(i, 4) = 1 Then
        Sheets("设置").Cells(i, 4).Interior.Color = 5287936
    End If
Next
    
    
 '总计
Dim sum1#, sum2# 'sum1为各桥梁板总数之和，sum2为各桥累计完成量之和


If Sheets("设置").Cells(last_row, 1) = "总计" Then

Else
    For i = 1 To n
        sum1 = sum1 + a(i, 2)
        sum2 = sum2 + a(i, 3)
    Next
    
    Sheets("设置").Cells(last_row + 1, 1) = "总计"
    Sheets("设置").Cells(last_row + 1, 2) = sum1
    Sheets("设置").Cells(last_row + 1, 3) = sum2
    Sheets("设置").Cells(last_row + 1, 4) = sum2 / sum1

    For i = 1 To 4
        Sheets("设置").Cells(last_row + 1, i).Font.Bold = True
        Sheets("设置").Cells(last_row + 1, i).Font.Color = 255
    Next
    
End If

End Function

