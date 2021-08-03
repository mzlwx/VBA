Option Base 1
Sub 人机检查表()
Dim i#, j#, n#, c#
n = Sheets("数据库").Cells(1, 2).End(xlDown).Row - 1
'c = Sheets("数据库").UsedRange.Columns.Count
c = 5

Dim l As Integer '字符串长度
Dim x As Integer

'将数据库的数据导入数组中
Dim data() As Variant
ReDim data(n, c) As Variant
For i = 1 To n
    For j = 1 To c
        data(i, j) = Sheets("数据库").Cells(i + 1, j)
    Next
Next

'将数组中的数据导入模板中
For i = 1 To n

        '方案名
        l = Len(data(i, 2))

        x = Int(l / 4)
        
        Sheets("模板").Range(Cells(7, 2), Cells(11, 10)) = ""
        Sheets("模板").Range(Cells(7, 2), Cells(7, 10)).UnMerge
        Sheets("模板").Range(Cells(7, 2), Cells(7, 10)).Font.Underline = xlUnderlineStyleNone
        Sheets("模板").Range(Cells(7, 2), Cells(7, 3 + x)).Merge
        Sheets("模板").Cells(7, 2) = data(i, 2)
        
        '加下划线
        Sheets("模板").Range(Cells(7, 2), Cells(7, 3 + x)).Font.Underline = xlUnderlineStyleSingle
        
        Sheets("模板").Range(Cells(7, 4 + x), Cells(7, 10)).Merge
        Sheets("模板").Cells(7, 4 + x) = ",请予审查和批准。"

        
        '附件1
        Sheets("模板").Cells(9, 2) = 1 & "、" & data(i, 3)

        '附件2
        If data(i, 4) = "" Then
        
        Else
            Sheets("模板").Cells(10, 2) = 2 & "、" & data(i, 4)
        End If
        
        
        '附件3
        If data(i, 5) = "" Then
        
        Else
            Sheets("模板").Cells(11, 2) = 3 & "、" & data(i, 5)
        End If
        
        
        '设置字体和对齐
        Sheets("模板").Range(Cells(7, 2), Cells(11, 10)).Font.Name = "宋体"
        Sheets("模板").Range(Cells(7, 2), Cells(11, 10)).Font.Size = 12
        Sheets("模板").Range(Cells(7, 2), Cells(9, 10)).HorizontalAlignment = xlHAlignLeft
Sheets("模板").Range(Cells(7, 2), Cells(7, 3 + x)).HorizontalAlignment = xlHAligncenter
        


    '标黄色的不打印，其余的打印
    If Sheets("数据库").Cells(i + 1, 2).Interior.Color = 65535 Then 'Yellow
    
    Else
        Sheets("模板").Range(Cells(1, 1), Cells(47, 10)).PrintOut
        'Sheets("模板").Range(Cells(1, 1), Cells(47, 10)).PrintOut
        'Sheets("模板").Range(Cells(1, 1), Cells(47, 10)).PrintOut
    End If
    
Next

End Sub
 

