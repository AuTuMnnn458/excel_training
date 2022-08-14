# Excel_training
This is my excel training including some tests for excel function using.
主要记录在学习vba过程中解决的一些excel实例。

## 1.找出文档左边数据中最近3次购买牛肉的记录并存放在右边。
![image](https://github.com/AuTuMnnn458/excel_training/blob/main/QQ%E6%88%AA%E5%9B%BE20220814161218.jpg)

```
Sub test()
n = Cells(Rows.Count, 1).End(xlUp).Row
For i = n To 2 Step -1
    If Cells(i, 2) = "牛肉" Then
        x = x + 1
        If x <= 3 Then
            Cells(i, 2).Offset(0, -1).Resize(1, 3).Copy Cells(Rows.Count, "e").End(xlUp).Offset(1, 0)
        Else
            Exit Sub
        End If
    End If
Next i
End Sub
```
代码优化：
```
Sub test()
n = Cells(Rows.Count, 1).End(xlUp).Row
For i = n To 2 Step -1
    If Cells(i, 2) = "牛肉" Then
        Cells(i, 2).Offset(0, -1).Resize(1, 3).Copy Cells(Rows.Count, "e").End(xlUp).Offset(1, 0)
        x = x + 1
        If x = 3 Then Exit For
        End If
    End If
Next i
MsgBox "数据处理完毕"
End Sub
```
