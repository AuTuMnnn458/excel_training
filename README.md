# Excel_training
This is my excel training including some tests for excel function using.
主要记录在学习vba过程中解决的一些excel实例。

## 1.找出文档左边数据中最近3次购买牛肉的记录并存放在右边。
![image](https://github.com/AuTuMnnn458/excel_training/blob/main/pictures/1.jpg)

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


## 2.写一个循环判断是否正确填写自己生日的对话框
特别注意如果漏写了end if，系统会提示loop没有do。
```
Sub test()
Dim n As Date
On Error Resume Next
Do
    n = InputBox("请输入我的生日(yyyy-mm-dd)")
    If Err.Number <> 0 Then
        MsgBox "输入格式有误": GoTo 1
    End If
    
    If n = "1997-08-03" Then
        MsgBox "回答正确，循环结束"
        Exit Do
    Else
        MsgBox "回答错误，请继续回答"
    End If
1:
    Err.Clear

Loop
End Sub
```


## 3.用工作表函数实现求平均和计数
![image](https://github.com/AuTuMnnn458/excel_training/blob/main/pictures/3.jpg)

```
Sub test()
[c23] = Application.WorksheetFunction.AverageIf([b:b], "牛肉", [c:c])
[c24] = Application.WorksheetFunction.CountIfs([b:b], "猪肉", [c:c], ">50")
End Sub
```


## 4.三数之和
![image](https://github.com/AuTuMnnn458/excel_training/blob/main/pictures/4.jpg)

经典算法题三数之和

方法一：暴力解法，随机数生成3个指针，然后使用do loop循环
```
Sub test()
Dim s1%, s2%, s3%, n%, h%, k%
n = Cells(Rows.Count, 1).End(xlUp).Row
Do
s1 = Int((n - 2 + 1) * Rnd + 2)
s2 = Int((n - 2 + 1) * Rnd + 2)
s3 = Int((n - 2 + 1) * Rnd + 2)
h = Cells(s1, 1) + Cells(s2, 1) + Cells(s3, 1)
k = k + 1
Loop Until h = [b2]
Cells(s1, 1).Interior.ColorIndex = 3
Cells(s2, 1).Interior.ColorIndex = 3
Cells(s3, 1).Interior.ColorIndex = 3
MsgBox "循环了" & k & "次"
End Sub
```


方法二：哈希表，用两个指针，要求的值减掉这两个指针所指向的值后，剩下的值与表中的数做匹配。
```
Dim i%, y%, x&, n&, c&
On Error Resume Next
s = Timer
n = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To n
    For y = 3 To n
        x = Cells(i, 1) + Cells(y, 1)
        If x < [b2] Then
            c = [b2] - x
            t = Application.WorksheetFunction.Match(c, [a:a], 0)
            If t >= 1 Then GoTo 1
        End If
    Next y
Next i
1:
Cells(i, 1).Interior.ColorIndex = 3
Cells(y, 1).Interior.ColorIndex = 3
Cells(t, 1).Interior.ColorIndex = 3
MsgBox "用时" & Timer - s & "秒"
End Sub
```
## 5.抓人分组
要求：不能重复抓。

注意：这里7可以整除21，也可以写一个判断处理不整除的情况。

![image](https://github.com/AuTuMnnn458/excel_training/blob/main/pictures/5.jpg)

```
Sub test()
Dim gp%, y%, s%, h%
y = Application.WorksheetFunction.CountA([a:a])
Do
s = Int((y - 2 + 1) * Rnd + 2)
h = Cells(Rows.Count, 2).End(xlUp).Row
Set b = range("b:b").Find(Cells(s, 1))
If b Is Nothing Then
    Cells(h + 1, 2) = Cells(s, 1)
    gp = gp + 1
End If
Loop Until gp = 7
End Sub
```
