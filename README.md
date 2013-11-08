VB
==

Method to one project

Option Explicit

Sub 点数模拟()
Dim 行 As Integer
Dim 列 As Integer
Dim 最大值 As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim temp As Integer
Dim 是否计算 As Integer


行 = Cells(3, "h").Value
Dim 列值(1 To 3) As Integer
'----清空上次数据
 For i = 11 To 11 + 行
  For j = 3 To 80
   Cells(j, i).Value = ""
 Next
Next

For i = 1 To 行

最大值 = Cells(2 + i, 4).Value
ReDim 表示范围(最大值) As Integer
'将表示范围初始化
For j = 0 To 最大值
表示范围(j) = 0
Next

For j = 1 To 3
列值(j) = Cells(2 + i, 4 + j).Value
Next

For j = 0 To 5
 For k = 0 To 5
  For l = 0 To 5
    temp = j * 列值(1) + k * 列值(2) + l * 列值(3)
    表示范围(temp) = 表示范围(temp) + 1
  Next
Next
Next


'打印最终结果
 For j = 1 To 最大值
  Cells(2 + j, 10 + i).Value = 表示范围(j)
 Next
 
Next
End Sub
