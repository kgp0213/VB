VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "BinaryFile"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   4395
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim c As Variant
    c = "abcd"
    Dim numFile As Integer
    numFile = FreeFile()
    Open App.Path + "\test.dat" For Binary Access Read Write As #numFile
    Put #numFile, 1, c
    Get #numFile, 1, c
    Form1.Print c
    Close #numFile
        
    Dim MyArray1(2, 1) As Integer
    Dim MyArray2(2, 1) As Integer
    '声明Integer类型数组

    MyArray1(0, 0) = 10
    MyArray1(0, 1) = 6

    MyArray1(1, 0) = 9
    MyArray1(1, 1) = 66

    MyArray1(2, 0) = 8
    MyArray1(2, 1) = 888
    '为数组MyArray1赋值
    numFile = FreeFile()
    Open App.Path + "\test.dat" For Binary Access Read Write As #numFile
    Put #numFile, Len(c) * 8, MyArray1
    '将数组MyArray1的内容存入test.dat
    '注意指定的存入位置
    'Len(c)得到的是字符串中字符的个数
    '而每个字符占8bit
    '所以MyArray1存储在先前写入的内容后
    '必须设置开始存入位置为Len(c)*8

    Get #numFile, Len(c) * 8, MyArray2
    '将指定位置的内容存入MyArray2
    Dim i, j As Integer
    For i = 0 To 2
        For j = 0 To 1
        Form1.Print MyArray2(i, j)
        Next
    Next
    '显示MyArray2内容
    Close #numFile
End Sub
