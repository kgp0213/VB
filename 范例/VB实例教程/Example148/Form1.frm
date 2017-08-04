VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5310
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command_Draw 
      Caption         =   "Draw"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim excel As Application
Dim excelbook As Workbook
Dim excelsheet As Worksheet
'声明变量
Dim x(1 To 4, 1 To 4) As Integer

Private Sub Command_Draw_Click()
    Randomize
    For i = 1 To 4
       For j = 1 To 4
          x(i, j) = 100 * Rnd()
          '数组元素赋值
       Next j
    Next i
    excel.Range("a1:d4").Value = x
    '填充数组
    excel.Range("a1:d4").Select
    '选中从第一行第一列到第四行第四列的数据
    Set excelchart = excel.Charts.Add
    '制成图表
    excelsheet.Application.Visible = True
    '使Excel处于可见的状态
    excelsheet.SaveAs App.Path + "\test.xls"
    '存储文件
End Sub

Private Sub Form_Load()
    Set excel = CreateObject("Excel.application")
    Set excelbook = excel.Workbooks().Add
    Set excelsheet = excelbook.Worksheets("sheet1")
    '声明并创建Microsoft Excel引用
End Sub

Private Sub Form_Unload(Cancel As Integer)
    excelsheet.Application.Quit
    '关闭 Excel。
    Set excelsheet = Nothing
    '释放对象
End Sub
