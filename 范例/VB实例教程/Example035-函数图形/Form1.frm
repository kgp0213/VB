VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "函数图形"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   6345
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      FillColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4635
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
    Picture1.Scale (-300, 300)-(300, -300)
    '设置Picture1控件使其中心点坐标为(0,0)
    
    Picture1.Cls             '清除其他图形
    
    '以下三行代码实现在Picture1中绘制黑色的水平坐标轴和垂直坐标轴
    Picture1.ForeColor = &H0
    Picture1.Line (-300, 0)-(300, 0)
    Picture1.Line (0, 300)-(0, -300)
    
    '开始绘制白色函数图形
    Picture1.ForeColor = &HFFFFFF
    For a = 0 To 200 Step 20
    For th = 0 To 2 * 3.1415926 + 0.1 Step 3.1415926 / 32
        x = a * Cos(th) ^ 3
        y = a * Sin(th) ^ 3
        Picture1.Line -(x, y)
    Next th
    Next
End Sub
