VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "函数过程"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "调用函数计算从1 到 N 的和"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "请输入 N 值"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'本程序主要练习子定义函数
'注意格式和参数
'注意一下调用函数的程序执行顺序

Private Sub Command1_Click()

    Dim a As Integer
    Dim b As Integer
    a = Form1.Text1.Text   'a 就是输入的 N 值，随便起的变量名字
    
    b = Plus_N(a) '调用函数过程，到这里后要跳到函数处执行函数。然后回来继续往下执行
    
    Form1.Text1.Text = b  'b 就是传出的累加和
End Sub
'这里定义了两个参数，N 用来传入一个数，M 用来传出累加的和,Plus_N 是自己起的名字
Private Function Plus_N(N As Integer)

    Dim I As Integer
    Dim Sum As Integer
    For I = 1 To N
        Sum = Sum + I
    Next I
    Plus_N = Sum
   
End Function




