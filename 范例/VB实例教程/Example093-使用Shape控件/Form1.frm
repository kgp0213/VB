VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shape"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6765
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "填充风格"
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   2775
      Begin VB.OptionButton Option3 
         Caption         =   "透明"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "交叉线"
         Height          =   495
         Index           =   6
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "水平线"
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "垂直线"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "边框类型"
      Height          =   1695
      Left            =   3480
      TabIndex        =   5
      Top             =   3000
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "透明"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "点划线"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "短划线"
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "点"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "实线"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "形状"
      Height          =   2415
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "圆角矩形"
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "圆"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "正方形"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "矩形"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   2655
      Left            =   240
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Option1_Click(Index As Integer)
'设置形状
Select Case Index
Case 0
    Shape1.Shape = 0
Case 1
    Shape1.Shape = 1
Case 2
    Shape1.Shape = 3
Case 3
    Shape1.Shape = 4
End Select
End Sub

Private Sub Option2_Click(Index As Integer)
'设置边框风格
Shape1.BorderWidth = 1
Select Case Index
Case 0
    Shape1.BorderStyle = 0
Case 1
    Shape1.BorderStyle = 1
Case 2
    Shape1.BorderStyle = 2
Case 3
    Shape1.BorderStyle = 3
Case 4
    Shape1.BorderStyle = 4
End Select

End Sub

Private Sub Option3_Click(Index As Integer)
'设置填充风格
Select Case Index
Case 1
    Shape1.FillStyle = 1
Case 2
    Shape1.FillStyle = 2
Case 3
    Shape1.FillStyle = 3
Case 6
    Shape1.FillStyle = 6
End Select

End Sub
