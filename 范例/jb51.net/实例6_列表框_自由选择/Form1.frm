VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   1710
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":001F
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "期待你的选择！"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "请你预测2006世界杯冠军："
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_Click()
Select Case List1.ListIndex
Case 0
Label2.Caption = "华而不实，焉能折桂？"
Case 1
Label2.Caption = "豪华阵容，时运不佳"
Case 2
Label2.Caption = "关键时刻腿软"
Case 3
Label2.Caption = "德国战车破旧不堪"
Case 4
Label2.Caption = "拉丁艺术仅供欣赏"
Case 5
Label2.Caption = "阴沟里翻船，不慎！"
Case 6
Label2.Caption = "现代足球不能靠防守度日"
Case 7
Label2.Caption = "难成大器！"
Case 8
Label2.Caption = "恭喜你，答对了！"
End Select

End Sub


