VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   2055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Off"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton Option3 
         Caption         =   "不能确定"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "错误"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "正确"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "可以利用框架控件把其他控件组织在一起，形成控件组？"
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frame1.Visible = True
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub
