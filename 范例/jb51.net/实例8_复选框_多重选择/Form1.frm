VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   1800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Caption         =   "其他"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "购物"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "音乐"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "体育活动"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "请你选择自己的爱好"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public txt As String
Public txt1 As String
Public txt2 As String
Public txt3 As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
txt = "体育活动 "
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
txt1 = "音乐 "
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
txt2 = "购物 "
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
txt3 = "其他 "
End If
End Sub

Private Sub Command1_Click()
MsgBox ("你的爱好是：") & txt & txt1 & txt2 & txt3
Unload Form1
End Sub

