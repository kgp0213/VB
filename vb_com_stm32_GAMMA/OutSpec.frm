VERSION 5.00
Begin VB.Form OutSpec 
   Appearance      =   0  'Flat
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "OutSpec"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11040
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "数据超规！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   4800
   End
End
Attribute VB_Name = "OutSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Command1.SetFocus
End Sub
