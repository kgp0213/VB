VERSION 5.00
Begin VB.Form frmmouse 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1620
   ClientLeft      =   3900
   ClientTop       =   2295
   ClientWidth     =   1320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   1320
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Image imgchange 
      Height          =   330
      Index           =   3
      Left            =   120
      Picture         =   "frmmouse.frx":0000
      Top             =   1200
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgchange 
      Height          =   330
      Index           =   2
      Left            =   120
      Picture         =   "frmmouse.frx":018A
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgchange 
      Height          =   330
      Index           =   1
      Left            =   120
      Picture         =   "frmmouse.frx":0314
      Top             =   360
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgchange 
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "frmmouse.frx":049E
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "文档"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   3
      Left            =   120
      Picture         =   "frmmouse.frx":0628
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "发信"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   2
      Left            =   120
      Picture         =   "frmmouse.frx":07B2
      Top             =   840
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderStyle     =   6  'Inside Solid
      X1              =   0
      X2              =   1320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "移动"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   1
      Left            =   120
      Picture         =   "frmmouse.frx":093C
      Top             =   360
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "frmmouse.frx":0AC6
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmmouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public former As Integer
Private Sub changecolor(index As Integer)


Image1(former).Visible = True
imgchange(former).Visible = False
Label1(former).ForeColor = &HFF0000

Image1(index).Visible = False
imgchange(index).Visible = True
Label1(index).ForeColor = QBColor(5)

former = index

End Sub


Private Sub Form_loadpicture()
Load frmmouse
End Sub

Private Sub Image1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
changecolor (index)
End Sub

Private Sub Label1_Click(index As Integer)
Select Case index
Case 0
del
Case 1
moveit
Case 2
mail
Case 3
documents
End Select
frmmouse.Hide
End Sub

Private Sub del()
'add codes here
End Sub

Private Sub moveit()
'add codes here
End Sub

Private Sub mail()
'add codes here
End Sub

Private Sub documents()
'nt下用这个路径
Shell "c:\winnt\notepad.exe", vbNormalFocus
'98下用这个路径
'Shell "c:\windows\notepad.exe", vbNormalFocus 'you could add other codes here
End Sub

Private Sub Label1_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
changecolor (index)
End Sub

