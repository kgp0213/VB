VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ascii码查询器"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   2070
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   840
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   8
      List            =   "Form1.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myWidth As Single
Private myHeight As Single
Private myShowed As Single
Private Sub Form_Load()
myWidth = Width
myHeight = Height
End Sub
Private Sub Form_Resize()
Width = myWidth
Height = myHeight
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub
Private Sub Text1_Change()
Dim myKeyAscii As String
If Text1.Text <> "" Then
myKeyAscii = Text1.Text + " ------- " + Text2.Text
List1.AddItem (myKeyAscii)
End If
End Sub
Private Sub Text1_GotFocus()
If myShowed <> 1 Then
    Form2.Show
    Form2.Cls
    Form2.Width = 1780
    Form2.Print "请在这里键入一个字符"
    myShowed = 1
    Form1.SetFocus
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
Text1.Text = ""
Text2.Text = KeyAscii
Text2.SetFocus
Text2.SelStart = 0
Text2.SelLength = 4
End Sub
Private Sub Text2_GotFocus()
If myShowed <> 2 Then
    Form2.Show
    Form2.Cls
    Form2.Width = 1690
    Form2.Print "这儿是相应的Ascii码"
    myShowed = 2
    Form1.SetFocus
End If
End Sub
Private Sub Timer1_Timer()
Form2.Left = Form1.Left + Form1.Width
Form2.Top = Form1.Top + Form1.Height
End Sub
