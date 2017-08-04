VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê¾Àý"
   ClientHeight    =   3480
   ClientLeft      =   1860
   ClientTop       =   1680
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5325
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Caption         =   "¹Ø±Õ"
      Height          =   495
      Left            =   1800
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "µã»÷ÓÒ¼ü"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With frmmouse
 If Button = 2 Then
   .Left = X + .Width + 500
   .Top = Y + .Height
   .Show
   Dim i As Integer
     For i = 0 To 3
     .Image1(i).Visible = True
      .imgchange(i).Visible = False
     .Label1(i).ForeColor = &HFF0000
     Next
   Else
  .Hide
 End If
End With


End Sub

Private Sub Form_Terminate()
Unload frmmouse
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmmouse
End Sub
