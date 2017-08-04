VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WindowText"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2250
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   2250
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command_Set 
      Caption         =   "SetWindowText"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command_Get 
      Caption         =   "GetWindowText"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowText Lib "user32" _
                Alias "SetWindowTextA" _
                (ByVal hwnd As Long, _
                ByVal lpString As String) _
                As Long
Private Declare Function GetWindowText Lib "user32" _
                Alias "GetWindowTextA" _
                (ByVal hwnd As Long, _
                ByVal lpString As String, _
                ByVal cch As Long) _
                As Long

Private Sub Command_Get_Click()
        Dim strText As String * 256
        Dim cch As Long
        cch = GetWindowText(Me.hwnd, strText, 256)
        Me.Text1.Text = Left(strText, cch)
End Sub

Private Sub Command_Set_Click()
        Call SetWindowText(Me.hwnd, Me.Text1.Text)
End Sub

