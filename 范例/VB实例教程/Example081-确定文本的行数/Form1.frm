VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "计算文本行数"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4095
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "计算行数"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib _
"user32" Alias "SendMessageA" _
( _
  ByVal hwnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long _
  ) As Long
Private Const EM_GETLINECOUNT = &HBA

Private Sub Command1_Click()
    Dim lineCount As Integer
    lineCount = SendMessageLong(Text1.hwnd, _
               EM_GETLINECOUNT, 0&, 0&)
    MsgBox Str(lineCount), vbInformation + vbOKOnly, "行数"
End Sub
