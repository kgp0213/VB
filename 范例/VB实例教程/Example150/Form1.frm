VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Unicode”ÎAnsi"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   2655
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command_Count 
      Caption         =   "Count"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Count_Click()
    Dim str_Input As String
    str_Input = Me.Text1.Text
    
    Dim strAnsi As String
    Dim strUnicode As String
    strUnicode = Me.Text1.Text
    strAnsi = StrConv(strUnicode, vbFromUnicode)
    str_Input = str_Input + Chr(13)
    str_Input = str_Input + "Unicode:" + str(Len(strUnicode)) + Chr(13)
    str_Input = str_Input + "Ansi:" + str(LenB(strAnsi))
    MsgBox str_Input
End Sub
