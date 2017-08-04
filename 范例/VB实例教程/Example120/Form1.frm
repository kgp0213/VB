VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CD-ROM"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1935
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   1935
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "CD-ROM"
      DownPicture     =   "Form1.frx":0000
      Height          =   735
      Left            =   360
      Picture         =   "Form1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" _
                Alias "mciSendStringA" _
                (ByVal lpstrCommand As String, _
                ByVal lpstrReturnString As String, _
                ByVal uReturnLength As Long, _
                ByVal hwndCallback As Long) _
                As Long

Private Sub Command1_Click()
   If Me.Command1.ToolTipText = "弹出光盘" Then
      retvalue = mciSendString("set CDAudio door open", _
                                returnstring, 127, 0)
      Me.Command1.ToolTipText = "关闭光驱"
    Else
      retvalue = mciSendString("set CDAudio door closed", _
                                returnstring, 127, 0)
      Me.Command1.ToolTipText = "弹出光盘"
    End If
End Sub

Private Sub Form_Load()
    Me.Command1.ToolTipText = "弹出光盘"
End Sub
