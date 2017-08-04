VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2565
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   1815
   ScaleWidth      =   2565
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "悬挂式窗口"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
 Dim res As Long
 If Check1.Value = 1 Then
    res = SetWindowPos(Me.hwnd, HWND_TOPMOST, _
                        0, 0, 0, 0, Flags)
 Else
    res = SetWindowPos(Me.hwnd, HWND_NOTOPMOST, _
                        0, 0, 0, 0, Flags)

 End If
End Sub
