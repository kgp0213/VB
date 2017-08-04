VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Keyboard"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   3150
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub keybd_event Lib "user32" _
                (ByVal bVk As Byte, _
                ByVal bScan As Byte, _
                ByVal dwFlags As Long, _
                ByVal dwExtraInfo As Long)
                
Private Const VK_CONTROL = &H11
Private Const VK_SHIFT = &H10
Private Const VK_MENU = &H12

Private Declare Function MapVirtualKey Lib "user32" _
                Alias "MapVirtualKeyA" _
                (ByVal wCode As Long, _
                ByVal wMapType As Long) _
                As Long

Private Sub Text1_Click()
    Call keybd_event(VK_CONTROL, MapVirtualKey(VK_CONTROL, 0), 0, 0)
    Call keybd_event(Asc("V"), MapVirtualKey(Asc("V"), 0), 0, 0)
    Call keybd_event(Asc("V"), MapVirtualKey(Asc("V"), 0), KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_CONTROL, MapVirtualKey(VK_CONTROL, 0), _
                     EYEVENTF_KEYUP, 0)
End Sub
